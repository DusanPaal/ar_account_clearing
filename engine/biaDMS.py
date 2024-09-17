# pylint: disable =C0103, C0123, W0603, W0703, W1203

"""
The 'biaDMS.py' module uses the standard SAP GUI UDM_DISPUTE transaction
in order to automate searching, export and updating disputed case data.

Version history:
1.0.20210526 - initial version
1.0.20210908 - removed SrchMask parameter from the Close() procedure
			   and any related logic
			 - added assertions as input check to public procedures
1.0.20220504 - removed unused virtual key mapping fom _vkeys {}
1.0.20220803 - fixed regression in modify_case_parameters() when automation
			   closed the DMS search mask and mistakenly returned to the
			   initial DMS screen after handling the dialog window
			   'Attributs may be overwritten later' that appears if
			   case Processor/Coordinator no longer exists in SAP.
1.0.20220906 - Minor code style improvements.
"""

from enum import IntEnum
import logging
from os.path import exists, isfile, split
from pyperclip import copy as copy_to_clipboard
from win32com.client import CDispatch

# custom exceptions
class CaseEditingError(Exception):
	"""
	Raised when the 'change' mode
	cannot be entered/escaped for
	a case.
	"""

class DataWritingError(Exception):
	"""
	Raised when writing of accounting
	data to file fails.
	"""

class FolderNotFoundError(Exception):
	"""
	Raised when the folder to which
	data should be exported doesn't exist.
	"""

class LayoutNotFoundError(Exception):
	"""
	Raised when the supplied layout
	name is not found in the list
	of available layouts.
	"""

class NoCaseFoundError(Exception):
	"""
	Raised when case searching
	returns no result.
	"""

class SavingChangesError(Exception):
	"""
	Raised when the changes to
	a case cannot be saved.
	"""

class StatusAcError(Exception):
	"""
	Raised when Status AC
	has incorrect value.
	"""

class TransactionNotStartedError(Exception):
	"""
	Raised when attempting to use a procedure
	before starting the transaction.
	"""

# module enums
class CaseStates(IntEnum):
	"""
	Available DMS case 'Status' values.
	"""
	Original = 0
	Open = 1
	Solved = 2
	Closed = 3
	Devaluated = 4

# keyboard to SAP virtual keys mapping
_vkeys = {
	"Enter":    0,
	"F3":       3,
	"F8":       8,
	"CtrlS":    11,
	"F12":      12,
	"ShiftF4":  16,
	"ShiftF12": 24
}

# private vars
_sess = None
_main_wnd = None
_stat_bar = None

_logger = logging.getLogger("master")

# procedures
def _is_error_message(sbar: CDispatch) -> bool:
	"""
	Checks if a status bar message
	is an error message.
	"""

	if sbar.messageType == "E":
		return True

	return False

def _is_popup_dialog() -> bool:
	"""
	Checks if the active window
	is a popup dialog window.
	"""

	if _sess.ActiveWindow.type == "GuiModalWindow":
		return True

	return False

def _close_popup_dialog(confirm: bool):
	"""
	Confirms or delines a pop-up dialog.
	"""

	if _sess.ActiveWindow.text == "Information":
		if confirm:
			_main_wnd.SendVKey(_vkeys["Enter"]) # confirm
		else:
			_main_wnd.SendVKey(_vkeys["F12"])   # decline
		return

	btn_caption = "Yes" if confirm else "No"

	for child in _sess.ActiveWindow.Children:
		for grandchild in child.Children:
			if grandchild.Type != "GuiButton":
				continue
			if btn_caption != grandchild.text.strip():
				continue
			grandchild.Press()
			return

def _get_grid_view():
	"""
	Returns a GuiGridView object representing
	the DMS window containing search results.
	"""

	splitter_shell = _main_wnd.FindByName("shell", "Guisplitter_shell")
	grid_view = splitter_shell.FindAllByName("shell", "GuiGridView")(6)

	return grid_view

def _get_param_mask() -> object:
	"""
	Returns a GuiGridView object representing
	the DMS case parameter mask containing editable fields.
	"""

	splitter_shell = _main_wnd.FindByName("shell", "Guisplitter_shell")
	param_mask = splitter_shell.FindAllByName("shell", "GuiGridView")(5)

	return param_mask

def _execute_query() -> int:
	"""
	Simulates pressing the 'Search' button
	located on the DMS main search mask.
	Returned is the number of cases found.
	"""

	splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
	qry_toolbar = splitter_shell.FindAllByName("shell", "GuiToolbarControl")(5)
	qry_toolbar.PressButton("DO_QUERY")

	num = _stat_bar.Text.split(" ")[0]
	num = num.strip().replace(".", "")

	return int(num)

def _find_and_click_node(tree: object, node: object, node_id: str) -> bool:
	"""
	Traverses the left-sided DMS menu tree to find the item with the given node ID.
	Once the item is found, the procedure simulates clicking on that item to open
	the corresponding subwindow.
	"""

	# find and double click the target root node
	if tree.IsFolder(node):
		tree.CollapseNode(node)
		tree.ExpandNode(node)

	# double clisk the target node
	if node.strip() == node_id:
		tree.DoubleClickNode(node)
		return True

	subnodes = tree.GetsubnodesCol(node)

	if subnodes is None:
		return False

	iter_subnodes = iter(subnodes)

	if _find_and_click_node(tree, next(iter_subnodes), node_id):
		return True

	try:
		next_node = next(iter_subnodes)
	except StopIteration:
		return False
	else:
		return _find_and_click_node(tree, next_node, node_id)

def _get_search_mask():
	"""
	Returns the GuiGridView object representing
	the DMS case search window.
	"""

	# find the target node by traversing the search tree
	tree = _main_wnd.findById(
		"shellcont/shell/shellcont[0]/shell/shellcont[1]/shell/shellcont[1]/shell"
	)

	nodes = tree.GetNodesCol()
	iter_nodes = iter(nodes)
	clicked = _find_and_click_node(tree, next(iter_nodes), node_id = "4")

	assert clicked, "Target node not found!"

	# get reference to the search mask object found
	splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
	srch_mask = splitter_shell.FindAllByName("shell", "GuiGridView")(4)

	return srch_mask

def _change_case_param(param_mask, cell_idx, cell_val_type, val):
	"""
	Changes the value of an editable case parameter field identified
	in the case parameter grid by the cell index and column type.
	"""

	param_mask.ModifyCell(cell_idx, cell_val_type, val)

	return

def _get_case_param(param_mask, cell_idx, cell_val_type):
	"""
	Returns the value of an editable case parameter field identified
	in the case parameter grid by the cell index and column type.
	"""

	cell_val = param_mask.GetCellValue(cell_idx, cell_val_type)

	return cell_val

def _get_control_toolbar() -> CDispatch:
	"""
	Returns GuiToolbarControl object representing the DMS control toolbar
	located in the transaction upper window.
	"""

	splitter_shell = _main_wnd.FindByName("shell", "Guisplitter_shell")
	toolbar = splitter_shell.FindAllByName("shell", "GuiToolbarControl")(3)

	return toolbar

def _save_changes(ctrl_toolbar: CDispatch) -> tuple:
	"""
	Simulates pressing the 'Save' button located
	in the DMS upper toolbar.
	"""

	ctrl_toolbar.PressButton("SAVE")

	if _is_statbar_error():
		err_msg = _stat_bar.Text
		return (False, err_msg)

	return (True, "")

def _change_case_status(param_mask: CDispatch, val: CaseStates):
	"""
	Changes the case 'Status' parameter.
	"""

	status_map = {
		CaseStates.Open: "Open",
		CaseStates.Solved: "Solved",
		CaseStates.Closed: "Closed",
	}

	prev_val = _get_case_param(param_mask, 0, "VALUE2")
	new_val = status_map[val]

	while prev_val != new_val:

		if prev_val == status_map[CaseStates.Open]:
			curr_val = status_map[CaseStates.Solved]
			_change_case_param(param_mask, 0, "VALUE2", curr_val)

			if new_val == status_map[CaseStates.Closed]:
				_save_changes(_get_control_toolbar())

		elif prev_val == status_map[CaseStates.Closed]:
			curr_val = status_map[CaseStates.Solved]
			_change_case_param(param_mask, 0, "VALUE2", curr_val)

			if new_val ==status_map[CaseStates.Open]:
				_save_changes(_get_control_toolbar())

		elif prev_val == status_map[CaseStates.Solved]:
			curr_val = new_val
			_change_case_param(param_mask, 0, "VALUE2", curr_val)

		prev_val = curr_val

def _apply_layout(grid_view: CDispatch, name: str):
	"""
	Searches a layout by name in the DMS layouts list. If the layout is
	found in the list of available layouts, this gets selected.
	"""

	# Open Change Layout Dialog
	grid_view.PressToolbarContextButton("&MB_VARIANT")
	grid_view.SelectContextMenuItem("&LOAD")
	apo_grid = _sess.findById("wnd[1]").findAllByName("shell", "GuiShell")(0)

	for row_idx in range(0, apo_grid.RowCount):
		if apo_grid.GetCellValue(row_idx, "VARIANT") == name:
			apo_grid.setCurrentCell(str(row_idx), "TEXT")
			apo_grid.clickCurrentCell()
			return True

	raise LayoutNotFoundError(f" Layout '{name}' not found!")

def _select_data_format(grid_view: CDispatch, idx: int):
	"""
	Selects the 'Unconverted' file format
	from file export format option window.
	"""

	grid_view.PressToolbarContextButton("&MB_EXPORT")
	grid_view.SelectContextMenuItem("&PC")
	option_wnd = _sess.FindById("wnd[1]")
	option_wnd.FindAllByName("SPOPLI-SELFLAG", "GuiRadioButton")(idx).Select()

def _is_statbar_error() -> bool:
	"""
	Checks if a status bar
	message is an error message.
	"""

	is_err = _stat_bar.MessageType == "E"

	return is_err

def _toggle_display_change(activate: bool) -> bool:
	"""
	Enables editing of a case.
	"""

	_get_control_toolbar().PressButton("TOGGLE_DISPLAY_CHANGE")

	msg = _stat_bar.Text

	if "display only" in msg:
		return (False, msg)

	if not activate:
		_main_wnd.SendVKey(_vkeys["F3"])

	# handle alert dialogs for non-editable cases
	if _is_popup_dialog():

		err_msg = _sess.ActiveWindow.children(1).children(1).text

		if err_msg == "Attributes may be overwritten later":
			_close_popup_dialog(confirm = True)
		else:
			_close_popup_dialog(confirm = False)
			_main_wnd.SendVKey(_vkeys["F3"])
			return (False, err_msg)

	elif _is_error_message(_stat_bar):
		err_msg = _stat_bar.Text
		return (False, err_msg)

	return (True, "")

def _copy_to_searchbox(srch_mask: CDispatch, cases: tuple):
	"""
	Copies case ID numbers into
	the search listbox.
	"""

	srch_mask.PressButton(0, "SEL_ICON1")
	cases = map(str, cases)
	_main_wnd.SendVKey(_vkeys["ShiftF4"])       # clear any previous values
	copy_to_clipboard("\r\n".join(cases))       # copy accounts to clipboard
	_main_wnd.SendVKey(_vkeys["ShiftF12"])      # confirm selection
	copy_to_clipboard("")                       # clear the clipboard content
	_main_wnd.SendVKey(_vkeys["F8"])            # confirm

def _set_hits_limit(srch_mask: CDispatch, n: int):
	"""
	Enters an iteger that restricts
	the he number of found records.
	"""

	MAX_DISPUTES = 5000

	if n > MAX_DISPUTES:
		raise ValueError(f"Cannot search more than {MAX_DISPUTES} cases!")

	if n == 0:
		raise ValueError("No case ID numbers provided!")

	srch_mask.ModifyCell(23, "VALUE1", MAX_DISPUTES)

def _set_case_id(srch_mask: CDispatch, val: int):
	"""
	Enters case ID value into the corresponding
	field located on the search mask.
	"""

	srch_mask.ModifyCell(0, "VALUE1", str(val))

def _export_to_file(grid_view: CDispatch, file_path: str, enc: str = "4120"):
	"""
	Exports loaded accounting data to a text file.
	"""

	folder_path, file_name = split(file_path)

	if not exists(folder_path):
		raise FolderNotFoundError(f"Export folder not found: {folder_path}")

	# select 'Unconverted' data format
	# and confirm the selection
	_select_data_format(grid_view, 0)
	_main_wnd.SendVKey(_vkeys["Enter"])

	# enter data export file name and folder path
	_sess.FindById("wnd[1]").FindByName("DY_PATH", "GuiCTextField").text = folder_path
	_sess.FindById("wnd[1]").FindByName("DY_FILENAME", "GuiCTextField").text = file_name
	_sess.FindById("wnd[1]").FindByName("DY_FILE_ENCODING", "GuiCTextField").text = enc

	# replace an exiting file
	_main_wnd.SendVKey(_vkeys["CtrlS"])

	# double check if data export succeeded
	if not isfile(file_path):
		raise DataWritingError(f"Failed to export data to file: {file_path}")

def start(sess: CDispatch) -> CDispatch:
	"""
	Starts UDM_DISPUTE transaction.

	Params:
	------
	sess:
		A GuiSession object.

	Returns:
	-------
	A GuiGridView object representing
	the transaction's search window.
	"""

	_logger.info("Starting UDM_DISPUTE ...")

	global _sess
	global _main_wnd
	global _stat_bar

	_sess = sess
	_main_wnd = _sess.findById("wnd[0]")
	_stat_bar = _main_wnd.findById("sbar")

	_sess.StartTransaction("UDM_DISPUTE")
	srch_mask = _get_search_mask()

	_logger.info("transaction running.")

	return srch_mask

def close():
	"""
	Closes a running UDM_DISPUTE transaction.

	Params:
	-------
	None.

	Returns:
	--------
	None.

	Raises:
	-------
	TransactionNotStartedError:
		When attempting to close
		UDM_DISPUTE when it's not running.
	"""

	_logger.info("Closing UDM_DISPUTE ...")

	global _sess
	global _main_wnd
	global _stat_bar

	if _sess is None:
		raise TransactionNotStartedError("Cannot close UDM_DISPUTE when it's actually not running!"
		"Use the biaDMS.start() procedure to run the transaction first of all.")

	_sess.EndTransaction()

	if _is_popup_dialog():
		_close_popup_dialog(confirm = True)

	_sess = None
	_main_wnd = None
	_stat_bar = None

def search_dispute(search_mask: CDispatch, case: int) -> CDispatch:
	"""
	Searches a disputed case in DMS.

	Params:
	-------
	search_mask:
		A GuiGridView object representing the transaction's search window.

	case:
		Identification number of the case in DMS.

	Returns:
	--------
	A GuiGridView object representing the search results.

	Raises:
	-------
	NoCaseFoundError:
		When searching returns no record for the given case.
	"""

	_set_case_id(search_mask, case)
	n_found = _execute_query()

	if n_found > 0:
		search_results = _get_grid_view()
	else:
		raise NoCaseFoundError(f"Case '{case}' not found! The entered value refers "
		"to an non-existing case or access to the case company code is missing.")

	return search_results

def search_disputes(search_mask: CDispatch, cases: list) -> CDispatch:
	"""
	Searches disputed cases in DMS.

	Params:
	-------
	search_mask:
		A GuiGridView object representing the transaction's search window.

	cases:
		Identification numbers of the cases in DMS.

	Returns:
	--------
	A GuiGridView object representing the search results.

	Raises:
	-------
	NoCaseFoundError:
		When searching returns no result for the given cases.
	"""

	n_total = len(cases)

	invalid_cases = []

	for case in cases:
		if not str(case).isnumeric():
			invalid_cases.append(case)

	if len(invalid_cases) != 0:
		vals = ';'.join(invalid_cases)
		raise ValueError(f"Argument 'cases' contains invalid value: {vals}")

	# hit limit should be equal to the num of cases
	_set_hits_limit(search_mask, n_total)

	# search cases
	_copy_to_searchbox(search_mask, cases)
	n_found = _execute_query()

	if n_found > 0:
		search_results = _get_grid_view()
	else:
		vals = "; ".join(cases)
		raise NoCaseFoundError("The entered values refer to non-existing cases or access "
		f"to the entire company code is missing. List of values:'{vals}' not found! ")

	if 0 < n_found < n_total:
		_logger.warning(f"Incorrect disputes detected: {n_total - n_found}. "
		"There might be a typo in the case ID provided in item 'Text' value(s).")

	return search_results

def modify_case_parameters(search_result: CDispatch, root_cause: str,
						   status_ac: str, status: CaseStates):
	"""
	Modifies the following parameters of a disputed case stored in DMS:
		- root cause code
		- status AC
		- case status

	Params:
	-------
	search_result:
		A GuiGridView object representing case(s) search result.

	root_cause:
		Represents the 'Root Cause Code' value of a disputed case.

	stat_sl:
		Represents the 'Status Sales' value of a disputed case.

	stat:
		Represents the 'Status' value of a disputed case.

	Returns:
	--------
	None.

	Raises:
	-------
	CaseEditingError:
		When the 'Change case' mode cannot be entered/escaped.

	SavingChangesError:
		When changes made to the case parameters cannot be saved.
	"""

	# check input
	valid_root_causes = ["L06", "L01"] # other vals are NA for account clearing
	valid_states = list(map(lambda st: st, CaseStates))

	if root_cause not in valid_root_causes:
		raise ValueError(f"Argument 'root_cause' has incorrect value: {root_cause}")

	MAX_STAT_AC_CHARS = 50

	if len(status_ac) > MAX_STAT_AC_CHARS:
		raise StatusAcError(f"Argument 'status_ac' too long! Number of chars is "
		f"{len(status_ac)} while a maximum of {MAX_STAT_AC_CHARS} chars is allowed!")

	if status not in valid_states:
		raise ValueError("Argument 'status' has incorrect value!")

	# open case details
	search_result.DoubleClickCurrentCell()

	# enter edit mode
	activated, err_msg = _toggle_display_change(activate = True)

	if not activated:
		_main_wnd.SendVKey(_vkeys["F3"])
		raise CaseEditingError(err_msg)

	param_mask = _get_param_mask()

	_change_case_param(param_mask, 10, "VALUE2", root_cause)
	_change_case_param(param_mask, 11, "VALUE2", status_ac)
	_change_case_status(param_mask, status)

	# save the modified params
	saved, err_msg = _save_changes(_get_control_toolbar())

	if not saved:
		_main_wnd.SendVKey(_vkeys["F3"])
		if _is_popup_dialog():
			_close_popup_dialog(confirm = False)
		raise SavingChangesError(err_msg.replace(" +", " "))

	# exit edit mode
	deactivated, err_msg = _toggle_display_change(activate = False)

	if not deactivated:
		raise CaseEditingError(err_msg)

def export(search_result: CDispatch, file_path: str, layout: str):
	"""
	Exports disputed data into a plain text file.

	Params:
	-------
	search_result:
		A GuiGridView object representing case(s) search result.

	file_path:
		Path to the file to which the data will be exported.

	layout:
		Name of the layout defining format of the loaded/exported data.

	Returns:
	--------
	None.

	Raises:
	-------
	FolderNotFoundError:
		When the folder to which data should be exported doesn't exist.

	LayoutNotFoundError:
		When the used layout is not found in the list of available layouts.

	DataWritingError:
		When writing of the data to a file fails.
	"""

	_apply_layout(search_result, layout)
	_export_to_file(search_result, file_path)
