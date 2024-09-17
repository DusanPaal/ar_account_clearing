# pylint: disable = C0103, C0123, W0603, W0703, W1203

"""
The 'biaFBL5N.py' module automates the standard SAP GUI FBL5N transaction
in order to load and export data located on customer accounts to a plain
text file.

Version history:
1.0.20210412 - Initial version.
1.0.20220112 - Removed dymamic layout creation upon data load.
               layout will now be applied through layout name
               in the main search mask.
1.0.20220504 - Removed unused virtual key mapping fom _vkeys {}.
1.0.20220906 - Minor code style improvements.
"""

import logging
from os.path import exists, split
from win32com.client import CDispatch

# custom warnings
class NoDataFoundWarning(Warning):
    """
    Warns that there are no open
    items available on account.
    """

# custom exceptions
class AbapRuntimeError(Exception):
    """
    Raised when SAP 'ABAP Runtime Error'
    occurs during communication with
    the transaction.
    """

class ConnectionLostError(Exception):
    """
    Raised when a connection to SAP
    is lost as a result of a network error.
    """

class FolderNotFoundError(Exception):
    """
    Raised when the folder to which
    data should be exported doesn't exist.
    """

class ItemsLoadingError(Exception):
    """
    Raised when loading of open
    items fails.
    """

class SapRuntimeError(Exception):
    """
    Raised when an unhanded general SAP
    error occurs during communication with
    the transaction.
    """

class WorklistNotFoundError(Exception):
    """
    Raised when FBL5N dispays an error
    message in the status bar indicating
    that the used worklist does not exist.
    """

# keyboard to SAP virtual keys mapping
_vkeys = {
    "Enter":  0,
    "F3":     3,
    "F8":     8,
    "F9":     9,
    "CtrlS":  11,
    "F12":    12,
    "CtrlF1": 25
}

_sess = None
_main_wnd = None
_stat_bar = None


def _is_popup_dialog() -> bool:
    """
    Checks if the active window
    is a popup dialog window.
    """

    if _sess.ActiveWindow.type == "GuiModalWindow":
        return True

    return False

def _is_sap_runtime_error(main_wnd: CDispatch) -> bool:
    """
    Checks if a SAP ABAP runtime error exists.
    """

    if main_wnd.text == "ABAP Runtime Error":
        return True

    return False

def _is_error_message(sbar: CDispatch) -> bool:
    """
    Checks if a status bar message
    is an error message.
    """

    if sbar.messageType == "E":
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

def _clear_customer_acc(worklist: bool):
    """
    Clears any value in the 'Customer account' field
    located on thetransaction main search window.
    """

    if worklist:
        fld_name = "PA_WLKUN"
    else:
        fld_name = "DD_KUNNR-LOW"

    _main_wnd.findByName(fld_name, "GuiCTextField").text = ""

def _set_layout(name: str):
    """
    Enters layout name into the 'Layout' field
    located on the transaction main window.
    """

    _main_wnd.findByName("PA_VARI", "GuiCTextField").text = name

def _set_company_code(val: str):
    """
    Enters company code into the 'Company code' field
    located on the transaction main window.
    """

    fld_tech_name = ""

    if _main_wnd.findAllByName("DD_BUKRS-LOW", "GuiCTextField").count > 0:
        fld_tech_name = "DD_BUKRS-LOW"
    elif _main_wnd.findAllByName("SO_WLBUK-LOW", "GuiCTextField").count > 0:
        fld_tech_name = "SO_WLBUK-LOW"

    _main_wnd.findByName(fld_tech_name, "GuiCTextField").text = val

def _set_worklist(name: str):
    """
    Enters worklist name into the 'Worklist' field
    located on the transaction main window.
    """

    _main_wnd.FindByName("PA_WLKUN", "GuiCTextField").text = name.upper()

def _toggle_worklist(activate: bool):
    """
    Activates or deactivates the 'Use worklist' option
    in the transaction main search mask.
    """

    used = _main_wnd.FindAllByName("PA_WLKUN", "GuiCTextField").Count > 0

    if (activate and not used) or (not activate and used):
        _main_wnd.SendVKey(_vkeys["CtrlF1"])

def _export_to_file(folder_path: str, file_name: str):
    """
    Exports loaded accounting data to a text file.
    """

    _main_wnd.SendVKey(_vkeys["F9"])     # open local data file export dialog
    _select_data_format(0)               # set plain text data export format
    _main_wnd.SendVKey(_vkeys["Enter"])  # confirm

    _sess.FindById("wnd[1]").FindByName("DY_PATH", "GuiCTextField").text = folder_path
    _sess.FindById("wnd[1]").FindByName("DY_FILENAME", "GuiCTextField").text = file_name
    _sess.FindById("wnd[1]").FindByName("DY_FILE_ENCODING", "GuiCTextField").text = "4120"

    _main_wnd.SendVKey(_vkeys["CtrlS"])  # replace an exiting file
    _main_wnd.SendVKey(_vkeys["F3"])     # Load main mask

def _select_data_format(idx: int):
    """
    Selects data export format from the export options dialog
    based on the option index on the list.
    """

    option_wnd = _sess.FindById("wnd[1]")
    option_wnd.FindAllByName("SPOPLI-SELFLAG", "GuiRadioButton")(idx).Select()

def _start(sess: CDispatch):
    """
    Starts FBL5N transaction.

    Params:
    ------
    sess:
        A GuiSession object.

    Returns:
    -------
    None.
    """

    global _sess
    global _main_wnd
    global _stat_bar

    _sess = sess
    _main_wnd = _sess.findById("wnd[0]")
    _stat_bar = _main_wnd.findById("sbar")

    _sess.StartTransaction("FBL5N")

def _close():
    """
    Closes a running FBL5N transaction.

    Params:
    -------
    None.

    Returns:
    --------
    None.
    """

    global _sess
    global _main_wnd
    global _stat_bar

    _sess.EndTransaction()

    if _is_popup_dialog():
        _close_popup_dialog(confirm = True)

    _sess = None
    _main_wnd = None
    _stat_bar = None

def _load_items():
    """
    Simulates pressing the 'Execute'
    button that triggers item loading.
    """

    _main_wnd.SendVKey(_vkeys["F8"])

    try: # SAP crash can be caught only after next statement following item loading
        msg = _stat_bar.Text
    except Exception as exc:
        raise ConnectionLostError("Connection to SAP lost due to an network error.") from exc

    if _is_sap_runtime_error(_main_wnd):
        raise SapRuntimeError("SAP runtime error!")

    if "items displayed" not in msg:
        raise NoDataFoundWarning(msg)

    if "The current transaction was reset" in msg:
        raise SapRuntimeError("FBL5N was unexpectedly terminated!")

    if _is_error_message(_stat_bar):
        raise ItemsLoadingError(msg)

    if _main_wnd.text == 'ABAP Runtime Error':
        raise AbapRuntimeError("Data loading failed due to an ABAP runtime error.")

    if "does not exist" in msg:
        raise WorklistNotFoundError(msg)

def export(sess: CDispatch, file_path: str, company_code: str,
           worklist: str = None, layout: str = ""):
    """
    Exports data from customer accounts into a plain text file.

    Params:
    -------
    file_path:
        Path to the file to which the data will be exported.

    company_code:
        Company code for which the data will be exported. \n
        If used, then any value passed to the 'worklist'
        param will be ignored.

    worklist:
        Name of a FBL5N worklist. \n
        If 'cocd' param is provided, then the 'worklist' value will be ignored.

    layout:
        Name of the layout (default "") defining format of the loaded/exported data.

    Returns:
    -------
    None.

    Raises:
    -------
    NoDataFoundWarning:
        If there are no open items available on account(s).

    AbapRuntimeError:
        If a SAP 'ABAP Runtime Error' occurs during transaction runtime.

    ConnectionLostError:
        If  a connection to SAP is lost
        as a result of a network error.

    FolderNotFoundError:
        When the folder to which data should be exported doesn't exist.

    ItemsLoadingError:
        If loading of open items fails.

    SapRuntimeError:
        If an unhanded general SAP error occurs
        during communication with the transaction.

    WorklistNotFoundError:
        If the used worklist does not exist in FBL5N.
    """

    folder_path, file_name = split(file_path)

    if not exists(folder_path):
        raise FolderNotFoundError(f"Export folder not found: {folder_path}")

    if not file_path.endswith(".txt"):
        raise ValueError(
            f"Invalid file type: {file_path}. "
            "Only '.txt' file types are supported."
        )

    _start(sess)

    if worklist is None:
        _toggle_worklist(activate = False)
        _clear_customer_acc(worklist = False)
    else:
        _toggle_worklist(activate = True)
        _clear_customer_acc(worklist = True)
        _set_worklist(worklist)

    _set_layout(layout)
    _set_company_code(company_code)

    try:
        _load_items()
    except:
        _close()
        raise

    _export_to_file(folder_path, file_name)
    _close()
