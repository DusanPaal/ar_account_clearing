# pylint: disable = C0103, C0123, R0913, R0914, W0603, W0703, W1203

"""
The 'biaF30.py' module uses the standard SAP GUI F-30 transaction
in order to load, select and post open items located on customer
accounts.

Version history:
    1.0.20210526 - Initial version.
    1.0.20220329 - Removed entering of business area to items being
                   transferred in '_transfer_rests()' procedure.
    1.0.20220504 - removed unused virtual key mapping fom _vkeys {}
    1.0.20220906 - Minor code style improvements and bugfixes.
"""

import logging
import time
from datetime import date, datetime
from typing import Collection
from win32com.client import CDispatch


# custom exceptions
class AccountBlockedError(Exception):
    """
    Raised when an acount is blocked
    by a user and cannot be accessed
    by the application.
    """

class DocumentNotFoundError(Exception):
    """
    Raised when F-30 fails to locate
    a specific document on a customer
    account.
    """

class ItemPostingError(Exception):
    """
    Raised when an error occured
    during posting of the selected
    open items.
    """

class ItemSelectionError(Exception):
    """
    Raised when selection of relevant
    open items from a list of all the
    loaded accounting items leaves
    non-zero amount as the final balance.
    """

class LayoutNotFoundError(Exception):
    """
    Raised when the supplied layout
    name is not found in the list
    of available layouts.
    """

class SapGeneralError(Exception):
    """
    Raised when F-30 displays
    an alert or warning for
    which no appropriata handler
    exists.
    """

class TransactionNotStartedError(Exception):
    """
    Raised when attempting to use a procedure
    before starting the transaction.
    """

# keyboard to SAP virtual keys mapping
_vkeys = {
    "Enter":   0,
    "F2":      2,
    "F6":      6,
    "F12":     12,
    "ShiftF2": 14,
    "ShiftF4": 16
}

_sess = None
_main_wnd = None
_stat_bar = None
_logger = logging.getLogger("master")

def _parse_amount(num: str) -> float:
    """
    Converts an amount value represented
    as SAP string format to a float data type.
    """

    parsed = num.strip()
    parsed = parsed.replace(".", "")
    parsed = parsed.replace(",", ".")

    if parsed.endswith("-"):
        parsed = parsed.replace("-", "")
        parsed = "".join(["-", parsed])

    return float(parsed)

def _is_alert_message() -> bool:
    """
    Checks if a message contained in
    the status bar is an alert message.
    """
    return _is_warning_message() or _is_error_message()

def _is_warning_message() -> bool:
    """
    Checks if a message contained in
    the status bar is a warning message.
    """
    return _stat_bar.MessageType == "W"

def _is_error_message() -> bool:
    """
    Checks if a message contained in
    the status bar is an error message.
    """
    return _stat_bar.MessageType == "E"

def _is_popup_dialog() -> bool:
    """
    Checks if the active window
    is a popup dialog window.
    """
    return _sess.ActiveWindow.type == "GuiModalWindow"

def _confirm() -> None:
    """Simulates pressing the 'Enter' button."""
    _main_wnd.SendVKey(_vkeys["Enter"])

def _decline() -> None:
    """Simulates pressing the 'F12' button."""
    _main_wnd.SendVKey(_vkeys["F12"])

def _close_popup_dialog(confirm: bool) -> None:
    """Confirms or delines a pop-up dialog."""

    if _sess.ActiveWindow.text == "Information":
        if confirm:
            _confirm()
        else:
            _decline()
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

def _select_open_items() -> bool:
    """
    Simulates pressing 'Choose open items' button
    located on the F-30 upper taskbar.
    """
    _main_wnd.SendVKey(_vkeys["F6"])
    return _is_alert_message()

def _choose_selection_method(name: str) -> None:
    """
    Chooses selection method from the F-30 side menu
    that will be applied for loading open items from
    account(s).
    """

    children = _main_wnd.FindByname(":SAPMF05A:0710", "GuiSimpleContainer").children

    for child in children:
        if name in child.Text:
            child.selected = True
            return

def _enter_values(vals: Collection) -> None:
    """
    Enters selection values based on which items will be loaded from
    a head office account. Selection will be performed based on the
    method name provided.
    """

    row_count = _main_wnd.FindByname(":SAPMF05A:0731", "GuiSimpleContainer").LoopRowCount

    for idx, val in enumerate(vals):

        _main_wnd.FindAllByname("RF05A-SEL01", "GuiTextField")[idx % row_count].text = val

        if idx % row_count == row_count - 1:
            _confirm()

def _set_head_office(val: str) -> None:
    """
    Enters head office account number into
    the corresponding field located on the
    'Select open items' submask.
    """
    _main_wnd.FindByname("RF05A-AGKON", "GuiCTextField").text = val

def _cancel_processing() -> None:
    """
    Exits transaction by double
    clicking the cancel button
    """
    _decline()
    _decline()

def _insert_document_numbers(hd_off: int, doc_nums: list) -> None:
    """
    Enters document numbers associated with a head office
    to the 'Enter selection criteria' fields.
    """

    _main_wnd.SendVKey(_vkeys["F6"])  # click 'choose open items

    while _is_warning_message():
        _confirm()

    _set_head_office(hd_off)
    _choose_selection_method(name = "Document Number")
    _main_wnd.SendVKey(_vkeys["ShiftF4"]) # click "Process open items" button
    msg = _stat_bar.Text

    if "is a branch of" in msg:
        _confirm()
    elif "is currently blocked by user" in msg:
        raise AccountBlockedError(msg)

    _enter_values(doc_nums)
    _main_wnd.SendVKey(_vkeys["ShiftF4"]) # Click Process Open Items

    # check SAP response in the status bar
    while _is_alert_message():

        msg = _stat_bar.Text

        if "No further open items were found" in msg:
            _main_wnd.SendVKey(_vkeys["ShiftF2"]) # doc overview
            if _select_open_items(): # choose open items
                correct_hd_off = msg[-7]
                _set_head_office(correct_hd_off)
                _main_wnd.SendVKey(_vkeys["ShiftF4"])
                _enter_values(doc_nums)
                _main_wnd.SendVKey(_vkeys["ShiftF4"]) # Click Process Open Items
            else:
                _decline()
                raise SapGeneralError(msg)

        if "No appropriate line item is contained in this document" in msg:
            _cancel_processing()
            if not _is_popup_dialog():
                _decline()
            _close_popup_dialog(confirm = True)
            raise DocumentNotFoundError(msg)

        _cancel_processing()
        _close_popup_dialog(confirm = True)
        raise SapGeneralError(msg)

    _main_wnd.SendVKey(_vkeys["ShiftF2"]) # Click Document Overview

def _get_guitable_control() -> CDispatch:
    """
    Returns a reference to the GuiTableControl object
    representing a table with items loaded from accounts.
    """
    return _main_wnd.FindByname("SAPDF05XTC_6102", "GuiTableControl")

def _get_unassigned_field() -> CDispatch:
    """
    Returns a reference to the GuiTextField object
    representing a field containing final balance
    of activated items amounts.
    """
    return _main_wnd.FindByname("RF05A-DIFFB", "GuiTextField")

def _select_layout(name: str) -> bool:
    """
    Selects display layout for loaded items
    from the list of available global layouts.
    """

    _main_wnd.findById("mbar/menu[3]/menu[0]").select()
    layouts = _sess.findById("wnd[1]").findAllByname("RF05A-XPOS1", "GuiRadioButton")

    for lay in layouts:
        if lay.text == name:
            lay.select()
            _confirm()
            return True

    return False

def _set_profit_center(head_off: int) -> None:
    """Sets item profit center value."""

    if _main_wnd.FindAllByname("COBL_XERGO", "GuiButton").Count == 0:
        return

    _main_wnd.FindByname("COBL_XERGO", "GuiButton").Press()
    wnd = _sess.FindById("wnd[1]")
    coll = wnd.FindAllByname("RKEAK-FIELD", "GuiCTextField")
    coll.ElementAt(0).text = str(head_off)
    _confirm()

def _set_amount(val: float) -> None:
    """Sets item amount value."""

    amnt = f"{abs(val):.2f}".replace(".", ",")
    _main_wnd.FindByname("BSEG-WRBTR", "GuiTextField").text = amnt

def _set_text(val: str) -> None:
    """Sets item text value."""

    if len(val) > 50:
        val = val.replace("D ", "D")

    _main_wnd.FindByname("BSEG-SGTXT", "GuiCTextField").text = val

def _set_tax_code(val: str) -> None:
    """Sets tax code value."""

    if _main_wnd.FindAllByName("BSEG-MWSKZ", "GuiCTextField").count != 0:
        _main_wnd.FindByname("BSEG-MWSKZ", "GuiCTextField").text = val

def _set_cost_center(val: str) -> None:
    """Sets cost center value."""

    if _main_wnd.FindAllByname("COBL-KOSTL", "GuiCTextField").count > 0:
        if _main_wnd.FindByname("COBL-KOSTL", "GuiCTextField").text == "":
            _main_wnd.FindByname("COBL-KOSTL", "GuiCTextField").text = val

def _set_gl_account(val: int) -> None:
    """Sets GL account value."""

    _main_wnd.FindByname("RF05A-NEWKO", "GuiCTextField").text = str(val)

def _set_posting_key(amnt: float) -> None:
    """Sets posting key value."""

    pst_key = "40" if amnt > 0 else "50"
    _main_wnd.FindByname("RF05A-NEWBS", "GuiCTextField").text = pst_key

def _set_assignment(val: int) -> None:
    """Sets assignment value."""

    if _main_wnd.FindAllByname("BSEG-ZUONR", "GuiTextField").Count > 0:
        _main_wnd.FindByname("BSEG-ZUONR", "GuiTextField").text = str(val)

def _next_item() -> None:
    """Moves mask to the next item."""

    _main_wnd.SendVKey(_vkeys["ShiftF4"])

def _check_calculate_tax() -> None:
    """Checks the 'Calculate Tax' option."""

    _main_wnd.findByname("BKPF-XMWST", "GuiCheckBox").selected = True

def _transfer_rests(rest_data: dict) -> None:
    """
    Transfers rests of matched items
    to the correcponding GL accounts.
    """

    tx_calc_checked = False

    for params in rest_data.values():

        if params["Rest_Amount"] == 0:
            continue

        _set_posting_key(params["Rest_Amount"])
        _set_gl_account(params["GL_Account"])
        _next_item()

        if _is_alert_message():
            if "999" in _stat_bar.Text:
                _confirm()

        if not tx_calc_checked:
            _check_calculate_tax()
            tx_calc_checked = True

        _set_amount(params["Rest_Amount"])
        _set_text(params["Posting_Text"])
        _set_tax_code(params["Tax_Code"])
        _set_cost_center(params["Cost_Center"])
        _set_assignment(params["Assignment"])
        _set_profit_center(params["Head_Office"])

def _get_field_indices(items: CDispatch) -> dict:
    """
    Returns a map of field names: field indices for
    the applied layout.
    """

    col_to_idx = {}

    for idx, col in enumerate(items.Columns):
        col_to_idx.update({col.name: idx})

    return col_to_idx

def _get_posting_period(pst_date: date) -> str:
    """
    Calulates posting period with respect
    to the document posting date.
    """

    curr_date = datetime.date(datetime.now())
    period = _main_wnd.findByname("BKPF-MONAT", "GuiTextField").text

    # adjust period if the current date in F-03
    # does not equal to the (assuming past) clearing date
    if curr_date != pst_date:
        period = int(period)
        if period == 1:
            period = 12
        else:
            period -= 1
        period = str(period)

    return period

def start(sess: CDispatch) -> None:
    """
    Starts F-30 transaction.

    Params:
    ------
    sess: A GuiSession object.

    Returns:
    --------
    None.
    """

    _logger.info("Starting F-30 ...")

    global _sess
    global _main_wnd
    global _stat_bar

    _sess = sess
    _main_wnd = _sess.findById("wnd[0]")
    _stat_bar = _main_wnd.findById("sbar")

    _sess.StartTransaction("F-30")

def close() -> None:
    """
    Closes a running F-30 transaction.

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
        F-30 when it's not running.
    """

    _logger.info("Closing F-30 ...")

    global _sess
    global _main_wnd
    global _stat_bar

    if _sess is None:
        _logger.warning("Attempt to close a connection that doesn't exist!")
        return

    _sess.EndTransaction()

    if _is_popup_dialog():
        _close_popup_dialog(confirm = True)

    _sess = None
    _main_wnd = None
    _stat_bar = None

def load_account_items(
        hq_docs: dict, cocd: str, curr: str, doc_date: date,
        pst_date: date, doc_type: str = "DA") -> CDispatch:
    """
    Loads open items from head office account(s) into F-30.

    Params:
    -------
    hq_docs:
        A map of head office accounts to document numbers which \n
        serve as uniqe identificators of open items on accounts.

    cocd:
        Ccompany code of the cleared open items.

    curr:
        Currency of the cleared open items.

    doc_date:
        Document date of the clearing document header.

    pst_date:
        Posting date of the clearing document header.

    doc_type:
        Type of the clearing document (default 'DA' - posting with clearing)

    Returns:
    --------
    A GuiTableControl object representing the list of loaded items.

    Raises:
    --------
    TransactionNotStartedError:
        When attempting to use the procedure
        before starting F-30.

    DocumentNotFoundError:
        When F-30 fails to find a specific
        document on a customer account.

    AccountBlockedError:
        When an acount is blocked by a user
        and cannot be accessed by the application.

    SapGeneralError:
        When F-30 displays an alert or warning
        for which no appropriata handler exists.
    """

    if _sess is None:
        raise TransactionNotStartedError(
            "Cannot load open items to F-30 when it's actually "
            "not running! Use the biaF30.start() procedure to "
            "run the transaction first of all.")

    period = _get_posting_period(pst_date)
    pst_date = pst_date.strftime("%d.%m.%Y")
    doc_date = doc_date.strftime("%d.%m.%Y")

    # set clearing document header data
    _main_wnd.findByname("BKPF-BLDAT", "GuiCTextField").text = doc_date
    _main_wnd.findByname("BKPF-BUDAT", "GuiCTextField").text = pst_date
    _main_wnd.findByname("BKPF-MONAT", "GuiTextField").text = period
    _main_wnd.findByname("BKPF-BUKRS", "GuiCTextField").text = cocd
    _main_wnd.findByname("BKPF-WAERS", "GuiCTextField").text = curr
    _main_wnd.findByname("BKPF-BLART", "GuiCTextField").text = doc_type

    # select transaction to be processed: transfer posting with clearing
    _main_wnd.FindAllByname("RF05A-XPOS1", "GuiRadioButton").ElementAt(3).Select()

    # load document numbers
    for hd_off, doc_nums in hq_docs.items():
        _insert_document_numbers(hd_off, doc_nums)

    _main_wnd.SendVKey(_vkeys["ShiftF4"])
    items = _get_guitable_control()

    return items

def select_and_transfer(
        items: CDispatch, transf: dict, n_cleared: int,
        cases: list, layout: str) -> CDispatch:
    """
    Selects relevant items from the list of loaded open items \n
    and transfers any amount rests to the appropriate GL account(s).

    Params:
    -------
    items:
        A GuiTableControl object representing the list of loaded items.

    transf:
        Rests of the open items amounts and their accounting params. \n
        These values are entered to 'Add Customer Item' submask fields \n
        durig their transfer to GL account(s).

    n_cleared:
        Number of open items that are expected to be cleared.

    cases:
        List of case ID numbers that identify cases in DMS.

    layout:
        Name of the layout defining columns for the loaded items.

    Returns:
    --------
    A GuiButton object that represents the 'Post' button \n
    located on the top toolbar of the F-30 transaction.

    Raises:
    -------
    ItemSelectionError:
        When selection of relevant open items from
        a list results in non-zero open balance.

    LayoutNotFoundError:
        When the supplied layout name is not found
        on the list of available data layouts.
    """

    loaded_item_count = items.VerticalScrollbar.Maximum + 1
    visible_row_count = items.VisibleRowCount

    # get layout feild indices
    fld_to_index = _get_field_indices(items)

    # check of the required field names are contained
    # and if not, select appropriate layout
    if "RFOPS_DK-SGTXT" not in fld_to_index:

        if not _select_layout(layout):
            raise LayoutNotFoundError(f"Layout '{layout}' not found "
            "in the list of global layouts! Reason: The name of the "
            "layout is incorrect or the layout has non-global scope!")

        items = _get_guitable_control()
        fld_to_index = _get_field_indices(items)

    doc_txt_idx = fld_to_index["RFOPS_DK-SGTXT"]
    doc_amnt_idx = fld_to_index["DF05B-PSBET"]

    # deactivate discount amounts for all the loaded items
    _main_wnd.FindByname("ICON_SELECT_ALL", "GuiButton").Press()
    _main_wnd.findByname("IC_Z-S", "GuiButton").press()

    # deactivate/activate items before selection
    # based on the cleared:uncleared ratio
    _main_wnd.FindByname("ICON_SELECT_ALL", "GuiButton").Press()

    if loaded_item_count - n_cleared > n_cleared:
        activated = False
        _main_wnd.findByname("IC_Z-", "GuiButton").press()
    else:
        _main_wnd.findByname("IC_Z+", "GuiButton").press()
        activated = True

    for row_idx in range(0, loaded_item_count):

        # get the GuiTableControl object and the
        # index of the current visible row
        items = _get_guitable_control()
        visible_row_idx = row_idx % visible_row_count

        # scroll down on large list to unhide rows
        if visible_row_idx == 0 and row_idx > 0:
            items.VerticalScrollbar.position = row_idx
            items = _get_guitable_control()

        # get document text value and return to the list table
        item_doc_txt = items.GetCell(visible_row_idx, doc_txt_idx).text

        # flag item indexes that met the criteria as processed
        # and select the item in the list for posting
        id_found = False

        for case_id in list(map(str, cases)):
            if case_id in item_doc_txt:
                id_found = True
                break

        if (id_found and not activated) or (not id_found and activated):
            items.GetCell(row_idx % visible_row_count, doc_amnt_idx).SetFocus()
            _main_wnd.SendVKey(_vkeys["F2"])

    # transfer rests if any
    fin_bal_fld = _get_unassigned_field()

    if _parse_amount(fin_bal_fld.text) != 0 and transf is not None:
        _main_wnd.SendVKey(_vkeys["ShiftF2"])
        _transfer_rests(transf)
        _main_wnd.SendVKey(_vkeys["ShiftF4"])
        # get a new reference to the final balance object since
        #  the prev one gets reset when switching content windows
        fin_bal_fld = _get_unassigned_field()

    # perform final balance check
    if _parse_amount(fin_bal_fld.text) != 0:
        _cancel_processing()
        _close_popup_dialog(True)
        raise ItemSelectionError(f"A difference found in the final balance: {fin_bal_fld.text}")

    post_btn = _main_wnd.findById("tbar[0]/btn[11]")

    return post_btn

def post_items(post_btn: CDispatch) -> int:
    """
    Posts the selected open items.

    Params:
    -------
    post_btn:
        A GuiButton object that represents the 'Post' button \n
        located on the top toolbar of F-30.

    Returns:
    --------
    The number of the posting document.

    Raises:
    -------
    ItemPostingError:
        When an error occured during posting
        of the selected open items.
    """

    # post selected items
    post_btn.press()

    if _is_warning_message():
        _confirm()

    # pause execution in order for SAP to apply changes in DB
    time.sleep(3)

    # better raise an exception to the caller in case of an error/warning
    if _is_error_message():
        _decline()
        _close_popup_dialog(True)
        raise ItemPostingError(_stat_bar.text)

    # get the posting number from the status bar text
    tokens = _stat_bar.text.split()
    pst_num = int(tokens[1])

    return pst_num
