# pylint: disable = C0103, C0123, W0603, W0703, W1203

"""
The 'biaQM02' module automates the standard SAP GUI QM02 transaction
in order to search, open and/or close service notifications and their
tasks.

Versioning
1.0.20220316 - initial version
1.0.20220504 - removed unused virtual key mapping fom _vkeys {}
"""

import logging
from typing import Union
from win32com.client import CDispatch

# custom warnings
class NotificationCompletionWarning(Warning):
    """
    Raised when a notification
    is already completed.
    """

# custom errors
class NotificationCompletionError(Exception):
    """
    Raised when a notification
    is cannot be completed.
    """

class NotificationSearchError(Exception):
    """
    Exception raised when no notification
    is found for a given notification ID.
    """

class TransactionNotStartedError(Exception):
    """
    Raised when attempting to use a procedure
    before starting the transaction.
    """

_sess = None
_main_wnd = None
_stat_bar = None

_logger = logging.getLogger("master")

_vkeys = {
    "Enter": 0,
    "CtrlS": 11,
    "F12":   12
}

def _is_alert_message() -> bool:
    """
    Checks whether status bar message \n
    is either an error or a warning.
    """

    ERROR = "E"
    WARNING = "W"

    if _stat_bar.MessageType in (ERROR, WARNING):
        return True

    return False

def _is_popup_dialog() -> bool:
    """
    Checks if the active window
    is a popup dialog window.
    """

    is_popup = (_sess.ActiveWindow.type == "GuiModalWindow")

    return is_popup

def _confirm():
    """
    Simulates pressing the 'Enter' button.
    """

    _main_wnd.SendVKey(_vkeys["Enter"])

def _decline():
    """
    Simulates pressing the 'F12' button.
    """

    _main_wnd.SendVKey(_vkeys["F12"])

def _close_popup_dialog(confirm: bool):
    """
    Confirms or delines a pop-up dialog.
    """

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

    return

def _get_task_viewer() -> CDispatch:
    """
    Returns a table object containing notification tasks.
    """

    tsk_vwr = _main_wnd.FindByName("SAPLIQS0MASSNAHMEN_VIEWER", "GuiTableControl")

    return tsk_vwr

def _set_notification_id(val: int):
    """
    Enters notification ID into the initial QM02 seach window.
    """

    _main_wnd.FindByName("RIWO00-QMNUM", "GuiCTextField").text = str(val)

def _activate_data_tab():
    """
    Activates the 'Data' tab located \n
    in the QM02 transaction.
    """

    _main_wnd.FindByName("TAB_GROUP_10", "GuiTabStrip").children[3].select()

def _select_task(row_idx: int):
    """
    Selects a task contained in QM02 task table \n
    based on the task row index.
    """

    # set scrollbar to tha appropriate position
    task_viewer = _get_task_viewer()
    task_viewer.VerticalScrollbar.position = row_idx

    # mark the task row as selected
    task_viewer = _get_task_viewer()
    task_viewer.getAbsoluteRow(row_idx).selected = True

def _append_task_text(num: int, val: str):
    """
    Appends a text value to an existing
    task description text.
    """

    _select_task(num)

    task_viewer = _get_task_viewer()
    task_viewer.VerticalScrollbar.position = num

    task_viewer = _get_task_viewer()
    task_viewer.GetCell(0, 4).text += val

def _complete_task(num: int):
    """
    Selects a task from task table based on the task \n
    row number, then completes the task by pressing \n
    the 'Complete' button.
    """

    _select_task(num)
    _main_wnd.findByName("FC_ERLEDIGT", "GuiButton").press()

    while _is_popup_dialog():
        _close_popup_dialog(confirm = True)

def _get_task_number(row_idx: int, task_viewer: CDispatch) -> str:
    """
    Returns the number of a task located at a given row index.
    """

    val = task_viewer.GetCell(row_idx, 0).text

    return val

def _get_task_text(row_idx: int, task_viewer: CDispatch) -> str:
    """
    Returns task description text.
    """

    val = task_viewer.GetCell(row_idx, 4).text

    return val

def _get_task_competion_date(row_idx: int, task_viewer: CDispatch) -> str:
    """
    Returns task completion date.
    """

    val = task_viewer.GetCell(row_idx, 16).text

    return val

def _get_active_tasks(task_vwr: CDispatch) -> dict:
    """
    Returns active tasks and their params
    such as row index and text.
    """

    row_idx = 0
    tasks = {}

    # do not use the VisibleRowCount property since task
    # count may be greater than those visible in the grid
    last_row_idx = task_vwr.RowCount - 1

    while row_idx < last_row_idx:

        visible_row_idx = row_idx % task_vwr.visibleRowCount

        # move down the scrollbar so that next positions appear in the table
        if visible_row_idx == 0 and row_idx > 0:
            task_vwr.VerticalScrollbar.position = row_idx
            task_vwr = _get_task_viewer()

        # program reaches the end of task list (the row contains no ID)
        if _get_task_number(visible_row_idx, task_vwr) == "":
            break

        # ignore completed tasks
        if _get_task_competion_date(visible_row_idx, task_vwr) != "":
            row_idx += 1
            continue

        text = _get_task_text(visible_row_idx, task_vwr)

        # store task parameters
        tasks[row_idx] = {
            "task_text": text,
            "row_idx": row_idx
        }

        row_idx += 1

    return tasks

def _complete_tasks(tasks: dict, cases: list, flag: str) -> tuple:
    """
    Identifies tasks related to given case IDs, \n
    appends a specific text to the task 'Text' value, \n
    and completes the tasks.
    """

    open_count = 0
    released_count = 0

    for task_num, task in tasks.items():

        # get the case ID contained in the CS task text if available
        if not task["task_text"].isnumeric():
            continue

        task_case_id = int(task["task_text"])

        if task_case_id not in cases:
            # leave it open
            open_count += 1
        else:
            # flag & release
            _append_task_text(task_num, flag)
            _complete_task(task_num)
            released_count += 1

    return (open_count, released_count)

def start(sess: CDispatch):
    """
    Starts QM02 transaction.

    Params:
    -------
    sess:
        A SAP GuiSession object.

    Returns:
    --------
    None.
    """

    _logger.info("Starting QM02 ...")

    global _sess
    global _main_wnd
    global _stat_bar

    _sess = sess
    _main_wnd = _sess.findById("wnd[0]")
    _stat_bar = _main_wnd.findById("sbar")

    _sess.StartTransaction("QM02")

def close():
    """
    Closes a running QM02 transaction.

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
        QM02 when it's not running.
    """

    _logger.info("Closing QM02 ...")

    global _sess
    global _main_wnd
    global _stat_bar

    if _sess is None:
        raise TransactionNotStartedError(
            "Cannot close QM02 when it's actually not running!"
            "Use the biaQM02.start() procedure to run the transaction first of all."
        )

    _sess.EndTransaction()

    if _is_popup_dialog():
        _close_popup_dialog(confirm = True)

    _sess = None
    _main_wnd = None
    _stat_bar = None

def search_notification(num: int) -> CDispatch:
    """
    Finds and opens a service notification in QM02.

    Params:
    -------
        num: Service notification ID number.

    Returns:
    --------
    A SAP GuiTableControl object representing
    the list of active notification tasks.

    Raises:
    -------
    TransactionNotStartedError:
        When attempting to use the
        procedure before starting QM02.

    NotificationSearchError:
        When no notification is found
        for a given notification ID number.

    NotificationCompletedWarning:
        When a notification is found
        but has already been completed.
    """

    if _sess is None:
        raise RuntimeError(
            "Cannot create a notification when QM02 has not started!"
            "Use the biaQM02.start() procedure to run the transaction first of all."
        )

    # enter the notification id into
    # the main search mask and confirm
    _set_notification_id(num)
    _confirm()

    if _is_popup_dialog():
        _confirm()

    msg = _stat_bar.text

    if _is_alert_message():
        _set_notification_id("") # clear entered value
        raise NotificationSearchError(msg)

    if "can only be displayed" in msg:
        _decline()
        _set_notification_id("") # clear entered value
        raise NotificationCompletionWarning(
            f"Could not open the notification. Reason: {msg}"
        )

    if _stat_bar.messagetype == "E":
        raise NotificationSearchError(msg)

    _activate_data_tab()
    task_viewer = _get_task_viewer()

    return task_viewer

def complete_notification(tasks: CDispatch, case: Union[int,list], flag: str = " Repaid"):
    """
    Completes an opened service notification.

    Params:
    -------
    tasks:
        A reference to task-containing table object.

    case:
        DMS ID number of case(s) that identify tasks \n
        to release before completing the notification.

    flag:
        Text value that will be appended to the tasks text \n
        before they are completed (default = ' Repaid').

    Returns:
    --------
    None.

    Raises:
    -------
    NotificationCompletionError:
        When a notification cannot be completed.
    """

    active_tasks = _get_active_tasks(tasks)

    # analyze params and process tasks accogringly
    if len(active_tasks) == 0:

        # a task cannot be just case ID or CS task if debit amount > thresh
        # then go back to initial mask and raise warning to the caller
        _decline()

        if _is_popup_dialog():
            _close_popup_dialog(confirm = False)

        raise NotificationCompletionError(
            "Unexpected number of tasks found in the notification!"
        )

    # at least 2 active tasks exist
    n_open = 0
    n_released = 0

    if isinstance(case, int):
        cases = [case]
    elif isinstance(case, list):
        cases = case
    else:
        raise TypeError(f"Argument 'cases' has incorrect type '{type(cases)}'!")

    n_open, n_released = _complete_tasks(active_tasks, cases, flag)

    # saves changes (if any were made)
    # and returns back to the initial window
    _main_wnd.SendVKey(_vkeys["CtrlS"])

    if _is_popup_dialog():
        _close_popup_dialog(confirm = True)

    if n_released == 0:
        # no task could be released
        raise NotificationCompletionError(
            "Could not release active task(s)! "
            "The task text contians no case ID, making the program "
            "unable to associate the task with a DMS case."
        )

    if n_open != 0:
        # notification contains still some active tasks
        raise NotificationCompletionError(
            "Cannot complete a notification "
            "since there are still some active tasks left!"
        )

    # press complete button
    _main_wnd.findById("tbar[1]/btn[20]").press()
