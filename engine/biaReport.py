# pylint: disable = C0103, C0301, C0123, E0110, E1101, W0602, W0603, W0703, W1203

"""
The 'biaReport.py' module contains procedures
for generation of clearing output excel reports
that will be sent to users.

Version history:
    1.0.20210526 - Initial version.
    1.0.20220906 - Minor code style improvements.
"""

from datetime import datetime, date
from glob import glob
import logging
from os import mkdir
from os.path import exists, isfile, join, split
from shutil import move
import pandas as pd
from pandas import Index, DataFrame, Series, ExcelWriter

_logger = logging.getLogger("master")

_NO_ITEMS_TO_CLEAR = (1, 1)

_all_items_data_fields = [
    "Warnings",
    "Document_Number",
    "Document_Type",
    "DC_Amount",
    "Currency",
    "Tax",
    "Document_Date",
    "Due_Date",
    "Head_Office",
    "Branch",
    "Debitor",
    "Assignment_Acc",
    "Text",
    "ID",
    "ID_Match",
    "Amount_Match",
    "Tax_Match",
    "Status",
    "Status_Sales",
    "Status_AC",
    "Assignment_Disp",
    "Notification",
    "Category",
    "Category_Desc",
    "Root_Cause",
    "Autoclaims_Note",
    "Fax_Number",
    "Created_On",
    "Processor"
]

_all_items_data_fields_alt = [
    "Warnings",
    "Document_Number",
    "Document_Type",
    "DC_Amount",
    "Currency",
    "Tax",
    "Document_Date",
    "Due_Date",
    "Head_Office",
    "Branch",
    "Assignment_Acc",
    "Text",
    "ID",
    "ID_Match",
    "Amount_Match",
    "Tax_Match"
]

_cleared_fields = [
    "ID",
    "Rest_Amount",
    "Head_Office",
    "Currency",
    "Tax_Code",
    "GL_Account",
    "Posting_Text",
    "Posting_Number",
    "F30_Clearing_Status",
    "DMS_Closing_Status",
    "QM_Closing_Status"
]

def _get_col_width(data_fld_vals: Series, data_fld_name: str):
    """
    Returns excel column width calculated as
    the maximum count of characters contained
    in the column name and column data strings.
    """

    vals = data_fld_vals.astype("string").dropna().str.len()
    vals = list(vals)
    vals.append(len(str(data_fld_name)))

    return max(vals)

def _col_to_rng(data: DataFrame, first_col: str, last_col: str = None, row: int = -1, last_row: int = -1) -> str:
    """
    Generates excel data range notation (e.g. 'A1:D1', 'B2:G2'). \n
    If 'last_col' is None, then only single-column range will be \n
    generated (e.g. 'A:A', 'B1:B1'). if 'row' is '-1', then the generated \n
    range will span through all column(s) rows (e.g. 'A:A', 'E:E').

    Params:
    -------
    data:
        Data for which colum names should be converted to a range.

    first_col:
        Name of the first column.

    last_col:
        Name of the last column.

    row:
        index of the row for which the range will be generated.

    Returns:
    --------
    Excel data range notation.
    """

    if isinstance(first_col, str):
        first_col_idx = data.columns.get_loc(first_col)
    elif isinstance(first_col, int):
        first_col_idx = first_col
    else:
        assert False, "Argument 'first_col' has invalid type!"

    first_col_idx += 1
    prim_lett_idx = first_col_idx // 26
    sec_lett_idx = first_col_idx % 26

    lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
    lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
    lett = "".join([lett_a, lett_b])

    if last_col is None:
        last_lett = lett
    else:

        if isinstance(last_col, str):
            last_col_idx = data.columns.get_loc(last_col)
        elif isinstance(last_col, int):
            last_col_idx = last_col
        else:
            assert False, "Argument 'last_col' has invalid type!"

        last_col_idx += 1
        prim_lett_idx = last_col_idx // 26
        sec_lett_idx = last_col_idx % 26

        lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
        lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
        last_lett = "".join([lett_a, lett_b])

    if row == -1:
        rng = ":".join([lett, last_lett])
    elif first_col == last_col and row != -1 and last_row == -1:
        rng = f"{lett}{row}"
    elif first_col == last_col and row != -1 and last_row != -1:
        rng = ":".join([f"{lett}{row}", f"{lett}{last_row}"])
    elif first_col != last_col and row != -1 and last_row == -1:
        rng = ":".join([f"{lett}{row}", f"{last_lett}{row}"])
    elif first_col != last_col and row != -1 and last_row != -1:
        rng = ":".join([f"{lett}{row}", f"{last_lett}{last_row}"])
    else:
        assert False, "Undefined argument combination!"

    return rng

def _to_excel_serial(day: date) -> int:
    """
    Converts a datetime object into
    an excel-compatible date integer
    (serial) format.
    """

    delta = day - datetime(1899, 12, 30).date()
    days = delta.days

    return days

def _replace_col_char(col_names: Index, char: str, repl: str) -> Index:
    """
    Replaces a given character in data column names with a string.
    """

    col_names = col_names.str.replace(char, repl, regex = False)

    return col_names

def _write_all_items(wrtr: ExcelWriter, data: DataFrame, sht_name: str, head_idx: int, **formats):
    """
    Creates a report sheet containing all accounting data.
    """

    data.drop(columns = ["Virtual_ID"], inplace = True)

    # reorder fields - if there is at least one case ID in the 'ID' field
    # then use the field list for the consolidated dataset, otherwise for
    # FBL5N exported data only
    if data["ID"].notna().any():
        data = data.reindex(columns = _all_items_data_fields)
        data.loc[:, "Category"] = pd.to_numeric(data["Category"]).astype("UInt8")
    else:
        data = data.reindex(columns = _all_items_data_fields_alt)

    # prep data by converting to beter displayable data types
    data.loc[:, "Document_Number"] = pd.to_numeric(data["Document_Number"]).astype("UInt64")

    # delete False in Match fields where there were no IDs found
    qry = data.query("ID.isna()")
    data.loc[qry.index, "ID_Match"] = pd.NA
    data.loc[qry.index, "Amount_Match"] = pd.NA
    data.loc[qry.index, "Tax_Match"] = pd.NA

    date_fields = ("Document_Date", "Due_Date", "Created_On")

    for col in date_fields:

        if col not in data.columns:
            continue

        data.loc[:, col] = data[col].apply(
            lambda x: _to_excel_serial(x) if not pd.isna(x) else x
        )

    # format headers
    data.columns = _replace_col_char(data.columns, "_", " ")
    data.to_excel(wrtr, index = False, sheet_name = sht_name)
    data.columns = _replace_col_char(data.columns, " ", "_")

    data_sht = wrtr.sheets[sht_name]

    # apply formats to cleared items column data
    for idx, col in enumerate(data.columns):

        # select column data format
        if col == "DC_Amount":
            col_fmt = formats["monetary"]
        elif col == "Category":
            col_fmt = formats["categorical"]
        elif col in ("Document_Date", "Due_Date", "Created_On"):
            col_fmt = formats["date"]
        else:
            col_fmt = formats["general"]

        # calculate column width
        if col == "Warnings":
            col_width = _get_col_width(data[col], col) + 7
        else:
            col_width = _get_col_width(data[col], col) + 6

        # apply new column params
        data_sht.set_column(idx, idx, col_width, col_fmt)

    # apply formats to headers
    first_col, last_col = data.columns[0], data.columns[-1]

    data_sht.conditional_format(
        _col_to_rng(data, first_col, last_col, head_idx),
        {"type": "no_errors", "format": formats["header"]}
    )

    # freeze data header row and set autofiler on all fields
    data_sht.freeze_panes(head_idx, 0)
    data_sht.autofilter(_col_to_rng(data, first_col, last_col, row = head_idx))

def _write_cleared_items(wrtr: ExcelWriter, data: DataFrame, sht_name: str, head_idx: int, **formats):
    """
    Creates a report sheet containing cleared items only.
    """

    if data.shape == _NO_ITEMS_TO_CLEAR:
        data.to_excel(wrtr, index = False, header = False, sheet_name = sht_name)
        cleared_sht = wrtr.sheets[sht_name]
        width = _get_col_width(data["Evaluation_result"], "Evaluation_result") + 2
        cleared_sht.set_column(0, 0, width, formats["general"])
        return

    # reorder fields
    data = data.reindex(columns = _cleared_fields)

    # prep data by converting to beter displayable data types
    data["ID"] = pd.to_numeric(data["ID"]).astype("UInt64")

    # format headers
    data.columns = _replace_col_char(data.columns, "_", " ")
    data.to_excel(wrtr, index = False, sheet_name = sht_name)
    data.columns = _replace_col_char(data.columns, " ", "_")

    cleared_sht = wrtr.sheets[sht_name]

    # apply formats to all items column data
    for idx, col in enumerate(data.columns):

        col_width = _get_col_width(data[col], col) + 2

        if col == "Rest_Amount":
            col_fmt = formats["monetary"]
        else:
            col_fmt = formats["general"]

        cleared_sht.set_column(idx, idx, col_width, col_fmt)

    clr_first_col, clr_last_col = data.columns[0], data.columns[-1]

    cleared_sht.conditional_format(
        _col_to_rng(data, clr_first_col, clr_last_col, head_idx),
        {"type": "no_errors", "format": formats["header"]}
    )

    # freeze data header row and set autofiler on all fields
    cleared_sht.freeze_panes(head_idx, 0)

def create(evaluated: DataFrame, cleared: DataFrame, file_path: str, **sht_names):
    """
    Creates an excel report that contais evaluated and cleared accounting items.

    Params:
    -------
    evaluated:
        A dataset containig accounting items
        created by merging and evaluation of FBL5N and DMS data.

    cleared:
        A dataset containing items cleared in F-30.

    file_path:
        Path to the excel report file.

    sht_names:
        Names of particular report sheets.

    Returns:
    --------
    None.
    """

    HEADER_ROW_IDX = 1

    # print all and cleared items to separate sheets of a workbook
    with ExcelWriter(file_path, engine = "xlsxwriter") as wrtr:

        report = wrtr.book
        date_fmt = report.add_format({"num_format": "dd.mm.yyyy", "align": "center"})
        money_fmt = report.add_format({"num_format": "#,##0.00", "align": "center"})
        categ_fmt = report.add_format({"num_format": "000", "align": "center"})
        general_fmt = report.add_format({"align": "center"})
        header_fmt = report.add_format({
            "bg_color": "black", "font_color": "white", "bold": True, "align": "center"
        })

        _write_all_items(wrtr, evaluated, sht_names["all_items"], HEADER_ROW_IDX,
            header = header_fmt, monetary = money_fmt, categorical = categ_fmt,
            date = date_fmt, general = general_fmt
        )

        _write_cleared_items(wrtr, cleared, sht_names["cleared_items"], HEADER_ROW_IDX,
            header = header_fmt, monetary = money_fmt, general = general_fmt
        )

def upload(src_dir: str, dst_dir: str, subdir: str):
    """
    Uploads user excel reports to a network folder.

    A new subfolder is created in the destination folder \n
    if such doesn't exist, then the excel report is moved \n
    to the newly created subfolder.

    Params:
    -------
    src_dir:
        Path to a local folder containing the report file(s).

    dst_dir:
        Path to a network folder.

    subdir:
        Name of the network folder subdirectory.

    Returns:
    --------
    None.

    Raises:
    -------
    FileNotFoundError:
        When no excel files are found
        in the specified report folder.
    """

    dst_dir_path = join(dst_dir, subdir)
    report_paths = glob(join(src_dir, "*.xlsx"))

    if len(report_paths) == 0:
        raise FileNotFoundError(f"No .xlsx files were found in the specified folder: {src_dir}")

    if not exists(dst_dir_path):
        mkdir(dst_dir_path)

    for loc_rep in report_paths:

        rep_name = split(loc_rep)[1].replace(" ", "_") # replace spaces for a better readability
        net_rep = join(dst_dir_path, rep_name)

        if isfile(net_rep):
            _logger.warning(f"A file already exists and  will be overwritten: {net_rep}.")

        _logger.debug(f"Moving report: {loc_rep} -> {net_rep} ...")
        move(loc_rep, net_rep)

def append_summary(summ: str, evaluated: DataFrame, cleared: DataFrame, cocd: str, ent: str) -> str:
    """
    Appends a new row to an existing HTML table that summarizes
    selected paramaters of an account clearing report.

    Params:
    -------
    summ:
        An existing data summary to which a new summary row will be added.

    evaluated:
        Evaluated accounting items consisting of merged FBL5N and DMS data.

    cleared:
        A dataset containing items cleared in F-30.

    cocd:
        Company code of the entity for which the summary will be updated.

    ent:
        Entity for which the summary will be updated.

    Returns:
    --------
    Updated data summary.
    """

    f30_cleared_count = 0
    dms_closed_count = 0
    qm_closed_count = 0
    f30_err_count = 0
    qm_err_count = 0
    dms_err_count = 0
    skipped_count = 0
    total_errs = 0
    total_warns = 0

    cleared_cases = []

    if cleared.shape != _NO_ITEMS_TO_CLEAR:

        skipped_count = cleared[cleared["F30_Clearing_Status"].str.contains("skipped", False)].shape[0]

        f30_err_count = cleared[cleared["F30_Clearing_Status"].str.contains("error", False)].shape[0]
        qm_err_count = cleared[cleared["QM_Closing_Status"].str.contains("error", False)].shape[0]
        dms_err_count = cleared[cleared["DMS_Closing_Status"].str.contains("error", False)].shape[0]

        cleared_cases = cleared["ID"][cleared["F30_Clearing_Status"] == "Item cleared."]
        f30_cleared_count = evaluated[evaluated["ID"].isin(cleared_cases)].shape[0]
        dms_closed_count = cleared[cleared["DMS_Closing_Status"] == "Case closed."].shape[0]

        # one notification may be assigned to multiple case IDs
        id_notif_closed = cleared[cleared["QM_Closing_Status"] == "Notification closed."]["ID"]
        qm_closed_count = evaluated[evaluated["ID"].isin(id_notif_closed)]["Notification"].nunique()

        total_warns += cleared[cleared["F30_Clearing_Status"].str.contains("warning", False)].shape[0]
        total_warns += cleared[cleared["DMS_Closing_Status"].str.contains("warning", False)].shape[0]
        total_warns += cleared[cleared["QM_Closing_Status"].str.contains("warning", False)].shape[0]

        total_errs = f30_err_count + qm_err_count + dms_err_count

    total_itms_left = evaluated.shape[0] - f30_cleared_count
    total_warns += evaluated[evaluated["Warnings"].str.len() > 0].shape[0]

    curr_date = datetime.date(datetime.now())

    due_with_id = (
        ~pd.isnull(evaluated["ID"])) & \
        (evaluated["Due_Date"] <= curr_date) & \
        (~evaluated["ID"].isin(cleared_cases)
    )

    due_with_id_count = evaluated[due_with_id].shape[0]

    due_without_id = (
        pd.isnull(evaluated["ID"])) & \
        (evaluated["Due_Date"] <= curr_date) & \
        (evaluated["Document_Type"] != "DR") & \
        (~evaluated["ID"].isin(cleared_cases)
    )

    due_without_id_count = evaluated[due_without_id].shape[0]

    tbl_row = f"""
    <tr>
        <td style="border: purple 2px solid; padding: 5px">{ent}</td>
        <td style="border: purple 2px solid; padding: 5px">{cocd}</td>
        <td style="border: purple 2px solid; padding: 5px">{total_itms_left}</td>
        <td style="border: purple 2px solid; padding: 5px">{due_with_id_count}</td>
        <td style="border: purple 2px solid; padding: 5px">{due_without_id_count}</td>
        <td style="border: purple 2px solid; padding: 5px">{skipped_count}</td>
        <td style="border: purple 2px solid; padding: 5px">{f30_cleared_count}</td>
        <td style="border: purple 2px solid; padding: 5px">{dms_closed_count}</td>
        <td style="border: purple 2px solid; padding: 5px">{qm_closed_count}</td>
        <td style="border: purple 2px solid; padding: 5px">{total_warns}</td>
        <td style="border: purple 2px solid; padding: 5px">{total_errs}</td>
    </tr>
    """

    updated = "".join([summ, tbl_row])

    return updated
