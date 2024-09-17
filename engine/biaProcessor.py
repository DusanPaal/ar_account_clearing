# pylint: disable = C0103, C0123, C0301, C0302, W0603, W0703, W1203

"""
The 'biaProcessor' module performs all data-associated operations,such as
parsing, cleaning, conversion, evaluation and querying.

Version history:
1.0.20210803 - initial version
1.0.20220303 - fixed a regression in create_clearing_input() when attempt
               to assign a tax code to items with unused tax code was falsely
               evaluated as unsuccessful.
1.0.20220329 - Removed evaluation of business area for items being cleared
               in 'create_clearing_input()' procedure.
1.0.20220531 - Added item difference indicator as assignment value for France
               in 'create_clearing_input()' procedure.
1.0.20221013 - Updated docstrings.
1.0.20221013 - Added cleaning of user-entered reserved chars that would otherwise
               be interpreted as EOF by the pandas.read_csv() parser resulting in
               parsing error.
"""

from io import StringIO
import json
import logging
import re
from typing import Any, Union
import pandas as pd
from pandas import DataFrame, Series

_NOT_APPLICABLE = "NA"  # constant for 'not applcable' vals
_NULL_TAX = ""          # represents missing tax code

_logger = logging.getLogger("master")

_accum = {
    "fbl5n_data": {},
    "dms_data": {},
    "consolidated_data": {},
    "analyzed_data": {},
    "matched_data": {},
    "clearing_input": {},
    "clearing_output": {},
    "dms_closing_output": {},
    "qm_closing_output": {}
}

def store_to_accum(ent: str, key: str, data: Any = None, force: bool = False):
    """
    Stores application data to the global accumulator.

    For each entity, the data is stored under a specfic key serving \n
    as a descriptor of the data origin and role. By default, ovewriting \n
    of data stored in the accumulator is not allowed. The overwriting \n
    can be however enforced by setting the 'force' parameter to 'True'.

    Params:
    -------
    ent:
        Name of the entity (country, worklist, ...).

    key:
        Descriptor repesenting the name of the data.

    data:
        Data to store to the accumulator, eitehr a DataFrame object or None.

    force:
        Indicates whether data overwriting is allowed.

    Returns:
    --------
    None.
    """

    _logger.debug(f"Storing data to accumulator: entity = '{ent}'; key = '{key}'")

    if ent in _accum[key] and not force:
        assert False, "Cannot modify data already stored in the accumulator!"

    _accum[key].update({ent: data})

def get_from_accum(ent: str, key: str) -> DataFrame:
    """
    Returns application data from the global accumulator.

    For each entity, the data is stored under a specfic key \n
    serving as a descriptor of the data origin and role.

    Params:
    -------
    ent:
        Name of the entity (country, worklist, ...)

    key:
        Descriptor repesenting the name of the data.

    Returns:
    --------
    A DataFame object containing stored records, or None.
    """

    data = _accum[key][ent]

    return data

def store_to_serial(data: dict, file_path: str):
    """
    Stores a python dict object to a JSON file.

    Params:
    -------
    data:
        A dict object containing data to store.

    file_path:
        Path to the .json file to which the
        serialzed data will be stored.

    Returns:
    --------
    None.
    """

    with open(file_path, 'w', encoding = "utf-8") as stream:
        json.dump(data, stream, indent = 4)

def read_serial(file_path: str) -> dict:
    """
    Reads serialized data from a json file.

    Params:
    -------
    file_path:
        Path to the .json file containing
        serialized data.

    Returns:
    --------
    A dict object containing the loaded json data.
    """

    with open(file_path, 'r', encoding = "utf-8") as stream:
        data = json.loads(stream.read())

    if len(data) == 0:
        return None

    return data

def store_to_binary(data: DataFrame, file_path: str):
    """
    Stores a DataFrame object into a binary file.

    Params:
    -------
    data:
        A DataFrame object containing data records.

    file_path:
        Path to the binary file to which the data
        will be stored.

    Returns:
    --------
    None.
    """

    data.to_pickle(file_path)

def read_binary(file_path: str) -> DataFrame:
    """
    Reads the content of a binary \n
    file with stored panel dataset.

    Params:
    -------
    file_path:
        Path to the binary file \n
        from which data will be read.

    Returns:
    --------
    A DataFrame object containing data records.
    """

    data = pd.read_pickle(file_path)

    return data

def _read_text_file(file_path: str) -> str:
    """
    Reads the content of a text file.
    """

    with open(file_path, 'r', encoding = "utf-8") as stream:
        txt = stream.read()

    return txt

def _get_cases(source: Union[Series, str], case_rx: str) -> Union[Series, int]:
    """
    Extracts case numbers from string values.

    Params:
    -------
    source:
        A string or a Series of strings containing case ID numbers.

    case_rx:
        A Regex pattern matchng the numerical format
        of case numbers located in item description text.

    Returns:
    --------
    A case ID number or a Series of case ID numbers, both integers.
    """

    rx_patt = fr"(\A|[^a-zA-Z])D[P]?\s*[-_/]?\s*({case_rx})"

    if isinstance(source, Series):
        result = source.str.findall(rx_patt, re.I)
    elif isinstance(source, str):
        matches = re.findall(rx_patt, source)
        result = [int(m[1]) for m in matches]
    else:
        raise TypeError(f"Unsupported data type: {type(source)}")

    return result

def _parse_amount(val: str) -> float:
    """
    Converts string amount in SAP
    numeric format to a float number.
    """

    parsed = val.replace(".", "").replace(",", ".")

    if parsed.endswith("-"):
        parsed = "".join(["-" , parsed.replace("-", "")])

    return parsed

def _compact_fbl5n_data(text: str) -> str:
    """
    Extracts relevant accounting lines
    from the raw data strings exported
    from FBL5N.
    """

    matches = re.findall(
        pattern = r"^\|\s*\d+.*\|$",
        string = text,
        flags = re.M
    )

    replaced_a = re.sub(
        pattern = r"^\|",
        repl = "",
        string = "\n".join(matches),
        flags = re.M
    )

    del matches

    replaced_b = re.sub(
        pattern = r"\|$",
        repl = "",
        string = replaced_a,
        flags = re.M
    )

    del replaced_a

    # remove any reserved characters
    # possibly entered by users
    compacted = re.sub(
        pattern = "\"",
        repl = "",
        string = replaced_b,
        flags = re.M
    )

    del replaced_b

    return compacted

def _compact_dms_data(text: str) -> str:
    """
    Extracts relevant accounting lines
    from the raw data strings exported
    from DMS.
    """

    matches = re.findall(
        pattern = r"^\|.*?\|.*?\|\d+.*$",
        string = text,
        flags = re.M
    )

    replaced = re.sub(
        pattern = r"^\|",
        repl = "",
        string = "\n".join(matches),
        flags = re.M
    )

    del matches

    cleaned = re.sub(r"\|$", "", replaced, flags = re.M)
    del replaced

    cleaned = re.sub(
        pattern = "\"",
        repl = "",
        string = cleaned,
        flags = re.M
    )

    # remove any reserved characters
    # possibly entered by users
    compacted = re.sub(
        pattern = r"\|$",
        repl = "",
        string = cleaned,
        flags = re.M
    )

    return compacted

def _parse_fbl5n_data(data: str, case_rx: str) -> DataFrame:
    """
    Parses compacted FBL5N data strings.

    New fields are added to data:
        - ID
        - ID_Match
        - Amount_Match
        - Tax_Match
        - Warnings
        - Virtual_ID

    Case ID numbers are extracted form the \n
    'Text' field and placed into the 'ID' field.

    Params:
    -------
    data:
        Compacted FBL5N data strings.

    case_rx:
        A Regex pattern matchng the numerical format
        of case numbers located in item description text.

    Returns:
    --------
    A DataFrame object. The result of parsing.
    """

    parsed = pd.read_csv(StringIO(data),
        sep = "|",
        dtype = "string",
        names = [
            "Document_Number",
            "Assignment_Acc",
            "Document_Type",
            "Document_Date",
            "Due_Date",
            "DC_Amount",
            "Currency",
            "Tax",
            "Text",
            "Branch",
            "Head_Office"
        ]
    )

    # remove non-printable leading and trailing chars
    parsed = parsed.apply(lambda x: x.str.strip())

    # valid for Switzerland, Italy, maybe Austria
    parsed["Tax"].replace("**", "", inplace = True)

    # spare some RAM for our low mem server
    # by adding fields directly to an existing
    # DataFrame rather than using DataFrame.assign()
    # method to avoid creating a new data copy

    # indicates whether and item matches with some
    # others on ID (Case ID, Virtual ID)
    parsed["ID_Match"] = False

    # indicates whether the sum of the amounts
    # of ID-matched items is within the given threshold
    parsed["Amount_Match"] = False

    # indicates whether items that match on
    # IDs have the same tax symbol
    parsed["Tax_Match"] = False

    # contains information about inconsistencies
    # found in accounting data
    parsed["Warnings"] = ""

    # items with > 1 case ID get assigned a virtual ID
    # meaning that they belong together
    parsed["Virtual_ID"] = pd.NA

    # find and extract case IDs form items 'Text' values
    parsed["ID"] = _get_cases(parsed["Text"], case_rx)
    parsed["ID"].mask(parsed["ID"].str.len() == 0, pd.NA, inplace = True)
    parsed["ID"].mask(parsed["ID"].notna(), parsed["ID"].str[0], inplace = True)
    parsed["ID"].mask(parsed["ID"].notna(), parsed["ID"].str[1], inplace = True)

    return parsed

def _parse_dms_data(data: str) -> DataFrame:
    """
    Parses compacted DMS data strings.

    Params:
    -------
    data:
        Compacted DMS data strings.

    Returns:
    --------
    A DataFrame object. The result of parsing.
    """

    parsed = pd.read_csv(StringIO(data),
        sep = "|",
        dtype = "string",
        names = [
            "Debitor",
            "Case_ID",
            "Notification",
            "Status_Sales",
            "Assignment_Disp",
            "Status",
            "Created_On",
            "Status_AC",
            "Processor",
            "Category_Desc",
            "Root_Cause",
            "Autoclaims_Note",
            "Fax_Number",
            "Category"
        ]
    )

    # remove non-printable leading and trailing chars
    parsed = parsed.apply(lambda x: x.str.strip())

    return parsed

def _convert_fbl5n_data(data: DataFrame) -> DataFrame:
    """
    Converts fields of cleaned FBL5N data
    into their approprite data types.
    """

    conv = data.copy()

    conv["Head_Office"] = conv["Head_Office"].astype("object")
    numeric_mask = conv["Head_Office"].str.isnumeric()
    conv.loc[numeric_mask, "Head_Office"] = pd.to_numeric(conv.loc[numeric_mask, "Head_Office"])

    conv["DC_Amount"] = conv["DC_Amount"].apply(_parse_amount)
    conv["Branch"] = pd.to_numeric(conv["Branch"]).astype("UInt32")
    conv["DC_Amount"] = pd.to_numeric(conv["DC_Amount"])
    conv["Document_Number"] = pd.to_numeric(conv["Document_Number"]).astype("UInt64")
    conv["Document_Date"] = pd.to_datetime(conv["Document_Date"], dayfirst = True).dt.date
    conv["Due_Date"] = pd.to_datetime(conv["Due_Date"], dayfirst = True).dt.date
    conv["ID"] = pd.to_numeric(conv["ID"]).astype("UInt32")

    return conv

def _convert_dms_data(data: DataFrame) -> DataFrame:
    """
    Converts fields of cleaned DMS data
    into their approprite data types.
    """

    conv = data.copy()

    conv["Debitor"] = pd.to_numeric(conv["Debitor"]).astype("UInt32")
    conv["Notification"] = conv["Notification"].astype("UInt64")
    conv["Case_ID"] = conv["Case_ID"].astype("UInt32")
    conv["Created_On"] = pd.to_datetime(conv["Created_On"], dayfirst = True).dt.date
    conv["Root_Cause"] = conv["Root_Cause"].astype("category")
    conv["Category"] = conv["Category"].astype("category")
    conv["Status"] = conv["Status"].astype("UInt8")

    return conv

def _order_data(data: DataFrame, fields: list) -> DataFrame:
    """
    Orders data o DMS data based on specific fields.
    """

    ordered = data.sort_values(fields, ascending = False)

    return ordered

def preprocess_fbl5n_data(file_path: str, case_rx: str) -> DataFrame:
    """
    Converts plain FBL5N text data into a DataFrame object.

    Params:
    -------
    file_path:
        Path to the text file with FBL5N textual data to parse.

    case_rx:
        A Regex pattern matchng the numerical format
        of case numbers located in item description text.

    Returns:
    --------
    A DataFrame object representing the preprocessed
    data on success, None if preprocessing fails.
    """

    content = _read_text_file(file_path)
    compacted = _compact_fbl5n_data(content)
    parsed = _parse_fbl5n_data(compacted, case_rx)
    converted = _convert_fbl5n_data(parsed)

    return converted

def preprocess_dms_data(file_path: str) -> DataFrame:
    """
    Converts plain DMS text data to a panel dataset.

    Params:
    -------
    file_path: Path to the text file with DMS data to parse.

    Returns:
    --------
    A DataFrame object representing the preprocessed
    data on success, None if preprocessing fails.
    """

    content = _read_text_file(file_path)
    compacted = _compact_dms_data(content)
    parsed = _parse_dms_data(compacted)
    converted = _convert_dms_data(parsed)
    ordered = _order_data(converted, fields = ["Case_ID"])

    return ordered

def _virtualize(data: DataFrame, case_id_rx: str, base: int) -> DataFrame:
    """
    Generates virtual IDs in accounting data where
    multiple Case IDs are present in the item text.
    """

    # virtual ID generator
    def _init_generator(base: int) -> int:

        while True:
            yield base
            base += 1

    # find items with more than ID
    virt_id_generator = _init_generator(base)
    case_IDs = _get_cases(data["Text"], case_id_rx)
    multi_id_recs = data.index[case_IDs.str.len() > 1]

    if multi_id_recs.empty:
        return data

    virtual = data.copy()

    for idx in multi_id_recs:

        virt_id = next(virt_id_generator)
        virtual.loc[idx, "Virtual_ID"] = virt_id
        case_IDs = _get_cases(data["Text"][idx], case_id_rx)

        for case_id in case_IDs:
            virtual.loc[virtual["ID"] == case_id, "Virtual_ID"] = virt_id

    return virtual

def _detect_inconsistencies(data: DataFrame, valid_taxes: list) -> DataFrame:
    """
    Detects inconsistencies in data by validating critical accounting params.
    """

    checked = data.copy()

    checked.loc[
        checked.query("Debitor != Branch and ID.notna()").index, "Warnings"
    ] = "FBL5N and DMS debitors not equal!"

    checked.loc[
        checked.query("Debitor != Branch and ID.notna()").index, "Warnings"
    ] = "FBL5N and DMS debitors not equal!"

    checked.loc[
        checked.query(f"~Tax.isin({valid_taxes})").index, "Warnings"
    ] = "Unexpected tax code detected!"

    checked.loc[
        checked.query("Status == 4").index, "Warnings"
    ] = "Devaluated case ID assigned to an open item!"

    return checked

def consolidate(fbl5n: DataFrame, dms: DataFrame, cust: DataFrame,
                case_rx: str, valid_taxes: list) -> DataFrame:
    """
    Conslidates FBL5N and DMS data by:
    - merging of FBL5N and DMS datasets
    - generating of virtual ID numbers (if more than 1 case ID number/ item is detected)
    - checking the data for any incnsistencies with respect to accounting
    - ordering the data bysed on the 'ID' field

    Params:
    -------
    fbl5n:
        Preprocessed FBL5N data.

    dms:
        Preprocessed DMS data.

    cust:
        Nme of the customer to whom the data relate.

    case_rx:
        A Regex pattern matchng the numerical format
        of case numbers located in item description text.

    valid_taxes:
        A list of valid tax codes.

    Returns:
    --------
    A DataFrame object. The result of consolidation.
    """

    merged = pd.merge(fbl5n, dms, how = "left", left_on = "ID", right_on = "Case_ID")
    assert len(merged) > 0, "Data unmerged! Check the correctness of the megring key!"

    # add customer name and channel to data if applicable
    if cust is not None:

        merged = pd.merge(merged, cust, how = "left",
            left_on = "Head_Office", right_on = "Account"
        )

        merged.drop("Account", axis = 1, inplace = True)

        if merged["Customer_Name"].isna().any():
            return None

    virtualized = _virtualize(merged, case_rx, base = 10000000)

    if virtualized["Virtual_ID"].any():
        # apply virtual IDs through swapping Case ID and virtual ID fields
        virt_id_mask = virtualized["Virtual_ID"].notna()
        virtualized.loc[virt_id_mask, ["ID", "Virtual_ID"]] = (
            virtualized.loc[virt_id_mask, ["Virtual_ID", "ID"]].values
        )

    checked = _detect_inconsistencies(virtualized, valid_taxes)
    ordered = _order_data(checked, ["ID"])

    return ordered

def get_matched_items(data: DataFrame) -> DataFrame:
    """
    Extracts items where the below
    fields have 'True' value:
        - ID_Match
        - Tax_Match
        - Amount_Match

    These items represent the open items
    that will be cleared in F-30.

    Params:
    -------
    data:
        A DataFrame object representing
        the result of data evaluation.

    Returns:
    --------
    A subset of the data representing the matched items.
    """

    matched = data.query(
        "ID_Match == True and "
        "Tax_Match == True and "
        "Amount_Match == True"
    )

    return matched

def evaluate_items(consolid: DataFrame, **tax_rules) -> DataFrame:
    """
    Evaluates the consolidated FBL5N and DMS data
    using specific accounting criteria.

    Any items (data rows) match if and ony if:
    - their ID values are equal ('ID_Match' value will be set to 'True')
    - the sum of their amounts is within a defined threshold ('Amount_Match' value will be set to 'True')
    - their tax code values are compatible ('Amount_Match' value will be set to 'True')

    Params:
    -------
    consolid:
        Consolidated FBL5N and DMS data.

    tax_rules:
        Taxing params that apply to a specific country / kind of item: \n
        - base_threshold (float): The country-spciic tax base.
        - tax_thresholds (dict): Maps tax codes (str) to their threshold amounts (float).

    Returns:
    --------
    A DataFrame object. The result of evaluation.
    """

    INVALID_THRESHOLD = -1
    base_thresh = tax_rules["base_threshold"]
    tax_threshs = tax_rules["tax_thresholds"]

    if base_thresh == 0:
        base_thresh += 0.01

    # taxes that are allowed to pair with the null tax code
    allowed_tax_codes = [
        "YR", "YN", "TT", "TZ", "YO",
        "C3", "IG", "K6", "AU", "UU"
    ]

    data = consolid.copy()

    # get all records where duplicated IDs occur
    id_values = data.loc[data["ID"].notna(), "ID"]
    duplicated = id_values.duplicated(keep = False)

    # stop processing if there are no ID duplicates
    if not duplicated.any():
        return data

    # indicate match on 'ID' for the duplicated case ID / virtual ID records
    data.loc[duplicated.index[duplicated], "ID_Match"] = True

    # check if the IDs match on tax codes and amounts
    for id_num in id_values[duplicated].unique():

        id_recs = data[data["ID"] == id_num]

        # get list of used 'Tax' codes
        id_taxes = id_recs["Tax"].unique()
        tax_code = "" # init with an empty str
        thresh = INVALID_THRESHOLD

        if len(id_taxes) == 1:
            tax_code = id_taxes[0]
            data.loc[(data["ID"] == id_num), "Tax_Match"] = True
        elif len(id_taxes) == 2:
            # in case there is a standard and null tax code
            # check if the both are compatible, otherwise keep the initial tax value
            if _NULL_TAX in id_taxes:

                tax_code = id_taxes[id_taxes != _NULL_TAX][0]

                if tax_code in allowed_tax_codes:
                    data.loc[(data["ID"] == id_num), "Tax_Match"] = True

        if tax_code in tax_threshs:
            thresh = tax_threshs[tax_code]
        else:
            thresh = base_thresh

        assert thresh != INVALID_THRESHOLD, "Invalid threshold value!"

        # check if sum of credit amount(s) and debit amount(s) falls below threshold
        # excluding cases, where the sum of just credit or just debit amounts would fall
        # below threshold and thus falsely considered as match
        dc_amnts = id_recs["DC_Amount"]

        if abs(dc_amnts.sum()) < thresh and any(dc_amnts > 0) and any(dc_amnts < 0):
            data.loc[(data["ID"] == id_num), "Amount_Match"] = True

    return data

def _get_tax_code(curr: str, tax_codes: list, cocd_rules: dict,
                  hd_off, categ, ent_rules: dict) -> str:
    """
    Determnes item posting tax code based on the given params.
    """

    tax_code = "".join(tax_codes)

    if cocd_rules["diff_universal_tax_code"] != _NOT_APPLICABLE:
        tax_code = cocd_rules["diff_universal_tax_code"]

    if tax_code != _NULL_TAX:
        return tax_code

    if curr in cocd_rules["currency_taxes"]:
        tax_code = cocd_rules["currency_taxes"][curr]
    elif hd_off in ent_rules["head_office_taxes"]:
        tax_code = ent_rules["head_office_taxes"][hd_off]
    elif categ in cocd_rules["category_taxes"]:
        tax_code = cocd_rules["category_taxes"][categ]
    else:
        if cocd_rules["unused_tax_code"] == _NOT_APPLICABLE:
            tax_code = "" # empty str
        else:
            tax_code = cocd_rules["unused_tax_code"]

    return tax_code

def _get_gl_account_params(rest_amnt, ent_rules, categ) -> dict:
    """
    Returns GL account-specific params based on given params.
    """

    if rest_amnt == 0:
        gl_acc_type = None
    elif "penalties" in ent_rules["gl_accounts"] and categ in ("010", "011", "012"):
        gl_acc_type = ent_rules["gl_accounts"]["penalties"]
    elif "write_off_debits" in ent_rules["gl_accounts"] and rest_amnt > 0:
        gl_acc_type = ent_rules["gl_accounts"]["write_off_debits"]
    elif "write_off_credits" in ent_rules["gl_accounts"] and rest_amnt < 0:
        gl_acc_type = ent_rules["gl_accounts"]["write_off_credits"]
    else:
        gl_acc_type = ent_rules["gl_accounts"]["write_off_common"]

    return gl_acc_type

def _get_new_root_cause(prev_rt_cause: str, doc_types) -> str:
    """
    Returns new root cause code considering a previous
    root cause code, and a current document type.
    """

    if prev_rt_cause in ("L06", "L01"):
        new_root_cause = prev_rt_cause
    elif "DG" in doc_types:
        new_root_cause = "L06"
    elif "DZ" in doc_types or "DA" in doc_types:
        new_root_cause = "L01"

    return new_root_cause

def _get_posting_text(rest_amnt, cocd_rules, cust_name, case_IDs) -> str:
    """
    Compiles item posting text based on given params.
    """

    if rest_amnt == 0.0:
        pst_text = _NOT_APPLICABLE
    else:
        loc_diff_name = cocd_rules["local_diff_name"]
        loc_diff_name = loc_diff_name.replace("$customer$", cust_name)
        pst_text = loc_diff_name + " D " + " D ".join(list(map(str, case_IDs)))

    return pst_text

def _get_assignment(cocd_rules, id_num):
    """
    Returns item posting assignment based on given params.
    """

    if cocd_rules["country"] == "France":
        # item difference indicator, defined by credit mngr for France (Alexandra P.)
        assign = "2"
    else:
        assign = id_num

    return assign

def _update_head_office_list(uniq_hd_offs, hd_offs, hd_off_doc_nums) -> dict:
    """
    Adds new head offices to an existing list of unique head offices.
    """

    for ho in hd_offs:

        if ho not in uniq_hd_offs:
            uniq_hd_offs.update({ho: []})

        uniq_doc_nums = hd_off_doc_nums["Document_Number"][hd_off_doc_nums["Head_Office"] == ho]
        uniq_hd_offs[ho] += list(map(int, uniq_doc_nums.unique()))
        uniq_hd_offs[ho] = list(set(uniq_hd_offs[ho]))

    return uniq_hd_offs

def _update_case_id_list(uniq_case_IDs, case_IDs) -> list:
    """
    Adds case IDs to an existing list of unique case IDs.
    """

    updated = uniq_case_IDs + case_IDs

    return updated

def _update_output_currency(curr, output) -> dict:
    """
    Updates F-30 input on a new currency.
    """

    if curr not in output:
        output.update({curr: {}})
        output[curr].update({"records": {}})

    return output

def _get_unique_cases(curr_items) -> list:
    """
    Returns a list of unique case ID numbers.
    """

    nums = list(map(int, curr_items["ID"].unique()))

    return nums

def _get_match_count(curr_items) -> int:
    """
    Returns the number of matched items.
    """

    num = curr_items.shape[0]

    return num

def _get_customer_name(curr_items, id_mask) -> str:
    """
    Returns name of the customer.
    """

    cust_name = ""

    if "Customer_Name" in curr_items.columns:
        cust_name = curr_items[id_mask]["Customer_Name"].unique()

    return cust_name

def create_clearing_input(matched: DataFrame, cocd_rules: dict,
                          ent_rules: dict, cust_data: DataFrame = None) -> dict:
    """
    Transforms matched items into a specific data structure that contains \n
    accounting params that control the processing of open items in F-30.

    Params:
    -------
    matched:
        A DataFrame object that
        contains matched items.

    cocd_rules:
        Accounting rule valid for an entire company code.

    ent_rules:
        Accounting rule valid for a given  entity.

    cust_data:
        Customer-specific data map.

    Returns:
    --------
    Account clearig data input.
    """

    output = {}

    # create a lit of available document currencies
    currencies = matched["Currency"].unique()

    # process records separately for each currency found
    for curr in currencies:

        uniq_hd_offs = {}
        uniq_case_IDs = []

        curr_items = matched[matched["Currency"] == curr]
        match_count = _get_match_count(curr_items)
        output = _update_output_currency(curr, output)

        for id_num in _get_unique_cases(curr_items):

            id_mask = (curr_items["ID"] == id_num)
            cust_accs = curr_items[id_mask]["Branch"].unique()
            hd_offs = curr_items[id_mask]["Head_Office"].unique()
            hd_off_doc_nums = curr_items[id_mask][["Document_Number", "Head_Office"]]
            amounts = curr_items[id_mask]["DC_Amount"] # must not be unique!
            tax_codes = curr_items[id_mask]["Tax"].unique()
            doc_types = curr_items[id_mask]["Document_Type"].unique()
            virt_IDs = curr_items[id_mask]["Virtual_ID"].dropna().unique()
            categ = curr_items[id_mask]["Category"].dropna().unique()
            rt_cause = curr_items[id_mask]["Root_Cause"].dropna().unique()
            notif = curr_items[id_mask]["Notification"].dropna().unique()
            cust_name = _get_customer_name(curr_items, id_mask)

            # remap numpy/pandas integer types to native python
            # int as this is required by json serializers
            hd_offs = list(map(int, hd_offs))
            virt_IDs = list(map(int, virt_IDs))

            categ = categ[0]
            rt_cause = rt_cause[0]
            hd_off = hd_offs[0] # use always first available if multiple are present
            notif = int(notif[0])
            rest_amnt = float(amounts.sum().round(2))

            skipped = False
            msg = "" # init with empty str

            tax_code = _get_tax_code(curr, tax_codes, cocd_rules, hd_off, categ, ent_rules)

            # above replacement of null tax code won't succeed
            if tax_code == _NULL_TAX:
                skipped = True
                msg += "No tax code used! Program attemted to assign a valid tax code, but failed to find a suitable accounting rule."

            if len(virt_IDs) == 0: # positions which contain Virtual_IDs
                case_IDs = [id_num]
            else:
                case_IDs = virt_IDs

            # get new root cause or leave the prev one
            new_root_cause = _get_new_root_cause(rt_cause, doc_types)
            assert new_root_cause in ("L01", "L06"), "The new root cause is not applicable for DMS closing!"

            # get target GL account and cost center
            gl_acc_type = _get_gl_account_params(rest_amnt, ent_rules, categ)

            if gl_acc_type is None:
                gl_acc = _NOT_APPLICABLE
                cost_center = _NOT_APPLICABLE
            elif gl_acc_type["cost_center"]["trade"] == gl_acc_type["cost_center"]["retail"]:
                cost_center = set(gl_acc_type["cost_center"].values())
                cost_center = list(cost_center)[0]
                gl_acc = gl_acc_type["number"]
            else:
                assert cust_data is not None, f"Customer data is needed to categorize the account '{cust_acc}' as trade or retail!"
                gl_acc = gl_acc_type["number"]
                cust_acc = cust_accs[0]
                cust_info = cust_data.query(f"Account == {cust_acc}")
                assert not cust_info.empty, f"Could not find the account '{cust_acc}' in customer data!"

                if cust_info["Channel"] == "trade":
                    cost_center = gl_acc_type["cost_center"]["trade"]
                elif cust_info["Channel"] == "retail":
                    cost_center = gl_acc_type["cost_center"]["retail"]

            if tax_code in (ent_rules["skipped_taxes"] + cocd_rules["skipped_taxes"]):
                skipped = True
                msg += "Clearing skipped based on tax exclusion criteria defined in accounting rules."

            uniq_hd_offs = _update_head_office_list(uniq_hd_offs, hd_offs, hd_off_doc_nums)
            uniq_case_IDs = _update_case_id_list(uniq_case_IDs, case_IDs)
            pst_text = _get_posting_text(rest_amnt, cocd_rules, cust_name, case_IDs)
            assign = _get_assignment(cocd_rules, id_num)

            record = {
                "Skipped": skipped,
                "Message": msg,
                "Case_IDs": case_IDs,
                "Currency": curr,
                "Assignment": assign,
                "Head_Office": hd_off,
                "Tax_Code": tax_code,
                "Root_Cause": new_root_cause,
                "GL_Account": gl_acc,
                "Cost_Center": cost_center,
                "Posting_Text": pst_text,
                "Rest_Amount": rest_amnt,
                "Notification": notif,
                "DMS_Closing_Status": "",
                "QM_Closing_Status": ""
            }

            output[curr]["records"].update({id_num: record})

        # clearing is all or nothing, hence the F30_Clearing_Status
        # param used per currency rathe than per case
        output[curr]["F30_Clearing_Status"] = ""
        output[curr]["Head_Offs_To_Docs"] = uniq_hd_offs
        output[curr]["Case_IDs"] = uniq_case_IDs
        output[curr]["Posting_Number"] = None
        output[curr]["Cleared"] = False
        output[curr]["Matched_Count"] = match_count

    return output

def convert_processing_output(data: dict) -> DataFrame:
    """
    Converts dict data resulting from items \n
    clearing into a DataFrame object.

    Params:
    ------
    data:
        Output of open items clearing in F-30.

    Returns:
    --------
    A DataFrame object containing
    cleared items, their accounting
    params and a message for the user.
    """

    if data is None:
        output = DataFrame({
            "Evaluation_result": "No items to clear found."
        }, index = [0])

        return output

    data_frames = []

    for curr, params in data.items():

        records = params["records"]
        pst_num = params["Posting_Number"]

        for id_num in records:
            row = DataFrame.from_dict(records[id_num])
            row["Currency"] = curr
            row["ID"] = id_num
            row = row.drop_duplicates(subset=["ID"], ignore_index=True)
            row["F30_Clearing_Status"] = params["F30_Clearing_Status"]
            data_frames.append(row)

    output = pd.concat(data_frames)
    output["Posting_Number"] = pst_num if pst_num is not None else pd.NA
    output["ID"] = pd.to_numeric(output["ID"]).astype("UInt64")

    return output

def generate_status_ac(data: DataFrame, case: int, pst_num: int) -> str:
    """
    Generates a new text for 'Status AC' field in DMS from the used params.

    Params:
    -------
    data:
        A DataFrame object that contains matched items.

    case:
        ID number of the case stored in DMS.

    pst_num:
        Number of a clearing document generated \n
        by F-30 when posting the open items.

    Returns:
    --------
    A new 'Status AC' text.
    """

    # search the Case ID in 'ID' field
    states_ac = data["Status_AC"][data["ID"] == case]

    # if nothing was found, then search the Case ID in 'Virtual_ID' field
    if states_ac.empty:
        states_ac = data["Status_AC"][data["Virtual_ID"] == case]

    max_chars = 50 # DMS text length limit for this field
    status_ac = str(states_ac.unique()[0]).strip()
    new_status_ac = " ".join([status_ac, str(pst_num)])
    new_status_ac = new_status_ac.strip()

    if len(new_status_ac) > max_chars:
        _logger.error(f"The new status AC exceeds the limit {max_chars} chars!")
        _logger.warning(f"The original status AC value will be retained.")
        return status_ac

    return new_status_ac
