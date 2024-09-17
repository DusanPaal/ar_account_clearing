# pylint: disable = C0103, C0123, C0301, C0302, W0603, W0703, W1203

"""
The 'biaController.py' module represents the main communication
channel that manages data and control flow between the connected
highly specialized modules.

Version history:
1.0.20220325 - Initial version.
"""

import logging
import sys
from datetime import datetime as dt
from glob import glob
from logging import config
from os import remove
from os.path import exists, join
import re

import pandas as pd
import yaml
from pandas import DataFrame
from win32com.client import CDispatch

from . import biaDates as dat
from . import biaDMS as dms
from . import biaF30 as f30
from . import biaFBL5N as fbl5n
from . import biaMail as mail
from . import biaProcessor as proc
from . import biaQM02 as qm02
from . import biaRecovery as rec
from . import biaReport as report
from . import biaSAP as sap

_logger = logging.getLogger("master")

def _clean_dump(dir_path: str):
	"""
	Deletes all files contained
	in the application dump folder.
	"""

	file_paths = glob(join(dir_path, "*.*"))

	if len(file_paths) == 0:
		_logger.warning("Dump folder contains no files.")
		return

	_logger.info("Deleting dump files ...")

	for file_path in file_paths:
		try:
			remove(file_path)
		except Exception as exc:
			_logger.exception(exc)

def _exist_cases(fbl5n_data: DataFrame) -> bool:
	"""
	Checks if FBL5N data contains
	at least one case ID.
	"""

	cases = fbl5n_data["ID"][fbl5n_data["ID"].notna()].unique()

	if len(cases) == 0:
		return False

	return True

def clean_temp(dir_path: str):
	"""
	Deletes all files contained
	in the application temp folder.

	Params:
	-------
	dir_path:
		Path to the folder where
		temporary files are stored.

	Returns:
	--------
	None.
	"""

	file_paths = glob(join(dir_path, "**", "*.*"), recursive = True)

	if len(file_paths) == 0:
		_logger.warning("No temporary files found!")
		return

	_logger.info("Deleting temporaty data ...")

	for file_path in file_paths:
		try:
			remove(file_path)
		except Exception as exc:
			_logger.exception(exc)

def get_current_date(fmt: str) -> str:
	"""
	Returns a formatted current date.

	Params:
	-------
	fmt:
		The string that controls
		the ouptut date format.

	Returns:
	--------
	A string representing the current date.
	"""

	ctime = dat.get_current_date().strftime(fmt)

	return ctime

def configure_logger(cfg_path: str, log_path: str, header: dict, debug: bool = False) :
	"""
	Creates a new or clears an existing
	log file and prints the log header.

	Params:
	---------
	cfg_path:
		Path to a file with logging configuration params.

	log_path:
		Path to the application log file.

	header:
		Log header strings represented by parameter
		name (key) and description (value).

	debug:
		Indicates whether debug-level messages should be logged (default False).

	Returns:
	--------
	None.
	"""

	with open(cfg_path, 'r', encoding = "utf-8") as stream:
		content = stream.read()

	log_cfg = yaml.safe_load(content)
	config.dictConfig(log_cfg)

	if debug:
		_logger.setLevel(logging.DEBUG)
	else:
		_logger.setLevel(logging.INFO)

	prev_file_handler = _logger.handlers.pop(1)
	new_file_handler = logging.FileHandler(log_path)
	new_file_handler.setFormatter(prev_file_handler.formatter)
	_logger.addHandler(new_file_handler)

	# create a new / clear an existing log file
	with open(log_path, 'w', encoding = "utf-8"):
		pass

	# write log header
	for i, (key, val) in enumerate(header.items(), start = 1):
		line = f"{key}: {val}"
		if i == len(header):
			line += "\n"
		_logger.info(line)

def load_app_config(file_path: str) -> dict:
	"""
	Reads application configuration
	parameters from a file.

	Params:
	-------
	file_path:
		Path to a .yaml file containing
		the configuration params.

	Returns:
	--------
	Application configuration parameters.
	"""

	_logger.info("Configuring application ...")

	with open(file_path, 'r', encoding = "utf-8") as file:
		txt = file.read()

	txt = txt.replace("$appdir$", sys.path[0])
	cfg = yaml.safe_load(txt)

	return cfg

def load_clearing_rules(clear_cfg: dict) -> dict:
	"""
	Reads and parses a file that contains \n
	accounting parameters used for evalation and \n
	subsequent clearing of accounting open items.

	Params:
	-------
	clear_cfg:
		Application 'clearing' configuration parameters.

	Returns:
	--------
	A dict that maps countries (keys) to
	their respective clearing rules (values).
	"""

	_logger.info("Loading clearing rules ...")

	file_path = clear_cfg["rules_path"]

	with open(file_path, 'r', encoding = "utf-8") as file:
		rules = yaml.safe_load(file.read())

	return rules

def get_user_info(msg_cfg: dict, email_id: str) -> tuple:
	"""
	Returns the user's processing params and data.

	Params:
	-------
	msg_cfg:
		Application 'messages' configuration params.

	email_id:
		The string ID of the message.

	Returns:
	--------
	A tupe of the entity to process, and the user email address.
	"""

	_logger.info("Fetching user message ...")

	acc = mail.get_account(msg_cfg["mailbox"], msg_cfg["account"], msg_cfg["server"])
	msg = mail.get_message(acc, email_id)


	if msg is None:
		raise RuntimeError(f"Message not found using message ID: '{email_id}'!")

	_logger.info("Scanning message for a valid entity name ...")
	match = re.search(r"entity\s*:\s*(.*?)\r\n", msg.text_body, re.I)

	if match is None:
		raise RuntimeError("The message contains no valid entity name!")

	entity = match.group(1).upper()
	user_email = msg.sender.email_address

	_logger.info(f"User identified: '{user_email}'")
	_logger.info(f"Entity identified: '{entity}'")

	return (entity, user_email)

def get_active_entities(rules: dict, user_entity: str = None) -> tuple:
	"""
	Extracts active entites from clearing rules.

	Params:
		rules: Clearing rules for all countries.
		user_entity: Entity to clear requested by the user.

	Returns: A dict of active countries and their company codes.
	"""

	_logger.info("Searching for active entities ...")

	entits = {}

	for cocd, ccparams in rules.items():

		cntry = ccparams["country"]

		if not ccparams["active"]:
			_logger.warning(
				f"Country '{cntry}' is excluded from clearing "
				"according to the clearing rules.")
			continue

		for ent, enparams in ccparams["entities"].items():

			if user_entity is not None:

				if user_entity == ent:
					entits.update({ent: cocd})
					break
				else:
					continue

			if not enparams["active"]:
				_logger.warning(
					f"Entity '{ent}' is excluded from clearing "
					"according to the clearing rules.")
				continue

			entits.update({ent: cocd})

	ent_count = len(entits)

	_logger.info(f"Active entities found: {ent_count}")

	return entits

def connect_to_sap(sap_cfg: dict) -> CDispatch:
	"""
	Manages connecting of the application
	to the SAP GUI scripting engine.

	Params:
	-------
	sap_cfg:
		Application 'sap' configuration params.

	Returns:
	--------
	A SAP GuiSession object if a connection is created. \n
	If the attempt to connect fails due to an error, then None is returned.
	"""

	if sap_cfg["system"] == "P25":
		system = sap.SYS_P25
	elif sap_cfg["system"] == "Q25":
		system = sap.SYS_Q25

	_logger.info("Logging to SAP ... ")
	_logger.debug(f"System: '{system}'")

	try:
		sess = sap.login(sap_cfg["gui_path"], system)
	except Exception as exc:
		_logger.error(str(exc))
		return None

	return sess

def disconnect_from_sap(sess: CDispatch):
	"""
	Manages disconnecting from
	the SAP GUI scripting engine.

	Params:
	-------
	sess:
		A SAP GuiSession object.

	Returns:
	--------
	None.
	"""

	sap.logout(sess)

	return sess

def initialize_recovery(rec_cfg: dict, data_cfg: dict, entits: list):
	"""
	Manages initialization of the recovery funcionality.

	Params:
	-------
	rec_cfg:
		Application 'recovery' configuration parameters.

	data_cfg:
		Application 'data' configuration parameters.

	entits:
		List of entity names for which recovery will be initiaized.

	Returns:
	-------
	None.
	"""

	_logger.info("Initializing application recovery ...")

	is_prev_failure = rec.initialize(
		join(sys.path[0], rec_cfg["recovery_name"]), entits
	)

	# dumps may be deleted on start
	# if there was no prev app failure
	if not is_prev_failure:
		_clean_dump(data_cfg["dump_dir"])

def reset_recovery():
	"""
	Manages resetting
	the recovery funcionality.

	Params:
	-------
	None.

	Returns:
	--------
	None.
	"""

	_logger.info("Resetting recovery ...")

	rec.reset()

def export_fbl5n_data(data_cfg: dict, sap_cfg: dict, entits: dict, rules: dict, sess: CDispatch):
	"""
	Manages data export from customer accounts into a local file.

	Params:
	-------
	data_cfg:
		Application 'data' configuration parameters.

	sap_cfg:
		Application 'sap' configuration parameters.

	entits:
		List of entity names for wich data will be exported.

	rules:
		Clearing rules for all countries.

	sess:
		A SAP GuiSession object.

	Returns:
	--------
	None.
	"""

	success = False

	for ent, cocd in entits.items():

		group_type = rules[cocd]["entities"][ent]["type"]
		exp_name = data_cfg["fbl5n_data_export_name"].replace("$entity$", ent)
		exp_path = join(data_cfg["fbl5n_export_dir"], exp_name)

		if rec.get_entity_state(ent, "fbl5n_data_exported"):
			success = True
			_logger.warning(
				f"Skipping '{ent}' since the data "
				"was already exported in the previous run."
			)
			continue

		if group_type == "worklist":
			worklist = ent
		elif group_type == "company_code":
			worklist = None

		_logger.info(f" Exporting data for '{ent}' ...")

		try:
			fbl5n.export(sess, exp_path, cocd, worklist, sap_cfg["fbl5n_layout"])
		except fbl5n.NoDataFoundWarning as wng:
			_logger.warning(wng)
			continue
		except (fbl5n.AbapRuntimeError, fbl5n.WorklistNotFoundError) as exc:
			_logger.exception(exc)
		except fbl5n.ConnectionLostError as exc:
			_logger.error(str(exc))
			fbl5n.export(exp_path, cocd, worklist, sap_cfg["fbl5n_layout"])

		rec.save_entity_state(ent, "fbl5n_data_exported", True)
		success = True

	return success

def load_fbl5n_data(data_cfg: dict, rules: dict, entits: dict) -> bool:
	"""
	Manages conversion of data exported from FBL5N.

	Params:
	-------
	data_cfg:
		Application 'data' configuration parameters.

	sap_cfg:
		Application 'sap' configuration parameters.

	entits:
		List of entity names for wich data will be exported

	Returns:
	--------
	True if data conversion succeeds for at least
	one entity, False if it completely fails.
	"""

	_logger.info("Converting FBL5N data ...")

	success = False

	for ent, cocd in entits.items():

		exp_dir = data_cfg["fbl5n_export_dir"]
		bin_dir = data_cfg["dump_dir"]
		exp_file_name = data_cfg["fbl5n_data_export_name"].replace("$entity$", ent)
		bin_file_name = data_cfg["fbl5n_data_binary_name"].replace("$entity$", ent)

		exp_file_path = join(exp_dir, exp_file_name)
		bin_file_path = join(bin_dir, bin_file_name)

		if not rec.get_entity_state(ent, "fbl5n_data_exported"):
			_logger.warning(f"Skipping '{ent}' since "
			"there were no open items found on accounts.")
			proc.store_to_accum(ent, "fbl5n_data", None)
			continue

		if rec.get_entity_state(ent, "fbl5n_data_converted"):
			_logger.warning(f"Skipping '{ent}' since "
			"the data was already converted in the previous run.")
			conv = proc.read_binary(bin_file_path)
			success = True
		else:

			_logger.info(f" Converting data for '{ent}' ...")
			conv = proc.preprocess_fbl5n_data(exp_file_path, rules[cocd]["case_id_rx"])

		if conv is not None:
			proc.store_to_binary(conv, bin_file_path)
			rec.save_entity_state(ent, "fbl5n_data_converted", True)

			if not _exist_cases(conv):
				rec.save_entity_state(ent, "fbl5n_data_no_case", True)

		proc.store_to_accum(ent, "fbl5n_data", conv, force = True)
		success = True

	return success

def export_dms_data(data_cfg: dict, sap_cfg: str, sess: CDispatch, entits: dict) -> bool:
	"""
	Manages data export from DMS into a local file.

	Params:
	-------
	data_cfg:
		Application 'data' configuration parameters.

	sap_cfg:
		Application 'sap' configuration parameters.

	sess:
		A SAP GuiSession object.

	entits:
		List of entity names for wich data will be exported.

	Returns:
	--------
	True if data export succeeds for at least
	one entity, False if it completely fails.
	"""

	search_mask = dms.start(sess)

	for ent in entits:

		exp_dir = data_cfg["dms_export_dir"]
		exp_file_name = data_cfg["dms_data_export_name"].replace("$entity$", ent)
		exp_file_path = join(exp_dir, exp_file_name)

		if rec.get_entity_state(ent, "fbl5n_data_no_case"):
			_logger.warning(f"Skipping '{ent}' since FBL5N data contained no case ID.")
			continue

		if not rec.get_entity_state(ent, "fbl5n_data_exported"):
			_logger.warning(f"Skipping '{ent}' since no FBL5N data was exported.")
			continue

		fbl5n_data = proc.get_from_accum(ent, "fbl5n_data")
		case_nums = fbl5n_data["ID"][fbl5n_data["ID"].notna()].unique()

		if rec.get_entity_state(ent, "dms_data_exported"):
			_logger.warning(f"Skipping '{ent}' since "
			"the data was already exported in the previous run.")
			continue

		_logger.info(f" Exporting data for '{ent}' ...")

		try:
			grid_view = dms.search_disputes(search_mask, case_nums)
		except dms.NoCaseFoundError as exc:
			_logger.exception(exc)
			continue

		try:
			dms.export(grid_view, exp_file_path, sap_cfg["dms_layout"])
		except Exception as exc:
			_logger.error(str(exc))
			dms.close()
			return False

		rec.save_entity_state(ent, "dms_data_exported", True)

	dms.close()

	return True

def load_dms_data(data_cfg: dict, entits: dict) -> bool:
	"""
	Manages conversion of data exported from DMS.

	Params:
	-------
	data_cfg:
		Application 'data' configuration parameters.

	entits:
		List of entity names for which DMS data will be converted.

	Returns:
	--------
	True if data conversion succeeds,
	False if it completely fails.
	"""

	_logger.info("Converting DMS data ...")

	success = False

	for ent in entits:

		exp_dir = data_cfg["dms_export_dir"]
		bin_dir = data_cfg["dump_dir"]
		exp_file_name = data_cfg["dms_data_export_name"].replace("$entity$", ent)
		bin_file_name = data_cfg["dms_data_binary_name"].replace("$entity$", ent)

		exp_file_path = join(exp_dir, exp_file_name)
		bin_file_path = join(bin_dir, bin_file_name)

		if not rec.get_entity_state(ent, "dms_data_exported"):
			_logger.warning(f"Skipping '{ent}' since no DMS data was exported.")
			success = True # not a failure
			continue

		if not rec.get_entity_state(ent, "dms_data_converted"):
			_logger.info(f" Converting data for '{ent}' ...")
			converted = proc.preprocess_dms_data(exp_file_path)
			proc.store_to_binary(converted, bin_file_path)
			rec.save_entity_state(ent, "dms_data_converted", True)
		else:
			_logger.warning(f"Skipping '{ent}' since "
			"the data was already converted in the previous run.")
			converted = proc.read_binary(bin_file_path)

		proc.store_to_accum(ent, "dms_data", converted, force = True)
		success = converted is not None

	return success

def consolidate_data(data_cfg: dict, entits: dict, rules: dict) -> bool:
	"""
	Manages FBL5N and DMS data consolidation.

	Params:
	-------
	data_cfg:
		Application 'data' configuration parameters.

	entits:
		List of entity names for which cleatring input will be consolidated.

	rules:
		Clearing rules for all countries.

	Returns:
	--------
	True if data consolidation succeeds for at
	least one entity, False if it completely fails.
	"""

	_logger.info("Consolidating data ...")

	success = False

	for ent, cocd in entits.items():

		case_id_rx = rules[cocd]["case_id_rx"]
		valid_taxes = rules[cocd]["entities"][ent]["valid_taxes"]
		consolid_name = data_cfg["consolidated_data_name"].replace("$entity$", ent)
		consolid_path = join(data_cfg["dump_dir"], consolid_name)
		cust_data_name = data_cfg["customer_data_name"].replace("$comp_code$", cocd)
		cust_data_path = join(data_cfg["data_dir"], cust_data_name)
		cust_data = None

		if not rec.get_entity_state(ent, "dms_data_exported"):
			_logger.warning(f"Skipping '{ent}' since no DMS data was exported.")
			proc.store_to_accum(ent, "consolidated_data", None)
			success = True
			continue

		if rec.get_entity_state(ent, "data_consolidated"):
			_logger.warning(f"Skipping '{ent}' since "
			"the data was already consolidated in the previous run.")
			consolid = proc.read_binary(consolid_path)
			proc.store_to_accum(ent, "consolidated_data", consolid)
			success = True
			continue

		_logger.info(f" Consolidating data for '{ent}' ...")

		if exists(cust_data_path):
			_logger.info("Loading customer data ...")
			cust_data = pd.read_excel(cust_data_path)

		fbl5n_data = proc.get_from_accum(ent, "fbl5n_data")
		dms_data = proc.get_from_accum(ent, "dms_data")

		consolid = proc.consolidate(fbl5n_data,
			dms_data, cust_data, case_id_rx, valid_taxes
		)

		if consolid is None:
			# continue with next entity instead of returning
			# False if customer account is not found
			continue

		proc.store_to_binary(consolid, consolid_path)
		proc.store_to_accum(ent, "consolidated_data", consolid, force = True)
		rec.save_entity_state(ent, "data_consolidated", True)
		success = True

	return success

def generate_clearing_input(data_cfg: dict, entits: dict, rules: dict) -> bool:
	"""
	Manages generation of data input for open items clearing in F-30.

	Params:
	-------
	data_cfg:
		Application 'data' configuration parameters.

	entits:
		List of entity names for which cleatring input will be generated.

	rules:
		Clearing rules for all countries.

	Returns:
	True if any items to clear were matched,
	False if there were no items to clear detected.
	"""

	_logger.info("Generating clearig input ...")

	exist_matches = False

	for ent, cocd in entits.items():

		clr_input_path = join(
			data_cfg["dump_dir"],
			data_cfg["clearing_input_name"].replace("$entity$", ent)
		)

		matched_path = join(
			data_cfg["dump_dir"],
			data_cfg["matched_data_name"].replace("$entity$", ent)
		)

		analyzed_path = join(
			data_cfg["dump_dir"],
			data_cfg["analyzed_data_name"].replace("$entity$", ent)
		)

		# in case no cases were found (missing DMS access, no case id in acc text...)
		if not rec.get_entity_state(ent, "data_consolidated"):
			_logger.warning(
				f"Skipping '{ent}' since the data "
				"consolidation was not performed."
			)
			proc.store_to_accum(ent, "clearing_input", None)
			proc.store_to_accum(ent, "analyzed_data", None)
			proc.store_to_accum(ent, "matched_data", None)
			continue

		if rec.get_entity_state(ent, "data_analyzed"):
			_logger.warning(f"Data for '{ent}' was evaluated in the previous run.")
			analyzed = proc.read_binary(analyzed_path)
			matched = proc.read_binary(matched_path)
		else:
			consolid = proc.get_from_accum(ent, "consolidated_data")
			_logger.info(f"Detecting items to clear for '{ent}'...")
			analyzed = proc.evaluate_items(consolid,
				base_threshold = rules[cocd]["base_threshold"],
				tax_thresholds = rules[cocd]["tax_thresholds"]
			)

			matched = proc.get_matched_items(analyzed)
			_logger.info(f" Found {matched.shape[0]} items to clear.")

		proc.store_to_accum(ent, "analyzed_data", analyzed)
		proc.store_to_accum(ent, "matched_data", matched)
		proc.store_to_binary(analyzed, analyzed_path)
		proc.store_to_binary(matched, matched_path)
		rec.save_entity_state(ent, "data_analyzed", True)

		if matched.shape[0] == 0:
			proc.store_to_accum(ent, "clearing_input", None)
			proc.store_to_serial({}, clr_input_path)
			continue

		exist_matches = True

		if rec.get_entity_state(ent, "f30_input_generated"):
			_logger.warning(f"Clearing inut for '{ent}' already generated in the previous run.")
			clearing_input = proc.read_serial(clr_input_path)
			proc.store_to_accum(ent, "clearing_input", clearing_input)
			continue

		_logger.info(" Generating clearing input ...")
		cocd_rules = rules[cocd]
		ent_rules = rules[cocd]["entities"][ent]
		clr_input = proc.create_clearing_input(matched, cocd_rules, ent_rules)
		proc.store_to_serial(clr_input, clr_input_path)
		proc.store_to_accum(ent, "clearing_input", clr_input)
		rec.save_entity_state(ent, "f30_input_generated", True)

	return exist_matches

def clear_open_items(data_cfg: dict, sap_cfg: dict, clear_cfg: dict, entits: dict, sess: CDispatch) -> bool:
	"""
	Manages clearing of open items on customer accounts.

	Params:
	-------
	data_cfg:
		Application 'data' configuration parameters.

	sap_cfg:
		Application 'sap' configuration parameters.

	clear_cfg:
		Configuration parameters for open items clearing.

	entits:
		List of entity names for which clearing will be performed.

	sess:
		A SAP GuiSession object.

	Returns:
	--------
	True if clearing succeeds, False if it fails.
	"""

	f30.start(sess)

	success = False
	clr_date = dat.calculate_clearing_date(clear_cfg["holidays"])

	_logger.info(f"Clearing open items (clearing date: {clr_date.strftime('%d.%m.%Y')}) ... ")

	for ent, cocd in entits.items():

		clr_output_path = join(
			data_cfg["dump_dir"],
			data_cfg["clearing_output_name"].replace("$entity$", ent)
		)

		if proc.get_from_accum(ent, "clearing_input") is None:
			_logger.warning(f"Skipping '{ent}' since "
			"there were no items to clear found.")
			proc.store_to_accum(ent, "clearing_output", None)
			success = True
			continue

		if rec.get_entity_state(ent, "f30_items_cleared"):
			_logger.warning(f"Skipping '{ent}' since "
			"the items were already cleared in the previous run.")
			clr_output = proc.read_serial(clr_output_path)
			proc.store_to_accum(ent, "clearing_output", clr_output)
			success = True
			continue

		clr_input = proc.get_from_accum(ent, "clearing_input")
		clr_output = clr_input.copy() # make a data copy to preserve original vals
		posted = False

		for curr, params in clr_output.items():

			hd_off_docs = params["Head_Offs_To_Docs"]
			records = params["records"]
			clr_count = params["Matched_Count"]
			clr_records = {}
			case_nums = []

			# identify records that are not excluded from clearing
			for id_num, record in records.items():

				if not record["Skipped"]:
					clr_records[id_num] = record
					case_nums += record["Case_IDs"]
				else:
					reason = record["Message"]
					record["F30_Clearing_Status"] = f"WARNING: {reason}"
					_logger.warning(f"Skipping '{ent}' with ID '{id_num}'. Reason: {reason}")

			if len(clr_records) == 0:
				# all clearing records are to be skipped
				continue

			_logger.info(f"Clearing open items for '{ent}'; currency = {curr}")

			try:
				items = f30.load_account_items(hd_off_docs, cocd, curr, clr_date, clr_date)
			except Exception as exc:
				_logger.exception(exc)
				params["Cleared"] = False
				params["F30_Clearing_Status"] = "ERROR: Could not load items from account(s)"
				_logger.error(f"Loading failed. Reason: {exc}")
				continue

			pst_button = f30.select_and_transfer(items,
				clr_records, clr_count, case_nums, sap_cfg["f30_layout"]
			)

			if pst_button is None:
				params["Cleared"] = False
				params["F30_Clearing_Status"] = "ERROR: Item selection failed"
				continue

			try:
				pst_num = f30.post_items(pst_button)
			except f30.ItemPostingError as exc:
				params["Cleared"] = False
				params["F30_Clearing_Status"] = f"ERROR: {exc}."
				_logger.error(f"Posting failed. Reason: {exc}")
				continue

			_logger.debug(f"Posting number: {pst_num}")
			params["Cleared"] = True
			params["F30_Clearing_Status"] = "Item cleared."
			params["Posting_Number"] = pst_num
			posted = True

		rec.save_entity_state(ent, "f30_items_cleared", posted)
		proc.store_to_accum(ent, "clearing_output", clr_output)
		proc.store_to_serial(clr_output, clr_output_path)
		success |= posted

	f30.close()

	return success

def close_disputes(data_cfg: dict, entits: dict, sess: CDispatch):
	"""
	Manages closing of disputed cases.

	Params:
	-------
	data_cfg:
		Application 'data' configuration parameters.

	entits:
		List of entity names for which closing will be performed.

	sess:
		A SAP GuiSession object.

	Returns:
	--------
	None.
	"""

	search_mask = dms.start(sess)

	_logger.info("Closing DMS disputes ...")

	for ent in entits:

		dms_output_path = join(
			data_cfg["dump_dir"],
			data_cfg["dms_closing_output_name"].replace("$entity$", ent)
		)

		if rec.get_entity_state(ent, "fbl5n_data_no_case"):
			_logger.warning(f"Skipping '{ent}' since there were no items to clear found.")
			continue

		if rec.get_entity_state(ent, "dms_cases_processed"):
			_logger.warning(f"Skipping '{ent}' since the cases were already processed in the previous run.")
			dms_output = proc.read_serial(dms_output_path)
			proc.store_to_accum(ent, "dms_closing_output", dms_output)
			continue

		if not rec.get_entity_state(ent, "f30_items_cleared"):
			_logger.warning(f"Skipping '{ent}' since no items were cleared in F-30.")
			continue

		clearing_out = proc.get_from_accum(ent, "clearing_output")
		assert clearing_out is not None, "Error loading correct F-30 clearing input from the accumulator!"
		dms_output = clearing_out.copy()

		for curr, c_params in clearing_out.items():

			if not c_params["Cleared"]:
				# leaving this check here as some currency clearings may fail while others not
				_logger.warning(f"Skipping '{ent}'; currency: {curr} since no items were cleared in F-30.")
				continue

			_logger.info(f"Closing dispute(s) for '{ent}' ...")
			records = c_params["records"]
			pst_num = c_params["Posting_Number"]
			matched = proc.get_from_accum(ent, "matched_data")

			for id_num, record in records.items():

				if record["Skipped"]:
					records[id_num]["DMS_Closing_Status"] = "WARNING: Closing skipped due to the accouting exclusion criteria."
					_logger.warning(f"Skipping ID '{id_num}' as per settings defined in 'rules.yaml'.")
					continue

				for case in record["Case_IDs"]:

					_logger.info(f"Processing case: {case} ...")
					search_result = dms.search_dispute(search_mask, int(case))

					try:
						dms.modify_case_parameters(search_result,
							root_cause = record["Root_Cause"],
							status_ac = proc.generate_status_ac(matched, case, pst_num),
							status = dms.CaseStates.Closed
						)
					except dms.SavingChangesError as exc:
						_logger.exception(exc)
						usr_msg = "ERROR: Could not save changes. Check if Coordinator/Processor is correct."
					except dms.StatusAcError as exc:
						_logger.exception(exc)
						usr_msg = "ERROR: Status AC length limit exceeded!"
					except dms.CaseEditingError as exc:
						_logger.exception(exc)
						usr_msg = "ERROR: Could not edit the case!"
					except Exception as exc:
						_logger.error("Unhanded expection occured!", exc_info = exc)
						usr_msg = "ERROR: Could not close the case!"
						search_mask = dms.start(sess)
					else:
						usr_msg = "Case closed."

					record["DMS_Closing_Status"] = usr_msg

		rec.save_entity_state(ent, "dms_cases_processed", True)
		proc.store_to_accum(ent, "dms_closing_output", dms_output)
		proc.store_to_serial(dms_output, dms_output_path)

	dms.close()

def close_notifications(data_cfg: dict, entits: dict, sess: CDispatch) -> bool:
	"""
	Manages closing of service notifications.

	Params:
	-------

	data_cfg:
		Application 'data' configuration parameters.

	entits:
		List of entity names for which closing will be performed.

	sess:
		A SAP GuiSession object.

	Returns:
	--------
	True if closing succeeds, False if it fails.
	"""

	_logger.info("Closing QM02 notifications ...")

	success = True
	qm02.start(sess)

	for ent in entits:

		qm_output_path = join(
			data_cfg["dump_dir"],
			data_cfg["qm_closing_output_name"].replace("$entity$", ent)
		)

		if rec.get_entity_state(ent, "fbl5n_data_no_case"):
			_logger.warning(f"Closing skipped for '{ent}' "
			"since there were no case IDs found in FLB5N data.")
			continue

		if not rec.get_entity_state(ent, "f30_items_cleared"):
			_logger.warning(f"Closing skipped for '{ent}' "
			"since there were no items cleared.")
			continue

		if rec.get_entity_state(ent, "qm_notifications_processed"):
			_logger.warning(f"Closing skipped for '{ent}'. "
			"since cases were already processed in the previous run.")
			qm_output = proc.read_serial(qm_output_path)
			proc.store_to_accum(ent, "qm_closing_output", qm_output)
			success = True
			continue

		_logger.info(f"Closing notification(s) for '{ent}' ...")

		dms_closing_out = proc.get_from_accum(ent, "dms_closing_output")
		assert dms_closing_out is not None, "Error loading correct F-30 input from the accumulator!"
		qm_closing_output = dms_closing_out.copy()

		for curr, params in dms_closing_out.items():

			if not params["Cleared"]:
				# leaving this check here as some currency
				# clearings may fail while others not
				_logger.warning(f"Skipping '{ent}'; currency: {curr}. "
				"since no items were cleared in F-30.")
				continue

			records = params["records"]

			for id_num, record in records.items():

				if record["Skipped"]:
					msg = "WARNING: Closing skipped due to the accouting exclusion criteria."
					qm_closing_output[curr]["records"][id_num]["QM_Closing_Status"] = msg
					_logger.warning(f"Notification closing for ID: '{id_num}' skipped "
					"since excluded from clearing as per settings defined in 'rules.yaml'.")
					continue

				cases = record["Case_IDs"]
				notif_id = record["Notification"]

				if str(notif_id).startswith("301"):
					msg = "WARNING: Closing skipped due to invalid notification type for QM02."
					qm_closing_output[curr]["records"][id_num]["QM_Closing_Status"] = msg
					_logger.warning(f"Notification '{notif_id}' skipped "
					"for having an invalid notification type for QM02.")
					continue

				if record["Root_Cause"] == "L06":
					msg = "WARNING: Manual closing expected for credited L06 items."
					qm_closing_output[curr]["records"][id_num]["QM_Closing_Status"] = msg
					_logger.info(f"Closing of notification '{notif_id}' skipped since "
					"the case was cleared with a credit note, manual closing at CS assumed.")
					continue

				_logger.info(f" Searching notification '{notif_id}' ...")

				try:
					tasks = qm02.search_notification(notif_id)
				except qm02.TransactionNotStartedError as exc:
					_logger.critical(str(exc))
					return False
				except qm02.NotificationSearchError as exc:
					_logger.exception(exc)
					qm_closing_output[curr]["records"][id_num]["QM_Closing_Status"] = f"ERROR: {exc}."
					continue
				except qm02.NotificationCompletionWarning as wng:
					_logger.warning(wng)
					qm_closing_output[curr]["records"][id_num]["QM_Closing_Status"] = "WARNING: Notification already closed."
					continue

				_logger.info(" Completing notification ...")

				try:
					qm02.complete_notification(tasks, cases)
				except Exception as exc:
					_logger.error(str(exc))
					msg = "ERROR: Attempt to complete the notification failed."
					qm02.start(sess)
				else:
					msg = "Notification closed."

				params["records"][id_num]["QM_Closing_Status"] = msg

		rec.save_entity_state(ent, "qm_notifications_processed", True)
		proc.store_to_accum(ent, "qm_closing_output", qm_closing_output)
		proc.store_to_serial(qm_closing_output, qm_output_path)

	qm02.close()

	return success

def report_output(
		rep_cfg: dict, notif_cfg: dict, entits: dict,
		rules: dict, user_email: str) -> bool:
	"""
	Manages creation and uploading of reports to a network folder,
	as well as sending notification with clearing summary to users.

	Params:
	-------
	rep_cfg:
		Application 'reports' configuration parameters.

	notif_cfg:
		Application 'notifications' configuration parameters.

	entits:
		List of entity names for which reports will be generated.

	rules:
		Clearing rules for all countries.

	Returns:
	--------
	True if user notification is sent to the users,
	False if the sending fails.
	"""

	net_subdir_fmt = rep_cfg["net_subdir_format"]
	net_subdir = dt.now().date().strftime(net_subdir_fmt)
	user_report_dir = join(rep_cfg["net_dir"], net_subdir)
	summ = "" # start with an empty summary text
	recips = []

	for ent, cocd in entits.items():

		_logger.info(f"Generating user report and summary for '{ent}' ...")

		recips += [usr["mail"] for usr in rules[cocd]["entities"][ent]["accountants"]]
		rep_name = rep_cfg["name"].replace("$comp_code$", cocd).replace("$entity$", ent)
		qm_closing_output = None

		if rec.get_entity_state(ent, "qm_notifications_processed"):
			qm_closing_output = proc.get_from_accum(ent, "qm_closing_output")

		if not rec.get_entity_state(ent, "fbl5n_data_exported"):
			continue

		if not rec.get_entity_state(ent, "data_analyzed"):
			# for fbl5n data containing no case id in texts
			all_itms = proc.get_from_accum(ent, "fbl5n_data")
		else:
			all_itms = proc.get_from_accum(ent, "analyzed_data")

		clr_itms = proc.convert_processing_output(qm_closing_output)
		loc_rep_path = join(rep_cfg["local_dir"], rep_name)

		report.create(all_itms, clr_itms, loc_rep_path,
			all_items = "All Items", cleared_items = "Cleared"
		)

		summ = report.append_summary(summ, all_itms, clr_itms, cocd, ent)

	if user_email is None:
		try:
			_logger.info("Uploading reports ...")
			report.upload(rep_cfg["local_dir"], rep_cfg["net_dir"], net_subdir)
		except Exception as exc:
			_logger.exception(exc)
			# leave reports where they are

	if not notif_cfg["send"]:
		_logger.warning("Sending notification to users disabled in 'appconfig.yaml'.")
		return True

	_logger.info("Compiling notification ...")
	if user_email is None:
		template_path = notif_cfg["team_template_path"]
	else:
		template_path = notif_cfg["user_template_path"]

	with open(template_path, encoding = "utf-8") as stream:
		notif = stream.read()

	notif = notif.replace("$ReportPath$", user_report_dir).replace("$TblRows$", summ)

	with open(notif_cfg["html_email_path"], 'w', encoding = "utf-8") as stream:
		stream.write(notif)

	curr_date = dt.now().date().strftime("%d-%b-%Y")
	subj = notif_cfg["subject"].replace("$date$", curr_date)
	recips = list(set(recips)) if user_email is None else user_email

	if user_email is None:
		msg = mail.create_message(notif_cfg["sender"], recips, subj, notif)
	else:
		msg = mail.create_message(notif_cfg["sender"], recips, subj, notif, att = loc_rep_path)

	_logger.info("Sending user notification ...")

	try:
		mail.send_smtp_message(msg, notif_cfg["host"], notif_cfg["port"], debug = 1)
	except Exception as exc:
		_logger.exception(exc)
		return False

	_logger.info("Notification successfully sent to users.")

	return True
