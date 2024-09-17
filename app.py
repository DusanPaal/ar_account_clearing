# pylint: disable = C0103, W0703, W1203

"""
The 'app.py' module contains the main entry of the program
that controls application data and operation flow via
the high level wrapper-like biaController module.
Application event logger is initiated upon module import.

Version history:
----------------
1.2.20210730 - Initial version.
1.2.20220315 - Logging initialization moved to biaController.init_logger() procedure.
1.2.20220906 - Minor bugfixes and code style improvements across all modules.
1.2.20221014 - Code refactored across all modules.
			 - Updated / added docstrings for most of procedures.
1.3.20230830 - Added clearing of open items only for a user-specified entity.
"""

import argparse
import logging
import sys
from os.path import join
import engine.biaCore as core

_logger = logging.getLogger("master")

def main(args: dict) -> int:
	"""
	Program entry point.
	Controls the overall program execution.

	Returns:
	--------
	Program completion state.
	"""

	log_cfg_path = join(sys.path[0], "log_config.yaml")
	log_file_path = join(sys.path[0], "log.log")
	log_header = {
		"Application name": "AR Account Clearing",
		"Application version": "1.3.20230830",
		"Log date": core.get_current_date("%d-%b-%Y")
	}

	_logger.info("=== Initialization ===")

	try:
		core.configure_logger(log_cfg_path, log_file_path, log_header, debug = False)
		cfg = core.load_app_config(join(sys.path[0], "app_config.yaml"))
		rules = core.load_clearing_rules(cfg["clearing"])
	except Exception as exc:
		_logger.critical(str(exc))
		return 1

	if args["email_id"] is None:
		entits = core.get_active_entities(rules)
		user_email = None
	else:
		user_entity, user_email = core.get_user_info(cfg["mails"]["requests"], args["email_id"])
		entits = core.get_active_entities(rules, user_entity)

	if len(entits) == 0:
		_logger.warning("No entity to process detected! Applicaiton will quit.")
		return 0 # not considered an error, exit with success

	try:
		core.initialize_recovery(cfg["recovery"], cfg["data"], entits.keys())
		sess = core.connect_to_sap(cfg["sap"])
	except Exception as exc:
		_logger.critical(str(exc))
		return 2

	_logger.info("=== Initialization OK ===\n")

	try:
		_logger.info("=== Processing ===")
		core.export_fbl5n_data(cfg["data"], cfg["sap"], entits, rules, sess)
	except Exception as exc:
		_logger.exception(exc)
		_logger.critical("Data export from FBL5N failed!")
		_logger.info("=== Cleanup ===")
		core.disconnect_from_sap(sess)
		_logger.info("=== Cleanup OK ===\n")
		return 4

	if not core.load_fbl5n_data(cfg["data"], rules, entits):
		_logger.critical("Prerpocessing of data exported from FBL5N failed.")
		_logger.info("=== Cleanup ===")
		core.disconnect_from_sap(sess)
		return 5

	if not core.export_dms_data(cfg["data"], cfg["sap"], sess, entits):
		_logger.critical("Data export from DMS failed.")
		_logger.info("=== Cleanup ===")
		core.disconnect_from_sap(sess)
		return 6

	if not core.load_dms_data(cfg["data"], entits):
		_logger.critical("Prerpocessing of data exported from DMS failed.")
		_logger.info("=== Cleanup ===")
		core.disconnect_from_sap(sess)
		return 7

	if not core.consolidate_data(cfg["data"], entits, rules):
		_logger.critical("Data consolidation failed.")
		_logger.info("=== Cleanup ===")
		core.disconnect_from_sap(sess)
		return 8

	if core.generate_clearing_input(cfg["data"], entits, rules):

		if not core.clear_open_items(cfg["data"], cfg["sap"], cfg["clearing"], entits, sess):
			_logger.critical("Clearing of open items in F-30 failed.")
			_logger.info("=== Cleanup ===")
			core.disconnect_from_sap(sess)
			return 9

		core.close_disputes(cfg["data"], entits, sess)

		if not core.close_notifications(cfg["data"], entits, sess):
			_logger.critical("Closing of notification(s) in QM02 failed.")
			_logger.info("=== Cleanup ===")
			core.disconnect_from_sap(sess)
			return 10

	_logger.info("=== Reporting ===")
	if not core.report_output(cfg["reports"], cfg["mails"]["notifications"], entits, rules, user_email):
		_logger.info("=== Cleanup ===")
		core.disconnect_from_sap(sess)
		return 11

	_logger.info("=== Success ===\n")

	_logger.info("=== Cleanup ===")
	core.disconnect_from_sap(sess)
	core.clean_temp(cfg["data"]["temp_dir"])
	core.reset_recovery()
	_logger.info("=== Success ===\n")

	return 0

if __name__ == "__main__":

	parser = argparse.ArgumentParser()

	parser.add_argument(
		"-e", "--email_id",
		required = False,
		help = "Sender email id."
	)

	exit_code = main(vars(parser.parse_args()))
	_logger.info(f"=== System shutdown with return code: {exit_code} ===")
	logging.shutdown()
	sys.exit(exit_code)
