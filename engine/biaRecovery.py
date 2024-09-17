# pylint: disable = C0103, C0123, W0602, W0603, W0703, W1203

"""
The 'biaRecovery.py' module provides interface for setup
and management of the application recovery functionality.

Version history:
1.0.20222501 - initial version
1.0.20220906 - Minor code style improvements.
"""

import json
from os.path import exists
from typing import Any

_ent_states = None
_rec_path = None

def initialize(rec_path: str, entits: list) -> bool:
    """
    Initializes the application recovery functionality.

    A check for any previous application failure is performed. \n
    If a failure is detected, then the recovery file containing \n
    saved processing checkpoints will be loaded. If no failure \n
    is found, then a new recovery file with default entity states \n
    is created.

    Params:
    -------
    rec_path:
        Path to the file containing checkpoints for application recovery \n
        should the app crash or be terminated due to a critical error.

    entits:
        List of entities to which the recovery mechanism
        will be applied.

    Returns:
    --------
    True, if a previous application run ended with an error. \n
    False, if a previous application run ended without any \n
    error (return code 0).
    """

    global _ent_states
    global _rec_path

    _rec_path = rec_path

    if not exists(_rec_path):
        clear_entity_states()

    with open(_rec_path, 'r', encoding = "utf-8") as stream:
        ent_states = json.loads(stream.read())

    # use previous states to recover app
    if len(ent_states) != 0:
        _ent_states = ent_states
        return True

    # no failure - init app with default states
    _ent_states = reset_entity_states(entits)

    return False

def reset():
    """
    Releases all resources allocated \n
    by the recovery functionality and \n
    resets entity states to defaults.

    Params:
    -------
    None.

    Returns:
    --------
    None.
    """

    global _rec_path
    global _ent_states

    clear_entity_states()

    _rec_path = None
    _ent_states = None

def clear_entity_states():
    """
    Clears all application
    recovery data.

    Params:
    --------
    None.

    Returns:
    --------
    None.
    """

    global _ent_states

    _ent_states = {}

    with open(_rec_path, 'w', encoding = "UTF-8") as stream:
        json.dump(_ent_states, stream)

def reset_entity_states(entits: list) -> dict:
    """
    Sets default values to recovery data for active entities.

    Params:
    -------
    entits:
        A list of entity names.

    Returns:
    --------
    Default processing checkpoints per entity.
    """

    new_ent_states = {}

    for ent in entits:

        new_ent_states[ent] = {}
        new_ent_states[ent]["fbl5n_data_exported"] = False
        new_ent_states[ent]["fbl5n_data_converted"] = False
        new_ent_states[ent]["fbl5n_data_no_case"] = False
        new_ent_states[ent]["dms_data_exported"] = False
        new_ent_states[ent]["dms_data_converted"] = False
        new_ent_states[ent]["f30_input_generated"] = False
        new_ent_states[ent]["f30_items_cleared"] = False
        new_ent_states[ent]["dms_cases_processed"] = False
        new_ent_states[ent]["data_consolidated"] = False
        new_ent_states[ent]["data_analyzed"] = False
        new_ent_states[ent]["qm_notifications_processed"] = False

    with open(_rec_path, 'w', encoding = "utf-8") as stream:
        json.dump(new_ent_states, stream, indent = 4)

    return new_ent_states

def save_entity_state(ent: str, key: str, val: Any):
    """
    Stores a new recovery state for a given entity and parameter.

    Params:
    -------
    ent:
        Name of the entity (r.g. company code, worklist, ...)

    key:
        Parameter name for which the new value will be stored.

    val:
        A new value to store.

    Returns:
    --------
    None.
    """

    global _ent_states

    _ent_states[ent][key] = val

    with open(_rec_path, 'w', encoding = "utf-8") as stream:
        json.dump(_ent_states, stream, indent = 4)

def get_entity_state(ent: str, state: str) -> bool:
    """
    Returns a program runtime checkpoint
    state for a given entity.

    Params:
    -------
    ent:
        Name of the entity for which
        a recovery value will be searched.

    param:
        Name of the state for which
        a recovery value will be searched.

    Returns:
    --------
    True if a checkpoint was reached, othewise False.
    """

    state = _ent_states[ent][state]

    return state
