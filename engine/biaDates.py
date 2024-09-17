# pylint: disable = C0103

"""
the 'biaDates.py' module provides procedures
that that perform date-related calculations
for the application.

Version history:
    1.0.20221005 - Initial version.
"""

from datetime import date, datetime, timedelta
import numpy as np

def _end_of_month(day: date) -> date:
    """
    Calculates last day of the month for a given day.
    """

    next_mon = day.replace(day=28) + timedelta(days=4)
    first_day_next_mon = next_mon - timedelta(days=next_mon.day)

    return first_day_next_mon

def _start_of_month(day: date) -> date:
    """
    Calculates first day of the month for a given day.
    """

    first_day = day.replace(day = 1)

    return first_day

def _get_month_ultimo(day: date, off_days: list) -> date:
    """
    Calculates ultimo date for the month of a given day.
    """

    ultimo = _end_of_month(day)

    while not np.is_busday(ultimo, holidays = off_days):
        ultimo -= timedelta(1)

    return ultimo

def _get_month_uplusone(day: date, off_days: list) -> date:
    """
    Calculates ultimo plus one date for the month of a given day.
    """

    upone = _start_of_month(day)

    while not np.is_busday(upone, holidays = off_days):
        upone += timedelta(1)

    return upone

def _get_prev_ultimo(uplusone: date, off_days: list) -> date:
    """
    Calculates ultimo date corresponding to a given
    ultimo plus one day.
    """

    ultimo = uplusone - timedelta(1)

    while not np.is_busday(ultimo, holidays = off_days):
        ultimo -= timedelta(1)

    return ultimo

def _get_actual_off_days(day: date, off_days: list) -> list:
    """
    Returns a list of company's calculated out of office days.
    """

    actual = []

    for item in off_days:
        curr_day = date(day.year, item.month, item.day)
        actual.append(curr_day)

    return actual

def get_current_date() -> date:
    """
    Returns a current date.

    Params:
    -------
    None.

    Returns:
    --------
    A datetime.date object
    representing a current date.
    """

    curr_date = datetime.now().date()

    return curr_date

def calculate_clearing_date(off_days: list) -> date:
    """
    Returns a calculated clearing date for items to post.

    Params:
    -------
    off_days:
        A list of datetime.date objects that represent \n
        out of office dates according to the company's \n
        fiscal year calendar.

    Returns:
    ---------
    Calculated clearing date.
    """

    day = get_current_date()
    actual_off_days = _get_actual_off_days(day, off_days)
    uplusone = _get_month_uplusone(day, actual_off_days)
    ultimo = _get_month_ultimo(day, actual_off_days)

    if ultimo < day:
        clr_date = ultimo
    elif day <= uplusone:
        clr_date = _get_prev_ultimo(uplusone, actual_off_days)
    else:
        clr_date = day

    return clr_date
