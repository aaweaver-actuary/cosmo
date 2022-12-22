"""
is_xlsb.py
"""

import re
import openpyxl
import pyxlsb

def is_xlsb(wb):
    """
    Description
    -----------
    This function takes a workbook object as input and returns True if the
    workbook file extension ends with ".xlsb" and False if not.

    Parameters
    ----------
    wb : openpyxl.Workbook or pyxlsb.Workbook
        Workbook object to be checked.

    Returns
    -------
    bool
        True if the workbook file extension ends with ".xlsb" and False if not.

    Raises
    ------
    ValueError
        If the wb is not a workbook object.

    Imports
    -------
    openpyxl
    pyxlsb
    re

    Examples
    --------
    >>> is_xlsb(wb)
    True

    >>> is_xlsb(245)
    False
    """
    # test if the object passed is a workbook object,
    # and if not, raise a value error
    if (
        not isinstance(wb, openpyxl.Workbook) and
        not isinstance(wb, pyxlsb.Workbook)
        ):
        raise ValueError("wb is not a workbook object")
    # otherwise, test if the workbook is an xlsb file by
    # checking the file extension
    # if the workbook is an xlsb file, return True
    # if the workbook is not an xlsb file, return False
    else:
        # test if the workbook is an xlsb file by checking the file extension
        # using a regular expression
        return re.search(r"\.xlsb$", wb.path) is not None
