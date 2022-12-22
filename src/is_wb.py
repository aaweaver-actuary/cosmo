"""
is_wb.py
"""
import openpyxl
import pyxlsb

def is_wb(wb):
    """
    Description
    -----------
    This function takes a workbook object as input
    and returns True if it is either
    a pyxlsb or openpyxl workbook object
    and False if it is not a pyxlsb or openpyxl workbook object.

    Parameters
    ----------
    wb : pyxlsb.Workbook or openpyxl.Workbook
        Workbook object to be checked.

    Returns
    -------
    bool
        True if the wb is a pyxlsb or openpyxl workbook object and
        False if it is not a pyxlsb or openpyxl workbook object.

    Raises
    ------
    ValueError
        If the wb is not a pyxlsb or openpyxl workbook object.

    Imports
    -------
    openpyxl
    pyxlsb

    Examples
    --------
    >>> is_wb(wb)
    True

    >>> is_wb(245)
    False
    """
    # test if the object passed is a workbook object,
    # and if not, return False
    if (
        not isinstance(wb, openpyxl.Workbook) and
        not isinstance(wb, pyxlsb.Workbook)
        ):
        return False
    # otherwise, return True
    else:
        return True
