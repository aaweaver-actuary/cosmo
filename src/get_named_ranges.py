"""
get_named_ranges.py
"""

import re
import openpyxl
import pyxlsb

from .is_xlsb import is_xlsb

def get_named_ranges_pyxlsb(wb):
    """
    Description
    -----------
    This function uses the pyxlsb library to
    take a workbook object as input and
    return a dictionary of the named ranges in the workbook,
    where the keys are the names of the named ranges and
    the values are the values of the named ranges.

    Parameters
    ----------
    wb : pyxlsb.Workbook
        Workbook object to be checked.

    Returns
    -------
    dict
        Dictionary of the named ranges in the workbook,
        where the keys are the names of the named ranges and
        the values are the values of the named ranges.

    Raises
    ------
    ValueError
        If the wb is not an xlsb file.

    Imports
    -------
    pyxlsb

    Examples
    --------
    >>> get_named_ranges_pyxlsb(wb)
    {'named_range_1': 'Sheet1!$A$1:$A$2', 'named_range_2': 'Sheet1!$B$1:$B$2'}
    """
    # test if the workbook is an xlsb file
    if is_xlsb(wb):
        # return the named ranges in the workbook
        named_ranges = {}

        # loop through the named ranges in the workbook
        for name in wb.defined_names:
            # add the name and value of the named range to the dictionary
            named_ranges[name.name] = name.value
        return named_ranges
    # if the workbook is not an xlsb file, return nothing and raise a value error
    else:
        raise ValueError("wb is not an xlsb file")


def get_named_ranges_openpyxl(wb):
    """
    Description
    -----------
    This function uses the openpyxl library to take a workbook object as input
    and return a dictionary of the named ranges in the workbook,
    where the keys are the names of the named ranges and
    the values are the values of the named ranges.

    Parameters
    ----------
    wb : openpyxl.Workbook
        Workbook object to be checked.

    Returns
    -------
    dict
        Dictionary of the named ranges in the workbook,
        where the keys are the names of the named ranges and
        the values are the values of the named ranges.

    Raises
    ------
    ValueError
        If the wb is not an openpyxl workbook object,
        with file extension ".xlsx", ".xlsm", or ".xltx".

    Imports
    -------
    openpyxl
    re

    Examples
    --------
    >>> get_named_ranges_openpyxl(wb)
    {'named_range_1': 'Sheet1!$A$1:$A$2', 'named_range_2': 'Sheet1!$B$1:$B$2'}
    """
    # test if the workbook is an openpyxl workbook object
    if isinstance(wb, openpyxl.Workbook):
        # test if the workbook is an openpyxl workbook object
        # with file extension ".xlsx", ".xlsm", or ".xltx"
        if re.search(r"\.(xlsx|xlsm|xltx)$", wb.path) is not None:
            # return the named ranges in the workbook
            named_ranges = {}

            # loop through the named ranges in the workbook
            for name in wb.defined_names.definedName:
                # add the name and value of the named range to the dictionary
                named_ranges[name.name] = name.value
            return named_ranges
        # if the workbook is not an openpyxl workbook object
        # with file extension ".xlsx", ".xlsm", or ".xltx",
        # return nothing and raise a value error
        else:
            raise ValueError("""wb is not an openpyxl workbook object,
            with file extension '.xlsx', '.xlsm', or '.xltx'""")
    # if the workbook is not an openpyxl workbook object,
    # return nothing and raise a value error
    else:
        raise ValueError("wb is not an openpyxl workbook object")


def get_named_ranges(wb):
    """
    Description
    -----------
    This function takes a workbook object as input and
    returns a dictionary of the named ranges in the workbook,
    where the keys are the names of the named ranges and
    the values are the values of the named ranges.

    Parameters
    ----------
    wb : openpyxl.Workbook or pyxlsb.Workbook
        Workbook object to be checked.

    Returns
    -------
    dict
        Dictionary of the named ranges in the workbook,
        where the keys are the names of the named ranges and
        the values are the values of the named ranges.

    Raises
    ------
    ValueError
        If the wb is not a workbook object.

    Imports
    -------
    openpyxl
    pyxlsb

    Examples
    --------
    >>> get_named_ranges(wb)
    {'named_range_1': 'Sheet1!$A$1:$A$2',
    'named_range_2': 'Sheet1!$B$1:$B$2'}
    """
    # first, test if the workbook is an xlsb file and
    # use the appropriate function to get the named ranges
    if isinstance(wb, openpyxl.Workbook):
        return get_named_ranges_openpyxl(wb)
    # if not, test if the workbook is an xlsb file and
    # use the appropriate function to get the named ranges
    elif isinstance(wb, pyxlsb.Workbook):
        return get_named_ranges_pyxlsb(wb)
    else:
        raise ValueError("wb is not a workbook object")
