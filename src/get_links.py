"""
get_links.py
"""

import openpyxl
import pyxlsb

from .is_wb import is_wb
from .is_xlsb import is_xlsb


def get_links_pyxlsb(wb):
    """
    Description
    -----------
    Get a list of all the links to other excel files in the wb object,
    including their full paths.

    Parameters
    ----------
    wb : object
        The wb object.

    Returns
    -------
    links : list
        The list of links.

    Raises
    ------
    ValueError
        If the wb object is not a wb object.

    Imports
    -------
    pyxlsb

    Examples
    --------
    >>> import pyxlsb
    >>> wb = pyxlsb.open_workbook("test.xlsb")
    >>> links = get_links_pyxlsb(wb)
    >>> links
    ['C:\\Users\\test\\test2.xlsb']

    >>> get_links_pyxlsb("test.xlsx")
    ValueError: The wb object is not a wb object.
    """
    # checks that the wb object input is a wb object
    # either of these packages can use
    # if the wb object input is not a wb object
    # either of these packages can use, raise a value error
    if not isinstance(wb, pyxlsb.Workbook):
        raise ValueError("The wb object is not a wb object.")

    # get the links
    links = wb.links

    # return the links
    return links


def get_links_openpyxl(wb):
    """
    Description
    -----------
    Get a list of all the links to other excel files in the wb object,
    including their full paths.

    Parameters
    ----------
    wb : object
        The wb object.

    Returns
    -------
    links : list
        The list of links.

    Raises
    ------
    ValueError
        If the wb object is not a wb object.

    Imports
    -------
    openpyxl

    Examples
    --------
    >>> import openpyxl
    >>> wb = openpyxl.load_workbook("test.xlsx")
    >>> links = get_links_openpyxl(wb)
    >>> links
    ['C:\\Users\\test\\test2.xlsx']

    >>> get_links_openpyxl("test.xlsb")
    ValueError: The wb object is not a wb object.
    """
    # checks that the wb object input is a wb object openpyxl can use
    # if the wb object input is not a wb object openpyxl
    # can use (an openpyxl.Workbook object), raise a value error
    if not isinstance(wb, openpyxl.Workbook):
        raise ValueError("The wb object is not a wb object.")

    # checks that the wb object points to a file that
    # ends in .xlsx, .xlsm, or .xltx
    # if the wb object does not point to a file that
    # ends in .xlsx, .xlsm, or .xltx, raise a value error
    if not wb.path.endswith((".xlsx", ".xlsm", ".xltx")):
        raise ValueError("The wb object does not point to a file " +
        "that ends in .xlsx, .xlsm, or .xltx.")

    # get the links
    links = wb._external_links

    # return the links
    return links


def get_links(wb):
    """
    Description
    -----------
    Get a list of all the links to other excel files in the wb object,
    including their full paths.

    Parameters
    ----------
    wb : object
        The wb object.

    Returns
    -------
    links : list
        The list of links.

    Raises
    ------
    ValueError
        If the wb object is not a wb object.

    Imports
    -------
    openpyxl
    pyxlsb

    Examples
    --------
    >>> import pyxlsb
    >>> wb = pyxlsb.open_workbook("test.xlsb")
    >>> links = get_links(wb)
    >>> links
    ['C:\\Users\\test\\test2.xlsb']
    """
    # first test to be sure the wb object is
    # a wb object either of these packages can use
    # if the wb object is not a wb object either of
    # these packages can use, raise a value error
    if not is_wb(wb):
        raise ValueError("The wb object is not a wb object.")

    # if you use pyxlsb, call the get_links_pyxlsb function
    if is_xlsb(wb):
        return get_links_pyxlsb(wb)
    # if you use openpyxl, call the get_links_openpyxl function
    else:
        return get_links_openpyxl(wb)
