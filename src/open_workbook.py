"""
open_workbook.py
"""
import openpyxl
import pyxlsb

from .is_xlsb import is_xlsb
from .is_wb import is_wb

# function to open an excel workbook
# takes an input file name as a string
# returns a workbook object
# this version is for use with the openpyxl module
def open_workbook_openpyxl(file_name):
    """
    Description
    -----------
    Open an Excel workbook.

    Parameters
    ----------
    file_name : str
        The name of the file to open.
        Must have the file extension .xlsx or .xlsm.

    Returns
    -------
    wb : Workbook
        The workbook object.

    Raises
    ------
    None

    Imports
    -------
    openpyxl

    Examples
    --------
    >>> wb = open_workbook_openpyxl('test.xlsx')
    """
    # check if the file is an xlsb file
    # if it is, raise an error
    if is_xlsb(file_name):
        raise Exception('File is an xlsb file. Use open_workbook_pyxlsb() instead.')

    # open the workbook
    wb = openpyxl.load_workbook(file_name)
    return wb

# similar funciton to above but for use with the pyxlsb module
def open_workbook_pyxlsb(file_name):
    """
    Description
    -----------
    Open an Excel xlsb workbook.

    Parameters
    ----------
    file_name : str
        The name of the file to open.
        Must be an xlsb file.

    Returns
    -------
    wb : Workbook
        The workbook object.

    Raises
    ------
    None

    Imports
    -------
    pyxlsb

    Examples
    --------
    >>> wb = open_workbook_pyxlsb('test.xlsb')
    """
    # check if the file is an xlsb file
    # if it is NOT, raise an error
    if not is_xlsb(file_name):
        raise Exception('File is not an xlsb file. Use open_workbook_openpyxl() instead.')

    # open the workbook
    wb = pyxlsb.open_workbook(file_name)
    return wb

# function to open an excel workbook
# takes an input file name as a string
# returns a workbook object
# this version combines the above two functions
# first test if the file is an xlsb file using is_xlsb()
# if it is, use open_workbook_pyxlsb()
# if it is not, use open_workbook_openpyxl()
def open_workbook(file_name):
    """
    Description
    -----------
    Open an Excel workbook.

    Parameters
    ----------
    file_name : str
        The name of the file to open.
        Must have the file extension .xlsx, .xlsm, or .xlsb.

    Returns
    -------
    wb : Workbook
        The workbook object.

    Raises
    ------
    None

    Imports
    -------
    openpyxl
    pyxlsb

    Examples
    --------
    >>> wb = open_workbook('test.xlsx')
    >>> wb = open_workbook('test.xlsb')
    """
    # check if the file is an xlsb file
    # if it is, use open_workbook_pyxlsb()
    if is_xlsb(file_name):
        wb = open_workbook_pyxlsb(file_name)
    # if it is not, use open_workbook_openpyxl()
    else:
        wb = open_workbook_openpyxl(file_name)
    return wb
