"""
This module contains utility functions that are used in the package,
and in the class definitions. These functions are not intended to be
used by the end user.

This module is a container that collects funcitons from the following modules:

    utils._get_utils
    utils._transform_utils
    utils._testing_utils
    utils._update_utils

Does not import functions that are considered support functions,
such as functions that are different depending on whether the
file is an xlsb or xlsx file.

Does not contain any classes or functions that are intended to be used by the end user.
"""
import pyxlsb
import openpyxl

# Import functions from the following modules:
from .utils._testing_utils import (
    is_wb
    , is_xlsb
    , is_a1_cell
)
from .utils._get_utils import (
    get_cells_from_range
    , get_named_ranges
    , get_cells_a1
    , find_links
    )
from .utils._transform_utils import (
    column_letter_from_index
    , column_index_from_string
    , to_a1_cell
    , to_rc_cell
    )
from .utils._update_utils import (
    update_range
    , update_named_range
    , update_links
    )

# function that takes a path to an excel workbook as input and returns the wb object
# similar docstring as before
# returns the wb object
# the wb object is created differently depending on whether the file is an xlsb or an xlsx, xlsm, xls, or xltx file, 
# so first need to use is_xlsb function to determine which method to use, either pyxlsb or openpyxl
def open_workbook(filepath):
    """
    Open the workbook.

    Parameters
    ----------
    filepath : str
        The file path.

    Returns
    -------
    wb : object
        The workbook object.

    Raises
    ------
    ValueError
        If the file path does not end with ".xlsb", ".xlsx", ".xlsm", ".xls", or ".xltx".
    
    Examples
    --------
    >>> filepath = "C:\\Users\\test\\test1Q2018.xlsb"
    >>> wb = open_workbook(filepath)
    """

    # check to make sure the file path ends with ".xlsb", ".xlsx", ".xlsm", ".xls", or ".xltx"
    # if the file path does not end with ".xlsb", ".xlsx", ".xlsm", ".xls", or ".xltx", raise a value error
    if not filepath.endswith((".xlsb", ".xlsx", ".xlsm", ".xls", ".xltx")):
        raise ValueError("The file path {} does not end with '.xlsb', '.xlsx', '.xlsm', '.xls', or '.xltx'.".format(filepath))

    # if the file is an xlsb file, use pyxlsb to open the workbook
    if is_xlsb(filepath):
        # use pxlsb to open the workbook and create the wb object
        return pyxlsb.open_workbook(filepath)
    # if the file is not an xlsb file, use openpyxl to open the workbook
    else:
        return openpyxl.load_workbook(filepath)