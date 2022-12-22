"""
update_named_range.py
"""

from .get_cells_a1 import get_cells_a1

def update_named_range_pyxlsb(wb, named_range, value):
    """
    Description
    -----------
    This function takes a wb object as input,
    named range as input and a value as input,
    and updates the named range with the value.
    This is for a ".xlsb" file only.

    Parameters
    ----------
    wb : wb object
        wb object.
    named_range : str
        Named range.
    value : list
        Value.

    Returns
    -------
    None

    Raises
    ------
    ValueError
        If the named range is not found.
    ValueError
        If the value is not a list of the same length as the named range.
    ValueError
        If the wb object is not a wb object.
    ValueError
        If the wb object does not refer to a ".xlsb" file.

    Imports
    -------
    pyxlsb

    Examples
    --------
    >>> update_named_range_xlsb(wb, "named_range", [1, 2, 3])
    None
    >>> update_named_range_xlsb(wb, "named_range", [1, 2, 3, 4])
    ValueError: The value is not a list of the same length as the named range.
    >>> update_named_range_xlsb(wb, "named_range", 1)
    None
    >>> update_named_range_xlsb(wb, "named_range", [1])
    None
    >>> update_named_range_xlsb(wb, "named_range", [1, 2, 3])
    None
    >>> update_named_range_xlsb(wb, "named_range", [1, 2, 3, 4])
    ValueError: The value is not a list of the same length as the named range.
    >>> update_named_range_xlsb(wb, "named_range", 1)
    None
    """
    # check that the wb object is a wb object
    if not isinstance(wb, wb):
        raise ValueError("The wb object is not a wb object.")

    # check that the wb object refers to a ".xlsb" file
    if not wb.filename.endswith(".xlsb"):
        raise ValueError("The wb object does not refer to a .xlsb file.")

    # get the workbook object
    workbook = wb.active

    # get the named ranges
    named_ranges = workbook.get_named_ranges()

    # check that the named range is found
    if named_range not in named_ranges:
        raise ValueError("The named range is not found.")

    # get the named range
    named_range = named_ranges[named_range]

    # get the cells in the named range
    cells = named_range.destinations

    # get the cells in A1 notation
    cells_a1 = get_cells_a1(cells)

    # get the number of cells in the named range
    number_of_cells = len(cells_a1)

    # check that the value is a list of the same length as the named range
    if not isinstance(value, list):
        value = [value]
    if len(value) != number_of_cells:
        raise ValueError("The value is not a list of the same length as the named range.")

    # loop through the cells in the named range
    for i in range(number_of_cells):
        # get the cell in A1 notation
        cell_a1 = cells_a1[i]
        # get the value
        value_i = value[i]
        # update the cell with the value
        workbook[cell_a1] = value_i


def update_named_range_openpyxl(wb, named_range, value):
    """This function takes a wb object as input,
    named range as input and a value as input,
    and updates the named range with the value.
    This is for a openpyxl format-friendly files only.
    .xlsx, .xlsm, .xltx, .xltm

    Parameters
    ----------
    wb : wb object
        wb object.
    named_range : str
        Named range.
    value : list
        Value.

    Returns
    -------
    None

    Raises
    ------
    ValueError
        If the named range is not found.
    ValueError
        If the value is not a list of the same length as the named range.
    ValueError
        If the wb object is not a wb object.
    ValueError
        If the wb object does not refer to a file that is one of .xlsx, .xlsm, .xltx, or .xltm.

    Examples
    --------
    >>> update_named_range_xlsx(wb, "named_range", [1, 2, 3])
    None
    >>> update_named_range_xlsx(wb, "named_range", [1, 2, 3, 4])
    ValueError: The value is not a list of the same length as the named range.
    >>> update_named_range_xlsx(wb, "named_range", 1)
    None
    >>> update_named_range_xlsx(wb, "named_range", [1])
    None
    >>> update_named_range_xlsx(wb, "named_range", [1, 2, 3])
    None
    >>> update_named_range_xlsx(wb, "named_range", [1, 2, 3, 4])
    ValueError: The value is not a list of the same length as the named range.
    >>> update_named_range_xlsx(wb, "named_range", 1)
    None
    """
    # check that the wb object is a wb object
    if not isinstance(wb, wb):
        raise ValueError("The wb object is not a wb object.")

    # get the workbook object
    workbook = wb.active

    # get the named ranges
    named_ranges = workbook.defined_names.definedName

    # check that the named range is found
    if named_range not in named_ranges:
        raise ValueError("The named range is not found.")

    # get the named range
    named_range = named_ranges[named_range]

    # get the cells in the named range
    cells = named_range.destinations

    # get the cells in A1 notation
    cells_a1 = get_cells_a1(cells)

    # get the number of cells in the named range
    number_of_cells = len(cells_a1)

    # check that the value is a list of the same length as the named range
    # if the named range is a single cell, the value does not need to be a list
    if not isinstance(value, list):
        value = [value]
    if len(value) != number_of_cells:
        raise ValueError("The value is not a list of the same length as the named range.")

    # loop through the cells in the named range
    for i in range(number_of_cells):
        # get the cell in A1 notation
        cell_a1 = cells_a1[i]
        # get the value
        value_i = value[i]
        # update the cell with the value
        workbook[cell_a1] = value_i


def update_named_range(wb, named_range, value):
    """This function combines the two functions above.
    Takes a wb object as input, named range as input and a value as input,
    and updates the named range with the value.
    First determines which of `update_named_range_openpyxl`
    or `update_named_range_pyxlsb` to use based on
    the file extension of the wb object.

    Parameters
    ----------
    wb : wb object
        wb object.
    named_range : str
        Named range.
    value : list
        Value.

    Returns
    -------
    None

    Raises
    ------
    ValueError
        If the wb object is not a wb object.
    ValueError
        If the wb object does not refer to a file that is
        one of .xlsx, .xlsm, .xltx, .xltm, or .xlsb.

    Examples
    --------
    >>> update_named_range(wb, "named_range", [1, 2, 3])
    None
    >>> update_named_range(wb, "named_range", [1, 2, 3, 4])
    ValueError: The value is not a list of the same length as the named range.
    >>> update_named_range(wb, "named_range", 1)
    None
    >>> update_named_range(wb, "named_range", [1])
    None
    >>> update_named_range(wb, "named_range", [1, 2, 3])
    None
    >>> update_named_range(wb, "named_range", [1, 2, 3, 4])
    ValueError: The value is not a list of the same length as the named range.
    >>> update_named_range(wb, "named_range", 1)
    None
    """
    # check that the wb object is a wb object
    if not isinstance(wb, wb):
        raise ValueError("The wb object is not a wb object.")

    # get the file name
    file_name = wb.file_name

    # get the file extension
    file_extension = file_name.split(".")[-1]

    # if the file extension is .xlsb, use the pyxlsb function
    if file_extension == "xlsb":
        update_named_range_pyxlsb(wb, named_range, value)
    # if the file extension is .xlsx, .xlsm, .xltx, or .xlt use the openpyxl function
    elif file_extension in ["xlsx", "xlsm", "xltx", "xltm"]:
        update_named_range_openpyxl(wb, named_range, value)
    # otherwise, raise a value error
    else:
        raise ValueError("""The wb object does not refer to a file that
        is one of .xlsx, .xlsm, .xltx, .xltm, or .xlsb.""")
