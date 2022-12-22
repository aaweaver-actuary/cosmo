"""
get_cells_r1c1.py
"""
import re

from .column_index_from_string import column_index_from_string

def get_cells_r1c1(cells):
    """
    Description
    -----------
    This function takes a list of cells as input
    and returns a list of the cells in R1C1 notation
    as tuples of the form (row, column).

    Parameters
    ----------
    cells : list
        List of cells.

    Returns
    -------
    list
        List of the cells in R1C1 notation as
        tuples of the form (row, column).

    Raises
    ------
    ValueError
        If the cells is not a list.
    ValueError
        If the cells is not a list of cells.

    Imports
    -------
    re


    Examples
    --------
    >>> get_cells_r1c1(["A1"])
    [(1, 1)]
    >>> get_cells_r1c1(["A1", "A2"])
    [(1, 1), (2, 1)]
    >>> get_cells_r1c1(["A1", "A2", "B1", "B2"])
    [(1, 1), (2, 1), (1, 2), (2, 2)]
    """
    # check that the input is a list
    if not isinstance(cells, list):
        raise ValueError(f"cells {cells} is not a list")

    # check that the input is a list of cells
    # a cell should be 1-3 upper case letters followed by 1-7 digits
    if not all([re.match(r"^[A-Z]{1,3}[0-9]{1,7}$", cell) for cell in cells]):
        raise ValueError(f"cells {cells} is not a list of cells")

    # initialize a list to store the cells in R1C1 notation
    cells_r1c1 = []

    # loop through the cells
    for cell in cells:
        # get the column letter
        column_letter = cell[0]
        # get the row number
        row_number = int(cell[1:])
        # get the column number
        column_number = column_index_from_string(column_letter)
        # add the cell in R1C1 notation to the list of cells in R1C1 notation
        cells_r1c1.append((row_number, column_number))

    # return the list of cells in R1C1 notation
    return cells_r1c1
