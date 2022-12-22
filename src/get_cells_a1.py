"""
get_cells_a1.py
"""

from .column_letter_from_index import column_letter_from_index

def get_cells_a1(cells_r1c1):
    """
    Description
    -----------
    This function takes a list of cells in R1C1 notation
    as tuples of the form (row, column) as input and
    returns a list of the cells in A1 notation as strings.

    Parameters
    ----------
    cells_r1c1 : list
        List of cells in R1C1 notation as tuples of the form (row, column).

    Returns
    -------
    list
        List of the cells in A1 notation as strings.

    Raises
    ------
    ValueError
        If the cells_r1c1 is not a list.
    ValueError
        If the cells_r1c1 is not a list of cells in R1C1 notation.

    Imports
    -------
    column_letter_from_index
        This function takes a column index as input and returns the column letter as a string.
    get_column_letter
        This function takes a column index as input and returns the column letter as a string.

    Examples
    --------
    >>> get_cells_a1([(1, 1)])
    ["A1"]
    >>> get_cells_a1([(1, 1), (2, 1)])
    ["A1", "A2"]
    >>> get_cells_a1([(1, 1), (2, 1), (1, 2), (2, 2)])
    ["A1", "A2", "B1", "B2"]
    """
    # check that the input is a list
    if not isinstance(cells_r1c1, list):
        raise ValueError("cells_r1c1 is not a list")

    # check that the input is a list of cells in R1C1 notation
    # a cell in R1C1 notation should be a tuple of two integers
    # define a function to check that a cell is a tuple of two integers in the form (row, column)
    def cond(cell):
        return (
            # check that the cell is a tuple
            isinstance(cell, tuple) and

            # check that the cell is a tuple of two elements
            len(cell) == 2 and

            # check that the first element is an integer
            isinstance(cell[0], int) and

            # check that the second element is an integer
            isinstance(cell[1], int)
            )

    # check that the input is a list of cells in R1C1 notation
    if not all([cond(cell) for cell in cells_r1c1]):
        raise ValueError(f"cells_r1c1 {cells_r1c1} is not a list of cells in R1C1 notation")

    # initialize a list to store the cells in A1 notation
    cells_a1 = []

    # loop through the cells in R1C1 notation
    for cell_r1c1 in cells_r1c1:
        # get the row number
        row_number = cell_r1c1[0]
        # get the column number
        column_number = cell_r1c1[1]
        # get the column letter
        column_letter = column_letter_from_index(column_number)
        # add the cell in A1 notation to the list of cells in A1 notation
        cells_a1.append(column_letter + str(row_number))

    # return the list of cells in A1 notation
    return cells_a1
