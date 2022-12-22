"""
is_a1_cell.py
"""

import re

def is_a1_cell(cell):
    """
    Description
    -----------
    This function takes either a tuple, string, or list,
    and returns either a boolean or a list of booleans.
    If the input is a tuple or string, returns a boolean,
    and if the input is a list, returns a list of booleans.
    The boolean is true if the input is an excel cell reference
    in A1 notation, or an excel cell range reference,
    and false otherwise.

    Parameters
    ----------
    cell : tuple, str, or list
        Cell.

    Returns
    -------
    bool or list
        Boolean or list of booleans.

    Imports
    -------
    re

    Examples
    --------
    >>> is_a1_cell("A1")
    True
    >>> is_a1_cell("A1:B2")
    True
    >>> is_a1_cell("A1:B")
    False
    >>> is_a1_cell("A1B2")
    False
    >>> is_a1_cell("A1B")
    False
    >>> is_a1_cell("A1:B2", "A1:B2")
    [True, True]
    >>> is_a1_cell("A1:B2", "A1:B")
    [True, False]
    >>> is_a1_cell("A1:B2", "A1B2")
    [True, False]
    >>> is_a1_cell("A1:B2", "A1B")
    [True, False]
    >>> is_a1_cell("A1:B2", ["A1:B2", "A1:B", "A1B2", "A1B"])
    [True, False, False, False]
    """
    # if the input is a tuple or string, return a boolean
    if isinstance(cell, tuple) or isinstance(cell, str):
        # check that the cell is in A1 notation
        # if the cell is in A1 notation, return True
        # if the cell is not in A1 notation, return False
        return re.match(r"^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$", cell) is not None
    # if the input is a list, return a list of booleans
    elif isinstance(cell, list):
        # loop through the list
        # check that each cell is in A1 notation
        # if the cell is in A1 notation, append True to the
        # list of booleans
        # if the cell is not in A1 notation, append False to the
        # list of booleans
        # return the list of booleans
        return [re.match(r"^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$", cell) is not None for cell in cell]
    # otherwise, raise a value error
    else:
        raise ValueError("The cell is not a tuple, string, or list.")
