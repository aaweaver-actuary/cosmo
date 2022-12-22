"""
to_a1_cell.py
"""
import re

def to_a1_cell(cell):
    """
    Description
    -----------
    This function uses the above function
    to transform a list of excel cell references to A1 notation.
    It checks that a list of cells are in A1 notation or (row, column) notation,
    and returns a list where every element is a string in A1 notation.

    Parameters
    ----------
    cell : tuple, str, or list
        Cell.

    Returns
    -------
    str or list
        String or list of strings.

    Imports
    -------
    re

    Examples
    --------
    >>> to_a1_cell("A1")
    "A1"
    >>> to_a1_cell("A1:B2")
    "A1:B2"
    >>> to_a1_cell("A1:B")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_a1_cell("A1B2")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_a1_cell("A1B")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_a1_cell("A1:B2", "A1:B2")
    ["A1:B2", "A1:B2"]
    >>> to_a1_cell("A1:B2", "A1:B")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_a1_cell("A1:B2", "A1B2")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_a1_cell("A1:B2", "A1B")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_a1_cell("A1:B2", ["A1:B2", "A1:B", "A1B2", "A1B"])
    ["A1:B2",
    ValueError: The cell is not in A1 notation or (row, column) notation.,
    ValueError: The cell is not in A1 notation or (row, column) notation.,
    ValueError: The cell is not in A1 notation or (row, column) notation.]
    >>> to_a1_cell((1, 1))
    "A1"
    >>> to_a1_cell((1, 1), (1, 2))
    ["A1", "A2"]
    >>> to_a1_cell((1, 1), (1, 2), (1, 3))
    ["A1", "A2", "A3"]
    """


    # if the input is a tuple or string, return a string
    if isinstance(cell, tuple) or isinstance(cell, str):
        # check that the cell is in A1 notation
        # if the cell is in A1 notation, return the cell
        # if the cell is not in A1 notation,
        # check that the cell is in (row, column) notation
        # if the cell is in (row, column) notation,
        # transform the cell to A1 notation and return the cell
        # if the cell is not in (row, column) notation, raise a value error
        if re.match(r"^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$", cell) is not None:
            return cell
        elif re.match(r"^\([0-9]+, [0-9]+\)(:\([0-9]+, [0-9]+\))?$", cell) is not None:
            def out1(cell):
                cell = cell[1:-1].split(", ")
                return chr(int(cell[1]) + 64) + cell[0]
            out = [out1(cell) for cell in cell.split(":")]
            return "".join(out)
        else:
            raise ValueError("The cell is not in A1 notation or (row, column) notation.")
    # if the input is a list, return a list of strings
    elif isinstance(cell, list):
        # loop through the list
        # check that each cell is in A1 notation
        # if the cell is in A1 notation, append the cell to the
        # list of strings
        # if the cell is not in A1 notation,
        # check that the cell is in (row, column) notation
        # if the cell is in (row, column) notation,
        # transform the cell to A1 notation and append the cell to the
        # list of strings
        # if the cell is not in (row, column) notation, raise a value error

        # function that uses the above function to
        # transform a list of excel cell references to A1 notation
        def chg_to_a1(cell):
            out = chr(int(cell[1]) + 64) + cell[3]
            out1 = chr(int(cell[6]) + 64) + cell[8]
            out = out + (":" + out1 if len(cell) > 9 else "")
            return out
        # return the list of strings

        join_1 = [chg_to_a1(cell) for cell in cell.split(":")]
        re_match = re.match(r"^\([0-9]+, [0-9]+\)(:\([0-9]+, [0-9]+\))?$", cell)
        return ["".join(join_1) if re_match is not None else cell for cell in cell]
