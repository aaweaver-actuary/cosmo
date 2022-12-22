"""
to_rc_cell.py
"""

import re

def to_rc_cell(cell):
    """This function is the opposite of `to_a1_cell`,
    and instead transforms to (row, column) notation.
    It uses get_cells_r1c1 to transform a list of excel cell references
    to a list of every (row, column) cell in that range.

    Parameters
    ----------
    cell : tuple, str, or list
        Cell.

    Returns
    -------
    tuple or list
        Tuple or list of tuples.

    Examples
    --------
    >>> to_rc_cell("A1")
    (1, 1)
    >>> to_rc_cell("A1:B2")
    [(1, 1), (1, 2), (2, 1), (2, 2)]
    >>> to_rc_cell("A1:B")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_rc_cell("A1B2")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_rc_cell("A1B")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_rc_cell("A1:B2", "A1:B2")
    [(1, 1), (1, 2), (2, 1), (2, 2), (1, 1), (1, 2), (2, 1), (2, 2)]
    >>> to_rc_cell("A1:B2", "A1:B")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_rc_cell("A1:B2", "A1B2")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_rc_cell("A1:B2", "A1B")
    ValueError: The cell is not in A1 notation or (row, column) notation.
    >>> to_rc_cell("A1:B2", ["A1:B2", "A1:B", "A1B2", "A1B"])
    [[(1, 1), (1, 2), (2, 1), (2, 2)],
    ValueError: The cell is not in A1 notation or (row, column) notation.,
    ValueError: The cell is not in A1 notation or (row, column) notation.,
    ValueError: The cell is not in A1 notation or (row, column) notation.]
    """
    # if the input is a tuple or string, return a tuple
    if isinstance(cell, tuple) or isinstance(cell, str):
        # check that the cell is in A1 notation
        # if the cell is in A1 notation, return the cell
        # if the cell is not in A1 notation, check that the cell is in (row, column) notation
        # if the cell is in (row, column) notation,
        # transform the cell to A1 notation and return the cell
        # if the cell is not in (row, column) notation, raise a value error
        if re.match(r"^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$", cell) is not None:
            def cell_out(cell):
                out = [int(cell[1]) if
                len(cell) == 2 else
                (int(cell[1:3]), int(cell[4:6]))
                for cell in cell.split(":")]
                return out

            return tuple(out for out in cell_out(cell))
        elif re.match(r"^\([0-9]+, [0-9]+\)(:\([0-9]+, [0-9]+\))?$", cell) is not None:
            return (tuple([int(cell[1]) if
            len(cell) == 3 else
            (int(cell[1]), int(cell[3])) for
            cell in cell.split(":")]))
        else:
            raise ValueError("The cell is not in A1 notation or (row, column) notation.")
    # if the input is a list, return a list of tuples
    elif isinstance(cell, list):
        # loop through the list
        # check that each cell is in A1 notation
        # if the cell is in A1 notation, append the cell to the
        # list of tuples
        # if the cell is not in A1 notation,
        # check that the cell is in (row, column) notation
        # if the cell is in (row, column) notation,
        # transform the cell to A1 notation and append the cell to the
        # list of tuples
        # if the cell is not in (row, column) notation, raise a value error

        def get_list(cell):
            return int(cell[1]) if len(cell) == 2 else (int(cell[1:3]), int(cell[4:6]))

        # return the list of tuples
        return [
            tuple(
                [get_list(cell) for cell in cell.split(":")]
                ) if
                re.match(r"^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$"
                , cell) is not None else
                (int(cell[1]) if
                len(cell) == 3 else
                (int(cell[1]), int(cell[3]))) if
                re.match(r"^\([0-9]+, [0-9]+\)(:\([0-9]+, [0-9]+\))?$"
                , cell) is not None else
                ValueError("The cell is not in A1 notation or (row, column) notation."
                ) for cell in cell]
