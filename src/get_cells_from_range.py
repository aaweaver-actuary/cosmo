"""
get_cells_from_range.py
"""
import re

def get_cells_from_range(range_string):
    """
    Description
    -----------
    This function takes a string representing a cell range in Excel as input
    and returns a list of the cells in the range.

    Parameters
    ----------
    range_string : str
        String representing a cell range in Excel.

    Returns
    -------
    list
        List of the cells in the range.

    Raises
    ------
    ValueError
        If the range_string is not a string.
    ValueError
        If the range_string is not a cell range in Excel.

    Imports
    -------
    re

    Examples
    --------
    >>> get_cells_from_range("A1")
    ['A1']
    >>> get_cells_from_range("A1:A2")
    ['A1', 'A2']
    >>> get_cells_from_range("A1:A2, B1:B2")
    ['A1', 'A2', 'B1', 'B2']
    """
    # check that the input is a string
    if not isinstance(range_string, str):
        raise ValueError("range_string is not a string")

    # check that the input is a cell range in Excel
    # a cell range should be 1-3 upper case letters followed by 1-7 digits,
    # optionally separated by a colon
    # and followed by a second cell range
    re_match_str = r"^[A-Z]{1,3}[0-9]{1,7}(:[A-Z]{1,3}[0-9]{1,7})?"
    re_match_str = re_match_str + r"(,[A-Z]{1,3}[0-9]{1,7}(:[A-Z]{1,3}[0-9]{1,7})?)*$"
    if not re.match(re_match_str, range_string):
        raise ValueError("range_string is not a cell range in Excel")

    # split the input string on commas
    range_list = range_string.split(",")

    # initialize a list to store the cells
    cells = []

    # loop through the list of ranges
    for xlrange in range_list:
        # split xlrange on the colon
        xlrange = xlrange.split(":")

        # if the range only has one element, then it's a single cell
        if len(xlrange) == 1:

            # add the cell to the list of cells
            cells.append(xlrange[0])

        # if the range has two elements, then it's a range of cells
        elif len(xlrange) == 2:

            # get the first cell
            cell_start = xlrange[0]

            # get the last cell
            cell_end = xlrange[1]

            # get the column letter of the first cell
            column_letter_start = cell_start[0]

            # get the column letter of the last cell
            column_letter_end = cell_end[0]

            # get the row number of the first cell
            row_number_start = int(cell_start[1:])

            # get the row number of the last cell
            row_number_end = int(cell_end[1:])

            # loop through the columns
            for column_letter in column_letter_start + column_letter_end:

                # loop through the rows
                for row_number in range(row_number_start, row_number_end + 1):

                    #  add the cell to the list of cells
                    cells.append(column_letter + str(row_number))

        # if the range has more than two elements, then it's not a cell range in Excel
        else:
            raise ValueError("range_string is not a cell range in Excel")
    # return the list of cells
    return cells
