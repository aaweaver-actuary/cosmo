"""
column_index_from_string.py
"""

import re

def column_index_from_string(column_string):
    """
    Description
    -----------
    This function takes a string representing
    a column in Excel as input and returns the column index.

    Parameters
    ----------
    column_string : str
        String representing a column in Excel.

    Returns
    -------
    int
        Column index.

    Raises
    ------
    ValueError
        If the column_string is not a string.
    ValueError
        If the column_string is not a column in Excel.

    Imports
    -------
    re

    Examples
    --------
    >>> column_index_from_string("A")
    1
    >>> column_index_from_string("AA")
    27
    >>> column_index_from_string("AAA")
    703
    """
    # check that the input is a string
    if not isinstance(column_string, str):
        raise ValueError(f"column_string {column_string} is not a string")

    # check that the input is a column in Excel
    # a column should be 1-3 upper case letters
    if not re.match(r"^[A-Z]{1,3}$", column_string):
        raise ValueError(f"column_string {column_string} is not a column in Excel")

    # initialize the column index
    column_index = 0

    # loop through the characters in the string
    for i, char in enumerate(column_string):
        # add the column index for the character to the column index
        column_index += (ord(char) - 64) * 26 ** (len(column_string) - i - 1)

    # return the column index
    return column_index
    