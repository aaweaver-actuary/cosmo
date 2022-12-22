"""
column_letter_from_index.py
"""


def column_letter_from_index(column_index):
    """
    Description
    -----------
    This function takes an integer representing
    a column in Excel as input and returns the column letter.

    Parameters
    ----------
    column_index : int
        Integer representing a column in Excel.

    Returns
    -------
    str
        Column letter.

    Raises
    ------
    ValueError
        If the column_index is not an integer.
    ValueError
        If the column_index is not a column in Excel.

    Imports
    -------
    None

    Examples
    --------
    >>> column_letter_from_index(1)
    'A'
    >>> column_letter_from_index(27)
    'AA'
    >>> column_letter_from_index(702)
    'ZZ'
    """
    # check that the input is an integer
    if not isinstance(column_index, int):
        raise ValueError(f"column_index {column_index} is not an integer")

    # check that the input is a column in Excel
    if not 1 <= column_index <= 18278:
        raise ValueError(f"column_index {column_index} is not a column in Excel")

    # initialize a list to store the column letters
    column_letters = []

    # loop through the column index
    while column_index > 0:
        # subtract 1 from the column index
        column_index -= 1
        # get the column letter
        column_letter = chr(ord("A") + column_index % 26)
        # add the column letter to the list of column letters
        column_letters.append(column_letter)
        # divide the column index by 26
        column_index //= 26

    # return the column letter
    return "".join(column_letters[::-1])
    