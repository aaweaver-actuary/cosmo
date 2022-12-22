"""
get_quarter_year.py
"""
import re

def get_quarter_year(filepath):
    """
    Description
    -----------
    Extract the quarter and year from the file name.

    Parameters
    ----------
    filepath : str
        The file path.

    Returns
    -------
    quarter : int
        The quarter.
    year : int
        The year.

    Raises
    ------
    ValueError
        If the file path does not have the substring number 1-4,
        "Q", and the year (4-digit number) with no spaces in between.

    Imports
    -------
    re

    Examples
    --------
    >>> filepath = "C:\\Users\\test\\test1Q2018.xlsb"
    >>> quarter, year = get_quarter_year(filepath)
    >>> quarter
    1
    >>> year
    2018

    >>> filepath = "C:\\Users\\test\\3Q2023 Analysis\\test_file.xlsb"
    >>> quarter, year = get_quarter_year(filepath)
    >>> quarter
    3
    >>> year
    2023

    >>> filepath = "C:\\Users\\test\\1999q2 Analysis\\test_file.xlsb"
    >>> quarter, year = get_quarter_year(filepath)
    >>> quarter
    2
    >>> year
    1999
    """
    # checks that the file path is a string
    # if the file path is not a string, raise a value error
    if not isinstance(filepath, str):
        raise ValueError(f"The file path {filepath} is not a string.")

    # get the quarter and year
    # the quarter and year are extracted using re.search
    # the quarter may come first, or the year may come first,
    # but either way they are separated by the "Q"

    # expects the file name to have the substring number 1-4,
    # "Q", and the year (4-digit number) with no spaces in between
    # or the 4-digit year, "Q", and the quarter number 1-4
    # with no spaces in between
    # the "Q" may or may not be uppercase,
    # so need to use re.IGNORECASE
    # you know it is the quarter because it is a number 1-4,
    # you know it is the year because it is a 4-digit number

    # search for the quarter and year with quarter coming first,
    # return None if it is not found
    quarter_year = re.search(r"([1-4])Q(\d{4})", filepath, re.IGNORECASE)

    # search for the quarter and year with year coming first
    # if quarter_year is None,
    # return None if it is not found
    quarter_year = (
        re.search(r"(\d{4})Q([1-4])", filepath, re.IGNORECASE)
        if quarter_year is None
        else quarter_year
        )

    # if both searches returned None, raise a value error
    if quarter_year is None:
        raise ValueError(f"""The file path {filepath}
        does not have the substring with a number 1-4,
        'Q', and the year (4-digit number) with no spaces in between.""")

    # otherwise convert both to integers
    quarter = int(quarter_year.group(1))
    year = int(quarter_year.group(2))

    # return the quarter and year
    return quarter, year
