"""
get_prior_quarter_year.py
"""



def get_prior_quarter_year(year, quarter):
    """
    Description
    -----------
    Get the quarter and year of the previous quarter.

    Parameters
    ----------
    year : int
        The year.
    quarter : int
        The quarter.

    Returns
    -------
    prior_year : int
        The year of the previous quarter.
    prior_qtr : int
        The quarter of the previous quarter.

    Raises
    ------
    ValueError
        If the year is not an integer.
    ValueError
        If the quarter is not an integer.
    ValueError
        If the quarter is not a number 1-4.

    Imports
    -------
    None

    Examples
    --------
    >>> year = 2018
    >>> quarter = 1
    >>> prior_year, prior_qtr = get_prior_quarter_year(year, quarter)
    >>> prior_year
    2017
    >>> prior_qtr
    4

    >>> year = 2018
    >>> quarter = 4
    >>> prior_year, prior_qtr = get_prior_quarter_year(year, quarter)
    >>> prior_year
    2018
    >>> prior_qtr
    3
    """
    # checks that the year is an integer
    # if the year is not an integer, raise a value error
    if not isinstance(year, int):
        raise ValueError(f"The year {year} is not an integer.")

    # checks that the quarter is an integer
    # if the quarter is not an integer, raise a value error
    if not isinstance(quarter, int):
        raise ValueError(f"The quarter {quarter} is not an integer.")

    # checks that the quarter is a number 1-4
    # if the quarter is not a number 1-4, raise a value error
    if not 1 <= quarter <= 4:
        raise ValueError(f"The quarter {quarter} is not a number 1-4.")

    # if the quarter is 1, the previous quarter is 4 of the previous year
    if quarter == 1:
        prior_year = year - 1
        prior_qtr = 4

    # if the quarter is not 1, the previous quarter is
    # the current quarter minus 1, in the current year
    else:
        prior_year = year
        prior_qtr = quarter - 1

    # return the quarter and year of the previous quarter
    return prior_year, prior_qtr
