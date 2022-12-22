"""
get_years.py
"""


def get_years(current_year, n_years=20):
    """
    Description
    -----------
    Get a list of the years from (n_years - 1) years prior to
    the year input to the year input,
    inclusive, for a length of n_years years total.

    Parameters
    ----------
    current_year : int
        The current year.
    n_years : int, optional
        The number of years.

    Returns
    -------
    years : list
        The list of years.

    Raises
    ------
    ValueError
        If the current year is not an integer.
    ValueError
        If the number of years is not an integer.

    Imports
    -------
    None

    Examples
    --------
    >>> get_years(2020, 20)
    [
        2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010,
        2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020
        ]

    >>> get_years(2019, 10)
    [2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019]
    """
    # test whether the current year is an integer
    if not isinstance(current_year, int):
        raise ValueError(f"The current year {current_year} is not an integer.")

    # test whether the number of years is an integer
    if not isinstance(n_years, int):
        raise ValueError(f"The number of years {n_years} is not an integer.")

    # get the current year
    current_year = int(current_year)

    # get the number of years
    n_years = int(n_years)

    # get the list of years
    years = list(range(current_year - n_years + 1, current_year + 1))

    # return the list of years
    return years
