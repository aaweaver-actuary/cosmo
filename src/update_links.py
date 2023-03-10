"""
update_links.py
"""

import pyxlsb
import openpyxl

from .find_links import find_links
from .is_xlsb import is_xlsb

def update_links_pyxlsb(wb, links):
    """
    Description
    -----------
    Update the links in the workbook from
    the current links to the desired links.

    Parameters
    ----------
    wb : object
        The workbook object.
        Must be able to be read by pyxlsb.
    links : dict
        The dictionary with the current links as keys
        and the desired links as values.
        Dictionary should be of the form:
            {
                current_link1: desired_link1,
                current_link2: desired_link2,
                ...
            }

    Returns
    -------
    wb : object
        The workbook object.

    Raises
    ------
    ValueError
        If the workbook object is not able to be read by pyxlsb.

    Imports
    -------
    pyxlsb
    openpyxl
    re

    Examples
    --------
    >>> import pyxlsb
    >>> wb = pyxlsb.open_workbook("test.xlsb")
    >>> links = {
        "C:\\Users\\test\\test2.xlsb": "C:\\Users\\test\\test3.xlsb"
    }
    >>> wb = update_links_pyxlsb(wb, links)
    """
    # try to read the workbook with pyxlsb
    # if the workbook is not able to be read by pyxlsb,
    # raise a value error
    try:
        # use pyxlsb to read the workbook and return the wb object
        wb = pyxlsb.open_workbook(wb)

        # get the list of links in the workbook
        current_links = [link[0] for link in wb.links]
    except:
        raise ValueError("The workbook object is not able to be read by pyxlsb.")

    # for each link in the dictionary of links
    for link in links:
        # if the link is in the list of links in the workbook
        if link in current_links:
            # get the index of the link in the list of links in the workbook
            index = current_links.index(link)
            # update the link in the list of links in the workbook
            current_links[index] = links[link]
        # if the link is not in the list of links in the workbook
        else:
            # pass a message to the user
            print(f"Warning: The link {link} is not in the workbook.")
            # continue to the next link
            continue

    # update the links in the workbook
    wb.links = current_links

    # return the workbook object
    return wb


def update_links_openpyxl(wb, links):
    """
    Update the links in the workbook from
    the current links to the desired links.

    Parameters
    ----------
    wb : object
        The workbook object. Must be able to be read by openpyxl.
    links : dict
        The dictionary with the current links as keys
        and the desired links as values.
        Dictionary should be of the form:
            {
                current_link1: desired_link1,
                current_link2: desired_link2,
                ...
            }

    Returns
    -------
    wb : object
        The workbook object.

    Raises
    ------
    ValueError
        If the workbook object is not able to be read by openpyxl.
    """
    # try to read the workbook with openpyxl
    # if the workbook is not able to be read by openpyxl, raise a value error
    try:
        wb = openpyxl.load_workbook(wb)
    except:
        raise ValueError("The workbook object wb is " +
        "not able to be read by openpyxl.")

    # get the list of links in the workbook
    current_links = [link.target for link in wb._external_links]

    # for each link in the dictionary of links
    for link in links:
        # if the link is in the list of links in the workbook
        if link in current_links:
            # get the index of the link in the list of links in the workbook
            index = current_links.index(link)
            # update the link in the list of links in the workbook
            current_links[index] = links[link]
        # if the link is not in the list of links in the workbook
        else:
            # pass a message to the user
            print(f"The link {link} is not in the workbook.")
            # continue to the next link
            continue

    # update the links in the workbook
    wb._external_links = current_links

    # return the workbook object
    return wb


def update_links(wb, links):
    """
    Description
    -----------
    Update the links in the workbook from
    the current links to the desired links.

    Parameters
    ----------
    wb : object
        The workbook object. Must be able to be read by openpyxl or pyxlsb.
    links : dict
        The dictionary with the current links as keys and the desired links as values.
        Dictionary should be of the form:
            {
                current_link1: desired_link1,
                current_link2: desired_link2,
                ...
            }

    Returns
    -------
    wb : object
        The workbook object.

    Imports
    -------
    openpyxl
    pyxlsb

    Raises
    ------
    ValueError
        If the workbook object is not able to be read by openpyxl or pyxlsb.

    Examples
    --------
    >>> import pyxlsb
    >>> wb = pyxlsb.open_workbook("test.xlsb")
    >>> links = {
        "C:\\Users\\test\\test2.xlsb": "C:\\Users\\test\\test3.xlsb"
    }
    >>> wb = update_links(wb, links)
    """
    # check to make sure the workbook can be read by openpyxl or pyxlsb
    # if the workbook is not able to be read by openpyxl or pyxlsb, raise a value error
    try:
        find_links(wb)
    except:
        raise ValueError("The workbook object wb is not able to be read by openpyxl or pyxlsb.")

    # if the workbook is an xlsb file
    if is_xlsb(wb):
        # use the update_links_pyxlsb function to update the links
        wb = update_links_pyxlsb(wb, links)
    # if the workbook is not an xlsb file
    else:
        # use the update_links_openpyxl function to update the links
        wb = update_links_openpyxl(wb, links)

    # return the workbook object
    return wb
