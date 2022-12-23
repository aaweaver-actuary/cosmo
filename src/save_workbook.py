"""
save_workbook.py
"""
# do not test these imports with pylance
# pylance does not recognize the imports
# pylint: disable=E0401
# pylint: disable=E0611
import pyxlsb
import openpyxl
import datetime
from .is_xlsb import is_xlsb

# function that takes wb object as input and saves the workbook
# starts with extremely detailed docstring
# tests with is_xlsb function to determine if the workbook is an xlsb file
# if it is, then it saves the workbook as an xlsb file using pyxlsb
# if it is not, then it saves the workbook using openpyxl
# takes is_copy as input, which is a boolean that determines whether
# the workbook is a copy of the original workbook
# if it is a copy, then it saves the workbook to the original file path
# with the original file name plus a timestamp
# if it is not a copy, then it saves the workbook to the original file path
# with the original file name
# tests the wb object to determine if it is a wb object that pyxlsb or openpyxl
# can read
# if it is, then it saves the workbook
# if it is not, then it raises an error
# either way, prints a message to the console with the file path of the saved workbook
def save_workbook(wb, is_copy=True, new_filename=None):
    """
    Description
    -----------
    Save the workbook.
    If the workbook is an xlsb file, then save it using pyxlsb.
    If the workbook is not an xlsb file, then save it using openpyxl.


    Parameters
    ----------
    wb : openpyxl.Workbook or pyxlsb.Workbook
        The workbook to save.
        Must be a wb object that pyxlsb or openpyxl can read.
    is_copy : bool
        Whether the workbook is a copy of the original workbook.
        Default is True.
    new_filename : str
        The new filename to save the workbook as.
        Default is None.
        If None, then the workbook is saved with the original filename.
        If not None, then the workbook is saved with the new filename.

    Returns
    -------
    None

    Raises
    ------
    TypeError
        If the workbook is not a wb object that pyxlsb or openpyxl can read.

    Notes
    -----
    If the workbook is a copy, then it saves the workbook to the original file path
    with the original file name plus a timestamp.
    If the workbook is not a copy, then it saves the workbook to the original file path
    with the original file name.


    Imports
    -------
    pyxlsb
    openpyxl
    datetime
    .is_xlsb


    Examples
    --------
    >>> import pyxlsb
    >>> wb = pyxlsb.open_workbook('test.xlsb')
    >>> wb.file
    'test.xlsb'
    >>> wb.original_file_path
    'C:\\Users\\username\\Documents\\'
    >>> save_workbook(wb, is_copy=True)
    Saved workbook to C:\\Users\\username\\Documents\\test_2021-08-01_12-00-00.xlsb

    >>> import openpyxl
    >>> wb = openpyxl.Workbook('./test.xlsx')
    >>> wb.file
    'test.xlsx'
    >>> wb.original_file_path
    'C:\\Users\\username\\Documents\\'
    >>> save_workbook(wb, is_copy=True)
    Saved workbook to C:\\Users\\username\\Documents\\test_2021-08-01_12-00-00.xlsx

    >>> wb = 'test'
    >>> save_workbook(wb)
    Traceback (most recent call last):
    ...
    TypeError: The workbook is not a wb object that pyxlsb or openpyxl can read.

    >>> wb = pyxlsb.open_workbook('test.xlsb')
    >>> wb.original_file_path
    'C:\\Users\\username\\Documents\\'
    >>> save_workbook(wb, is_copy=True, new_filename='test2.xlsb')
    Saved workbook to C:\\Users\\username\\Documents\\test2.xlsb
    """
    # test if the workbook is an xlsb file
    if is_xlsb(wb):
        # if is_copy is True,
        # make a copy of the file and save it to the original file path
        # with a new filename
        if is_copy:
            # test whether or not new_filename is None
            # if it is, then then save the workbook to the original file path
            # with the original file name plus a timestamp as the new filename
            # if it is not, then use the new_filename as the new filename
            # with no timestamp
            if new_filename is None:
                wb.save(wb.original_file_path + wb.file + str(datetime.now()))
            else:
                # save a copy of the workbook to the original file path
                # with the new filename
                wb.save(wb.original_file_path + new_filename)
        # if is_copy is False,
        # save the workbook to the original file path
        # with the original file name
        else:
            wb.save(wb.original_file_path)

    # if the workbook is not an xlsb file
    else:
        # if is_copy is True,
        # make a copy of the file and save it to the original file path
        # with a new filename
        if is_copy:
            # test whether or not new_filename is None
            # if it is, then then save the workbook to the original file path
            # with the original file name plus a timestamp as the new filename
            # if it is not, then use the new_filename as the new filename
            # with no timestamp
            # this uses openpyxl instead of pyxlsb
            if new_filename is None:
                wb.save(wb.original_file_path + wb.file + str(datetime.now()))
            else:
                # save a copy of the workbook to the original file path
                # with the new filename
                wb.save(wb.original_file_path + new_filename)
        # if is_copy is False,
        # save the workbook to the original file path
        # with the original file name
        else:
            wb.save(wb.original_file_path)

    # print a message to the console with the file path of the saved workbook
    print("Workbook saved to: " + wb.original_file_path)
