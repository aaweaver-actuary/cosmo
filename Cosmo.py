import datetime

from .src.is_xlsb import is_xlsb
from .src.open_workbook import open_workbook
from .src.get_named_ranges import get_named_ranges
from .src.get_links import get_links
from .src.save_workbook import save_workbook


# define a class to hold the data from the excel file
# this class will be used to update the excel file
# convention that methods that using underscores between words are used to retrieve data
# convention that methods that using camel case are used to update data/perform actions/change the workbook
class Cosmo:
    """
    Description
    -----------
    A class to hold the data from the excel file and to update the excel file.

    Creates a dicitonary to hold the cosmo macro as it is built. This cosmo
    macro can be saved to a json file using the SaveCosmoMacro method.

    Can load and run a cosmo macro from a .cosmomacro file.

    Logs the actions performed on the workbook to a cosmo_log list. This list
    can be saved to a json file using the SaveCosmoLog method.

    Parameters
    ----------
    workbook_file_path : str
        The file path of the workbook to open.

    Attributes
    ----------
    workbook_file_path : str
        The file path of the workbook.
    is_xlsb : bool
        Whether the workbook is an xlsb file or not.
    wb : openpyxl.Workbook or pyxlsb.Workbook
        The workbook object.
    book : openpyxl.Workbook or pyxlsb.Workbook
        Alias for wb.
    workbook_obj : openpyxl.Workbook or pyxlsb.Workbook
        Alias for wb.
    sheet_names : list
        The names of the sheets in the workbook.
    sheets : list
        Alias for sheet_names.
    worksheets : list
        Alias for sheet_names.
    tabs : list
        Alias for sheet_names.
    named_ranges : list
        The names of the named ranges in the workbook.
    links : list
        The links in the workbook.
    external_links : list
        Alias for links.
    linked_files : list
        Alias for links.
    cosmo_macro : dict
        A dictionary to hold the cosmo macro as it is built.
        This cosmo macro can be saved to a json file using 
        the SaveCosmoMacro method.
    cosmo_log : list
        A list to hold the actions performed on the workbook.
        This list can be saved to a json file using the SaveCosmoLog method.

    Methods
    -------
    ### Workbook update methods:
    Save
        Save the workbook.
        If the workbook is an xlsb file, then save it using pyxlsb.
        If the workbook is not an xlsb file, then save it using openpyxl.
        After saving the workbook, prints a message to the console
        with the file path of the saved workbook.


    ### Macro methods:
    SaveCosmoMacro
        Save the cosmo macro to a json file.
    LoadCosmoMacro
        Load a cosmo macro from a json file.
    RunCosmoMacro
        Run a cosmo macro.
    SaveCosmoLog
        Save the cosmo log to a json file.

    



    """
    def __init__(self, workbook_file_path):
        self.workbook_file_path = workbook_file_path

        # boolean for whether the workbook is .xlsb or not
        self.is_xlsb = is_xlsb(workbook_file_path)

        # open the workbook
        self.wb = open_workbook(workbook_file_path)

        # alias the wb to book, workbook_obj
        # this is to make it easier to remember the variable name
        self.book = self.wb
        self.workbook_obj = self.wb

        # get the sheet names
        self.sheet_names = self.wb.sheetnames

        # alias the sheet names to sheets, worksheets, tabs
        self.sheets = self.sheet_names
        self.worksheets = self.sheet_names
        self.tabs = self.sheet_names

        # get the named ranges
        self.named_ranges = get_named_ranges(self.wb)

        # get the links in the workbook
        self.links = get_links(self.wb)

        # a few aliases for the links
        self.external_links = self.links
        self.linked_files = self.links

        # initialize an empty dictionary to hold
        # the cosmo macro as it is built
        self.cosmo_macro = {}

        # initialize an empty list to hold the cosmo log
        self.cosmo_log = []

    # function to save the workbook
    def Save(self, is_copy=True, new_filename=None):
        """
        Description
        -----------
        Save the workbook.
        If the workbook is an xlsb file, then save it using pyxlsb.
        If the workbook is not an xlsb file, then save it using openpyxl.
        After saving the workbook, prints a message to the console
        with the file path of the saved workbook.
        Also logs the action to the cosmo log and updates the cosmo macro.

        Parameters
        ----------
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

        Notes
        -----
        If the workbook is an xlsb file,
        then the workbook is saved using pyxlsb.
        If the workbook is not an xlsb file,
        then the workbook is saved using openpyxl.

        Imports
        -------
        from .src.save_workbook import save_workbook

        Examples
        --------
        >>> # get current working directory
        >>> import os
        >>> os.getcwd()
        'C:\\Users\\username\\Documents\\Python Scripts\\Cosmo'

        >>> # save a copy of the workbook
        >>> cosmo.Save(is_copy=True)
        Saved workbook as
        C:\\Users\\username\\Documents\\Python Scripts\\Cosmo\\test_copy.xlsb
        """
        # save the workbook
        save_workbook(
            # the workbook object
            self.wb
            # whether the workbook is a copy of the original workbook
            , is_copy=is_copy
            # the new filename to save the workbook as
            # if None, then the workbook is saved with the original filename
            # if not None, then the workbook is saved with the new filename
            , new_filename=new_filename
            )
        ### DROP THE LOG FUNCTIONALITY FOR NOW
        # # log the action to the cosmo log
        # # first check if the cosmo log already has a save action anywhere
        # # in the log
        # # if it does, then prompt the user to overwrite the save action
        # # if it does not, then add the save action to the cosmo log
        # # and then add to the cosmo macro as well
        # if any(action['action'] == 'save' for action in self.cosmo_log):
        #     # notify the user that the cosmo log already has a save action 
        #     # and give the timestamp of the previous save action
        #     time_stamp = [action['timestamp'] for action in self.cosmo_log if action['action'] == 'save'][0]
        #     print('The cosmo log already has a save action.')
        #     print(f'Timestamp of previous save action: {time_stamp}')

        #     # prompt the user to overwrite the save action
        #     overwrite = input('Overwrite previous save action? (y/n): ')
        #     if overwrite.lower() == 'y':
        #         # remove the previous save action from the cosmo log
        #         self.cosmo_log = [action for action in self.cosmo_log if action['action'] != 'save']

        #         # define the save action
        #         save_action = {
        #             'action': 'save'
        #             , 'is_copy': is_copy
        #             , 'new_filename': new_filename
        #             , "timestamp": datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        #         }

        #         # add the save action to the cosmo log
        #         self.cosmo_log.append(save_action)

        #         # add the save action to the cosmo macro
        #         self.cosmo_macro['cosmo_log'] = self.cosmo_log
        #     else:
        #         # do not overwrite the previous save action
        #         pass

    # function to close the workbook
    def Close(self):
        """
        Close the workbook.

        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        eu.close_workbook(self.wb)

    # function to update links
    def UpdateLinks(self, links):
        """
        Description
        -----------
        Update the links in the workbook from the current links to the desired links, 
        where the current links are the keys and the desired links are the values of the dictionary
        that is passed to the function.

        Parameters
        ----------
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
        None
        """
        self.wb = eu.update_links(self.wb, links)

    # function to update the named ranges
    # takes a dictionary called named_ranges as input where the keys are the named ranges and the values are the new values
    # dictionary should be of the form:
    # {
    #     named_range1: new_value1,
    #     named_range2: new_value2,
    #     ...
    # }
    def UpdateNamedRanges(self, named_ranges):
        """
        Update the named ranges in the workbook to the new values, 
        where the name of a range is the keys and the new values are the values of the dictionary
        that is passed to the function.

        Parameters
        ----------
        named_ranges : dict
            The dictionary with the current values as keys and the desired values as values.
            Dictionary should be of the form:
                {
                    named_range1: new_value1,
                    named_range2: new_value2,
                    ...
                }

        Returns
        -------
        None
        """
        self.wb = eu.update_named_ranges(self.wb, named_ranges)

