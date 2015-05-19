# Copyright (C) 2003 - 2006 The Board of Regents of the University of Wisconsin System 
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of version 2 of the GNU General Public License as
# published by the Free Software Foundation.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
#

"""This file implements the Filtering Dialog Box for the Transana application."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "FilterDialog DEBUG is ON!!"

# Import wxPython
import wx
# Import wxPython's CheckListCtrl Mixin
from wx.lib.mixins.listctrl import CheckListCtrlMixin
# import Python's cPickle module
import cPickle
# import Python's os and sys module
import os, sys
# Import Transana's Collection Object (for the Clip List)
import Collection
# Import Transana's DBInterface
import DBInterface
# import Transana's Dialogs module for the ErrorDialog.
import Dialogs
# import Transana's Miscellaneous Functions
import Misc
# Import Transana's Globals
import TransanaGlobal

# Declare Constants for the Toolbar Button IDs
T_FILE_OPEN    =  wx.NewId()
T_FILE_SAVE    =  wx.NewId()
T_FILE_DELETE  =  wx.NewId()
T_CHECK_ALL    =  wx.NewId()
T_CHECK_NONE   =  wx.NewId()
T_HELP_HELP    =  wx.NewId()
T_FILE_EXIT    =  wx.NewId()

class CheckListCtrl(wx.ListCtrl, CheckListCtrlMixin):
    """ This class turns a normal ListCtrl into a CheckListCtrl. """
    def __init__(self, parent):
        # Create a ListCtrl in Report View that only allows single selection
        wx.ListCtrl.__init__(self, parent, -1, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        # Make it a CheckList using the CheckListCtrlMixin
        CheckListCtrlMixin.__init__(self)

class FilterDialog(wx.Dialog):
    """ This window implements Episode, Clip, and Keyword filtering for Transana Reports.
        Required parameters are:
          parent
          id            (should be -1)
          title
          reportType     1 = Keyword Map  (requires episodeNum)
        Optional parameters are:
          configName    (current Configuration Name)
          episodeNum    (required for Keyword Map Filter)
          episodeFilter (boolean)
          episodeSort   (boolean)
          clipFilter    (boolean)
          clipSort      (boolean)
          keywordFilter (boolean)
          keywordSort   (boolean)
          timeRange     (boolean)
          startTime     (number of milliseconds)
          endTime       (number of milliseconds) """
    def __init__(self, parent, id, title, reportType, **kwargs):
        """ Initialize the Transana Filter Dialog Box """
        # Create a Dialog Box
        wx.Dialog.__init__(self, parent, id, title, size = (500,500),
                           style=wx.CAPTION | wx.CLOSE_BOX | wx.RESIZE_BORDER | wx.NO_FULL_REPAINT_ON_RESIZE)
        # Make the background White
        self.SetBackgroundColour(wx.WHITE)

        # Remember the report type
        self.reportType = reportType
        # Remember the keyword arguments
        self.kwargs = kwargs
        # Initialize the Report Configuration Name if one is passed in, if not, initialize to an empty string
        if kwargs.has_key('configName'):
            self.configName = kwargs['configName']
        else:
            self.configName = ''

        if kwargs.has_key('startTime') and kwargs['startTime']:
            self.startTime = Misc.time_in_ms_to_str(kwargs['startTime'])
        else:
            self.startTime = Misc.time_in_ms_to_str(0)
        if kwargs.has_key('endTime') and kwargs['endTime']:
            self.endTime = Misc.time_in_ms_to_str(kwargs['endTime'])
        else:
            self.endTime = Misc.time_in_ms_to_str(parent.MediaLength)

        # Create BoxSizers for the Dialog
        vBox = wx.BoxSizer(wx.VERTICAL)
        hBox = wx.BoxSizer(wx.HORIZONTAL)

        # Add the Tool Bar
        self.toolBar = wx.ToolBar(self, -1, style=wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_TEXT)
        # Get the image for File Open
        bmp = wx.ArtProvider_GetBitmap(wx.ART_FILE_OPEN, wx.ART_TOOLBAR, (16,16))
        # Create the File Open button
        btnFileOpen = self.toolBar.AddTool(T_FILE_OPEN, bmp, shortHelpString=_("Load Filter Configuration"))
        # Create the File Save button
        btnFileSave = self.toolBar.AddTool(T_FILE_SAVE, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Save16.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_("Save Filter Configuration"))
        # Get the image for File Open
        bmp = wx.ArtProvider_GetBitmap(wx.ART_DELETE, wx.ART_TOOLBAR, (16,16))
        # Create the Config Delete button
        btnFileDelete = self.toolBar.AddTool(T_FILE_DELETE, bmp, shortHelpString=_("Delete Filter Configuration"))
        # Create the Check All button
        btnCheckAll = self.toolBar.AddTool(T_CHECK_ALL, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Check.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Check All'))
        # Create the Uncheck All button
        btnCheckNone = self.toolBar.AddTool(T_CHECK_NONE, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "NoCheck.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Uncheck All'))
        # Get the graphic for Help
        bmp = wx.ArtProvider_GetBitmap(wx.ART_HELP, wx.ART_TOOLBAR, (16,16))
        # create the Help button
        btnHelp = self.toolBar.AddTool(T_HELP_HELP, bmp, shortHelpString=_("Help"))
        # Create the Exit button
        btnExit = self.toolBar.AddTool(T_FILE_EXIT, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Exit.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Exit'))
        # Realize the Toolbar
        self.toolBar.Realize()
        # Bind events to the Toolbar Buttons
        self.Bind(wx.EVT_MENU, self.OnFileOpen, btnFileOpen)
        self.Bind(wx.EVT_MENU, self.OnFileSave, btnFileSave)
        self.Bind(wx.EVT_MENU, self.OnFileDelete, btnFileDelete)
        self.Bind(wx.EVT_MENU, self.OnCheckAll, btnCheckAll)
        self.Bind(wx.EVT_MENU, self.OnCheckAll, btnCheckNone)
        self.Bind(wx.EVT_MENU, self.OnHelp, btnHelp)
        self.Bind(wx.EVT_MENU, self.OnClose, btnExit)
        # Add a spacer to the Sizer that allows for the Toolbar
        vBox.Add((1,26))

        # Everything in this Dialog goes inside a Notebook.  Create that notebook control
        self.notebook = wx.Notebook(self, -1)

        # If Episode Filtering is requested ...
        if kwargs.has_key('episodeFilter') and kwargs['episodeFilter']:
            # ... build a Panel for Episodes ...
            self.episodesPanel = wx.Panel(self.notebook, -1)  # , style=wx.WANTS_CHARS
            # ... place the Episode Panel on the Notebook, creating an Episodes tab ...
            self.notebook.AddPage(self.episodesPanel, _("Episodes"))
            # ... 
            epTxt = wx.StaticText(self.episodesPanel, -1, "Episode Filtering not yet implemented.")

            # If Episode Sorting capacity has been requested ...
            if kwargs.has_key('episodeSort') and kwargs['episodeSort']:
                print "Episode Sorting has not yet been implemented."
                
        # If Clip Filtering is requested ...
        if kwargs.has_key('clipFilter') and kwargs['clipFilter']:
            # ... build a Panel for Clips ...
            self.clipsPanel = wx.Panel(self.notebook, -1)  # , style=wx.WANTS_CHARS
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Clips Panel on the Notebook, creating a Clips tab ...
            self.notebook.AddPage(self.clipsPanel, _("Clips"))
            # ... place a Check List Ctrl on the Clips Panel ...
            self.clipList = CheckListCtrl(self.clipsPanel)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.clipList, 1, wx.EXPAND)
            # The clip List needs two columns, Clip ID and Collection nesting.
            self.clipList.InsertColumn(0, _("Clip ID"))
            self.clipList.InsertColumn(1, _("Collection ID(s)"))

            # If Clip Sorting capacity has been requested ...
            if kwargs.has_key('clipSort') and kwargs['clipSort']:
                print "Clip Sorting has not yet been implemented."
                # self.originalClipData will need to be reordered the same way that self.clipList is.
                # Otherwise, GetClips won't work right!

            # NOTE:  The actual list of Clips needs to be provided by the calling routine using
            #        the SetClips method.  This is because only the calling routine knows which
            #        Clips are legal and should be included in the list.  We don't necessarily
            #        want all Clips all the time.  The Clip List should be in the form of an
            #        ordered list of tuples made up of (collectionNum, clipID, checked) information.

            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.clipsPanel.SetSizer(pnlVSizer)

        # If Keyword Filtering is requested ...
        if kwargs.has_key('keywordFilter') and kwargs['keywordFilter']:
            # ... build a Panel for Keywords ...
            self.keywordsPanel = wx.Panel(self.notebook, -1)  # , style=wx.WANTS_CHARS
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Keywords Panel on the Notebook, creating a Keywords tab ...
            self.notebook.AddPage(self.keywordsPanel, _("Keywords"))
            # ... place a Check List Ctrl on the Keywords Panel ...
            self.keywordList = CheckListCtrl(self.keywordsPanel)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.keywordList, 1, wx.EXPAND)
            # The keyword List needs two columns, Keyword Group and Keyword.
            self.keywordList.InsertColumn(0, _("Keyword Group"))
            self.keywordList.InsertColumn(1, _("Keyword"))

            # If Keyword Sorting capacity has been requested ...
            if kwargs.has_key('keywordSort') and kwargs['keywordSort']:
                # create a vertical sizer to hold the sort buttons
                pnlBtnSizer = wx.BoxSizer(wx.VERTICAL)
                # get the graphic for the Up arrow
                bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_UP, wx.ART_TOOLBAR, (16,16))
                # create a bitmap button for the Move Up button
                self.btnKwUp = wx.BitmapButton(self.keywordsPanel, -1, bmp)
                # Set the Tool Tip for the Move Up button
                self.btnKwUp.SetToolTipString(_("Move keyword up"))
                # Bind the button event to a method
                self.btnKwUp.Bind(wx.EVT_BUTTON, self.OnButton)
                # Insert some expandable space into the button sizer at the top
                pnlBtnSizer.Add((1,1), 1, wx.EXPAND)
                # Add the Move Up button to the button sizer
                pnlBtnSizer.Add(self.btnKwUp, 0, wx.ALIGN_CENTER | wx.ALL, 5)
                # Add a spacer to the button sizer to increase the amount of space between the buttons
                pnlBtnSizer.Add((1, 10))
                # Get the graphic for the Down arrow
                bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_DOWN, wx.ART_TOOLBAR, (16,16))
                # create a bitmap button for the Move Down button
                self.btnKwDown = wx.BitmapButton(self.keywordsPanel, -1, bmp)
                # Set the Tool Tip for the Move Down button
                self.btnKwDown.SetToolTipString(_("Move keyword down"))
                # Bind the button event to a method
                self.btnKwDown.Bind(wx.EVT_BUTTON, self.OnButton)
                # Add the Move Down button to the button sizer
                pnlBtnSizer.Add(self.btnKwDown, 0, wx.ALIGN_CENTER | wx.ALL, 5)
                # Add some expandable space below the button to keep the buttons centered vertically
                pnlBtnSizer.Add((1,1), 1, wx.EXPAND)
                # Add the button sizer to the panel's Horizontal sizer
                pnlHSizer.Add(pnlBtnSizer, 0, wx.EXPAND)
                
            # NOTE:  The actual list of Keywords needs to be provided by the calling routine using
            #        the SetKeywords method.  This is because only the calling routine knows which
            #        Keywords are legal and should be included in the list.  We don't necessarily
            #        want all Keywords all the time.  The Keyword List should be in the form of an
            #        ordered list of tuples made up of (keywordgroup, keyword, checked) information.

            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.keywordsPanel.SetSizer(pnlVSizer)

        # If Time Range selection is requested ...
        if kwargs.has_key('timeRange') and kwargs['timeRange']:
            # ... build a Panel for Time Range ...
            self.timeRangePanel = wx.Panel(self.notebook, -1)  # , style=wx.WANTS_CHARS
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            # ... place the Time Range Panel on the Notebook, creating a Time Range tab ...
            self.notebook.AddPage(self.timeRangePanel, _("Time Range"))
            # Add a label for the Start Time field
            startTimeTxt = wx.StaticText(self.timeRangePanel, -1, _("Start Time"))
            pnlVSizer.Add(startTimeTxt, 0, wx.TOP | wx.LEFT, 10)
            # Add the Start Time field
            self.startTime = wx.TextCtrl(self.timeRangePanel, -1, self.startTime)
            pnlVSizer.Add(self.startTime, 0, wx.LEFT, 10)
            # Add a label for the End Time field
            endTimeTxt = wx.StaticText(self.timeRangePanel, -1, _("End Time"))
            pnlVSizer.Add(endTimeTxt, 0, wx.TOP | wx.LEFT, 10)
            # Add the End Time field
            self.endTime = wx.TextCtrl(self.timeRangePanel, -1, self.endTime)
            pnlVSizer.Add(self.endTime, 0, wx.LEFT, 10)
            # Add a note that says this data does not get saved.
            tRTxt = wx.StaticText(self.timeRangePanel, -1, _("NOTE:  Setting the End Time to 0 will set it to the end of the Media File.\nTime Range data is not saved as part of the Filter Configuration data."))
            pnlVSizer.Add(tRTxt, 0, wx.ALL, 10)

            # Now declare the panel's vertical sizer as the panel's official sizer
            self.timeRangePanel.SetSizer(pnlVSizer)

        # Place the Notebook in the Dialog Horizontal Sizer
        hBox.Add(self.notebook, 1, wx.EXPAND, 10)
        # Now place the Dialog's horizontal Sizer in the Dialog's Vertical sizer so we can expand in two dimensions
        vBox.Add(hBox, 1, wx.EXPAND, 10)

        # Create another horizontal sizer for the Dialog's buttons
        btnBox = wx.BoxSizer(wx.HORIZONTAL)
        # Create the OK button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Make this the default button
        btnOK.SetDefault()
        # Add the OK button to the dialog's Button sizer
        btnBox.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT, 10)
        # Create the Cancel button
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        # Add the Cancel button to the dialog's Button sizer
        btnBox.Add(btnCancel, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT, 10)
        # Create the Help button
        btnHelp = wx.Button(self, -1, _("Help"))
        # Add the Help Button to the Dialog's Button sizer
        btnBox.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT, 10)
        # Bind the Help button to the appropriate method
        btnHelp.Bind(wx.EVT_BUTTON, self.OnHelp)
        # Now add the dialog's button sizer to the dialog's vertical sizer
        vBox.Add(btnBox, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM)
        # Declare the dialog's vertical sizer as the dialog's official sizer
        self.SetSizer(vBox)
        # Tell the dialog to auto-layout
        self.SetAutoLayout(True)
        # NOTE:  The calling routine must provide additional data via the SetEpisodes, SetClips, and / or SetKeywords methods
        #        before this dialog is displayed.  Therefore, we don't complete the Layout or call ShowModal here.

    def OnClose(self, event):
        """ Close the Filter Dialog """
        self.Close()

    def GetReportScope(self):
        """ Different Report versions are based on different originating objects.  This method returns the
            correct record number (that is, the Report's Scope) for the originating object based on the Report Type. """
        # Initialize Report Scope to None
        reportScope = None
        # If we are dealing with a Keyword Map Report ...
        if self.reportType == 1:
            # ... check to see that an episodeNum parameter was passed ...
            if self.kwargs.has_key('episodeNum'):
                # ... and return the Episode Number as the Report's Scope.
                reportScope = self.kwargs['episodeNum']
        # Return the Report Scope as the function result
        return reportScope

    def GetConfigNames(self):
        """ Get a list of Configuration Names for the current Report Type and Report Scope. """
        # Create a blank list
        resList = []
        # Get the Report's Scope
        reportScope = self.GetReportScope()
        # Make sure we have a legal Report
        if reportScope != None:
            # Get a Database Cursor
            DBCursor = DBInterface.get_db().cursor()
            # Build the database query
            query = """ SELECT ConfigName FROM Filters2
                          WHERE ReportType = %s AND
                                ReportScope = %s
                          GROUP BY ConfigName
                          ORDER BY ConfigName """
            # Set up the data values that match the query
            values = (self.reportType, reportScope)
            # Execute the query with the data values
            DBCursor.execute(query, values)
            # Iterate through the report results
            for (configName,) in DBCursor.fetchall():
                # Decode the data
                configName = DBInterface.ProcessDBDataForUTF8Encoding(configName)
                # Add the data to the Results list
                resList.append(configName)
        # return the Results List
        return resList

    def ReconcileLists(self, formList, fileList, listIsOrdered=False):
        """ Take lists from the form and from the database and reconcile them, returning the results  """
        # We need to compare the form data to the file data and reconcile differences.
        #   1.  detect items that are in the file that are not in the form and drop them
        #   2.  detect items that are in the form that are not in the file.  If the list is
        #       ordered, we can add them to the bottom.  Otherwise, we need to insert them
        #       in place.  Either way, we make sure they are checked!
        #
        # As long as the list item tuples match and the last item in each tuple is the "checked" item,
        # we don't need to know the contents of the list.
        #

        # Create an empty list for the results
        resList = []

        # Iterate through the File List, which should control initial order and checked status
        for items in fileList:
            # Get all items except the last one
            mostItems = items[:-1]
            # Check to see if the file list items are in the form list, checked as True
            if (mostItems + (True,) in formList):
                resList.append(items)
                formList.remove(mostItems + (True,))
                
            # Check to see if the file list items are in the form list, checked as False
            elif (mostItems + (False,) in formList):
                resList.append(items)
                formList.remove(mostItems + (False,))

            elif DEBUG:
                print "In File, not Form:", items

        # If the file list item doesn't appear in the form list, it is NOT placed in
        # the results list.

        # All that should be left in the Form List after the File List items have been removed
        # are items that existed in the form that didn't exist in the file.  These items should
        # be appended to the end of the list.

        # First, let's figure out how to place items in an unordered list.
        # For the moment, we'll look at the first column of the first item.  If it's an integer,
        # we'll use the second column.  Otherwise, we can use the first column.
        if (len(formList) > 0) and (isinstance(formList[0][0], long)):
            lookupCol = 1
        else:
            lookupCol = 0

        # Iterate through the items left in the Form List, the ones left after processing the File list.
        for items in formList:
            # If a item is in the form but not in the file, check to see if it is unchecked.
            if items[-1] == False:
                # If it is not checked, we need to rebuild the tuple to ensure it is checked when we're done!
                items = items[:-1] + (True,)
            # If we have a user-ordered list ...
            if listIsOrdered:
                # ... just append the items to the end of the list.
                resList.append(items)
            # If we have a list that is not user-ordered, it is alphabetical.  We need to figure out where to place
            # the items that need to be added.
            else:
                # Indicate that the item has not been placed yet.
                placed = False
                # Iterate through the Results List ...
                for x in range(len(resList)):
                    # ... looking for where the current formList item should be located.
                    if items[lookupCol] < resList[x][lookupCol]:
                        # Insert the Form List item into the Results List ...
                        resList.insert(x, items)
                        # ... indicate that the item WAS placed ...
                        placed = True
                        # ... and stop looking.
                        break
                # If the item was never placed, ...  
                if not placed:
                    # ... add it to the end of the Results List.
                    resList.append(items)

            if DEBUG:
                print "In Form, not File:", items

        return resList

    def OnFileOpen(self, event):
        """ Load a Filter Configuration appropriate to the current Report specifications """
        # Get the current Report's Scope
        reportScope = self.GetReportScope()
        # Make sure we have a legal Report
        if reportScope != None:
            # Get a list of legal Report Names from the Database
            configNames = self.GetConfigNames()
            # Create a Choice Dialog so the user can select the Configuration they want
            dlg = wx.SingleChoiceDialog(self, _("Choose a Configuration to load"), _("Filter Configuration"), configNames, wx.CHOICEDLG_STYLE)
            # Center the Dialog on the screen
            dlg.CentreOnScreen()
            # Show the Choice Dialog and see if the user chooses OK
            if dlg.ShowModal() == wx.ID_OK:
                # Remember the Configuration Name
                self.configName = dlg.GetStringSelection()
                # Get a Database Cursor
                DBCursor = DBInterface.get_db().cursor()
                # Build a query to get the selected report's data
                query = """ SELECT FilterDataType, FilterData FROM Filters2
                              WHERE ReportType = %s AND
                                    ReportScope = %s AND
                                    ConfigName = %s """
                # Build the data values that match the query
                values = (self.reportType, reportScope, self.configName)
                # Execute the query with the appropriate data values
                DBCursor.execute(query, values)
                # We may get multiple records, one for each tab on the Filter Dialog.
                for (filterDataType, filterData) in DBCursor.fetchall():
                    # If the data is for the Clips Tab (filterDataType 2) ...
                    if filterDataType == 2:
                        # Get the current Clip data from the Form
                        formClipData = self.GetClips()
                        # Get the Clip data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileClipData = cPickle.loads(filterData.tostring())
                        else:
                            fileClipData = cPickle.loads(filterData)
                        # Clear the Clip List
                        self.clipList.DeleteAllItems()
                        # Determine if this list is Ordered
                        if self.kwargs.has_key('clipSort') and self.kwargs['clipSort']:
                            orderedList = True
                        else:
                            orderedList = False
                        # We need to compare the file data to the form data and reconcile differences,
                        # then feed the results to the Clip Tab.
                        self.SetClips(self.ReconcileLists(formClipData, fileClipData, listIsOrdered=orderedList))
                        
                    # If the data is for the Keywords Tab (filterDataType 3) ...
                    elif filterDataType == 3:
                        # Get the current Keyword data from the Form
                        formKeywordData = self.GetKeywords()
                        # Get the Keyword data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileKeywordData = cPickle.loads(filterData.tostring())
                        else:
                            fileKeywordData = cPickle.loads(filterData)
                        # Clear the Keyword List
                        self.keywordList.DeleteAllItems()
                        # Determine if this list is Ordered
                        if self.kwargs.has_key('keywordSort') and self.kwargs['keywordSort']:
                            orderedList = True
                        else:
                            orderedList = False
                        # We need to compare the file data to the form data and reconcile differences,
                        # then feed the results to the Keyword Tab.
                        self.SetKeywords(self.ReconcileLists(formKeywordData, fileKeywordData, listIsOrdered=orderedList))

            # Destroy the Choice Dialog
            dlg.Destroy()
        
    def OnFileSave(self, event):
        """ Save a Configuration to the Filter table """
        # Remember the original Configuration Name
        originalConfigName = self.configName
        # Set a bogus error message to get us into the While loop
        errorMsg = "Begin"
        # Keep trying as long as there are error messages!
        while errorMsg != '':
            # Clear the error message, assuming the user will succeed in loading a configuration
            errorMsg = ''
            # Create a Dialog where the Configuration can be named
            dlg = wx.TextEntryDialog(self, _("Save Configuration As"), _("Filter Configuration"), originalConfigName)
            # Show the dialog to the user and see how they respond
            if dlg.ShowModal() == wx.ID_OK:
                # Remove leading and trailing white space
                configName = dlg.GetValue().strip()
                # Get a list of existing Configuration Names from the database
                configNames = self.GetConfigNames()
                # If the user changed the Configuration Name and the new name already exists ...
                if (configName != originalConfigName) and (configName in configNames):
                    # Build and encode the error prompt
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('A configuration named "%s" already exists.  Do you want to replace it?'), 'utf8')
                    else:
                        prompt = _('A configuration named "%s" already exists.  Do you want to replace it?')
                    # Build a dialog to notify the user of the duplication and ask what to do
                    dlg2 = wx.MessageDialog(self, prompt % configName, _('Transana Confirmation'),
                                            style = wx.YES_NO | wx.ICON_QUESTION | wx.STAY_ON_TOP)
                    # Center the dialog on the screen
                    dlg2.CentreOnScreen()
                    # Display the dialog and get the user's response
                    if dlg2.ShowModal() != wx.ID_YES:
                        # If the user doesn't say to replace, create an error message (never displayed) to signal there's a problem.
                        errorMsg = 'Duplicate Filename'
                    # Clean up after prompting the user for feedback
                    dlg2.Destroy()
                # To proceed, we need a report name and we need the error message to still be blank
                if (configName != '') and (errorMsg == ''):
                    # Get the current report's Scope
                    reportScope = self.GetReportScope()
                    # If we have a legal report
                    if reportScope != None:
                        # update the Configuration Name
                        self.configName = configName
                        # Get a database cursor
                        DBCursor = DBInterface.get_db().cursor()
                        # Encode the Report Name if we're using Unicode
                        if 'unicode' in wx.PlatformInfo:
                            configName = self.configName.encode(TransanaGlobal.encoding)
                        else:
                            configName = self.configName

                        # Check to see if the Configuration record already exists
                        if DBInterface.record_match_count('Filters2',
                                                         ('ReportType', 'ReportScope', 'ConfigName'),
                                                         (self.reportType, reportScope, configName)) > 0:
                            # Update existing record.  Note that each report may generate up to 3 records in the database,
                            # FilterDataType 1 = Episodes, FilterDataType 2 = Clips, FilterDataType 3 = Keywords
                            # If we have a report that's not yet implemented that needs it, update Episode Data (FilterDataType 1)
                            if self.reportType in []:
                                print "Update Episodes not implemented"
                            # If we have a Keyword Map (reportType 1), update Clip Data (FilterDataType 2)
                            if self.reportType in [1]:
                                # Build the Update Query for Clip Data
                                query = """ UPDATE Filters2
                                              SET FilterData = %s
                                              WHERE ReportType = %s AND
                                                    ReportScope = %s AND
                                                    ConfigName = %s AND
                                                    FilterDataType = %s """
                                # Pickle the Clip Data
                                clips = cPickle.dumps(self.GetClips())
                                # Build the values to match the query, including the pickled Clip data
                                values = (clips, self.reportType, reportScope, configName, 2)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Keyword Map (reportType 1), update Keyword Data (FilterDataType 3)
                            if self.reportType in [1]:
                                # Build the Update Query for Keyword Data
                                query = """ UPDATE Filters2
                                              SET FilterData = %s
                                              WHERE ReportType = %s AND
                                                    ReportScope = %s AND
                                                    ConfigName = %s AND
                                                    FilterDataType = %s """
                                # Pickle the Keyword Data
                                keywords = cPickle.dumps(self.GetKeywords())
                                # Build the values to match the query, including the pickled Keyword data
                                values = (keywords, self.reportType, reportScope, configName, 3)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                        else:
                            # Insert new record.  Note that each report may generate up to 3 records in the database,
                            # FilterDataType 1 = Episodes, FilterDataType 2 = Clips, FilterDataType 3 = Keywords
                            # If we have a report that's not yet implemented that needs it, insert Episode Data (FilterDataType 1)
                            if self.reportType in []:
                                print "Insert Episodes not implemented"
                            # If we have a Keyword Map (reportType 1), insert Clip Data (FilterDataType 2)
                            if self.reportType in [1]:
                                # Build the Insert Query for Clip Data
                                query = """ INSERT INTO Filters2
                                                (ReportType, ReportScope, ConfigName, FilterDataType, FilterData)
                                              VALUES
                                                (%s, %s, %s, %s, %s) """
                                # Pickle the Clip Data
                                clips = cPickle.dumps(self.GetClips())
                                # Build the values to match the query, including the pickled Clip data
                                values = (self.reportType, reportScope, configName, 2, clips)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Keyword Map (reportType 1), insert Keyword Data (FilterDataType 3)
                            if self.reportType in [1]:
                                # Build the Insert Query for Keyword Data
                                query = """ INSERT INTO Filters2
                                                (ReportType, ReportScope, ConfigName, FilterDataType, FilterData)
                                              VALUES
                                                (%s, %s, %s, %s, %s) """
                                # Pickle the Keyword Data
                                keywords = cPickle.dumps(self.GetKeywords())
                                # Build the values to match the query, including the pickled Keyword data
                                values = (self.reportType, reportScope, configName, 3, keywords)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                # If we can't proceed with the save ...
                else:
                    # If there's already an error message, we DON'T show it.
                    if errorMsg == '':
                        # If there's NOT an error messge, the User failed to enter a Report Name.
                        # Create an appropriate error message
                        errorMsg = _("You must enter a Configuration Name to save this filter configuration.")
                        # Build an error dialog
                        dlg2 = Dialogs.ErrorDialog(self, errorMsg)
                        # Display the error message
                        dlg2.ShowModal()
                        # Destroy the error dialog.
                        dlg2.Destroy()
            # Destroy the Dialog that allows the user to name the configuration
            dlg.Destroy()

    def OnFileDelete(self, event):
        """ Delete a Filter Configuration appropriate to the current Report specifications """
        # Get the current Report's Scope
        reportScope = self.GetReportScope()
        # Make sure we have a legal Report
        if reportScope != None:
            # Get a list of legal Report Names from the Database
            configNames = self.GetConfigNames()
            # Create a Choice Dialog so the user can select the Configuration they want
            dlg = wx.SingleChoiceDialog(self, _("Choose a Configuration to delete"), _("Filter Configuration"), configNames, wx.CHOICEDLG_STYLE)
            # Center the Dialog on the screen
            dlg.CentreOnScreen()
            # Show the Choice Dialog and see if the user chooses OK
            if dlg.ShowModal() == wx.ID_OK:
                # Remember the Configuration Name
                localConfigName = dlg.GetStringSelection()
                # Better confirm this.
                dlg2 = wx.MessageDialog(self, _('Are you sure you want to delete Filter Configuration "%s"?') % self.configName,
                                        _("Transana Confirmation"), wx.YES_NO | wx.ICON_INFORMATION)
                if dlg2.ShowModal() == wx.ID_YES:
                    # Clear the global configuration name
                    self.configName = ''
                    # Get a Database Cursor
                    DBCursor = DBInterface.get_db().cursor()
                    # Build a query to delete the selected report
                    query = """ DELETE FROM Filters2
                                  WHERE ReportType = %s AND
                                        ReportScope = %s AND
                                        ConfigName = %s """
                    # Build the data values that match the query
                    values = (self.reportType, reportScope, localConfigName)
                    # Execute the query with the appropriate data values
                    DBCursor.execute(query, values)
                dlg2.Destroy()
            # Destroy the Choice Dialog
            dlg.Destroy()
        
    def OnCheckAll(self, event):
        # Remember the ID of the button that triggered this event
        btnID = event.GetId()
        # Determine which Tab is currently showing
        selectedTab = self.notebook.GetPageText(self.notebook.GetSelection())
        # If we're looking at the Episodes tab ...
        if selectedTab == _("Episodes"):
            print "FilterDialog.OnCheckAll() does not work on the Episodes Tab at this time."
        # if we're looking at the Clips tab ...
        elif selectedTab == _("Clips"):
            # ... iterate through the Clips in the Clip List
            for x in range(self.clipList.GetItemCount()):
                # If the clip List Item's checked status does not match the desired status ...
                if self.clipList.IsChecked(x) != (btnID == T_CHECK_ALL):
                    # ... then toggle the item so it will match
                    self.clipList.ToggleItem(x)
        # if we're looking at the Keywords tab ...
        elif selectedTab == _("Keywords"):
            # ... iterate through the Keywords in the Keyword List
            for x in range(self.keywordList.GetItemCount()):
                # If the keyword List Item's checked status does not match the desired status ...
                if self.keywordList.IsChecked(x) != (btnID == T_CHECK_ALL):
                    # ... then toggle the item so it will match
                    self.keywordList.ToggleItem(x)

    def OnButton(self, event):
        """ Process Button Events for the Filter Dialog """
        # Get the ID of the button that triggered this event
        btnID = event.GetId()
        
        # If the Move Keyword Up button is pressed ...
        if btnID == self.btnKwUp.GetId():
            # Determine which Keyword is selected (-1 is None)
            item = self.keywordList.GetNextItem(-1, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
            # If a Keyword is selected ...
            if item > -1:
                # Extract the data from the Keyword List
                kwg = self.keywordList.GetItem(item, 0).GetText()
                kw = self.keywordList.GetItem(item, 1).GetText()
                checked = self.keywordList.IsChecked(item)
                # Check to make sure there's room to move up
                if item > 0:
                    # Delete the current item
                    self.keywordList.DeleteItem(item)
                    # Insert a new item one position up and add the extracted data
                    self.keywordList.InsertStringItem(item - 1, kwg)
                    self.keywordList.SetStringItem(item - 1, 1, kw)
                    if checked:
                        self.keywordList.CheckItem(item - 1)
                    # Make sure the new item is selected
                    self.keywordList.SetItemState(item - 1, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)
                    # Make sure the item is visible
                    self.keywordList.EnsureVisible(item - 1)

        # If the Move Keyword Down button is pressed ...
        elif btnID == self.btnKwDown.GetId():
            # Determine which Keyword is selected (-1 is None)
            item = self.keywordList.GetNextItem(-1, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
            # If a Keyword is selected ...
            if item > -1:
                # Extract the data from the Keyword List
                kwg = self.keywordList.GetItem(item, 0).GetText()
                kw = self.keywordList.GetItem(item, 1).GetText()
                checked = self.keywordList.IsChecked(item)
                # Check to make sure there's room to move down
                if item < self.keywordList.GetItemCount() - 1:
                    # Delete the current item
                    self.keywordList.DeleteItem(item)
                    # Insert a new item one position down and add the extracted data
                    self.keywordList.InsertStringItem(item + 1, kwg)
                    self.keywordList.SetStringItem(item + 1, 1, kw)
                    if checked:
                        self.keywordList.CheckItem(item + 1)
                    # Make sure the new item is selected
                    self.keywordList.SetItemState(item + 1, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)
                    # Make sure the item is visible
                    self.keywordList.EnsureVisible(item + 1)

    def SetClips(self, clipList):
        """ Allows the calling routine to provide a list of Clips that should be included on the Clips Tab.
            A sorted list of (clipName, collectionNumber, checked(boolean)) information should be passed in. """
        # Save the original data.  We'll need it to retrieve the Collection Number data
        self.originalClipData = clipList
        # Iterate through the clip list that was passed in
        for (clipID, clipNum, checked) in clipList:
            # Create a new Item in the Clip List at the end of the list.  Add the Clip ID data.
            index = self.clipList.InsertStringItem(sys.maxint, clipID)
            # Add the Collection data
            tempColl = Collection.Collection(clipNum)
            tempStr = ""
            for coll in tempColl.GetNodeData():
                if tempStr != "":
                    tempStr += " > "
                tempStr += coll
            self.clipList.SetStringItem(index, 1, tempStr)
            # If the item should be checked, check it!
            if checked:
                self.clipList.CheckItem(index)
        # Set the column width for Clip IDs
        self.clipList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Collection Data
        self.clipList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # lay out the Keywords Panel
        self.clipsPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()

    def GetClips(self):
        """ Allows the calling routine to retrieve the clip data from the Filter Dialog.  A sorted list
            of (clipID, collectNum, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty clip list
        clipList = []
        # Iterate throught the ClipListCtrl items
        for x in range(self.clipList.GetItemCount()):
            # Append the list item's data to the clip list
            clipList.append((self.clipList.GetItem(x, 0).GetText(), self.originalClipData[x][1], self.clipList.IsChecked(x)))
        # Return the clip list
        return clipList
            
    def SetKeywords(self, kwList):
        """ Allows the calling routine to provide a list of keywords that should be included on the Keywords Tab.
            A sorted list of (keywordgroup, keyword, checked(boolean)) information should be passed in. """
        # Iterate through the keyword list that was passed in
        for (kwg, kw, checked) in kwList:
            # Create a new Item in the Keyword List at the end of the list.  Add the Keyword Group data.
            index = self.keywordList.InsertStringItem(sys.maxint, kwg)
            # Add the Keyword data
            self.keywordList.SetStringItem(index, 1, kw)
            # If the item should be checked, check it!
            if checked:
                self.keywordList.CheckItem(index)
        # Set the column width for Keyword Groups
        self.keywordList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Keywords
        self.keywordList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # lay out the Keywords Panel
        self.keywordsPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()

    def GetKeywords(self):
        """ Allows the calling routine to retrieve the keyword data from the Filter Dialog.  A sorted list
            of (keywordgroup, keyword, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty keyword list
        keywordList = []
        # Iterate throught the KeywordListCtrl items
        for x in range(self.keywordList.GetItemCount()):
            # Append the list item's data to the keyword list
            keywordList.append((self.keywordList.GetItem(x, 0).GetText(), self.keywordList.GetItem(x, 1).GetText(), self.keywordList.IsChecked(x)))
        # Return the keyword list
        return keywordList

    def GetStartTime(self):
        """ Return the value of the Start Time text control on the Time Range tab """
        return self.startTime.GetValue()

    def GetEndTime(self):
        """ Return teh value of the End Time text control on the Time Range tab """
        return self.endTime.GetValue()

    def OnHelp(self, event):
        """ Implement the Filter Dialog Box's Help function """
        # Define the Help Context
        HelpContext = "Filter Dialog"
        # If a Help Window is defined ...
        if TransanaGlobal.menuWindow != None:
            # ... call Help!
            TransanaGlobal.menuWindow.ControlObject.Help(HelpContext)
