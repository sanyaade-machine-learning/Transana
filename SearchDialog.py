# Copyright (C) 2003 - 2010 The Board of Regents of the University of Wisconsin System 
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

""" This module implements the Search Interface. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx

# import Transana's Database interface
import DBInterface
# import Transana's common Dialogs
import Dialogs
# import Transana's Globals
import TransanaGlobal
# import Python's os module
import os


class SearchDialog(wx.Dialog):
    """ Dialog Box that implements Transana's Search Interface. """

    def __init__(self, searchName=''):
        """ Initialize the Search Dialog, passing in the default Search Name. """
        # Define the SearchDialog as a resizable wxDialog Box
        wx.Dialog.__init__(self, TransanaGlobal.menuWindow, -1, _("Boolean Keyword Search"), wx.DefaultPosition, wx.Size(500, 480), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Set the lineStarted Indicator to False
        self.lineStarted = False
        # Set the parensOpen Indicator to False
        self.parensOpen = 0
        # Let's create an "Undo Stack" for the search terms
        self.ClearSearchStack()
        self.configName = ''
        self.reportType = 15

        # Specify the minimum acceptable width and height for this window
        self.SetSizeHints(500, 440)

        # Define all GUI Elements for the Form
        
        # Add Search Name Label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self, wx.Top, 10)
        lay.left.SameAs(self, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        searchNameText = wx.StaticText(self, -1, _('Search Name:'))
        searchNameText.SetConstraints(lay)
        
        # Add Search Name Text Box
        lay = wx.LayoutConstraints()
        lay.top.Below(searchNameText, 3)
        lay.left.SameAs(self, wx.Left, 10)
        lay.right.SameAs(self, wx.Right, 10)
        lay.height.AsIs()
        self.searchName = wx.TextCtrl(self, -1)
        self.searchName.SetValue(searchName)
        self.searchName.SetConstraints(lay)

        # Add Boolean Operators Label
        lay = wx.LayoutConstraints()
        lay.top.Below(self.searchName, 10)
        lay.left.SameAs(self, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        operatorsText = wx.StaticText(self, -1, _('Operators:'))
        operatorsText.SetConstraints(lay)
        
        # Add AND Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.left.SameAs(self, wx.Left, 10)
        lay.width.Absolute(50)  # AsIs()
        lay.height.AsIs()
        self.btnAnd = wx.Button(self, -1, _('AND'))
        self.btnAnd.SetConstraints(lay)
        self.btnAnd.Enable(False)
        wx.EVT_BUTTON(self, self.btnAnd.GetId(), self.OnBtnClick)

        # Add OR Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.left.RightOf(self.btnAnd, 10)
        lay.width.Absolute(50)  # AsIs()
        lay.height.AsIs()
        self.btnOr = wx.Button(self, -1, _('OR'))
        self.btnOr.SetConstraints(lay)
        self.btnOr.Enable(False)
        wx.EVT_BUTTON(self, self.btnOr.GetId(), self.OnBtnClick)

        # Add NOT Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.left.PercentOf(self, wx.Width, 30)
        lay.width.Absolute(50)  # AsIs()
        lay.height.AsIs()
        self.btnNot = wx.Button(self, -1, _('NOT'))
        self.btnNot.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnNot.GetId(), self.OnBtnClick)

        # Add Left Parenthesis Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.left.PercentOf(self, wx.Width, 50)
        lay.width.Absolute(30)  # AsIs()
        lay.height.AsIs()
        self.btnLeftParen = wx.Button(self, -1, '(')
        self.btnLeftParen.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnLeftParen.GetId(), self.OnBtnClick)

        # Add Right Parenthesis Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.left.RightOf(self.btnLeftParen, 10)
        lay.width.Absolute(30)  # AsIs()
        lay.height.AsIs()
        self.btnRightParen = wx.Button(self, -1, ')')
        self.btnRightParen.SetConstraints(lay)
        self.btnRightParen.Enable(False)
        wx.EVT_BUTTON(self, self.btnRightParen.GetId(), self.OnBtnClick)

        # Add "Undo" Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.right.SameAs(self, wx.Right, 100)
        lay.width.Absolute(32)  # AsIs()
        lay.height.AsIs()
        # Get the image for Undo
        bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Undo16.xpm"), wx.BITMAP_TYPE_XPM)
        self.btnUndo = wx.BitmapButton(self, -1, bmp)
        self.btnUndo.SetToolTip(wx.ToolTip(_('Undo')))
        self.btnUndo.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnUndo.GetId(), self.OnBtnClick)

        # Add Reset Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.right.SameAs(self, wx.Right, 10)
        lay.width.Absolute(80)  # AsIs()
        lay.height.AsIs()
        self.btnReset = wx.Button(self, -1, _('Reset'))
        self.btnReset.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnReset.GetId(), self.OnBtnClick)

        # Add Keyword Groups Label
        lay = wx.LayoutConstraints()
        lay.top.Below(self.btnAnd, 10)
        lay.left.SameAs(self, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        keywordGroupsText = wx.StaticText(self, -1, _('Keyword Groups:'))
        keywordGroupsText.SetConstraints(lay)
        
        # Add Keywords Label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(keywordGroupsText, wx.Top, 0)
        lay.left.PercentOf(self, wx.Width, 52)
        lay.width.AsIs()
        lay.height.AsIs()
        keywordsText = wx.StaticText(self, -1, _('Keywords:'))
        keywordsText.SetConstraints(lay)
        
        # Add Keyword Groups
        lay = wx.LayoutConstraints()
        lay.top.Below(keywordGroupsText, 3)
        lay.bottom.SameAs(self, wx.Bottom, 226)
        lay.left.SameAs(keywordGroupsText, wx.Left, 0)
        lay.width.PercentOf(self, wx.Width, 46)
        # Get the Keyword Groups from the Database Interface
        self.kw_groups = DBInterface.list_of_keyword_groups()
        self.kw_group_lb = wx.ListBox(self, -1, wx.DefaultPosition, wx.DefaultSize, self.kw_groups)
        self.kw_group_lb.SetConstraints(lay)
        # Select the first item in the list (required for Mac)
        if len(self.kw_groups) > 0:
            self.kw_group_lb.SetSelection(0)
        # Define the "Keyword Group Select" behavior
        wx.EVT_LISTBOX(self, self.kw_group_lb.GetId(), self.OnKeywordGroupSelect)

        # Add Keywords
        lay = wx.LayoutConstraints()
        lay.top.Below(keywordGroupsText, 3)
        lay.bottom.SameAs(self, wx.Bottom, 226)
        lay.left.SameAs(keywordsText, wx.Left, 0)
        lay.width.PercentOf(self, wx.Width, 46)
        # If there are defined Keyword Groups, load the Keywords for the first Group in the list
        if len(self.kw_groups) > 0:
            self.kw_list = DBInterface.list_of_keywords_by_group(self.kw_groups[0])
        else:
            self.kw_list = []
        self.kw_lb = wx.ListBox(self, -1, wx.DefaultPosition, wx.DefaultSize, self.kw_list)
        # Select the first item in the list (required for Mac)
        if len(self.kw_list) > 0:
            self.kw_lb.SetSelection(0)
        self.kw_lb.SetConstraints(lay)
        # Define the "Keyword Select" behavior
        wx.EVT_LISTBOX(self, self.kw_lb.GetId(), self.OnKeywordSelect)
        # Double-clicking a Keyword is equivalent to selecting it and pressing the "Add Keyword to Query" button
        wx.EVT_LISTBOX_DCLICK(self, self.kw_lb.GetId(), self.OnBtnClick)

        # Add "Add Keyword to Query" Button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 190)
        lay.centreX.SameAs(self, wx.CentreX, 0)
        lay.width.Absolute(240)  # AsIs()
        lay.height.AsIs()
        self.btnAdd = wx.Button(self, -1, _('Add Keyword to Query'))
        self.btnAdd.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnAdd.GetId(), self.OnBtnClick)

        # Add Search Query Label
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 170)
        lay.left.SameAs(self, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        searchQueryText = wx.StaticText(self, -1, _('Search Query:'))
        searchQueryText.SetConstraints(lay)
        
        # Add Search Query Text Box
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 46)
        lay.left.SameAs(self, wx.Left, 10)
        lay.right.SameAs(self, wx.Right, 10)
        lay.height.Absolute(120)
        # The Search Query is Read-Only
        self.searchQuery = wx.TextCtrl(self, -1, style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.searchQuery.SetConstraints(lay)

        # Add the File Open button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.left.SameAs(self, wx.Left, 10)
        lay.width.Absolute(32)  # AsIs()
        lay.height.AsIs()
        # Get the image for File Open
        bmp = wx.ArtProvider_GetBitmap(wx.ART_FILE_OPEN, wx.ART_TOOLBAR, (16,16))
        # Create the File Open button
        self.btnFileOpen = wx.BitmapButton(self, -1, bmp)
        self.btnFileOpen.SetToolTip(wx.ToolTip(_('Load a Search')))
        self.btnFileOpen.SetConstraints(lay)
        self.btnFileOpen.Bind(wx.EVT_BUTTON, self.OnFileOpen)

        # Add the File Save button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.left.SameAs(self, wx.Left, 45)
        lay.width.Absolute(32)  # AsIs()
        lay.height.AsIs()
        # Get the image for File Save
        bmp = wx.ArtProvider_GetBitmap(wx.ART_FILE_SAVE, wx.ART_TOOLBAR, (16,16))
        # Create the File Save button
        self.btnFileSave = wx.BitmapButton(self, -1, bmp)
        self.btnFileSave.SetToolTip(wx.ToolTip(_('Save a Search')))
        self.btnFileSave.SetConstraints(lay)
        self.btnFileSave.Bind(wx.EVT_BUTTON, self.OnFileSave)
        self.btnFileSave.Enable(False)

        # Add the File Delete button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.left.SameAs(self, wx.Left, 80)
        lay.width.Absolute(32)  # AsIs()
        lay.height.AsIs()
        # Get the image for File Delete
        bmp = wx.ArtProvider_GetBitmap(wx.ART_DELETE, wx.ART_TOOLBAR, (16,16))
        # Create the File Delete button
        self.btnFileDelete = wx.BitmapButton(self, -1, bmp)
        self.btnFileDelete.SetToolTip(wx.ToolTip(_('Delete a Saved Search')))
        self.btnFileDelete.SetConstraints(lay)
        self.btnFileDelete.Bind(wx.EVT_BUTTON, self.OnFileDelete)

        # Add "Search" Button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.right.SameAs(self, wx.Right, 190)
        lay.width.Absolute(80)  # AsIs()
        lay.height.AsIs()
        self.btnSearch = wx.Button(self, -1, _('Search'))
        self.btnSearch.SetConstraints(lay)
        self.btnSearch.Enable(False)
        wx.EVT_BUTTON(self, self.btnSearch.GetId(), self.OnBtnClick)

        # Add "Cancel" Button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.right.SameAs(self, wx.Right, 100)
        lay.width.Absolute(80)  # AsIs()
        lay.height.AsIs()
        self.btnCancel = wx.Button(self, -1, _('Cancel'))
        self.btnCancel.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnCancel.GetId(), self.OnBtnClick)

        # Add "Help" Button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.right.SameAs(self, wx.Right, 10)
        lay.width.Absolute(80)  # AsIs()
        lay.height.AsIs()
        self.btnHelp = wx.Button(self, -1, _('Help'))
        self.btnHelp.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnHelp.GetId(), self.OnBtnClick)

        # Highlight the default Search name and set the focus to the Search Name.
        # (This is done automatically on Windows, but needs to be done explicitly on the Mac.)
        self.searchName.SetSelection(-1, -1)
        # put the cursor firmly in the Search Name field
        self.searchName.SetFocus()
        # Lay out the form
        self.Layout()
        # Have the form handle layout changes automatically
        self.SetAutoLayout(True)
        # Center the dialog on screen
        self.CenterOnScreen()
        # Refresh the display
        self.Refresh()
        

    def OnBtnClick(self, event):
        """ This method handles all Button Clicks for the Search Dialog. """
        
        # "AND" Button
        if event.GetId() == self.btnAnd.GetId():
            # Add the appropriate text to the Search Query
            self.searchQuery.AppendText(' AND\n')
            # Enable the "Add" button
            self.btnAdd.Enable(True)
            # Disable the "And" button
            self.btnAnd.Enable(False)
            # Disable the "Or" button
            self.btnOr.Enable(False)
            # Enable the "Not" button
            self.btnNot.Enable(True)
            # Enable the "(" (Left Paren) button
            self.btnLeftParen.Enable(True)
            # Disable the ")" (Right Paren) button
            self.btnRightParen.Enable(False)
            # Disable the "Search" button
            self.btnSearch.Enable(False)
            self.btnFileSave.Enable(False)
            # Add to the Search Stack
            self.SaveSearchStack()

        # "OR" Button
        elif event.GetId() == self.btnOr.GetId():
            # Add the appropriate text to the Search Query
            self.searchQuery.AppendText(' OR\n')
            # Enable the "Add" button
            self.btnAdd.Enable(True)
            # Disable the "And" button
            self.btnAnd.Enable(False)
            # Disable the "Or" button
            self.btnOr.Enable(False)
            # Enable the "Not" button
            self.btnNot.Enable(True)
            # Enable the "(" (Left Paren) button
            self.btnLeftParen.Enable(True)
            # Disable the ")" (Right Paren) button
            self.btnRightParen.Enable(False)
            # Disable the "Search" button
            self.btnSearch.Enable(False)
            self.btnFileSave.Enable(False)
            # Add to the Search Stack
            self.SaveSearchStack()
            
        # "NOT" Button
        elif event.GetId() == self.btnNot.GetId():
            # Add the appropriate text to the Search Query
            self.searchQuery.AppendText('NOT ')
            # Disable the "Not" button
            self.btnNot.Enable(False)
            # Disable the "(" (Left Paren) button
            self.btnLeftParen.Enable(False)
            # Disable the ")" (Right Paren) button
            self.btnRightParen.Enable(False)
            # Disable the "Search" button
            self.btnSearch.Enable(False)
            self.btnFileSave.Enable(False)
            # Add to the Search Stack
            self.SaveSearchStack()

        # "(" (Open Paren) Button
        elif event.GetId() == self.btnLeftParen.GetId():
            # Add the appropriate text to the Search Query
            self.searchQuery.AppendText('(')
            self.parensOpen += 1
            # Disable the "And" button
            self.btnAnd.Enable(False)
            # Disable the "Or" button
            self.btnOr.Enable(False)
            # Disable the "Search" button
            self.btnSearch.Enable(False)
            self.btnFileSave.Enable(False)
            # Add to the Search Stack
            self.SaveSearchStack()

        # ")" (Close Paren) Button
        elif event.GetId() == self.btnRightParen.GetId():
            # Add the appropriate text to the Search Query
            self.searchQuery.AppendText(')')
            self.parensOpen -= 1
            if self.parensOpen == 0:
                # Disable the ")" (Right Paren) button
                self.btnRightParen.Enable(False)
                # Enable the "Search" button
                self.btnSearch.Enable(True)
                self.btnFileSave.Enable(True)
            # Add to the Search Stack
            self.SaveSearchStack()

        # "Reset" Button
        elif event.GetId() == self.btnReset.GetId():
            # Clear the Search Query Terms
            self.searchQuery.Clear()
            # No Line has been Started
            self.lineStarted = False
            # You are starting over, so enable "Add"
            self.btnAdd.Enable(True)
            # You can't add a Boolean Operator
            self.btnAnd.Enable(False)
            self.btnOr.Enable(False)
            # You can add a NOT Operator
            self.btnNot.Enable(True)
            # You can't perform a Search yet
            self.btnSearch.Enable(False)
            self.btnFileSave.Enable(False)
            # No Parens are Open
            self.parensOpen = 0
            # Reset the Search Stack
            self.ClearSearchStack()
            # Reset the Parens Buttons
            self.btnLeftParen.Enable(True)
            self.btnRightParen.Enable(False)
            # Deselect the current Keyword (We don't need to deselect the Keyword Group)
            try:
                self.kw_lb.SetSelection(self.kw_lb.GetSelection(), False)
            except:
                pass

        # "Add Keyword to Query" Button or (Keyword Double-Clicked when Add Button is enabled)
        elif (event.GetId() == self.btnAdd.GetId()) or \
             ((event.GetId() == self.kw_lb.GetId()) and (self.btnAdd.IsEnabled())):
            # Get the keyword group and keyword values
            keywordGroup = self.kw_group_lb.GetStringSelection()
            keyword = self.kw_lb.GetStringSelection()
            # A Keyword MUST be selected!
            if keyword <> '':
                # Add the appropriate text to the Search Query
                self.searchQuery.AppendText(keywordGroup + ':' + keyword)
                # Disable the "Add" button
                self.btnAdd.Enable(False)
                # Enable the "And" button
                self.btnAnd.Enable(True)
                # Enable the "Or" button
                self.btnOr.Enable(True)
                # Disable the "Not" button
                self.btnNot.Enable(False)
                # See if there are still parens that need to be closed
                if self.parensOpen > 0:
                    # If there are parens that need to be closed, enable the ")" (Right Paren) button
                    self.btnRightParen.Enable(True)
                else:
                    # If there are no parens that need to be closed, enable the "Search" button
                    self.btnSearch.Enable(True)
                    # and the save button
                    self.btnFileSave.Enable(True)
                # Disable the "(" (Left Paren) button
                self.btnLeftParen.Enable(False)
                # Add to the Search Stack
                self.SaveSearchStack()

        # "Undo" Button
        elif event.GetId() == self.btnUndo.GetId():
            # If the search stack has data ...
            if len(self.searchStack) > 1:
                # ... drop the last value added
                self.searchStack.pop()
                # Restore the Search Dialog based on the last data element in the Search Stack
                # Restore the Search Query
                self.searchQuery.SetValue(self.searchStack[-1][0])
                # Restore the Search Button
                self.btnSearch.Enable(self.searchStack[-1][1])
                # Restore the Save Button
                self.btnFileSave.Enable(self.searchStack[-1][1])
                # Restore the Add Button
                self.btnAdd.Enable(self.searchStack[-1][2])
                # Restore the And button
                self.btnAnd.Enable(self.searchStack[-1][3])
                # Restore the Or button
                self.btnOr.Enable(self.searchStack[-1][3])
                # Restore the Not button
                self.btnNot.Enable(self.searchStack[-1][4])
                # Restore the Left Parens button
                self.btnLeftParen.Enable(self.searchStack[-1][5])
                # Restore the Right Parens button
                self.btnRightParen.Enable(self.searchStack[-1][6])
                # Restore the counter of open paren pairs
                self.parensOpen = self.searchStack[-1][7]

        # "Search" Button
        elif event.GetId() == self.btnSearch.GetId():
            # Close with a wxID_OK result
            self.EndModal(wx.ID_OK)

        # "Cancel" Button
        elif event.GetId() == self.btnCancel.GetId():
            # Close with a wxID_CANCEL result
            self.EndModal(wx.ID_CANCEL)

        # "Help" Button
        elif event.GetId() == self.btnHelp.GetId():
            if TransanaGlobal.menuWindow != None:
                TransanaGlobal.menuWindow.ControlObject.Help('Search')

    def OnKeywordGroupSelect(self, event):
        """ Implement Interface Changes needed when a Keyword Group is selected. """
        # Get the List of Keywords for the selected Keyword Group from the Database Interface
        self.kw_list = DBInterface.list_of_keywords_by_group(event.GetString())
        # Reset the list of Keywords on screen
        self.kw_lb.Set(self.kw_list)


    def OnKeywordSelect(self, event):
        """ Implement Interface Changes needed when a Keyword is selected. """
        # Do nothing (!)
        pass

    def ClearSearchStack(self):
        """ Initialize the Search Undo Stack """
        # Set the initial values for the search stack:
        #   No Search Query
        #   Search and Save buttons disabled
        #   Add Keyword button enabled
        #   And and Or buttons disabled
        #   Not button enabled
        #   Left paren button enabled
        #   Right parent button disabled
        #   0 paren pairs open
        self.searchStack = [('', False, True, False, True, True, False, 0)]

    def SaveSearchStack(self):
        """ Add the current state of the dialog to the Search Undo Stack """
        # Save the state of all elements on the Search Dialog that need to be reset:
        #   Current Search Query
        #   Search and Save buttons states
        #   Add Keyword button state
        #   And and Or buttons states
        #   Not button state
        #   Left paren button state
        #   Right parent button state
        #   number of open paren pairs
        self.searchStack.append((self.searchQuery.GetValue(),
                                 self.btnSearch.IsEnabled(),
                                 self.btnAdd.IsEnabled(),
                                 self.btnAnd.IsEnabled(),
                                 self.btnNot.IsEnabled(),
                                 self.btnLeftParen.IsEnabled(),
                                 self.btnRightParen.IsEnabled(),
                                 self.parensOpen))


    def GetConfigNames(self):
        """ Get a list of Configuration Names for the current Report Type and Report Scope. """
        # Create a blank list
        resList = []
        # Get a Database Cursor
        DBCursor = DBInterface.get_db().cursor()
        # Build the database query
        query = """ SELECT ConfigName FROM Filters2
                      WHERE ReportType = %s
                      GROUP BY ConfigName
                      ORDER BY ConfigName """
        # Set up the data values that match the query
        values = (self.reportType)
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

    def OnFileOpen(self, event):
        """ Load a saved Search Query from the Filter Table """
        # Get a list of legal Report Names from the Database
        configNames = self.GetConfigNames()
        # Create a Choice Dialog so the user can select the saved Search they want
        dlg = SearchLoadDialog(self, _("Choose a saved Search to load"), _("Search"), configNames, self.reportType)
        # Center the Dialog on the screen
        dlg.CentreOnScreen()
        # Show the Choice Dialog and see if the user chooses OK
        if dlg.ShowModal() == wx.ID_OK:
            # Set the Cursor to the Hourglass while the filter is loaded
            TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
            # Remember the Configuration Name
            self.configName = dlg.GetStringSelection()
            # Get a Database Cursor
            DBCursor = DBInterface.get_db().cursor()
            # Build a query to get the selected report's data.  The Order By clause is
            # necessary so that when keyword colors are updated, they will have been loaded
            # by the time the list in the keyword control is re-populated by the KeywordList
            # (that is, we need FilterDataType 4 to be processed BEFORE FilterDataType 3)
            query = """ SELECT FilterData FROM Filters2
                          WHERE ReportType = %s AND
                                ConfigName = %s
                          ORDER BY FilterDataType DESC"""
            # Build the data values that match the query
            values = (self.reportType, self.configName.encode(TransanaGlobal.encoding))
            # Execute the query with the appropriate data values
            DBCursor.execute(query, values)
            # Initialize the Search Stack
            self.ClearSearchStack()
            # We shouldn't get multiple records, just one.
            (filterData,) = DBCursor.fetchone()
            # Get the Search Text data from the Database.
            # (If MySQLDB returns an Array, convert it to a String!)
            if type(filterData).__name__ == 'array':
                searchText = filterData.tostring()
            else:
                searchText = filterData
            # Replace the Search Query with the value from the database.
            self.searchQuery.SetValue(searchText)
            # Only valid searches can be saved, so we know the desired state of the interface buttons
            # Disable the "Add" button
            self.btnAdd.Enable(False)
            # Enable the "And" button
            self.btnAnd.Enable(True)
            # Enable the "Or" button
            self.btnOr.Enable(True)
            # Disable the "Not" button
            self.btnNot.Enable(False)
            # Enable the "Search" button
            self.btnSearch.Enable(True)
            # and the Save button
            self.btnFileSave.Enable(True)
            # Disable the "(" (Left Paren) button
            self.btnLeftParen.Enable(False)
            # Disable the ")" (Left Paren) button
            self.btnRightParen.Enable(False)
            # Reset the number of open paren pairs to NONE
            self.parensOpen = 0
            # Add to the Search Stack
            self.SaveSearchStack()

    def OnFileSave(self, event):
        """ Save a Search Query to the Filter table """
        # Get the data to be saved in the Filter table
        filterData = self.searchQuery.GetValue()
        # Remember the original Search Name
        originalSearchName = self.configName
        # Set a bogus error message to get us into the While loop
        errorMsg = "Begin"
        # Keep trying as long as there are error messages!
        while errorMsg != '':
            # Clear the error message, assuming the user will succeed in loading a configuration
            errorMsg = ''
            # Create a Dialog where the Configuration can be named
            dlg = wx.TextEntryDialog(self, _("Save Search As"), _("Search"), originalSearchName)
            # Show the dialog to the user and see how they respond
            if dlg.ShowModal() == wx.ID_OK:
                # Remove leading and trailing white space
                configName = dlg.GetValue().strip()
                # Get a list of existing Configuration Names from the database
                configNames = self.GetConfigNames()
                # If the user changed the Configuration Name and the new name already exists ...
                if (configName != originalSearchName) and (configName in configNames):
                    # Build and encode the error prompt
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('A search named "%s" already exists.  Do you want to replace it?'), 'utf8')
                    else:
                        prompt = _('A search named "%s" already exists.  Do you want to replace it?')
                    # Build a dialog to notify the user of the duplication and ask what to do
                    dlg2 = Dialogs.QuestionDialog(self, prompt % configName)
                    # Center the dialog on the screen
                    dlg2.CentreOnScreen()
                    # Display the dialog and get the user's response
                    if dlg2.LocalShowModal() != wx.ID_YES:
                        # If the user doesn't say to replace, create an error message (never displayed) to signal there's a problem.
                        errorMsg = 'Duplicate Search'
                    # Clean up after prompting the user for feedback
                    dlg2.Destroy()

                # Encode the data for saving
                configNameEnc = configName.encode(TransanaGlobal.encoding)
                filterDataEnc = filterData.encode(TransanaGlobal.encoding)
                # To proceed, we need a report name and we need the error message to still be blank
                if (configName != '') and (errorMsg == ''):

                    # Check to see if the Configuration record already exists
                    if DBInterface.record_match_count('Filters2',
                                                     ('ReportType', 'ConfigName'),
                                                     (self.reportType, configNameEnc)) > 0:
                        # Build the Update Query for Data
                        query = """ UPDATE Filters2
                                      SET FilterData = %s
                                      WHERE ReportType = %s AND
                                            ConfigName = %s """
                        values = (filterDataEnc, self.reportType, configNameEnc)
                    else:
                        query = """ INSERT INTO Filters2
                                        (ReportType, ConfigName, FilterData)
                                      VALUES
                                        (%s, %s, %s) """
                        values = (self.reportType, configNameEnc, filterDataEnc)
                    # Get a database cursor
                    DBCursor = DBInterface.get_db().cursor()
                    # Execute the query with the appropriate data
                    DBCursor.execute(query, values)
                    # Close the cursor
                    DBCursor.close()

                # If we can't proceed with the save ...
                else:
                    # If there's already an error message, we DON'T show it.
                    if errorMsg == '':
                        # If there's NOT an error messge, the User failed to enter a Report Name.
                        # Create an appropriate error message
                        errorMsg = _("You must enter a Name to save this Search.")
                        # Build an error dialog
                        dlg2 = Dialogs.ErrorDialog(self, errorMsg)
                        # Display the error message
                        dlg2.ShowModal()
                        # Destroy the error dialog.
                        dlg2.Destroy()
            # Destroy the Dialog that allows the user to name the configuration
            dlg.Destroy()

    def OnFileDelete(self, event):
        """ Delete a saved Search from the Filter table """
        # Get a list of legal Search Names from the Database
        configNames = self.GetConfigNames()
        # Create a Choice Dialog so the user can select the saved Search they want
        dlg = wx.SingleChoiceDialog(self, _("Choose a saved Search to delete"), _("Search"), configNames, wx.CHOICEDLG_STYLE)
        # Center the Dialog on the screen
        dlg.CentreOnScreen()
        # Show the Choice Dialog and see if the user chooses OK
        if dlg.ShowModal() == wx.ID_OK:
            # Remember the Configuration Name
            localConfigName = dlg.GetStringSelection()
            # Better confirm this.
            if 'unicode' in wx.PlatformInfo:
                prompt = unicode(_('Are you sure you want to delete saved Search "%s"?'), 'utf8')
            else:
                prompt = _('Are you sure you want to delete saved Search "%s"?')
            dlg2 = Dialogs.QuestionDialog(self, prompt % localConfigName)
            if dlg2.LocalShowModal() == wx.ID_YES:
                # Clear the global configuration name
                self.configName = ''
                # Get a Database Cursor
                DBCursor = DBInterface.get_db().cursor()
                # Build a query to delete the selected report
                query = """ DELETE FROM Filters2
                              WHERE ReportType = %s AND
                                    ConfigName = %s """
                # Build the data values that match the query
                values = (self.reportType, localConfigName.encode(TransanaGlobal.encoding))
                # Execute the query with the appropriate data values
                DBCursor.execute(query, values)
            dlg2.Destroy()
        # Destroy the Choice Dialog
        dlg.Destroy()

class SearchLoadDialog(wx.Dialog):
    """ Emulate wx.SingleChoiceDialog(). """
    def __init__(self, parent, prompt, title, configNames, reportType):
        # Note the dialog title
        self.title = title
        # Remember the list items for later comparison
        self.configNames = configNames
        # Note the report type
        self.reportType = reportType
        # Create a dialog box
        wx.Dialog.__init__(self, parent, -1, title, size=(350, 150), style=wx.CAPTION | wx.RESIZE_BORDER | wx.CLOSE_BOX | wx.STAY_ON_TOP)
        # Set "Window Variant" to small only for Mac to make fonts match better
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Set a vertical Box Sizer for the dialog
        box = wx.BoxSizer(wx.VERTICAL)

        # Show the prompt text
        message = wx.StaticText(self, -1, prompt)
        box.Add(message, 0, wx.EXPAND | wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.RIGHT | wx.TOP, 10)

        # Display a list of choices the user can select from.
        self.choices = wx.ListBox(self, -1, size=(350, 200), choices=configNames, style=wx.LB_SINGLE | wx.LB_ALWAYS_SB | wx.LB_SORT)
        box.Add(self.choices, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        # Select the first item, if the list isn't empty
        if len(configNames) > 0:
            self.choices.SetSelection(0)
        # Capture the double-click so it can emulate "Click and OK" function like the SingleChoiceDialog()
        self.choices.Bind(wx.EVT_LISTBOX_DCLICK, self.OnListDoubleClick)

        # Add a Horizontal Sizer for the buttons
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer on the left to expand, allowing the buttons to be right-justified
        box2.Add((1, 1), 1, wx.LEFT, 9)
        # Add an OK button and a Cancel button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))

        # Put the buttons in the  Button sizer
        box2.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        box2.Add(btnCancel, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Button Size to the Dialog Sizer
        box.Add(box2, 0, wx.EXPAND)

        # Turn AutoLayout on
        self.SetAutoLayout(True)

        # Set the main Dialog sizer
        self.SetSizer(box)
        # Fit the dialog to the sizer elements
        self.Fit()
        # Layout the Dialog
        self.Layout()
        # Center on the screen
        self.CentreOnScreen()

    def GetStringSelection(self):
        """ Return the string for the current selection in the ListBox """
        # Return the selection in the List Box as the selection in the Dialog
        return self.choices.GetStringSelection()

    def OnListDoubleClick(self, event):
        """ Handle double-click in the ListBox, which should emulate list selection and OK """
        # End the modal display with a return code matching the OK button
        self.EndModal(wx.ID_OK)
