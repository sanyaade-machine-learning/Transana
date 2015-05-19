# Copyright (C) 2003 - 2007 The Board of Regents of the University of Wisconsin System 
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
# import Transana's Globals
import TransanaGlobal


class SearchDialog(wx.Dialog):
    """ Dialog Box that implements Transana's Search Interface. """

    def __init__(self, searchName=''):
        """ Initialize the Search Dialog, passing in the default Search Name. """
        # Define the SearchDialog as a resizable wxDialog Box
        wx.Dialog.__init__(self, TransanaGlobal.menuWindow, -1, _("Boolean Keyword Search"), wx.DefaultPosition, wx.Size(450, 480), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Set the lineStarted Indicator to False
        self.lineStarted = False
        # Set the parensOpen Indicator to False
        self.parensOpen = 0

        # Specify the minimum acceptable width and height for this window
        self.SetSizeHints(400, 440)

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
        self.btnAnd = wx.Button(self, -1, _('AND'))  #  , size=wx.Size(50, 24))
        self.btnAnd.SetConstraints(lay)
        self.btnAnd.Enable(False)
        wx.EVT_BUTTON(self, self.btnAnd.GetId(), self.OnBtnClick)

        # Add OR Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.left.RightOf(self.btnAnd, 10)
        lay.width.Absolute(50)  # AsIs()
        lay.height.AsIs()
        self.btnOr = wx.Button(self, -1, _('OR'))  # , size=wx.Size(50, 24))
        self.btnOr.SetConstraints(lay)
        self.btnOr.Enable(False)
        wx.EVT_BUTTON(self, self.btnOr.GetId(), self.OnBtnClick)

        # Add NOT Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.left.PercentOf(self, wx.Width, 35)
        lay.width.Absolute(50)  # AsIs()
        lay.height.AsIs()
        self.btnNot = wx.Button(self, -1, _('NOT'))  # , size=wx.Size(50, 24))
        self.btnNot.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnNot.GetId(), self.OnBtnClick)

        # Add Left Parenthesis Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.left.PercentOf(self, wx.Width, 55)
        lay.width.Absolute(30)  # AsIs()
        lay.height.AsIs()
        self.btnLeftParen = wx.Button(self, -1, '(')  # , size=wx.Size(30, 24))
        self.btnLeftParen.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnLeftParen.GetId(), self.OnBtnClick)

        # Add Right Parenthesis Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.left.RightOf(self.btnLeftParen, 10)
        lay.width.Absolute(30)  # AsIs()
        lay.height.AsIs()
        self.btnRightParen = wx.Button(self, -1, ')')  # , size=wx.Size(30, 24))
        self.btnRightParen.SetConstraints(lay)
        self.btnRightParen.Enable(False)
        wx.EVT_BUTTON(self, self.btnRightParen.GetId(), self.OnBtnClick)

        # Add Reset Button
        lay = wx.LayoutConstraints()
        lay.top.Below(operatorsText, 3)
        lay.right.SameAs(self, wx.Right, 10)
        lay.width.Absolute(100)  # AsIs()
        lay.height.AsIs()
        self.btnReset = wx.Button(self, -1, _('Reset'))  # , size=wx.Size(100, 24))
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
        self.btnAdd = wx.Button(self, -1, _('Add Keyword to Query'))  # , size=wx.Size(240, 24))
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

        # Add "Search" Button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.right.SameAs(self, wx.Right, 190)
        lay.width.Absolute(80)  # AsIs()
        lay.height.AsIs()
        self.btnSearch = wx.Button(self, -1, _('Search'))  # , size=wx.Size(80, 24))
        self.btnSearch.SetConstraints(lay)
        self.btnSearch.Enable(False)
        wx.EVT_BUTTON(self, self.btnSearch.GetId(), self.OnBtnClick)

        # Add "Cancel" Button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.right.SameAs(self, wx.Right, 100)
        lay.width.Absolute(80)  # AsIs()
        lay.height.AsIs()
        self.btnCancel = wx.Button(self, -1, _('Cancel'))  # , size=wx.Size(80, 24))
        self.btnCancel.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnCancel.GetId(), self.OnBtnClick)

        # Add "Help" Button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.right.SameAs(self, wx.Right, 10)
        lay.width.Absolute(80)  # AsIs()
        lay.height.AsIs()
        self.btnHelp = wx.Button(self, -1, _('Help'))  # , size=wx.Size(80, 24))
        self.btnHelp.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnHelp.GetId(), self.OnBtnClick)

        # Highlight the default Search name and set the focus to the Search Name.
        # (This is done automatically on Windows, but needs to be done explicitly on the Mac.)
        self.searchName.SetSelection(-1, -1)
        self.searchName.SetFocus()
        # Lay out the form
        self.Layout()
        # Have the form handle layout changes automatically
        self.SetAutoLayout(True)
        self.CenterOnScreen()
        
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
            # No Parens are Open
            self.parensOpen = 0
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
            keywordGroup = self.kw_group_lb.GetStringSelection()
            keyword = self.kw_lb.GetStringSelection()
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
            # Disable the "(" (Left Paren) button
            self.btnLeftParen.Enable(False)

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

#        else:
#            print "Unknown Button Pressed."


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
