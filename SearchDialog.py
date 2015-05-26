# Copyright (C) 2003 - 2014 The Board of Regents of the University of Wisconsin System 
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

USE_NOTEBOOK = True

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx
# import the CustomTreeCtrl
import wx.lib.agw.customtreectrl as CT
# import Python's cPickle module
import cPickle

# import Transana's Database interface
import DBInterface
# import Transana's common Dialogs
import Dialogs
# import Transana's Constants
import TransanaConstants
# import Transana's Globals
import TransanaGlobal
# import Transana's Images
import TransanaImages
# import Python's os module
import os

T_CHECK_ALL    =  wx.NewId()
T_CHECK_NONE   =  wx.NewId()

class SearchDialog(wx.Dialog):
    """ Dialog Box that implements Transana's Search Interface. """

    def __init__(self, searchName=''):
        """ Initialize the Search Dialog, passing in the default Search Name. """
        # Define the SearchDialog as a resizable wxDialog Box
        wx.Dialog.__init__(self, TransanaGlobal.menuWindow, -1, _("Boolean Keyword Search"), wx.DefaultPosition, wx.Size(500, 600),
                           style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

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
        self.SetSizeHints(500, 550)

        # Define all GUI Elements for the Form
        
        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Add Search Name Label
        searchNameText = wx.StaticText(self, -1, _('Search Name:'))
        mainSizer.Add(searchNameText, 0, wx.TOP | wx.LEFT, 10)
        mainSizer.Add((0, 3))
        
        # Add Search Name Text Box
        self.searchName = wx.TextCtrl(self, -1)
        self.searchName.SetValue(searchName)
        mainSizer.Add(self.searchName, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add Scope Label
        scopeText = wx.StaticText(self, -1, _('Search Scope:'))
        mainSizer.Add(scopeText, 0, wx.LEFT | wx.RIGHT, 10)
        mainSizer.Add((0, 3))
        
        # Create a Row Sizer for the "include" checkboxes
        includeSizer = wx.BoxSizer(wx.HORIZONTAL)

        # Add a spacer
        includeSizer.Add((1,1), 1, wx.EXPAND)

        # Add a checkbox for including Episodes in the Search Results
        self.includeEpisodes = wx.CheckBox(self, -1, _('Search Episodes'))
        # Include Episodes by default
        self.includeEpisodes.SetValue(True)
        # Bind the Check event
        self.includeEpisodes.Bind(wx.EVT_CHECKBOX, self.OnBtnClick)
        # Add the checkbox to the Include Sizer
        includeSizer.Add(self.includeEpisodes, 0)

        # Add a spacer
        includeSizer.Add((1,1), 1, wx.EXPAND)

        # Add a checkbox for including Clips in the Search Results
        self.includeClips = wx.CheckBox(self, -1, _('Search Clips'))
        # Include Clips by default
        self.includeClips.SetValue(True)
        # Bind the Check event
        self.includeClips.Bind(wx.EVT_CHECKBOX, self.OnBtnClick)
        # Add the checkbox to the Include Sizer
        includeSizer.Add(self.includeClips, 0)

        if TransanaConstants.proVersion:
            # Add a spacer
            includeSizer.Add((1,1), 1, wx.EXPAND)

            # Add a checkbox for including Snapshots in the Search Results
            self.includeSnapshots = wx.CheckBox(self, -1, _('Search Snapshots'))
            # Include Snapshots by default
            self.includeSnapshots.SetValue(True)
            # Bind the Check event
            self.includeSnapshots.Bind(wx.EVT_CHECKBOX, self.OnBtnClick)
            # Add the checkbox to the Include Sizer
            includeSizer.Add(self.includeSnapshots)
        else:
            # Add a checkbox for including Snapshots in the Search Results
            self.includeSnapshots = wx.CheckBox(self, -1, _('Search Snapshots'))
            # Don't Display this checkbox!
            self.includeSnapshots.Show(False)
            # Include Snapshots by default
            self.includeSnapshots.SetValue(False)

        # Add a spacer
        includeSizer.Add((1,1), 1, wx.EXPAND)

        # Add the Include Sizer on the Main Sizer
        mainSizer.Add(includeSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)


        # experimental NOTEBOOK INTERFACE
        if USE_NOTEBOOK:

            # Create a Notebook Control to allow different types of Search Information
            selectionNotebook = wx.Notebook(self, -1)
            # Set the Notebook Background to White (probably prevents a visible anomoly in Arabic!)
            selectionNotebook.SetBackgroundColour(wx.WHITE)


            # *********************************************************************************************************
            # Add a Panel to the Notebook for Text Search information
#            panelText = wx.Panel(selectionNotebook, -1)
#            panelText.SetBackgroundColour(wx.RED)

            # Add a Sizer to the Text Search Panel
#            panelTextSizer = wx.BoxSizer(wx.VERTICAL)



            # Add the Text Search Sizer to the Collections Panel
#            panelText.SetSizer(panelTextSizer)

            # Add the Text Search Panel to the Notebook as the initial page
#            selectionNotebook.AddPage(panelText, _("Text Search"), True)
            # *********************************************************************************************************

            # Add a Panel to the Notebook for Collection information
            panelCollections = wx.Panel(selectionNotebook, -1)
#            panelCollections.SetBackgroundColour(wx.BLUE)

            # Add a Sizer to the Collections Panel
            panelCollectionsSizer = wx.BoxSizer(wx.VERTICAL)

            collToolBar = wx.ToolBar(panelCollections, -1, style=wx.TB_HORIZONTAL | wx.NO_BORDER)
            # Toggle Button to indicate if child nodes should follow parent node
            self.btnChildFollow = collToolBar.AddCheckLabelTool(-1, "Checkable", TransanaImages.CheckTree.GetBitmap(), shortHelp=_("Change Nested Collections"))
            # have this be toggled by default
            self.btnChildFollow.Toggle()
            # Create the Check All button
            self.btnCheckAll = collToolBar.AddTool(T_CHECK_ALL, TransanaImages.Check.GetBitmap(), shortHelpString=_('Check All'))
            # Create the Uncheck All button
            self.btnCheckNone = collToolBar.AddTool(T_CHECK_NONE, TransanaImages.NoCheck.GetBitmap(), shortHelpString=_('Uncheck All'))
            collToolBar.Realize()
            panelCollectionsSizer.Add(collToolBar, 0, wx.EXPAND, 0)

            self.Bind(wx.EVT_MENU, self.OnCollectionSelectAll, self.btnCheckAll)
            self.Bind(wx.EVT_MENU, self.OnCollectionSelectAll, self.btnCheckNone)

            # Create a Tree Control with Checkboxes.
            # (I experimented with the TR_AUTO_CHECK_CHILD style, but that is not appropriate.  Not wanting data
            # from a certain collection doesn't necessarily imply we don't want it from the children.
            self.ctcCollections = CT.CustomTreeCtrl(panelCollections, agwStyle=wx.TR_DEFAULT_STYLE )
            # Set the TreeCtrl's background to white so it looks better
            self.ctcCollections.SetBackgroundColour(wx.WHITE)
            # Add the TreeCtrl to the Notebook Tab's Sizer
            panelCollectionsSizer.Add(self.ctcCollections, 1, wx.EXPAND | wx.ALL, 10)

            # Define the TreeCtrl's Images
            image_list = wx.ImageList(16, 16, 0, 2)
            image_list.Add(TransanaImages.Collection16.GetBitmap())
            image_list.Add(TransanaImages.db.GetBitmap())
            self.ctcCollections.SetImageList(image_list)

            # Add the TreeCtrl's Root Node
            self.ctcRoot = self.ctcCollections.AddRoot(_("Collections"))
            self.ctcCollections.SetItemImage(self.ctcRoot, 1, wx.TreeItemIcon_Normal)
            self.ctcCollections.SetItemImage(self.ctcRoot, 1, wx.TreeItemIcon_Selected)
            self.ctcCollections.SetItemImage(self.ctcRoot, 1, wx.TreeItemIcon_Expanded)
            self.ctcCollections.SetItemImage(self.ctcRoot, 1, wx.TreeItemIcon_SelectedExpanded)

            # Create a Mapping Dictionary for the Collections so we can place Nested Collections quickly
            mapDict = {}
            # Add the Root Node to the Mapping Dictionary
            mapDict[0] = self.ctcRoot

            # Get all the Collections from the Database and iterate through them
            for (collNo, collID, parentCollNo) in DBInterface.list_of_all_collections():
                # If the Mapping Dictionary has the PARENT Collection ...
                # (NOTE:  The query's ORDER BY clause means the parent should ALWAYS exist!!)
                if mapDict.has_key(parentCollNo):
                    # Get the Parent Node
                    parentItem = mapDict[parentCollNo]
                    # Create a new Checkbox Node for the current item a a child to the Parent Node
                    item = self.ctcCollections.AppendItem(parentItem, collID, ct_type=CT.TREE_ITEMTYPE_CHECK)
                    self.ctcCollections.SetItemImage(item, 0, wx.TreeItemIcon_Normal)
                    self.ctcCollections.SetItemImage(item, 0, wx.TreeItemIcon_Selected)
                    self.ctcCollections.SetItemImage(item, 0, wx.TreeItemIcon_Expanded)
                    self.ctcCollections.SetItemImage(item, 0, wx.TreeItemIcon_SelectedExpanded)
                    # Check the item
                    self.ctcCollections.CheckItem(item, True)
                    # Set the item's PyData to the Collection Number
                    self.ctcCollections.SetPyData(item, collNo)
                    # Add the new item to the Mapping Dictionary
                    mapDict[collNo] = item

                else:
                    print "SearchDialog.__init__(): Adding Collections.  ", collNo, collID, parentCollNo, "*** NOT ADDED ***"

            # Expand the tree's Root Node
            self.ctcCollections.Expand(self.ctcRoot)
            # Define the Item Check event handler.
            self.ctcCollections.Bind(CT.EVT_TREE_ITEM_CHECKED, self.OnCollectionsChecked)

            # Add the Collections Sizer to the Collections Panel
            panelCollections.SetSizer(panelCollectionsSizer)

            # Add the Collections Panel to the Notebook as the initial page
            selectionNotebook.AddPage(panelCollections, _("Collections"), True)

            # Add a Panel to the Notebook for the Keywords information
            panelKeywords = wx.Panel(selectionNotebook, -1)

            # Add a Sizer to the Keywords Panel
            panelKeywordsSizer = wx.BoxSizer(wx.VERTICAL)

            # Add Boolean Operators Label
            operatorsText = wx.StaticText(panelKeywords, -1, _('Operators:'))
            panelKeywordsSizer.Add(operatorsText, 0, wx.LEFT | wx.TOP, 10)
            panelKeywordsSizer.Add((0, 3))

            # Create a HORIZONTAL sizer for the first row
            r1Sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Add AND Button
            self.btnAnd = wx.Button(panelKeywords, -1, _('AND'), size=wx.Size(50, 24))
            r1Sizer.Add(self.btnAnd, 0)
            self.btnAnd.Enable(False)
            wx.EVT_BUTTON(self, self.btnAnd.GetId(), self.OnBtnClick)

            r1Sizer.Add((10, 0))

            # Add OR Button
            self.btnOr = wx.Button(panelKeywords, -1, _('OR'), size=wx.Size(50, 24))
            r1Sizer.Add(self.btnOr, 0)
            self.btnOr.Enable(False)
            wx.EVT_BUTTON(self, self.btnOr.GetId(), self.OnBtnClick)

            r1Sizer.Add((1, 0), 1, wx.EXPAND)

            # Add NOT Button
            self.btnNot = wx.Button(panelKeywords, -1, _('NOT'), size=wx.Size(50, 24))
            r1Sizer.Add(self.btnNot, 0)
            wx.EVT_BUTTON(self, self.btnNot.GetId(), self.OnBtnClick)

            r1Sizer.Add((1, 0), 1, wx.EXPAND)

            # Add Left Parenthesis Button
            self.btnLeftParen = wx.Button(panelKeywords, -1, '(', size=wx.Size(30, 24))
            r1Sizer.Add(self.btnLeftParen, 0)
            wx.EVT_BUTTON(self, self.btnLeftParen.GetId(), self.OnBtnClick)

            r1Sizer.Add((10, 0))

            # Add Right Parenthesis Button
            self.btnRightParen = wx.Button(panelKeywords, -1, ')', size=wx.Size(30, 24))
            r1Sizer.Add(self.btnRightParen, 0)
            self.btnRightParen.Enable(False)
            wx.EVT_BUTTON(self, self.btnRightParen.GetId(), self.OnBtnClick)

            r1Sizer.Add((1, 0), 1, wx.EXPAND)

            # Add "Undo" Button
            # Get the image for Undo
            bmp = TransanaImages.Undo16.GetBitmap()
            self.btnUndo = wx.BitmapButton(panelKeywords, -1, bmp, size=wx.Size(30, 24))
            self.btnUndo.SetToolTip(wx.ToolTip(_('Undo')))
            r1Sizer.Add(self.btnUndo, 0)
            wx.EVT_BUTTON(self, self.btnUndo.GetId(), self.OnBtnClick)

            r1Sizer.Add((10, 0))

            # Add Reset Button
            self.btnReset = wx.Button(panelKeywords, -1, _('Reset'), size=wx.Size(80, 24))
            r1Sizer.Add(self.btnReset, 0)
            wx.EVT_BUTTON(self, self.btnReset.GetId(), self.OnBtnClick)

            panelKeywordsSizer.Add(r1Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            r2Sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Add Keyword Groups Label
            keywordGroupsText = wx.StaticText(panelKeywords, -1, _('Keyword Groups:'))
            r2Sizer.Add(keywordGroupsText, 1, wx.EXPAND)

            r2Sizer.Add((10, 0))
            
            # Add Keywords Label
            keywordsText = wx.StaticText(panelKeywords, -1, _('Keywords:'))
            r2Sizer.Add(keywordsText, 1, wx.EXPAND)

            panelKeywordsSizer.Add(r2Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
            panelKeywordsSizer.Add((0, 3))
            
            r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

            # Add Keyword Groups
            self.kw_group_lb = wx.ListBox(panelKeywords, -1, wx.DefaultPosition, wx.DefaultSize, [])
            r3Sizer.Add(self.kw_group_lb, 1, wx.EXPAND)
            # Define the "Keyword Group Select" behavior
            wx.EVT_LISTBOX(self, self.kw_group_lb.GetId(), self.OnKeywordGroupSelect)

            r3Sizer.Add((10, 0))
            
            # Add Keywords
            self.kw_lb = wx.ListBox(panelKeywords, -1, wx.DefaultPosition, wx.DefaultSize, [])
            r3Sizer.Add(self.kw_lb, 1, wx.EXPAND)
            # Define the "Keyword Select" behavior
            wx.EVT_LISTBOX(self, self.kw_lb.GetId(), self.OnKeywordSelect)
            # Double-clicking a Keyword is equivalent to selecting it and pressing the "Add Keyword to Query" button
            wx.EVT_LISTBOX_DCLICK(self, self.kw_lb.GetId(), self.OnBtnClick)

            panelKeywordsSizer.Add(r3Sizer, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            # Add "Add Keyword to Query" Button
            self.btnAdd = wx.Button(panelKeywords, -1, _('Add Keyword to Query'), size=wx.Size(240, 24))
            panelKeywordsSizer.Add(self.btnAdd, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
            wx.EVT_BUTTON(self, self.btnAdd.GetId(), self.OnBtnClick)

            # Add Search Query Label
            searchQueryText = wx.StaticText(panelKeywords, -1, _('Search Query:'))
            panelKeywordsSizer.Add(searchQueryText, 0, wx.LEFT | wx.RIGHT, 10)
            panelKeywordsSizer.Add((0, 3))
            
            # Add Search Query Text Box
            # The Search Query is Read-Only
            self.searchQuery = wx.TextCtrl(panelKeywords, -1, size = wx.Size(200, 120), style=wx.TE_MULTILINE | wx.TE_READONLY)
            panelKeywordsSizer.Add(self.searchQuery, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            # Add the Keyword Sizer to the Keyword Panel
            panelKeywords.SetSizer(panelKeywordsSizer)

            # Add the Keywords Panel to tne Notebook and select it
            selectionNotebook.AddPage(panelKeywords, _("Keywords"), True)

            # Add the Notebook to the form's Main Sizer
            mainSizer.Add(selectionNotebook, 5, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            # *****************************************************************************************************
            # Set the Text Search Panel to AutoLayout
#            panelText.SetAutoLayout(True)
            # Lay Out the Text Search Panel
#            panelText.Layout()
            # *****************************************************************************************************

            # Set the Collection Panel to AutoLayout
            panelCollections.SetAutoLayout(True)
            # Lay Out the Collections Panel
            panelCollections.Layout()

            # Set the Keyword Panel to AutoLayout
            panelKeywords.SetAutoLayout(True)
            # Lay Out the Keyword Panel
            panelKeywords.Layout()

        # TRADITIONAL Keywords Only Interface
        else:
            
            # Add Boolean Operators Label
            operatorsText = wx.StaticText(self, -1, _('Operators:'))
            mainSizer.Add(operatorsText, 0, wx.LEFT, 10)
            mainSizer.Add((0, 3))
            
            # Create a HORIZONTAL sizer for the first row
            r1Sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Add AND Button
            self.btnAnd = wx.Button(self, -1, _('AND'), size=wx.Size(50, 24))
            r1Sizer.Add(self.btnAnd, 0)
            self.btnAnd.Enable(False)
            wx.EVT_BUTTON(self, self.btnAnd.GetId(), self.OnBtnClick)

            r1Sizer.Add((10, 0))

            # Add OR Button
            self.btnOr = wx.Button(self, -1, _('OR'), size=wx.Size(50, 24))
            r1Sizer.Add(self.btnOr, 0)
            self.btnOr.Enable(False)
            wx.EVT_BUTTON(self, self.btnOr.GetId(), self.OnBtnClick)

            r1Sizer.Add((1, 0), 1, wx.EXPAND)

            # Add NOT Button
            self.btnNot = wx.Button(self, -1, _('NOT'), size=wx.Size(50, 24))
            r1Sizer.Add(self.btnNot, 0)
            wx.EVT_BUTTON(self, self.btnNot.GetId(), self.OnBtnClick)

            r1Sizer.Add((1, 0), 1, wx.EXPAND)

            # Add Left Parenthesis Button
            self.btnLeftParen = wx.Button(self, -1, '(', size=wx.Size(30, 24))
            r1Sizer.Add(self.btnLeftParen, 0)
            wx.EVT_BUTTON(self, self.btnLeftParen.GetId(), self.OnBtnClick)

            r1Sizer.Add((10, 0))

            # Add Right Parenthesis Button
            self.btnRightParen = wx.Button(self, -1, ')', size=wx.Size(30, 24))
            r1Sizer.Add(self.btnRightParen, 0)
            self.btnRightParen.Enable(False)
            wx.EVT_BUTTON(self, self.btnRightParen.GetId(), self.OnBtnClick)

            r1Sizer.Add((1, 0), 1, wx.EXPAND)

            # Add "Undo" Button
            # Get the image for Undo
            bmp = TransanaImages.Undo16.GetBitmap()
            self.btnUndo = wx.BitmapButton(self, -1, bmp, size=wx.Size(30, 24))
            self.btnUndo.SetToolTip(wx.ToolTip(_('Undo')))
            r1Sizer.Add(self.btnUndo, 0)
            wx.EVT_BUTTON(self, self.btnUndo.GetId(), self.OnBtnClick)

            r1Sizer.Add((10, 0))

            # Add Reset Button
            self.btnReset = wx.Button(self, -1, _('Reset'), size=wx.Size(80, 24))
            r1Sizer.Add(self.btnReset, 0)
            wx.EVT_BUTTON(self, self.btnReset.GetId(), self.OnBtnClick)

            mainSizer.Add(r1Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            r2Sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Add Keyword Groups Label
            keywordGroupsText = wx.StaticText(self, -1, _('Keyword Groups:'))
            r2Sizer.Add(keywordGroupsText, 1, wx.EXPAND)

            r2Sizer.Add((10, 0))
            
            # Add Keywords Label
            keywordsText = wx.StaticText(self, -1, _('Keywords:'))
            r2Sizer.Add(keywordsText, 1, wx.EXPAND)

            mainSizer.Add(r2Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
            mainSizer.Add((0, 3))
            
            r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

            # Add Keyword Groups
            self.kw_group_lb = wx.ListBox(self, -1, wx.DefaultPosition, wx.DefaultSize, [])
            r3Sizer.Add(self.kw_group_lb, 1, wx.EXPAND)
            # Define the "Keyword Group Select" behavior
            wx.EVT_LISTBOX(self, self.kw_group_lb.GetId(), self.OnKeywordGroupSelect)

            r3Sizer.Add((10, 0))
            
            # Add Keywords
            self.kw_lb = wx.ListBox(self, -1, wx.DefaultPosition, wx.DefaultSize, [])
            r3Sizer.Add(self.kw_lb, 1, wx.EXPAND)
            # Define the "Keyword Select" behavior
            wx.EVT_LISTBOX(self, self.kw_lb.GetId(), self.OnKeywordSelect)
            # Double-clicking a Keyword is equivalent to selecting it and pressing the "Add Keyword to Query" button
            wx.EVT_LISTBOX_DCLICK(self, self.kw_lb.GetId(), self.OnBtnClick)

            mainSizer.Add(r3Sizer, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            # Add "Add Keyword to Query" Button
            self.btnAdd = wx.Button(self, -1, _('Add Keyword to Query'), size=wx.Size(240, 24))
            mainSizer.Add(self.btnAdd, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
            wx.EVT_BUTTON(self, self.btnAdd.GetId(), self.OnBtnClick)

            # Add Search Query Label
            searchQueryText = wx.StaticText(self, -1, _('Search Query:'))
            mainSizer.Add(searchQueryText, 0, wx.LEFT | wx.RIGHT, 10)
            mainSizer.Add((0, 3))
            
            # Add Search Query Text Box
            # The Search Query is Read-Only
            self.searchQuery = wx.TextCtrl(self, -1, size = wx.Size(200, 120), style=wx.TE_MULTILINE | wx.TE_READONLY)
            mainSizer.Add(self.searchQuery, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Create a Row sizer for the buttons at the bottom of the form
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)

        # Add the File Open button
        # Get the image for File Open
        bmp = wx.ArtProvider_GetBitmap(wx.ART_FILE_OPEN, wx.ART_TOOLBAR, (16,16))
        # Create the File Open button
        self.btnFileOpen = wx.BitmapButton(self, -1, bmp, size = wx.Size(32, 24))
        self.btnFileOpen.SetToolTip(wx.ToolTip(_('Load a Search')))
        btnSizer.Add(self.btnFileOpen, 0)
        self.btnFileOpen.Bind(wx.EVT_BUTTON, self.OnFileOpen)

        btnSizer.Add((10, 0))

        # Add the File Save button
        # Get the image for File Save
        bmp = wx.ArtProvider_GetBitmap(wx.ART_FILE_SAVE, wx.ART_TOOLBAR, (16,16))
        # Create the File Save button
        self.btnFileSave = wx.BitmapButton(self, -1, bmp, size = wx.Size(32, 24))
        self.btnFileSave.SetToolTip(wx.ToolTip(_('Save a Search')))
        btnSizer.Add(self.btnFileSave, 0)
        self.btnFileSave.Bind(wx.EVT_BUTTON, self.OnFileSave)
        self.btnFileSave.Enable(False)

        btnSizer.Add((10, 0))

        # Add the File Delete button
        # Get the image for File Delete
        bmp = wx.ArtProvider_GetBitmap(wx.ART_DELETE, wx.ART_TOOLBAR, (16,16))
        # Create the File Delete button
        self.btnFileDelete = wx.BitmapButton(self, -1, bmp, size = wx.Size(32, 24))
        self.btnFileDelete.SetToolTip(wx.ToolTip(_('Delete a Saved Search')))
        btnSizer.Add(self.btnFileDelete, 0)
        self.btnFileDelete.Bind(wx.EVT_BUTTON, self.OnFileDelete)

        btnSizer.Add((1, 0), 1, wx.EXPAND)

        # Add "Search" Button
        self.btnSearch = wx.Button(self, -1, _('Search'), size = wx.Size(80, 24))
        btnSizer.Add(self.btnSearch, 0)
        self.btnSearch.Enable(False)
        wx.EVT_BUTTON(self, self.btnSearch.GetId(), self.OnBtnClick)

        btnSizer.Add((10, 0))

        # Add "Cancel" Button
        self.btnCancel = wx.Button(self, -1, _('Cancel'), size = wx.Size(80, 24))
        btnSizer.Add(self.btnCancel, 0)
        wx.EVT_BUTTON(self, self.btnCancel.GetId(), self.OnBtnClick)

        btnSizer.Add((10, 0))

        # Add "Help" Button
        self.btnHelp = wx.Button(self, -1, _('Help'), size = wx.Size(80, 24))
        btnSizer.Add(self.btnHelp, 0)
        wx.EVT_BUTTON(self, self.btnHelp.GetId(), self.OnBtnClick)

        mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Highlight the default Search name and set the focus to the Search Name.
        # (This is done automatically on Windows, but needs to be done explicitly on the Mac.)
        self.searchName.SetSelection(-1, -1)
        # put the cursor firmly in the Search Name field
        self.searchName.SetFocus()

        self.SetSizer(mainSizer)
        # Have the form handle layout changes automatically
        self.SetAutoLayout(True)
        # Lay out the form
        self.Layout()
        # Center the dialog on screen
        self.CenterOnScreen()

        # Get the Keyword Groups from the Database Interface
        self.kw_groups = DBInterface.list_of_keyword_groups()
        for kwg in self.kw_groups:
            self.kw_group_lb.Append(kwg)
        # Select the first item in the list (required for Mac)
        if len(self.kw_groups) > 0:
            self.kw_group_lb.SetSelection(0)

        # If there are defined Keyword Groups, load the Keywords for the first Group in the list
        if len(self.kw_groups) > 0:
            self.kw_list = DBInterface.list_of_keywords_by_group(self.kw_groups[0])
        else:
            self.kw_list = []
        for kw in self.kw_list:
            self.kw_lb.Append(kw)
        # Select the first item in the list (required for Mac)
        if len(self.kw_list) > 0:
            self.kw_lb.SetSelection(0)

    def OnCollectionsChecked(self, event):
        """ Event Handler for when a Collection is checked un-checked """
        # Get the item that has been checked/unchecked
        sel = event.GetItem()
        # Make sure that item is expanded in the tree, so the user will see that nested collections have NOT been checked as well.
        self.ctcCollections.Expand(sel)
        # if the toggle is set so that check changes should be shared with children ...
        if self.btnChildFollow.IsToggled():
            # ... determine whether we're enabling or disabling
            enable = self.ctcCollections.IsItemChecked(sel)
            # ... and apply the setting to the children
            self.ctcCollections.CheckChilds(sel, enable)

    def EnableCollections(self, collNode, enable):
        """ A recursive method for enabling or disabling all Collections in the Collections Tree """
        # Get the First Child record
        (childNode, cookieItem) = self.ctcCollections.GetFirstChild(collNode)
        # While there are valid Child records ...
        while childNode.IsOk():
            # ... get the Collection Number out of the PyData
            collNum = self.ctcCollections.GetPyData(childNode)
            # Set the node's Checked status
            self.ctcCollections.CheckItem(childNode, enable)
            # If the Node has children ...
            if self.ctcCollections.HasChildren(childNode):
                # ... recursively call this method to set the Enabled status of children
                self.EnableCollections(childNode, enable)
            # If this node is not the LAST child ...
            if childNode != self.ctcCollections.GetLastChild(collNode):
                # ... then get the next child
                (childNode, cookieItem) = self.ctcCollections.GetNextChild(collNode, cookieItem)
            # if we're at the last child ...
            else:
                # ... we can quit
                break
            
    def OnCollectionSelectAll(self, event):
        """ An event handler for the Check All Collections and Uncheck All Collections Toolbar Buttons """
        # If we're Checking ALL ...
        if event.GetId() == self.btnCheckAll.GetId():
            # ... set Enable to True
            enable = True
        # If we're UnChecking ALL ...
        elif event.GetId() == self.btnCheckNone.GetId():
            # ... set Enable to False
            enable = False
        # Call the recursive Method for enabling/disabling all Collections
        self.EnableCollections(self.ctcRoot, enable)
        
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

        # The "include" checkboxes
        elif event.GetId() in [self.includeEpisodes.GetId(), self.includeClips.GetId(), self.includeSnapshots.GetId()]:
            # If we are CHECKING one of the boxes ...
            if event.IsChecked():
                # Detect the current status of the query to see if it's a valid search.  If so ...
                if (len(self.searchQuery.GetValue()) > 0) and (self.parensOpen == 0) and \
                   not (self.searchQuery.GetValue().rstrip().upper()[-4:] == ' AND') and \
                   not (self.searchQuery.GetValue().rstrip().upper()[-3:] == ' OR') and \
                   not (self.searchQuery.GetValue().rstrip().upper()[-4:] in ['\nNOT', '(NOT']):
                    # Enable the Search Button
                    self.btnSearch.Enable(True)
                    # and the save button
                    self.btnFileSave.Enable(True)

        # At least one of the "Include" checkboxes MUST be checked for the Search to be valid
        if not (self.includeEpisodes.IsChecked() or self.includeClips.IsChecked() or self.includeSnapshots.IsChecked()):
            # If no results are included, disable the "Search" button
            self.btnSearch.Enable(False)
            # and the save button
            self.btnFileSave.Enable(False)
            

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
        # Adjust query for sqlite, if needed
        query = DBInterface.FixQuery(query)
        # Set up the data values that match the query
        values = (self.reportType, )
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


    def UpdateCollectionList(self, collTree, collNode, uncheckedList):
        """ Recursively traverse through all nodes for the Search Collections checkbox tree,
            unchecking those in the "uncheckedList" and checking all other.
            This is for loading the Saved Search data for Collections.  """
        # Initialize a list of results to hold the Checked Collection records
        results = []
        # Get the First Child record
        (childNode, cookieItem) = collTree.GetFirstChild(collNode)
        # While there are valid Child records ...
        while childNode.IsOk():
            # ... get the Collection Number out of the PyData
            collNum = collTree.GetPyData(childNode)
            collName = collTree.GetItemText(childNode)
            # If the node is checked ...
            if (collNum, collName) in uncheckedList:
                collTree.CheckItem(childNode, False)
            else:
                collTree.CheckItem(childNode, True)
            # If the Node has children ... (The Node does NOT have to be checked.  You can check a child of an unchecked parent!)
            if collTree.HasChildren(childNode):
                # ... recursively call this method
                self.UpdateCollectionList(collTree, childNode, uncheckedList)
            # If this node is not the LAST child ...
            if childNode != collTree.GetLastChild(collNode):
                # ... then get the next child
                (childNode, cookieItem) = collTree.GetNextChild(collNode, cookieItem)
            # if we're at the last child ...
            else:
                # ... we can quit
                break
        # Return the results to the calling method
        return results


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
            # Build a query to get the Saved Search's Query data.
            query = """ SELECT FilterData FROM Filters2
                          WHERE ReportType = %s AND
                                ReportScope = 0 AND
                                ConfigName = %s
                          ORDER BY FilterDataType DESC"""
            # Adjust query for sqlite, if needed
            query = DBInterface.FixQuery(query)
            # Build the data values that match the query
            values = (self.reportType, self.configName.encode(TransanaGlobal.encoding))
            # Execute the query with the appropriate data values
            DBCursor.execute(query, values)
            # Initialize the Search Stack
            self.ClearSearchStack()
            # Get the query results
            data = DBCursor.fetchall()
            if len(data) > 0:
                # We shouldn't get multiple records, just one.
                (filterData,) = data[0]
                # Get the Search Text data from the Database.
                # (If MySQLDB returns an Array, convert it to a String!)
                if type(filterData).__name__ == 'array':
                    searchText = filterData.tostring()
                else:
                    searchText = filterData
                # Decode the text
                searchText = searchText.decode(TransanaGlobal.encoding)
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
            
            # Initialize CollectionsToSkip
            collectionsToSkip = []
            # Build a query to get the Saved Search's Collections data.
            query = """ SELECT FilterData FROM Filters2
                          WHERE ReportType = %s AND
                                ReportScope = 1 AND
                                ConfigName = %s
                          ORDER BY FilterDataType DESC"""
            # Adjust query for sqlite, if needed
            query = DBInterface.FixQuery(query)
            # Build the data values that match the query
            values = (self.reportType, self.configName.encode(TransanaGlobal.encoding))
            # Execute the query with the appropriate data values
            DBCursor.execute(query, values)

            # Get the data from the query
            data = DBCursor.fetchall()
            # Iterate through the data (There should only be one record!)
            for datum in data:
                # We shouldn't get multiple records, just one.
                (filterData,) = datum
                # Get the list of Unchecked Collections from the Database.
                # (If MySQLDB returns an Array, convert it to a String!)
                if type(filterData).__name__ == 'array':
                    collectionsDataEnc = cPickle.loads(filterData.tostring())
                else:
                    collectionsDataEnc = cPickle.loads(filterData)
                # Decode the data
                for coll in collectionsDataEnc:
                    collectionsToSkip.append((coll[0], coll[1].decode(TransanaGlobal.encoding)))
            # update the Collections Check Tree based on this configuration
            self.UpdateCollectionList(self.ctcCollections, self.ctcRoot, collectionsToSkip)

    def GetCollectionList(self, collTree, collNode, checkedVal):
        """ Recursively builds a list of all nodes for the Search Collections checkbox tree that match checkedVal """
        # Initialize a list of results to hold the Checked Collection records
        results = []
        # Get the First Child record
        (childNode, cookieItem) = collTree.GetFirstChild(collNode)
        # While there are valid Child records ...
        while childNode.IsOk():
            # ... get the Collection Number out of the PyData
            collNum = collTree.GetPyData(childNode)
            # If the node is checked ...
            if childNode.IsChecked() == checkedVal:
                # ... add the Collection Number and collection Name to the Results List
                results.append((collNum, collTree.GetItemText(childNode)))
            # If the Node has children ... (The Node does NOT have to be checked.  You can check a child of an unchecked parent!)
            if collTree.HasChildren(childNode):
                # ... recursively call this method to get the results of this node's child nodes, adding those results to these
                results += self.GetCollectionList(collTree, childNode, checkedVal)
            # If this node is not the LAST child ...
            if childNode != collTree.GetLastChild(collNode):
                # ... then get the next child
                (childNode, cookieItem) = collTree.GetNextChild(collNode, cookieItem)
            # if we're at the last child ...
            else:
                # ... we can quit
                break
        # Return the results to the calling method
        return results

    def OnFileSave(self, event):
        """ Save a Search Query to the Filter table """
        # Get a list of the Collections that are UN-CHECKED
        collectionsToSkip = self.GetCollectionList(self.ctcCollections, self.ctcRoot, False)
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
                collectionsDataEnc = []
                for coll in collectionsToSkip:
                    collectionsDataEnc.append((coll[0], coll[1].encode(TransanaGlobal.encoding)))
                # To proceed, we need a report name and we need the error message to still be blank
                if (configName != '') and (errorMsg == ''):

                    # Check to see if the Configuration record for the Query already exists
                    if DBInterface.record_match_count('Filters2',
                                                     ('ReportType', 'ReportScope', 'ConfigName'),
                                                     (self.reportType, 0, configNameEnc)) > 0:
                        # Build the Update Query for Data
                        query = """ UPDATE Filters2
                                      SET FilterData = %s
                                      WHERE ReportType = %s AND
                                            ReportScope = 0 AND
                                            ConfigName = %s """
                        values = (filterDataEnc, self.reportType, configNameEnc)
                    else:
                        query = """ INSERT INTO Filters2
                                        (ReportType, ReportScope, ConfigName, FilterData)
                                      VALUES
                                        (%s, 0, %s, %s) """
                        values = (self.reportType, configNameEnc, filterDataEnc)
                    # Adjust query for sqlite, if needed
                    query = DBInterface.FixQuery(query)
                    # Get a database cursor
                    DBCursor = DBInterface.get_db().cursor()
                    # Execute the query with the appropriate data
                    DBCursor.execute(query, values)

                    # Check to see if the Configuration record for the Collections already exists
                    if DBInterface.record_match_count('Filters2',
                                                     ('ReportType', 'ReportScope', 'ConfigName'),
                                                     (self.reportType, 1, configNameEnc)) > 0:
                        # Build the Update Query for Data
                        query = """ UPDATE Filters2
                                      SET FilterData = %s
                                      WHERE ReportType = %s AND
                                            ReportScope = 1 AND
                                            ConfigName = %s """
                        values = (filterDataEnc, self.reportType, configNameEnc)
                    else:
                        query = """ INSERT INTO Filters2
                                        (ReportType, ReportScope, ConfigName, FilterData)
                                      VALUES
                                        (%s, 1, %s, %s) """
                        values = (self.reportType, configNameEnc, cPickle.dumps(collectionsDataEnc))
                    # Adjust the query for sqlite if needed
                    query = DBInterface.FixQuery(query)
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
                # Adjust query for sqlite, if needed
                query = DBInterface.FixQuery(query)
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
