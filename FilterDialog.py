# Copyright (C) 2003 - 2008 The Board of Regents of the University of Wisconsin System 
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

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Kathleen Liston'

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
# Import Transana's Custom ColorListCtrl
import ColorListCtrl
# Import Transana's DBInterface
import DBInterface
# import Transana's Dialogs module for the ErrorDialog.
import Dialogs
# import Transana's Miscellaneous Functions
import Misc
# import Transana's Constants
import TransanaConstants
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
        # Bind the Item Activated method
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnItemActivated)

    def OnItemActivated(self, event):
        self.ToggleItem(event.m_itemIndex)

class FilterDialog(wx.Dialog):
    """ This window implements Episode, Clip, and Keyword filtering for Transana Reports.
        Required parameters are:
          parent
          id            (should be -1)
          title
          reportType     1 = Keyword Map  (reportScope is the Episode Number)
                         2 = Keyword Visualization  (reportScope is the Episode Number)
                         3 = Episode Clip Data Export (reportScope is the Episode Number)
                         4 = Collection Clip Data Export (reportScope is the Collection Number, or 0 for Collection Root)
                         5 = Series Keyword Sequence Map (reportScope is the Series Number)
                         6 = Series Keyword Bar Graph (reportScope is the Series Number)
                         7 = Series Keyword Percentage Map (reportScope is the Series Number)
                         8 = Episode Clip Data Coder Reliabiltiy Export (reportScope is the Episode Number)
                         9 = Keyword Summary Report for all Keyword Groups (configSave not yet implemented) (reportScope is not yet defined)
                        10 = Series Report (reportScope is Series Number)
                        11 = Episode Report (reportScope is Episode Number)
                        12 = Collection Report (reportScope is Collection Number)
                        13 = Notes Report (reportScope is 1 for all notes, 2 for Series, 3 for Episodes, 4 for Transcripts, 5 for Collections,
                                           6 for Clips)
                        14 = Series Clip Data Export (reportScope is the Series Number)
                        15 = Search Save (Search Saves have NO reportScope! or FilterDataType!!)

            *** ADDING A REPORT TYPE?  Remember to add the delete_filter_records() call to the appropriate
                object's db_delete() method!

                ALSO, remember to add the ReportScope conversion to XMLImport for Filter Imports! ***
                         
        Optional parameters are:
          loadDefault       (boolean) -- silently load a profile named "Default", if one exists
          configName        (current Configuration Name)
          reportScope       (required for Configuration Save/Load)
          episodeFilter     (boolean)
          episodeSort       (boolean)
          transcriptFilter  (boolean)
          collectionFilter  (boolean)
          collectionSort    (boolean)  * NOT FULLY IMPLEMENTED
          clipFilter        (boolean)
          clipSort          (boolean)
          keywordGroupFilter(boolean)
          keywordGroupColor (boolean)  * NOT FULLY IMPLEMENTED
          keywordFilter     (boolean)
          keywordSort       (boolean)
          keywordColor      (boolean)
          notesFilter       (boolean)
          options           (boolean)
          startTime         (number of milliseconds)
          endTime           (number of milliseconds)
          barHeight         (integer)
          whitespace        (integer)
          hGridLines        (boolean)
          vGridLines        (boolean)
          singleLineDisplay (boolean)
          showLegend        (boolean)
          colorOutput       (boolean)
          showClipTranscripts (boolean)
          showClipKeywords    (boolean)
          showNestedData      (boolean)
          showFile            (boolean)
          showTime            (boolean)
          showComments        (boolean)
          showCollectionNotes (boolean)
          showClipNotes       (boolean) """
    
    def __init__(self, parent, id, title, reportType, **kwargs):
        """ Initialize the Transana Filter Dialog Box """
        # Create a Dialog Box

        # Remember the keyword arguments
        self.kwargs = kwargs

        # See if we have a loadDefault argument and save it if we do.
        if self.kwargs.has_key('loadDefault'):
            self.loadDefault = self.kwargs['loadDefault']
        else:
            self.loadDefault = False
        
        # Remember the report type
        self.reportType = reportType

        # Initialize the Report Configuration Name if one is passed in, if not, initialize to an empty string
        if self.kwargs.has_key('configName'):
            self.configName = self.kwargs['configName']
        else:
            self.configName = ''

        # Remember the title
        self.title = title
        # If a configName exists ...
        if self.configName != '':
            # ... add it to the title for the window (but not what gets saved!)
            title += ' - ' + self.configName

        if self.kwargs.has_key('startTime') and self.kwargs['startTime']:
            self.startTimeVal = Misc.time_in_ms_to_str(self.kwargs['startTime'])
        else:
            self.startTimeVal = Misc.time_in_ms_to_str(0)
        if self.kwargs.has_key('endTime'):
            if self.kwargs['endTime']:
                self.endTimeVal = Misc.time_in_ms_to_str(self.kwargs['endTime'])
            else:
                self.endTimeVal = Misc.time_in_ms_to_str(parent.MediaLength)
        else:
            self.endTimeVal = self.startTimeVal

        # Initialize the dialog box.
        # Just to be clear, if we're loading the default, we don't actually SEE the dialog, but we still
        # create it and populate all of its fields.  That's the easiest way to load the default config!!
        wx.Dialog.__init__(self, parent, id, title, size = (500,500),
                           style=wx.CAPTION | wx.CLOSE_BOX | wx.RESIZE_BORDER | wx.NO_FULL_REPAINT_ON_RESIZE)
        # Make the background White
        self.SetBackgroundColour(wx.WHITE)

        # Create BoxSizers for the Dialog
        vBox = wx.BoxSizer(wx.VERTICAL)
        hBox = wx.BoxSizer(wx.HORIZONTAL)

        # Add the Tool Bar
        self.toolBar = wx.ToolBar(self, -1, style=wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_TEXT)
        # If there is a "reportScope" parameter, we should add Configuration Open, Save, and Delete buttons
        if self.kwargs.has_key('reportScope'):
            # Get the image for File Open
            bmp = wx.ArtProvider_GetBitmap(wx.ART_FILE_OPEN, wx.ART_TOOLBAR, (16,16))
            # Create the File Open button
            btnFileOpen = self.toolBar.AddTool(T_FILE_OPEN, bmp, shortHelpString=_("Load Filter Configuration"))
            # Create the File Save button
            btnFileSave = self.toolBar.AddTool(T_FILE_SAVE, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Save16.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_("Save Filter Configuration"))
            # Get the image for File Delete
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
        # If there is a "reportScope" parameter, we should implement Configuration Open, Save, and Delete buttons
        if self.kwargs.has_key('reportScope'):
            self.Bind(wx.EVT_MENU, self.OnFileOpen, btnFileOpen)
            self.Bind(wx.EVT_MENU, self.OnFileSave, btnFileSave)
            self.Bind(wx.EVT_MENU, self.OnFileDelete, btnFileDelete)
        self.Bind(wx.EVT_MENU, self.OnCheckAll, btnCheckAll)
        self.Bind(wx.EVT_MENU, self.OnCheckAll, btnCheckNone)
        self.Bind(wx.EVT_MENU, self.OnHelp, btnHelp)
        self.Bind(wx.EVT_MENU, self.OnClose, btnExit)
        # Add a spacer to the Sizer that allows for the Toolbar
        vBox.Add(self.toolBar, 0, wx.BOTTOM | wx.EXPAND, 5)

        # Everything in this Dialog goes inside a Notebook.  Create that notebook control
        self.notebook = wx.Notebook(self, -1)

        # If Report Contents Specification is requested ...
        if self.kwargs.has_key('reportContents') and self.kwargs['reportContents']:
            # ... build a Panel for Report Contents ...
            self.reportContentsPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            # ... place the Time Range Panel on the Notebook, creating a Time Range tab ...
            self.notebook.AddPage(self.reportContentsPanel, _("Report Contents"))
            
            # The Episode and Collection Reports need options for
            # showing Clip Transcripts, showing Clip Keywords, and showing Nested Data
            if self.kwargs.has_key('showNestedData'):
                self.showNestedData = wx.CheckBox(self.reportContentsPanel, -1, _("Include Clips from Nested Collections"))
                self.showNestedData.SetValue(self.kwargs['showNestedData'])
                pnlVSizer.Add(self.showNestedData, 0, wx.TOP | wx.LEFT, 10)
                text1 = wx.StaticText(self.reportContentsPanel, -1, _("(Unchecking this will cause clips from nested collections to\nbe skipped even if checked on the Clips tab.)"))
                pnlVSizer.Add(text1, 0, wx.LEFT, 30)
            if self.kwargs.has_key('showFile'):
                self.showFile = wx.CheckBox(self.reportContentsPanel, -1, _("Show Media File Name"))
                self.showFile.SetValue(self.kwargs['showFile'])
                pnlVSizer.Add(self.showFile, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showTime'):
                self.showTime = wx.CheckBox(self.reportContentsPanel, -1, _("Show Clip Time"))
                self.showTime.SetValue(self.kwargs['showTime'])
                pnlVSizer.Add(self.showTime, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showClipTranscripts'):
                self.showClipTranscripts = wx.CheckBox(self.reportContentsPanel, -1, _("Show Clip Transcripts"))
                self.showClipTranscripts.SetValue(self.kwargs['showClipTranscripts'])
                pnlVSizer.Add(self.showClipTranscripts, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showClipKeywords'):
                if self.reportType == 10:
                    prompt = _("Show Keywords")
                else:
                    prompt = _("Show Clip Keywords")
                self.showClipKeywords = wx.CheckBox(self.reportContentsPanel, -1, prompt)
                self.showClipKeywords.SetValue(self.kwargs['showClipKeywords'])
                pnlVSizer.Add(self.showClipKeywords, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showComments'):
                self.showComments = wx.CheckBox(self.reportContentsPanel, -1, _("Show Comments"))
                self.showComments.SetValue(self.kwargs['showComments'])
                pnlVSizer.Add(self.showComments, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showCollectionNotes'):
                self.showCollectionNotes = wx.CheckBox(self.reportContentsPanel, -1, _("Show Collection Notes"))
                self.showCollectionNotes.SetValue(self.kwargs['showCollectionNotes'])
                pnlVSizer.Add(self.showCollectionNotes, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showClipNotes'):
                self.showClipNotes = wx.CheckBox(self.reportContentsPanel, -1, _("Show Clip Notes"))
                self.showClipNotes.SetValue(self.kwargs['showClipNotes'])
                pnlVSizer.Add(self.showClipNotes, 0, wx.TOP | wx.LEFT, 10)

            # Now declare the panel's vertical sizer as the panel's official sizer
            self.reportContentsPanel.SetSizer(pnlVSizer)

        # If Episode Filtering is requested ...
        if self.kwargs.has_key('episodeFilter') and self.kwargs['episodeFilter']:
            # ... build a Panel for Episodes ...
            self.episodesPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Episode Panel on the Notebook, creating an Episodes tab ...
            self.notebook.AddPage(self.episodesPanel, _("Episodes"))
            # ... place a Check List Ctrl on the Episodes Panel ...
            self.episodeList = CheckListCtrl(self.episodesPanel)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.episodeList, 1, wx.EXPAND)
            # The episode List needs two columns, Episode ID and Series ID.
            self.episodeList.InsertColumn(0, _("Episode ID"))
            self.episodeList.InsertColumn(1, _("Series ID"))

            # If Episode Sorting capacity has been requested ...
            if self.kwargs.has_key('episodeSort') and self.kwargs['episodeSort']:
                # create a vertical sizer to hold the sort buttons
                pnlBtnSizer = wx.BoxSizer(wx.VERTICAL)
                # get the graphic for the Up arrow
                bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_UP, wx.ART_TOOLBAR, (16,16))
                # create a bitmap button for the Move Up button
                self.btnEpUp = wx.BitmapButton(self.episodesPanel, -1, bmp)
                # Set the Tool Tip for the Move Up button
                self.btnEpUp.SetToolTipString(_("Move episode up"))
                # Bind the button event to a method
                self.btnEpUp.Bind(wx.EVT_BUTTON, self.OnButton)
                # Insert some expandable space into the button sizer at the top
                pnlBtnSizer.Add((1,1), 1, wx.EXPAND)
                # Add the Move Up button to the button sizer
                pnlBtnSizer.Add(self.btnEpUp, 0, wx.ALIGN_CENTER | wx.ALL, 5)
                # Add a spacer to the button sizer to increase the amount of space between the buttons
                pnlBtnSizer.Add((1, 10))
                # Get the graphic for the Down arrow
                bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_DOWN, wx.ART_TOOLBAR, (16,16))
                # create a bitmap button for the Move Down button
                self.btnEpDown = wx.BitmapButton(self.episodesPanel, -1, bmp)
                # Set the Tool Tip for the Move Down button
                self.btnEpDown.SetToolTipString(_("Move episode down"))
                # Bind the button event to a method
                self.btnEpDown.Bind(wx.EVT_BUTTON, self.OnButton)
                # Add the Move Down button to the button sizer
                pnlBtnSizer.Add(self.btnEpDown, 0, wx.ALIGN_CENTER | wx.ALL, 5)
                # Add some expandable space below the button to keep the buttons centered vertically
                pnlBtnSizer.Add((1,1), 1, wx.EXPAND)
                # Add the button sizer to the panel's Horizontal sizer
                pnlHSizer.Add(pnlBtnSizer, 0, wx.EXPAND)
                
            # NOTE:  The actual list of Episodes needs to be provided by the calling routine using
            #        the SetEpisodes method.  This is because only the calling routine knows which
            #        Episodes are legal and should be included in the list.  We don't necessarily
            #        want all Episodes all the time.  The Episode List should be in the form of an
            #        ordered list of tuples made up of (EpisodeID, SeriesID, checked) information.

            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.episodesPanel.SetSizer(pnlVSizer)

        #If transcript filter is requested... (Kathleen)
        if self.kwargs.has_key('transcriptFilter') and self.kwargs['transcriptFilter']:
            # ... build a Panel for Transcript List ...
            self.transcriptPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Transcripts Panel on the Notebook, creating a Transcripts tab ...
            self.notebook.AddPage(self.transcriptPanel, _("Transcripts"))
            # ... place a Check List Ctrl on the Transcripts Panel ...
            self.transcriptList = CheckListCtrl(self.transcriptPanel)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.transcriptList, 1, wx.EXPAND)
            # The Transcripts List needs four columns, Series, Episode, Transcript, and Number of Clips.
            self.transcriptList.InsertColumn(0, _("Series"))
            self.transcriptList.InsertColumn(1, _("Episode"))
            self.transcriptList.InsertColumn(2, _("Transcript"))
            self.transcriptList.InsertColumn(3, _("# Clips"))
            
            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND, 0)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.transcriptPanel.SetSizer(pnlVSizer)
            
        #If collection filter is requested... (Kathleen)
        if self.kwargs.has_key('collectionFilter') and self.kwargs['collectionFilter']:
            # ... build a Panel for Collections ...
            self.collectionPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Collections Panel on the Notebook, creating a Collections tab ...
            self.notebook.AddPage(self.collectionPanel, _("Collections"))
            # ... place a Check List Ctrl on the Collections Panel ...
            self.collectionList = CheckListCtrl(self.collectionPanel)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.collectionList, 1, wx.EXPAND)
            # The Collections List needs one column, Collection.
            self.collectionList.InsertColumn(0, _("Collection"))
            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND, 0)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.collectionPanel.SetSizer(pnlVSizer)
            
        # If Clip Filtering is requested ...
        if self.kwargs.has_key('clipFilter') and self.kwargs['clipFilter']:
            # ... build a Panel for Clips ...
            self.clipsPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Clips Panel on the Notebook, creating a Clips tab ...
            self.notebook.AddPage(self.clipsPanel, _("Clips"))
            # ... place a Check List Ctrl on the Clips Panel ...
            self.clipList = CheckListCtrl(self.clipsPanel)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.clipList, 1, wx.EXPAND)
            # The Clip List needs two columns, Clip ID and Collection nesting.
            self.clipList.InsertColumn(0, _("Clip ID"))
            self.clipList.InsertColumn(1, _("Collection ID(s)"))

            # If Clip Sorting capacity has been requested ...
            if self.kwargs.has_key('clipSort') and self.kwargs['clipSort']:
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

        #If Keyword Group filter is requested...  (Kathleen)
        if self.kwargs.has_key('keywordGroupFilter') and self.kwargs['keywordGroupFilter']:
            # ... build a Panel for Keyword Groups ...
            self.keywordGroupPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Keyword Groups Panel on the Notebook, creating a Keyword Groups tab ...
            self.notebook.AddPage(self.keywordGroupPanel, _("Keyword Groups"))
            # Determine if Keyword Group Color specification is enabled or disabled.
            if self.kwargs.has_key('keywordGroupColor') and self.kwargs['keywordGroupColor']:
                self.keywordGroupColor = True
            else:
                self.keywordGroupColor = False
            # If Keyword Group Color Specification is enabled ...
            if self.keywordGroupColor:
                # ... we need to use the ColorListCtrl for the Keyword Groups List.
                self.keywordGroupList = ColorListCtrl.ColorListCtrl(self.keywordGroupPanel)
            # If Keyword Group Color specification is disabled ...
            else:
                # ... place a Check List Ctrl on the Keyword Group Panel ...
                self.keywordGroupList = CheckListCtrl(self.keywordGroupPanel)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.keywordGroupList, 1, wx.EXPAND)
            # The Keyword Groups List needs one column, Keyword Group.
            self.keywordGroupList.InsertColumn(0, _("Keyword Group"))
            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND, 0)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.keywordGroupPanel.SetSizer(pnlVSizer)
        
        # If Keyword Filtering is requested ...
        if self.kwargs.has_key('keywordFilter') and self.kwargs['keywordFilter']:
            # ... build a Panel for Keywords ...
            self.keywordsPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Keywords Panel on the Notebook, creating a Keywords tab ...
            self.notebook.AddPage(self.keywordsPanel, _("Keywords"))
            # Determine if Keyword Color specification is enabled or disabled.
            if self.kwargs.has_key('keywordColor') and self.kwargs['keywordColor']:
                self.keywordColor = True
            else:
                self.keywordColor = False
            # If Keyword Color Specification is enabled ...
            if self.keywordColor:
                # ... we need to use the ColorListCtrl for the Keywords List.
                self.keywordList = ColorListCtrl.ColorListCtrl(self.keywordsPanel)
            # If Keyword Color specification is disabled ...
            else:
                # ... place a Check List Ctrl on the Keywords Panel ...
                self.keywordList = CheckListCtrl(self.keywordsPanel)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.keywordList, 1, wx.EXPAND)
            # The keyword List needs two columns, Keyword Group and Keyword.
            self.keywordList.InsertColumn(0, _("Keyword Group"))
            self.keywordList.InsertColumn(1, _("Keyword"))

            # If Keyword Sorting capacity has been requested ...
            if self.kwargs.has_key('keywordSort') and self.kwargs['keywordSort']:
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
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND, 0)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.keywordsPanel.SetSizer(pnlVSizer)

        else:
            self.keywordColor = False
        
        #If Notes filter is requested...
        if self.kwargs.has_key('notesFilter') and self.kwargs['notesFilter']:
            # ... build a Panel for Notes ...
            self.notesPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Notes Panel on the Notebook, creating a Notes tab ...
            self.notebook.AddPage(self.notesPanel, _("Notes"))
            # ... place a Check List Ctrl on the Notes Panel ...
            self.notesList = CheckListCtrl(self.notesPanel)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.notesList, 1, wx.EXPAND)
            # The Notes List needs two columns, Note ID and Parent.
            self.notesList.InsertColumn(0, _("Note"))
            self.notesList.InsertColumn(1, _("Parent"))
            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND, 0)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.notesPanel.SetSizer(pnlVSizer)

        # If Options Specification is requested ...
        if self.kwargs.has_key('options') and self.kwargs['options']:
            # ... build a Panel for Options ...
            self.optionsPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            # ... place the Time Range Panel on the Notebook, creating a Time Range tab ...
            self.notebook.AddPage(self.optionsPanel, _("Options"))
            
            # This gets a bit convoluted, as different reports can have different options.  But it shouldn't be too bad.
            # Keyword Map Report and the Series Keyword Sequence Map
            # have Start and End times on the Options tab
            if self.reportType in [1, 5]:
                # Add a label for the Start Time field
                startTimeTxt = wx.StaticText(self.optionsPanel, -1, _("Start Time"))
                pnlVSizer.Add(startTimeTxt, 0, wx.TOP | wx.LEFT, 10)
                # Add the Start Time field
                self.startTime = wx.TextCtrl(self.optionsPanel, -1, self.startTimeVal)
                pnlVSizer.Add(self.startTime, 0, wx.LEFT, 10)
                # Add a label for the End Time field
                endTimeTxt = wx.StaticText(self.optionsPanel, -1, _("End Time"))
                pnlVSizer.Add(endTimeTxt, 0, wx.TOP | wx.LEFT, 10)
                # Add the End Time field
                self.endTime = wx.TextCtrl(self.optionsPanel, -1, self.endTimeVal)
                pnlVSizer.Add(self.endTime, 0, wx.LEFT, 10)
                # Add a note that says this data does not get saved.
                tRTxt = wx.StaticText(self.optionsPanel, -1, _("NOTE:  Setting the End Time to 0 will set it to the end of the Media File.\nTime Range data is not saved as part of the Filter Configuration data."))
                pnlVSizer.Add(tRTxt, 0, wx.ALL, 10)

            # Keyword Map Report, Keyword Visualization,and the Series Keyword Sequence Map
            # have Bar height and Whitespace parameters as well as horizontal and vertical grid lines
            if self.reportType in [1, 2, 5, 6, 7]:
                if self.kwargs.has_key('barHeight'):
                    # Add a label for the Bar Height field
                    barHeightTxt = wx.StaticText(self.optionsPanel, -1, _("Keyword Bar Height"))
                    pnlVSizer.Add(barHeightTxt, 0, wx.TOP | wx.LEFT, 10)
                    # Create a list of options for Bar Height
                    barHeightOptions = []
                    for x in range(2, 21):
                        barHeightOptions.append(str(x))
                    # Create a Choice Control of bar heights
                    self.barHeight = wx.Choice(self.optionsPanel, -1, choices=barHeightOptions)
                    # Set the initial value of the bar height
                    if self.kwargs.has_key('barHeight'):
                        self.barHeight.SetStringSelection(str(self.kwargs['barHeight']))
                    pnlVSizer.Add(self.barHeight, 0, wx.LEFT, 10)

                if self.kwargs.has_key('whitespace'):
                    # Add a label for the Whitespace field
                    whitespaceTxt = wx.StaticText(self.optionsPanel, -1, _("Space Between Bars"))
                    pnlVSizer.Add(whitespaceTxt, 0, wx.TOP | wx.LEFT, 10)
                    # Create a list of options for whitespace
                    whitespaceOptions = []
                    for x in range(0, 6):
                        whitespaceOptions.append(str(x))
                    # Create a Choice Control of whitespace
                    self.whitespace = wx.Choice(self.optionsPanel, -1, choices=whitespaceOptions)
                    # Set the initial value of the whitespace
                    if self.kwargs.has_key('whitespace'):
                        self.whitespace.SetStringSelection(str(self.kwargs['whitespace']))
                    pnlVSizer.Add(self.whitespace, 0, wx.LEFT, 10)

                if self.kwargs.has_key('hGridLines'):
                    # Add a check box for Horizontal Grid Lines
                    self.hGridLines = wx.CheckBox(self.optionsPanel, -1, _("Horizontal Grid Lines"))
                    if self.kwargs.has_key('hGridLines') and self.kwargs['hGridLines']:
                        self.hGridLines.SetValue(True)
                    pnlVSizer.Add(self.hGridLines, 0, wx.TOP | wx.LEFT, 10)

                if self.kwargs.has_key('vGridLines'):
                    # Add a check box for Vertical Grid Lines
                    self.vGridLines = wx.CheckBox(self.optionsPanel, -1, _("Vertical Grid Lines"))
                    if self.kwargs.has_key('vGridLines') and self.kwargs['vGridLines']:
                        self.vGridLines.SetValue(True)
                    pnlVSizer.Add(self.vGridLines, 0, wx.TOP | wx.LEFT, 10)

            # If we have a Series Keyword Sequence Map
            if self.reportType in [5] and self.kwargs.has_key('singleLineDisplay'):
                # ... add a check box for the Single Line Display option
                self.singleLineDisplay = wx.CheckBox(self.optionsPanel, -1, _("Single-line display"))
                self.singleLineDisplay.SetValue(self.kwargs['singleLineDisplay'])
                pnlVSizer.Add(self.singleLineDisplay, 0, wx.TOP | wx.LEFT, 10)

            # If we have a Series Keyword Sequence Map, a Series Keyword Bar Graph, or a Series Keyword
            # Percentage Graph ...
            if self.reportType in [5, 6, 7] and self.kwargs.has_key('showLegend'):
                # ... add a check box for the Show Legend option
                self.showLegend = wx.CheckBox(self.optionsPanel, -1, _("Show Legend"))
                self.showLegend.SetValue(self.kwargs['showLegend'])
                pnlVSizer.Add(self.showLegend, 0, wx.TOP | wx.LEFT, 10)

            # If the Single Line Display option is present ...
            if self.reportType in [5] and self.kwargs.has_key('singleLineDisplay'):
                # ... and is un-checked ...
                if not self.kwargs['singleLineDisplay']:
                    # ... then disable the Show Legend option.  There's no Legend for the multi-line display!
                    self.showLegend.Enable(False)
                # Also, connect the Single Line Display with an event that handles enabling and disabling the Legend option
                self.Bind(wx.EVT_CHECKBOX, self.OnSingleLineDisplay, self.singleLineDisplay)

            # If we have a Keyword Map, a Series Keyword Sequence Map, a Series Keyword Bar Graph, or 
            # a Series Keyword Percentage Graph ...
            if self.reportType in [1, 5, 6, 7] and self.kwargs.has_key('colorOutput'):
                # ... add a check box for Color (vs. GrayScale) output
                self.colorOutput = wx.CheckBox(self.optionsPanel, -1, _("Color Output"))
                self.colorOutput.SetValue(self.kwargs['colorOutput'])
                pnlVSizer.Add(self.colorOutput, 0, wx.TOP | wx.LEFT, 10)

            # Now declare the panel's vertical sizer as the panel's official sizer
            self.optionsPanel.SetSizer(pnlVSizer)

        # Place the Notebook in the Dialog Horizontal Sizer
        hBox.Add(self.notebook, 1, wx.EXPAND, 10)
        # Now place the Dialog's horizontal Sizer in the Dialog's Vertical sizer so we can expand in two dimensions
        vBox.Add(hBox, 1, wx.EXPAND, 10)

        # Create another horizontal sizer for the Dialog's buttons
        btnBox = wx.BoxSizer(wx.HORIZONTAL)
        # Create the OK button
        self.btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Make this the default button
        self.btnOK.SetDefault()
        # Add the OK button to the dialog's Button sizer
        btnBox.Add(self.btnOK, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT, 10)
        # Bind an event to OK.  (We need to override the default behavior or just closing.)
        self.btnOK.Bind(wx.EVT_BUTTON, self.OnOK)
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
        # ... check to see that an reportScope parameter was passed ...
        if self.kwargs.has_key('reportScope'):
            # ... and return the Episode Number as the Report's Scope.
            reportScope = self.kwargs['reportScope']
        
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
                # The default config name has been saved in English.  We need to translate it!
                if configName == 'Default':
                    configName = unicode(_('Default'), TransanaGlobal.encoding)
                # Decode the data
#                configName = DBInterface.ProcessDBDataForUTF8Encoding(configName)
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
            # If we're auto-loading the Default Config, we don't want to prompt for a Config Name.
            if self.loadDefault:
                # Fake that the user was shown a dialog and pressed OK
                result = wx.ID_OK
                # We'll use the Default config
                configName = 'Default'
                # Remember the Configuration Name
                self.configName = unicode(_(configName), TransanaGlobal.encoding)
            # If we're coming here through normal channels, we DO want to prompt for a config name.
            else:
                # Get a list of legal Report Names from the Database
                configNames = self.GetConfigNames()
                # Create a Choice Dialog so the user can select the Configuration they want
                dlg = FilterLoadDialog(self, _("Choose a Configuration to load"), _("Filter Configuration"), configNames, self.reportType, reportScope)
                # Center the Dialog on the screen
                dlg.CentreOnScreen()
                # Show the dialog and get the results.
                result = dlg.ShowModal()
                # Get the config name the user chose.
                configName = dlg.GetStringSelection()
                # Remember the Configuration Name
                self.configName = configName
                # Destroy the Choice Dialog
                dlg.Destroy()
            
            # Show the Choice Dialog and see if the user chooses OK
            if result == wx.ID_OK:
                # Set the Cursor to the Hourglass while the filter is loaded
                TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
                # Update the Window title to reflect the config name
                title = self.title + ' - ' + self.configName
                self.SetTitle(title)

                # Get a Database Cursor
                DBCursor = DBInterface.get_db().cursor()
                # Build a query to get the selected report's data.  The Order By clause is
                # necessary so that when keyword colors are updated, they will have been loaded
                # by the time the list in the keyword control is re-populated by the KeywordList
                # (that is, we need FilterDataType 4 to be processed BEFORE FilterDataType 3)
                query = """ SELECT FilterDataType, FilterData FROM Filters2
                              WHERE ReportType = %s AND
                                    ReportScope = %s AND
                                    ConfigName = %s
                              ORDER BY FilterDataType DESC"""
                # Build the data values that match the query
                values = (self.reportType, reportScope, configName.encode(TransanaGlobal.encoding))
                # Execute the query with the appropriate data values
                DBCursor.execute(query, values)
                # We may get multiple records, one for each tab on the Filter Dialog.
                for (filterDataType, filterData) in DBCursor.fetchall():
                    # If the data is for the Episodes Tab (filterDataType 1) ...
                    if filterDataType == 1:
                        # Get the current Episode data from the Form
                        formEpisodeData = self.GetEpisodes()
                        # Get the Episode data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileEpisodeData = cPickle.loads(filterData.tostring())
                        else:
                            fileEpisodeData = cPickle.loads(filterData)
                        # Clear the Episode List
                        self.episodeList.DeleteAllItems()
                        # Determine if this list is Ordered
                        if self.kwargs.has_key('episodeSort') and self.kwargs['episodeSort']:
                            orderedList = True
                        else:
                            orderedList = False
                        # We need to compare the file data to the form data and reconcile differences,
                        # then feed the results to the Episode Tab.
                        self.SetEpisodes(self.ReconcileLists(formEpisodeData, fileEpisodeData, listIsOrdered=orderedList))
                        
                    # If the data is for the Clips Tab (filterDataType 2) ...
                    elif filterDataType == 2:
                        # Get the current Clip data from the Form
                        formClipData = self.GetClips()
                        # Due to the BLOB / LONGBLOB problem, we can have bad data in the database.  Better catch it.
                        try:
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
                        # If the pickled data got truncated in the database, we'll get an Unpickling error here!
                        except cPickle.UnpicklingError, e:
                            # Construct and display an error message here
                            errormsg = _("Transana was unable to load your clip filter data from the database.  Please select your clips again and re-save the filter configuration.")
                            errorDlg = Dialogs.ErrorDialog(self, errormsg)
                            errorDlg.ShowModal()
                            errorDlg.Destroy()
                        
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

                    # If the data is for the Keyword Colors (filterDataType 4) ...
                    elif filterDataType == 4:
                        # Get the Keyword data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileKeywordColorData = cPickle.loads(filterData.tostring())
                        else:
                            fileKeywordColorData = cPickle.loads(filterData)
                        # Now over-ride existing Keyword Color data with the data from the file
                        for kwPair in fileKeywordColorData.keys():
                            self.keywordColors[kwPair] = fileKeywordColorData[kwPair]
                    
                    # If the data is for the Keywords Group Tab (filterDataType 5) ...
                    elif filterDataType == 5:
                        # Get the current Keyword Group data from the Form
                        formKeywordGroupData = self.GetKeywordGroups()
                        # Get the Keyword data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileKeywordGroupData = cPickle.loads(filterData.tostring())
                        else:
                            fileKeywordGroupData = cPickle.loads(filterData)
                        # Clear the Keyword Group List
                        self.keywordGroupList.DeleteAllItems()
                        # Determine if this list is Ordered
                        if self.kwargs.has_key('keywordGroupSort') and self.kwargs['keywordGroupSort']:
                            orderedList = True
                        else:
                            orderedList = False
                        # We need to compare the file data to the form data and reconcile differences,
                        # then feed the results to the Keyword Tab.
                        self.SetKeywordGroups(self.ReconcileLists(formKeywordGroupData, fileKeywordGroupData, listIsOrdered=orderedList))
                    
                    # If the data is for the Transcripts Tab (filterDataType 6) ...
                    elif filterDataType == 6:
                        # Get the current Transcript data from the Form
                        formTranscriptData = self.GetTranscripts()
                        # Get the Transcript data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileTranscriptData = cPickle.loads(filterData.tostring())
                        else:
                            fileTranscriptData = cPickle.loads(filterData)
                        # Clear the Transcript List
                        self.TranscriptList.DeleteAllItems()
                        # Determine if this list is Ordered
                        if self.kwargs.has_key('transcriptSort') and self.kwargs['transcriptSort']:
                            orderedList = True
                        else:
                            orderedList = False
                        # We need to compare the file data to the form data and reconcile differences,
                        # then feed the results to the Keyword Tab.
                        self.SetTranscripts(self.ReconcileLists(formTranscriptData, fileTranscriptData, listIsOrdered=orderedList))
                    
                    # If the data is for the Collections Tab (filterDataType 7) ...
                    elif filterDataType == 7:
                        # Get the current Collection data from the Form
                        formCollectionData = self.GetCollections()
                        # Get the Collection data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileCollectionData = cPickle.loads(filterData.tostring())
                        else:
                            fileCollectionData = cPickle.loads(filterData)
                        # Clear the Collections List
                        self.collectionList.DeleteAllItems()
                        # Determine if this list is Ordered
                        if self.kwargs.has_key('CollectionSort') and self.kwargs['CollectionSort']:
                            orderedList = True
                        else:
                            orderedList = False
                        # We need to compare the file data to the form data and reconcile differences,
                        # then feed the results to the Keyword Tab.
                        self.SetCollections(self.ReconcileLists(formCollectionData, fileCollectionData, listIsOrdered=orderedList))
                        
                    # If the data is for the Notes Tab (filterDataType 8) ...
                    elif filterDataType == 8:
                        # Get the current Notes data from the Form
                        formNotesData = self.GetNotes()
                        # Get the Notes data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileNotesData = cPickle.loads(filterData.tostring())
                        else:
                            fileNotesData = cPickle.loads(filterData)

                        # The Notes List is special, in that it contains the object type as part of the data
                        # that is displayed to the user.  This has to be localized, or Notes Filters won't work
                        # if loaded in a language different than they were saved in.  (The Save translates this
                        # part of the data to English.)
                        
                        # We need to iterate through the list
                        for recNum in range(len(fileNotesData)):
                            # See what type of record we have and replace English with the localized version
                            if fileNotesData[recNum][2][:len('Series')] == 'Series':
                                fileNotesData[recNum] = (fileNotesData[recNum][0], fileNotesData[recNum][1], unicode(_('Series'), 'utf8') + fileNotesData[recNum][2][len('Series'):], fileNotesData[recNum][3])
                            elif fileNotesData[recNum][2][:len('Episode')] == 'Episode':
                                fileNotesData[recNum] = (fileNotesData[recNum][0], fileNotesData[recNum][1], unicode(_('Episode'), 'utf8') + fileNotesData[recNum][2][len('Episode'):], fileNotesData[recNum][3])
                            elif fileNotesData[recNum][2][:len('Transcript')] == 'Transcript':
                                fileNotesData[recNum] = (fileNotesData[recNum][0], fileNotesData[recNum][1], unicode(_('Transcript'), 'utf8') + fileNotesData[recNum][2][len('Transcript'):], fileNotesData[recNum][3])
                            elif fileNotesData[recNum][2][:len('Collection')] == 'Collection':
                                fileNotesData[recNum] = (fileNotesData[recNum][0], fileNotesData[recNum][1], unicode(_('Collection'), 'utf8') + fileNotesData[recNum][2][len('Collection'):], fileNotesData[recNum][3])
                            elif fileNotesData[recNum][2][:len('Clip')] == 'Clip':
                                fileNotesData[recNum] = (fileNotesData[recNum][0], fileNotesData[recNum][1], unicode(_('Clip'), 'utf8') + fileNotesData[recNum][2][len('Clip'):], fileNotesData[recNum][3])

                        # Clear the Notes List
                        self.notesList.DeleteAllItems()
                        # Determine if this list is Ordered
                        if self.kwargs.has_key('NotesSort') and self.kwargs['NotesSort']:
                            orderedList = True
                        else:
                            orderedList = False
                        # We need to compare the file data to the form data and reconcile differences,
                        # then feed the results to the Notes Tab.
                        self.SetNotes(self.ReconcileLists(formNotesData, fileNotesData, listIsOrdered=orderedList))

                    # If the data is for the Start Time (filterDataType 9) ...
                    elif filterDataType == 9:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Only do this if startTime exists!  (Copied configs from ReportType 5 may have this.)
                        if self.kwargs.has_key('startTime'):
                            # Set the start time value
                            self.startTime.SetValue(filterData)

                    # If the data is for the End Time (filterDataType 10) ...
                    elif filterDataType == 10:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Only do this if endTime exists!  (Copied configs from ReportType 5 may have this.)
                        if self.kwargs.has_key('endTime'):
                            # Set the end time value
                            self.endTime.SetValue(filterData)

                    # If the data is for the Bar Height (filterDataType 11) ...
                    elif filterDataType == 11:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the bar height value
                        self.barHeight.SetStringSelection(filterData)

                    # If the data is for the Bar White Space (filterDataType 12) ...
                    elif filterDataType == 12:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the white space value
                        self.whitespace.SetStringSelection(filterData)

                    # If the data is for the Horizontal Grid Lines (filterDataType 13) ...
                    elif filterDataType == 13:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Horizontal Grid Lines value
                        self.hGridLines.SetValue(filterData == 'True')

                    # If the data is for the Vertical Grid Lines (filterDataType 14) ...
                    elif filterDataType == 14:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Vertical Grid Lines value
                        self.vGridLines.SetValue(filterData == 'True')

                    # If the data is for the Single Line Display (filterDataType 15) ...
                    elif filterDataType == 15:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Only do this if singleLineDisplay exists!  (Copied configs from ReportType 5 may have this.)
                        if self.kwargs.has_key('singleLineDisplay'):
                            # Set the Single Line Display value
                            self.singleLineDisplay.SetValue(filterData == 'True')

                    # If the data is for the Show Legend value (filterDataType 16) ...
                    elif filterDataType == 16:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Legend value
                        self.showLegend.SetValue(filterData == 'True')

                    # If the data is for the Color Output value (filterDataType 17) ...
                    elif filterDataType == 17:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Color Output value
                        self.colorOutput.SetValue(filterData == 'True')

                    # If the data is for the Include Nested Collection Data value (filterDataType 101) ...
                    elif filterDataType == 101:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Nested Collection Data value
                        self.showNestedData.SetValue(filterData == 'True')

                    # If the data is for the Show Media Filename Data value (filterDataType 102) ...
                    elif filterDataType == 102:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Media Filename Data value
                        self.showFile.SetValue(filterData == 'True')

                    # If the data is for the Show Clip Time Data value (filterDataType 103) ...
                    elif filterDataType == 103:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Clip Time Data value
                        self.showTime.SetValue(filterData == 'True')

                    # If the data is for the Show Clip Transcripts value (filterDataType 104) ...
                    elif filterDataType == 104:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Clip Transcripts value
                        self.showClipTranscripts.SetValue(filterData == 'True')

                    # If the data is for the Show Clip Keywords value (filterDataType 105) ...
                    elif filterDataType == 105:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Clip Keywords value
                        self.showClipKeywords.SetValue(filterData == 'True')

                    # If the data is for the Show Comments value (filterDataType 106) ...
                    elif filterDataType == 106:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Comments value
                        self.showComments.SetValue(filterData == 'True')

                    # If the data is for the Show Collection Notes value (filterDataType 107) ...
                    elif filterDataType == 107:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Collection Notes value
                        self.showCollectionNotes.SetValue(filterData == 'True')

                    # If the data is for the Show Clip Notes value (filterDataType 108) ...
                    elif filterDataType == 108:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Clip Notes value
                        self.showClipNotes.SetValue(filterData == 'True')

                    # If we have an unknown filterDataType ...
                    else:
                        print "Unknown Filter Data:", self.reportType, reportScope, self.configName, filterDataType, type(filterData), filterData
                # Set the Cursor to the Arrow now that the filter is loaded
                TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))


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
            # If Save was called through the OnOK event, we DON'T want to ask for a Config name.
            if event.GetId() == self.btnOK.GetId():
                # Fake that the user was given a dialog and pressed OK ...
                result = wx.ID_OK
                # ... and just keep the config name the same.
                configName = self.configName
            # If Save was called normally ...
            else:
                # Create a Dialog where the Configuration can be named
                dlg = wx.TextEntryDialog(self, _("Save Configuration As"), _("Filter Configuration"), originalConfigName)
                # Get the results from the user
                result = dlg.ShowModal()
                # Get the config name.  Remove leading and trailing white space
                configName = dlg.GetValue().strip()
                # If this is the defaut config ...
                if configName.capitalize() == 'Default':
                    # ... make sure the case usage is standard.  (I don't dare try this with language other than English.)
                    configName = configName.capitalize()
                # Destroy the Dialog that allows the user to name the configuration
                dlg.Destroy()
            # Show the dialog to the user and see how they respond
            if result == wx.ID_OK:
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
                    dlg2 = Dialogs.QuestionDialog(self, prompt % configName)
                    # Center the dialog on the screen
                    dlg2.CentreOnScreen()
                    # Display the dialog and get the user's response
                    if dlg2.LocalShowModal() != wx.ID_YES:
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
                        # Update the Window title to reflect the config name
                        title = self.title + ' - ' + self.configName
                        self.SetTitle(title)
                        # Get a database cursor
                        DBCursor = DBInterface.get_db().cursor()
                        # Encode the Report Name if we're using Unicode
                        if 'unicode' in wx.PlatformInfo:
                            configName = self.configName.encode(TransanaGlobal.encoding)
                        else:
                            configName = self.configName
                        # We need to save the "Default" config name in ENGLISH for mixed-language environments.
                        # So if we are dealing with the Default config ...
                        if configName == _('Default'):
                            # ... convert the config name to English.
                            configName = 'Default'
                            
                        # Check to see if the Configuration record already exists
                        if DBInterface.record_match_count('Filters2',
                                                         ('ReportType', 'ReportScope', 'ConfigName'),
                                                         (self.reportType, reportScope, configName)) > 0:
                            # Update existing record.  Note that each report may generate multiple records in the database,
                            # FilterDataTypes
                            #   1 = Episodes,
                            #   2 = Clips,
                            #   3 = Keywords,
                            #   4 = Keyword Colors,
                            #   5 = Keyword Groups (not implemented -- Fix XMLExport and XMLImport when implementing!)
                            #   6 = Transcripts    (not implemented -- Fix XMLExport and XMLImport when implementing!)
                            #   7 = Collections    (not implemented -- Fix XMLExport and XMLImport when implementing!)
                            #   8 = Notes

                            # Build the Update Query for Data
                            query = """ UPDATE Filters2
                                          SET FilterData = %s
                                          WHERE ReportType = %s AND
                                                ReportScope = %s AND
                                                ConfigName = %s AND
                                                FilterDataType = %s """

                            # If we have a Series Keyword Sequence Map (reportType 5), or
                            # a Series Keyword Bar Graph (reportType 6), or
                            # a Series Keyword Percentage Graph (reportType 7), or
                            # a Series Report (reportType 10), or
                            # a Series Clip Data Export (reportType 14),
                            # update Episode Data (FilterDataType 1)
                            if self.reportType in [5, 6, 7, 10, 14]:
                                # Pickle the Episode Data
                                episodes = cPickle.dumps(self.GetEpisodes())
                                # Build the values to match the query, including the pickled Episode data
                                values = (episodes, self.reportType, reportScope, configName, 1)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Keyword Map (reportType 1), or
                            # an Episode Clip Data Export (reportType 3), or
                            # a Collection Clip Data Export (reportType 4), or
                            # a Series Keyword Sequence Map (reportType 5), or
                            # a Series Keyword Bar Graph (reportType 6), or
                            # a Series Keyword Percentage Graph (reportType 7),
                            # an Episode Report (reportType 11),
                            # a Collection Report (reportType 12), or
                            # a Series Clip Data Export (reportType 14),
                            # update Clip Data (FilterDataType 2)
                            if self.reportType in [1, 3, 4, 5, 6, 7, 11, 12, 14]:
                                # Pickle the Clip Data
                                clips = cPickle.dumps(self.GetClips())
                                # Build the values to match the query, including the pickled Clip data
                                values = (clips, self.reportType, reportScope, configName, 2)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Keyword Map (reportType 1), or
                            # a Keyword Visualization (reportType 2), or
                            # an Episode Clip Data Export (reportType 3), or
                            # a Collection Clip Data Export (reportType 4), or
                            # a Series Keyword Sequence Map (reportType 5), or
                            # a Series Keyword Bar Graph (reportType 6), or
                            # a Series Keyword Percentage Graph (reportType 7),
                            # a Series Report (reportType 10)
                            # an Episode Report (reportType 11),
                            # a Collection Report (reportType 12), or
                            # a Series Clip Data Export (reportType 14),
                            # update Keyword Data (FilterDataType 3)
                            if self.reportType in [1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 14]:
                                # Pickle the Keyword Data
                                keywords = cPickle.dumps(self.GetKeywords())
                                # Build the values to match the query, including the pickled Keyword data
                                values = (keywords, self.reportType, reportScope, configName, 3)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Keyword Visualization (reportType 2), or
                            # a Series Keyword Sequence Map (reportType 5), or
                            # a Series Keyword Bar Graph (reportType 6), or
                            # a Series Keyword Percentage Graph (reportType 7),
                            # update Keyword Color Data (FilterDataType 4)
                            if self.reportType in [2, 5, 6, 7]:
                                # Pickle the Keyword Colors Data
                                keywordColors = cPickle.dumps(self.GetKeywordColors())
                                # Build the values to match the query, including the pickled Keyword Colors data
                                values = (keywordColors, self.reportType, reportScope, configName, 4)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Notes Report (reportType 13)
                            # update Notes Data (FilterDataType 8)
                            if self.reportType in [13]:
                                # Pickle the Notes Data
                                notes = cPickle.dumps(self.GetNotes())
                                # Build the values to match the query, including the pickled Notes data
                                values = (notes, self.reportType, reportScope, configName, 8)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                        else:
                            # Insert new record.  Note that each report may generate up to 4 records in the database,
                            # FilterDataType 1 = Episodes, FilterDataType 2 = Clips, FilterDataType 3 = Keywords, 4 = Keyword Colors
                            
                            # Build the Insert Query for the Data
                            query = """ INSERT INTO Filters2
                                            (ReportType, ReportScope, ConfigName, FilterDataType, FilterData)
                                          VALUES
                                            (%s, %s, %s, %s, %s) """

                            # If we have a Series Keyword Sequence Map (reportType 5), or
                            # a Series Keyword Bar Graph (reportType 6), or
                            # a Series Keyword Percentage Graph (reportType 7),
                            # a Series Report (reportType 10), or
                            # a Series Clip Data Export (reportType 14),
                            # insert Episode Data (FilterDataType 1)
                            if self.reportType in [5, 6, 7, 10, 14]:
                                # Pickle the Episode Data
                                episodes = cPickle.dumps(self.GetEpisodes())
                                # Build the values to match the query, including the pickled Episode data
                                values = (self.reportType, reportScope, configName, 1, episodes)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Keyword Map (reportType 1), or
                            # an Episode Clip Data Export (reportType 3), or
                            # a Collection Clip Data Export (reportType 4), or
                            # a Series Keyword Sequence Map (reportType 5), or
                            # a Series Keyword Bar Graph (reportType 6), or
                            # a Series Keyword Percentage Graph (reportType 7),
                            # an Episode Report (reportType 11),
                            # a Collection Report (reportType 12), or
                            # a Series Clip Data Export (reportType 14),
                            # insert Clip Data (FilterDataType 2)
                            if self.reportType in [1, 3, 4, 5, 6, 7, 11, 12, 14]:
                                # Pickle the Clip Data
                                clips = cPickle.dumps(self.GetClips())
                                # Build the values to match the query, including the pickled Clip data
                                values = (self.reportType, reportScope, configName, 2, clips)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Keyword Map (reportType 1), or
                            # a Keyword Visualization (reportType 2), or
                            # an Episode Clip Data Export (reportType 3), or
                            # a Collection Clip Data Export (reportType 4), or
                            # a Series Keyword Sequence Map (reportType 5), or
                            # a Series Keyword Bar Graph (reportType 6), or
                            # a Series Keyword Percentage Graph (reportType 7),
                            # a Series Report (reportType 10)
                            # an Episode Report (reportType 11),
                            # a Collection Report (reportType 12), or
                            # a Series Clip Data Export (reportType 14),
                            # insert Keyword Data (FilterDataType 3)
                            if self.reportType in [1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 14]:
                                # Pickle the Keyword Data
                                keywords = cPickle.dumps(self.GetKeywords())
                                # Build the values to match the query, including the pickled Keyword data
                                values = (self.reportType, reportScope, configName, 3, keywords)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Keyword Visualization (reportType 2), or
                            # a Series Keyword Sequence Map (reportType 5), or
                            # a Series Keyword Bar Graph (reportType 6), or
                            # a Series Keyword Percentage Graph (reportType 7),
                            # insert Keyword Color Data (FilterDataType 4)
                            if self.reportType in [2, 5, 6, 7]:
                                # Pickle the Keyword Colors Data
                                keywordColors = cPickle.dumps(self.GetKeywordColors())
                                # Build the values to match the query, including the pickled Keyword Colors data
                                values = (self.reportType, reportScope, configName, 4, keywordColors)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)
                            # If we have a Notes Report (reportType 13)
                            # insert Notes Data (FilterDataType 8)
                            if self.reportType in [13]:
                                # The Notes List has some translated content!  Specifically, the first part of each Note's PARENT
                                # indicates what type of record the note is attached to in the LOCAL language.  We have to translate
                                # this part of the note's parent to English before saving it.  Otherwise, you can't save a configuration
                                # in one language and then use it in another.

                                # First, let's get the list of notes.
                                notesList = self.GetNotes()
                                # Now we need to iterate through the list
                                for recNum in range(len(notesList)):
                                    # See what type of record we have and replace the localized version with English
                                    if notesList[recNum][2][:len(unicode(_('Series'), 'utf8'))] == unicode(_('Series'), 'utf8'):
                                        notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Series' + notesList[recNum][2][len(unicode(_('Series'), 'utf8')):], notesList[recNum][3])
                                    elif notesList[recNum][2][:len(unicode(_('Episode'), 'utf8'))] == unicode(_('Episode'), 'utf8'):
                                        notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Episode' + notesList[recNum][2][len(unicode(_('Episode'), 'utf8')):], notesList[recNum][3])
                                    elif notesList[recNum][2][:len(unicode(_('Transcript'), 'utf8'))] == unicode(_('Transcript'), 'utf8'):
                                        notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Transcript' + notesList[recNum][2][len(unicode(_('Transcript'), 'utf8')):], notesList[recNum][3])
                                    elif notesList[recNum][2][:len(unicode(_('Collection'), 'utf8'))] == unicode(_('Collection'), 'utf8'):
                                        notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Collection' + notesList[recNum][2][len(unicode(_('Collection'), 'utf8')):], notesList[recNum][3])
                                    elif notesList[recNum][2][:len(unicode(_('Clip'), 'utf8'))] == unicode(_('Clip'), 'utf8'):
                                        notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Clip' + notesList[recNum][2][len(unicode(_('Clip'), 'utf8')):], notesList[recNum][3])

                                # Pickle the Notes Data
                                notes = cPickle.dumps(notesList)
                                # Build the values to match the query, including the pickled Notes data
                                values = (self.reportType, reportScope, configName, 8, notes)
                                # Execute the query with the appropriate data
                                DBCursor.execute(query, values)

                            # We need a debugging message if the save is requested for an unknown reportType
                            if self.reportType not in [1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 13, 14]:
                                tmpDlg = Dialogs.ErrorDialog(self, "FilterDialog.OnFileSave() doesn't yet implement Save for reportType %d." % self.reportType)
                                tmpDlg.ShowModal()
                                tmpDlg.Destroy()

                            # Close the cursor
                            DBCursor.close()

                        # Filter Data Options information may not already exist, so
                        # need to be checked individually.  These Filter Data Types include:
                            #   9 = Start Time
                            #  10 = End Time
                            #  11 = Keyword Bar Height
                            #  12 = Bar Whitespace
                            #  13 = Horizontal Grid Lines
                            #  14 = Vertical Grid Lines
                            #  15 = Single Line Display
                            #  16 = Show Legend
                            #  17 = Color Output
                            
                        # If we have a Keyword Map (reportType 1), or
                        # a Series Keyword Sequence Map (reportType 5), 
                        # insert Start Time data (FilterDataType 9)
                        # and End Time data (FilterDataType 10)
                        if self.reportType in [1, 5]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 9, self.startTime.GetValue())
                            self.SaveFilterData(self.reportType, reportScope, configName, 10, self.endTime.GetValue())

                        # If we have a Keyword Map (reportType 1), or
                        # a Keyword Visualization (reportType 2), or
                        # a Series Keyword Sequence Map (reportType 5), or
                        # a Series Keyword Bar Graph (reportType 6), or
                        # a Series Keyword Percentage Graph (reportType 7),
                        # insert Keyword Bar Height Data (FilterDataType 11),
                        # Space Between Bars data (FilterDataType 12),
                        # Horizontal Grid Lines data (FilterDataType 13), and
                        # Vertical Grid Lines data (FilterDataType 14)
                        if self.reportType in [1, 2, 5, 6, 7]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 11, self.barHeight.GetStringSelection())
                            self.SaveFilterData(self.reportType, reportScope, configName, 12, self.whitespace.GetStringSelection())
                            self.SaveFilterData(self.reportType, reportScope, configName, 13, self.hGridLines.IsChecked())
                            self.SaveFilterData(self.reportType, reportScope, configName, 14, self.vGridLines.IsChecked())

                        # a Series Keyword Sequence Map (reportType 5), 
                        # insert Single Line Display Data (FilterDataType 15)
                        if self.reportType in [5]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 15, self.singleLineDisplay.IsChecked())

                        # a Series Keyword Sequence Map (reportType 5), or
                        # a Series Keyword Bar Graph (reportType 6), or
                        # a Series Keyword Percentage Graph (reportType 7),
                        # insert Show Legend Data (FilterDataType 16)
                        if self.reportType in [5, 6, 7]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 16, self.showLegend.IsChecked())

                        # If we have a Keyword Map (reportType 1), or
                        # a Series Keyword Sequence Map (reportType 5), or
                        # a Series Keyword Bar Graph (reportType 6), or
                        # a Series Keyword Percentage Graph (reportType 7),
                        # insert Color output Data (FilterDataType 17)
                        if self.reportType in [1, 5, 6, 7]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 17, self.colorOutput.IsChecked())

                        # Report Content Options information may not already exist, so
                        # need to be checked individually.  These Filter Data Types include:
                            # 101 = Include Nested Collections
                            # 102 = Show Media File Names
                            # 103 = Show Clip Times
                            # 104 = Show Clip Transcripts
                            # 105 = Show Keywords
                            # 106 = Show Comments
                            # 107 = Show Collection Notes
                            # 108 = Show Clip Notes

                        # If we have a Collection Report (reportType 12), or
                        # a Collection Clip Data Export (reportType 4) AND reportScope != 0
                        #   (Collection Clip Data Export NOT from Root Collection Node), 
                        # insert Include Nested Collections Data (FilterDataType 101)
                        if (self.reportType in [12]) or ((self.reportType == 4) and (reportScope != 0)):
                            self.SaveFilterData(self.reportType, reportScope, configName, 101, self.showNestedData.IsChecked())

                        # If we have an Episode Report (reportType 11), or
                        # a Collection Report (reportType 12), 
                        # Show Clip Time (FilterDataType 103),
                        # Show Clip Transcripts (FilterDataType 104),
                        # Show Comments (FilterDataType 106), and
                        # Show Clip Notes (Filter Data Type 108)
                        if self.reportType in [11, 12]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 103, self.showTime.IsChecked())
                            self.SaveFilterData(self.reportType, reportScope, configName, 104, self.showClipTranscripts.IsChecked())
                            self.SaveFilterData(self.reportType, reportScope, configName, 106, self.showComments.IsChecked())
                            self.SaveFilterData(self.reportType, reportScope, configName, 108, self.showClipNotes.IsChecked())

                        # If we have a Series Report (reportType 10), or
                        # an Episode Report (reportType 11), or
                        # a Collection Report (reportType 12), 
                        # insert Show Media Filename Data (FilterDataType 102),
                        # Show Clip Time (FilterDataType 103),
                        # Show Clip Transcripts (FilterDataType 104),
                        # Show Clip Keywords (FilterDataType 105),
                        # Show Comments (FilterDataType 106), and
                        # Show Clip Notes (Filter Data Type 108)
                        if self.reportType in [10, 11, 12]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 102, self.showFile.IsChecked())
                            self.SaveFilterData(self.reportType, reportScope, configName, 105, self.showClipKeywords.IsChecked())

                         # If we have a Collection Report (reportType 12), 
                        # insert Show Collection Notes (Filter Data Type 107)
                        if self.reportType in [12]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 107, self.showCollectionNotes.IsChecked())

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

    def SaveFilterData(self, reportType, reportScope, configName, filterDataType, filterData):
        """  """
        # Check to see if the Configuration record already exists
        if DBInterface.record_match_count('Filters2',
                                         ('ReportType', 'ReportScope', 'ConfigName', 'FilterDataType'),
                                         (reportType, reportScope, configName, filterDataType)) > 0:
            # Build the Update Query for Data
            query = """ UPDATE Filters2
                          SET FilterData = %s
                          WHERE ReportType = %s AND
                                ReportScope = %s AND
                                ConfigName = %s AND
                                FilterDataType = %s """
            values = (filterData, reportType, reportScope, configName, filterDataType)
        else:
            query = """ INSERT INTO Filters2
                            (ReportType, ReportScope, ConfigName, FilterDataType, FilterData)
                          VALUES
                            (%s, %s, %s, %s, %s) """
            values = (reportType, reportScope, configName, filterDataType, filterData)
        # Get a database cursor
        DBCursor = DBInterface.get_db().cursor()
        # Execute the query with the appropriate data
        DBCursor.execute(query, values)
        # Close the cursor
        DBCursor.close()


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
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(_('Are you sure you want to delete Filter Configuration "%s"?'), 'utf8')
                else:
                    prompt = _('Are you sure you want to delete Filter Configuration "%s"?')
                dlg2 = Dialogs.QuestionDialog(self, prompt % localConfigName)
                if dlg2.LocalShowModal() == wx.ID_YES:
                    # If we're deleting the Default configuration ...
                    if localConfigName == unicode(_('Default'), TransanaGlobal.encoding):
                        # ... we need to convert the configuration name to English!
                        localConfigName = 'Default'
                    # Clear the global configuration name
                    self.configName = ''
                    # Update the Window title to reflect the disappearance of the config name
                    self.SetTitle(self.title)
                    # Get a Database Cursor
                    DBCursor = DBInterface.get_db().cursor()
                    # Build a query to delete the selected report
                    query = """ DELETE FROM Filters2
                                  WHERE ReportType = %s AND
                                        ReportScope = %s AND
                                        ConfigName = %s """
                    # Build the data values that match the query
                    values = (self.reportType, reportScope, localConfigName.encode(TransanaGlobal.encoding))
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
        if selectedTab == unicode(_("Episodes"), 'utf8'):
            # ... iterate through the Episodes in the Episode List
            for x in range(self.episodeList.GetItemCount()):
                # If the Episode List Item's checked status does not match the desired status ...
                if self.episodeList.IsChecked(x) != (btnID == T_CHECK_ALL):
                    # ... then toggle the item so it will match
                    self.episodeList.ToggleItem(x)
        # if we're looking at the Transcripts tab ...
        elif selectedTab == unicode(_("Transcripts"), 'utf8'):
            # ... iterate through the Transcript items in the Transcript List
            for x in range(self.transcriptList.GetItemCount()):
                # If the Transcript List Item's checked status does not match the desired status ...
                if self.transcriptList.IsChecked(x) != (btnID == T_CHECK_ALL):
                    # ... then toggle the item so it will match
                    self.transcriptList.ToggleItem(x)
        # if we're looking at the Collections tab ...
        elif selectedTab == unicode(_("Collections"), 'utf8'):
            # ... iterate through the Collection items in the Collection List
            for x in range(self.collectionList.GetItemCount()):
                # If the Collection List Item's checked status does not match the desired status ...
                if self.collectionList.IsChecked(x) != (btnID == T_CHECK_ALL):
                    # ... then toggle the item so it will match
                    self.collectionList.ToggleItem(x)
        # if we're looking at the Clips tab ...
        elif selectedTab == unicode(_("Clips"), 'utf8'):
            # ... iterate through the Clips in the Clip List
            for x in range(self.clipList.GetItemCount()):
                # If the Clip List Item's checked status does not match the desired status ...
                if self.clipList.IsChecked(x) != (btnID == T_CHECK_ALL):
                    # ... then toggle the item so it will match
                    self.clipList.ToggleItem(x)
        # if we're looking at the Keywords tab ...
        elif selectedTab == unicode(_("Keywords"), 'utf8'):
            # ... iterate through the Keywords in the Keyword List
            for x in range(self.keywordList.GetItemCount()):
                # If the Keyword List Item's checked status does not match the desired status ...
                if self.keywordList.IsChecked(x) != (btnID == T_CHECK_ALL):
                    # ... then toggle the item so it will match
                    self.keywordList.ToggleItem(x)
        # if we're looking at the Keywords Group tab ...
        elif selectedTab == unicode(_("Keyword Groups"), 'utf8'):
            # ... iterate through the Keyword Groups in the Keyword Group List
            for x in range(self.keywordGroupList.GetItemCount()):
                # If the Keyword Group List Item's checked status does not match the desired status ...
                if self.keywordGroupList.IsChecked(x) != (btnID == T_CHECK_ALL):
                    # ... then toggle the item so it will match
                    self.keywordGroupList.ToggleItem(x)
        # if we're looking at the Notes tab ...
        elif selectedTab == unicode(_("Notes"), 'utf8'):
            # ... iterate through the Notes in the Notes List
            for x in range(self.notesList.GetItemCount()):
                # If the Notes List Item's checked status does not match the desired status ...
                if self.notesList.IsChecked(x) != (btnID == T_CHECK_ALL):
                    # ... then toggle the item so it will match
                    self.notesList.ToggleItem(x)
                    
    def OnButton(self, event):
        """ Process Button Events for the Filter Dialog """
        # Get the ID of the button that triggered this event
        btnID = event.GetId()
        # The buttons don't always exist, which makes some of the "if" statements here a bit trickier.
        # Therefore, we determine what buttons should exist, what was pressed, and what list it should affect.
        # First, check the Episode-related sort buttons.
        if self.kwargs.has_key('episodeFilter') and (btnID in [self.btnEpUp.GetId(), self.btnEpDown.GetId()]):
            # If it's one of these, we're working on the Episode List
            listAffected = self.episodeList
            if btnID == self.btnEpUp.GetId():
                btnPressed = 'btnUp'
            else:
                btnPressed = 'btnDown'
        # Next, check the Keyword-related Sort Buttons.
        elif self.kwargs.has_key('keywordFilter') and (btnID in [self.btnKwUp.GetId(), self.btnKwDown.GetId()]):
            listAffected = self.keywordList
            if btnID == self.btnKwUp.GetId():
                btnPressed = 'btnUp'
            else:
                btnPressed = 'btnDown'
        # It it's not one of those, we have a coding problem!  In this case, we shouldn't do anything.
        else:
            btnPressed = 'Unknown'
            
        # If the Move Item Up button is pressed ...
        if btnPressed == 'btnUp':
            # Determine which item is selected (-1 is None)
            item = listAffected.GetNextItem(-1, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
            # If an item is selected ...
            if item > -1:
                # Extract the data from the List
                col1 = listAffected.GetItem(item, 0).GetText()
                col2 = listAffected.GetItem(item, 1).GetText()
                checked = listAffected.IsChecked(item)
                # We store the color code in the ItemData!
                colorSpec = listAffected.GetItemData(item)
                # Check to make sure there's room to move up
                if item > 0:
                    # Delete the current item
                    listAffected.DeleteItem(item)
                    # Insert a new item one position up and add the extracted data
                    listAffected.InsertStringItem(item - 1, col1)
                    listAffected.SetStringItem(item - 1, 1, col2)
                    if checked:
                        listAffected.CheckItem(item - 1)
                    # We store the color code in the ItemData!
                    listAffected.SetItemData(item - 1, colorSpec)
                    # Make sure the new item is selected
                    listAffected.SetItemState(item - 1, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)
                    # Make sure the item is visible
                    listAffected.EnsureVisible(item - 1)

        # If the Move Item Down button is pressed ...
        elif btnPressed == 'btnDown':
            # Determine which item is selected (-1 is None)
            item = listAffected.GetNextItem(-1, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
            # If an item is selected ...
            if item > -1:
                # Extract the data from the  List
                col1 = listAffected.GetItem(item, 0).GetText()
                col2 = listAffected.GetItem(item, 1).GetText()
                checked = listAffected.IsChecked(item)
                # We store the color code in the ItemData!
                colorSpec = listAffected.GetItemData(item)
                # Check to make sure there's room to move down
                if item < listAffected.GetItemCount() - 1:
                    # Delete the current item
                    listAffected.DeleteItem(item)
                    # Insert a new item one position down and add the extracted data
                    listAffected.InsertStringItem(item + 1, col1)
                    listAffected.SetStringItem(item + 1, 1, col2)
                    if checked:
                        listAffected.CheckItem(item + 1)
                    # We store the color code in the ItemData!
                    listAffected.SetItemData(item + 1, colorSpec)
                    # Make sure the new item is selected
                    listAffected.SetItemState(item + 1, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)
                    # Make sure the item is visible
                    listAffected.EnsureVisible(item + 1)

    def OnSingleLineDisplay(self, event):
        """ Implements GUI changes needed when the Single Line Display option is used """
        # If singleLineDisplay has been checked ...
        if self.singleLineDisplay.IsChecked():
            # ... then a Legend is an option.
            self.showLegend.Enable(True)
        # If singleLineDisplay is NOT checked ...
        else:
            # ... then the Legend is not available.
            self.showLegend.Enable(False)
            
    def GetShowClipTranscripts(self):
        """ Return the value of showClipTranscripts on the Report Contents tab """
        return self.showClipTranscripts.GetValue()

    def GetShowClipKeywords(self):
        """ Return the value of showClipKeywords on the Report Contents tab """
        return self.showClipKeywords.GetValue()

    def GetShowFile(self):
        """ Return the value of showFile on the Report Contents tab """
        return self.showFile.GetValue()

    def GetShowTime(self):
        """ Return the value of showTime on the Report Contents tab """
        return self.showTime.GetValue()

    def GetShowComments(self):
        """ Return the value of showComments on the Report Contents tab """
        return self.showComments.GetValue()

    def GetShowCollectionNotes(self):
        """ Return the value of showCollectionNotes on the Report Contents tab """
        return self.showCollectionNotes.GetValue()

    def GetShowClipNotes(self):
        """ Return the value of showClipNotes on the Report Contents tab """
        return self.showClipNotes.GetValue()

    def GetShowNestedData(self):
        """ Return the value of showNestedData on the Report Contents tab """
        return self.showNestedData.GetValue()

    def SetEpisodes(self, episodeList):
        """ Allows the calling routine to provide a list of Episodes that should be included on the Episodes Tab.
            A sorted list of (episodeID, seriesID, checked(boolean)) information should be passed in. """
        # Iterate through the Episode list that was passed in
        for (episodeID, seriesID, checked) in episodeList:
            # Create a new Item in the Episode List at the end of the list.  Add the Episode ID data.
            index = self.episodeList.InsertStringItem(sys.maxint, episodeID)
            # Add the Series data
            self.episodeList.SetStringItem(index, 1, seriesID)
            # If the item should be checked, check it!
            if checked:
                self.episodeList.CheckItem(index)
        # Set the column width for Episode IDs
        self.episodeList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Series Data
        self.episodeList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # lay out the Episode Panel
        self.episodesPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()

    def GetEpisodes(self):
        """ Allows the calling routine to retrieve the episode data from the Filter Dialog.  A sorted list
            of (episodeID, seriesID, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty Episode list
        episodeList = []
        # Iterate throught the EpisodeListCtrl items
        for x in range(self.episodeList.GetItemCount()):
            # Append the list item's data to the episode list
            episodeList.append((self.episodeList.GetItem(x, 0).GetText(), self.episodeList.GetItem(x, 1).GetText(), self.episodeList.IsChecked(x)))
        # Return the episode list
        return episodeList

    def SetTranscripts(self, transcriptList):
        """ Allows the calling routine to provide a list of Transcripts that should be included on the Transcripts Tab.
            A sorted list of (seriesID,episodeID, transcriptID, numClips,checked(boolean)) information should be passed in. """
        # Iterate through the Transcript list that was passed in
        for (seriesID, episodeID, transcriptID, numClips, checked) in transcriptList:
            # Create a new Item in the Transcript List at the end of the list.  Add the Transcript ID data.
            index = self.transcriptList.InsertStringItem(sys.maxint, seriesID)
            # Add the Episode data
            self.transcriptList.SetStringItem(index, 1, episodeID)
            # Add the Transcript data
            self.transcriptList.SetStringItem(index,2, transcriptID)
            # Add the NumClips data
            self.transcriptList.SetStringItem(index,3, str(numClips))
            # If the item should be checked, check it!
            if checked:
                self.transcriptList.CheckItem(index)
        # Set the column width for Episode IDs
        self.transcriptList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Transcript Data
        self.transcriptList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # Set the Column with for Number of Clips Data
        #self.transcriptList.SetColumnWith(2, wx.LIST_AUTOSIZE)
        # lay out the transcript Panel
        self.transcriptPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()

    def GetTranscripts(self):
        """ Allows the calling routine to retrieve the Transcript data from the Filter Dialog.  A sorted list
            of (episodeID, transcriptID, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty Transcript list
        transcriptList = []
        # Iterate throught the TranscriptListCtrl items
        for x in range(self.transcriptList.GetItemCount()):
            # Append the list item's data to the Transcript list
            transcriptList.append((self.transcriptList.GetItem(x, 0).GetText(), self.transcriptList.GetItem(x, 1).GetText(), self.transcriptList.GetItem(x,2).GetText(),self.transcriptList.GetItem(x,3).GetText(),self.transcriptList.IsChecked(x)))
        # Return the Transcript list
        return transcriptList
    
    def SetCollections(self, collectionList):
        """ Allows the calling routine to provide a list of Collections that should be included on the Collections Tab.
            A sorted list of (collectionID, checked(boolean)) information should be passed in. """
        # Iterate through the Collection list that was passed in
        for (collID, checked) in collectionList:
            # Create a new Item in the Collection List at the end of the list.  Add the Collection ID data.
            index = self.collectionList.InsertStringItem(sys.maxint, collID)
            # If the item should be checked, check it!
            if checked:
                self.collectionList.CheckItem(index)
        # Set the column width for Collection Name
        self.collectionList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        
        # lay out the Collections Panel
        self.collectionPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()
    
    def GetCollections(self):
        """ Allows the calling routine to retrieve the Collections data from the Filter Dialog.  A sorted list
            of (collectionID, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty Collections list
        collectionList = []
        # Iterate throught the CollectionListCtrl items
        for x in range(self.collectionList.GetItemCount()):
            # Append the list item's data to the Collections list
            collectionList.append((self.collectionList.GetItem(x, 0).GetText(),self.collectionList.IsChecked(x)))
        # Return the Collections list
        return collectionList
    
    def SetClips(self, clipList):
        """ Allows the calling routine to provide a list of Clips that should be included on the Clips Tab.
            A sorted list of (clipName, collectionNumber, checked(boolean)) information should be passed in. """
        # Save the original data.  We'll need it to retrieve the Collection Number data
        self.originalClipData = clipList
        # Iterate through the Clip list that was passed in
        for (clipID, collectNum, checked) in clipList:
            # Create a new Item in the Clip List at the end of the list.  Add the Clip ID data.
            index = self.clipList.InsertStringItem(sys.maxint, clipID)
            # Add the Collection data
            tempColl = Collection.Collection(collectNum)
            # Add the Collection's Node String
            self.clipList.SetStringItem(index, 1, tempColl.GetNodeString())
            # If the item should be checked, check it!
            if checked:
                self.clipList.CheckItem(index)
        # Set the column width for Clip IDs
        self.clipList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Collection Data
        self.clipList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # lay out the Clips Panel
        self.clipsPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()

    def GetClips(self):
        """ Allows the calling routine to retrieve the Clip data from the Filter Dialog.  A sorted list
            of (clipID, collectNum, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty Clip list
        clipList = []
        # Iterate throught the ClipListCtrl items
        for x in range(self.clipList.GetItemCount()):
            # Append the list item's data to the Clip list
            clipList.append((self.clipList.GetItem(x, 0).GetText(), self.originalClipData[x][1], self.clipList.IsChecked(x)))
        # Return the clip list
        return clipList

    def SetKeywordGroups(self, kwGroupList):
        """ Allows the calling routine to provide a list of keyword groups that should be included on the Keywords Tab.
            A sorted list of (keywordgroup, checked(boolean)) information should be passed in.
            If Keyword Group Colors are enabled, SetKeywordGroupColors() should be called before SetKeywordGroups(). """
        # Iterate through the keyword group list that was passed in
        for kwrec in kwGroupList:
            # Unpack the data record
            (kwg, checked) = kwrec
            # If keyword colors are enabled ...
            if self.keywordGroupColor:
                # ... get the keyword color data
                colorSpec = self.keywordColors[(kwg)]
            else:
                colorSpec = 0
            # Create a new Item in the Keyword Group List at the end of the list.  Add the Keyword Group data.
            index = self.keywordGroupList.InsertStringItem(sys.maxint, kwg)
            # If the item should be checked, check it!
            if checked:
                self.keywordGroupList.CheckItem(index)
            # We store the color code in the ItemData!
            self.keywordGroupList.SetItemData(index, colorSpec)
        # Set the column width for Keyword Groups
        self.keywordGroupList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # lay out the Keyword Groups Panel
        self.keywordGroupPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()
    
    def GetKeywordGroups(self):
        """ Allows the calling routine to retrieve the keyword Group data from the Filter Dialog.  A sorted list
            of (keywordgroup, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty keyword groups list
        keywordGroupList = []
        # Iterate throught the KeywordGroupListCtrl items
        for x in range(self.keywordGroupList.GetItemCount()):
            # Append the list item's data to the keyword Groups list
            keywordGroupList.append((self.keywordGroupList.GetItem(x, 0).GetText(), self.keywordGroupList.IsChecked(x)))
        # Return the keyword groups list
        return keywordGroupList
    
    def SetKeywords(self, kwList):
        """ Allows the calling routine to provide a list of keywords that should be included on the Keywords Tab.
            A sorted list of (keywordgroup, keyword, checked(boolean)) information should be passed in.
            If Keyword Colors are enabled, SetKeywordColors() should be called before SetKeywords(). """
        # Iterate through the keyword list that was passed in
        for kwrec in kwList:
            # Unpack the data record
            (kwg, kw, checked) = kwrec
            # If keyword colors are enabled ...
            if self.keywordColor:
                # ... get the keyword color data
                colorSpec = self.keywordColors[(kwg, kw)]
            else:
                colorSpec = 0
            # Create a new Item in the Keyword List at the end of the list.  Add the Keyword Group data.
            index = self.keywordList.InsertStringItem(sys.maxint, kwg)
            # Add the Keyword data
            self.keywordList.SetStringItem(index, 1, kw)
            # If the item should be checked, check it!
            if checked:
                self.keywordList.CheckItem(index)
            # We store the color code in the ItemData!
            self.keywordList.SetItemData(index, colorSpec)
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

    def SetKeywordColors(self, keywordColors):
        """ Allows the calling routine to provide a dictionary of keyword colors that should be indicated on the Keywords Tab.
            This method should be called prior to SetKeywords(). """
        self.keywordColors = keywordColors

    def GetKeywordColors(self):
        """ Allows the calling routine to retrieve the keyword color data from the Filter Dialog.  A dictionary
            object is returned. """
        # Iterate through the keyword list ...
        for x in range(self.keywordList.GetItemCount()):
            # add to or update the dictionary to reflect that for the (kwg, kw) key, the Item Data is the color value
            self.keywordColors[(self.keywordList.GetItem(x, 0).GetText(), self.keywordList.GetItem(x, 1).GetText())] = self.keywordList.GetItemData(x)
        # return the updated dictionary object
        return self.keywordColors

    def SetNotes(self, notesList):
        """ Allows the calling routine to provide a list of Notes that should be included on the Notes tab. """
        # Iterate through the note list that was passed in
        for (noteNum, noteID, noteParent, checked) in notesList:
            # Create a new Item in the Note List at the end of the list.  Add the Note ID data.
            index = self.notesList.InsertStringItem(sys.maxint, noteID)
            # Add the Note Parent data
            self.notesList.SetStringItem(index, 1, noteParent)
            # If the item should be checked, check it!
            if checked:
                self.notesList.CheckItem(index)
            # Add the note number to the note list's Item Data
            self.notesList.SetItemData(index, noteNum)
        # Set the column widths for Note Name and Parent 
        self.notesList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        self.notesList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        
        # lay out the Notes Panel
        self.notesPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()
    
    def GetNotes(self):
        """ Allows the calling routine to retrieve Notes data from the Filter Dialog.  """
        noteList = []
        # Iterate through the notes list ...
        for x in range(self.notesList.GetItemCount()):
            # build the list to return
            noteList.append((self.notesList.GetItemData(x), self.notesList.GetItem(x, 0).GetText(), self.notesList.GetItem(x, 1).GetText(), self.notesList.IsChecked(x)))
        # return the updated dictionary object
        return noteList
        
    def GetStartTime(self):
        """ Return the value of the Start Time text control on the Options tab """
        return self.startTime.GetValue()

    def GetEndTime(self):
        """ Return the value of the End Time text control on the Options tab """
        return self.endTime.GetValue()

    def GetBarHeight(self):
        """ Return the value of the Bar Heigth on the Options tab """
        return int(self.barHeight.GetStringSelection())

    def GetWhitespace(self):
        """ Return the value of the Whitespace on the Options tab """
        return int(self.whitespace.GetStringSelection())

    def GetHGridLines(self):
        """ Return the value of hGridLines on the Options tab """
        return self.hGridLines.GetValue()

    def GetVGridLines(self):
        """ Return the value of vGridLines on the Options tab """
        return self.vGridLines.GetValue()

    def GetSingleLineDisplay(self):
        """ Return the value of singleLineDisplay on the Option tab """
        return self.singleLineDisplay.GetValue()

    def GetShowLegend(self):
        """ Return the value of showLegend on the Options tab """
        return self.showLegend.GetValue()

    def GetColorOutput(self):
        """ Return the value of colorOutput on the Options tab """
        return self.colorOutput.GetValue()

    def OnOK(self, event):
        """ implement the Filter Dialog's OK button """
        # If we're using a Default configuration ...
        if self.configName == unicode(_('Default'), TransanaGlobal.encoding):
            # ... then we automatically save!
            self.OnFileSave(event)
        # Process the Dialog's default OK behavior
        event.Skip()

    def OnHelp(self, event):
        """ Implement the Filter Dialog Box's Help function """
        # Define the Help Context
        HelpContext = "Filter Dialog"
        # If a Help Window is defined ...
        if TransanaGlobal.menuWindow != None:
            # ... call Help!
            TransanaGlobal.menuWindow.ControlObject.Help(HelpContext)

class FilterLoadDialog(wx.Dialog):
    """ Emulate wx.SingleChoiceDialog() but with a button allowing the user to copy a configuration file. """
    def __init__(self, parent, prompt, title, configNames, reportType, reportScope):
        # Note the dialog title
        self.title = title
        # Remember the list items for later comparison
        self.configNames = configNames
        # Note the report type
        self.reportType = reportType
        # Note the scope of the current report
        self.reportScope = reportScope
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
        # If this is for the appropriate report type ...
        # (Copy Configuration has been implemented for Keyword Map, Keyword Visualization, Series Keyword Sequence Map,
        # Series Keyword Bar Graph, Series Keyword Percentage Graph, and Collection Report.)
        if self.reportType in [1, 2, 5, 6, 7]:
            # ... display the Copy Configuration button
            btnCopy = wx.Button(self, -1, _("Copy Configuration"))
            # Give the Copy Button an event handler
            btnCopy.Bind(wx.EVT_BUTTON, self.OnCopyConfig)
            # put the Copy button on the Button Sizer
            box2.Add(btnCopy, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 10)
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

    def OnCopyConfig(self, event):
        """ View and select a configuration from another selection """
        # Get a list of configuration names that can be copied
        configList = self.GetConfigNames()
        # Create the Configuration Copy dialog box
        dlg = FilterCopyConfigDialog(self, _("Select a configuration to copy:"), self.title, configList)
        # Display the dialog and get feedback from the user
        if dlg.ShowModal() == wx.ID_OK:
            # Get the data from the dialog
            returnVal = dlg.GetSelection()
            # If a problem came up, the return value is None.  Otherwise, we want to continue.
            if returnVal != None:
                # Break the return value into its components
                (copyReportType, copyReportScope, copyConfigName) = returnVal
                # Let's call the new configuration by the same name if we can.  If we can't, we'll change this later.
                newConfigName = copyConfigName

                # First let's get the configuration data from the database
                # Get a Database Cursor
                DBCursor = DBInterface.get_db().cursor()
                # Different reportTypes have different rules for what to copy.
                # For the Keyword Map and the Keyword Visualization ...
                if self.reportType in [1, 2]:
                    # Build the database query.  We do not duplicate 
                    # FilterDataType 1 (Episodes) or FilterDataType 2 (Clips).  Different Scopes will have the same Keywords and
                    # Colors, but never the same Episodes or Clips.
                    query = """ SELECT ReportType, ReportScope, ConfigName, FilterDataType, FilterData
                                  FROM Filters2
                                  WHERE ReportType = %s AND
                                        ReportScope = %s AND
                                        ConfigName = %s AND
                                        FilterDataType <> 1 AND
                                        FilterDataType <> 2 """
                    # Set up the data values that match the query
                    values = (self.reportType, copyReportScope, copyConfigName.encode(TransanaGlobal.encoding))
                # For the Keyword Series Sequence Map, the Keyword Series Bar Graph, and the Keyword Series Percentage Graph ...
                elif self.reportType in [5, 6, 7]:
                    # Build the database query.  We can duplicate all FilterDataTypes.
                    query = """ SELECT ReportType, ReportScope, ConfigName, FilterDataType, FilterData
                                  FROM Filters2
                                  WHERE ReportType = %s AND
                                        ReportScope = %s AND
                                        ConfigName = %s"""
                    # Set up the data values that match the query
                    values = (copyReportType, copyReportScope, copyConfigName.encode(TransanaGlobal.encoding))

                # Execute the query with the data values
                DBCursor.execute(query, values)
                
                # Assume the process should not be interrupted until proven otherwise.
                cont = True

                # First, let's see if the Config Name is already being used.  The user may change it, but doesn't have to.
                if newConfigName in self.configNames:
                    # Build a dialog to prompt for a new name
                    dlg2 = wx.TextEntryDialog(self, _("Duplicate Configuration Name.  Please enter a new name, or press OK to overwrite the existing Configuration."),
                                              _("Transana Information"), newConfigName, style= wx.OK | wx.CANCEL | wx.CENTRE)
                    # Center it on the screen
                    dlg2.CentreOnScreen()
                    # Prompt the user.
                    if dlg2.ShowModal() == wx.ID_OK:
                        newConfigName = dlg2.GetValue()
                    # If the user presses Cancel ...
                    else:
                        # ... we need to signal that the duplication process should be interrupted!
                        cont = False

                # If we haven't indicated that we should stop ...
                if cont:
                    # ... then we can iterate through the query results and process them.
                    for (rowReportType, rowReportScope, rowConfigName, rowFilterDataType, rowFilterData) in DBCursor.fetchall():
                        # The name could still be the same, or it could be different.  We need to Insert or Update, depending.
                        if newConfigName in self.configNames:
                            query = """ UPDATE Filters2
                                          SET FilterData = %s
                                          WHERE ReportType = %s AND
                                                ReportScope = %s AND
                                                ConfigName = %s AND
                                                FilterDataType = %s """
                            values = (rowFilterData, self.reportType, self.reportScope, newConfigName.encode(TransanaGlobal.encoding), rowFilterDataType)
                        else:
                            query = """ INSERT INTO Filters2
                                            (ReportType, ReportScope, ConfigName, FilterDataType, FilterData)
                                          VALUES
                                            (%s, %s, %s, %s, %s) """
                            values = (self.reportType, self.reportScope, newConfigName.encode(TransanaGlobal.encoding), rowFilterDataType, rowFilterData)
                        # Execute the query with the data values
                        DBCursor.execute(query, values)

                    if not newConfigName in self.configNames:
                        # Add the new Config Name to the dialog box's list of configurations ...
                        self.choices.InsertItems([newConfigName], self.choices.GetCount())
                        # ... and select it
                        self.choices.SetSelection(self.choices.GetCount() - 1)
                        # Also add the new config name to self.configNames so it can be detected as a duplicate if the user
                        # repeats the process!
                        self.configNames.append(newConfigName)
                    else:
                        # Select the just-copied configuration
                        self.choices.SetSelection(self.choices.FindString(newConfigName))

        # Destroy the Configuration Copy dialog box
        dlg.Destroy()

    def GetConfigNames(self):
        """ Get a list of all Configuration Names for the current Report Type that excludes the current Report Scope. """
        # Create a blank list
        resList = []
        # Initialize a query variable
        query = ''
        # Get a Database Cursor
        DBCursor = DBInterface.get_db().cursor()
        # For Report Types 1 and 2, Keyword Map and Keyword Visualization, use an Episode-based Query.
        # The Report Type must be the SAME.
        # The Report Scope must be DIFFERENT.
        if self.reportType in [1, 2]:
            # Build the database query
            query = """ SELECT s.SeriesID, e.EpisodeID, f.ReportType, f.ReportScope, f.ConfigName
                          FROM Series2 s, Episodes2 e, Filters2 f
                          WHERE ReportType = %s AND
                                ReportScope <> %s AND
                                EpisodeNum = ReportScope AND
                                e.SeriesNum = s.SeriesNum
                          GROUP BY ReportScope, ConfigName
                          ORDER BY SeriesID, EpisodeID, ConfigName """
            # Set up the data values that match the query
            values = (self.reportType, self.reportScope)
        # For Report Types 5, 6, and 7, Series Keyword Sequence Map, Series Keyword Bar Graph, and Series Keyword Percentage
        # Graph, use a Series-based Query.
        # The Report Type must be DIFFERENT.
        # The Report Scope must be the SAME.
        elif self.reportType in [5, 6, 7]:
            # Build the database query
            query = """ SELECT s.SeriesID, f.ReportType, f.ReportScope, f.ConfigName
                          FROM Series2 s, Filters2 f
                          WHERE ("""
            # We need the query to include reportTypes in [5, 6, 7] that differ from the current reportType.
            # Building this query correctly is a bit of a pain.  This is just brute force, not elegant.
            if self.reportType in [6, 7]:
                query += '(ReportType = 5) OR '
            if self.reportType in [5, 7]:
                query += '(ReportType = 6) '
            if self.reportType == 5:
                query += 'OR '
            if self.reportType in [5, 6]:
                query += '(ReportType = 7)'
            query += """       ) AND
                                ReportScope = %s AND
                                s.SeriesNum = ReportScope
                          GROUP BY ReportScope, ConfigName
                          ORDER BY SeriesID, ConfigName """
            # Set up the data values that match the query
            values = (self.reportScope)
            
        if query != '':
            # Execute the query with the data values
            DBCursor.execute(query, values)
            # Iterate through the report results
            for rowValues in DBCursor.fetchall():
                # Different queries produce different result sets (including or excluding Episode ID, depending on if it's relevant.)
                # For the Keyword Map and the Keyword Visualization, Episode ID is included.
                if self.reportType in [1, 2]:
                    # Decode the data
                    seriesID = DBInterface.ProcessDBDataForUTF8Encoding(rowValues[0])
                    episodeID = DBInterface.ProcessDBDataForUTF8Encoding(rowValues[1])
                    reportType = rowValues[2]
                    reportScope = rowValues[3]
                    configName = DBInterface.ProcessDBDataForUTF8Encoding(rowValues[4])
                    # Add the data to the Results list
                    resList.append((seriesID, episodeID, reportType, reportScope, configName))
                # for the Series Keyword Sequence Map, the Series Keyword Bar Graph, and the Series Keyword Percentage Graph,
                # Episode ID is not included because these are Series-based reports.
                elif self.reportType in [5, 6, 7]:
                    # Decode the data
                    seriesID = DBInterface.ProcessDBDataForUTF8Encoding(rowValues[0])
                    reportType = rowValues[1]
                    reportScope = rowValues[2]
                    configName = DBInterface.ProcessDBDataForUTF8Encoding(rowValues[3])
                    # Add the data to the Results list
                    resList.append((seriesID, reportType, reportScope, configName))
        # return the Results List
        return resList

class FilterCopyConfigDialog(wx.Dialog):
    """ Emulate wx.SingleChoiceDialog() but with a grid instead of a listbox. """
    def __init__(self, parent, prompt, title, configNames):
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

        # Display a list of choices the user can select from.  Building the control takes several steps.
        # First, use an Auto Width List Control
        self.choices = AutoWidthListCtrl(self, -1, size=(350, 200), style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_HRULES)
        # Currently, all configs are Episode Based.  This may need to be parameterized later.
        self.choices.InsertColumn(0, _('Series'))
        # We figure out if Episode Name is included by looking at the length of the data records sent.
        if (len(configNames) > 0) and len(configNames[0]) == 5:
            self.choices.InsertColumn(1, _('Episode'))
            self.choices.InsertColumn(2, _('Configuration Name'))
            # Note how many colums we have
            maxCol = 3
        else:
            self.choices.InsertColumn(1, _('Configuration Name'))
            # Note how many colums we have
            maxCol = 2

        # We'll need to be able to look up return data based on List Item Index.  We'll use a dictionary to hold the data.
        self.dataLookup = {}

        # For each row in the Configuration Names data passed in ...
        for row in configNames:
            # Create a Row, with the Series data in the first column
            index = self.choices.InsertStringItem(sys.maxint, row[0])
            # We figure out if Episode Name is included by looking at the length of the data records sent.
            if len(row) == 5:
                # Add the Episode Data to the second column
                self.choices.SetStringItem(index, 1, row[1])
                # Add the Configuration Name data to the third column
                self.choices.SetStringItem(index, 2, row[4])
                # The lookup data for this List Row needs to be the ReportScope data and the ConfigName
                self.dataLookup[index] = (row[2], row[3], row[4])
            else:
                # Add the Configuration Name data to the second column
                self.choices.SetStringItem(index, 1, row[3])
                # The lookup data for this List Row needs to be the ReportScope data and the ConfigName
                self.dataLookup[index] = (row[1], row[2], row[3])

        # We need to know how wide the window needs to be.  Start with room for the margins and the ScrollBar
        windowWidth = 40
        # For each column ...
        for x in range(maxCol):
            # Set the width to the widest item ...
            self.choices.SetColumnWidth(x, wx.LIST_AUTOSIZE)
            colWidth1 = self.choices.GetColumnWidth(x)
            # Note that the column header might be wider than any item!
            self.choices.SetColumnWidth(x, wx.LIST_AUTOSIZE_USEHEADER)
            colWidth2 = self.choices.GetColumnWidth(x)
            # Pick the wider of the two, but then make the column just a little bit wider!
            colWidth = max(colWidth1, colWidth2) + 5
            # Set the column width to what we decided was appropriate
            self.choices.SetColumnWidth(x, colWidth)
            # Add that width to our total window width value
            windowWidth += colWidth

        # Add the ListCtrl to the Sizer
        box.Add(self.choices, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        # Select the first item, if the list isn't empty
        if len(configNames) > 0:
            self.choices.SetItemState(0, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)
        # Capture the double-click so it can emulate "Click and OK" function like the SingleChoiceDialog()
        self.choices.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnListDoubleClick)

        # Add a Horizontal Sizer for the buttons
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer on the left to expand, allowing the buttons to be right-justified
        box2.Add((10, 1), 1)
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
        # Fit() often produces a dialog that is too narrow.  If it should be wider, make it wider!
        self.SetSize((max(windowWidth, self.GetSize()[0]), self.GetSize()[1]))
        # Layout the Dialog
        self.Layout()
        # Center on the screen
        self.CentreOnScreen()

    def GetSelection(self):
        # Get the first (only) "Selected" item in the list
        itemNum = self.choices.GetNextItem(-1, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
        # Return that item's data, or None if there is no data.
        if itemNum > -1:
            return self.dataLookup[itemNum]
        else:
            return None

    def OnListDoubleClick(self, event):
        """ Handle double-click in the ListBox, which should emulate list selection and OK """
        # End the modal display with a return code matching the OK button
        self.EndModal(wx.ID_OK)

class AutoWidthListCtrl(wx.ListCtrl, wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin):
    """ Create a ListCtrl class that uses the AutoWidthMixin """
    def __init__(self, parent, ID, pos=wx.DefaultPosition, size=wx.DefaultSize, style=0):
        wx.ListCtrl.__init__(self, parent, ID, pos, size, style)
        wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)
