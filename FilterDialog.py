# Copyright (C) 2003 - 2015 The Board of Regents of the University of Wisconsin System 
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
# Import Transana's Images
import TransanaImages

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
    def __init__(self, parent, multSelect=False):
        # If multSelect is requested ...
        if multSelect:
            # ... create a ListCtrl in Report View that allows multiple selection
            wx.ListCtrl.__init__(self, parent, -1, style=wx.LC_REPORT)
        # If multSelect is NOT requested ...
        else:
            # ... create a ListCtrl in Report View that only allows single selection
            wx.ListCtrl.__init__(self, parent, -1, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        # Make it a CheckList using the CheckListCtrlMixin
        CheckListCtrlMixin.__init__(self)
        # Bind the Item Activated method
        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.OnItemActivated)

    def OnItemActivated(self, event):
        self.ToggleItem(event.m_itemIndex)

class FilterDialog(wx.Dialog):
    """ This window implements Document, Episode, Quote, Clip, and Keyword filtering for Transana Reports.
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
                        10 = Library Report (reportScope is Library Number)
                        11 = Episode Report (reportScope is Episode Number)
                        12 = Collection Report (reportScope is Collection Number)
                        13 = Notes Report (reportScope is 1 for all notes, 2 for Series, 3 for Episodes, 4 for Transcripts, 5 for Collections,
                                           6 for Clips)
                        14 = Series Clip Data Export (reportScope is the Series Number)
                        15 = Search Save (Search Saves have NO reportScope! or FilterDataType!!)
                        16 = Collection Keyword Map  (reportScope is the CollectionNumber)
                        17 = Document Keyword Map (reportScope is the Document Number)
                        18 = Document Keyword Visualization (reportScope is the Document Number)
                        19 = Document Report (reportScope is the Document Number)
                        20 = Document Analytic Data Export (reportScope is the Document Number)

            *** ADDING A REPORT TYPE?  Remember to add the delete_filter_records() call to the appropriate
                object's db_delete() method!

                ALSO, remember to add the ReportScope conversion to XMLImport for Filter Imports! ***
                         
        Optional parameters are:
          loadDefault         (boolean) -- silently load a profile named "Default", if one exists
          configName          (current Configuration Name)
          reportScope         (required for Configuration Save/Load)
          episodeFilter       (boolean)
          episodeSort         (boolean)
          transcriptFilter    (boolean)
          documentFilter      (boolean)
          collectionFilter    (boolean)
          collectionSort      (boolean)  * NOT FULLY IMPLEMENTED
          quoteFilter         (boolean)
          clipFilter          (boolean)
          clipSort            (boolean)
          snapshotFilter      (boolean)
          keywordGroupFilter  (boolean)
          keywordGroupColor   (boolean)  * NOT FULLY IMPLEMENTED
          keywordFilter       (boolean)
          keywordSort         (boolean)
          keywordColor        (boolean)
          notesFilter         (boolean)
          options             (boolean)
          startTime           (number of milliseconds)
          endTime             (number of milliseconds)
          barHeight           (integer)
          whitespace          (integer)
          hGridLines          (boolean)
          vGridLines          (boolean)
          singleLineDisplay   (boolean)
          showLegend          (boolean)
          colorOutput         (boolean)
          colorAsKeywords     (boolean)
          showSourceInfo      (boolean)
          showQuoteText       (boolean)
          showClipTranscripts (boolean)
          showSnapshotImage   (integer  0 = Full Size, 1 = Medium, 2 = Small, 3 = Don't Show)
          showSnapshotCoding  (boolean)
          showKeywords        (boolean)
          showNestedData      (boolean)
          showHyperlink       (boolean)
          showFile            (boolean)
          showTime            (boolean)
          showComments        (boolean)
          showCollectionNotes (boolean)
          showQuoteNotes      (boolean)
          showClipNotes       (boolean)
          showSnapshotNotes   (boolean) """
    
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

        if self.kwargs.has_key('startTime'):
            if reportType in [17]:
                self.startTimeVal = "%s" % self.kwargs['startTime']
            else:
                self.startTimeVal = Misc.time_in_ms_to_str(self.kwargs['startTime'])
        else:
            self.startTimeVal = Misc.time_in_ms_to_str(0)
        if self.kwargs.has_key('endTime'):
            if self.kwargs['endTime']:
                if reportType in [17]:
                    self.endTimeVal = "%s" % self.kwargs['endTime']
                else:
                    self.endTimeVal = Misc.time_in_ms_to_str(self.kwargs['endTime'])
            else:
                if reportType in [17]:
                    self.endTimeVal = "%s" % parent.CharacterLength
                else:
                    self.endTimeVal = Misc.time_in_ms_to_str(parent.MediaLength)
        else:
            self.endTimeVal = self.startTimeVal

        # Initialize the dialog box.
        # The form needs to be a bit larger on OSX
        if 'wxMac' in wx.PlatformInfo:
            formHeight = 610
        else:
            formHeight = 575
        # Just to be clear, if we're loading the default, we don't actually SEE the dialog, but we still
        # create it and populate all of its fields.  That's the easiest way to load the default config!!
        wx.Dialog.__init__(self, parent, id, title, size = (500, formHeight),
                           style= wx.DEFAULT_DIALOG_STYLE | wx.MAXIMIZE_BOX | wx.RESIZE_BORDER | wx.NO_FULL_REPAINT_ON_RESIZE)
        # Make the background White
        self.SetBackgroundColour(wx.WHITE)

        # Create BoxSizers for the Dialog
        vBox = wx.BoxSizer(wx.VERTICAL)
        hBox = wx.BoxSizer(wx.HORIZONTAL)

        # Add the Tool Bar
        self.toolBar = wx.ToolBar(self, -1, style=wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_TEXT)
        # If there is a "reportScope" parameter, we should add Configuration Open, Save, and Delete buttons
        if self.kwargs.has_key('reportScope'):
            # Create the File Open button
            btnFileOpen = self.toolBar.AddTool(T_FILE_OPEN, TransanaImages.ArtProv_FILEOPEN.GetBitmap(), shortHelpString=_("Load Filter Configuration"))
            # Create the File Save button
            btnFileSave = self.toolBar.AddTool(T_FILE_SAVE, TransanaImages.Save16.GetBitmap(), shortHelpString=_("Save Filter Configuration"))
            # Create the Config Delete button
            btnFileDelete = self.toolBar.AddTool(T_FILE_DELETE, TransanaImages.ArtProv_DELETE.GetBitmap(), shortHelpString=_("Delete Filter Configuration"))
        # Create the Check All button
        btnCheckAll = self.toolBar.AddTool(T_CHECK_ALL, TransanaImages.Check.GetBitmap(), shortHelpString=_('Check All'))
        # Create the Uncheck All button
        btnCheckNone = self.toolBar.AddTool(T_CHECK_NONE, TransanaImages.NoCheck.GetBitmap(), shortHelpString=_('Uncheck All'))
        # create the Help button
        btnHelp = self.toolBar.AddTool(T_HELP_HELP, TransanaImages.ArtProv_HELP.GetBitmap(), shortHelpString=_("Help"))
        # Create the Exit button
        btnExit = self.toolBar.AddTool(T_FILE_EXIT, TransanaImages.Exit.GetBitmap(), shortHelpString=_('Exit'))
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

        # If Document Filtering is requested ...
        if self.kwargs.has_key('documentFilter') and self.kwargs['documentFilter']:
            # Set a flag for the presence of the Documents tab
            self.documentFilter = True
        else:
            # Set a flag for the absence of the Documents tab
            self.documentFilter = False

        # If Episode Filtering is requested ...
        if self.kwargs.has_key('episodeFilter') and self.kwargs['episodeFilter']:
            # Set a flag for the presence of the Episodes tab
            self.episodeFilter = True
        else:
            # Set a flag for the absence of the Episodes tab
            self.episodeFilter = False

        # If Quote Filtering is requested ...
        if self.kwargs.has_key('quoteFilter') and self.kwargs['quoteFilter']:
            # Set a flag for the presence of the Quotes tab
            self.quoteFilter = True
        else:
            # Set a flag for the absence of the Quotes tab
            self.quoteFilter = False

        # If Clip Filtering is requested ...
        if self.kwargs.has_key('clipFilter') and self.kwargs['clipFilter']:
            # Set a flag for the presence of the Clips tab
            self.clipFilter = True
        else:
            # Set a flag for the absence of the Clips tab
            self.clipFilter = False

        # If Snapshot Filtering is requested ...
        if self.kwargs.has_key('snapshotFilter') and self.kwargs['snapshotFilter']:
            # Set a flag for the presence of the Snapshots tab
            self.snapshotFilter = True
        else:
            # Set a flag for the absence of the Snapshots tab
            self.snapshotFilter = False

        # If Keyword Filtering is requested ...
        if self.kwargs.has_key('keywordFilter') and self.kwargs['keywordFilter']:
            # Set a flag for the presence of the Keywords tab
            self.keywordFilter = True
        else:
            # Set a flag for the absence of the Keywords tab
            self.keywordFilter = False

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
                self.showNestedData = wx.CheckBox(self.reportContentsPanel, -1, _("Include Items from Nested Collections"))
                self.showNestedData.SetValue(self.kwargs['showNestedData'])
                pnlVSizer.Add(self.showNestedData, 0, wx.TOP | wx.LEFT, 10)
                text1 = wx.StaticText(self.reportContentsPanel, -1, _("(Unchecking this will cause items from nested collections to\nbe skipped even if checked on the Quote, Clip or Snapshot tabs.)"))
                pnlVSizer.Add(text1, 0, wx.LEFT, 30)
            if self.kwargs.has_key('showHyperlink') and (self.quoteFilter or self.clipFilter or self.snapshotFilter):
                self.showHyperlink = wx.CheckBox(self.reportContentsPanel, -1, _("Enable Hyperlinks"))
                self.showHyperlink.SetValue(self.kwargs['showHyperlink'])
                pnlVSizer.Add(self.showHyperlink, 0, wx.TOP | wx.LEFT, 10)
            else:
                self.showHyperlink = False
            if self.kwargs.has_key('showFile') and (self.episodeFilter or self.documentFilter or self.quoteFilter or self.clipFilter or self.snapshotFilter):
                self.showFile = wx.CheckBox(self.reportContentsPanel, -1, _("Show File Name"))
                self.showFile.SetValue(self.kwargs['showFile'])
                pnlVSizer.Add(self.showFile, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showTime') and \
               (self.quoteFilter or self.episodeFilter or self.clipFilter or self.snapshotFilter):
                if self.reportType == 10:
                    self.showTime = wx.CheckBox(self.reportContentsPanel, -1, _("Show Episode Length"))
                elif self.reportType == 19:
                    self.showTime = wx.CheckBox(self.reportContentsPanel, -1, _("Show Quote Position"))
                else:
                    self.showTime = wx.CheckBox(self.reportContentsPanel, -1, _("Show Item Time / Position"))
                self.showTime.SetValue(self.kwargs['showTime'])
                pnlVSizer.Add(self.showTime, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showDocImportDate') and (self.documentFilter):
                self.showDocImportDate = wx.CheckBox(self.reportContentsPanel, -1, _("Show Document Import Date"))
                self.showDocImportDate.SetValue(self.kwargs['showDocImportDate'])
                pnlVSizer.Add(self.showDocImportDate, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showSourceInfo') and (self.quoteFilter or self.clipFilter or self.snapshotFilter):
                self.showSourceInfo = wx.CheckBox(self.reportContentsPanel, -1, _("Show Source Information"))
                self.showSourceInfo.SetValue(self.kwargs['showSourceInfo'])
                pnlVSizer.Add(self.showSourceInfo, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showQuoteText') and self.quoteFilter:
                self.showQuoteText = wx.CheckBox(self.reportContentsPanel, -1, _("Show Quote Text"))
                self.showQuoteText.SetValue(self.kwargs['showQuoteText'])
                pnlVSizer.Add(self.showQuoteText, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showClipTranscripts') and self.clipFilter:
                self.showClipTranscripts = wx.CheckBox(self.reportContentsPanel, -1, _("Show Clip Transcripts"))
                self.showClipTranscripts.SetValue(self.kwargs['showClipTranscripts'])
                pnlVSizer.Add(self.showClipTranscripts, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showSnapshotImage') and self.snapshotFilter:
                choices = [_('Large'), _('Medium'), _('Small'), _("Don't Show")]
                self.showSnapshotImage = wx.RadioBox(self.reportContentsPanel, -1, _('Show Snapshots'), choices=choices)
                self.showSnapshotImage.SetSelection(self.kwargs['showSnapshotImage'])
                pnlVSizer.Add(self.showSnapshotImage, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showSnapshotCoding') and self.snapshotFilter:
                self.showSnapshotCoding = wx.CheckBox(self.reportContentsPanel, -1, _('Show Snapshot Coding Key'))
                self.showSnapshotCoding.SetValue(self.kwargs['showSnapshotCoding'])
                pnlVSizer.Add(self.showSnapshotCoding, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showKeywords') and self.keywordFilter:
                if self.reportType == 10:
                    prompt = _("Show Keywords")
                else:
                    prompt = _("Show Item Keywords")
                self.showKeywords = wx.CheckBox(self.reportContentsPanel, -1, prompt)
                self.showKeywords.SetValue(self.kwargs['showKeywords'])
                pnlVSizer.Add(self.showKeywords, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showComments'):
                self.showComments = wx.CheckBox(self.reportContentsPanel, -1, _("Show Comments"))
                self.showComments.SetValue(self.kwargs['showComments'])
                pnlVSizer.Add(self.showComments, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showCollectionNotes'):
                self.showCollectionNotes = wx.CheckBox(self.reportContentsPanel, -1, _("Show Collection Notes"))
                self.showCollectionNotes.SetValue(self.kwargs['showCollectionNotes'])
                pnlVSizer.Add(self.showCollectionNotes, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showQuoteNotes') and self.quoteFilter:
                self.showQuoteNotes = wx.CheckBox(self.reportContentsPanel, -1, _("Show Quote Notes"))
                self.showQuoteNotes.SetValue(self.kwargs['showQuoteNotes'])
                pnlVSizer.Add(self.showQuoteNotes, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showClipNotes') and self.clipFilter:
                self.showClipNotes = wx.CheckBox(self.reportContentsPanel, -1, _("Show Clip Notes"))
                self.showClipNotes.SetValue(self.kwargs['showClipNotes'])
                pnlVSizer.Add(self.showClipNotes, 0, wx.TOP | wx.LEFT, 10)
            if self.kwargs.has_key('showSnapshotNotes') and self.snapshotFilter:
                self.showSnapshotNotes = wx.CheckBox(self.reportContentsPanel, -1, _("Show Snapshot Notes"))
                self.showSnapshotNotes.SetValue(self.kwargs['showSnapshotNotes'])
                pnlVSizer.Add(self.showSnapshotNotes, 0, wx.TOP | wx.LEFT, 10)

            # Now declare the panel's vertical sizer as the panel's official sizer
            self.reportContentsPanel.SetSizer(pnlVSizer)

        # Multiple selection has been implemented everywhere!  Signal that it is allowed!!!
        multSelect = True
        
        # If Document Filtering is requested ...
        if self.kwargs.has_key('documentFilter') and self.kwargs['documentFilter']:
            # ... build a Panel for Documents ...
            self.documentsPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Document Panel on the Notebook, creating a Documents tab ...
            self.notebook.AddPage(self.documentsPanel, _("Documents"))
            # ... place a Check List Ctrl on the Documents Panel ...
            self.documentList = CheckListCtrl(self.documentsPanel, multSelect=multSelect)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.documentList, 1, wx.EXPAND)
            # The document List needs two columns, Document ID and Library ID.
            self.documentList.InsertColumn(0, _("Document ID"))
            self.documentList.InsertColumn(1, _("Library ID"))

            # NOTE:  The actual list of Documents needs to be provided by the calling routine using
            #        the SetDocuments method.  This is because only the calling routine knows which
            #        Documents are legal and should be included in the list.  We don't necessarily
            #        want all Documents all the time.  The Document List should be in the form of an
            #        ordered list of tuples made up of (DocumentID, SeriesID, checked) information.

            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.documentsPanel.SetSizer(pnlVSizer)

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
            self.episodeList = CheckListCtrl(self.episodesPanel, multSelect=multSelect)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.episodeList, 1, wx.EXPAND)
            # The episode List needs two columns, Episode ID and Library ID.
            self.episodeList.InsertColumn(0, _("Episode ID"))
            self.episodeList.InsertColumn(1, _("Library ID"))

            # If Episode Sorting capacity has been requested ...
            if self.kwargs.has_key('episodeSort') and self.kwargs['episodeSort']:
                # create a vertical sizer to hold the sort buttons
                pnlBtnSizer = wx.BoxSizer(wx.VERTICAL)
                # create a bitmap button for the Move Up button
                self.btnEpUp = wx.BitmapButton(self.episodesPanel, -1, TransanaImages.ArtProv_UP.GetBitmap())
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
                # create a bitmap button for the Move Down button
                self.btnEpDown = wx.BitmapButton(self.episodesPanel, -1, TransanaImages.ArtProv_DOWN.GetBitmap())
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

        if self.kwargs.has_key('transcriptFilter'):
            self.transcriptFilter = True
        else:
            self.transcriptFilter = False
        #If transcript filter is requested... (Kathleen)
        if self.transcriptFilter and self.kwargs['transcriptFilter']:
            # ... build a Panel for Transcript List ...
            self.transcriptPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Transcripts Panel on the Notebook, creating a Transcripts tab ...
            self.notebook.AddPage(self.transcriptPanel, _("Transcripts"))
            # ... place a Check List Ctrl on the Transcripts Panel ...
            self.transcriptList = CheckListCtrl(self.transcriptPanel, multSelect=multSelect)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.transcriptList, 1, wx.EXPAND)
            # The Transcripts List needs four columns, Series, Episode, Transcript, and Number of Clips.
            self.transcriptList.InsertColumn(0, _('Library'))
            self.transcriptList.InsertColumn(1, _("Episode"))
            self.transcriptList.InsertColumn(2, _("Transcript"))
            self.transcriptList.InsertColumn(3, _("# Clips"))
            
            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND, 0)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.transcriptPanel.SetSizer(pnlVSizer)
            
        if self.kwargs.has_key('collectionFilter'):
            self.collectionFilter = True
        else:
            self.collectionFilter = False
        #If collection filter is requested... (Kathleen)
        if self.collectionFilter and self.kwargs['collectionFilter']:
            # ... build a Panel for Collections ...
            self.collectionPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Collections Panel on the Notebook, creating a Collections tab ...
            self.notebook.AddPage(self.collectionPanel, _("Collections"))
            # ... place a Check List Ctrl on the Collections Panel ...
            self.collectionList = CheckListCtrl(self.collectionPanel, multSelect=multSelect)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.collectionList, 1, wx.EXPAND)
            # The Collections List needs one column, Collection.
            self.collectionList.InsertColumn(0, _("Collection"))
            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND, 0)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.collectionPanel.SetSizer(pnlVSizer)
            
        # If Quote Filtering is requested ...
        if self.quoteFilter:
            # ... build a Panel for Quotes ...
            self.quotesPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Quotes Panel on the Notebook, creating a Quotes tab ...
            self.notebook.AddPage(self.quotesPanel, _("Quotes"))
            # ... place a Check List Ctrl on the Quotes Panel ...
            self.quoteList = CheckListCtrl(self.quotesPanel, multSelect=multSelect)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.quoteList, 1, wx.EXPAND)
            # The Quote List needs two columns, Quote ID and Collection nesting.
            self.quoteList.InsertColumn(0, _("Quote ID"))
            self.quoteList.InsertColumn(1, _("Collection ID(s)"))

            # NOTE:  The actual list of Quotes needs to be provided by the calling routine using
            #        the SetQuotes method.  This is because only the calling routine knows which
            #        Quotes are legal and should be included in the list.  We don't necessarily
            #        want all Quotes all the time.  The Quote List should be in the form of an
            #        ordered list of tuples made up of (collectionNum, quoteID, checked) information.

            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.quotesPanel.SetSizer(pnlVSizer)

        # If Clip Filtering is requested ...
        if self.clipFilter:
            # ... build a Panel for Clips ...
            self.clipsPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Clips Panel on the Notebook, creating a Clips tab ...
            self.notebook.AddPage(self.clipsPanel, _("Clips"))
            # ... place a Check List Ctrl on the Clips Panel ...
            self.clipList = CheckListCtrl(self.clipsPanel, multSelect=multSelect)
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

        # If Snapshot Filtering is requested ...
        if self.snapshotFilter:
            # ... build a Panel for Snapshots ...
            self.snapshotsPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Snapshots Panel on the Notebook, creating a Snapshots tab ...
            self.notebook.AddPage(self.snapshotsPanel, _("Snapshots"))
            # ... place a Check List Ctrl on the Snapshots Panel ...
            self.snapshotList = CheckListCtrl(self.snapshotsPanel, multSelect=multSelect)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.snapshotList, 1, wx.EXPAND)
            # The Snapshot List needs two columns, Snapshot ID and Collection nesting.
            self.snapshotList.InsertColumn(0, _("Snapshot ID"))
            self.snapshotList.InsertColumn(1, _("Collection ID(s)"))

            # If Snapshot Sorting capacity has been requested ...
            if self.kwargs.has_key('snapshotSort') and self.kwargs['snapshotSort']:
                print "Snapshot Sorting has not yet been implemented."
                # self.originalSnapshotData will need to be reordered the same way that self.snapshotList is.
                # Otherwise, GetSnapshots won't work right!

            # NOTE:  The actual list of Snapshots needs to be provided by the calling routine using
            #        the SetSnapshots method.  This is because only the calling routine knows which
            #        Snapshots are legal and should be included in the list.  We don't necessarily
            #        want all Snapshots all the time.  The Snapshot List should be in the form of an
            #        ordered list of tuples made up of (snapshotNum, snapshotID, checked) information.

            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.snapshotsPanel.SetSizer(pnlVSizer)

        if self.kwargs.has_key('keywordGroupFilter'):
            self.keywordGroupFilter = True
        else:
            self.keywordGroupFilter = False
        #If Keyword Group filter is requested...  (Kathleen)
        if self.keywordGroupFilter and self.kwargs['keywordGroupFilter']:
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
                self.keywordGroupList = ColorListCtrl.ColorListCtrl(self.keywordGroupPanel, multSelect=multSelect)
            # If Keyword Group Color specification is disabled ...
            else:
                # ... place a Check List Ctrl on the Keyword Group Panel ...
                self.keywordGroupList = CheckListCtrl(self.keywordGroupPanel, multSelect=multSelect)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.keywordGroupList, 1, wx.EXPAND)
            # The Keyword Groups List needs one column, Keyword Group.
            self.keywordGroupList.InsertColumn(0, _("Keyword Group"))
            # Add the panel's horizontal sizer to the panel's vertical sizer so we can expand in two dimensions
            pnlVSizer.Add(pnlHSizer, 1, wx.EXPAND, 0)
            # Now declare the panel's vertical sizer as the panel's official sizer
            self.keywordGroupPanel.SetSizer(pnlVSizer)
        
        # If Keyword Filtering is requested ...
        if self.keywordFilter:
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
                self.keywordList = ColorListCtrl.ColorListCtrl(self.keywordsPanel, multSelect=multSelect)
            # If Keyword Color specification is disabled ...
            else:
                # ... place a Check List Ctrl on the Keywords Panel ...
                self.keywordList = CheckListCtrl(self.keywordsPanel, multSelect=multSelect)
            # ... and place it on the panel's horizontal sizer.
            pnlHSizer.Add(self.keywordList, 1, wx.EXPAND)
            # The keyword List needs two columns, Keyword Group and Keyword.
            self.keywordList.InsertColumn(0, _("Keyword Group"))
            self.keywordList.InsertColumn(1, _("Keyword"))

            # If Keyword Sorting capacity has been requested ...
            if self.kwargs.has_key('keywordSort') and self.kwargs['keywordSort']:
                # create a vertical sizer to hold the sort buttons
                pnlBtnSizer = wx.BoxSizer(wx.VERTICAL)
                # create a bitmap button for the Move Up button
                self.btnKwUp = wx.BitmapButton(self.keywordsPanel, -1, TransanaImages.ArtProv_UP.GetBitmap())
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
                # create a bitmap button for the Move Down button
                self.btnKwDown = wx.BitmapButton(self.keywordsPanel, -1, TransanaImages.ArtProv_DOWN.GetBitmap())
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

        if self.kwargs.has_key('notesFilter'):
            self.notesFilter = True
        else:
            self.notesFilter = False
        #If Notes filter is requested...
        if self.notesFilter and self.kwargs['notesFilter']:
            # ... build a Panel for Notes ...
            self.notesPanel = wx.Panel(self.notebook, -1)
            # Create vertical and horizontal Sizers for the Panel
            pnlVSizer = wx.BoxSizer(wx.VERTICAL)
            pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
            # ... place the Notes Panel on the Notebook, creating a Notes tab ...
            self.notebook.AddPage(self.notesPanel, _("Notes"))
            # ... place a Check List Ctrl on the Notes Panel ...
            self.notesList = CheckListCtrl(self.notesPanel, multSelect=multSelect)
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
            # Episode Keyword Map Report, the Series Keyword Sequence Map, the Collection Keyword Map, and the Document
            # Keyword Map have Start and End times on the Options tab
            if self.reportType in [1, 5, 16, 17]:
                # Add a label for the Start Time field
                if self.reportType in [17]:
                    startTimeTxt = wx.StaticText(self.optionsPanel, -1, _("Start Position"))
                else:
                    startTimeTxt = wx.StaticText(self.optionsPanel, -1, _("Start Time"))
                pnlVSizer.Add(startTimeTxt, 0, wx.TOP | wx.LEFT, 10)
                # Add the Start Time field
                self.startTime = wx.TextCtrl(self.optionsPanel, -1, self.startTimeVal)
                pnlVSizer.Add(self.startTime, 0, wx.LEFT, 10)
                # Add a label for the End Time field
                if self.reportType in [17]:
                    endTimeTxt = wx.StaticText(self.optionsPanel, -1, _("End Position"))
                else:
                    endTimeTxt = wx.StaticText(self.optionsPanel, -1, _("End Time"))
                pnlVSizer.Add(endTimeTxt, 0, wx.TOP | wx.LEFT, 10)
                # Add the End Time field
                self.endTime = wx.TextCtrl(self.optionsPanel, -1, self.endTimeVal)
                pnlVSizer.Add(self.endTime, 0, wx.LEFT, 10)
                # Add a note about using 0 for end-of-file position.
                if self.reportType in [17]:
                    tRTxt = wx.StaticText(self.optionsPanel, -1, _("NOTE:  Setting the End Position to 0 will set it to the end of the Document."))
                else:
                    tRTxt = wx.StaticText(self.optionsPanel, -1, _("NOTE:  Setting the End Time to 0 will set it to the end of the Media File."))
                pnlVSizer.Add(tRTxt, 0, wx.ALL, 10)

            # Keyword Map Report, Keyword Visualization, the Series Keyword Sequence Map, the Collection Keyword Map,
            # the Document Keyword Map, and the Document Keyword Visualization 
            # have Bar height and Whitespace parameters as well as horizontal and vertical grid lines
            if self.reportType in [1, 2, 5, 6, 7, 16, 17, 18]:
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

            # Keyword Map Report, Keyword Visualization, the Series Keyword Sequence Map, the Collection Keyword Map,
            # and the Document Keyword Visualization have Bar height and Whitespace parameters as well as horizontal
            # and vertical grid lines
            if self.reportType in [1, 2, 5, 6, 7, 16, 17, 18]:
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

            # If we have a Keyword Map, a Series Keyword Sequence Map, a Series Keyword Bar Graph, 
            # a Series Keyword Percentage Graph, or a Collection Keyword Map ...
            if self.reportType in [1, 5, 6, 7, 16, 17] and self.kwargs.has_key('colorOutput'):
                # ... add a check box for Color (vs. GrayScale) output
                self.colorOutput = wx.CheckBox(self.optionsPanel, -1, _("Color Output"))
                self.colorOutput.SetValue(self.kwargs['colorOutput'])
                pnlVSizer.Add(self.colorOutput, 0, wx.TOP | wx.LEFT, 10)

            # If we have a Keyword Map or a Collection Keyword Map, we ask if color represents Clips or Keywords ...
            if self.reportType in [1, 16, 17] and self.kwargs.has_key('colorAsKeywords'):
                # Create a Horizontal Panel
                pnlHSizer = wx.BoxSizer(wx.HORIZONTAL)
                # Create a text label for the radiobox
                txtLabel = wx.StaticText(self.optionsPanel, -1, _('Color represents:'))
                # Add a horizontal spacer for the horizontal Sizer
                pnlHSizer.Add((10, 0))
                # Determine which platform we're on, and select the appropriate top offset for the text, based on platform
                if 'wxMac' in wx.PlatformInfo:
                    topOffset = 6
                else:
                    topOffset = 16
                # Add the text label to the horizontal sizer with the appropriate top offset
                pnlHSizer.Add(txtLabel, 0, wx.TOP, topOffset)
                # ... add a Radio Box for Color as Clips vs. Color as Keywords
                if self.reportType in [17]:
                    choices = [_('Quotes'), _('Keywords')]
                else:
                    choices = [_('Clips'), _('Keywords')]
                self.colorAsKeywords = wx.RadioBox(self.optionsPanel, -1, choices=choices,
                                                     majorDimension=2, style=wx.RA_SPECIFY_COLS)
                # Set the appropriate radio button selection
                self.colorAsKeywords.SetSelection(self.kwargs['colorAsKeywords'])
                # Place the radio button on the horizontal sizer
                pnlHSizer.Add(self.colorAsKeywords, 0, wx.LEFT, 10)
                # Place the horizontal sizer into the main vertical sizer
                pnlVSizer.Add(pnlHSizer, 0)

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
        # NOTE:  The calling routine must provide additional data via the SetEpisodes, SetDocuments, SetQuotes, SetClips, and / or 
        #        SetKeywords methods before this dialog is displayed.  Therefore, we don't complete the Layout or call ShowModal here.
        TransanaGlobal.CenterOnPrimary(self)

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
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Execute the query with the data values
            DBCursor.execute(query, values)
            # Iterate through the report results
            for (configName,) in DBCursor.fetchall():
                # The default config name has been saved in English.  We need to translate it!
                if configName == 'Default':
                    configName = unicode(_('Default'), TransanaGlobal.encoding)
                else:
                    # Decode the data
                    configName = DBInterface.ProcessDBDataForUTF8Encoding(configName)
                # Add the data to the Results list
                resList.append(configName)

            DBCursor.close()
            DBCursor = None

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
                # If the user pressed OK ...
                if result == wx.ID_OK:
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
                # Adjust the query for sqlite if needed
                query = DBInterface.FixQuery(query)
                # Build the data values that match the query
                values = (self.reportType, reportScope, configName.encode(TransanaGlobal.encoding))
                # Execute the query with the appropriate data values
                DBCursor.execute(query, values)
                
                # If there are Episodes in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileEpisodes = self.episodeFilter
                # If there are Documents in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileDocuments = self.documentFilter
                # If there are Quotes in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileQuotes = self.quoteFilter
                # If there are Clips in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileClips = self.clipFilter
                # If there are Keywords in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileKeywords = self.keywordFilter
                # If there are Keyword Groups in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileKeywordGroups = self.keywordGroupFilter
                # If there are Transcripts in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileTranscripts = self.transcriptFilter
                # If there are Collections in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileCollections = self.collectionFilter
                # If there are Notes in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileNotes = self.notesFilter
                # If there are Snapshots in the form, we need to make sure they get reconciled even if they
                # don't occur in the saved Filter Configuration.  Let's make sure this happens.
                needToReconcileSnapshots = self.snapshotFilter
                
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
                        # Note that we have reconciled Episodes
                        needToReconcileEpisodes = False

                    # If the data is for the Clips Tab (filterDataType 2) ...
                    elif self.clipFilter and (filterDataType == 2):
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
                            # Note that we have reconciled Clips
                            needToReconcileClips = False
                        # If the pickled data got truncated in the database, we'll get an Unpickling error here!
                        except cPickle.UnpicklingError, e:
                            # Construct and display an error message here
                            errormsg = _("Transana was unable to load your clip filter data from the database.  Please select your clips again and re-save the filter configuration.")
                            errorDlg = Dialogs.ErrorDialog(self, errormsg)
                            errorDlg.ShowModal()
                            errorDlg.Destroy()
                        
                    # If the data is for the Keywords Tab (filterDataType 3) ...
                    elif self.keywordFilter and (filterDataType == 3):
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
                        # Note that we have reconciled Keywords
                        needToReconcileKeywords = False

                    # If the data is for the Keyword Colors (filterDataType 4) ...
                    elif self.keywordFilter and (filterDataType == 4):
                        # Get the Keyword data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileKeywordColorData = cPickle.loads(filterData.tostring())
                        else:
                            fileKeywordColorData = cPickle.loads(filterData)
                        # Now over-ride existing Keyword Color data with the data from the file
                        for kwPair in fileKeywordColorData.keys():
                            self.keywordColors[kwPair] = fileKeywordColorData[kwPair] % len(TransanaGlobal.keywordMapColourSet)

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
                        # Note that we have reconciled Keyword Groups
                        needToReconcileKeywordGroups = False

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
                        # Note that we have reconciled Transcripts
                        needToReconcileTranscripts = False

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
                        # Note that we have reconciled Collections
                        needToReconcileCollections = False

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
                            if fileNotesData[recNum][2][:len('Libraries')] == 'Libraries':
                                fileNotesData[recNum] = (fileNotesData[recNum][0], fileNotesData[recNum][1], unicode(_('Libraries'), 'utf8') + fileNotesData[recNum][2][len('Libraries'):], fileNotesData[recNum][3])
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
                        # Note that we have reconciled Notes
                        needToReconcileNotes = False

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
                        self.hGridLines.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Vertical Grid Lines (filterDataType 14) ...
                    elif filterDataType == 14:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Vertical Grid Lines value
                        self.vGridLines.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Single Line Display (filterDataType 15) ...
                    elif filterDataType == 15:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Only do this if singleLineDisplay exists!  (Copied configs from ReportType 5 may have this.)
                        if self.kwargs.has_key('singleLineDisplay'):
                            # Set the Single Line Display value
                            self.singleLineDisplay.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Legend value (filterDataType 16) ...
                    elif filterDataType == 16:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Legend value
                        self.showLegend.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Color Output value (filterDataType 17) ...
                    elif filterDataType == 17:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Color Output value
                        self.colorOutput.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Snapshots Tab (filterDataType 18) ...
                    elif self.snapshotFilter and (filterDataType == 18):
                        # Get the current Snapshot data from the Form
                        formSnapshotData = self.GetSnapshots()
                        # Get the Snapshot data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileSnapshotData = cPickle.loads(filterData.tostring())
                        else:
                            fileSnapshotData = cPickle.loads(filterData)
                        # Clear the Snapshot List
                        self.snapshotList.DeleteAllItems()
                        # We need to compare the file data to the form data and reconcile differences,
                        # then feed the results to the Snapshots Tab.
                        self.SetSnapshots(self.ReconcileLists(formSnapshotData, fileSnapshotData))
                        # Note that we have reconciled Snapshots
                        needToReconcileSnapshots = False

                    # If the data is for the Documents Tab (filterDataType 19) ...
                    elif filterDataType == 19:
                        # Get the current Document data from the Form
                        formDocumentData = self.GetDocuments()
                        # Get the Document data from the Database.
                        # (If MySQLDB returns an Array, convert it to a String!)
                        if type(filterData).__name__ == 'array':
                            fileDocumentData = cPickle.loads(filterData.tostring())
                        else:
                            fileDocumentData = cPickle.loads(filterData)
                        # Clear the Document List
                        self.documentList.DeleteAllItems()
                        # We need to compare the file data to the form data and reconcile differences,
                        # then feed the results to the Documents Tab.
                        self.SetDocuments(self.ReconcileLists(formDocumentData, fileDocumentData, listIsOrdered=False))
                        # Note that we have reconciled Documents
                        needToReconcileDocuments = False

                    # If the data is for the Quotes Tab (filterDataType 20) ...
                    elif self.quoteFilter and (filterDataType == 20):
                        # Get the current Quote data from the Form
                        formQuoteData = self.GetQuotes()
                        # Due to the BLOB / LONGBLOB problem, we can have bad data in the database.  Better catch it.
                        # (Probably NOT for Quotes, but the try ... except block does no harm!)
                        try:
                            # Get the Quote data from the Database.
                            # (If MySQLDB returns an Array, convert it to a String!)
                            if type(filterData).__name__ == 'array':
                                fileQuoteData = cPickle.loads(filterData.tostring())
                            else:
                                fileQuoteData = cPickle.loads(filterData)
                            # Clear the Quote List
                            self.quoteList.DeleteAllItems()
                            # This is NOT an ordered list!
                            orderedList = False
                            # We need to compare the file data to the form data and reconcile differences,
                            # then feed the results to the Quote Tab.
                            self.SetQuotes(self.ReconcileLists(formQuoteData, fileQuoteData, listIsOrdered=orderedList))
                            # Note that we have reconciled Quotes
                            needToReconcileQuotes = False
                        # If the pickled data got truncated in the database, we'll get an Unpickling error here!
                        except cPickle.UnpicklingError, e:
                            # Construct and display an error message here
                            errormsg = _("Transana was unable to load your quote filter data from the database.  Please select your quotes again and re-save the filter configuration.")
                            errorDlg = Dialogs.ErrorDialog(self, errormsg)
                            errorDlg.ShowModal()
                            errorDlg.Destroy()
                        
                    # If the data is for the Include Nested Collection Data value (filterDataType 101) ...
                    elif filterDataType == 101:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Nested Collection Data value
                        self.showNestedData.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Media Filename Data value (filterDataType 102) ...
                    elif (self.quoteFilter or self.clipFilter or self.snapshotFilter) and (filterDataType == 102):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Media Filename Data value
                        self.showFile.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Item Time Data value (filterDataType 103) ...
                    elif (self.quoteFilter or self.clipFilter or self.snapshotFilter) and (filterDataType == 103):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Clip Time Data value
                        self.showTime.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Clip Transcripts value (filterDataType 104) ...
                    elif self.clipFilter and (filterDataType == 104):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Clip Transcripts value
                        self.showClipTranscripts.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Keywords value (filterDataType 105) ...
                    elif (self.episodeFilter or self.documentFilter or self.quoteFilter or self.clipFilter or self.snapshotFilter) and (filterDataType == 105):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # If we delete the LAST Keyword from a report, and that report has a
                        # Default Configuration, there is no "showKeyword" control to set!!
                        try:
                            # Set the Show Clip Keywords value
                            self.showKeywords.SetValue((filterData == 'True') or (filterData == '1'))
                        except:
                            pass

                    # If the data is for the Show Comments value (filterDataType 106) ...
                    elif filterDataType == 106:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Comments value
                        self.showComments.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Collection Notes value (filterDataType 107) ...
                    elif filterDataType == 107:
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Collection Notes value
                        self.showCollectionNotes.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Clip Notes value (filterDataType 108) ...
                    elif self.clipFilter and (filterDataType == 108):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Clip Notes value
                        self.showClipNotes.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Source Information value (filterDataType 109) ...
                    elif (self.quoteFilter or self.clipFilter or self.snapshotFilter) and (filterDataType == 109):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Source Information value
                        self.showSourceInfo.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Snapshot Image value (filterDataType 110) ...
                    elif self.snapshotFilter and (filterDataType == 110):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Snapshot Image value
                        self.showSnapshotImage.SetSelection(int(filterData))

                    # If the data is for the Show Snapshot Notes value (filterDataType 111) ...
                    elif self.snapshotFilter and (filterDataType == 111):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Snapshot Notes value
                        self.showSnapshotNotes.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Snapshot Coding value (filterDataType 112) ...
                    elif self.snapshotFilter and (filterDataType == 112):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Snapshot Coding value
                        self.showSnapshotCoding.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Enable Hyperlink value (filterDataType 113) ...
                    elif (filterDataType == 113):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Enable Hyperlink value
                        self.showHyperlink.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Doc Import Date value (filterDataType 114) ...
                    elif (filterDataType == 114):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Doc Import Date value
                        self.showDocImportDate.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Quote Notes value (filterDataType 115) ...
                    elif self.quoteFilter and (filterDataType == 115):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Quote Notes value
                        self.showQuoteNotes.SetValue((filterData == 'True') or (filterData == '1'))

                    # If the data is for the Show Quote Text value (filterDataType 116) ...
                    elif self.quoteFilter and (filterDataType == 116):
                        if type(filterData).__name__ == 'array':
                            filterData = filterData.tostring()
                        # Set the Show Quote Text value
                        self.showQuoteText.SetValue((filterData == 'True') or (filterData == '1'))

                    # If we have an unknown filterDataType ...
                    else:
                        print "Unknown Filter Data:", self.reportType, reportScope, self.configName.encode('utf8'), filterDataType, type(filterData), filterData

                # If there are Episodes on the form but not in the loaded Filter Data ...
                if needToReconcileEpisodes:
                    # Get the current Episode data from the Form
                    formEpisodeData = self.GetEpisodes()
                    # There was no data from the file!
                    fileEpisodeData = []
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

                # If there are Quotes on the form but not in the loaded Filter Data ...
                if needToReconcileQuotes:
                    # Get the current Quote data from the Form
                    formQuoteData = self.GetQuotes()
                    # There was no data from the file!
                    fileQuoteData = []
                    # Clear the Quote List
                    self.quoteList.DeleteAllItems()
                    # Determine if this list is Ordered
                    orderedList = False
                    # We need to compare the file data to the form data and reconcile differences,
                    # then feed the results to the Quote Tab.
                    self.SetQuotes(self.ReconcileLists(formQuoteData, fileQuoteData, listIsOrdered=orderedList))

                # If there are Clips on the form but not in the loaded Filter Data ...
                if needToReconcileClips:
                    # Get the current Clip data from the Form
                    formClipData = self.GetClips()
                    # There was no data from the file!
                    fileClipData = []
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

                # If there are Keywords on the form but not in the loaded Filter Data ...
                if needToReconcileSnapshots:
                    # Get the current Keyword data from the Form
                    formKeywordData = self.GetKeywords()
                    # There was no data from the file!
                    fileKeywordData = []
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

                # If there are Documents on the form but not in the loaded Filter Data ...
                if needToReconcileDocuments:
                    # Get the current Document data from the Form
                    formDocumentData = self.GetDocuments()
                    # There was no data from the file!
                    fileDocumentData = []
                    # Clear the Document List
                    self.documentList.DeleteAllItems()
                    # We need to compare the file data to the form data and reconcile differences,
                    # then feed the results to the Documents Tab.
                    self.SetDocuments(self.ReconcileLists(formDocumentData, fileDocumentData, listIsOrdered=False))

                # If there are Keyword Groups on the form but not in the loaded Filter Data ...
                if needToReconcileKeywordGroups:
                    # Get the current Keyword Group data from the Form
                    formKeywordGroupData = self.GetKeywordGroups()
                    # There was no data from the file!
                    fileKeywordGroupData = []
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

                # If there are Transcripts on the form but not in the loaded Filter Data ...
                if needToReconcileTranscripts:
                    # Get the current Transcript data from the Form
                    formTranscriptData = self.GetTranscripts()
                    # There was no data from the file!
                    fileTranscriptData = []
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

                # If there are Collections on the form but not in the loaded Filter Data ...
                if needToReconcileCollections:
                    # Get the current Collection data from the Form
                    formCollectionData = self.GetCollections()
                    # There was no data from the file!
                    fileCollectionData = []
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

                # If there are Notes on the form but not in the loaded Filter Data ...
                if needToReconcileNotes:
                    # Get the current Notes data from the Form
                    formNotesData = self.GetNotes()
                    # There was no data from the file!
                    fileNotesData = []
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

                # If there are Snapshots on the form but not in the loaded Filter Data ...
                if needToReconcileSnapshots:
                    # Get the current Snapshot data from the Form
                    formSnapshotData = self.GetSnapshots()
                    # There was no data from the file!
                    fileSnapshotData = []
                    # Clear the Snapshot List
                    self.snapshotList.DeleteAllItems()
                    # We need to compare the file data to the form data and reconcile differences,
                    # then feed the results to the Snapshots Tab.
                    self.SetSnapshots(self.ReconcileLists(formSnapshotData, fileSnapshotData))

                # Set the Cursor to the Arrow now that the filter is loaded
                TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

                DBCursor.close()
                DBCursor = None

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
                            
                        # FilterDataTypes
                        #   1 = Episodes,
                        #   2 = Clips,
                        #   3 = Keywords,
                        #   4 = Keyword Colors,
                        #   5 = Keyword Groups (not implemented -- Fix XMLExport and XMLImport when implementing!)
                        #   6 = Transcripts    (not implemented -- Fix XMLExport and XMLImport when implementing!)
                        #   7 = Collections    (not implemented -- Fix XMLExport and XMLImport when implementing!)
                        #   8 = Notes
                        #   9 = Options > Start Time
                        #  10 = Options > End Time
                        #  11 = Options > Keyword Bar Height
                        #  12 = Options > Bar Whitespace
                        #  13 = Options > Horizontal Grid Lines
                        #  14 = Options > Vertical Grid Lines
                        #  15 = Options > Single Line Display
                        #  16 = Options > Show Legend
                        #  17 = Options > Color Output
                        #  18 = Snapshots
                        #  19 = Documents
                        #  20 = Quotes
                        # 101 = Contents > Include Nested Collections
                        # 102 = Contents > Show Media File Names
                        # 103 = Contents > Show Clip Times / Quote Positions
                        # 104 = Contents > Show Clip Transcripts
                        # 105 = Contents > Show Keywords
                        # 106 = Contents > Show Comments
                        # 107 = Contents > Show Collection Notes
                        # 108 = Contents > Show Clip Notes
                        # 109 = Contents > Show Source Information
                        # 110 = Contents > Show Snapshot Image
                        # 111 = Contents > Show Snapshot Notes
                        # 112 = Contents > Show Snapshot Coding Key
                        # 113 = Contents > Enable Hyperlinks
                        # 114 = Contents > Show Document Import Date
                        # 115 = Contents > Show Quote Notes
                        # 116 = Contents > Show Quote Text

                        # If we have a Series Keyword Sequence Map (reportType 5), or
                        # a Series Keyword Bar Graph (reportType 6), or
                        # a Series Keyword Percentage Graph (reportType 7), or
                        # a Library Report (reportType 10), or
                        # a Series Clip Data Export (reportType 14),
                        # update Episode Data (FilterDataType 1)
                        if self.reportType in [5, 6, 7, 10, 14]:
                            # Pickle the Episode Data
                            episodes = cPickle.dumps(self.GetEpisodes())
                            # Save the Filter Data
                            self.SaveFilterData(self.reportType, reportScope, configName, 1, episodes)

                        # If we have an Episode Keyword Map (reportType 1), or
                        # a Keyword Visualization (reportType 2), or
                        # an Episode Clip Data Export (reportType 3), or
                        # a Collection Clip Data Export (reportType 4), or
                        # a Series Keyword Sequence Map (reportType 5), or
                        # a Series Keyword Bar Graph (reportType 6), or
                        # a Series Keyword Percentage Graph (reportType 7),
                        # an Episode Report (reportType 11),
                        # a Collection Report (reportType 12), or
                        # a Series Clip Data Export (reportType 14), or
                        # a Collection Keyword Map (reportType 16),
                        # update Clip Data (FilterDataType 2)
                        if self.clipFilter and self.reportType in [1, 2, 3, 4, 5, 6, 7, 11, 12, 14, 16]:
                            # Pickle the Clip Data
                            clips = cPickle.dumps(self.GetClips())
                            # Save the Filter Data
                            self.SaveFilterData(self.reportType, reportScope, configName, 2, clips)
                            
                        # If we have a Keyword Map (reportType 1), or
                        # a Keyword Visualization (reportType 2), or
                        # an Episode Clip Data Export (reportType 3), or
                        # a Collection Clip Data Export (reportType 4), or
                        # a Series Keyword Sequence Map (reportType 5), or
                        # a Series Keyword Bar Graph (reportType 6), or
                        # a Series Keyword Percentage Graph (reportType 7),
                        # a Library Report (reportType 10)
                        # an Episode Report (reportType 11),
                        # a Collection Report (reportType 12), 
                        # a Series Clip Data Export (reportType 14), 
                        # a Collection Keyword Map (reportType 16),
                        # a Document Keyword Map (reportType 17),
                        # a Document Keyword Visualization (reportType 18),
                        # a Document Report (reportType 19), or
                        # a Document Analytic Data Export (reportType 20),
                        # update Keyword Data (FilterDataType 3)
                        if self.keywordFilter and self.reportType in [1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 14, 16, 17, 18, 19, 20]:
                            # Pickle the Keyword Data
                            keywords = cPickle.dumps(self.GetKeywords())
                            # Save the Filter Data
                            self.SaveFilterData(self.reportType, reportScope, configName, 3, keywords)
                                
                        # If we have a Keyword Map (reportType 1),
                        # a Keyword Visualization (reportType 2),
                        # a Series Keyword Sequence Map (reportType 5),
                        # a Series Keyword Bar Graph (reportType 6),
                        # a Series Keyword Percentage Graph (reportType 7),
                        # a Collection Keyword Map (reportType 16),
                        # a Document Keyword Map (reportType 17), or
                        # a Document Keyword Visualization (reportType 18),
                        # update Keyword Color Data (FilterDataType 4)
                        if self.keywordFilter and self.reportType in [1, 2, 5, 6, 7, 16, 17, 18]:
                            # Pickle the Keyword Colors Data
                            keywordColors = cPickle.dumps(self.GetKeywordColors())
                            # Save the Filter Data
                            self.SaveFilterData(self.reportType, reportScope, configName, 4, keywordColors)

                        # If we have a Notes Report (reportType 13)
                        # update Notes Data (FilterDataType 8)
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
                                if notesList[recNum][2][:len(unicode(_('Libraries'), 'utf8'))] == unicode(_('Libraries'), 'utf8'):
                                    notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Libraries' + notesList[recNum][2][len(unicode(_('Libraries'), 'utf8')):], notesList[recNum][3])
                                elif notesList[recNum][2][:len(unicode(_('Episode'), 'utf8'))] == unicode(_('Episode'), 'utf8'):
                                    notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Episode' + notesList[recNum][2][len(unicode(_('Episode'), 'utf8')):], notesList[recNum][3])
                                elif notesList[recNum][2][:len(unicode(_('Transcript'), 'utf8'))] == unicode(_('Transcript'), 'utf8'):
                                    notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Transcript' + notesList[recNum][2][len(unicode(_('Transcript'), 'utf8')):], notesList[recNum][3])
                                elif notesList[recNum][2][:len(unicode(_('Collection'), 'utf8'))] == unicode(_('Collection'), 'utf8'):
                                    notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Collection' + notesList[recNum][2][len(unicode(_('Collection'), 'utf8')):], notesList[recNum][3])
                                elif notesList[recNum][2][:len(unicode(_('Clip'), 'utf8'))] == unicode(_('Clip'), 'utf8'):
                                    notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Clip' + notesList[recNum][2][len(unicode(_('Clip'), 'utf8')):], notesList[recNum][3])
                                elif notesList[recNum][2][:len(unicode(_('Snapshot'), 'utf8'))] == unicode(_('Snapshot'), 'utf8'):
                                    notesList[recNum] = (notesList[recNum][0], notesList[recNum][1], u'Snapshot' + notesList[recNum][2][len(unicode(_('Snapshot'), 'utf8')):], notesList[recNum][3])
                            # Pickle the Notes Data
                            notes = cPickle.dumps(self.GetNotes())
                            # Save the Filter Data
                            self.SaveFilterData(self.reportType, reportScope, configName, 8, notes)
                            
                        # Filter Data Options information may not already exist, so
                        # need to be checked individually.  These Filter Data Types include:
                            # (1 - 8, 18, 19, and 101 - 114 are also taken)
                            
                        # If we have a Keyword Map (reportType 1), or
                        # a Series Keyword Sequence Map (reportType 5),  or
                        # a Collection Keyword Map (reportType 16),
                        # a Document Keyword Map (reportType 17), 
                        # insert Start Time data (FilterDataType 9)
                        # and End Time data (FilterDataType 10)
                        if self.reportType in [1, 5, 16, 17]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 9, self.startTime.GetValue())
                            self.SaveFilterData(self.reportType, reportScope, configName, 10, self.endTime.GetValue())

                        # If we have a Keyword Map (reportType 1), or
                        # a Keyword Visualization (reportType 2), or
                        # a Series Keyword Sequence Map (reportType 5), or
                        # a Series Keyword Bar Graph (reportType 6), or
                        # a Series Keyword Percentage Graph (reportType 7), or
                        # a Collection Keyword Map (reportType 16),
                        # a Document Keyword Map (reportType 17), or
                        # a Document Keyword Visualization (reportType 18),
                        # insert Keyword Bar Height Data (FilterDataType 11),
                        # Space Between Bars data (FilterDataType 12),
                        if self.reportType in [1, 2, 5, 6, 7, 16, 17, 18]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 11, self.barHeight.GetStringSelection())
                            self.SaveFilterData(self.reportType, reportScope, configName, 12, self.whitespace.GetStringSelection())

                        # If we have a Keyword Map (reportType 1), or
                        # a Keyword Visualization (reportType 2), or
                        # a Series Keyword Sequence Map (reportType 5), or
                        # a Series Keyword Bar Graph (reportType 6), or
                        # a Series Keyword Percentage Graph (reportType 7), or
                        # a Collection Keyword Map (reportType 16),
                        # a Document Keyword Map (reportType 17), or
                        # a Document Keyword Visualization (reportType 18),
                        # Horizontal Grid Lines data (FilterDataType 13), and
                        # Vertical Grid Lines data (FilterDataType 14)
                        if self.reportType in [1, 2, 5, 6, 7, 16, 17, 18]:
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
                        # a Collection Keyword Map (reportType 16), or
                        # a Document Keyword Map (reportType 17),
                        # insert Color output Data (FilterDataType 17)
                        if self.reportType in [1, 5, 6, 7, 16, 17]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 17, self.colorOutput.IsChecked())

                        # If we have a Keyword Map (reportType 1),
                        # a Keyword Visualization (reportType 2),
                        # a Series Keyword Sequence Map (reportType 5), or
                        # a Series Keyword Bar Graph (reportType 6), or
                        # a Series Keyword Percentage Graph (reportType 7),
                        # an Episode Report (reportType 11),
                        # a Collection Report (reportType 12),
                        # a Collection Keyword Map (reportType 16),
                        # update Snapshot Data (FilterDataType 18)
                        if self.snapshotFilter and self.reportType in [1, 2, 5, 6, 7, 11, 12, 16]:
                            # Pickle the Snapshot Data
                            snapshots = cPickle.dumps(self.GetSnapshots())
                            self.SaveFilterData(self.reportType, reportScope, configName, 18, snapshots)
                            
                        # If we have a Library Report (reportType 10)
                        # update Document Data (FilterDataType 19)
                        if self.reportType in [10]:
                            # Pickle the Document Data
                            documents = cPickle.dumps(self.GetDocuments())
                            self.SaveFilterData(self.reportType, reportScope, configName, 19, documents)

                        # If we have a Collection Report (reportType 12),
                        # a Document Keyword Map (reportType 17),
                        # a Document Keyword Visualization (reportType 18),
                        # a Document Report (reportType 19), or
                        # a Document Analytic Data Export (reportType 20),
                        # update Quote Data (FilterDataType 20)
                        if self.reportType in [12, 17, 18, 19, 20]:
                            # Pickle the Document Data
                            quotes = cPickle.dumps(self.GetQuotes())
                            self.SaveFilterData(self.reportType, reportScope, configName, 20, quotes)

                        # If we have a Collection Report (reportType 12), or
                        # a Collection Clip Data Export (reportType 4) AND reportScope != 0
                        #   (Collection Clip Data Export NOT from Root Collection Node), 
                        # insert Include Nested Collections Data (FilterDataType 101)
                        if (self.reportType in [12]) or ((self.reportType == 4) and (reportScope != 0)):
                            self.SaveFilterData(self.reportType, reportScope, configName, 101, self.showNestedData.IsChecked())

                        # If we have a Library Report (reportType 10), or
                        # an Episode Report (reportType 11),
                        # a Collection Report (reportType 12), or
                        # a Document Report (reportType 19),
                        # insert Show Media / Source Document Filename Data (FilterDataType 102),
                        # Show Document Position / Clip Time (FilterDataType 103),
                        # Show Quote / Clip Keywords (FilterDataType 105),
                        if self.reportType in [10, 11, 12, 19]:
                            if self.quoteFilter or self.clipFilter or self.snapshotFilter:
                                self.SaveFilterData(self.reportType, reportScope, configName, 102, self.showFile.IsChecked())
                                self.SaveFilterData(self.reportType, reportScope, configName, 103, self.showTime.IsChecked())
                            if self.keywordFilter:
                                self.SaveFilterData(self.reportType, reportScope, configName, 105, self.showKeywords.IsChecked())

                        # If we have an Episode Report (reportType 11),
                        # a Collection Report (reportType 12), or
                        # a Document Report (reportType 19), 
                        # Show Clip Transcripts (FilterDataType 104),
                        # Show Comments (FilterDataType 106), and
                        # Show Clip Notes (Filter Data Type 108)
                        # Show Source Info (Filter Data Type 109)
                        # Show Snapshot Image (Filter Data Type 110)
                        # Show Snapshot Notes (Filter Data Type 111)
                        # Show Snapshot Coding Key (Filter Data Type 112)
                        # Enable Hyperlink (Filter Data Type 113)
                        # Show Quote Notes (Filter Data Type 115)
                        if self.reportType in [11, 12, 19]:
                            if self.clipFilter:
                                self.SaveFilterData(self.reportType, reportScope, configName, 104, self.showClipTranscripts.IsChecked())
                                self.SaveFilterData(self.reportType, reportScope, configName, 108, self.showClipNotes.IsChecked())
                            self.SaveFilterData(self.reportType, reportScope, configName, 106, self.showComments.IsChecked())
                            if self.quoteFilter or self.clipFilter or self.snapshotFilter:
                                self.SaveFilterData(self.reportType, reportScope, configName, 109, self.showSourceInfo.IsChecked())
                            if self.snapshotFilter:
                                self.SaveFilterData(self.reportType, reportScope, configName, 110, self.showSnapshotImage.GetSelection())
                                self.SaveFilterData(self.reportType, reportScope, configName, 111, self.showSnapshotNotes.IsChecked())
                                self.SaveFilterData(self.reportType, reportScope, configName, 112, self.showSnapshotCoding.IsChecked())
                            if self.showHyperlink:
                                self.SaveFilterData(self.reportType, reportScope, configName, 113, self.showHyperlink.IsChecked())
                            if self.quoteFilter:
                                self.SaveFilterData(self.reportType, reportScope, configName, 115, self.showQuoteNotes.IsChecked())

                        # If we have a Collection Report (reportType 12), 
                        # insert Show Collection Notes (Filter Data Type 107)
                        if self.reportType in [12]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 107, self.showCollectionNotes.IsChecked())

                        # If we have a Library Report (reportType 10), 
                        # insert Show Document Import Date (Filter Data Type 114)
                        if self.reportType in [10]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 114, self.showDocImportDate.IsChecked())

                        # If we have a Collection Report (reportType 12),
                        # a Document Report (reportType 19), 
                        # insert Show Quote Text (Filter Data Type 116)
                        if self.reportType in [12, 19]:
                            self.SaveFilterData(self.reportType, reportScope, configName, 116, self.showQuoteText.IsChecked())

                        # We need a debugging message if the save is requested for an unknown reportType
                        if self.reportType not in [1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20]:
                            tmpDlg = Dialogs.ErrorDialog(self, "FilterDialog.OnFileSave() doesn't yet implement Save for reportType %d." % self.reportType)
                            tmpDlg.ShowModal()
                            tmpDlg.Destroy()

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
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
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
                    if localConfigName == unicode(_('Default'), 'utf8'):
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
                    # Adjust the query for sqlite if needed
                    query = DBInterface.FixQuery(query)
                    # Build the data values that match the query
                    values = (self.reportType, reportScope, localConfigName.encode(TransanaGlobal.encoding))
                    # Execute the query with the appropriate data values
                    DBCursor.execute(query, values)

                    DBCursor.close()
                    DBCursor = None

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
                # If the Episode List Item's checked status does not match the desired status AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.episodeList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.episodeList.GetSelectedItemCount() < 2) or
                    (self.episodeList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
                    # ... then toggle the item so it will match
                    self.episodeList.ToggleItem(x)
        # if we're looking at the Transcripts tab ...
        elif selectedTab == unicode(_("Transcripts"), 'utf8'):
            # ... iterate through the Transcript items in the Transcript List
            for x in range(self.transcriptList.GetItemCount()):
                # If the Transcript List Item's checked status does not match the desired status  AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.transcriptList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.transcriptList.GetSelectedItemCount() < 2) or
                    (self.transcriptList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
                    # ... then toggle the item so it will match
                    self.transcriptList.ToggleItem(x)
        # If we're looking at the Documents tab ...
        if selectedTab == unicode(_("Documents"), 'utf8'):
            # ... iterate through the Documents in the Document List
            for x in range(self.documentList.GetItemCount()):
                # If the Document List Item's checked status does not match the desired status AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.documentList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.documentList.GetSelectedItemCount() < 2) or
                    (self.documentList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
                    # ... then toggle the item so it will match
                    self.documentList.ToggleItem(x)
        # if we're looking at the Collections tab ...
        elif selectedTab == unicode(_("Collections"), 'utf8'):
            # ... iterate through the Collection items in the Collection List
            for x in range(self.collectionList.GetItemCount()):
                # If the Collection List Item's checked status does not match the desired status  AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.collectionList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.collectionList.GetSelectedItemCount() < 2) or
                    (self.collectionList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
                    # ... then toggle the item so it will match
                    self.collectionList.ToggleItem(x)
        # if we're looking at the Quotes tab ...
        elif selectedTab == unicode(_("Quotes"), 'utf8'):
            # ... iterate through the Quotes in the Quote List
            for x in range(self.quoteList.GetItemCount()):
                # If the Quote List Item's checked status does not match the desired status  AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.quoteList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.quoteList.GetSelectedItemCount() < 2) or
                    (self.quoteList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
                    # ... then toggle the item so it will match
                    self.quoteList.ToggleItem(x)
        # if we're looking at the Clips tab ...
        elif selectedTab == unicode(_("Clips"), 'utf8'):
            # ... iterate through the Clips in the Clip List
            for x in range(self.clipList.GetItemCount()):
                # If the Clip List Item's checked status does not match the desired status  AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.clipList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.clipList.GetSelectedItemCount() < 2) or
                    (self.clipList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
                    # ... then toggle the item so it will match
                    self.clipList.ToggleItem(x)
        # if we're looking at the Snapshots tab ...
        elif selectedTab == unicode(_("Snapshots"), 'utf8'):
            # ... iterate through the Snapshots in the Snapshot List
            for x in range(self.snapshotList.GetItemCount()):
                # If the Snapshot List Item's checked status does not match the desired status  AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.snapshotList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.snapshotList.GetSelectedItemCount() < 2) or
                    (self.snapshotList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
                    # ... then toggle the item so it will match
                    self.snapshotList.ToggleItem(x)
        # if we're looking at the Keywords Group tab ...
        elif selectedTab == unicode(_("Keyword Groups"), 'utf8'):
            # ... iterate through the Keyword Groups in the Keyword Group List
            for x in range(self.keywordGroupList.GetItemCount()):
                # If the Keyword Group List Item's checked status does not match the desired status  AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.keywordGroupList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.keywordGroupList.GetSelectedItemCount() < 2) or
                    (self.keywordGroupList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
                    # ... then toggle the item so it will match
                    self.keywordGroupList.ToggleItem(x)
        # if we're looking at the Keywords tab ...
        elif selectedTab == unicode(_("Keywords"), 'utf8'):
            # ... iterate through the Keywords in the Keyword List
            for x in range(self.keywordList.GetItemCount()):
                # If the Keyword List Item's checked status does not match the desired status  AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.keywordList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.keywordList.GetSelectedItemCount() < 2) or
                    (self.keywordList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
                    # ... then toggle the item so it will match
                    self.keywordList.ToggleItem(x)
        # if we're looking at the Notes tab ...
        elif selectedTab == unicode(_("Notes"), 'utf8'):
            # ... iterate through the Notes in the Notes List
            for x in range(self.notesList.GetItemCount()):
                # If the Notes List Item's checked status does not match the desired status  AND
                # (one of fewer items is selected OR
                # the current item is selected) ...
                if (self.notesList.IsChecked(x) != (btnID == T_CHECK_ALL)) and \
                   ((self.notesList.GetSelectedItemCount() < 2) or
                    (self.notesList.GetItemState(x, wx.LIST_STATE_SELECTED) == wx.LIST_STATE_SELECTED)):
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
            # Determine which direction we're going
            if btnID == self.btnEpUp.GetId():
                btnPressed = 'btnUp'
            else:
                btnPressed = 'btnDown'
        # Next, check the Keyword-related Sort Buttons.
        elif self.kwargs.has_key('keywordFilter') and (btnID in [self.btnKwUp.GetId(), self.btnKwDown.GetId()]):
            # If it's one of these, we're working on the Keyword List
            listAffected = self.keywordList
            # Determine which direction we're going
            if btnID == self.btnKwUp.GetId():
                btnPressed = 'btnUp'
            else:
                btnPressed = 'btnDown'
        # It it's not one of those, we have a coding problem!  In this case, we shouldn't do anything.
        else:
            btnPressed = 'Unknown'
            
        # If the Move Item Up button is pressed ...
        if btnPressed == 'btnUp':
            # We need to know the first UN-SELECTED item.  Initialize to the top of the list.
            topItem = 0
            # Determine which item is selected (-1 is None)
            item = listAffected.GetNextItem(-1, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
            # If an item is selected ...
            while item > -1:
                # If the current selected item is at the top of the list ...
                if item - topItem == 0:
                    # ... move the top indicator down to keep this item fixed in place
                    topItem += 1
                # Extract the data from the List
                col1 = listAffected.GetItem(item, 0).GetText()
                col2 = listAffected.GetItem(item, 1).GetText()
                checked = listAffected.IsChecked(item)
                # We store the color code in the ItemData!
                colorSpec = listAffected.GetItemData(item)
                # Check to make sure there's room to move up
                if item > topItem:
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
                # Determine which item is selected (-1 is None)
                item = listAffected.GetNextItem(item, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)

        # If the Move Item Down button is pressed ...
        elif btnPressed == 'btnDown':
            # We need to go from the bottom of the list up, which the ListCtrl doesn't seem to support right.
            # Therefore, we'll collect a list of selected item numbers, reverse that list, and then move the items
            # based on the list.  First, we'll initialize the list.
            sortList = []
            # Set the bottom-of-list indicator to the last item in the list
            topItem = listAffected.GetItemCount() - 1
            # Find the first selected item
            item = listAffected.GetNextItem(-1, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
            # While there are selected items to process ...
            while item > -1:
                # ... add the item to the list
                sortList.append(item)
                # ... and look for the next selected item
                item = listAffected.GetNextItem(item, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
            # When moving down, we need to work in reverse order to prevent sorting problems
            sortList.reverse()
            # If an item is selected ...
            for item in sortList:
                # If the current item is at the bottom of the list ...
                if item - topItem == 0:
                    # ... adjust the bottom position indicator so that it cannot be moved
                    topItem -= 1
                # Extract the data from the  List
                col1 = listAffected.GetItem(item, 0).GetText()
                col2 = listAffected.GetItem(item, 1).GetText()
                checked = listAffected.IsChecked(item)
                # We store the color code in the ItemData!
                colorSpec = listAffected.GetItemData(item)
                # Check to make sure there's room to move down
                if item < topItem:
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
            
    def GetShowSourceInfo(self):
        """ Return the value of showSourceInfo on the Report Contents tab """
        return self.showSourceInfo.GetValue()

    def GetShowQuoteText(self):
        """ Return the value of showQuoteText on the Report Contents tab """
        return self.showQuoteText.GetValue()

    def GetShowClipTranscripts(self):
        """ Return the value of showClipTranscripts on the Report Contents tab """
        return self.showClipTranscripts.GetValue()

    def GetShowSnapshotImage(self):
        """ Return the value of showSnapShotImage on the Report Contents tab """
        return self.showSnapshotImage.GetSelection()

    def GetShowSnapshotCoding(self):
        """ Return the value of showSnapshotCoding on the Report Contents tab """
        return self.showSnapshotCoding.GetValue()

    def GetShowKeywords(self):
        """ Return the value of showKeywords on the Report Contents tab """
        return self.showKeywords.GetValue()

    def GetShowFile(self):
        """ Return the value of showFile on the Report Contents tab """
        return self.showFile.GetValue()

    def GetShowTime(self):
        """ Return the value of showTime on the Report Contents tab """
        return self.showTime.GetValue()

    def GetShowDocImportDate(self):
        """ Return the value of showDocImportDate on the Report Contents tab """
        return self.showDocImportDate.GetValue()

    def GetShowComments(self):
        """ Return the value of showComments on the Report Contents tab """
        return self.showComments.GetValue()

    def GetShowCollectionNotes(self):
        """ Return the value of showCollectionNotes on the Report Contents tab """
        return self.showCollectionNotes.GetValue()

    def GetShowQuoteNotes(self):
        """ Return the value of showQuoteNotes on the Report Contents tab """
        return self.showQuoteNotes.GetValue()

    def GetShowClipNotes(self):
        """ Return the value of showClipNotes on the Report Contents tab """
        return self.showClipNotes.GetValue()

    def GetShowSnapshotNotes(self):
        """ Return the value of showSnapshotNotes on the Report Contents tab """
        return self.showSnapshotNotes.GetValue()

    def GetShowNestedData(self):
        """ Return the value of showNestedData on the Report Contents tab """
        return self.showNestedData.GetValue()

    def GetShowHyperlink(self):
        """ Return the value of the showHyperlink on the Report Contents tab """
        # If we delete the LAST Quote / Clip / Snapshot from a report, and that report has a
        # Default Configuration, there is no "showHyperlink" control to get the value of!!
        try:
            return self.showHyperlink.GetValue()
        except:
            return True

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
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Episode IDs based on the Header
        self.episodeList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Set the Column width for Series Data based on the Header
        self.episodeList.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.episodeList.GetColumnWidth(0) + 10
        w12 = self.episodeList.GetColumnWidth(1) + 10
        # Set the column width for Episode IDs based on the Data
        self.episodeList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Series Data based on the Data
        self.episodeList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # Remember these widths
        w21 = self.episodeList.GetColumnWidth(0) + 10
        w22 = self.episodeList.GetColumnWidth(1) + 10
        # Set the proper column width for Episode IDs 
        self.episodeList.SetColumnWidth(0, max(w11, w21, 100))
        # Set the proper Column width for Series Data
        self.episodeList.SetColumnWidth(1, max(w21, w22, 100))
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
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Episode IDs based on the Header
        self.transcriptList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Set the Column width for Transcript Data based on the Header
        self.transcriptList.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.transcriptList.GetColumnWidth(0) + 10
        w12 = self.transcriptList.GetColumnWidth(1) + 10
        # Set the column width for Episode IDs based on the Data
        self.transcriptList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Transcript Data based on the Data
        self.transcriptList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # Remember these widths
        w21 = self.transcriptList.GetColumnWidth(0) + 10
        w22 = self.transcriptList.GetColumnWidth(1) + 10
        # Set the proper column width for Episode IDs 
        self.transcriptList.SetColumnWidth(0, max(w11, w21, 100))
        # Set the proper Column width for Transcript Data
        self.transcriptList.SetColumnWidth(1, max(w21, w22, 100))
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
    
    def SetDocuments(self, documentList):
        """ Allows the calling routine to provide a list of Documents that should be included on the Documents Tab.
            A sorted list of (DocumentID, seriesID, checked(boolean)) information should be passed in. """
        # Iterate through the Document list that was passed in
        for (documentID, seriesID, checked) in documentList:
            # Create a new Item in the Document List at the end of the list.  Add the Document ID data.
            index = self.documentList.InsertStringItem(sys.maxint, documentID)
            # Add the Library data
            self.documentList.SetStringItem(index, 1, seriesID)
            # If the item should be checked, check it!
            if checked:
                self.documentList.CheckItem(index)
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Document IDs based on the Header
        self.documentList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Set the Column width for Library Data based on the Header
        self.documentList.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.documentList.GetColumnWidth(0) + 10
        w12 = self.documentList.GetColumnWidth(1) + 10
        # Set the column width for Document IDs based on the Data
        self.documentList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Library Data based on the Data
        self.documentList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # Remember these widths
        w21 = self.documentList.GetColumnWidth(0) + 10
        w22 = self.documentList.GetColumnWidth(1) + 10
        # Set the proper column width for Document IDs 
        self.documentList.SetColumnWidth(0, max(w11, w21, 100))
        # Set the proper Column width for Library Data
        self.documentList.SetColumnWidth(1, max(w21, w22, 100))
        # lay out the Document Panel
        self.documentsPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()

    def GetDocuments(self):
        """ Allows the calling routine to retrieve the Document data from the Filter Dialog.  A sorted list
            of (documentID, seriesID, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty Document list
        documentList = []
        # Iterate throught the DocumentListCtrl items
        for x in range(self.documentList.GetItemCount()):
            # Append the list item's data to the document list
            documentList.append((self.documentList.GetItem(x, 0).GetText(), self.documentList.GetItem(x, 1).GetText(), self.documentList.IsChecked(x)))
        # Return the docuemnt list
        return documentList

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
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Collection Name based on the Header
        self.collectionList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.collectionList.GetColumnWidth(0) + 10
        # Set the column width for Collection Name based on the Data
        self.collectionList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Remember these widths
        w21 = self.collectionList.GetColumnWidth(0) + 10
        # Set the proper column width for Collection Name
        self.collectionList.SetColumnWidth(0, max(w11, w21, 100))
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
    
    def SetQuotes(self, quoteList):
        """ Allows the calling routine to provide a list of Quotes that should be included on the Quotes Tab.
            A sorted list of (quoteName, collectionNumber, checked(boolean)) information should be passed in. """
        # Save the original data.  We'll need it to retrieve the Collection Number data
        self.originalQuoteData = quoteList
        # Iterate through the Quote list that was passed in
        for (quoteID, collectNum, checked) in quoteList:
            # Create a new Item in the Quote List at the end of the list.  Add the Quote ID data.
            index = self.quoteList.InsertStringItem(sys.maxint, quoteID)
            # Add the Collection data
            tempColl = Collection.Collection(collectNum)
            # Add the Collection's Node String
            self.quoteList.SetStringItem(index, 1, tempColl.GetNodeString())
            # If the item should be checked, check it!
            if checked:
                self.quoteList.CheckItem(index)
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Quote IDs based on the Header
        self.quoteList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Set the Column width for Collection Data based on the Header
        self.quoteList.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.quoteList.GetColumnWidth(0) + 10
        w12 = self.quoteList.GetColumnWidth(1) + 10
        # Set the column width for Quote IDs based on the Data
        self.quoteList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Collection Data based on the Data
        self.quoteList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # Remember these widths
        w21 = self.quoteList.GetColumnWidth(0) + 10
        w22 = self.quoteList.GetColumnWidth(1) + 10
        # Set the proper column width for Quote IDs 
        self.quoteList.SetColumnWidth(0, max(w11, w21, 100))
        # Set the proper Column width for Collection Data
        self.quoteList.SetColumnWidth(1, max(w21, w22, 100))
        # lay out the Quotes Panel
        self.quotesPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()

    def GetQuotes(self):
        """ Allows the calling routine to retrieve the Quote data from the Filter Dialog.  A sorted list
            of (quoteID, collectNum, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty Quote list
        quoteList = []
        # Iterate throught the QuoteListCtrl items
        for x in range(self.quoteList.GetItemCount()):
            # Append the list item's data to the Quote list
            quoteList.append((self.quoteList.GetItem(x, 0).GetText(), self.originalQuoteData[x][1], self.quoteList.IsChecked(x)))
        # Return the quote list
        return quoteList

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
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Clip IDs based on the Header
        self.clipList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Set the Column width for Collection Data based on the Header
        self.clipList.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.clipList.GetColumnWidth(0) + 10
        w12 = self.clipList.GetColumnWidth(1) + 10
        # Set the column width for Clip IDs based on the Data
        self.clipList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Collection Data based on the Data
        self.clipList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # Remember these widths
        w21 = self.clipList.GetColumnWidth(0) + 10
        w22 = self.clipList.GetColumnWidth(1) + 10
        # Set the proper column width for Clip IDs 
        self.clipList.SetColumnWidth(0, max(w11, w21, 100))
        # Set the proper Column width for Collection Data
        self.clipList.SetColumnWidth(1, max(w21, w22, 100))
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

    def SetSnapshots(self, snapshotList):
        """ Allows the calling routine to provide a list of Snapshots that should be included on the Snapshots Tab.
            A sorted list of (snapshotName, collectionNumber, checked(boolean)) information should be passed in. """
        # Save the original data.  We'll need it to retrieve the Collection Number data
        self.originalSnapshotData = snapshotList
        # Iterate through the Snapshot list that was passed in
        for (snapshotID, collectNum, checked) in snapshotList:
            # Create a new Item in the Snapshot List at the end of the list.  Add the Snapshot ID data.
            index = self.snapshotList.InsertStringItem(sys.maxint, snapshotID)
            # Add the Collection data
            tempColl = Collection.Collection(collectNum)
            # Add the Collection's Node String
            self.snapshotList.SetStringItem(index, 1, tempColl.GetNodeString())
            # If the item should be checked, check it!
            if checked:
                self.snapshotList.CheckItem(index)
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Snapshot IDs based on the Header
        self.snapshotList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Set the Column width for Collection Data based on the Header
        self.snapshotList.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.snapshotList.GetColumnWidth(0) + 10
        w12 = self.snapshotList.GetColumnWidth(1) + 10
        # Set the column width for Snapshot IDs based on the Data
        self.snapshotList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Collection Data based on the Data
        self.snapshotList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # Remember these widths
        w21 = self.snapshotList.GetColumnWidth(0) + 10
        w22 = self.snapshotList.GetColumnWidth(1) + 10
        # Set the proper column width for Snapshot IDs 
        self.snapshotList.SetColumnWidth(0, max(w11, w21, 100))
        # Set the proper Column width for Collection Data
        self.snapshotList.SetColumnWidth(1, max(w21, w22, 100))
        # lay out the Snapshots Panel
        self.snapshotsPanel.Layout()
        # Lay out the Dialog
        self.Layout()
        # Center the Dialog on the screen
        self.CenterOnScreen()

    def GetSnapshots(self):
        """ Allows the calling routine to retrieve the Snapshot data from the Filter Dialog.  A sorted list
            of (snapshotID, collectNum, checked(boolean)) information is returned.  (Unchecked items ARE included,
            as we don't want to lose their information for later processing.) """
        # Create an empty Snapshot list
        snapshotList = []
        # Iterate throught the SnapshotListCtrl items
        for x in range(self.snapshotList.GetItemCount()):
            # Append the list item's data to the Snapshot list
            snapshotList.append((self.snapshotList.GetItem(x, 0).GetText(), self.originalSnapshotData[x][1], self.snapshotList.IsChecked(x)))
        # Return the Snapshot list
        return snapshotList

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
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Keyword Groups based on the Header
        self.keywordGroupList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.keywordGroupList.GetColumnWidth(0) + 10
        # Set the column width for Keyword Groups based on the Data
        self.keywordGroupList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Remember these widths
        w21 = self.keywordGroupList.GetColumnWidth(0) + 10
        # Set the proper column width for Keyword Groups 
        self.keywordGroupList.SetColumnWidth(0, max(w11, w21, 100))
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
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Keyword Groups based on the Header
        self.keywordList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Set the Column width for Keywords based on the Header
        self.keywordList.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.keywordList.GetColumnWidth(0) + 10
        w12 = self.keywordList.GetColumnWidth(1) + 10
        # Set the column width for Keyword Groups based on the Data
        self.keywordList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Keywords based on the Data
        self.keywordList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        self.keywordsPanel.Layout()
        # Remember these widths
        w21 = self.keywordList.GetColumnWidth(0) + 10
        w22 = self.keywordList.GetColumnWidth(1) + 10
        # Set the proper column width for Keyword Groups 
        self.keywordList.SetColumnWidth(0, max(w11, w21, 100))
        # Set the proper Column width for Keywords
        self.keywordList.SetColumnWidth(1, max(w21, w22, 100))
        # lay out the panel
        self.keywordsPanel.Layout()
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
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # Set the column width for Note Name based on the Header
        self.notesList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
        # Set the Column width for Note Parent based on the Header
        self.notesList.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        # Remember these widths
        w11 = self.notesList.GetColumnWidth(0) + 10
        w12 = self.notesList.GetColumnWidth(1) + 10
        # Set the column width for Note Name based on the Data
        self.notesList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
        # Set the Column width for Note Parent based on the Data
        self.notesList.SetColumnWidth(1, wx.LIST_AUTOSIZE)
        # Remember these widths
        w21 = self.notesList.GetColumnWidth(0) + 10
        w22 = self.notesList.GetColumnWidth(1) + 10
        # Set the proper column width for Note Name 
        self.notesList.SetColumnWidth(0, max(w11, w21, 100))
        # Set the proper Column width for Note Parent
        self.notesList.SetColumnWidth(1, max(w21, w22, 100))
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

    def GetColorAsKeywords(self):
        """ Return the value of colorAsKeywords on the Options tab """
        return self.colorAsKeywords.GetSelection()

    def OnOK(self, event):
        """ implement the Filter Dialog's OK button """
        # If we're using a Default configuration ...
        if self.configName == unicode(_('Default'), 'utf8'):
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
        # Series Keyword Bar Graph, Series Keyword Percentage Graph, Collection Report.)
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
        # Create the Configuration Copy dialog box.
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
                # Adjust the query for sqlite if needed
                query = DBInterface.FixQuery(query)
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
                        # Adjust the query for sqlite if needed
                        query = DBInterface.FixQuery(query)
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

                DBCursor.close()
                DBCursor = None

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
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
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

        DBCursor.close()
        DBCursor = None

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
        # For Series-based configs ...
        # ... label the first column "Libraries"
        self.choices.InsertColumn(0, _('Libraries'))
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
            self.choices.SetColumnWidth(x, wx.LIST_AUTOSIZE_USEHEADER)
            colWidth1 = self.choices.GetColumnWidth(x)
            # Note that the column header might be wider than any item!
            self.choices.SetColumnWidth(x, wx.LIST_AUTOSIZE_USEHEADER)
            colWidth2 = self.choices.GetColumnWidth(x)
            # Pick the wider of the two, but then make the column just a little bit wider!
            colWidth = max(colWidth1, colWidth2) + 10
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
