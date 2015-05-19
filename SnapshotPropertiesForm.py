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

"""  This module implements the Snapshot Properties form.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import Clip
import Collection
import DBInterface
import Dialogs
import Episode
import KWManager
import MediaConvert
import Misc
import Snapshot
import TransanaConstants
import TransanaGlobal
# Import Transana's Images
import TransanaImages
import Transcript

import wx
import os
import string
import sys

class SnapshotPropertiesForm(Dialogs.GenForm):
    """Form containing Snapshot fields."""

    def __init__(self, parent, id, title, snapshot_object):
        # ... use the default Clip Properties Dialog size passed in from the config object.  (This does NOT get saved.)
        size = TransanaGlobal.configData.clipPropertiesSize
        # Define the Help Context
        HelpContext='Snapshot Properties'

        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style.  Signal that Propogation is included.
        Dialogs.GenForm.__init__(self, parent, id, title, size=size, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext=HelpContext)
        # Define the minimum size for this dialog as the initial size
        minWidth = 750
        minHeight = 570
        # Remember the Parent Window
        self.parent = parent
        # Remember the original Snapshot Object passed in
        self.obj = snapshot_object

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Snapshot ID
        self.id_edit = self.new_edit_box(_("Snapshot ID"), v1, self.obj.id, maxLen=100)
        # Add the element to the sizer
        r1Sizer.Add(v1, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r1Sizer.Add((10, 0))

        # Update the Snapshot's Collection ID based on the Snapshot's Collection Number
        self.obj._sync_snapshot()

        # Create a VERTICAL sizer for the next element
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Collection ID
        collection_edit = self.new_edit_box(_("Collection ID"), v2, self.obj.GetNodeString(False))
        # Add the element to the sizer
        r1Sizer.Add(v2, 2, wx.EXPAND)
        collection_edit.Enable(False)
        
        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v3 = wx.BoxSizer(wx.VERTICAL)

        # Image Filename
        # If the image filename path is not empty, we should normalize the path specification
        if self.obj.image_filename == '':
            filePath = self.obj.image_filename
        else:
            filePath = os.path.normpath(self.obj.image_filename)

        self.fname_edit = self.new_edit_box(_("Image Filename"), v3, filePath)
        r2Sizer.Add(v3, 5, wx.EXPAND)

        r2Sizer.Add((10, 0))

        # Browse button layout
        btnBrowse = wx.Button(self.panel, wx.ID_FILE1, _("Browse"), wx.DefaultPosition)
        r2Sizer.Add(btnBrowse, 0, wx.EXPAND | wx.TOP, 20)
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)


        # Figure out if the Snapshot Button should be displayed
        if ((isinstance(self.parent.ControlObject.currentObj, Episode.Episode)) or \
            (isinstance(self.parent.ControlObject.currentObj, Clip.Clip))) and \
           (len(self.parent.ControlObject.VideoWindow.mediaPlayers) == 1):

            # Add a horizontal spacer
            r2Sizer.Add((10, 0))

            # Snapshot button layout
            # Create the Snapshot button
            btnSnapshot = wx.BitmapButton(self.panel, -1, TransanaImages.Snapshot.GetBitmap(), size=(48, 24))
            # Set the Help String
            btnSnapshot.SetToolTipString(_("Capture Snapshot for Coding"))
            r2Sizer.Add(btnSnapshot, 0, wx.EXPAND | wx.TOP, 20)
            # Bind the Snapshot button to its event handler
            btnSnapshot.Bind(wx.EVT_BUTTON, self.OnSnapshot)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r4Sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.seriesList = DBInterface.list_of_series()
        seriesRecs = ['']
        serDefault = 0
        for (seriesNum, seriesName) in self.seriesList:
            seriesRecs.append(seriesName)
            if self.obj.series_num == seriesNum:
                serDefault = len(seriesRecs) - 1

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Series ID
        self.series_cb = self.new_choice_box(_("Series ID"), v4, seriesRecs, default = serDefault)
        # Add the element to the sizer
        r4Sizer.Add(v4, 1, wx.EXPAND)
        self.series_cb.Bind(wx.EVT_CHOICE, self.OnSeriesChoice)

        # Add a horizontal spacer to the row sizer        
        r4Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v5 = wx.BoxSizer(wx.VERTICAL)
        # Episode ID
        self.episode_cb = self.new_choice_box(_("Episode ID"), v5, [''])
        # Add the element to the sizer
        r4Sizer.Add(v5, 1, wx.EXPAND)
        self.episode_cb.Bind(wx.EVT_CHOICE, self.OnEpisodeChoice)

        self.episodeList = []
        if self.obj.series_id != '':
            self.PopulateEpisodeChoiceBasedOnSeries(self.obj.series_id)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r4Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r5Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v6 = wx.BoxSizer(wx.VERTICAL)
        # Episode ID
        self.transcript_cb = self.new_choice_box(_("Transcript ID"), v6, [''])
        # Add the element to the sizer
        r5Sizer.Add(v6, 2, wx.EXPAND)
        self.transcript_cb.Bind(wx.EVT_CHOICE, self.OnTranscriptChoice)

        self.transcriptList = []
        if self.obj.episode_id != '':
            self.PopulateTranscriptChoiceBasedOnEpisode(self.obj.series_id, self.obj.episode_id)

        # Add a horizontal spacer to the row sizer        
        r5Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v7 = wx.BoxSizer(wx.VERTICAL)
        # Episode Time Code.  Convert to HH:MM:SS.mm
        self.episode_start_edit = self.new_edit_box(_("Episode Position"), v7, Misc.time_in_ms_to_str(self.obj.episode_start))
        # Add the element to the sizer
        r5Sizer.Add(v7, 1, wx.EXPAND)
        if self.episode_cb.GetStringSelection() == '':
            self.episode_start_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r5Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v8 = wx.BoxSizer(wx.VERTICAL)
        # Episode Duration.  Convert to HH:MM:SS.mm
        self.episode_duration_edit = self.new_edit_box(_("Duration"), v8, Misc.time_in_ms_to_str(self.obj.episode_duration))
        # Add the element to the sizer
        r5Sizer.Add(v8, 1, wx.EXPAND)
        if self.episode_cb.GetStringSelection() == '':
            self.episode_duration_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r5Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r6Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v9 = wx.BoxSizer(wx.VERTICAL)
        # Comment
        comment_edit = self.new_edit_box(_("Comment"), v9, self.obj.comment, maxLen=255)
        # Add the element to the sizer
        r6Sizer.Add(v9, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r6Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r7Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v10 = wx.BoxSizer(wx.VERTICAL)
        # Keyword Group [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword Group"))
        v10.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword Group [list box]

        # Create an empty Keyword Group List for now.  We'll populate it later (for layout reasons)
        self.kw_groups = []
        self.kw_group_lb = wx.ListBox(self.panel, -1, choices = self.kw_groups)
        v10.Add(self.kw_group_lb, 1, wx.EXPAND)
        
        # Add the element to the sizer
        r7Sizer.Add(v10, 1, wx.EXPAND)

        # Create an empty Keyword List for now.  We'll populate it later (for layout reasons)
        self.kw_list = []
        wx.EVT_LISTBOX(self, self.kw_group_lb.GetId(), self.OnGroupSelect)

        # Add a horizontal spacer
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v11 = wx.BoxSizer(wx.VERTICAL)
        # Keyword [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword"))
        v11.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword [list box]
        self.kw_lb = wx.ListBox(self.panel, -1, choices = self.kw_list, style=wx.LB_EXTENDED)
        v11.Add(self.kw_lb, 1, wx.EXPAND)

        wx.EVT_LISTBOX_DCLICK(self, self.kw_lb.GetId(), self.OnAddKW)

        # Add the element to the sizer
        r7Sizer.Add(v11, 1, wx.EXPAND)

        # Add a horizontal spacer
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v12 = wx.BoxSizer(wx.VERTICAL)
        # Keyword transfer buttons
        add_kw = wx.Button(self.panel, wx.ID_FILE2, ">>", wx.DefaultPosition)
        v12.Add(add_kw, 0, wx.EXPAND | wx.TOP, 20)
        wx.EVT_BUTTON(self, wx.ID_FILE2, self.OnAddKW)

        rm_kw = wx.Button(self.panel, wx.ID_FILE3, "<<", wx.DefaultPosition)
        v12.Add(rm_kw, 0, wx.EXPAND | wx.TOP, 10)
        wx.EVT_BUTTON(self, wx.ID_FILE3, self.OnRemoveKW)

        kwm = wx.BitmapButton(self.panel, wx.ID_FILE4, TransanaImages.KWManage.GetBitmap())
        v12.Add(kwm, 0, wx.EXPAND | wx.TOP, 10)
        # Add a spacer to increase the height of the Keywords section
        v12.Add((0, 60))
        kwm.SetToolTipString(_("Keyword Management"))
        wx.EVT_BUTTON(self, wx.ID_FILE4, self.OnKWManage)

        # Add the element to the sizer
        r7Sizer.Add(v12, 0)

        # Add a horizontal spacer
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v13 = wx.BoxSizer(wx.VERTICAL)

        # Whole Snapshot Keywords [label]
        txt = wx.StaticText(self.panel, -1, _("Whole Snapshot Keywords"))
        v13.Add(txt, 0, wx.BOTTOM, 3)

        # Clip Keywords [list box]
        # Create an empty ListBox.  We'll populate it later for layout reasons.
        self.ekw_lb = wx.ListBox(self.panel, -1, style=wx.LB_EXTENDED)
        v13.Add(self.ekw_lb, 1, wx.EXPAND)

        self.ekw_lb.Bind(wx.EVT_KEY_DOWN, self.OnKeywordKeyDown)

        # Add the element to the sizer
        r7Sizer.Add(v13, 2, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r7Sizer, 5, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the buttons
        self.create_buttons(sizer=btnSizer)
        # Add the button sizer to the main sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)

##        # Because of the way Clips are created (with Drag&Drop / Cut&Paste functions), we have to trap the missing
##        # ID error here.  Therefore, we need to override the EVT_BUTTON for the OK Button.
##        # Since we don't have an object for the OK Button, we use FindWindowById to find it based on its ID.
##        self.Bind(wx.EVT_BUTTON, self.OnOK, self.FindWindowById(wx.ID_OK))
##        # We also need to intercept the Cancel button.
##        self.Bind(wx.EVT_BUTTON, self.OnCancel, self.FindWindowById(wx.ID_CANCEL))
##
##        self.Bind(wx.EVT_SIZE, self.OnSize)
##
        # Set the main sizer
        self.panel.SetSizer(mainSizer)
        # Tell the panel to auto-layout
        self.panel.SetAutoLayout(True)
        # Lay out the Panel
        self.panel.Layout()
        # Lay out the form
        self.Layout()
        # Resize the form to fit the contents
        self.Fit()

        # Get the new size of the form
        (width, height) = self.GetSizeTuple()
        # Reset the form's size to be at least the specified minimum width
        self.SetSize(wx.Size(max(minWidth, width), max(minHeight, height)))
        # Define the minimum size for this dialog as the current size
        self.SetSizeHints(max(minWidth, width), max(minHeight, height))
        # Center the form on screen
        self.CenterOnScreen()

        # We need to set some minimum sizes so the sizers will work right
        self.kw_group_lb.SetSizeHints(minW = 50, minH = 20)
        self.kw_lb.SetSizeHints(minW = 50, minH = 20)
        self.ekw_lb.SetSizeHints(minW = 50, minH = 20)

        # We populate the Keyword Groups, Keywords, and Clip Keywords lists AFTER we determine the Form Size.
        # Long Keywords in the list were making the form too big!

        self.kw_groups = DBInterface.list_of_keyword_groups()
        for keywordGroup in self.kw_groups:
            self.kw_group_lb.Append(keywordGroup)

        # Populate the Keywords ListBox
        # Load the parent Collection in order to determine the default Keyword Group
        tempCollection = Collection.Collection(self.obj.collection_num)
        # Select the Collection Default Keyword Group in the Keyword Group list
        if (tempCollection.keyword_group != '') and (self.kw_group_lb.FindString(tempCollection.keyword_group) != wx.NOT_FOUND):
            self.kw_group_lb.SetStringSelection(tempCollection.keyword_group)
        # If no Default Keyword Group is defined, select the first item in the list
        else:
            # but only if there IS a first item.
            if len(self.kw_groups) > 0:
                self.kw_group_lb.SetSelection(0)
        # If there's a selected keyword group ...
        if self.kw_group_lb.GetSelection() != wx.NOT_FOUND:
            # populate the Keywords list
            self.kw_list = \
                DBInterface.list_of_keywords_by_group(self.kw_group_lb.GetStringSelection())
        else:
            # If not, create a blank one
            self.kw_list = []
        for keyword in self.kw_list:
            self.kw_lb.Append(keyword)

        # Populate the Snapshot Keywords ListBox
        # If the snapshot object has keywords ... 
        for snapshotKeyword in self.obj.keyword_list:
            # ... add them to the keyword list
            self.ekw_lb.Append(snapshotKeyword.keywordPair)

        # Set initial focus to the Clip ID
        self.id_edit.SetFocus()

    def refresh_keyword_groups(self):
        """Refresh the keyword groups listbox."""
        self.kw_groups = DBInterface.list_of_keyword_groups()
        self.kw_group_lb.Clear()
        self.kw_group_lb.InsertItems(self.kw_groups, 0)
        if len(self.kw_groups) > 0:
            self.kw_group_lb.SetSelection(0)

    def refresh_keywords(self):
        """Refresh the keywords listbox."""
        sel = self.kw_group_lb.GetStringSelection()
        if sel:
            self.kw_list = \
                DBInterface.list_of_keywords_by_group(sel)
            self.kw_lb.Clear()
            self.kw_lb.InsertItems(self.kw_list, 0)

    def highlight_bad_keyword(self):
        """ Highlight the first bad keyword in the keyword list """
        # Get the Keyword Group name
        sel = self.kw_group_lb.GetStringSelection()
        # If there was a selected Keyword Group ...
        if sel:
            # ... initialize a list of keywords
            kwlist = []
            # Iterate through the current keyword group's keywords ...
            for item in range(self.kw_lb.GetCount()):
                # ... and add them to the list of keywords 
                kwlist.append("%s : %s" % (sel, self.kw_lb.GetString(item)))
            # Now iterate through the list of Episode Keywords
            for item in range(self.ekw_lb.GetCount()):
                # If the keyword is from the current Keyword Group AND the keyword is not in the keyword list ...
                if (self.ekw_lb.GetString(item)[:len(sel)] == sel) and (not self.ekw_lb.GetString(item) in kwlist):
                    # ... select the current item in the Episode Keywords control ...
                    self.ekw_lb.SetSelection(item)
                    # ... and stop looking for bad keywords.  (We just highlight the first!)
                    break

    def OnBrowse(self, event):
        """ Browse to get the Image Filename """
        # Get the current image file name
        imgFile = self.fname_edit.GetValue()
        # Split the file into path and filename
        (imgPath, imgFileName) = os.path.split(imgFile)
        # If the path is blank, use the Video Root
        if imgPath == '':
            imgPath = TransanaGlobal.configData.videoPath
        # Create a dialog box for requesting an image file
        dlg = wx.FileDialog(self, defaultDir=imgPath, defaultFile=imgFileName,
                            wildcard=TransanaConstants.imageFileTypesString, style=wx.OPEN)
        # Set the File Filter to acceptable graphics types
        dlg.SetFilterIndex(1)
        # Get a file selection from the user
        if dlg.ShowModal() == wx.ID_OK:
            # If the user clicks OK, get the name of the selected file
            imgFile = dlg.GetPath()
            # Set the field value to the file name selected
            self.fname_edit.SetValue(imgFile)
            # If the form's ID field is empty ...
            if self.id_edit.GetValue() == '':
                # ... get the base filename
                tempFilename = os.path.basename(imgFile)
                # ... separate the filename root and extension
                (self.obj.id, tempExt) = os.path.splitext(tempFilename)
                # ... and set the ID to match the base file name
                self.id_edit.SetValue(self.obj.id)
        # Destroy the File Dialog
        dlg.Destroy()
        
    def OnSnapshot(self, event):
        """ Handle the Snapshot button press """
        # Create the Media Conversion dialog, including Clip Information so we export only the clip segment
        convertDlg = MediaConvert.MediaConvert(self, self.parent.ControlObject.currentObj.media_filename,
                                               self.parent.ControlObject.GetVideoPosition(), snapshot=True)
        # Show the Media Conversion Dialog
        convertDlg.ShowModal()
        # If the user took a snapshop and the image was successfully created ...
        if convertDlg.snapshotSuccess and os.path.exists(convertDlg.txtDestFileName.GetValue() % 1):
            imgFile = convertDlg.txtDestFileName.GetValue() % 1
            # Set the field value to the file name selected
            self.fname_edit.SetValue(imgFile)
            # If the form's ID field is empty ...
            if self.id_edit.GetValue() == '':
                # ... get the base filename
                tempFilename = os.path.basename(imgFile)
                # ... separate the filename root and extension
                (self.obj.id, tempExt) = os.path.splitext(tempFilename)
                # ... and set the ID to match the base file name
                self.id_edit.SetValue(self.obj.id)
            # If the Snapshot comes from an Episode ...
            if isinstance(self.parent.ControlObject.currentObj, Episode.Episode):
                self.obj.series_num = self.parent.ControlObject.currentObj.series_num
                self.obj.series_id = self.parent.ControlObject.currentObj.series_id
                self.obj.episode_num = self.parent.ControlObject.currentObj.number
                self.obj.episode_id = self.parent.ControlObject.currentObj.id
                self.obj.transcript_num = self.parent.ControlObject.TranscriptNum[self.parent.ControlObject.activeTranscript]
                self.series_cb.SetStringSelection(self.obj.series_id)
                self.PopulateEpisodeChoiceBasedOnSeries(self.obj.series_id)
                self.episode_cb.SetStringSelection(self.obj.episode_id)
                self.PopulateTranscriptChoiceBasedOnEpisode(self.obj.series_id, self.obj.episode_id)
                tmpTranscript = Transcript.Transcript(self.obj.transcript_num)
                self.transcript_cb.SetStringSelection(tmpTranscript.id)
                self.obj.episode_start = self.parent.ControlObject.GetVideoPosition()
                self.episode_start_edit.SetValue(Misc.time_in_ms_to_str(self.obj.episode_start))
                self.episode_start_edit.Enable(True)
                self.obj.episode_duration = 10000
                self.episode_duration_edit.SetValue(Misc.time_in_ms_to_str(self.obj.episode_duration))
                self.episode_duration_edit.Enable(True)
            # If the Snapshot comes from a Clip ...
            elif isinstance(self.parent.ControlObject.currentObj, Clip.Clip):
                self.obj.episode_num = self.parent.ControlObject.currentObj.episode_num
                tmpEpisode = Episode.Episode(self.obj.episode_num)
                self.obj.series_num = tmpEpisode.series_num
                self.obj.series_id = tmpEpisode.series_id
                self.obj.episode_id = tmpEpisode.id
                # We need the Clip's SOURCE TRANSCRIPT for the Active Transcript
                self.obj.transcript_num = self.parent.ControlObject.currentObj.transcripts[self.parent.ControlObject.activeTranscript].source_transcript
                self.series_cb.SetStringSelection(self.obj.series_id)
                self.PopulateEpisodeChoiceBasedOnSeries(self.obj.series_id)
                self.episode_cb.SetStringSelection(self.obj.episode_id)
                self.PopulateTranscriptChoiceBasedOnEpisode(self.obj.series_id, self.obj.episode_id)
                tmpTranscript = Transcript.Transcript(self.obj.transcript_num)
                self.transcript_cb.SetStringSelection(tmpTranscript.id)
                self.obj.episode_start = self.parent.ControlObject.currentObj.clip_start
                self.episode_start_edit.SetValue(Misc.time_in_ms_to_str(self.obj.episode_start))
                self.episode_start_edit.Enable(True)
                self.obj.episode_duration = self.parent.ControlObject.currentObj.clip_stop - self.parent.ControlObject.currentObj.clip_start - 0.1
                self.episode_duration_edit.SetValue(Misc.time_in_ms_to_str(self.obj.episode_duration))
                self.episode_duration_edit.Enable(True)

        # We need to explicitly Close the conversion dialog here to force cleanup of temp files in some circumstances
        convertDlg.Close()
        # Destroy the Media Conversion Dialog
        convertDlg.Destroy()

    def OnSeriesChoice(self, event):
        """ Handle the selection of a Series """
        self.episode_start_edit.Enable(self.episode_cb.GetStringSelection() != '')
        self.episode_duration_edit.Enable(self.episode_cb.GetStringSelection() != '')
        self.transcript_cb.SetItems([''])
        self.transcript_cb.SetStringSelection('')
        self.obj.episode_num = 0
        self.obj.episode_id = ''
        self.obj.transcript_num = 0
        self.episode_start_edit.Enable(False)
        self.episode_duration_edit.Enable(False)
        self.PopulateEpisodeChoiceBasedOnSeries(event.GetString())

    def PopulateEpisodeChoiceBasedOnSeries(self, seriesName):
        # Get a list of all of the Episodes for the selected Series
        self.episodeList = DBInterface.list_of_episodes_for_series(seriesName)
        # Initialize the Episode Records list with a blank entry
        episodeRecs = ['']
        # For each entry in the Episode List ...
        for (episodeNum, episodeID, seriesNum) in self.episodeList:
            # ... add the Episode Name to the Episode Records list
            episodeRecs.append(episodeID)
        self.episode_cb.SetItems(episodeRecs)
        self.episode_cb.SetStringSelection(self.obj.episode_id)

    def OnEpisodeChoice(self, event):
        if event.GetString() != '':
            for (episodeNum, episodeID, seriesNum) in self.episodeList:
                if episodeID == event.GetString():
                    self.obj.episode_num = episodeNum
                    self.obj.episode_id = episodeID
                    self.obj.series_num = seriesNum
                    self.obj.series_id = self.series_cb.GetStringSelection()
                    break
        else:
            self.obj.episode_num = 0
            self.obj.episode_id = ''
        self.obj.transcript_num = 0
        self.episode_start_edit.Enable(False)
        self.episode_duration_edit.Enable(False)
        self.PopulateTranscriptChoiceBasedOnEpisode(self.series_cb.GetStringSelection(), event.GetString())

    def PopulateTranscriptChoiceBasedOnEpisode(self, seriesName, episodeName):
        # Get a list of all of the Transcripts for the selected Episide
        self.transcriptList = DBInterface.list_transcripts(seriesName, episodeName)
        # Initialize the Transcript Records list with a blank entry
        transcriptRecs = ['']
        # Initialize a Transcript Lookup dictionary
        transcriptLookup = {0 : ''}
        # For each entry in the Transcript List ...
        for (transcriptNum, transcriptID, episodeNum) in self.transcriptList:
            # ... add the Transcript Name to the Transcript Records list
            transcriptRecs.append(transcriptID)
            # ... and add the Transcript Rec to the Transcript Lookup dictionary
            transcriptLookup[transcriptNum] = transcriptID
        self.transcript_cb.SetItems(transcriptRecs)
        self.transcript_cb.SetStringSelection(transcriptLookup[self.obj.transcript_num])

    def OnTranscriptChoice(self, event):
        self.episode_start_edit.Enable(event.GetString() != '')
        self.episode_duration_edit.Enable(event.GetString() != '')
        if event.GetString() != '':
            # For each entry in the Transcript List ...
            for (transcriptNum, transcriptID, episodeNum) in self.transcriptList:
                if transcriptID == event.GetString():
                    self.obj.transcript_num = transcriptNum
        else:
            self.obj.transcript_num = 0

    def OnAddKW(self, evt):
        """Invoked when the user activates the Add Keyword (>>) button."""
        # For each selected Keyword ...
        for item in self.kw_lb.GetSelections():
            # ... get the keyword group name ...
            kwg_name = self.kw_group_lb.GetStringSelection()
            # ... get the keyword name ...
            kw_name = self.kw_lb.GetString(item)
            # ... build the kwg : kw combination ...
            ep_kw = "%s : %s" % (kwg_name, kw_name)
            # ... and if it's NOT already in the Episode Keywords list ...
            if self.ekw_lb.FindString(ep_kw) == -1:
                # ... add the keyword to the Episode object ...
                self.obj.add_keyword(kwg_name, kw_name)
                # ... and add it to the Episode Keywords list box
                self.ekw_lb.Append(ep_kw)
        
    def OnRemoveKW(self, evt):
        """Invoked when the user activates the Remove Keyword (<<) button."""
        # Get the selection(s) from the Episode Keywords list box
        kwitems = self.ekw_lb.GetSelections()
        # The items are returned as an immutable tuple.  Convert this to a list.
        kwitems = list(kwitems)
        # Now sort the list.  For reasons that elude me, the list is arbitrarily ordered on the Mac, which causes
        # deletes to be done out of order so the wrong elements get deleted, which is BAD.
        kwitems.sort()
        # We have to go through the list items BACKWARDS so that item numbers don't change on us as we delete items!
        for item in range(len(kwitems), 0, -1):
            sel = kwitems[item - 1]
            # Separate out the Keyword Group and the Keyword
            kwlist = string.split(self.ekw_lb.GetString(sel), ':')
            kwg = string.strip(kwlist[0])
            kw = ':'.join(kwlist[1:]).strip()
            # If the selected keyword is in the current clip object ...
            if self.obj.has_keyword(kwg, kw):
                # Remove the Keyword from the Clip Object (this CAN be overridden by the user!)
                delResult = self.obj.remove_keyword(kwg, kw)
                # Remove the Keyword from the Keywords list
                if (delResult != 0) and (sel >= 0):
                    self.ekw_lb.Delete(sel)
                    # If what we deleted was a Keyword Example, remember the crucial information
                    if delResult == 2:
                        self.keywordExamplesToDelete.append((kwg, kw, self.obj.number))
            # If the selected keyword is NOT in the current object (and therefore has been merged into it) ...
            else:
                # See if it's a keyword example.
                if (kwg, kw) in self.mergedKeywordExamples:
                    # If so, prompt before removing it.
                    prompt = _('Clip "%s" has been defined as a Keyword Example for Keyword "%s : %s".')
                    data = (self.mergeList[self.mergeItemIndex][1], kwg, kw)
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(prompt, 'utf8') % data
                        prompt = prompt + unicode(_('\nAre you sure you want to delete it?'), 'utf8')
                    else:
                        prompt = prompt % data
                        prompt = prompt + _('\nAre you sure you want to delete it?')
                    dlg = Dialogs.QuestionDialog(None, prompt, _('Delete Clip'))
                    # If the user does NOT want to delete the keyword ...
                    if (dlg.LocalShowModal() == wx.ID_NO):
                        dlg.Destroy()
                        # ... then exit this method now before it gets removed.
                        return
                    else:
                        dlg.Destroy()
                        
               # ... merely remove it from the keyword list.  (Merged keywords are not stored in the original Clip object
                # in case the user changes the merge clip selection.)
                self.ekw_lb.Delete(sel)
 
    def OnKWManage(self, evt):
        """Invoked when the user activates the Keyword Management button."""
        # find out if there is a default keyword group
        if self.kw_group_lb.IsEmpty():
            sel = None
        else:
            sel = self.kw_group_lb.GetStringSelection()
        # Create and display the Keyword Management Dialog
        kwm = KWManager.KWManager(self, defaultKWGroup=sel, deleteEnabled=False)
        # Refresh the Keyword Groups list, in case it was changed.
        self.refresh_keyword_groups()
        # Make sure the last Keyword Group selected in the Keyword Management is selected when it gets closed.
        selPos = self.kw_group_lb.FindString(kwm.kw_group.GetStringSelection())
        if selPos == -1:
            selPos = 0
        if not self.kw_group_lb.IsEmpty():
            self.kw_group_lb.SetSelection(selPos)
        # Refresh the Keyword List, in case it was changed.
        self.refresh_keywords()
        # We must refresh the Keyword List in the DBTree to reflect changes made in the
        # Keyword Management.

        #This doesn't work when we create a new clip, as the Parent isn't the DatabaseTreeTab!!
        #self.parent.tree.refresh_kwgroups_node()
        TransanaGlobal.menuWindow.ControlObject.DataWindow.DBTab.tree.refresh_kwgroups_node()

    def OnGroupSelect(self, evt):
        """Invoked when the user selects a keyword group in the listbox."""
        self.refresh_keywords()
        
    def OnKeywordKeyDown(self, event):
        """ Process the Key Down event for the Keyword field """
        # Begin exception handling, as some keys will raise an exception here
        try:
            # Get the key code of the pressed key
            c = event.GetKeyCode()
            # If it's the DELETE key ...
            if c == wx.WXK_DELETE:
                # ... see if there is a selection ...
                if self.ekw_lb.GetSelection() != wx.NOT_FOUND:
                    # ... and if so, delete this keyword!
                    self.OnRemoveKW(event)
        # If an exception occurs, ignore it!
        except:
            pass  # ignore non-ASCII keys

    def OnItemSelected(self, event):
        """ Process the selection of a clip to be merged with the original clip """
        # Identify the selected item
        self.mergeItemIndex = event.GetIndex()
        # Get the Clip Data for the selected item
        mergeClip = Clip.Clip(self.mergeList[self.mergeItemIndex][0])
        # re-initialize the TRANSCRIPT start and stop times
        self.transcript_clip_start = []
        self.transcript_clip_stop = []
        
        # Merge the clips.
        # First, determine whether the merging clip comes BEFORE or after the original
        if mergeClip.clip_start < self.obj.clip_start:
            # The start value comes from the merge clip
            self.clip_start_edit.SetValue(Misc.time_in_ms_to_str(mergeClip.clip_start))
            # Update the merged clip Start Time
            self.clip_start = mergeClip.clip_start
            # The stop value comes from the original clip
            self.clip_stop_edit.SetValue(Misc.time_in_ms_to_str(self.obj.clip_stop))
            # Update the Clip Length
            self.clip_length_edit.SetValue(Misc.time_in_ms_to_str(self.obj.clip_stop - mergeClip.clip_start))
            # Update the merged clip Stop Time
            self.clip_stop = self.obj.clip_stop
            # For each of the original clip's Transcripts ...
            for x in range(len(self.obj.transcripts)):
                # We get the TRANSCRIPT start time from the merge clip
                self.transcript_clip_start.append(mergeClip.transcripts[x].clip_start)
                # We get the TRANSCRIPT end time from the original clip
                self.transcript_clip_stop.append(self.obj.transcripts[x].clip_stop)
                # If we're using the Rich Text Ctrl ...
                if TransanaConstants.USESRTC:
                    # ... clear the transcript ...
                    self.text_edit[x].ClearDoc(skipUnlock = True)
                    # ... turn off read-only ...
                    self.text_edit[x].SetReadOnly(0)
                    # Create the Transana XML to RTC Import Parser.  This is needed so that we can
                    # pull XML transcripts into the existing RTC.  Pass the RTC to be edited.
                    handler = PyXML_RTCImportParser.XMLToRTCHandler(self.text_edit[x])
                    # Parse the merge clip transcript text, adding it to the RTC
                    xml.sax.parseString(mergeClip.transcripts[x].text, handler)
                    # Add a blank line
                    self.text_edit[x].Newline()
                    # ... trap exceptions here ...
                    try:
                        # ... insert a time code at the position of the clip break ...
                        self.text_edit[x].insert_timecode(self.obj.clip_start)
                    # If there were exceptions (duplicating time codes, for example), just skip it.
                    except:
                        pass
                    # Parse the original transcript text, adding it to the RTC
                    xml.sax.parseString(self.obj.transcripts[x].text, handler)
                # If we're using the Styled Text Ctrl
                else:
                    # ... clear the transcript ...
                    self.text_edit[x].ClearDoc()
                    # ... turn off read-only ...
                    self.text_edit[x].SetReadOnly(0)
                    # ... insert the merge clip's text, skipping whitespace at the end ...
                    self.text_edit[x].InsertRTFText(mergeClip.transcripts[x].text.rstrip())
                    # ... add a couple of line breaks ...
                    self.text_edit[x].InsertStyledText('\n\n', len('\n\n'))
                    # ... trap exceptions here ...
                    try:
                        # ... insert a time code at the position of the clip break ...
                        self.text_edit[x].insert_timecode(self.obj.clip_start)
                    # If there were exceptions (duplicating time codes, for example), just skip it.
                    except:
                        pass
                    # ... now add the original clip's text ...
                    self.text_edit[x].InsertRTFText(self.obj.transcripts[x].text)

                # ... signal that time codes will be visible, which they always are in the Clip Properties ...
                self.text_edit[x].codes_vis = 0
                # ... scan transcript for Time Codes ...
                self.text_edit[x].load_timecodes()
                # ... display the time codes
                self.text_edit[x].show_codes()
        # If the merging clip comes AFTER the original ...
        else:
            # The start value comes from the original clip
            self.clip_start_edit.SetValue(Misc.time_in_ms_to_str(self.obj.clip_start))
            # Update the merged clip Start Time
            self.clip_start = self.obj.clip_start
            # The stop value comes from the merge clip
            self.clip_stop_edit.SetValue(Misc.time_in_ms_to_str(mergeClip.clip_stop))
            # Update the Clip Length
            self.clip_length_edit.SetValue(Misc.time_in_ms_to_str(mergeClip.clip_stop - self.obj.clip_start))
            # Update the merged clip Stop Time
            self.clip_stop = mergeClip.clip_stop
            # For each of the original clip's Transcripts ...
            for x in range(len(self.obj.transcripts)):
                # We get the TRANSCRIPT start time from the original clip
                self.transcript_clip_start.append(self.obj.transcripts[x].clip_start)
                # We get the TRANSCRIPT end time from the merge clip
                self.transcript_clip_stop.append(mergeClip.transcripts[x].clip_stop)
                # If we're using the Rich Text Ctrl ...
                if TransanaConstants.USESRTC:
                    # ... clear the transcript ...
                    self.text_edit[x].ClearDoc(skipUnlock = True)
                    # ... turn off read-only ...
                    self.text_edit[x].SetReadOnly(0)
                    # Create the Transana XML to RTC Import Parser.  This is needed so that we can
                    # pull XML transcripts into the existing RTC.  Pass the RTC in.
                    handler = PyXML_RTCImportParser.XMLToRTCHandler(self.text_edit[x])
                    # Parse the original clip transcript text, adding it to the reportText RTC
                    xml.sax.parseString(self.obj.transcripts[x].text, handler)
                    # Add a blank line
                    self.text_edit[x].Newline()
                    # ... trap exceptions here ...
                    try:
                        # ... insert a time code at the position of the clip break ...
                        self.text_edit[x].insert_timecode(mergeClip.clip_start)
                    # If there were exceptions (duplicating time codes, for example), just skip it.
                    except:
                        pass
                    # ... now add the merge clip's text ...
                    # Parse the merge clip transcript text, adding it to the RTC
                    xml.sax.parseString(mergeClip.transcripts[x].text, handler)
                # If we're using the Styled Text Ctrl
                else:
                    # ... clear the transcript ...
                    self.text_edit[x].ClearDoc()
                    # ... turn off read-only ...
                    self.text_edit[x].SetReadOnly(0)
                    # ... insert the original clip's text, skipping whitespace at the end ...
                    self.text_edit[x].InsertRTFText(self.obj.transcripts[x].text.rstrip())
                    # ... add a couple of line breaks ...
                    self.text_edit[x].InsertStyledText('\n\n', len('\n\n'))
                    # ... trap exceptions here ...
                    try:
                        # ... insert a time code at the position of the clip break ...
                        self.text_edit[x].insert_timecode(mergeClip.clip_start)
                    # If there were exceptions (duplicating time codes, for example), just skip it.
                    except:
                        pass
                    # ... now add the merge clip's text ...
                    self.text_edit[x].InsertRTFText(mergeClip.transcripts[x].text)

                # ... signal that time codes will be visible, which they always are in the Clip Properties ...
                self.text_edit[x].codes_vis = 0
                # ... scan transcript for Time Codes ...
                self.text_edit[x].load_timecodes()
                # ... display the time codes
                self.text_edit[x].show_codes()
        # Remember the Merged Clip's Clip Number
        self.obj.mergeNumber = mergeClip.number
        # Create a list object for merging the keywords
        kwList = []
        # Add all the original keywords
        for kws in self.obj.keyword_list:
            kwList.append(kws.keywordPair)
        # Iterate through the merge clip keywords.
        for kws in mergeClip.keyword_list:
            # If they're not already in the list, add them.
            if not kws.keywordPair in kwList:
                kwList.append(kws.keywordPair)
        # Sort the keyword list
        kwList.sort()
        # Clear the keyword list box.
        self.ekw_lb.Clear()
        # Add the merged keywords to the list box.
        for kws in kwList:
            self.ekw_lb.Append(kws)
        # Get the Keyword Examples for the merged clip, if any.
        kwExamples = DBInterface.list_all_keyword_examples_for_a_clip(mergeClip.number)
        # Initialize a variable to hold merged keyword examples.  (This clears the variable if the user changes merge clips.)
        self.mergedKeywordExamples = []
        # Iterate through the keyword examples and add them to the list.
        for kw in kwExamples:
            self.mergedKeywordExamples.append(kw[:2])

    def get_input(self):
        """Custom input routine."""
        # Inherit base input routine
        gen_input = Dialogs.GenForm.get_input(self)
        # If OK ...
        if gen_input:
            # Get main data values from the form, ID, Comment, and Image Filename
            self.obj.id = gen_input[_('Snapshot ID')]
            self.obj.image_filename = gen_input[_('Image Filename')]
            if gen_input[_('Episode Position')] == '':
                self.obj.episode_start = 0.0
            else:
                self.obj.episode_start = Misc.time_in_str_to_ms(gen_input[_('Episode Position')])
            if gen_input[_('Duration')] == '':
                self.obj.episode_duration = 0.0
            else:
                self.obj.episode_duration = Misc.time_in_str_to_ms(gen_input[_('Duration')])
            self.obj.comment = gen_input[_('Comment')]

        else:
            self.obj = None

        return self.obj


class AddSnapshotDialog(SnapshotPropertiesForm):
    """Dialog used when adding a new Snapshot."""
    # NOTE:  AddSnapshotDialog is not like the AddDialog for other objects.  The Add for other objects
    #        creates the object and then opens the form.  Here, the Snapshot Object can be created based on
    #        data specified in signalling the desire to create a Snapshot, so the object is created and
    #        mostly populated elsewhere and passed in here.  It's still an Add, though.

    def __init__(self, parent, id, snapshot_obj):
        SnapshotPropertiesForm.__init__(self, parent, id, _("Add Snapshot"), snapshot_obj)

class EditSnapshotDialog(SnapshotPropertiesForm):
    """Dialog used when editing Snapshot properties."""

    def __init__(self, parent, id, snapshot_obj):
        SnapshotPropertiesForm.__init__(self, parent, id, _("Snapshot Properties"), snapshot_obj)
