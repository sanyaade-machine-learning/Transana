# Copyright (C) 2003 - 2012 The Board of Regents of the University of Wisconsin System 
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

"""  This module implements the Clip Properties form.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

import Episode
import Collection
import Clip
import DBInterface
import Dialogs
import EpisodePropertiesForm
import KWManager
import Misc
import PyXML_RTCImportParser
import TransanaConstants
import TransanaGlobal

import wx
import os
import string
import sys
import xml.sax
if TransanaConstants.USESRTC:
    import TranscriptEditor_RTC
else:
    import TranscriptEditor_STC

# Define the maximum number of video files allowed.  (This could change!)
MEDIAFILEMAX = EpisodePropertiesForm.MEDIAFILEMAX

class ClipPropertiesForm(Dialogs.GenForm):
    """Form containing Clip fields."""

    def __init__(self, parent, id, title, clip_object, mergeList=None):
        # If no Merge List is passed in ...
        if mergeList == None:
            # ... use the default Clip Properties Dialog size passed in from the config object.  (This does NOT get saved.)
            size = TransanaGlobal.configData.clipPropertiesSize
            # ... we can enable Clip Change Propagation
            propagateEnabled = True
            HelpContext='Clip Properties'
        # If we DO have a merge list ...
        else:
            # ... increase the size of the dialog to allow for the list of mergeable clips.
            size = (TransanaGlobal.configData.clipPropertiesSize[0], TransanaGlobal.configData.clipPropertiesSize[1] + 130)
            # ... and Clip Change Propagation should be disabled
            propagateEnabled = False
            HelpContext='Clip Merge'

        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style.  Signal that Propogation is included.
        Dialogs.GenForm.__init__(self, parent, id, title, size=size, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 propagateEnabled=propagateEnabled, useSizers = True, HelpContext=HelpContext)
        if mergeList == None:
            # Define the minimum size for this dialog as the initial size
            minWidth = 750
            minHeight = 570
        else:
            # Define the minimum size for this dialog as the initial size
            minWidth = 750
            minHeight = 650
        # Remember the Parent Window
        self.parent = parent
        # Remember the original Clip Object passed in
        self.obj = clip_object
        # Add a placeholder to the clip object for the merge clip number
        self.obj.mergeNumber = 0
        # Remember the merge list, if one is passed in
        self.mergeList = mergeList
        # Initialize the merge item as unselected
        self.mergeItemIndex = -1
        # if Keywords that server as Keyword Examples are removed, we will need to remember them.
        # Then, when OK is pressed, the Keyword Example references in the Database Tree can be removed.
        # We can't remove them immediately in case the whole Clip Properties Edit process is cancelled.
        self.keywordExamplesToDelete = []
        # Initialize a variable to hold merged keyword examples.
        self.mergedKeywordExamples = []

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # If we're merging Clips ...
        if self.mergeList != None:

            # ... display a label for the Merge Clips ...
            lblMergeClip = wx.StaticText(self.panel, -1, _("Clip to Merge"))
            mainSizer.Add(lblMergeClip, 0)
            
            # Add a vertical spacer to the main sizer        
            mainSizer.Add((0, 3))

            # Create a HORIZONTAL sizer for the merge information
            mergeSizer = wx.BoxSizer(wx.HORIZONTAL)
            
            # ... display a ListCtrl for the Merge Clips ...
            self.mergeClips = wx.ListCtrl(self.panel, -1, size=(300, 100), style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
            # Add the element to the sizer
            mergeSizer.Add(self.mergeClips, 1)
            # ... bind the Item Selected event for the List Control ...
            self.mergeClips.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected)

            # ... define the columns for the Merge Clips ...
            self.mergeClips.InsertColumn(0, _('Clip Name'))
            self.mergeClips.InsertColumn(1, _('Collection'))
            self.mergeClips.InsertColumn(2, _('Start Time'))
            self.mergeClips.InsertColumn(3, _('Stop Time'))
            self.mergeClips.SetColumnWidth(0, 244)
            self.mergeClips.SetColumnWidth(1, 244)
            # ... and populate the Merge Clips list from the mergeList data
            for (ClipNum, ClipID, CollectNum, CollectID, ClipStart, ClipStop, transcriptCount) in self.mergeList:
                index = self.mergeClips.InsertStringItem(sys.maxint, ClipID)
                self.mergeClips.SetStringItem(index, 1, CollectID)
                self.mergeClips.SetStringItem(index, 2, Misc.time_in_ms_to_str(ClipStart))
                self.mergeClips.SetStringItem(index, 3, Misc.time_in_ms_to_str(ClipStop))

            # Add the row sizer to the main vertical sizer
            mainSizer.Add(mergeSizer, 0, wx.EXPAND)

            # Add a vertical spacer to the main sizer        
            mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Clip ID
        self.id_edit = self.new_edit_box(_("Clip ID"), v1, self.obj.id, maxLen=100)
        # Add the element to the sizer
        r1Sizer.Add(v1, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r1Sizer.Add((10, 0))

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
        # Series ID
        series_edit = self.new_edit_box(_("Series ID"), v3, self.obj.series_id)
        # Add the element to the sizer
        r2Sizer.Add(v3, 1, wx.EXPAND)
        series_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Episode ID
        episode_edit = self.new_edit_box(_("Episode ID"), v4, self.obj.episode_id)
        # Add the element to the sizer
        r2Sizer.Add(v4, 1, wx.EXPAND)
        episode_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Media Filename(s) [label]
        txt = wx.StaticText(self.panel, -1, _("Media Filename(s)"))
        mainSizer.Add(txt, 0)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 3))

        # Create a HORIZONTAL sizer for the next row
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Media Filename(s)
        # If the media filename path is not empty, we should normalize the path specification
        if self.obj.media_filename == '':
            filePath = self.obj.media_filename
        else:
            filePath = os.path.normpath(self.obj.media_filename)
        # Initialize the list of media filenames with the first one.
        self.filenames = [filePath]
        # For each additional Media file ...
        for vid in self.obj.additional_media_files:
            # ... add it to the filename list
            self.filenames.append(vid['filename'])
        self.fname_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.Size(180, 60), self.filenames)
        r3Sizer.Add(self.fname_lb, 1, wx.EXPAND)
        self.fname_lb.SetDropTarget(ListBoxFileDropTarget(self.fname_lb))

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r3Sizer, 2, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r4Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v5 = wx.BoxSizer(wx.VERTICAL)
        # Clip Start.  Convert to HH:MM:SS.mm
        self.clip_start_edit = self.new_edit_box(_("Clip Start"), v5, Misc.time_in_ms_to_str(self.obj.clip_start))
        # Add the element to the sizer
        r4Sizer.Add(v5, 1, wx.EXPAND)
        # For merging, we need to remember the merged value of Clip Start
        self.clip_start = self.obj.clip_start
        self.clip_start_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r4Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v6 = wx.BoxSizer(wx.VERTICAL)
        # Clip Stop.  Convert to HH:MM:SS.mm
        self.clip_stop_edit = self.new_edit_box(_("Clip Stop"), v6, Misc.time_in_ms_to_str(self.obj.clip_stop))
        # Add the element to the sizer
        r4Sizer.Add(v6, 1, wx.EXPAND)
        # For merging, we need to remember the merged value of Clip Stop
        self.clip_stop = self.obj.clip_stop
        self.clip_stop_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r4Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v7 = wx.BoxSizer(wx.VERTICAL)
        # Clip Length.  Convert to HH:MM:SS.mm
        self.clip_length_edit = self.new_edit_box(_("Clip Length"), v7, Misc.time_in_ms_to_str(self.obj.clip_stop - self.obj.clip_start))
        # Add the element to the sizer
        r4Sizer.Add(v7, 1, wx.EXPAND)
        self.clip_length_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r4Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r5Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v8 = wx.BoxSizer(wx.VERTICAL)
        # Comment
        comment_edit = self.new_edit_box(_("Comment"), v8, self.obj.comment, maxLen=255)
        # Add the element to the sizer
        r5Sizer.Add(v8, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r5Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r6Sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.text_edit = []

        # Notebook for holding Transcripts
        # ... we need to display them within a notebook ...
        self.notebook = wx.Notebook(self.panel, -1)
        # Add the element to the sizer
        r6Sizer.Add(self.notebook, 2, wx.EXPAND)
        
        # Initialize a list of notebook pages
        self.notebookPage = []
        # Initialize a counter for notebood pages
        counter = 0
        # Initialize the TRANSCRIPT start and stop time variables
        self.transcript_clip_start = []
        self.transcript_clip_stop = []
        # Iterate through the clip transcripts
        for tr in self.obj.transcripts:
            # Initialize the transcript start and stop time values for the individual transcripts
            self.transcript_clip_start.append(tr.clip_start)
            self.transcript_clip_stop.append(tr.clip_stop)
            # Create a panel for each notebook page
            self.notebookPage.append(wx.Panel(self.notebook))

            # Add the notebook page to the notebook ...
            self.notebook.AddPage(self.notebookPage[counter], _('Transcript') + " %d" % (counter + 1))
            # ... and use this page as the transcript object's parent
            transcriptParent = self.notebookPage[counter]
            
            # Clip Text

            # Load the Transcript into an RTF Control so the RTF Encoding won't show.
            # We use a list of edit controls to handle multiple transcripts.
            if TransanaConstants.USESRTC:
                self.text_edit.append(TranscriptEditor_RTC.TranscriptEditor(transcriptParent))
            else:
                self.text_edit.append(TranscriptEditor_STC.TranscriptEditor(transcriptParent))

            ##  DKW EXPERIMENT 4/5/2011
            self.text_edit[len(self.text_edit) - 1].load_transcript(self.obj.transcripts[counter])
                                                                    
            self.text_edit[len(self.text_edit) - 1].SetReadOnly(False)
            self.text_edit[len(self.text_edit) - 1].Enable(True)

            # Increment the counter
            counter += 1

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r6Sizer, 4, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r7Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v9 = wx.BoxSizer(wx.VERTICAL)
        # Keyword Group [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword Group"))
        v9.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword Group [list box]

        # Create an empty Keyword Group List for now.  We'll populate it later (for layout reasons)
        self.kw_groups = []
        self.kw_group_lb = wx.ListBox(self.panel, -1, choices = self.kw_groups)
        v9.Add(self.kw_group_lb, 1, wx.EXPAND)
        
        # Add the element to the sizer
        r7Sizer.Add(v9, 1, wx.EXPAND)

        # Create an empty Keyword List for now.  We'll populate it later (for layout reasons)
        self.kw_list = []
        wx.EVT_LISTBOX(self, self.kw_group_lb.GetId(), self.OnGroupSelect)

        # Add a horizontal spacer
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v10 = wx.BoxSizer(wx.VERTICAL)
        # Keyword [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword"))
        v10.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword [list box]
        self.kw_lb = wx.ListBox(self.panel, -1, choices = self.kw_list, style=wx.LB_EXTENDED)
        v10.Add(self.kw_lb, 1, wx.EXPAND)

        wx.EVT_LISTBOX_DCLICK(self, self.kw_lb.GetId(), self.OnAddKW)

        # Add the element to the sizer
        r7Sizer.Add(v10, 1, wx.EXPAND)

        # Add a horizontal spacer
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v11 = wx.BoxSizer(wx.VERTICAL)
        # Keyword transfer buttons
        add_kw = wx.Button(self.panel, wx.ID_FILE2, ">>", wx.DefaultPosition)
        v11.Add(add_kw, 0, wx.EXPAND | wx.TOP, 20)
        wx.EVT_BUTTON(self, wx.ID_FILE2, self.OnAddKW)

        rm_kw = wx.Button(self.panel, wx.ID_FILE3, "<<", wx.DefaultPosition)
        v11.Add(rm_kw, 0, wx.EXPAND | wx.TOP, 10)
        wx.EVT_BUTTON(self, wx.ID_FILE3, self.OnRemoveKW)

        bitmap = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "KWManage.xpm"), wx.BITMAP_TYPE_XPM)
        kwm = wx.BitmapButton(self.panel, wx.ID_FILE4, bitmap)
        v11.Add(kwm, 0, wx.EXPAND | wx.TOP, 10)
        # Add a spacer to increase the height of the Keywords section
        v11.Add((0, 60))
        kwm.SetToolTipString(_("Keyword Management"))
        wx.EVT_BUTTON(self, wx.ID_FILE4, self.OnKWManage)

        # Add the element to the sizer
        r7Sizer.Add(v11, 0)

        # Add a horizontal spacer
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v12 = wx.BoxSizer(wx.VERTICAL)

        # Clip Keywords [label]
        txt = wx.StaticText(self.panel, -1, _("Clip Keywords"))
        v12.Add(txt, 0, wx.BOTTOM, 3)

        # Clip Keywords [list box]
        # Create an empty ListBox.  We'll populate it later for layout reasons.
        self.ekw_lb = wx.ListBox(self.panel, -1, style=wx.LB_EXTENDED)
        v12.Add(self.ekw_lb, 1, wx.EXPAND)

        self.ekw_lb.Bind(wx.EVT_KEY_DOWN, self.OnKeywordKeyDown)

        # Add the element to the sizer
        r7Sizer.Add(v12, 2, wx.EXPAND)

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

        # Because of the way Clips are created (with Drag&Drop / Cut&Paste functions), we have to trap the missing
        # ID error here.  Therefore, we need to override the EVT_BUTTON for the OK Button.
        # Since we don't have an object for the OK Button, we use FindWindowById to find it based on its ID.
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.FindWindowById(wx.ID_OK))
        # We also need to intercept the Cancel button.
        self.Bind(wx.EVT_BUTTON, self.OnCancel, self.FindWindowById(wx.ID_CANCEL))

        self.Bind(wx.EVT_SIZE, self.OnSize)

        # Set the PANEL's main sizer
        self.panel.SetSizer(mainSizer)
        # Tell the PANEL to auto-layout
        self.panel.SetAutoLayout(True)
        # Lay out the Panel
        self.panel.Layout()
        # Lay out the panel on the form
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

        # Populate the Clip Keywords ListBox
        # If the clip object has keywords ... 
        for clipKeyword in self.obj.keyword_list:
            # ... add them to the keyword list
            self.ekw_lb.Append(clipKeyword.keywordPair)

        # If we have a Notebook of text controls ...
        if self.notebookPage:
            # ... interate through the text controls ...
            for textCtrl in self.text_edit:
                # ... and set them to the size of the notebook page.
                textCtrl.SetSize(self.notebookPage[0].GetSizeTuple())

        # Set initial focus to the Clip ID
        self.id_edit.SetFocus()

    def OnOK(self, event):
        """ Intercept the OK button click and process it """
        # Remember the SIZE of the dialog box for next time.  Remember, the size has been altered if
        # we are merging Clips, and we need to remember the NON-MODIFIED size!
        if self.mergeList == None:
            TransanaGlobal.configData.clipPropertiesSize = self.GetSize()
        else:
            TransanaGlobal.configData.clipPropertiesSize = (self.GetSize()[0], self.GetSize()[1] - 130)
        # Because of the way Clips are created (with Drag&Drop / Cut&Paste functions), we have to trap the missing
        # ID error here.  Duplicate ID Error is handled elsewhere.
        if self.id_edit.GetValue().strip() == '':
            # Display the error message
            dlg2 = Dialogs.ErrorDialog(self, _('Clip ID is required.'))
            dlg2.ShowModal()
            dlg2.Destroy()
            # Set the focus on the widget with the error.
            self.id_edit.SetFocus()
        else:
            # Continue on with the form's regular Button event
            event.Skip(True)
        
    def OnCancel(self, event):
        """ Intercept the Cancel button click and process it """
        # Remember the SIZE of the dialog box for next time.  Remember, the size has been altered if
        # we are merging Clips, and we need to remember the NON-MODIFIED size!
        if self.mergeList == None:
            TransanaGlobal.configData.clipPropertiesSize = self.GetSize()
        else:
            TransanaGlobal.configData.clipPropertiesSize = (self.GetSize()[0], self.GetSize()[1] - 130)
        # Now process the event normally, as if it hadn't been intercepted
        event.Skip(True)
                
    def OnSize(self, event):
        """ Size Change event for hte Clip Properties Form """
        # lay out the form
        self.Layout()
        # The Text Controls need explicit re-sizing.  Their layout constraints don't work right.
        if self.notebookPage:
            for textCtrl in self.text_edit:
                textCtrl.SetSize(self.notebookPage[0].GetSizeTuple())

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
            # Get main data values from the form, ID, Comment, and Media Filename
            self.obj.id = gen_input[_('Clip ID')]
            self.obj.comment = gen_input[_('Comment')]
            # If we're merging clips, more data may have changed!
            if self.mergeList != None:
                # We need the post-merge Clip Start Time
                self.obj.clip_start = self.clip_start
                # We need the post-merge Clip Stop Time
                self.obj.clip_stop = self.clip_stop
                # When merging, keywords are not merged into the object automatically, so we need to process the Clip Keyword List
                for clkw in self.ekw_lb.GetStrings():
                    # Separate out the Keyword Group and the Keyword
                    kwlist = string.split(clkw, ':')
                    kwg = string.strip(kwlist[0])
                    kw = ':'.join(kwlist[1:]).strip()
                    # If the keyword is a Keyword Example in the Merged Keyword ...
                    if (kwg, kw) in self.mergedKeywordExamples:
                        # ... and if the keyword isn't already in the Clip object ...
                        if not self.obj.has_keyword(kwg, kw):
                            # ... then add the example keyword to the Clip Object AS AN EXAMPLE.
                            self.obj.add_keyword(kwg, kw, True)

                        # In this situation, the user has merged a clip with a keyword example into a clip that
                        # already has that keyword, but it's not an example.  In this case, we need to tell OTHER
                        # copies of Transana to remove the OLD keyword example.  We don't need to do anything to
                        # the local copy of the database tree, though.  Include the Merge Clip Number, so the right
                        # record gets deleted.
                        if not TransanaConstants.singleUserVersion:
                            msg = "%s >|< %s >|< %s >|< %s >|< %s >|< %s" % ("KeywordExampleNode", "Keywords", kwg, kw, self.mergeList[self.mergeItemIndex][1], self.obj.mergeNumber)
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("DN %s" % msg)
                    # If the keyword isn't a merged example ...
                    else:
                        # ... we need to add it to the list.  (Duplicates don't matter.)
                        self.obj.add_keyword(kwg, kw)

            # Initialize a counter
            counter = 0
            # For each defined Transcripts object ...
            for tr in self.obj.transcripts:
                if TransanaConstants.USESRTC:
                    # ... update the text from the appropriate text control ...
                    tr.text = self.text_edit[counter].GetFormattedSelection('XML')  # GetXMLBuffer()
                else:
                    # ... update the text from the appropriate text control ...
                    tr.text = self.text_edit[counter].GetRTFBuffer()
                # Also update the TRANSCRIPT start and stop times
                tr.clip_start = self.transcript_clip_start[counter]
                tr.clip_stop = self.transcript_clip_stop[counter]
                # ... and increment the counter.
                counter += 1
            # Keyword list is already updated via the OnAddKW() callback
        else:
            self.obj = None

        return self.obj
    

# This simple derrived class let's the user drop files onto a list box
class ListBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, listbox):
        wx.FileDropTarget.__init__(self)
        self.listbox = listbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        # If there are no files in the ListBox ...
        if (self.listbox.GetCount() == 0) or (self.listbox.GetString(0).strip() == ''):
            # ... clear it to prevent there being a blank line at the top of the list.
            self.listbox.Clear()
        # If we have not exceeded the maximum number of files allowed ...
        if self.listbox.GetCount() < MEDIAFILEMAX:
            # ... add the file name to the list box
            self.listbox.Append(files[0])
        # If we have the maximum number of media files already selected ...
        else:
            # ... Display an error message to the user.
             msg = _('A maximum of %d media files is allowed.')
             if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                 msg = unicode(msg, 'utf8')
             dlg = Dialogs.ErrorDialog(None, msg % MEDIAFILEMAX)
             dlg.ShowModal()
             dlg.Destroy()


class AddClipDialog(ClipPropertiesForm):
    """Dialog used when adding a new Clip."""
    # NOTE:  AddClipDialog is not like the AddDialog for other objects.  The Add for other objects
    #        creates the object and then opens the form.  Here, the Clip Object is created based on
    #        data specified in signalling the desire to create a Clip, so the object is created and
    #        mostly populated elsewhere and passed in here.  It's still an Add, though.

    def __init__(self, parent, id, clip_obj):
        ClipPropertiesForm.__init__(self, parent, id, _("Add Clip"), clip_obj)

class EditClipDialog(ClipPropertiesForm):
    """Dialog used when editing Clip properties."""

    def __init__(self, parent, id, clip_object):
        ClipPropertiesForm.__init__(self, parent, id, _("Clip Properties"), clip_object)
