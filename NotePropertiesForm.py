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

"""  This module implements the Note Properties form.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case <nacase@wisc.edu>'

import TransanaConstants
import DBInterface
import Dialogs
import KWManager
import Misc

import Series
import Episode
import Transcript
import Collection
import Clip
import Snapshot
import Note

import wx
import string

class NotePropertiesForm(Dialogs.GenForm):
    """Form containing Note fields."""

    def __init__(self, parent, id, title, note_object):
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(400, 260), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Notes')  # 'Notes' is the Help Context for Notes.  There is no entry for Note Properties at this time

        self.obj = note_object
        seriesID = ''
        episodeID = ''
        transcriptID = ''
        collectionID = ''
        clipID = ''
        snapshotID = ''
        if (self.obj.series_num != 0) and (self.obj.series_num != None):
            tempSeries = Series.Series(self.obj.series_num)
            seriesID = tempSeries.id
        elif (self.obj.episode_num != 0) and (self.obj.episode_num != None):
            tempEpisode = Episode.Episode(self.obj.episode_num)
            episodeID = tempEpisode.id
            tempSeries = Series.Series(tempEpisode.series_num)
            seriesID = tempSeries.id
        elif (self.obj.transcript_num != 0) and (self.obj.transcript_num != None):
            # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
            tempTranscript = Transcript.Transcript(self.obj.transcript_num, skipText=True)
            transcriptID = tempTranscript.id
            tempEpisode = Episode.Episode(tempTranscript.episode_num)
            episodeID = tempEpisode.id
            tempSeries = Series.Series(tempEpisode.series_num)
            seriesID = tempSeries.id
        elif (self.obj.collection_num != 0) and (self.obj.collection_num != None):
            tempCollection = Collection.Collection(self.obj.collection_num)
            collectionID = tempCollection.GetNodeString()
        elif (self.obj.clip_num != 0) and (self.obj.clip_num != None):
            # We can skip loading the Clip Transcript to save load time
            tempClip = Clip.Clip(self.obj.clip_num, skipText=True)
            clipID = tempClip.id
            tempCollection = Collection.Collection(tempClip.collection_num)
            collectionID = tempCollection.GetNodeString()
        elif (self.obj.snapshot_num != 0) and (self.obj.snapshot_num != None):
            tempSnapshot = Snapshot.Snapshot(self.obj.snapshot_num)
            snapshotID = tempSnapshot.id
            tempCollection = Collection.Collection(tempSnapshot.collection_num)
            collectionID = tempCollection.GetNodeString()
            
        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Note ID
        id_edit = self.new_edit_box(_("Note ID"), v1, self.obj.id, maxLen=100)
        # Add the element to the sizer
        r1Sizer.Add(v1, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Series ID
        seriesID_edit = self.new_edit_box(_("Series ID"), v2, seriesID)
        # Add the element to the row sizer
        r2Sizer.Add(v2, 1, wx.EXPAND)
        seriesID_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v3 = wx.BoxSizer(wx.VERTICAL)
        # Episode ID
        episodeID_edit = self.new_edit_box(_("Episode ID"), v3, episodeID)
        # Add the element to the row sizer
        r2Sizer.Add(v3, 1, wx.EXPAND)
        episodeID_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Transcript ID
        transcriptID_edit = self.new_edit_box(_("Transcript ID"), v4, transcriptID)
        # Add the element to the row sizer
        r2Sizer.Add(v4, 1, wx.EXPAND)
        transcriptID_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v5 = wx.BoxSizer(wx.VERTICAL)
        # Collection ID
        collectionID_edit = self.new_edit_box(_("Collection ID"), v5, collectionID)
        # Add the element to the row sizer
        r3Sizer.Add(v5, 2, wx.EXPAND)
        collectionID_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r3Sizer, 0, wx.EXPAND)

        # Create a HORIZONTAL sizer for the next row
        r4Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v6 = wx.BoxSizer(wx.VERTICAL)
        # Clip ID
        clipID_edit = self.new_edit_box(_("Clip ID"), v6, clipID)
        # Add the element to the row sizer
        r4Sizer.Add(v6, 1, wx.EXPAND)
        clipID_edit.Enable(False)

        if TransanaConstants.proVersion:
            # Add a horizontal spacer to the row sizer        
            r4Sizer.Add((10, 0))

            # Create a VERTICAL sizer for the next element
            v7 = wx.BoxSizer(wx.VERTICAL)
            # Snapshot ID
            snapshotID_edit = self.new_edit_box(_("Snapshot ID"), v7, snapshotID)
            # Add the element to the row sizer
            r4Sizer.Add(v7, 1, wx.EXPAND)
            snapshotID_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r4Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r5Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v8 = wx.BoxSizer(wx.VERTICAL)
        # Comment layout
        noteTaker_edit = self.new_edit_box(_("Note Taker"), v8, self.obj.author, maxLen=100)
        # Add the element to the row sizer
        r5Sizer.Add(v8, 2, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r5Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the buttons
        self.create_buttons(sizer=btnSizer)
        # Add the button sizer to the main sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            mainSizer.Add((0, 2))

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
        self.SetSize(wx.Size(max(400, width), height))
        # Define the minimum size for this dialog as the current size, and define height as unchangeable
        self.SetSizeHints(max(400, width), height, -1, height)
        # Center the form on screen
        self.CenterOnScreen()
        
        # Set focus to the Note ID
        id_edit.SetFocus()


    def get_input(self):
        """Custom input routine."""
        # Inherit base input routine
        gen_input = Dialogs.GenForm.get_input(self)
        if gen_input:
            self.obj.id = gen_input[_('Note ID')]
            self.obj.author = gen_input[_('Note Taker')]
        else:
            self.obj = None
        
        return self.obj



class AddNoteDialog(NotePropertiesForm):
    """Dialog used when adding a new Note."""

    def __init__(self, parent, id, seriesNum=0, episodeNum=0, transcriptNum=0, collectionNum=0, clipNum=0, snapshotNum=0):
        obj = Note.Note()
        obj.author = DBInterface.get_username()
        obj.series_num = seriesNum
        obj.episode_num = episodeNum
        obj.transcript_num = transcriptNum
        obj.collection_num = collectionNum
        obj.clip_num = clipNum
        obj.snapshot_num = snapshotNum
        NotePropertiesForm.__init__(self, parent, id, _("Add Note"), obj)

class EditNoteDialog(NotePropertiesForm):
    """Dialog used when editing Note properties."""

    def __init__(self, parent, id, note_object):
        NotePropertiesForm.__init__(self, parent, id, _("Note Properties"), note_object)
