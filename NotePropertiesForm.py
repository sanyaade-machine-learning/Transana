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

"""  This module implements the Note Properties form.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case <nacase@wisc.edu>'

import DBInterface
import Dialogs
import KWManager
import Misc

import Series
import Episode
import Transcript
import Collection
import Clip
import Note

import wx
import string

class NotePropertiesForm(Dialogs.GenForm):
    """Form containing Note fields."""

    def __init__(self, parent, id, title, note_object):
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(400, 260), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Notes')  # 'Notes' is the Help Context for Notes.  There is no entry for Note Properties at this time
        # Define the minimum size for this dialog as the initial size, and define height as unchangeable
        self.SetSizeHints(400, 260, -1, 260)

        self.obj = note_object
        seriesID = ''
        episodeID = ''
        transcriptID = ''
        collectionID = ''
        clipID = ''
        if (self.obj.series_num != 0) and (self.obj.series_num != None):
            tempSeries = Series.Series(self.obj.series_num)
            seriesID = tempSeries.id
        elif (self.obj.episode_num != 0) and (self.obj.episode_num != None):
            tempEpisode = Episode.Episode(self.obj.episode_num)
            episodeID = tempEpisode.id
            tempSeries = Series.Series(tempEpisode.series_num)
            seriesID = tempSeries.id
        elif (self.obj.transcript_num != 0) and (self.obj.transcript_num != None):
            tempTranscript = Transcript.Transcript(self.obj.transcript_num)
            transcriptID = tempTranscript.id
            tempEpisode = Episode.Episode(tempTranscript.episode_num)
            episodeID = tempEpisode.id
            tempSeries = Series.Series(tempEpisode.series_num)
            seriesID = tempSeries.id
        elif (self.obj.collection_num != 0) and (self.obj.collection_num != None):
            tempCollection = Collection.Collection(self.obj.collection_num)
            collectionID = tempCollection.id
        elif (self.obj.clip_num != 0) and (self.obj.clip_num != None):
            tempClip = Clip.Clip(self.obj.clip_num)
            clipID = tempClip.id
            tempCollection = Collection.Collection(tempClip.collection_num)
            collectionID = tempCollection.id
            
        ######################################################
        # Tedious GUI layout code follows
        ######################################################

        # Note ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        id_edit = self.new_edit_box(_("Note ID"), lay, self.obj.id, maxLen=100)

        # Series ID layout
        lay = wx.LayoutConstraints()
        lay.top.Below(id_edit, 10)                 # 10 under Note ID
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 30)  # 31% width
        lay.height.AsIs()
        seriesID_edit = self.new_edit_box(_("Series ID"), lay, seriesID)
        seriesID_edit.Enable(False)

        # Episode ID layout
        lay = wx.LayoutConstraints()
        lay.top.Below(id_edit, 10)                 # 10 under Note ID
        lay.left.RightOf(seriesID_edit, 10)        # 10 right of Series ID
        lay.width.PercentOf(self.panel, wx.Width, 30)  # 31% width
        lay.height.AsIs()
        episodeID_edit = self.new_edit_box(_("Episode ID"), lay, episodeID)
        episodeID_edit.Enable(False)

        # Transcript ID layout
        lay = wx.LayoutConstraints()
        lay.top.Below(id_edit, 10)                 # 10 under Note ID
        lay.left.RightOf(episodeID_edit, 10)       # 10 right of Episode ID
        lay.width.PercentOf(self.panel, wx.Width, 30)  # 31% width
        lay.height.AsIs()
        transcriptID_edit = self.new_edit_box(_("Transcript ID"), lay, transcriptID)
        transcriptID_edit.Enable(False)

        # Collection ID layout
        lay = wx.LayoutConstraints()
        lay.top.Below(seriesID_edit, 10)           # 10 under Series ID
        lay.left.SameAs(seriesID_edit, wx.Left)  # Same as Series ID
        lay.width.PercentOf(self.panel, wx.Width, 30)  # 31% width
        lay.height.AsIs()
        collectionID_edit = self.new_edit_box(_("Collection ID"), lay, collectionID)
        collectionID_edit.Enable(False)

        # Clip ID layout
        lay = wx.LayoutConstraints()
        lay.top.Below(seriesID_edit, 10)           # 10 under Series ID
        lay.left.RightOf(collectionID_edit, 10)    # 10 right of Collection ID
        lay.width.PercentOf(self.panel, wx.Width, 30)  # 31% width
        lay.height.AsIs()
        clipID_edit = self.new_edit_box(_("Clip ID"), lay, clipID)
        clipID_edit.Enable(False)

        # Title/Comment layout
        lay = wx.LayoutConstraints()
        lay.top.Below(collectionID_edit, 10)       # 10 under Collection ID
        
# We're not ready to show a Comment field here yet!
# The Commenting is laid out this way to place the Note Taker field under the Collection Field, but to
# preserve the positioning when we want to add Comments in.

#        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
#        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
#        lay.height.AsIs()
#        comment_edit = self.new_edit_box("Title/Comment", lay, self.obj.comment)
#        comment_edit.Enable(False)

        # Note taker layout
#        lay = wx.LayoutConstraints()
#        lay.top.Below(comment_edit, 10)       # 10 under Comment

        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        noteTaker_edit = self.new_edit_box(_("Note Taker"), lay, self.obj.author, maxLen=100)

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

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

    def __init__(self, parent, id, seriesNum=0, episodeNum=0, transcriptNum=0, collectionNum=0, clipNum=0):
        obj = Note.Note()
        obj.author = DBInterface.get_username()
        obj.series_num = seriesNum
        obj.episode_num = episodeNum
        obj.transcript_num = transcriptNum
        obj.collection_num = collectionNum
        obj.clip_num = clipNum
        NotePropertiesForm.__init__(self, parent, id, _("Add Note"), obj)

class EditNoteDialog(NotePropertiesForm):
    """Dialog used when editing Note properties."""

    def __init__(self, parent, id, note_object):
        NotePropertiesForm.__init__(self, parent, id, _("Note Properties"), note_object)
