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

""" This module implements the Transcript Properties form. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

import DBInterface
import Dialogs

import wx
from Transcript import *

class TranscriptPropertiesForm(Dialogs.GenForm):
    """Form containing Transcript fields."""

    def __init__(self, parent, id, title, transcript_object):
        self.width = 400
        self.height = 260
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(self.width, self.height), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Transcript Properties')
        # Define the minimum size for this dialog as the initial size, and define height as unchangeable
        self.SetSizeHints(self.width, self.height, -1, self.height)

        self.obj = transcript_object

        # Transcript ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 40)  # 40% width
        lay.height.AsIs()
        id_edit = self.new_edit_box(_("Transcript ID"), lay, self.obj.id)

        # Series ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.SameAs(id_edit, wx.Right, 10)   # 10 from id_edit
        lay.width.PercentOf(self.panel, wx.Width, 25)  # 25% width
        lay.height.AsIs()
        series_id_edit = self.new_edit_box(_("Series ID"), lay, self.obj.series_id)
        series_id_edit.Enable(False)

        # Episode ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)                # 10 from top
        lay.left.SameAs(series_id_edit, wx.Right, 10)   # 10 from series_id_edit
        lay.width.PercentOf(self.panel, wx.Width, 25)         # 25% width
        lay.height.AsIs()
        episode_id_edit = self.new_edit_box(_("Episode ID"), lay, self.obj.episode_id)
        episode_id_edit.Enable(False)

        # Transcriber layout
        lay = wx.LayoutConstraints()
        lay.top.Below(id_edit, 10)              # 10 below id_edit
        lay.left.SameAs(id_edit, wx.Left, 0)     # Same as id_edit (10 from left)
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        transcriber_edit = self.new_edit_box(_("Transcriber"), lay, self.obj.transcriber)

        # Title/Comment layout
        lay = wx.LayoutConstraints()
        lay.top.Below(transcriber_edit, 10)     # 10 under transcriber
        lay.left.SameAs(id_edit, wx.Left, 0)     # Same as id_edit (10 from left)
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        comment_edit = self.new_edit_box(_("Title/Comment"), lay, self.obj.comment)

        # File to Import layout
        lay = wx.LayoutConstraints()
        lay.top.Below(comment_edit, 10)         # 10 under comment
        lay.left.SameAs(id_edit, wx.Left, 0)     # Same as id_edit (10 from left)
        lay.right.SameAs(self.panel, wx.Right, 100)  # Leave room for the Browse button
        lay.height.AsIs()
        self.rtfname_edit = self.new_edit_box(_("RTF File to import  (optional)"), lay, '')

        # Browse button
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.rtfname_edit, wx.Top, 0)
        lay.left.SameAs(self.panel, wx.Right, -87)
        lay.height.AsIs()
        lay.width.AsIs()
        browse = wx.Button(self.panel, -1, _("Browse"))
        browse.SetConstraints(lay)
        wx.EVT_BUTTON(self, browse.GetId(), self.OnBrowseClick)

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

        id_edit.SetFocus()

    def OnBrowseClick(self, event):
        """ Method for when Browse button is clicked """
        dlg = wx.FileDialog(None, wildcard="*.rtf", style=wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.rtfname_edit.SetValue(dlg.GetPath())
        dlg.Destroy()

    def get_input(self):
        """Show the dialog and return the modified Series Object.  Result
        is None if user pressed the Cancel button."""
        d = Dialogs.GenForm.get_input(self)     # inherit parent method
        if d:
            self.obj.id = d[_('Transcript ID')]
            self.obj.transcriber = d[_('Transcriber')]
            self.obj.comment = d[_('Title/Comment')]
            fname = d[_('RTF File to import  (optional)')]
            if fname:
                try:
                    f = open(fname, "r")
                    self.obj.text = f.read()
                    f.close()
                except:
                    pass
        else:
            self.obj = None

        return self.obj
        
class AddTranscriptDialog(TranscriptPropertiesForm):
    """Dialog used when adding a new Transcript."""

    def __init__(self, parent, id, episode):
        obj = Transcript()
        obj.series_id = episode.series_id
        obj.episode_num = episode.number
        obj.episode_id = episode.id
        obj.clip_num = 0
        obj.transcriber = DBInterface.get_username()
        TranscriptPropertiesForm.__init__(self, parent, id, _("Add Transcript"), obj)


class EditTranscriptDialog(TranscriptPropertiesForm):
    """Dialog used when editing Transcript properties."""

    def __init__(self, parent, id, transcript_object):
        TranscriptPropertiesForm.__init__(self, parent, id, _("Transcript Properties"), transcript_object)

