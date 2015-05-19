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

""" This module implements the Transcript Properties form. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

# import wxPython
import wx
# import Transana's Database Interface
import DBInterface
# Import Transana's Dialogs
import Dialogs
# Import the Transcript Object
import Transcript
# import Python's os module
import os

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
        self.id_edit = self.new_edit_box(_("Transcript ID"), lay, self.obj.id, maxLen=100)

        # Series ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.SameAs(self.id_edit, wx.Right, 10)   # 10 from id_edit
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
        lay.top.Below(self.id_edit, 10)              # 10 below id_edit
        lay.left.SameAs(self.id_edit, wx.Left, 0)     # Same as id_edit (10 from left)
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        transcriber_edit = self.new_edit_box(_("Transcriber"), lay, self.obj.transcriber, maxLen=100)

        # Comment layout
        lay = wx.LayoutConstraints()
        lay.top.Below(transcriber_edit, 10)     # 10 under transcriber
        lay.left.SameAs(self.id_edit, wx.Left, 0)     # Same as id_edit (10 from left)
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        comment_edit = self.new_edit_box(_("Comment"), lay, self.obj.comment, maxLen=255)

        # File to Import layout
        lay = wx.LayoutConstraints()
        lay.top.Below(comment_edit, 10)         # 10 under comment
        lay.left.SameAs(self.id_edit, wx.Left, 0)     # Same as id_edit (10 from left)
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

        self.id_edit.SetFocus()

    def OnBrowseClick(self, event):
        """ Method for when Browse button is clicked """
        # Allow for RTF, TXT or *.* combinations
        dlg = wx.FileDialog(None, wildcard="Rich Text Format Files (*.rtf)|*.rtf|Text Files(*.txt)|*.txt|All Files (*.*)|*.*", style=wx.OPEN)
        # Get a file selection from the user
        if dlg.ShowModal() == wx.ID_OK:
            # If the user clicks OK, set the file to import to the selected path.
            self.rtfname_edit.SetValue(dlg.GetPath())
            # If the ID field is blank ...
            if self.id_edit.GetValue() == '':
                # Get the base file name just selected
                tempFilename = os.path.basename(dlg.GetPath())
                # Split off the file extension
                (tempobjid, tempExt) = os.path.splitext(tempFilename)
                # Name the Transcript object after the imported Transcript
                self.id_edit.SetValue(tempobjid)
        # Destroy the File Dialog
        dlg.Destroy()

    def get_input(self):
        """Show the dialog and return the modified Series Object.  Result
        is None if user pressed the Cancel button."""
        # inherit parent method from Dialogs.Gen(eric)Form
        d = Dialogs.GenForm.get_input(self)
        # If the Form is created (not cancelled?) ...
        if d:
            # Set the Transcript ID
            self.obj.id = d[_('Transcript ID')]
            # Set the Transcriber
            self.obj.transcriber = d[_('Transcriber')]
            # Set the Comment
            self.obj.comment = d[_('Comment')]
            # Get the Media File to import
            fname = d[_('RTF File to import  (optional)')]
            # If a media file is entered ...
            if fname:
                # ... start exception handling ...
                try:
                    # Open the file
                    f = open(fname, "r")
                    # Read the file straight into the Transcript Text
                    self.obj.text = f.read()
                    # if the text does NOT have an RTF header ...
                    if (self.obj.text[:5].lower() != '{\\rtf'):
                        # ... add "txt" to the start of the file to signal that it's probably a text file
                        self.obj.text = 'txt\n' + self.obj.text
                    # Close the file
                    f.close()
                # If exceptions are raised ...
                except:
                    # ... we don't need to do anything here.  (Error message??)
                    # The consequence is probably that the Transcript Text will be blank.
                    pass
        # If there's no input from the user ...
        else:
            # ... then we can set the Transcript Object to None to signal this.
            self.obj = None
        # Return the Transcript Object we've created / edited
        return self.obj
        
class AddTranscriptDialog(TranscriptPropertiesForm):
    """Dialog used when adding a new Transcript."""

    def __init__(self, parent, id, episode):
        obj = Transcript.Transcript()
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
