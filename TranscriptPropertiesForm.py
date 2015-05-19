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

""" This module implements the Transcript Properties form. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

# import wxPython
import wx
# import Transana's Database Interface
import DBInterface
# Import Transana's Dialogs
import Dialogs
# import Transana's Episode object
import Episode
# Import the Transcript Object
import Transcript
# import Python's os module
import os

class TranscriptPropertiesForm(Dialogs.GenForm):
    """Form containing Transcript fields."""

    def __init__(self, parent, id, title, transcript_object):
        """ Create the Transcript Properties form """
        self.width = 500
        self.height = 260
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(self.width, self.height), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Transcript Properties')

        # Define the form's main object
        self.obj = transcript_object

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Add the Transcript ID element
        self.id_edit = self.new_edit_box(_("Transcript ID"), v1, self.obj.id, maxLen=100)
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
        # Add the Series ID element
        series_id_edit = self.new_edit_box(_("Series ID"), v2, self.obj.series_id)
        # Add the element to the row sizer
        r2Sizer.Add(v2, 1, wx.EXPAND)
        # Disable Series ID
        series_id_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v3 = wx.BoxSizer(wx.VERTICAL)
        # Add the Episode ID element
        episode_id_edit = self.new_edit_box(_("Episode ID"), v3, self.obj.episode_id)
        # Add the element to the row sizer
        r2Sizer.Add(v3, 1, wx.EXPAND)
        # Disable Episode ID
        episode_id_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Add the Transcriber element
        transcriber_edit = self.new_edit_box(_("Transcriber"), v4, self.obj.transcriber, maxLen=100)
        # Add the element to the row sizer
        r3Sizer.Add(v4, 3, wx.EXPAND | wx.RIGHT, 10)

        # Create a VERTICAL sizer for the next element
        v7 = wx.BoxSizer(wx.VERTICAL)
        # Add the Min Transcript Width element
        mintranscriptwidth = self.new_edit_box(_("Min. Width"), v7, str(self.obj.minTranscriptWidth))
        # Add the element to the row sizer
        r3Sizer.Add(v7, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r3Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r4Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v5 = wx.BoxSizer(wx.VERTICAL)
        # Add the Comment element
        comment_edit = self.new_edit_box(_("Comment"), v5, self.obj.comment, maxLen=255)
        # Add the element to the row sizer
        r4Sizer.Add(v5, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r4Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r5Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v6 = wx.BoxSizer(wx.VERTICAL)
        # Add the Import File element
        self.rtfname_edit = self.new_edit_box(_("RTF/XML/TXT File to import  (optional)"), v6, '')
        # Add the element to the row sizer
        r5Sizer.Add(v6, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r5Sizer.Add((10, 0))

        # Add the Browse Button
        browse = wx.Button(self.panel, -1, _("Browse"))
        # Add the Browse Method to the Browse Button
        wx.EVT_BUTTON(self, browse.GetId(), self.OnBrowseClick)
        # Add the element to the sizer
        r5Sizer.Add(browse, 0, wx.ALIGN_BOTTOM)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            r5Sizer.Add((2, 0))

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
        self.SetSize(wx.Size(max(self.width, width), height))
        # Define the minimum size for this dialog as the current size, and define height as unchangeable
        self.SetSizeHints(max(self.width, width), height, -1, height)
        # Center the form on screen
        self.CenterOnScreen()
        # Set focus to the Transcript ID
        self.id_edit.SetFocus()

    def OnBrowseClick(self, event):
        """ Method for when Browse button is clicked """
        # Load the source episode
        tmpEpisode = Episode.Episode(num=self.obj.episode_num)
        # Get the directory for the MAIN media file name
        dirName = os.path.dirname(tmpEpisode.media_filename)
        # Allow for RTF, XML, TXT or *.* combinations
        dlg = wx.FileDialog(None, defaultDir=dirName,
                            wildcard="Transcript Import Formats (*.rtf, *.xml, *.txt)|*.rtf;*.xml;*.txt|Rich Text Format Files (*.rtf)|*.rtf|XML Files (*.xml)|*.xml|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                            style=wx.OPEN)
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
            # Set Minimum Transcript Width
            try:
                self.obj.minTranscriptWidth = int(d[_("Min. Width")])
            except:
                self.obj.minTranscriptWidth = 0
            # Get the Media File to import
            fname = d[_('RTF/XML/TXT File to import  (optional)')]
            # If a media file is entered ...
            if fname:
                # ... start exception handling ...
                try:
                    # Open the file
                    f = open(fname, "r")
                    # Read the file straight into the Transcript Text
                    self.obj.text = f.read()
                    # if the text does NOT have an RTF header ...
                    if (self.obj.text[:5].lower() != '{\\rtf') and (self.obj.text[:5].lower() != '<?xml'):
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
