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

""" This module implements the Batch Waveform Generator for Transana.
    Because of interface overlap, it also implements Batch Episode Generation! """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Jonathan Beavers <jonathan.beavers@gmail.com>'

DEBUG = False
if DEBUG:
    print "BatchWaveformGenerator DEBUG is ON!!"

import wx
import ctypes
# import Transana's Common Dialogs
import Dialogs
# Import Transana's Miscellaneous routines
import Misc
# import Transana's Constants
import TransanaConstants
# import Transana's Globals
import TransanaGlobal
# import Transana's waveform progress routines
import WaveformProgress
import locale                 # import locale so we can get the default system encoding for Unicode Waveforming
import os
import sys

class BatchFileProcessor(Dialogs.GenForm):
    """ Batch File Processor, used for Batch Waveform Generator and Batch Episode Creation """
    def __init__(self, parent, mode):
        """ Initialize the Batch Waveform Generator form.  "mode" is either "waveform" for the
            Batch Waveform Generator or "episode" for the Batch Episode Creation routine. """
        # Remember the mode passed in.
        self.mode = mode
        # Based on the mode passed in, set the title and help context for the File Selection form
        if self.mode == 'waveform':
            formTitle = _('Batch Waveform Generator')
            helpContext = 'Batch Waveform Generator'
        elif self.mode == 'episode':
            formTitle = _("Batch Episode Creation")
            helpContext = 'Batch Episode Creation'
        else:
            print "UNKNOWN BATCHFILEPROCESSOR MODE"

        # Create the Dialog box for the File Selection Form            
        Dialogs.GenForm.__init__(self, parent, -1, formTitle, (500, 550), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext=helpContext)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Define the minimum size for this dialog
        self.SetSizeHints(500, 500)
        # Use the Video Root for the initial file path, if there is a Video Root
        if TransanaGlobal.configData.videoPath <> '':
            self.lastPath = TransanaGlobal.configData.videoPath
        # If there is no Video Root, use the Transana Program Directory
        else:
            self.lastPath = os.path.dirname(sys.argv[0])

        # Create the controls that will populate the File Selection Dialog window.
        browse = wx.Button(self.panel, wx.ID_FILE1, _("Select Files"), wx.DefaultPosition)
        directories = wx.Button(self.panel, wx.ID_FILE2, _("Select Directory"), wx.DefaultPosition)
        label = wx.StaticText(self.panel, -1, _('Selected Files:'))
        self.fileList = wx.ListBox(self.panel, -1, style=wx.LB_MULTIPLE)
        if self.mode == "waveform":
            self.overwrite = wx.CheckBox(self.panel, -1, _('Overwrite existing wave files?'))
        remfile = wx.Button(self.panel, wx.ID_FILE3, _("Remove Selected File(s)"), wx.DefaultPosition)

        # Bind the events that we'll need.
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)
        wx.EVT_BUTTON(self, wx.ID_FILE2, self.BrowseDirectories)
        wx.EVT_BUTTON(self, wx.ID_FILE3, self.RemoveSelected)

        # Browse button layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.width.PercentOf(self.panel, wx.Width, 46)
        lay.height.AsIs()
        browse.SetConstraints(lay)

        # dirdialog button layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)
        lay.right.SameAs(self.panel, wx.Right, 10)
        lay.width.PercentOf(self.panel, wx.Width, 46)
        lay.height.AsIs()
        directories.SetConstraints(lay)

        # Place the label for the forthcoming file list.

        # 500-2*(margin_sizes=10)=480 - width of label
        # this provides us with a "good enough" number of pixels to
        # use so that we can indent the label properly.
        rightpos = 480 - label.GetBestSizeTuple()[0]
        lay = wx.LayoutConstraints()
        lay.top.Below(directories, 12)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.right.SameAs(self.panel, wx.Right, rightpos)
        lay.height.AsIs()
        label.SetConstraints(lay)

        if self.mode == "waveform":
            # place an option to overwrite existing wave files to the upper-right
            # of the ListBox.

            # 500-2*(margin_sizes=10)=480 - width of checkbox
            # this provides us with a "good enough" number of pixels to
            # use so that we can indent the CheckBox properly.
            leftpos = 480 - self.overwrite.GetBestSizeTuple()[0]
            lay = wx.LayoutConstraints()
            lay.top.Below(directories, 12)
            lay.right.SameAs(self.panel, wx.Right, 10)
            lay.left.SameAs(self.panel, wx.Left, leftpos)
            lay.height.AsIs()
            self.overwrite.SetConstraints(lay)

        # place the actual ListBox.
        lay = wx.LayoutConstraints()
        lay.top.Below(label, 2)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.right.SameAs(self.panel, wx.Right, 10)
        lay.bottom.SameAs(self.panel, wx.Bottom, 50)
        self.fileList.SetConstraints(lay)

        # place the remove file button!
        lay = wx.LayoutConstraints()
        lay.top.Below(self.fileList, 17)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        remfile.SetConstraints(lay)

        # handle screen layout
        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

    def get_input(self):
        """ Get the Input values from the Batch Waveform Generator form and process the selected files """
        # Show the form and get the user response to the form.
        val = self.ShowModal()
        # Get the File List data from the form
        data = self.fileList.GetStrings()
        # If the user pressed OK and there are files in the list ...
        if (val == wx.ID_OK) and (not self.fileList.IsEmpty()):
            # If we're doing waveforms, extract them.  We don't need to know about other objects
            # to do waveform extraction.  (This logic could just as easily be part of the calling
            # routine -- we'd just need to pass teh file list and the value of the "overwrite" checkbox.)
            if self.mode == "waveform":
                # Iterate through the file list
                for loop in range(0, self.fileList.GetCount()):
                    # Get the current filename
                    filename = self.fileList.GetString(loop)
                    # Remember the original File Name that is passed in
                    originalFilename = filename
                    # Split the path off of the file name
                    (path, filename) = os.path.split(filename)
                    # Split the extension off the file name
                    (filenameroot, extension) = os.path.splitext(filename)
                    # Build the progress box's label
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_("Extracting %s\nfrom %s"), 'utf8')
                    else:
                        prompt = _("Extracting %s\nfrom %s")
                    # Build the filename for the extracted audio out of the filename parts
                    self.waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
                    # If there is no extracted audio file, OR if we're over-writing extracted audio ...
                    if not(os.path.exists(self.waveFilename)) or self.overwrite.GetValue():
                        # Build the correct filename for the Waveform Graphic
                        self.waveformFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.png')
                        # Create the Waveform Progress Dialog
                        self.progressDialog = WaveformProgress.WaveformProgress(self, self.waveFilename, prompt % (self.waveFilename, originalFilename))
                        # Tell the Waveform Progress Dialog to handle the audio extraction modally.
                        self.progressDialog.Extract(originalFilename, self.waveFilename)
                        # Okay, we're done with the Progress Dialog here!
                        self.progressDialog.Destroy()
                        # We just have to assume that audio extraction worked.  Signal success!
                    else:
                        continue
            
            # We don't DO anything here for the Batch Episode Creation routine.  We just
            # return the File List to the calling routine for processing!!  The calling routine
            # knows about the database tree, whereas this object doesn't.
            
            return data
        else:
            return None     # Cancel

    def OnBrowse(self, evt):
        """ Invoked when the user presses the Get Files button. """
        # Get Transana's File Filter definitions
        fileTypesString = TransanaConstants.fileTypesString
        # Create a File Open dialog.
        # Changed from FileSelector to FileDialog to allow multiple file selections.
        fs = wx.FileDialog(self, _('Select a media file to process:'),
                        self.lastPath,
                        "",
                        fileTypesString, 
                        wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE)
        # Select "All Media Files" as the initial Filter
        fs.SetFilterIndex(1)
        # Show the dialog and get user response.  If OK ...
        if fs.ShowModal() == wx.ID_OK:
            # ... For all files selected ...
            for filenm in fs.GetPaths():
                # If the file is NOT already in the File List...
                if not(filenm in self.fileList.GetStrings()):
                    # ... add the filename to the file list.
                    self.fileList.Append(filenm)
            # Remember the path of the first file selected for use next time
            (self.lastPath, filename) = os.path.split(fs.GetPath())
        # Destroy the File Dialog
        fs.Destroy()

    def BrowseDirectories(self, evt):
        """ Invoked when the user presses the Get Directory button """
        # Build a dialog that requests that the user select a directory
        dlg = wx.DirDialog(self, _('Select a directory that contains video or audio files:'), self.lastPath, style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        # Show the dialog and see if the user pressed OK
        if dlg.ShowModal() == wx.ID_OK:
            # Find all teh medica files in the specified path
            self.FindMediaFiles(dlg.GetPath())
            # Remember the path for reuse
            self.lastPath = dlg.GetPath()
        # Destroy the Dialog
        dlg.Destroy

    def RemoveSelected(self, evt):
        """ Remove any selected files from self.fileList """
        # Get the selected files
        selected = self.fileList.GetSelections()
        # Get a list of ALL files
        tempItems = self.fileList.GetItems()
        # Create an empty list
        newList = []
        # Iterate through all the files
        for x in range(0, len(tempItems)):
            # If an item has not been selected for removal ...
            if x not in selected:
                # ... then add it to the new list (of items to retain)
                newList.append(tempItems[x])
        # Set the file List to the list of items that should be retained.
        self.fileList.SetItems(newList)

    def FindMediaFiles(self, directory):
        """ Find files in a given directory with extensions that match those found in TransanaConstants.mediaFileTypes. """
        # Starting with the specified directory, traverse through all files and subdirectories
        for root, dirs, files in os.walk(directory):
            # for all the files in the current directory ...
            for name in files:
                # ... get the file extension of the current file
                extension = name[name.rfind('.')+1:]
                # If the extension is in the list of supported media types ...
                if extension.lower() in TransanaConstants.mediaFileTypes:
                    # ... add the file to the File List
                    self.fileList.Append(os.path.join(root, name))
