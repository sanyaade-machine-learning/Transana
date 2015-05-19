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

""" This module implements the Batch Waveform Generator for Transana. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Jonathan Beavers <jonathan.beavers@gmail.com>'

DEBUG = False
if DEBUG:
    print "BatchWaveformGenerator DEBUG is ON!!"

import wx

import ctypes
import Dialogs
# Import Transana's Miscellaneous routines
import Misc
import TransanaConstants
import TransanaGlobal
import WaveformProgress
import locale                 # import locale so we can get the default system encoding for Unicode Waveforming
import os
import sys

class BatchWaveformGenerator(Dialogs.GenForm):
    """ Batch Waveform Generator """
    def __init__(self, parent):
        """ Initialize the Batch Waveform Generator form """
        Dialogs.GenForm.__init__(self, parent, -1, _('Batch Waveform Generator'), (500, 550), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Batch Waveform Generator')

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Define the minimum size for this dialog as the initial size
        self.SetSizeHints(500, 500)

        if TransanaGlobal.configData.videoPath <> '':
            self.lastPath = TransanaGlobal.configData.videoPath
        else:
            self.lastPath = os.path.dirname(sys.argv[0])

        # Create the controls that will populate the BatchWaveformGenerator window.
        browse = wx.Button(self.panel, wx.ID_FILE1, _("Select Files"), wx.DefaultPosition)
        directories = wx.Button(self.panel, wx.ID_FILE2, _("Select Directory"), wx.DefaultPosition)
        label = wx.StaticText(self.panel, -1, _('Selected Files:'))
        self.fileList = wx.ListBox(self.panel, -1, style=wx.LB_MULTIPLE)
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

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

    def get_input(self):
        """ Get the Input values from the Batch Waveform Generator form and process the selected files """

        val = self.ShowModal()
        data = self.fileList.GetStrings()

        if (val == wx.ID_OK) and (not self.fileList.IsEmpty()):
            for loop in range(0, self.fileList.GetCount()):
                #WaveformProgress.Extract(self.fileList.GetString(loop), 'test.wav')
                filename = self.fileList.GetString(loop)
                # Build the progress box's label
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("Extracting %s\nfrom %s"), 'utf8')
                else:
                    prompt = _("Extracting %s\nfrom %s")

                # Remember the original File Name that is passed in
                originalFilename = filename

                filename = filename
                (path, filename) = os.path.split(filename)

                # prepare the filenames for extraction
                (filenameroot, extension) = os.path.splitext(filename)

                self.waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
                #print os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')

                if not(os.path.exists(self.waveFilename)) or self.overwrite.GetValue():
                    # Build the correct filename for the Waveform Graphic
                    self.waveformFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.png')
                    # Build the file name for the extracted audio
                    self.waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
                    # Create the Waveform Progress Dialog
                    self.progressDialog = WaveformProgress.WaveformProgress(self, self.waveFilename, prompt % (self.waveFilename, originalFilename))
                    # Tell the Waveform Progress Dialog to handle the audio extraction modally.
                    self.progressDialog.Extract(originalFilename, self.waveFilename)
                    # Okay, we're done with the Progress Dialog here!
                    self.progressDialog.Destroy()
                    # We just have to assume that audio extraction worked.  Signal success!
                else:
                    continue
                
            return data
        else:
            return None     # Cancel

    def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        fileTypesString = TransanaConstants.fileTypesString
        fs = wx.FileSelector(_('Select a video file to process:'),
                        self.lastPath,
                        "",
                        "*.mpg", 
                        fileTypesString, 
                        wx.OPEN | wx.FILE_MUST_EXIST)
        # If user didn't cancel ..
        if fs != "":
            if self.fileList.FindString(fs) == wx.NOT_FOUND:
                self.fileList.Append(fs)
            # On the Mac with Unicode filenames, we're having file names improperly rejected.  This code tries
            # to correct that.
            else:
                # Do the comparison manually
                if self.fileList.GetString(self.fileList.FindString(fs)) != fs:
                    # if they're not REALLY the same, add the item.
                    self.fileList.Append(fs)
            # Remember the path of the last file selected for use next time
            (self.lastPath, filename) = os.path.split(fs)

    def RemoveSelected(self, evt):
        """Remove any selected files from self.fileList."""
        selected = self.fileList.GetSelections()
        tempItems = self.fileList.GetItems()
        newList = []

        for x in range(0, len(tempItems)):
            if x not in selected:
                newList.append(tempItems[x])

        self.fileList.SetItems(newList)

    def BrowseDirectories(self, evt):
        dlg = wx.DirDialog(self, _('Select a directory that contains video files:'), '/', style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        if dlg.ShowModal() == wx.ID_OK:

            if DEBUG:
                print "BatchWaveFormGenerator.BrowseDirectories()=", dlg.GetPath()

            self.FindMediaFiles(dlg.GetPath())

        # Destroy the Dialog
        dlg.Destroy

    def FindMediaFiles(self, directory):
        """Find files in a given directory with extensions that match those found in TransanaConstants.mediaFileTypes."""
        for root, dirs, files in os.walk(directory):
            for name in files:
                extension = name[name.rfind('.')+1:]
                if extension.lower() in TransanaConstants.mediaFileTypes:
                    self.fileList.Append(os.path.join(root, name))
