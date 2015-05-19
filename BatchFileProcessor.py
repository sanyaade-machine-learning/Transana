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

""" This module implements the Batch Waveform Generator for Transana.
    Because of interface overlap, it also implements Batch Episode Generation! """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Jonathan Beavers <jonathan.beavers@gmail.com>'

# import wxPython
import wx
# import Python's ctypes
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
# import Python's locale module
import locale
# import Python's os module
import os
# import Python's sys module
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
        Dialogs.GenForm.__init__(self, parent, -1, formTitle, (500, 550), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext=helpContext)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Use the Video Root for the initial file path, if there is a Video Root
        if TransanaGlobal.configData.videoPath <> '':
            self.lastPath = TransanaGlobal.configData.videoPath
        # If there is no Video Root, use the Transana Program Directory
        else:
            self.lastPath = os.path.dirname(sys.argv[0])

        # Create the controls that will populate the File Selection Dialog window.
        # Create a Select Files button
        browse = wx.Button(self.panel, wx.ID_FILE1, _("Select Files"), wx.DefaultPosition)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            r1Sizer.Add((2, 0))
        # Add the element to the sizer
        r1Sizer.Add(browse, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r1Sizer.Add((10, 0))

        # Create a Select Directory button
        directories = wx.Button(self.panel, wx.ID_FILE2, _("Select Directory"), wx.DefaultPosition)
        # Add the element to the sizer
        r1Sizer.Add(directories, 1, wx.EXPAND)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            r1Sizer.Add((2, 0))

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)
        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a label for the list of Files
        label = wx.StaticText(self.panel, -1, _('Selected Files:'))
        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Add the element to the sizer
        r2Sizer.Add(label, 0)

        # For Audio Extraction (waveform creation) ...
        if self.mode == "waveform":
            # ... create a checkbox for overwriting files
            self.overwrite = wx.CheckBox(self.panel, -1, _('Overwrite existing wave files?'))
            # place an option to overwrite existing wave files to the upper-right
            # of the ListBox.

            # Add a horizontal spacer to the row sizer        
            r2Sizer.Add((10, 0), 1, wx.EXPAND)

            # Add the element to the sizer
            r2Sizer.Add(self.overwrite, 0, wx.ALIGN_RIGHT)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)
        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create the List of Files listbox
        self.fileList = wx.ListBox(self.panel, -1, style=wx.LB_MULTIPLE)
        # Add the element to the sizer
        r3Sizer.Add(self.fileList, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r3Sizer, 1, wx.EXPAND)
        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a Remove Files button
        remfile = wx.Button(self.panel, wx.ID_FILE3, _("Remove Selected File(s)"), wx.DefaultPosition)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            btnSizer.Add((2, 0))
        # Add the element to the Button Sizer
        btnSizer.Add(remfile, 0)

        # Add the buttons
        self.create_buttons(sizer=btnSizer)
        # Add the button sizer to the main sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            mainSizer.Add((0, 2))

        # Bind the events that we'll need.
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)
        wx.EVT_BUTTON(self, wx.ID_FILE2, self.BrowseDirectories)
        wx.EVT_BUTTON(self, wx.ID_FILE3, self.RemoveSelected)

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
        self.SetSize(wx.Size(max(500, width), max(500, height)))
        # Define the minimum size for this dialog as the current size, and define height as unchangeable
        self.SetSizeHints(max(500, width), max(500, height))
        # Center the form on screen
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
                        self.progressDialog = WaveformProgress.WaveformProgress(self, prompt % (self.waveFilename, originalFilename))
                        # Tell the Waveform Progress Dialog to handle the audio extraction modally.
                        self.progressDialog.Extract(originalFilename, self.waveFilename)
                        # Get the Error Log that may have been created
                        errorLog = self.progressDialog.GetErrorMessages()
                        # Okay, we're done with the Progress Dialog here!
                        self.progressDialog.Destroy()
                        # If the user cancelled the audio extraction ...
                        if (len(errorLog) == 1) and (errorLog[0] == 'Cancelled'):
                            break
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
