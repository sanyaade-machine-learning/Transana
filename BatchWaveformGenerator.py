# Copyright (C) 2003 - 2005 The Board of Regents of the University of Wisconsin System 
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

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import wx

import ctypes
import Dialogs
import TransanaConstants
import TransanaGlobal
import WaveformProgress
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

        # Browse button layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.right.SameAs(self.panel, wx.Right, 10)
        lay.height.AsIs()
        browse = wx.Button(self.panel, wx.ID_FILE1, _("Select Files"), wx.DefaultPosition)
        browse.SetConstraints(lay)
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)

        lay = wx.LayoutConstraints()
        lay.top.Below(browse, 10)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.right.SameAs(self.panel, wx.Right, 10)
        lay.height.AsIs()
        label = wx.StaticText(self.panel, -1, _('Selected Files:'))
        label.SetConstraints(lay)

        lay = wx.LayoutConstraints()
        lay.top.Below(label, 10)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.right.SameAs(self.panel, wx.Right, 10)
        lay.bottom.SameAs(self.panel, wx.Bottom, 50)
        self.fileList = wx.ListBox(self.panel, -1)
        self.fileList.SetConstraints(lay)


        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

    def get_input(self):
        """ Get the Input values from the Batch Waveform Generator form and process the selected files """
        val = self.ShowModal()

        if val == wx.ID_OK:
            BWGProgress = wx.ProgressDialog(_('Batch Waveform Generator'), _('Extracting %s') % self.fileList.GetString(0), style=wx.PD_AUTO_HIDE)
            BWGProgress.CenterOnScreen()
            (xPos, yPos) = BWGProgress.GetPositionTuple()
            BWGProgress.SetPosition(wx.Point(xPos, yPos - 150))
            data = self.fileList.GetStrings()

            for loop in range(0, self.fileList.GetCount()):
                if loop > 0:
                    BWGProgress.Update(int((float(loop) / float(self.fileList.GetCount())) * 100), _('Extracting %s') % self.fileList.GetString(loop))
                
                # Remember the original File Name that is passed in
                originalFilename = self.fileList.GetString(loop)
                filename = self.fileList.GetString(loop)
                # Separate path and filename
                (path, filename) = os.path.split(filename)
                # break filename into root filename and extension
                (filenameroot, extension) = os.path.splitext(filename)
                # Build the correct filename for the Wave File
                self.waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
                # Build the correct filename for the Waveform Graphic
                # self.waveformFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.bmp')
                self.waveformFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.png')
                # Delete the wave file, if it exists.
                if os.path.exists(self.waveFilename):
                    os.remove(self.waveFilename)
                # Delete the waveform file, if it exists.
                if os.path.exists(self.waveformFilename):
                    os.remove(self.waveformFilename)
                
                try:
                    if not os.path.exists(TransanaGlobal.configData.visualizationPath):
                        os.makedirs(TransanaGlobal.configData.visualizationPath)
                    # Build the progress box's label
                    label = _("Extracting %s\nfrom %s") % (self.waveFilename, originalFilename)
                    # If the user accepts, create and display the Progress Dialog
                    progressDialog = WaveformProgress.WaveformProgress(self, self.waveFilename, label)
                    progressDialog.Show()

                    # Set the Extraction Parameters to produce small WAV files
                    bits = 8           # Use 8 for the bit rate (Legal values are 8 and 16)
                    decimation = 16    # Use the highest level of decimation (Legal values are 1, 2, 4, 8, and 16)
                    mono = 1           # 1 = mono, 0 = stereo

                    # Define callback function
                    def Progress(percent, seconds):
                        return progressDialog.Update(percent, seconds)

                    # Create the data structure needed by ctypes to pass the callback function as a Pointer to the DLL/Shared Library
                    callback = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_int, ctypes.c_int)
                    # Define a pointer to the callback function
                    callbackFunction = callback(Progress)
                    # Call the wcerAudio DLL/Shared Library's ExtractAudio function
                    if (os.name == "nt"):
                        dllvalue = ctypes.cdll.wceraudio.ExtractAudio(originalFilename, self.waveFilename, bits, decimation, mono, callbackFunction)
                    else:
                        wceraudio = ctypes.cdll.LoadLibrary("wceraudio.dylib")
                        dllvalue = wceraudio.ExtractAudio(originalFilename, self.waveFilename, bits, decimation, mono, callbackFunction)

                    if dllvalue != 0:
                        try:
                            os.remove(self.waveFilename)
                        except:
                            dlg = Dialogs.ErrorDialog(self, _('Unable to create waveform for file "%s"\nError Code: %s') % (originalFilename, dllvalue))
                            dlg.ShowModal()
                            dlg.Destroy()

                    # Close the Progress Dialog when the DLL call is complete
                    progressDialog.Close()
                    
                except:
                    dlg = Dialogs.ErrorDialog(self, _('Unable to create Waveform Directory.\n%s\n%s') % (sys.exc_info()[0], sys.exc_info()[1]))
                    dlg.ShowModal()
                    dlg.Destroy()
                    dllvalue = 1  # Signal that the WAV file was NOT created!                        

            BWGProgress.Update(100)
            BWGProgress.Destroy()
            
            return data
        else:
            return None     # Cancel

    def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        fileTypesString = _("""All supported media files (*.mpg, *.avi, *.mp3, *.wav)|*.mpg;*.mpeg;*.avi;*.mp3;*.wav|All video files (*.mpg, *.avi, *.mov)|*.mpg;*.mpeg;*.avi;*.mov|All audio files (*.mp3, *.wav, *.au, *.snd)|*.mp3;*.wav;*.au;*.snd|MPEG files (*.mpg)|*.mpg;*.mpeg|AVI files (*.avi)|*.avi|MOV files (*.mov)|*.mov|MP3 files (*.mp3)|*.mp3|WAV files (*.wav)|*.wav|All files (*.*)|*.*""")
        fs = wx.FileSelector(_("Select an XML file to import"),
                        self.lastPath,
                        "",
                        "", 
                        fileTypesString, 
                        wx.OPEN | wx.FILE_MUST_EXIST)
        # If user didn't cancel ..
        if fs != "":
            if self.fileList.FindString(fs) == wx.NOT_FOUND:
                self.fileList.Append(fs)
            # Remember the path of the last file selected for use next time
            (self.lastPath, filename) = os.path.split(fs)
