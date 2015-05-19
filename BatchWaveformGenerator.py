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

""" This module implements the Batch Waveform Generator for Transana. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

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

        if (val == wx.ID_OK) and (not self.fileList.IsEmpty()):
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Extracting %s'), 'utf8') % self.fileList.GetString(0)
            else:
                prompt = _('Extracting %s') % self.fileList.GetString(0)
            BWGProgress = wx.ProgressDialog(_('Batch Waveform Generator'), prompt, style=wx.PD_AUTO_HIDE)
            BWGProgress.CenterOnScreen()
            (xPos, yPos) = BWGProgress.GetPositionTuple()
            BWGProgress.SetPosition(wx.Point(xPos, yPos - 150))
            data = self.fileList.GetStrings()

            for loop in range(0, self.fileList.GetCount()):
                if loop > 0:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Extracting %s'), 'utf8') % self.fileList.GetString(loop)
                    else:
                        prompt = _('Extracting %s') % self.fileList.GetString(loop)
                    BWGProgress.Update(int((float(loop) / float(self.fileList.GetCount())) * 100), prompt)
                
                filename = self.fileList.GetString(loop)
                # Deal with Mac Filename Encoding
                if 'wxMac' in wx.PlatformInfo:
                    filename = Misc.convertMacFilename(filename)
                # Remember the original File Name that is passed in
                originalFilename = filename
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
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_("Extracting %s\nfrom %s"), 'utf8') % (self.waveFilename, originalFilename)
                    else:
                        prompt = _("Extracting %s\nfrom %s") % (self.waveFilename, originalFilename)
                    # If the user accepts, create and display the Progress Dialog
                    progressDialog = WaveformProgress.WaveformProgress(self, self.waveFilename, prompt)
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
                    
                    # Wave Extraction doesn't work for Unicode Filenames without a little encoding help.
                    # It appears that UTF-8 isn't a good way to interact with the DLLs, so we need to change
                    # the encoding of the filenames to avoid an exception from the DLL.  At least on Windows,
                    # changing to the default system encoding seems to work.
                    if 'unicode' in wx.PlatformInfo:
                        # Get the default system encoding, and encode the file names into that encoding.
                        defEnc = locale.getdefaultlocale()[1]
                        originalFilename = originalFilename.encode(defEnc)
                        self.waveFilename = self.waveFilename.encode(defEnc)

                    # Call the wcerAudio DLL/Shared Library's ExtractAudio function
                    if (os.name == "nt") and (extension in ['.mpg', '.mpeg', '.wav', '.mp3']):  # '.avi', 
                        dllvalue = ctypes.cdll.wceraudio.ExtractAudio(originalFilename, self.waveFilename, bits, decimation, mono, callbackFunction)
                    elif (os.name == "nt"):
                        import pyMedia_audio_extract

                        self.waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
                        tempFile = pyMedia_audio_extract.ExtractWaveFile(originalFilename, self.waveFilename, progressDialog)
                        if tempFile:
                            pyMedia_audio_extract.DecimateWaveFile(tempFile, self.waveFilename, decimation, progressDialog)
                            dllvalue = 0
                        else:
                            dllvalue = 1
                        os.remove(tempFile)
                    # Waveform Extraction on Mac
                    else:
                        wceraudio = ctypes.cdll.LoadLibrary("wceraudio.dylib")
                        # Let's see if we have a legal filename for waveforming on the Mac
                        # Create a string of legal characters for the file names
                        allowedChars = TransanaConstants.legalFilenameCharacters
                        msg = ''
                        # check each character in the file name string
                        for char in originalFilename.decode(defEnc):
                            # If the character is illegal ...
                            if allowedChars.find(char) == -1:
                                if TransanaConstants.singleUserVersion:
                                    msg = _(u'There is an unsupported character in the Media File Name.\n\n"%s" includes the "%s" character, \nwhich Transana on the Mac does not support at this time.  Please rename your folders \nand files so that they do not include characters that are not part of English.') % (originalFilename.decode(defEnc), char)
                                else:
                                    msg = _(u'There is an unsupported character in the Media File Name.\n\n"%s" includes the "%s" character, \nwhich Transana on the Mac does not support at this time.  Please arrange to use waveform \nfiles created on Windows or rename your folders and files so that they \ndo not include characters that are not part of English.') % (originalFilename.decode(defEnc), char)                                    
                                dllvalue = 0
                                contin = False
                                break
                        if msg == '':
                            dllvalue = wceraudio.ExtractAudio(originalFilename, self.waveFilename, bits, decimation, mono, callbackFunction)
                        else:
                            dlg = Dialogs.ErrorDialog(self, msg)
                            dlg.ShowModal()
                            dlg.Destroy()
                            

                    if dllvalue != 0:
                        try:
                            os.remove(self.waveFilename)
                        except:
                            if DEBUG:
                                import traceback
                                traceback.print_exc(file=sys.stdout)

                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Unable to create waveform for file "%s"\nError Code: %s'), 'utf8') % (originalFilename, dllvalue)
                            else:
                                prompt = _('Unable to create waveform for file "%s"\nError Code: %s') % (originalFilename, dllvalue)
                            dlg = Dialogs.ErrorDialog(self, prompt)
                            dlg.ShowModal()
                            dlg.Destroy()

                    # Close the Progress Dialog when the DLL call is complete
                    progressDialog.Close()
                    
                except UnicodeEncodeError:
                    # If this exception is raised, the media filename contains a character that the default system
                    # encoding can't cope with.  On Windows, let's see what pyMedia does with it.
                    if 'wxMSW' in wx.PlatformInfo:
                        import pyMedia_audio_extract

                        self.waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
                        tempFile = pyMedia_audio_extract.ExtractWaveFile(originalFilename, self.waveFilename, progressDialog)
                        if tempFile:
                            pyMedia_audio_extract.DecimateWaveFile(tempFile, self.waveFilename, decimation, progressDialog)
                            dllvalue = 0
                        else:
                            dllvalue = 1
                        os.remove(tempFile)
                        
                    else:
                        if DEBUG:
                            import traceback
                            traceback.print_exc(file=sys.stdout)

                        dllvalue = 1  # Signal that the WAV file was NOT created!
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('Unable to create waveform for file "%s"\nError Code: %s'), 'utf8') % (originalFilename, 'UnicodeEncodeError')
                        else:
                            prompt = _('Unable to create waveform for file "%s"\nError Code: %s') % (originalFilename, 'UnicodeEncodeError')
                        errordlg = Dialogs.ErrorDialog(self, prompt)
                        errordlg.ShowModal()
                        errordlg.Destroy()

                    
                    # Close the Progress Dialog when the DLL call is complete
                    progressDialog.Close()

                except:
                    if DEBUG:
                        import traceback
                        traceback.print_exc(file=sys.stdout)

                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Unable to create WaveformDirectory.\n%s\n%s'), 'utf8') % (sys.exc_info()[0], sys.exc_info()[1])
                    else:
                        prompt = _('Unable to create WaveformDirectory.\n%s\n%s') % (sys.exc_info()[0], sys.exc_info()[1])
                    dlg = Dialogs.ErrorDialog(self, prompt)
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
            # On the Mac with Unicode filenames, we're having file names improperly rejected.  This code tries
            # to correct that.
            else:
                # Do the comparison manually
                if self.fileList.GetString(self.fileList.FindString(fs)) != fs:
                    # if they're not REALLY the same, add the item.
                    self.fileList.Append(fs)
            # Remember the path of the last file selected for use next time
            (self.lastPath, filename) = os.path.split(fs)
