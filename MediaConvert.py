# -*- coding: cp1252 -*-
# Copyright (C) 2011 - 2014 The Board of Regents of the University of Wisconsin System 
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

""" This dialog implements the Transana Media Conversion Dialog class.
    It requires the Transana-specific FFMpeg build for the platform being used. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "MediaConvert DEBUG is ON!!!"

# import wxPython
import wx

# if running stand-alone (for testing)
if __name__ == '__main__':
    # This module expects i18n.  Enable it here.
    __builtins__._ = wx.GetTranslation
# if NOT running stand-alone
else:
    # Import Transana's Database Interface
    import DBInterface
# Import Transana's Dialogs
import Dialogs
# Import Transana's Miscellaneous Functions
import Misc
# Import Transana's Constants
import TransanaConstants
# import Transana's Global variables
import TransanaGlobal
# import Transana's Waveform Progress Dialog (no longer just for Waveform Audio Extration!)
import WaveformProgress

# import the Python exceptions module
import exceptions
# Import Python's os and sys modules
import os, sys
# import Python's shutil for fast file copies
import shutil

# We MUST disable MPEG-1, MPEG-2, and MP3 formats for legal reasons.
ENABLE_MPG = False  # DO NOT CHANGE THIS VALUE
# We never got a response regarding the legalities of using the MOV format.
# We've disabled it to be certain we are within the law.
ENABLE_MOV = False

if ENABLE_MPG or ENABLE_MOV:
    print "MediaConvert:  MPG or MOV format enabled!!"

# This simple derived class let's the user drop files onto an edit box
class EditBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, editbox):
        wx.FileDropTarget.__init__(self)
        self.editbox = editbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        self.editbox.SetValue(files[0])

class MediaConvert(wx.Dialog):
    """ Transana's Media Conversion Tool Dialog Box. """
    
    def __init__(self, parent, fileName='', clipStart=0, clipDuration=0, clipName='', snapshot=False):
        """ Initialize the MediaConvert Dialog Box. """

        # There's a bug.  I think it's an interaction between OS X 10.7.5 and earlier (but not 10.8.4), wxPython (version
        # unknown, but I've seen it in 2.8.12.1, 2.9.4.0, and a pre-release build of 2.9.5.0), and the build process.
        # Basically, if you open a wxFileDialog object, something bad happens with wxLocale such that subsequent calls
        # to wxProcess don't encode unicode file names correctly, so FFMpeg fails.  Only OS X 10.7.5, only when Transana
        # has been built (not running through the interpreter), only following the opening of a wxFileDialog, and only
        # with filenames with accented characters or characters from non-English languages.

        # The resolution to this bug is to reset the Locale here, based on Transana's current wxLocale setting in the
        # menuWindow object.
        self.locale = wx.Locale(TransanaGlobal.menuWindow.locale.Language)

        # Remember the File Name passed in
        self.fileName = fileName
        # If we're on Windows ...
        if 'wxMSW' in wx.PlatformInfo:
            # ... specify a Temporary Path for handling non-cp1252 files
            self.tmpPath = 'C:\\Temp_Transana'
        # Set the Temporary Filename to a blank.  (This is used for non-cp1252 filenames, which ffmpeg can't handle on Windows.)
        self.tmpFileName = ''
        # Remember Clip Start Point
        self.clipStart = clipStart
        # Remember Clip Duration
        self.clipDuration = clipDuration
        # Remember the Clip Name
        self.clipName = clipName
        # Remember snapshot setting (Are we capturing a still from a video file?)
        self.snapshot = snapshot
        # We need to know if we have successfully taken a snapshot
        self.snapshotSuccess = False
        # Initialize the process variable
        self.process = None

        # Initialize all media file variables
        self.Reset()
        
        # Create the Dialog
        wx.Dialog.__init__(self, parent, -1, _('Media File Conversion'), size=wx.Size(600, 700), style=wx.DEFAULT_DIALOG_STYLE | wx.THICK_FRAME)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the main Sizer, which will hold the box1, box2, etc. and boxButton sizers
        box = wx.BoxSizer(wx.VERTICAL)

        # Add the Source File label
        lblSource = wx.StaticText(self, -1, _("Source Media File:"))
        box.Add(lblSource, 0, wx.TOP | wx.LEFT | wx.RIGHT, 10)
        
        # Create the box1 sizer, which will hold the source file and its browse button
        box1 = wx.BoxSizer(wx.HORIZONTAL)

        # If we are in the DEMO version and are not taking a snapshot ...
        if TransanaConstants.demoVersion and not snapshot:
            # ... use the file name to indicate that the Media Conversion tool is disabled
            fileName = _('Media Conversion is disabled in the Demo version.')

        # Create the Source File text box
        self.txtSrcFileName = wx.TextCtrl(self, -1, fileName)
        # If we are exporting a Clip ...
        if (self.clipDuration > 0) or (TransanaConstants.demoVersion and not snapshot):
            # ... then we need to disable the Source Media File text box
            self.txtSrcFileName.Enable(False)
        # If we're NOT exporting a Clip ...
        else:
            # Make the Source File a File Drop Target
            self.txtSrcFileName.SetDropTarget(EditBoxFileDropTarget(self.txtSrcFileName))

        # Handle ALL changes to the source filename
        self.txtSrcFileName.Bind(wx.EVT_TEXT, self.OnSrcFileNameChange)
        box1.Add(self.txtSrcFileName, 1, wx.EXPAND)
        # Spacer
        box1.Add((4, 0))
        # Create the Source File Browse button
        self.srcBrowse = wx.Button(self, -1, _("Browse"))
        # If we are exporting a Clip ...
        if (self.clipDuration > 0) or (TransanaConstants.demoVersion and not snapshot):
            # ... then we need to disable the Source Media File Browse button
            self.srcBrowse.Enable(False)
        self.srcBrowse.Bind(wx.EVT_BUTTON, self.OnBrowse)
        box1.Add(self.srcBrowse, 0)
        # Add the Source Sizer to the Main Sizer
        box.Add(box1, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Format label
        lblFormat = wx.StaticText(self, -1, _("Format:"))
        box.Add(lblFormat, 0, wx.LEFT | wx.RIGHT, 10)

        # Add the Format Choice Box with no options initially
        self.format = wx.Choice(self, -1, choices=[])
        # Fix problems with Right-To-Left languages
        self.format.SetLayoutDirection(wx.Layout_LeftToRight)
        # Disable the Format box initially
        self.format.Enable(False)
        self.format.Bind(wx.EVT_CHOICE, self.OnFormat)
        box.Add(self.format, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Destination File label
        lblDest = wx.StaticText(self, -1, _("Destination Media File:"))
        box.Add(lblDest, 0, wx.LEFT | wx.RIGHT, 10)
        
        # Create the box2 horizontal sizer for the Destination File and its Browse button
        box2 = wx.BoxSizer(wx.HORIZONTAL)

        # Create the Destination File text box
        self.txtDestFileName = wx.TextCtrl(self, -1)
        # Set the layout Direction to LtR so file names don't get mangled
        self.txtDestFileName.SetLayoutDirection(wx.Layout_LeftToRight)
        # Disable the Destination File box initially
        self.txtDestFileName.Enable(False)
        box2.Add(self.txtDestFileName, 1, wx.EXPAND)
        # Spacer
        box2.Add((4, 0))
        # Create the Destination File Browse button
        self.destBrowse = wx.Button(self, -1, _("Browse"))
        # Disable the Destination File Browse button initially
        self.destBrowse.Enable(False)
        self.destBrowse.Bind(wx.EVT_BUTTON, self.OnBrowse)
        box2.Add(self.destBrowse, 0)

        # Add the Destination Sizer to the Main Sizer
        box.Add(box2, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Video Parameters label
        lblVideo = wx.StaticText(self, -1, _("Video Parameters:"))
        box.Add(lblVideo, 0, wx.LEFT | wx.RIGHT, 10)
        
        # Create the box3 horizontal sizer for the Video Parameters
        box3 = wx.BoxSizer(wx.HORIZONTAL)

        # Add the Video Size label
        lblVideoSize = wx.StaticText(self, -1, _("Size:"))
        box3.Add(lblVideoSize, 0, wx.RIGHT, 10)

#        This does not work!!!!
#        if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
#            style = wx.ALIGN_RIGHT
#        else:
#            style = 0
        # Add the Video Size choice box, initially empty
        self.videoSize = wx.Choice(self, -1, choices=[])   # , style=style
        # Fix problems with Right-To-Left languages
        self.videoSize.SetLayoutDirection(wx.Layout_LeftToRight)
        # Disable Video Size initially
        self.videoSize.Enable(False)
        box3.Add(self.videoSize, 1, wx.RIGHT, 10)

        # Add the Video Bit Rate label
        lblVideoBitrate = wx.StaticText(self, -1, _("Bit Rate: (kb/s)"))
        box3.Add(lblVideoBitrate, 0, wx.RIGHT, 10)

        # Add the Video Bitrate choice box
        self.videoBitrate = wx.Choice(self, -1, choices=[])
        # Disable the Video Bit Rate box initially
        self.videoBitrate.Enable(False)
        box3.Add(self.videoBitrate, 1, wx.RIGHT, 10)

        # Add the Video Parameters sizer to the Main Sizer
        box.Add(box3, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Audio Parameters label
        lblAudio = wx.StaticText(self, -1, _("Audio Parameters:"))
        box.Add(lblAudio, 0, wx.LEFT | wx.RIGHT, 10)
        
        # Create the box4 horizontal sizer for Audio Parameters
        box4 = wx.BoxSizer(wx.HORIZONTAL)

        # Add the Audio Bit Rate label
        lblAudioBitrate = wx.StaticText(self, -1, _("Bit Rate: (kb/s)"))
        box4.Add(lblAudioBitrate, 0, wx.RIGHT, 10)

        # Add the Audio Bit Rate choice box
        self.audioBitrate = wx.Choice(self, -1, choices=[])
        # Disable the Audio Choice Box initially
        self.audioBitrate.Enable(False)
        box4.Add(self.audioBitrate, 4, wx.RIGHT, 10)

        # Add the Audio Sample Rate label
        lblAudioSampleRate = wx.StaticText(self, -1, _("Sample Rate: (Hz)"))
        box4.Add(lblAudioSampleRate, 0, wx.RIGHT, 10)

        # Add the Audio Sample choice box
        self.audioSampleRate = wx.Choice(self, -1, choices=[])
        # Disable the Audio Sample box initially
        self.audioSampleRate.Enable(False)
        box4.Add(self.audioSampleRate, 4, wx.RIGHT, 10)

        # Add the Audio Parameters sizer to the Main Sizer
        box.Add(box4, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Still Images label
        lblStill = wx.StaticText(self, -1, _("Still Image Parameters:"))
        box.Add(lblStill, 0, wx.LEFT | wx.RIGHT, 10)
        
        # Create the box6 horizontal sizer for Still Image Parameters
        box6 = wx.BoxSizer(wx.HORIZONTAL)
        
        # Add the Still Frame Rate label
        lblStillFrameRate = wx.StaticText(self, -1, _("Seconds between images:"))
        box6.Add(lblStillFrameRate, 0, wx.RIGHT, 10)

        # Add the Still Frame Rate choice box
        stillFrameRateChoices = [_("20 seconds"), _("15 seconds"), _("10 seconds"), _("5 seconds"), _("1 second")]
        self.stillFrameRate = wx.Choice(self, -1, choices=stillFrameRateChoices)
        # Select the first entry
        self.stillFrameRate.SetSelection(0)
        # Disable the Still Framee box initially
        self.stillFrameRate.Enable(False)
        box6.Add(self.stillFrameRate, 4, wx.RIGHT, 10)

        # Add the Still Images Parameters sizer to the Main Sizer
        box.Add(box6, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # If we are exporting Clip Video ...
        if self.clipDuration > 0:
            # Add the Clip Parameters label
            lblClip = wx.StaticText(self, -1, _("Clip Parameters:"))
            box.Add(lblClip, 0, wx.LEFT | wx.RIGHT, 10)
        
            # Create the box5 horizontal sizer for the Clip Parameters
            box5 = wx.BoxSizer(wx.HORIZONTAL)

            # Add the Start Time label
            lblClipStart = wx.StaticText(self, -1, _("Clip Start Time:"))
            box5.Add(lblClipStart, 0, wx.RIGHT, 10)

            # Add the Clip Start Time box
            self.txtClipStartTime = wx.TextCtrl(self, -1, Misc.time_in_ms_to_str(self.clipStart, True))
            # Disable Clip Start Time
            self.txtClipStartTime.Enable(False)
            box5.Add(self.txtClipStartTime, 1, wx.RIGHT, 10)

            # Add the Duration label
            lblClipDuration = wx.StaticText(self, -1, _("Clip Duration:"))
            box5.Add(lblClipDuration, 0, wx.RIGHT, 10)

            # Add the Clip Duration box
            self.txtClipDuration = wx.TextCtrl(self, -1, Misc.time_in_ms_to_str(self.clipDuration, True))
            # Disable Clip Duration
            self.txtClipDuration.Enable(False)
            box5.Add(self.txtClipDuration, 1, wx.RIGHT, 10)

            # Add the Clip Parameters sizer to the Main Sizer
            box.Add(box5, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Information label
        lblMemo = wx.StaticText(self, -1, _("Information:"))
        box.Add(lblMemo, 0, wx.LEFT | wx.RIGHT, 10)

        # Add the Information text control
        self.memo = wx.TextCtrl(self, -1, style = wx.TE_MULTILINE)
        box.Add(self.memo, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the FFmpeg label
        lblFFmpeg = wx.StaticText(self, -1, _("Transana's Media File Conversion tool is powered by FFmpeg."))
        box.Add(lblFFmpeg, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Create the boxButtons sizer, which will hold the dialog box's buttons
        boxButtons = wx.BoxSizer(wx.HORIZONTAL)

        # Create a Convert button
        btnConvert = wx.Button(self, -1, _("Convert"))
        # Set this as the default button
        btnConvert.SetDefault()
        btnConvert.Bind(wx.EVT_BUTTON, self.OnConvert)
        boxButtons.Add(btnConvert, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT, 10)

        # If we are in the DEMO version and are not taking a snapshot ...        
        if TransanaConstants.demoVersion and not snapshot:
            # ... then disable the Convert button
            btnConvert.Enable(False)

        # Create a Close button
        btnClose = wx.Button(self, wx.ID_CANCEL, _("Close"))
        boxButtons.Add(btnClose, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT, 10)

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # Create a Help button
        btnHelp = wx.Button(self, -1, _("Help"))
        btnHelp.Bind(wx.EVT_BUTTON, self.OnHelp)
        boxButtons.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM, 10) 

        # Add the boxButtons sizer to the main box sizer
        box.Add(boxButtons, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Define box as the form's main sizer
        self.SetSizer(box)
        # Set this as the minimum size for the form.
        self.SetSizeHints(minW = self.GetSize()[0], minH = int(self.GetSize()[1] * 0.75))
        # Tell the form to maintain the layout and have it set the intitial Layout
        self.SetAutoLayout(True)
        self.Layout()
        # Position the form in the center of the screen
        self.CentreOnScreen()

        # If NOT running stand-alone (for testing) ...
        if __name__ != '__main__':
            # ... start with Transana's Video Root as the initial path
            self.lastPath = TransanaGlobal.configData.videoPath
        # If running stand-alone (for testing) ...
        else:
            # ... use the current directory as the initial path
            self.lastPath = os.path.dirname(sys.argv[0])

        # If a file name was passed in as a parameter ...
        if (self.fileName != '') and (not (TransanaConstants.demoVersion) or snapshot):
            # ... process that file to prepare for conversion
            self.ProcessMediaFile(self.fileName)

    def Reset(self):
        """ Initialize or Reset all variables associated with the Media File to be converted """
        # initialize the process variable
        self.process = None
        # Media File Duration
        self.duration = 0
        # Overall Bit Rate
        self.bitrate = 0
        # Number of Streams
        self.streams = 0
        # Contains a Video Stream?
        self.vidStream = False
        # Video Codec Name
        self.vidCodec = ''
        # Video Picture Format (see FFmpeg)
        self.vidPixFmt = ''
        # Video Bit Rate
        self.vidBitrate = 0
        # Video Picture Width
        self.vidSizeW = 0
        # Video Picture Height
        self.vidSizeH = 0
        # Video Frame Rate
        self.vidFrameRate = 0
        # Contains an Audio Stream?
        self.audStream = False
        # Audio Codec Name
        self.audCodec = ''
        # Audio Bit Rate
        self.audBitrate = 0
        # Audio Sample Rate
        self.audSampleRate = 0
        # Number of Audio Channels
        self.audChannels = 0
        # Temporary File Name
        self.tmpFileName = ''
        # File Extension for converted file
        self.ext = ''

    def ProcessMediaFile(self, inputFile):
        """ Process a Media File to see what it's made of, a pre-requisite to converting it.
            This process also populates the form's options. """
        
        # If we're messing with wxProcess, we need to define a function to clean up if we shut down!
        def __del__(self):

            # If a process has been defined ...
            if self.process is not None:
                # ... detach it
                self.process.Detach()
                #  ... Close its output
                self.process.CloseOutput()
                # ... and de-reference it
                self.process = None

        # Reset (re-initialize) all media file variables
        self.Reset()

        # Clear the Information box
        self.memo.Clear()

        # If we're on Windows ...
        if 'wxMSW' in wx.PlatformInfo:
            # Start exception handling
            try:
                # Initialize the file name for the error message
                testFileName = self.fileName
                # Find out if the file name can be converted to cp1252 encoding.
                # FFmpeg cannot handle files that are not cp1252 encodable on Windows!
                testFileName = self.fileName.encode('cp1252')
                # If the above doesn't trigger an exception ...
                # Separate path and file name
                (self.lastPath, filename) = os.path.split(self.fileName)
                # Find out what the output file name would be                
                self.SetOutputFilename(filename)
                # Get the Destination File Name from the text control
                testFileName = self.txtDestFileName.GetValue()
                # Re-disable the Destination File Name and Browse controls that were accidentally enabled by SetOutputFilename()
                self.txtDestFileName.Enable(False)
                self.destBrowse.Enable(False)
                # In Chinese etc., the output file name may include non-cp1252 characters even if the input file doesn't!
                testFileName = testFileName.encode('cp1252')
                
            # Handle Unicode Error exceptions
            except exceptions.UnicodeEncodeError:
                # If we end up here, we have a file or path that is NOT cp1252 encodable.
                # Inform the user
                prompt = unicode(_('The file name "%s" is not compatible with FFmpeg.  Your file is being temporarily copied for processing.'), 'utf8')
                self.memo.AppendText(prompt % testFileName + '\n\n')
                # Remember the ORIGINAL file name.  We'll need it later.
                self.tmpFileName = self.fileName
                # Break the file name into file name and extension
                (filename, ext) = os.path.splitext(self.fileName)
                # Create a temporary file name, made up of the cp1252 compatible tmpPath, "temp", and the file's correct extension
                self.fileName = os.path.join(self.tmpPath, 'Input' + ext)
                # We need to use this temporary file name for the inputFile too!
                inputFile = self.fileName
                # If the cp1252 compatible path does not exist ...
                if not os.path.exists(self.tmpPath):
                    # ... create it!
                    os.mkdir(self.tmpPath)
                # Copy the non-cp1252 file to the cp1252-compatible path with a cp1252-compatible file name.
                shutil.copyfile(self.tmpFileName, self.fileName)            

        # Reset all Conversion parameter items on the form
        # Clear and Disable the Destination File Name
        self.txtDestFileName.SetValue('')
        self.txtDestFileName.Enable(False)
        # Disable the Destination File Name Browse Button
        self.destBrowse.Enable(False)
        # Clear and Disable the Format choice box
        self.format.Clear()
        self.format.Enable(False)
        # Clear and Disable the Video Size choice box
        self.videoSize.Clear()
        self.videoSize.Enable(False)
        # Clear and Disable the Video Bit Rate choice box
        self.videoBitrate.Clear()
        self.videoBitrate.Enable(False)
        # Clear and Disable the Audio Bit Rate choice box
        self.audioBitrate.Clear()
        self.audioBitrate.Enable(False)
        # Clear and Disable the Audio Sampling Rate choice box
        self.audioSampleRate.Clear()
        self.audioSampleRate.Enable(False)

        # Be prepared to capture the wxProcess' EVT_END_PROCESS
        self.Bind(wx.EVT_END_PROCESS, self.OnEndProcess)
        
        # Windows requires that we change the default encoding for Python for the audio extraction code to work
        # properly with Unicode files (!!!)  This isn't needed on OS X, as its default file system encoding is utf-8.
        # See python's documentation for sys.getfilesystemencoding()
        if 'wxMSW' in wx.PlatformInfo:
            # Set the Python Encoding to match the File System Encoding
            wx.SetDefaultPyEncoding(sys.getfilesystemencoding())
        # Just use the File Name, no encoding needed
        tempMediaFilename = inputFile

        if DEBUG:
            self.memo.AppendText("MediaConvert.ProcessMediaFile():\n")
            self.memo.AppendText("%s\n  (%d) exists: %s\n" % (tempMediaFilename, len(tempMediaFilename), os.path.exists(tempMediaFilename)))

            statinfo = os.stat(tempMediaFilename)
            self.memo.AppendText("  size: %s\n\n" % statinfo.st_size)

            self.memo.AppendText("  type: %s, defaultPyEncoding: %s, filesystemencoding: %s, Transana's Encoding: %s\n\n" % (type(tempMediaFilename), wx.GetDefaultPyEncoding(), sys.getfilesystemencoding(), TransanaGlobal.encoding))

            for tmpX in range(len(tempMediaFilename)):
                self.memo.AppendText("  - %d  %s  %d\n" % ( tmpX, tempMediaFilename[tmpX], ord(tempMediaFilename[tmpX]) ))
            self.memo.AppendText("\n")

        # We need to build the Conversion command line.  Start with the executable path and name,
        # and add that we are using it embedded and want the second level of feedback (file information),
        # and specify the Input File name placeholder.
        process = '"' + TransanaGlobal.programDir + os.sep + 'ffmpeg_Transana" "-embedded" "2"'
        process += ' "-i" "%s"'
        # Create a wxProcess object
        self.process = wx.Process(self)
        # Call the wxProcess Object's Redirect method.  This allows us to capture the process's output!
        self.process.Redirect()
        # Encode the filenames to UTF8 so that unicode files are handled properly
        process = process.encode('utf8')

        if DEBUG:
            self.memo.AppendText("\n\nMedia Filename:\n")
            self.memo.AppendText("%s\n\n" % tempMediaFilename)
            self.memo.AppendText("\n\nProcess call:\n")
            self.memo.AppendText("%s\n\n" % process % tempMediaFilename)

        # Call the Audio Extraction program using wxExecute, capturing the output via wxProcess.  This call MUST be asynchronous. 
        self.pid = wx.Execute(process % tempMediaFilename.encode(sys.getfilesystemencoding()), wx.EXEC_ASYNC, self.process)

        # On Windows, we need to reset the encoding to UTF-8
        if 'wxMSW' in wx.PlatformInfo:
            wx.SetDefaultPyEncoding('utf_8')

    def OnEndProcess(self, event):
        """ End of wxProcess Event Handler """
        # If a process is defined ...
        if self.process is not None:

            if DEBUG:
                self.memo.AppendText("\n\nProcess pid %s calling OnEndProcess()\n\n" % self.pid)
                
            # Get the Process' Input Stream
            stream = self.process.GetInputStream()
            # If that stream can be read ...
            if stream.CanRead():

                if DEBUG:
                    self.memo.AppendText("stream.CanRead() call successful\n")
                    tmpParamCount = 0
                    
                # ... read it!
                text = stream.read()
                # Divide the text up into separate lines
                text = text.replace('\r\n', '\n')
                text = text.split('\n')
                # Process the input stream text one line at a time
                for line in text:
                    # Divide the line up into its separate parameters
                    param = line.split(' ')

                    if DEBUG:
                        tmpParamCount += 1
                        self.memo.AppendText("%8d '%s' " % (tmpParamCount, param[0]))
                        if len(param) > 1:
                            self.memo.AppendText("%s " % param[1])
                        if tmpParamCount < 16:
                            for par in param[2:]:
                                self.memo.AppendText("%s " % par)
                        self.memo.AppendText("(%s)" % len(line))
                        self.memo.AppendText("\n")
                        
                        
                    # If the line isn't blank and starts with an "x", indicating embedded feedback information ...
                    if (len(line) > 0) and (line[0] == 'x'):
                        # If the first parameter is just plain "x", we have a General Parameter
                        if param[0] == 'x':
                            # If Duration:
                            if param[1] == 'Duration:':
                                # Get the Media File Duration
                                self.duration = float(param[2])
                            # If Bitrate:
                            elif param[1] == 'Bitrate:':
                                # Get the General Bitrate
                                self.bitrate = int(param[2])
                            # If Streams:
                            elif param[1] == 'Streams:':
                                # Get the Number of Streams
                                self.streams = int(param[2])
                        # If the first paramer is "xv", we have a Video Parameter
                        elif param[0] == 'xv':
                            # If Stream
                            if param[1] == 'Stream':
                                # We have a Video Stream
                                self.vidStream = True
                            # If Codec:
                            elif param[1] == 'Codec:':
                                # Get the Video Codec
                                self.vidCodec = param[2]
                            # If Pix_Fmt:
                            elif param[1] == 'Pix_Fmt:':
                                # Get the Picture Format
                                self.vidPixFmt = param[2]
                            # If Bitrate:
                            elif param[1] == 'Bitrate:':
                                # Get the Video Bit Rate
                                self.vidBitrate = int(param[2])
                            # If FrameRate:
                            elif param[1] == 'FrameRate:':
                                # Get Video Frame Rate
                                self.vidFrameRate = float(param[2])
                                # If the Frame Rate is negative ...
                                if self.vidFrameRate < 0.0:
                                    # ... reset it to 0
                                    self.vidFrameRate == 0.0
                            # If Size:
                            elif param[1] == 'Size:':
                                # Get video Width ...
                                self.vidSizeW = int(param[2])
                                # ... and Height
                                self.vidSizeH = int(param[4])
                        # If the first paramer is "xa", we have an Audio Parameter
                        elif param[0] == 'xa':
                            # If Stream
                            if param[1] == 'Stream':
                                self.audStream = True
                                # We have an Audio Stream
                            # If Codec:
                            elif param[1] == 'Codec:':
                                # Get the Audio Codec
                                self.audCodec = param[2]
                            # If Bitrate:
                            elif param[1] == 'Bitrate:':
                                # Get the Audio Bit Rate
                                self.audBitrate = int(param[2])
                            # If SampleRate:
                            elif param[1] == 'SampleRate:':
                                # Get Audio Sample Rate
                                self.audSampleRate = int(param[2])
                            # If Channels:
                            elif param[1] == 'Channels:':
                                # Get the Number of Audio Channels
                                self.audChannels = int(param[2])
                        # Otherwise ...
                        else:
                            # ... we have an unknown parameter.  (This shouldn't occur.)
                            print "Unknown Parameter", param

            # Since the process has ended, destroy it.
            self.process.Destroy()
            # De-reference the process
            self.process = None

            # Report File Information to the user
            # Start by freezing the Information box to speed and smooth the process
            self.memo.Freeze()
            # Split the file name from the path, either for tmpFileName or self.fileName as appropriate
            if self.tmpFileName == '':
                (self.lastPath, filename) = os.path.split(self.fileName)
            else:
                (self.lastPath, filename) = os.path.split(self.tmpFileName)
            # Report Path and File Name
            self.memo.AppendText(unicode(_("File Path: %s\n"), 'utf8') % self.lastPath)
            self.memo.AppendText(unicode(_("File Name: %s\n\n"), 'utf8') % filename)
            # Report General Media File Parameters
            self.memo.AppendText(unicode(_('Duration: %s\n'), 'utf8') % Misc.time_in_ms_to_str(self.duration * 1000, True))
            self.memo.AppendText(unicode(_('Bitrate: %d kb/s\n'), 'utf8') % self.bitrate)
            self.memo.AppendText(unicode(_('Streams: %d\n'), 'utf8') % self.streams)
            # If there was a Video stream ...
            if self.vidStream:
                # ... report the Video settings
                self.memo.AppendText('\n')
                self.memo.AppendText(unicode(_('Video Stream:\n'), 'utf8'))
                self.memo.AppendText(unicode(_('  Codec: %s\n'), 'utf8') % self.vidCodec)
                self.memo.AppendText(unicode(_('  Picture Format: %s\n'), 'utf8') % self.vidPixFmt)
                self.memo.AppendText(unicode(_('  Size:  %d x %d\n'), 'utf8') % (self.vidSizeW, self.vidSizeH))
                if self.vidBitrate > 0.0:
                    self.memo.AppendText(unicode(_('  Bitrate: %d kb/s\n'), 'utf8') % self.vidBitrate)
                else:
                    self.memo.AppendText(unicode(_('  Bitrate: Unknown\n'), 'utf8'))
                if self.vidFrameRate > 0.0:
                    self.memo.AppendText(unicode(_('  Frame Rate: %0.2f fps\n'), 'utf8') % self.vidFrameRate)
                else:
                    self.memo.AppendText(unicode(_('  Frame Rate: Unknown\n'), 'utf8'))
            # If there was an Audio stream ...
            if self.audStream:
                # ... report the Audio settings
                self.memo.AppendText('\n')
                self.memo.AppendText(unicode(_('Audio Stream:\n'), 'utf8'))
                self.memo.AppendText(unicode(_('  Codec: %s\n'), 'utf8') % self.audCodec)
                self.memo.AppendText(unicode(_('  Bitrate: %d kb/s\n'), 'utf8') % self.audBitrate)
                self.memo.AppendText(unicode(_('  Sample Rate: %d Hz\n'), 'utf8') % self.audSampleRate)
                self.memo.AppendText(unicode(_('  Channels: %d\n'), 'utf8') % self.audChannels)
            self.memo.AppendText('\n')
            # Once this is finished, we can Thaw the control
            self.memo.Thaw()

            # If our file had at least one media stream ...
            if self.audStream or self.vidStream:
                # Clear the Format choice box
                self.format.Clear()
                # If we have a Video stream ...
                if self.vidStream:
                    # If we're not doing a video snapshot ...
                    if not self.snapshot:
                        if ENABLE_MPG:
                            # ... add the video options to the Format choice box
                            # The MPEG-1 option is NOT ALLOWED.  They want royalty fees of $15,000 a year minimum,
                            # which we cannot afford.
                            self.format.Append(_("MPEG-1 - An excellent choice for video files"))
                        # MPEG-4 is viable as long as we ship 50,000 units or fewer a year.
                        self.format.Append(_("MPEG-4 - Efficient video compression, with moderate responsiveness"))
                        if ENABLE_MOV:
                            # Add the MOV format option.
                            self.format.Append(_("MOV - An excellent choice for multiple simultaneous video files"))

                    # Clear the Video Size choice box
                    self.videoSize.Clear()
                    # Add sizes as long as they are SMALLER than source video size.  (Calculate Heights for
                    # "standard" Width options)
                    if self.vidSizeW >= 320:
                        st = _('%d x %d') % (320, int(self.vidSizeH * 320.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 400:
                        st = _('%d x %d') % (400, int(self.vidSizeH * 400.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 480:
                        st = _('%d x %d') % (480, int(self.vidSizeH * 480.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 560:
                        st = _('%d x %d') % (560, int(self.vidSizeH * 560.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 640:
                        st = _('%d x %d') % (640, int(self.vidSizeH * 640.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 720:
                        st = _('%d x %d') % (720, int(self.vidSizeH * 720.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 800:
                        st = _('%d x %d') % (800, int(self.vidSizeH * 800.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 1024:
                        st = _('%d x %d') % (1024, int(self.vidSizeH * 1024.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 1280:
                        st = _('%d x %d') % (1280, int(self.vidSizeH * 1280.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 1366:
                        st = _('%d x %d') % (1366, int(self.vidSizeH * 1366.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 1440:
                        st = _('%d x %d') % (1440, int(self.vidSizeH * 1440.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 1680:
                        st = _('%d x %d') % (1680, int(self.vidSizeH * 1680.0 / self.vidSizeW))
                        self.videoSize.Append(st)
                    if self.vidSizeW >= 1920:
                        st = _('%d x %d') % (1920, int(self.vidSizeH * 1920.0 / self.vidSizeW))
                        self.videoSize.Append(st)

                    # If the actual video size hasn't already been inserted, add it, UNLESS
                    # it is larger than 1920 pixels wide, which is our maximum
                    if (not self.vidSizeW in [1920, 1680, 1440, 1366, 1280, 1024, 800, 720, 640, 560, 480, 400, 320]) and (self.vidSizeW <= 1920):
                        st = str(self.vidSizeW) + ' x ' + str(self.vidSizeH)
                        self.videoSize.Append(st)
                    # If more than one Video Size option exists ...
                    if self.videoSize.GetCount() > 1:
                        # ... enable the Video Size choice box
                        self.videoSize.Enable(True)
                    # Set the last Size in the list, which should be the same as the original, or the
                    # largest allowable size if the original was too big.
                    self.videoSize.SetSelection(self.videoSize.GetCount() - 1)

                    # If the Video Bitrate is 0 (as seems to be true for AVI files) ...
                    if self.vidBitrate == 0:
                        # ... use the Overall Bit Rate as the Video Bit Rate.
                        # (I don't know if this is legit, but it's a best-I-can-do approximation.)
                        self.vidBitrate = self.bitrate
                    # If we are doing a video snapshot ...
                    if self.snapshot:
                        # ... let the user know this could take a while
                        self.memo.AppendText("\n" + _("Please note that the further into the media file your snapshot is, the longer this process will take.") + "\n")
                    # If we're NOT doing a snapshot, we should display Video Bitrate warnings if appropriate
                    else:
                        # If the Video Bit Rate exceeds 1500 kb/s ...
                        if self.vidBitrate > 1500:
                            # ... inform the user they may want to lower it
                            self.memo.AppendText(_("You should be able to reduce the Video Bit Rate setting to 1500 kb/s or less without noticable loss of quality.") + "\n")
                        # If the Video Bit Rate exceeds 500 kb/s ...
                        if self.vidBitrate > 500:
                            # ... inform the user they may want to lower it if planning on using Multiple Simultaneous Media Files
                            self.memo.AppendText(_("If you intend to use this video as one of multiple simultaneously displayed videos, you may want to reduce the Video Bit Rate setting to 1000 kb/s or less.") + "\n")
                            self.memo.AppendText("\n")

                    # Clear the Video Bit Rate choice box
                    self.videoBitrate.Clear()
                    # Start with a list of "default" video bit rates
                    bitrates = [100, 150, 200, 250, 300, 350, 500, 750, 1000, 1500, 2000, 3000, 5000]
                    # For each bit rate in the list ...
                    for bitrate in bitrates:
                        # ... if the File's Video Bit Rate is greater than the proposed bit rate setting ...
                        if self.vidBitrate >= bitrate:
                            # ... then add the proposed bit rate option to the choice box
                            self.videoBitrate.Append(str(bitrate))
                        # If not ...
                        else:
                            # ... we can stop looking at bit rates
                            break
                    # If the file's ACTUAL video bitrate wasn't in the default list ...
                    if not str(self.vidBitrate) in self.videoBitrate.GetStrings():
                        # ... then add it to the choice box too
                        self.videoBitrate.Append(str(self.vidBitrate))
                    # if the Video Bitrate exceeds 1500 ...
                    if self.vidBitrate >= 1500:
                        # ... then select 1500 as a reasonable bitrate
                        self.videoBitrate.SetSelection(9)
                    # if the video bitrate is less than 1500
                    else:
                        # ... select the highest video bit rate in the list, which should match the source file's
                        self.videoBitrate.SetSelection(self.videoBitrate.GetCount() - 1)
                    # if there is moe than one option in the list and we're not doing a video snapshot ...
                    if (self.videoBitrate.GetCount() > 1) and not self.snapshot:
                        # ... then enable the Video Bit Rate choice box
                        self.videoBitrate.Enable(True)
                # If we have an Audio stream ...
                if self.audStream:
                    # If we're not doing a video snapshot ...
                    if not self.snapshot:
                        if ENABLE_MPG:
                            # ... add the audio options to the Format choice box
                            # The MP3 option is NOT ALLOWED.  They want royalty fees of $15,000 a year minimum,
                            # which we cannot afford.
                            self.format.Append(_("MP3 - Compressed audio files"))
                        # Add the WAV file option
                        self.format.Append(_("WAV - Uncompressed audio files"))

                    # Clear the Audio Bit Rate choice box
                    self.audioBitrate.Clear()
                    # Start with a list of "default" audio bit rates
                    bitrates = [32, 48, 56, 64, 80, 96, 128, 144, 192, 224, 256, 320, 384]
                    # For each bit rate in the list ...
                    for bitrate in bitrates:
                        # ... if the File's Audio Bit Rate is greater than the proposed bit rate setting ...
                        if self.audBitrate >= bitrate:
                            # ... then add the proposed bit rate option to the choice box
                            self.audioBitrate.Append(str(bitrate))
                        # If not ...
                        else:
                            # ... we can stop looking at bit rates
                            break
                    # If the file's ACTUAL audio bitrate wasn't in the default list ...
                    if not str(self.audBitrate) in self.audioBitrate.GetStrings():
                        # ... then add it to the choice box too
                        self.audioBitrate.Append(str(self.audBitrate))
                    # If 192 kb/s is NOT among the audio bit rate options ...
                    if self.audioBitrate.FindString('192') == wx.NOT_FOUND:
                        # ... then select the highest audio bit rate in the list, which should match the source file's
                        self.audioBitrate.SetSelection(self.audioBitrate.GetCount() - 1)
                    # If 192 kb/s IS among the options ...
                    else:
                        # ... pick that.  It's "good enough" for analysis.
                        self.audioBitrate.SetSelection(self.audioBitrate.FindString('192'))
                    # If there are multiple audio bit rate options and we're not doing a video snapshot ...
                    if (self.audioBitrate.GetCount() > 1) and not self.snapshot:
                        # Enable the audio bit rate choice box
                        self.audioBitrate.Enable(True)

                    # Clear the Audio Sampling Rate choice box
                    self.audioSampleRate.Clear()
                    # Start with a list of "default" audio sampling rates
                    samplerates = [11025, 22050, 24000, 32000, 44100, 48000]
                    # For each sample rate in the list ...
                    for samplerate in samplerates:
                        # ... if the File's Audio Sample Rate is greater than the proposed sample rate setting ...
                        if self.audSampleRate >= samplerate:
                            # ... then add the proposed sample rate option to the choice box
                            self.audioSampleRate.Append(str(samplerate))
                        # If not ...
                        else:
                            # ... we can stop looking at sample rates
                            break
                    # If the file's ACTUAL audio sampling rate wasn't in the default list ...
                    if not str(self.audSampleRate) in self.audioSampleRate.GetStrings():
                        # ... then add it to the choice box too
                        self.audioSampleRate.Append(str(self.audSampleRate))
                    # Select the highest audio sampling rate in the list, which should match the source file's
                    self.audioSampleRate.SetSelection(self.audioSampleRate.GetCount() - 1)
                    # If there are multiple audio sampling rate options and we're not doing a video snapshot ...
                    if (self.audioSampleRate.GetCount() > 1) and not self.snapshot:
                        # Enable the audio sampling rate choice box
                        self.audioSampleRate.Enable(True)

                # If we have a video stream ...
                if self.vidStream:
                    # ... add the still images option to the Format choice box
                    self.format.Append(_("JPEG - Create still images from video files"))

                # If we have video and are on a Mac ...
                if ENABLE_MOV and self.vidStream and ('wxMac' in wx.PlatformInfo):
                    # Select the second item in the Format list, MOV video
                    self.format.SetSelection(1)
                # If we have audio only OR we're not on a Mac ...
                else:
                    # Select the first item in the Format list
                    self.format.SetSelection(0)
                # Enable the Format choice box
                self.format.Enable(True)

                # now that we have an input file and a format, we can generate an output file name
                self.SetOutputFilename(filename)
                
            # If we have no audio or video streams, we don't have a valid media file.
            else:
                # Report this to the user.
                self.memo.AppendText(_('The selected media file cannot be processed or converted.'))
                
    def OnConvert(self, event):
        """ Convert Button Press Event """
        # If we're converting to a still image ...
        if self.ext in ['.jpg']:
            # Check the file name, ensuring it has required the "%06d" parameter
            # First, let's break the file name up into path, filename, and ext
            (filename, ext) = os.path.splitext(self.txtDestFileName.GetValue())
            (path, filename) = os.path.split(filename)
            # if "%06d" isn't part of the filename ...
            if not ("%06d" in filename):
                # ... add it to the end of the file name ...
                filename += "_%06d"
                # ... and rebuild the file name from its components
                self.txtDestFileName.SetValue(os.path.join(path, filename) + ext)

        # If a numeric insertion is called for ...
        if '%06d' in self.txtDestFileName.GetValue():
            # Figure out the output file name
            tmpFileName = self.txtDestFileName.GetValue() % 1
        else:
            tmpFileName = self.txtDestFileName.GetValue()
        # See if the output file already exists
        if os.path.exists(tmpFileName):
            # If so, inform the user
            errmsg = unicode(_('File "%s" already exists.  Do you want to replace this file?'), 'utf8')
            errDlg = Dialogs.QuestionDialog(self, errmsg % tmpFileName, noDefault=True)
            result = errDlg.LocalShowModal()
            errDlg.Destroy()
            # If the user does not want to replace the file, ...
            if result == wx.ID_NO:
                # ... then exit this method immediately!
                return

        # Error Checking -- Initialize Error Message to blank
        errmsg = ""
        # See if the Source File exists (is available)
        if not os.path.exists(self.txtSrcFileName.GetValue()):
            errmsg += unicode(_('File "%s" not found.\n'), 'utf8') % self.txtSrcFileName.GetValue()
        # If we have a non-cp1252-compatible File and it didn't COPY correctly ...
        if (self.tmpFileName != '') and (not os.path.exists(self.fileName)):
            errmsg += unicode(_('File "%s" not found.  File "%s" did not copy correctly.\n'), 'utf8') % (self.fileName, self.txtSrcFileName.GetValue())
        # See if we have a valid Media File
        elif not self.audStream and not self.vidStream:
            errmsg += unicode(_('File "%s" is not a valid media file.\n'), 'utf8') % self.txtSrcFileName.GetValue()
        # Check the file extension
        if self.txtDestFileName.GetValue()[-4:].lower() != self.ext.lower():
            errmsg += unicode(_('The Destination File Name does not have the correct file extension.  It must end with "%s".'), 'utf8') % self.ext
        # Check for a DIFFERENT file name, so we're not over-writing media files!
        if self.txtSrcFileName.GetValue() == self.txtDestFileName.GetValue():
            errmsg += unicode(_("The Destination File Name cannot be the same as the Source File Name."), 'utf8')
        # If an error has been detected ...
        if errmsg != '':
            # ... create an Error Dialog and display the message to the user
            errDlg = Dialogs.ErrorDialog(self, errmsg)
            errDlg.ShowModal()
            errDlg.Destroy()

        # If no error has been detected ...
        else:
            # Windows requires that we change the default encoding for Python for the audio extraction code to work
            # properly with Unicode files (!!!)  This isn't needed on OS X, as its default file system encoding is utf-8.
            # See python's documentation for sys.getfilesystemencoding()
            if 'wxMSW' in wx.PlatformInfo:
                # Set the Python Encoding to match the File System Encoding
                wx.SetDefaultPyEncoding(sys.getfilesystemencoding())
            # We need to build the Extraction command line in stages.  Start with the executable path and name,
            # and add that we are using it embedded and want the first level of feedback (progress information),
            # and specify the Input File name placeholder.
            FFmpegCommand = '"' + TransanaGlobal.programDir + os.sep + 'ffmpeg_Transana" "-embedded" "1" "-i" "%s"'

            # Specify image size.  If we are creating a Video file ...
            if self.vidStream and (self.ext in ['.mpg', '.mp4', '.mov', '.jpg']):
                # ... Determine the current Video Size selection, and divide it up into its component parts
                size = self.videoSize.GetStringSelection().split(' ')
                # Supply the FFmpeg "-s" parameter and open the data quotes
                FFmpegCommand += ' "-s" "'
                # Build the size value.  (This essentially removes the internal spaces from the string)
                for x in size:
                    FFmpegCommand += x
                # Close the data quotes.
                FFmpegCommand += '"'

            # Specify video bitrate and some additional parameters.  If we are creating a Video file ...
            if self.vidStream and (self.ext in ['.mpg', '.mp4', '.mov']):
                # If we are creating an MPEG-1 file ...
                if self.ext == '.mpg':
                    # ... specify the video codec as mpeg1video
                    FFmpegCommand += ' "-vcodec" "mpeg1video"'
                # If we are creating an MPEG-4 file ...
                elif self.ext == '.mp4':
                    # ... specify Four Motion Vector (mpeg4) and h.263 advanced introacoding / mpeg2 ac prediction
                    FFmpegCommand += ' "-flags" "+mv4+aic"'

                # Add the Video Bit Rate specification
                FFmpegCommand += ' "-vb" "%dk"' % int(self.videoBitrate.GetStringSelection())

                # if the Frame Rate is not UNKNOWN ...
                if self.vidFrameRate > 0.0:
                    # HD video with high frame rates (eg. 59.96 fps) don't play smoothly.
                    # Frame Rate reduction causes problems if set to "29.97" or "30", but is okay at "29"
                    if self.vidFrameRate > 30:
                        # Let's max the Frame Rate out at 29 fps.
                        FFmpegCommand += ' "-r" "29"'
                        # Let's inform the user we changed their frame rate!
                        self.memo.AppendText("\n" + _("Frame Rate reduced from %0.2f fps to 29 fps.") % self.vidFrameRate)
                    # Otherwise ...
                    else:
                        # ... use the existing frame rate
                        FFmpegCommand += ' "-r" "%0.2f"' % self.vidFrameRate

            # If we have an Audio Stream to process ...
            if self.audStream and not self.ext in ['.jpg']:
                # If we are creating an MPEG-1 file ...
                if self.ext == '.mpg':
                    # ... specify the audio codes as mp2
                    FFmpegCommand += ' "-acodec" "mp2"'

                # Get the desired Audio Bitrate from the form
                tmpAudioBitrate = int(self.audioBitrate.GetStringSelection())
                # If we are creating an MPEG-1 file ...
                if self.ext == '.mpg':
                    # Check to see if the desired Audio Bit Rate is in the options allowed by the MP2 specification.  If not ...
                    if not int(self.audioBitrate.GetStringSelection()) in [32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384]:
                        # ... display a message to the user ...
                        self.memo.AppendText("\n" + _("MPEG-1 supports only limited audio bit rate options.  Over-riding Audio Bit Rate setting."))
                        # ... if the user has multiple options ...
                        if self.audioBitrate.GetSelection() > 0:
                            # ... pick the option just smaller than the user selected.  All the values in the list except the source file's
                            #     original setting are legal values, so this MUST be legal!
                            tmpAudioBitrate = int(self.audioBitrate.GetString(self.audioBitrate.GetSelection() - 1))
                        # If there's only one value in the options list ...
                        else:
                            # ... just use 64.  (This was somewhat arbitrary, but is unlikely to be used.)
                            tmpAudioBitrate = 64
                # Add the Audio Bit Rate to the Conversion Command
                FFmpegCommand += ' "-ab" "%dk"' % tmpAudioBitrate

                # If we are creating an MPEG-1 file and we are supposed to use a Sample Rate less than 44,100 Hz ...
                if (self.ext == '.mpg') and (int(self.audioSampleRate.GetStringSelection()) < 32000):
                    # ... inform the user of the smallest legal Sample Rate value for MP2 audio
                    self.memo.AppendText("\n" + _("MPEG-1 requires an Audio Sample Rate of at least 32,000.  Over-riding Audio Sample Rate.") + "\n")
                    # ... and set the value to 32000
                    FFmpegCommand += ' "-ar" "%d"' % 32000
                # Otherwise ...
                else:
                    # ... use the Sample Rate from the form
                    FFmpegCommand += ' "-ar" "%d"' % int(self.audioSampleRate.GetStringSelection())

                # If there are more than 2 audio channels ...
                if self.audChannels > 2:
                    # ... then let's reduce the number down to just 2.
                    FFmpegCommand += ' "-ac" "2"'
                    
            if not self.ext in ['.jpg']:
                # For best quality media files, the FFmpeg site suggests the following:
                #   -mbd rd                Macroblock Decision Algorithm "use best rate distortion"
                #   -trellis 2             rate-distortion optimal quantization (whatever that means)
                #   -cmp 2                 full pel me compare function (whatever that means)
                #   -subcmp 2              sub pel me compare function (whatever that means)
                FFmpegCommand += ' "-mbd" "rd" "-trellis" "2" "-cmp" "2" "-subcmp" "2"'

            # If we are creating an MPEG-1 file ...
            if self.ext == '.mpg':
                # ... the FFmpeg web site recommends a "group picture size" of 100 and a pass value of 1/2
                FFmpegCommand += ' "-g" "100" "-pass" "1/2"'
            # if we are creating an MPEG-4 file ...
            elif self.ext == '.mp4':
                # ... the FFmpeg web site recommends a "group picture size" of 300 and a pass value of 1/2
                FFmpegCommand += ' "-g" "300"'
                # When bundled on OS X, this argument causes problems.
                # Or perhaps it's due to network and permissions, as it returns a "permission denied" error.
                # Let's leave it off everywhere!
                if False or (not 'wxMac' in wx.PlatformInfo):
                    FFmpegCommand += ' "-pass" "1/2"'
            
            # The Transana Demo restricts the length of file conversion to 10 minutes
            if TransanaConstants.demoVersion and (self.duration > 600):
                FFmpegCommand += ' "-t" "600"'

            # If we're producing still images ...
            if self.ext in ['.jpg']:
                # Extract the extension of the source file name
                (srcName, srcExt) = os.path.splitext(self.txtSrcFileName.GetValue())
                # Some video formats have proven to be less reliable than others.  They seem to
                # work well enough if we request 4 frames.
                # Specifically, MPEG formats only seem to work with every third frame.  Weird.
                # The value 4 was determined through trial-and-error.
                numFramesForStill = 4

                # AVI and WMV formats appear to have a frame rate of float(-1.#IND00), which also shows up as
                # string('nan').  To check for this, we have to typecast the Frame Rate as a string.
                # If the Video Frame Rate is "not a number" ...
                if str(self.vidFrameRate) == 'nan':
                    # ... then a frame rate of 30 fps can be used.
                    tmpVidFrameRate = 30.0
                # Otherwise ...
                else:
                    # ... just use the frame rate extracted from the video
                    tmpVidFrameRate = self.vidFrameRate

                # If we're doing a video snapshot ...
                if self.snapshot:
                    # Set the Clip Duration to the frame rate times the number of frames divided by 1000.
                    # Hopefully, this will stop the DIVx Snapshot not stopping problem.  (It didn't.)
                    self.clipDuration = round(tmpVidFrameRate * numFramesForStill) / 1000.0
                    # ... then a frame rate of whatever the frame rate is and specifying the position of the desired frame is needed.
                    # We need to adjust the start time one FRAME earlier!  We need to adjust the end time  4 FRAMES later.  Otherwise,
                    # MPEG-1 video doesn't work every time!  I'm not sure why.  (This was determined experimentally.)
                    FFmpegCommand += ' "-r" "%0.2f" "-ss" "%0.5f" "-t" "%0.5f"' % (tmpVidFrameRate, (float(self.clipStart) - (1.5 * tmpVidFrameRate))/ 1000.0, self.clipDuration)

                # If we're NOT doing a snapshop, we need to get the proper frame rate to produce the correct pictures.
                elif self.stillFrameRate.GetStringSelection() == _("20 seconds"):
                    FFmpegCommand += ' "-r" "0.05"'
                elif self.stillFrameRate.GetStringSelection() == _("15 seconds"):
                    FFmpegCommand += ' "-r" "0.0666667"'
                elif self.stillFrameRate.GetStringSelection() == _("10 seconds"):
                    FFmpegCommand += ' "-r" "0.1"'
                elif self.stillFrameRate.GetStringSelection() == _("5 seconds"):
                    FFmpegCommand += ' "-r" "0.2"'
                elif self.stillFrameRate.GetStringSelection() == _("1 second"):
                    FFmpegCommand += ' "-r" "1"'

            # For CLIPS, add "-ss StartTime" and "-t Duration (seconds)"!!
            if (not self.ext in ['.jpg']) and (self.clipDuration > 0):
                FFmpegCommand += ' "-ss" "%0.5f" "-t" "%0.5f"' % (float(self.clipStart) / 1000.0, float(self.clipDuration) / 1000.0)

            # Add the "-y" parameter to over-write files, and append the destination file name placeholder
            FFmpegCommand += ' "-y" "%s"'

            # Create the prompt for the progress dialog
            prompt = unicode(_("Converting %s\n to %s"), 'utf8') % (self.txtSrcFileName.GetValue(), self.txtDestFileName.GetValue())
            # Create the Progress Dialog
            progressDlg = WaveformProgress.WaveformProgress(self, prompt, self.clipStart, self.clipDuration)

            if DEBUG:
                self.memo.AppendText("MediaConvert.OnConvert():  FFmpeg Command:")
                self.memo.AppendText(FFmpegCommand % (self.txtSrcFileName.GetValue(), self.txtDestFileName.GetValue()))
                self.memo.AppendText('\n\n')
            
            # Pass the Conversion Command we have created to the Progress Dialog
            progressDlg.SetProcessCommand(FFmpegCommand)
            # If we have a temporary file name because of the non-cp1252 FFmpeg issue...
            if self.tmpFileName != '':
                # ... use the modified file name as the input ...
                inputFile = self.fileName
                # ... and create the appropriate modified output file name
                outputFile = os.path.join(self.tmpPath, 'Output' + self.ext)
            # If we do NOT have a temporary file name ...
            else:
                # ... then we can use the input and output files currently showing on the form.
                inputFile = self.txtSrcFileName.GetValue()
                outputFile = self.txtDestFileName.GetValue()

            # Initiate the Conversion with the appropriate file names
            progressDlg.Extract(inputFile, outputFile, mode='CustomConvert')
            # Get the Error Log that may have been created
            errorLog = progressDlg.GetErrorMessages()
            # Destroy the Progess Dialog
            progressDlg.Destroy()
            
            # If the conversion was CANCELLED by the user ...
            if (len(errorLog) == 1) and (errorLog[0] == 'Cancelled'):
                msg = _("Conversion cancelled by user.") + "\n\n"
            # If the conversion was NOT cancelled ...
            else:
                # Inform the user that the conversion is complete
                msg = unicode(self.ext[1:].upper(), 'utf8') + unicode(_(" conversion completed."), 'utf8') + "\n\n"
                # If there are messages in the Error Log ...
                if len(errorLog) > 0:
                    # Create the message to the user
                    msg += unicode(_("Conversion Report:"), 'utf8') + "\n" + unicode(_("(These messages can be ignored unless problems arise.)"), 'utf8') + "\n\n"
                    # Add the Error Log contents to the user message
                    for line in errorLog:
                        msg += unicode(line, sys.getfilesystemencoding()) + "\n"
                # If we're taking a snapshot ...
                if self.snapshot:
                    # ... indicate that the snapshot was successful
                    self.snapshotSuccess = True
            # Display the user message
            self.memo.AppendText(msg)

            # When still images are created, FFmpeg seems to like to create extra images.  We need to clean that up here.
            # Start exception handling
            try:
                # If we're creating still images, have an Image #2, and DON'T have an Image #6, we can safely conclude that
                # a single still image was desired but more than one was created.
                if (self.ext == '.jpg') and \
                   os.path.exists(self.txtDestFileName.GetValue() % 2) and \
                   not os.path.exists(self.txtDestFileName.GetValue() % 6):
                    # Delete images 2 through 5, which are extraneous
                    for img in range(2, 6):
                        if os.path.exists(self.txtDestFileName.GetValue() % img):
                            # delete the image
                            os.remove(self.txtDestFileName.GetValue() % img)

                # If we have a temporary file name because of the non-cp1252 file name issue on Windows ...
                if self.tmpFileName != '':
                    # Determine the destination file name
                    destFile = self.txtDestFileName.GetValue()
                    # If we are doing a SnapShot (i.e. if the file has a NUMBER part) ...
                    if '%06d' in destFile:
                        # ... then substitute 1 in the number part of the destination file name ...
                        destFile = destFile % 1
                    # ... move the CONVERTED file, renaming it along the way
                    shutil.move(outputFile, destFile)

            # Handle exceptions ...
            except:

                if DEBUG:
                    print sys.exc_info()[0], sys.exc_info()[1]
                
                # ... by ignoring them!
                pass

            # On Windows, we need to reset the encoding to UTF-8
            if 'wxMSW' in wx.PlatformInfo:
                wx.SetDefaultPyEncoding('utf_8')

            # If we are embedded in Transana, not running stand-alone ...
            if __name__ != '__main__':
                # If we converted a raw media file (as opposed to exporting Clip Video)
                if (self.ext != '.jpg') and ((len(errorLog) != 1) or (errorLog[0] != 'Cancelled')) and (self.clipDuration == 0):
                    # ... prompt about updating media file references
                    updateDlg = Dialogs.QuestionDialog(self, _("Do you want to update all media file references in the database?"), noDefault=True)
                    # If the user wants to update all references ...
                    if updateDlg.LocalShowModal() == wx.ID_YES:
                        # ... separate paths from file names for both source and destination
                        (sourcePath, sourceFile) = os.path.split(self.txtSrcFileName.GetValue())
                        (destPath, destFile) = os.path.split(self.txtDestFileName.GetValue())
                        # If our file name is cp1252 compatible ...
                        if (self.tmpFileName == ''):
                            # ... then we need to process it like Database Data gets processed so the Queries will work!
                            sourceFile = DBInterface.ProcessDBDataForUTF8Encoding(sourceFile)
                        # Update the source file in the Database with the new File Path AND the new File Name
                        if not DBInterface.UpdateDBFilenames(self, destPath, [sourceFile], newName=destFile):
                            # Display an error message if the update failed
                            infodlg = Dialogs.InfoDialog(self, _('Update Failed.  Some records that would be affected may be locked by another user.'))
                            infodlg.ShowModal()
                            infodlg.Destroy()

            # If we are doing a snapshot ...
            if self.snapshot:
                # ... let's close the dialog automatically!
                self.Close()

    def OnClose(self, event):
        """ Close Button Press """
        # If we're on Windows ...
        if 'wxMSW' in wx.PlatformInfo:
            # If the temporary path exists ...
            if os.path.exists(self.tmpPath):
                # Start exception handling
                try:
                    # Get a list of files in the temporary directory
                    files = os.listdir(self.tmpPath)
                    # iterate through the files
                    for fil in files:
                        # If the file is called Input or Output ...
                        if ((len(fil) > 5) and (fil[:5] == 'Input')) or ((len(fil) > 6) and (fil[:6] == 'Output')):
                            # ... try to delete it
                            os.remove(os.path.join(self.tmpPath, fil))
                    # Remove the DIRECTORY
                    os.removedirs(self.tmpPath)
                # If an exception is raised ...
                except:

                    if DEBUG:
                        print sys.exc_info()[0]
                        print sys.exc_info()[1]
                        
                    # ... ignore it.  Transana might clean up after itself later
                    pass
        # Allow the form's Cancel event to fire to close the form
        event.Skip()
        
    def OnBrowse(self, event):
        """ Browse Button event handler (for both source and destination file names) """
        # If triggered by the Source File ...
        if event.GetId() == self.srcBrowse.GetId():

            if DEBUG:
                cwdbefore = os.getcwd()
                
            # Get Transana's File Filter definitions
            fileTypesString = _("All files (*.*)|*.*")
            # Create a File Open dialog.
            fs = wx.FileDialog(self, _('Select a media file to process:'),
                            self.lastPath,
                            "",
                            fileTypesString, 
                            wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
            # Select "All Media Files" as the initial Filter
            fs.SetFilterIndex(0)
            # Show the dialog and get user response.  If OK ...
            if fs.ShowModal() == wx.ID_OK:
                # ... get the selected file name
                self.fileName = fs.GetPath()
##            self.fileName = '/Users/davidwoods/Movies/Workshop Video/Leader/Demo/Demo.mpg'
            else:
                self.fileName = ''
            

            # Destroy the File Dialog
            fs.Destroy()

            self.txtSrcFileName.SetValue(self.fileName)

            if DEBUG:
                self.memo.AppendText('cwd BEFORE: %s  AFTER: %s\n' % (cwdbefore, os.getcwd()))

        # If triggered by the Destination File ...
        else:
            # Get the path and file name from the Destination File's current setting
            (path, filename) = os.path.split(self.txtDestFileName.GetValue())
            # Check the Conversion Type and limit the File Types to that Type
            if self.ext == '.mpg':
                fileTypesString = _("MPEG-1 files (*.mpg)|*.mpg")
            elif self.ext == '.mov':
                fileTypesString = _("MOV files (*.mov)|*.mov")
            elif self.ext == '.mp4':
                fileTypesString = _("MPEG-4 files (*.mp4)|*.mp4")
            elif self.ext == '.mp3':
                fileTypesString = _("MP3 audio files (*.mp3)|*.mp3")
            elif self.ext == '.wav':
                fileTypesString = _("WAV audio files (*.wav)|*.wav")
            else:
                fileTypesString = _("All files (*.*)|*.*")
            # Create a File Save dialog, using the path, file name, and file type determined above.
            fs = wx.FileDialog(self, _('Select a name for the output:'),
                            path,
                            filename,
                            fileTypesString, 
                            wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
            # Show the dialog and get user response.  If OK ...
            if fs.ShowModal() == wx.ID_OK:
                # ... place the selected file and path in the Destination File Name text control
                self.txtDestFileName.SetValue(fs.GetPath())

            # Destroy the File Dialog
            fs.Destroy()

    def OnSrcFileNameChange(self, event):
        """ Process any change in the Source File Name text control """
        # This is needed for when the user TYPES in the Source File Name control rather than using the Browse button

        # Clear the Memo
        self.memo.Clear()
        self.memo.Update()
        # If the current contents point to a file that exists ...
        if os.path.exists(self.txtSrcFileName.GetValue()):
            # ... set the selected file name
            self.fileName = self.txtSrcFileName.GetValue()
            # Process the selected file name to prepare for conversion
            self.ProcessMediaFile(self.txtSrcFileName.GetValue())
        # if the file does NOT exist ...
        else:
            prompt = unicode('File "%s" not found.  Try the Browse button.', 'utf8')
            self.memo.AppendText(prompt % self.txtSrcFileName.GetValue())
        # Call the default processor
        event.Skip()
        
    def OnFormat(self, event):
        """ Select Output Format OnChoice event handler """
        # Split the path from the file name
        (path, filename) = os.path.split(self.fileName)
        # Update the Output File Name to reflect the new Format
        self.SetOutputFilename(filename)
        # Enable or disable fields based on format type
        # Video formats first
        if self.ext in ['.mpg', '.mov', '.mp4']:
            # Enable video fields
            if self.videoSize.GetCount() > 1:
                self.videoSize.Enable(True)
            if self.videoBitrate.GetCount() > 1:
                self.videoBitrate.Enable(True)
            # Enable audio fields
            if self.audioBitrate.GetCount() > 1:
                self.audioBitrate.Enable(True)
            if self.audioSampleRate.GetCount() > 1:
                self.audioSampleRate.Enable(True)
            # Disable still image fields
            self.stillFrameRate.Enable(False)
        elif self.ext in ['.mp3', '.wav']:
            # Disable video fields
            self.videoSize.Enable(False)
            self.videoBitrate.Enable(False)
            # Enable audio fields
            if self.audioBitrate.GetCount() > 1:
                self.audioBitrate.Enable(True)
            if self.audioSampleRate.GetCount() > 1:
                self.audioSampleRate.Enable(True)
            # Disable still image fields
            self.stillFrameRate.Enable(False)
        elif self.ext in ['.jpg']:
            # Enable video size field
            if self.videoSize.GetCount() > 1:
                self.videoSize.Enable(True)
            # Disable video Bitrate field
            self.videoBitrate.Enable(False)
            # Disable audio fields
            self.audioBitrate.Enable(False)
            self.audioSampleRate.Enable(False)
            # Enable still image fields
            self.stillFrameRate.Enable(True)

    def SetOutputFilename(self, filename):
        """ Update the Output File Name """
        # If we have a temporary file name because of the FFmpeg non-cp1252 compatibility issue ...
        if self.tmpFileName != '':
            # ... then we want to use the ORIGINAL file name here, not the altered one!
            (path, filename) = os.path.split(self.tmpFileName)
        # Separate the File Name and the Extension
        (fn, ext) = os.path.splitext(filename)
        # If we have a Clip Name ...
        if self.clipName != '':
            # ... use that rather than the media file name
            fn = unicode("Clip", 'utf8') + "_" + self.clipName
        elif self.format.GetStringSelection()[:4] == 'JPEG':
            if self.snapshot:
                fn += "_" + unicode("Snapshot", 'utf8')
            fn += _('_%06d')
        # Otherwise ...
        else:
            # If we have a Left-To-Right language ...
            if TransanaGlobal.configData.LayoutDirection == wx.Layout_LeftToRight:
                # ... add the TRANSLATED word "Analysis" on to indicate that this is the low-res Analysis version of the media file
                fn += unicode(_('-Analysis'), 'utf8')
            # With Right-To-Left languages, we can't use the Translated version of the word "Analysis"!!
            else:
                fn = fn + unicode('-Analysis', 'utf8')
        # Based on the Format Selection (start of the text), determine the appropriate file extension.
        # Remember it as a proxy for destination file type for later processing
        if 'MPEG-1' in self.format.GetStringSelection():
            self.ext = '.mpg'
        elif 'MOV' in self.format.GetStringSelection():
            self.ext = '.mov'
        elif 'MPEG-4' in self.format.GetStringSelection():
            self.ext = '.mp4'
        elif 'MP3' in self.format.GetStringSelection():
            self.ext = '.mp3'
        elif 'WAV' in self.format.GetStringSelection():
            self.ext = '.wav'
        elif 'JPEG' in self.format.GetStringSelection():
            self.ext = '.jpg'
        
        # Build a new file name, starting with the last path used
        newFilename = self.lastPath
        # If we have a Left-To-Right language ...
        if TransanaGlobal.configData.LayoutDirection == wx.Layout_LeftToRight:
            # If that path does NOT end with a file separator ...
            if newFilename[-1] != os.sep:
                # ... add one
                newFilename += os.sep
            # Add the File Name, the "Analysis" tag which flags files that have been converted for Analysis, and the proper extension
            newFilename += fn + self.ext
        # With Right-To-Left languages, we have to build file names BACKWARDS!
        else:
            # If that path does NOT end with a file separator ...
            if newFilename[-1] != os.sep:
                # ... add one
                newFilename = newFilename + os.sep
            # Add the File Name, the "Analysis" tag which flags files that have been converted for Analysis, and the proper extension
            newFilename = newFilename + fn + self.ext

        # Update the Destination File Name text control with the new file name
        self.txtDestFileName.SetValue(newFilename)
        # Enable the Destination File Name and Browse controls
        self.txtDestFileName.Enable(True)
        self.destBrowse.Enable(True)

    def OnHelp(self, event):
        """ Help Button event handler """
        # If a MenuWindow is defined (which is should always be!)
        if TransanaGlobal.menuWindow != None:
            # ... then use the MenuWindow's ControlObject to call the Help infrastructure
            TransanaGlobal.menuWindow.ControlObject.Help('Media File Conversion')
        

# For testing purposes, this module can run stand-alone.
if __name__ == '__main__':
    
    # Create a simple app for testing.
    app = wx.PySimpleApp()
    
    # Create the form, no parent needed
    frame = MediaConvert(None) # , u'E:\\Vidëo\\Demo\\Demo.mpg')
    # Show the Dialog Box and process the result.
    frame.ShowModal()

    # Destroy the dialog box.
    frame.Destroy()
    # Call the app's MainLoop()
    app.MainLoop()
