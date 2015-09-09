# -*- coding: cp1252 -*-
# Copyright (C) 2003 - 2015 The Board of Regents of the University of Wisconsin System
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

"""This module implements the Progress Dialog for Waveform Creation."""

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "WaveformProgres DEBUG is ON!!!!!"

import wx      # import wxPython
import os, sys
# import Python's time module
import time

if __name__ == '__main__':
    # This module expects i18n.  Enable it here.
    __builtins__._ = wx.GetTranslation

# Import Transana's Miscellaneous functions
import Misc
# Import Transana's Global Variables
import TransanaGlobal

ID_BTNCANCEL    =  wx.NewId()


class WaveformProgress(wx.Dialog):
    """ This class implements the Progress Dialog for Waveform Creation. 
        This dialog works two ways.

        The old way is that it is used to display progress from an EXTERNAL process.  It needs to be non-Modal.
        To use it this way, you create it, call Show(), periodically update it using the Update() method, and then Destroy() it.

        The new way is that it is used to actually handle the audio extraction.  In this case it is modal.
        To use it this way, you create it, then call the Extract() method, which will show it modally and handle updating itself.
        Then Destroy() it. """

    def __init__(self, parent, label='', clipStart=0, clipDuration=0, showModally=True):
        """ Initialize the Progress Dialog """

        # There's a bug.  I think it's an interaction between OS X 10.7.5 and earlier (but not 10.8.4), wxPython (version
        # unknown, but I've seen it in 2.8.12.1, 2.9.4.0, and a pre-release build of 2.9.5.0), and the build process.
        # Basically, if you open a wxFileDialog object, something bad happens with wxLocale such that subsequent calls
        # to wxProcess don't encode unicode file names correctly, so FFMpeg fails.  Only OS X 10.7.5, only when Transana
        # has been built (not running through the interpreter), only following the opening of a wxFileDialog, and only
        # with filenames with accented characters or characters from non-English languages.

        # The resolution to this bug is to reset the Locale here, based on Transana's current wxLocale setting in the
        # menuWindow object.
        self.locale = wx.Locale(TransanaGlobal.menuWindow.locale.Language)

        # Remember the Parent
        self.parent = parent
        # Remember whether we're MODAL or ALLOWING MULTIPLE THREADS
        self.showModally = showModally
        # Remember the start time and duration, if they are passed in.
        self.clipStart = clipStart
        self.clipDuration = clipDuration

        # Define the process variable
        self.process = None
        # Initialize a list to collect error messages
        self.errorMessages = []

        # Encode the prompt
        prompt = unicode(_('Media File Conversion Progress'), 'utf8')
        # Define the Dialog Box.  wx.STAY_ON_TOP required because this dialog can't be modal, but shouldn't be hidden or worked behind.
        wx.Dialog.__init__(self, parent, -1, prompt, size=(400, 160), style=wx.CAPTION | wx.STAY_ON_TOP)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the main Sizer, which is Vertical
        sizer = wx.BoxSizer(wx.VERTICAL)
        # Create a horizontal sizer for the text below the progress bar.
        timeSizer = wx.BoxSizer(wx.HORIZONTAL)

        # For now, place an empty StaticText label on the Dialog
        self.lbl = wx.StaticText(self, -1, label)
        sizer.Add(self.lbl, 0, wx.ALIGN_LEFT | wx.ALL, 10)

        # Progress Bar
        self.progressBar = wx.Gauge(self, -1, 100, style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        sizer.Add(self.progressBar, 0, wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT | wx.EXPAND, 10)

        # Seconds Processed label
        # Encode the prompt
        prompt = unicode(_("%d:%02d:%02d of %d:%02d:%02d processed"), 'utf8')
        self.lblTime = wx.StaticText(self, -1, prompt % (0, 0, 0, 0, 0, 0), style=wx.ST_NO_AUTORESIZE)
        timeSizer.Add(self.lblTime, 0, wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT | wx.TOP, 10)
        timeSizer.Add((0, 5), 1, wx.EXPAND)
        
        # Percent Processed label
        self.lblPercent = wx.StaticText(self, -1, "%3d %%" % 0, style=wx.ST_NO_AUTORESIZE | wx.ALIGN_RIGHT)
        timeSizer.Add(self.lblPercent, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.TOP, 10)
        sizer.Add(timeSizer, 1, wx.EXPAND)

        elapsedSizer = wx.BoxSizer(wx.HORIZONTAL)

        # Time Elapsed label
        # Encode the prompt
        prompt = unicode(_("%s elapsed"), 'utf8')
        self.lblElapsed = wx.StaticText(self, -1, prompt % '0:00:00', style=wx.ST_NO_AUTORESIZE | wx.ALIGN_LEFT)
        elapsedSizer.Add(self.lblElapsed, 0, wx.ALIGN_LEFT | wx.LEFT, 10)
        elapsedSizer.Add((0, 5), 1, wx.EXPAND)

        # Time Remaining label
        # Encode the prompt
        prompt = unicode(_("%s remaining"), 'utf8')
        self.lblRemaining = wx.StaticText(self, -1, prompt % '0:00:00', style=wx.ST_NO_AUTORESIZE | wx.ALIGN_RIGHT)
        elapsedSizer.Add(self.lblRemaining, 0, wx.ALIGN_RIGHT | wx.RIGHT, 10)
        sizer.Add(elapsedSizer, 1, wx.EXPAND | wx.BOTTOM, 4)

        # Cancel button
        # Encode the prompt
        prompt = unicode(_("Cancel"), 'utf8')
        self.btnCancel = wx.Button(self, -1, prompt)
        sizer.Add(self.btnCancel, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        self.btnCancel.Bind(wx.EVT_BUTTON, self.OnInterrupt)

        self.SetSizer(sizer)
        # Set this as the minimum size for the form.
        sizer.SetMinSize(wx.Size(350, 160))
        # Call Layout to "place" the widgits
        self.Layout()
        self.SetAutoLayout(True)
        self.Fit()

        # Create a Timer that will check for progress feedback
        self.timer = wx.Timer()
        self.timer.Bind(wx.EVT_TIMER, self.OnTimer)

        TransanaGlobal.CenterOnPrimary(self)

    def OnInterrupt(self, event):
        """ Cancel Button Event Handler """
        # Disable the Cancel button to prevent multiple presses while processing occurs
        self.btnCancel.Enable(False)
        # If the process exists ...
        if self.process is not None:
            # ... kill the process
            result = self.process.Kill(self.pid, wx.SIGKILL, wx.KILL_NOCHILDREN)
            # If the kill was successful ...
            if result == 0:
                # ... signal the calling routine through the Error Message process
                self.errorMessages=['Cancelled']
                # Delete the Destination File, if it exists
                if os.path.exists(self.destFile):
                    os.remove(self.destFile)
            
    def Update(self, percent, seconds, total=0):
        """ This method allows the contents of this form to be "updated" in the Wave Extraction Callback process """

        # Update the Progress Bar
        self.progressBar.SetValue(percent)
        # Calculate the time values to be displayed
        timeProcessed = seconds
        secondsProcessed = (timeProcessed % 60)
        timeProcessed = timeProcessed - secondsProcessed
        hoursProcessed = (timeProcessed / (60 * 60))
        timeProcessed = timeProcessed - (hoursProcessed * 60 * 60)
        minutesProcessed = (timeProcessed / 60)
        # Update the amount of time processed
        if total == 0:
            self.lblTime.SetLabel(_("%d:%02d:%02d processed") % (hoursProcessed, minutesProcessed, secondsProcessed))
        else:
            # Calculate the time values to be displayed
            totTimeProcessed = total
            totSecondsProcessed = (totTimeProcessed % 60)
            totTimeProcessed = totTimeProcessed - totSecondsProcessed
            totHoursProcessed = (totTimeProcessed / (60 * 60))
            totTimeProcessed = totTimeProcessed - (totHoursProcessed * 60 * 60)
            totMinutesProcessed = (totTimeProcessed / 60)
            # Include total time in the label if it is known
            self.lblTime.SetLabel(_("%d:%02d:%02d of %d:%02d:%02d processed") % (hoursProcessed, minutesProcessed, secondsProcessed,
                                                                                 totHoursProcessed, totMinutesProcessed, totSecondsProcessed))
            
        # Display % processed
        self.lblPercent.SetLabel("%d %%" % percent)

        # If we've made SOME progress ...
        if percent > 0:
            # ... calculate time elapsed ...
            t1 = (time.time() - self.progressStartTime)
            # ... calculate total estimated time for completion
            t2 = t1 * 100.0 / percent
            # Display elapsed time
            self.lblElapsed.SetLabel(_("%s elapsed") % Misc.TimeMsToStr(t1 * 1000))
            # Display time remaining
            self.lblRemaining.SetLabel(_("%s remaining") % Misc.TimeMsToStr((t2 - t1) * 1000))

        # wxYield tells the OS to process all messages in the queue, such as GUI updates, so the label change will show up.
        # But under some circumstances, it produces an error saying it's been called recursively.  We'll have to trap that.
        try:
            wx.Yield()
        except:
            pass

        # If we have processed more video than we should have (probably Snapshot not ending properly!) ...
        if (self.clipDuration > 0) and (seconds > self.clipDuration):
            # ... then if the process exists ...
            if self.process is not None:
                # ... kill the process
                result = self.process.Kill(self.pid, wx.SIGKILL, wx.KILL_NOCHILDREN)
                # NOTE:  This is probably a success, so DON'T signal failure here!

        return 1

    def SetProcessCommand(self, processCommand):
        """ When using FFmpeg, we need a way to pass a complex Conversion Process Command. """
        # Set the Command used for the extraction / conversion process
        self.processCommand = processCommand

    def GetErrorMessages(self):
        """ When using FFmpeg, we need a way to return error message to the calling routine. """
        # Return the error messages list
        return self.errorMessages

    def Extract(self, inputFile, outputFile, mode='AudioExtraction'):
        """ Perform Audio Extraction, or Media Conversion with the new version. """
        # Remember the mode being used
        self.mode = mode

        # If we're messing with wxProcess, we need to define a function to clean up if we shut down!
        def __del__(self):
            if self.process is not None:
                self.process.Detach()
                self.process.CloseOutput()
                self.process = None

        # Be prepared to capture the wxProcess' EVT_END_PROCESS
        self.Bind(wx.EVT_END_PROCESS, self.OnEndProcess)
        
        # Windows requires that we change the default encoding for Python for the audio extraction code to work
        # properly with Unicode files (!!!)  This isn't needed on OS X, as its default file system encoding is utf-8.
        # See python's documentation for sys.getfilesystemencoding()
        if 'wxMSW' in wx.PlatformInfo:
            wx.SetDefaultPyEncoding('mbcs')
            
        # Build the command line for the appropriate media conversion call
        if mode == 'AudioExtraction':
            # -i              input file
            # -vn             disable video
            # -ar 2756        Audio Sampling rate 2756 Hz
            # -ab 8k          Audio Bitrate 8kb/s
            # -ac 1           Audio Channels 1 (mono)
            # -acodec pcm_u8  8-bit PCM audio codec
            # -y              overwrite existing destination file
            programStr = os.path.join(TransanaGlobal.programDir, 'ffmpeg_Transana')
            if 'wxMSW' in wx.PlatformInfo:
                programStr += '.exe'
            process = '"' +  programStr + '" "-embedded" "1" "-i" "%s" "-vn" "-ar" "2756" "-ab" "8k" "-ac" "1" "-acodec" "pcm_u8" "-y" "%s"'
            tempMediaFilename = inputFile
            tempWaveFilename = outputFile

        elif mode == 'AudioExtraction-OLD':
            programStr = os.path.join(TransanaGlobal.programDir, 'audioextract')
            if 'wxMSW' in wx.PlatformInfo:
                programStr += '.exe'
            process = '"' + programStr + '" "%s" "%s"'
            tempMediaFilename = inputFile.encode('utf8')
            tempWaveFilename = outputFile.encode('utf8')
        elif mode == 'CustomConvert':
            process = self.processCommand

            if DEBUG:
                print "WaveformProgress.Extract():", sys.getfilesystemencoding()
                print process % (inputFile, outputFile)
                print
            
            tempMediaFilename = inputFile.encode(sys.getfilesystemencoding())
            tempWaveFilename = outputFile.encode(sys.getfilesystemencoding())
        else:
            import TransanaExceptions
            raise TransanaExceptions.NotImplementedError

        self.destFile = outputFile
        
        # Create a wxProcess object
        self.process = wx.Process(self)
        # Call the wxProcess Object's Redirect method.  This allows us to capture the process's output!
        self.process.Redirect()
        # Encode the filenames to UTF8 so that unicode files are handled properly
        process = process.encode('utf8')

        if DEBUG:
            print "WaveformProgress.Extract():"
            st = process % (tempMediaFilename, tempWaveFilename)
            if isinstance(st, unicode):
                print st.encode('utf8')
            else:
                print st
            print

        # Call the Audio Extraction program using wxExecute, capturing the output via wxProcess.  This call MUST be asynchronous. 
        self.pid = wx.Execute(process % (tempMediaFilename, tempWaveFilename), wx.EXEC_ASYNC, self.process)

        # On Windows ...
        if 'wxMSW' in wx.PlatformInfo:
            # ... reset the default Python encoding to UTF-8
            wx.SetDefaultPyEncoding('utf_8')

        # Note the time when the progress bar started
        self.progressStartTime = time.time()
        # Start the timer to get feedback and post progress
        self.timer.Start(500)

        # If we're processing Modally ...
        if self.showModally:
            # ... show the Progress Dialog modally
            self.ShowModal()
        # If we're allowing multiple threads ...
        else:
            # ... show the Progress Dialog non-modally
            self.Show()

    def OnEndProcess(self, event):
        """ End of wx.Process event handler """
        # Stop the Progress Timer
        self.timer.Stop()
        # If the process exists ...
        if self.process is not None:
            # Get the Process Error Stream
            errStream = self.process.GetErrorStream()
            # If the stream exists and can be read ...
            if errStream and errStream.CanRead():
                # ... read the stream
                text = errStream.read()
                # If the newlines are in the form of \r\n, we need to replace them with \n only for Python.
                text = text.replace('\r\n', '\n')
                # Split the stream into individual lines
                text = text.split('\n')
                # For each line ...
                for line in text:
                    # ... if the line is not blank, and is not is not a buffer line ...
                    if (line != ''):  # and (not '[buffer @' in line):
                        # ... add the line to the error messages list
                        self.errorMessages.append(line)
            # Destroy the now-completed process
            if self.process is not None:
                self.process.Detach()
                self.process.CloseOutput()
            # De-reference the process
            self.process = None
            wx.YieldIfNeeded()
            # If we're allowing multiple threads ...
            if not self.showModally:
                # ... inform the PARENT that this thread is complete for cleanup
                self.parent.OnConvertComplete(self)
            # Close the Progress Dialog
            self.Close()

    def OnTimer(self, event):
        """ Handle the EVT_TIMER event, which updates the progress dialog """
        # If the process exists ...
        if self.process is not None:
            # Get the process input stream
            stream = self.process.GetInputStream()
            # If the stream can be read ...
            if stream and stream.CanRead():
                # Read the stream
                text = stream.read()
                # If the newlines are in the form of \r\n, we need to replace them with \n only for Python.
                text = text.replace('\r\n', '\n')
                # Split the stream into individual lines
                text = text.split('\n')
                # Lines by this point are in pairs, with one containing data followed by an empty one, ['xp ## ## ## etc', ''].
                # So we just want to look at the second to last line!
                if len(text) >= 2:
                    # Split this line into the component elements, which are separated by spaces
                    progress = text[-2].split(' ')
                # If we only have one line, something is wrong.  Maybe we're done?
                else:
                    # Just send "100%" to the progress dialog
                    progress = ['xp', 100, 0]
                    # and close the progress dialog.
                    self.Close()

                # If we're using the OLD audio extraction system ...
                if self.mode == 'AudioExtraction-OLD':
                    # ... supplement the data provided with the "xp" header and add a 0 on the end to signal that total time is unknown
                    progress = ['xp'] + progress + ['0']
                # If the first term is "xp", we have an Extraction Progress line.
                if progress[0] == 'xp':

                    # If we are only exporting PART of a media file ...
                    if self.clipDuration > 0:
                        # ... then we need to ALTER the values from FFmpeg to reflect the TRUE progress
                        seconds = float(progress[2])
                        # We don't want total media file time, but total CLIP time
                        total = float(self.clipDuration) / 1000.0
                        # Percent should be based on CLIP time, not total media file time
                        percent = seconds / total * 100.0
                        # Rebuild the Progress data structure to reflect these changes
                        progress = [progress[0], "%0.2f" % percent, progress[2], "%0.2f" % total]

                    # If so, report the three data elements to the Update method to be displayed on the Progress Dialog
                    result = self.Update(long(float(progress[1])), long(float(progress[2])), int(float(progress[3])))
                    
                # Otherwise ...
#                else:
                    # ... just do output so I'll notice during testing and handle it!
#                    print ' WaveformProgress.OnTimer() --> ', progress

# If running in stand-alone mode for testing ...
if __name__ == '__main__':
    # Create a PySimpleApp
    app = wx.PySimpleApp()
    # Indicate what file to process
    srcfilename = u'E:\\Vidëo\\Demo\\BBC News.mp4'
   # Create the Progress Dialog
    frame = WaveformProgress(None, "Converting\n  %s " % (srcfilename))
    # If the source file exists ...
    if os.path.exists(srcfilename):
        # Specify the Destination file
        destfilename = u'D:\\My Documents\\Python\\ffmpeg\\test.wav'
        # Perform the extaction / conversion to be tested
        frame.Extract(srcfilename, destfilename, mode='AudioExtraction')
    # If the file does not exist ...
    else:
        # ... report the failure
        print "File does not exist"
    # Destroy the Progress Dialog
    frame.Destroy()
    # Run the Application Main Loop
    app.MainLoop()
    
    
