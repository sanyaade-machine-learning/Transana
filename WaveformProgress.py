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

"""This module implements the Progress Dialog for Waveform Creation."""

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

import wx      # import wxPython
import os, sys
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

    def __init__(self, parent, waveFilename, label=''):
        """ Initialize the Progress Dialog """
        # Remember the filename
        self.waveFilename = waveFilename

        # Define the process variable
        self.process = None
        
        # Define the Dialog Box.  wx.STAY_ON_TOP required because this dialog can't be modal, but shouldn't be hidden or worked behind.
        wx.Dialog.__init__(self, parent, -1, _('Wave Extraction Progress'), size=(400, 160), style=wx.CAPTION | wx.STAY_ON_TOP)

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
        self.lblTime = wx.StaticText(self, -1, _("%d:%02d:%02d processed") % (0, 0, 0), style=wx.ST_NO_AUTORESIZE)
        timeSizer.Add(self.lblTime, 0, wx.ALIGN_LEFT | wx.ALL, 10)
        timeSizer.Add((0, 5), 1, wx.EXPAND)
        

        # Percent Processed label
        self.lblPercent = wx.StaticText(self, -1, "%3d %%" % 0, style=wx.ST_NO_AUTORESIZE | wx.ALIGN_RIGHT)
        timeSizer.Add(self.lblPercent, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        sizer.Add(timeSizer, 1, wx.EXPAND)

        self.SetSizer(sizer)
        self.Fit()
        # Set this as the minimum size for the form.
        self.SetSizeHints(minW = self.GetSize()[0], minH = self.GetSize()[1])

        # Call Layout to "place" the widgits
        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()
        

    def Update(self, percent, seconds):
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
        self.lblTime.SetLabel(_("%d:%02d:%02d processed") % (hoursProcessed, minutesProcessed, secondsProcessed))
        # Display % processed
        self.lblPercent.SetLabel("%d %%" % percent)
        # wxYield tells the OS to process all messages in the queue, such as GUI updates, so the label change will show up.
        # But under some circumstances, it produces an error saying it's been called recursively.  We'll have to trap that.
        try:
            wx.Yield()
        except:
            pass
        return 1

    def Extract(self, inputFile, outputFile):

        # If we're messing with wxProcess, we need to define a function to clean up if we shut down!
        def __del__(self):
            if self.process is not None:
                self.process.Detach()
                self.process.CloseOutput()
                self.process = None

        # Be prepared to capture the wxProcess' EVT_END_PROCESS
        self.Bind(wx.EVT_END_PROCESS, self.OnEndProcess)
        # We use the Idle event to capture wxProcess feedback
        wx.EVT_IDLE(self, self.OnIdle)
        
        # Windows requires that we change the default encoding for Python for the audio extraction code to work
        # properly with Unicode files (!!!)  This isn't needed on OS X, as its default file system encoding is utf-8.
        # See python's documentation for sys.getfilesystemencoding()
        if 'wxMSW' in wx.PlatformInfo:
            wx.SetDefaultPyEncoding('mbcs')
            
        # Define the command line for calling the Audio Extract code, which is platform-dependent
        if 'wxMSW' in wx.PlatformInfo:
            process = '"' + TransanaGlobal.programDir + '\\audioextract.exe" "%s" "%s"'
        elif 'wxMac' in wx.PlatformInfo:
            process = '"' + TransanaGlobal.programDir + '/audioextract" "%s" "%s"'
        else:
            import TransanaExceptions
            raise TransanaExceptions.NotImplementedError
        # Create a wxProcess object
        self.process = wx.Process(self)
        # Call the wxProcess Object's Redirect method.  This allows us to capture the process's output!
        self.process.Redirect()
        # Encode the filenames to UTF8 so that unicode files are handled properly
        tempMediaFilename = inputFile.encode('utf8')
        tempWaveFilename = outputFile.encode('utf8')
        # Call the Audio Extraction program using wxExecute, capturing the output via wxProcess.  This call MUST be asynchronous. 
        self.pid = wx.Execute(process % (tempMediaFilename, tempWaveFilename), wx.EXEC_ASYNC, self.process)

        if 'wxMSW' in wx.PlatformInfo:
            wx.SetDefaultPyEncoding('utf_8')

        self.ShowModal()

    def OnIdle(self, event):
        if self.process is not None:
            stream = self.process.GetInputStream()
            if stream and stream.CanRead():
                text = stream.read()
                # If the newlines are in the form of \r\n, we need to replace them with \n only for Python.
                text = text.replace('\r\n', '\n')
                text = text.split('\n')
                if len(text) >= 2:
                    progress = text[-2].split(' ')
                else:
                    progress = [100, 0]
                    self.Close()
                result = self.Update(int(progress[0]), int(float(progress[1])))

    def OnEndProcess(self, event):
        if self.process is not None:
            stream = self.process.GetInputStream()
            if stream.CanRead():
                text = stream.read()
                text = text.split('\r\n')
                
            self.process.Destroy()
            self.process = None
            self.Close()
