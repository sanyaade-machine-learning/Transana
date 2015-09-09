# Copyright (C) 2004 - 2015  The Board of Regents of the University of Wisconsin System 
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

"""This module implements the sFTP File Transfer logic for the File Management window for Transana.  It can 
   also be run as a stand-alone utility.  More specifically, this module presents a File Transfer Dialog
   and manages all the sFTP Connection logic behind moving files between the sFTP Server and a local computer.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import wx  # import wxPython
import os
import string
import sys
import ctypes
# import Python exceptions
import exceptions
import time
import Dialogs
import FileManagement
import TransanaGlobal

# Define Transfer Direction constants
sFTP_UPLOAD   = wx.NewId()
sFTP_DOWNLOAD = wx.NewId()



class sFTPFileTransfer(wx.Dialog):
    """ This object transfers files between the sFTP Server and the local file system and displays transfer progress. """
    def __init__(self, parent, title, fileName, localDir, remoteDir, direction):
        """ Set up the Dialog Box and all GUI Widgets. """
        # Initialize the last update time
        self.lastUpdate = time.time()
        
        # Set up local variables
        self.parent = parent

        size = (350,200)
        # Create the Dialog Box itself, with no minimize/maximize/close buttons
        wx.Dialog.__init__(self, parent, -1, title, size = size, style=wx.CAPTION)

        # Create a main VERTICAL sizer for the form
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # File label
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("File: %s"), 'utf8')
        else:
            prompt = _("File: %s")
        self.lblFile = wx.StaticText(self, -1, prompt % fileName, style=wx.ST_NO_AUTORESIZE)
        # Add the label to the Main Sizer
        mainSizer.Add(self.lblFile, 0, wx.ALL, 10)

        # Progress Bar
        self.progressBar = wx.Gauge(self, -1, 100, style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        # Add the element to the Main Sizer
        mainSizer.Add(self.progressBar, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Create a Row Sizer
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        # Bytes Transferred label
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("%d bytes of %d transferred"), 'utf8')
        else:
            prompt = _("%d bytes of %d transferred")
        self.lblBytes = wx.StaticText(self, -1, prompt % (100000000, 100000000), style=wx.ST_NO_AUTORESIZE)
        # Add the label to the Row Sizer
        r1Sizer.Add(self.lblBytes, 5, wx.EXPAND)

        # Percent Transferred label
        self.lblPercent = wx.StaticText(self, -1, "%5.1d %%" % 1000.1, style=wx.ST_NO_AUTORESIZE | wx.ALIGN_RIGHT)
        # Add the Element to the Row Sizer
        r1Sizer.Add(self.lblPercent, 1, wx.ALIGN_RIGHT)

        # Add the Row Sizer to the Main Sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Create a Row Sizer
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Elapsed Time label
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("Elapsed Time: %d:%02d:%02d"), 'utf8')
        else:
            prompt = _("Elapsed Time: %d:%02d:%02d")
        self.lblElapsedTime = wx.StaticText(self, -1, prompt % (0, 0, 0), style=wx.ST_NO_AUTORESIZE)
        # Add the element to the Row Sizer
        r2Sizer.Add(self.lblElapsedTime, 0)

        # Add a spacer
        r2Sizer.Add((1, 0), 1, wx.EXPAND)

        # Remaining Time label
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("Time Remaining: %d:%02d:%02d"), 'utf8')
        else:
            prompt = _("Time Remaining: %d:%02d:%02d")
        self.lblTimeRemaining = wx.StaticText(self, -1, prompt % (0, 0, 0), style=wx.ST_NO_AUTORESIZE | wx.ALIGN_RIGHT)
        # Add the element to the Row Sizer
        r2Sizer.Add(self.lblTimeRemaining, 0)

        # Add the Row Sizer to the Main Sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Transfer Speed label
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("Transfer Speed: %d k/sec"), 'utf8')
        else:
            prompt = _("Transfer Speed: %d k/sec")
        self.lblTransferSpeed = wx.StaticText(self, -1, prompt % 0, style=wx.ST_NO_AUTORESIZE)
        # Add the element to the Main Sizer
        mainSizer.Add(self.lblTransferSpeed, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Cancel Button
        self.btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel Remaining Files"))
        # Add the element to the Main Sizer
        mainSizer.Add(self.btnCancel, 0, wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        wx.EVT_BUTTON(self, wx.ID_CANCEL, self.OnCancel)

        # Attach the main sizer to the form
        self.SetSizer(mainSizer)
        # Turn Auto Layout on
        self.SetAutoLayout(True)
        # Lay out the form
        self.Layout()
        # Center on the Screen
        TransanaGlobal.CenterOnPrimary(self)
        # Show the form
        self.Show()
        # make sure the loca directory ends with the proper path seperator character
        if localDir[-1] != os.sep:
            localDir = localDir + os.sep
        # Initialize variables used in file transfer
        BytesRead = 0
        # "cancelled" is intialized to false.  If the user cancels the file transfer,
        # this variable gets set to true to signal the need to interrupt the transfer.
        self.cancelled = False
        # Note the starting time of the transfer for progress reporting purposes
        self.StartTime = time.time()

        # If we are sending files from the sFTP Server to the local File System...
        if direction == sFTP_DOWNLOAD:
            self.SetTitle(_('Downloading . . .'))

            try:
                FileResult = parent.sFTPClient.get(string.join([remoteDir, fileName], '/'), os.path.join(localDir, fileName), callback = self.UpdateDisplay)
            except IOError:
                prompt = unicode(_("Input / Output Error"), 'utf8') + "\n%s\n" + unicode(_('Please press "Refresh".'), 'utf8')
                dlg = Dialogs.ErrorDialog(self, prompt % (sys.exc_info()[1]))
                dlg.ShowModal()
                dlg.Destroy()

        # If we are sending data from the local file system to the sFTP Server ...
        elif direction == sFTP_UPLOAD:
            self.SetTitle(_('Uploading . . .'))

            try:
                # Detect Strings (instead of Unicode objects) and decode them if needed.
                if isinstance(localDir, str):
                    localDir = localDir.decode(TransanaGlobal.encoding)
                if isinstance(remoteDir, str):
                    remoteDir = remoteDir.decode(TransanaGlobal.encoding)
                # Put the file and capture the result
                FileResult = parent.sFTPClient.put(os.path.join(localDir, fileName), string.join([remoteDir, fileName], '/'), callback = self.UpdateDisplay)
            except IOError:
                prompt = unicode(_("Input / Output Error"), 'utf8') + "\n%s\n" + unicode(_('Please press "Refresh".'), 'utf8')
                dlg = Dialogs.ErrorDialog(self, prompt % (sys.exc_info()[1]))
                dlg.ShowModal()
                dlg.Destroy()

    def TransferSuccessful(self):
        # This class needs to return whether the transfer succeeded so that the delete portion of
        # a cancelled Move can be skipped.
        return not self.cancelled

    def HoursMinutesSeconds(self, time):
        if time > 3600:
            hours = int(time / (3600))
        else:
            hours = 0
        time = time - (hours * 3600)
        if time > 60:
            mins = int(time / 60)
        else:
            mins = 0
        secs = time % 60
        return (hours, mins, secs)

    def UpdateDisplay(self, bytesTransferred, totalBytes):
        """ Update the Transfer Dialog Box labels and progress bar """

        # If the dialog has updated within the half last second ...
        if (time.time() - self.lastUpdate)  < 0.5:
            
            # ... then skip this routine.  Let's not add unnecessary inefficiency to the transfer.
            return

        else:
            self.lastUpdate = time.time()

        # Display Number of Bytes tranferred
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("%d bytes of %d transferred"), 'utf8')
        else:
            prompt = _("%d bytes of %d transferred")
        self.lblBytes.SetLabel(prompt % (bytesTransferred, totalBytes))
        # Avoiding division by zero problems ...
        if bytesTransferred > 0:
            # ... display % of total bytes transferred
            self.lblPercent.SetLabel("%5.1f %%" % (float(bytesTransferred) / float(totalBytes) * 100))
            # Update the Progress Bar
            self.progressBar.SetValue(int(float(bytesTransferred) / float(totalBytes) * 100))
        # Calculate Elapsed Time and display it
        elapsedTime = time.time() - self.StartTime

        (hours, mins, secs) = self.HoursMinutesSeconds(elapsedTime)
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("Elapsed Time: %d:%02d:%02d"), 'utf8')
        else:
            prompt = _("Elapsed Time: %d:%02d:%02d")
        self.lblElapsedTime.SetLabel(prompt % (hours, mins, secs))
        if elapsedTime > 0:
            # Calculate the Transfer Speed.  (Dividing by 1024 gives k/sec rather than bytes/sec).
            # (kilobytes transferred divided by elapsed time = rate of transfer in k/sec)
            speed = (bytesTransferred / 1024.0) / elapsedTime
            # Estimate the amount of time it will take to transfer the remaining data.
            # (kilobytes of data remaining divided by the transfer speed = time remaining for transfer in seconds)
            timeRemaining = (((totalBytes - bytesTransferred) / 1024.0) / (speed))
            # Display the results of these calculations

            (hours, mins, secs) = self.HoursMinutesSeconds(timeRemaining)
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Time Remaining: %d:%02d:%02d"), 'utf8')
            else:
                prompt = _("Time Remaining: %d:%02d:%02d")
            self.lblTimeRemaining.SetLabel(prompt % (hours, mins, secs))
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Transfer Speed: %6.1f k/sec"), 'utf8')
            else:
                prompt = _("Transfer Speed: %6.1f k/sec")
            self.lblTransferSpeed.SetLabel(prompt % speed)
        # Allow the screen to update and accept User Input (Cancel Button press, for example).
        # (wxYield allows Windows to process Windows Messages)
        wx.Yield()
        
    def OnCancel(self, event):
        """ Respond to the user pressing the "Cancel" button to interrupt the file transfer """
        # Cancel is accomplished by setting this local variable.  The File Transfer logic detects this and responds appropriately
        self.cancelled = True
        self.btnCancel.SetLabel(_('Cancel requested.'))
        self.btnCancel.Enable(False)
