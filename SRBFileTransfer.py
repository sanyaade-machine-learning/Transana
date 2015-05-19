# Copyright (C) 2004 - 2013  The Board of Regents of the University of Wisconsin System 
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

"""This module implements the SRB File Transfer logic for the File Management window for Transana.  It can 
   also be run as a stand-alone utility.  More specifically, this module presents a File Transfer Dialog
   and manages all the SRB Connection logic behind moving files between the SRB and a local computer.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import wx  # import wxPython
import os
import sys
import ctypes
# import Python exceptions
import exceptions
import time
import FileManagement
import TransanaGlobal

# Define Transfer Direction constants
srb_UPLOAD   = wx.NewId()
srb_DOWNLOAD = wx.NewId()



class SRBFileTransfer(wx.Dialog):
    """ This object transfers files between the SRB and the local file system and displays transfer progress. """
    def __init__(self, parent, title, fileName, fileSize, localDir, connectionID, collectionName, direction, bufferSize):
        """ Set up the Dialog Box and all GUI Widgets. """
        # Set up local variables
        self.parent = parent
        self.fileSize = fileSize
        self.bufferSize = int(bufferSize)

        # Create the Dialog Box itself, with no minimize/maximize/close buttons
        wx.Dialog.__init__(self, parent, -1, title, size = (350,200), style=wx.CAPTION)

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
        btn = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        # Add the element to the Main Sizer
        mainSizer.Add(btn, 0, wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        wx.EVT_BUTTON(self, wx.ID_CANCEL, self.OnCancel)

        # Attach the main sizer to the form
        self.SetSizer(mainSizer)
        # Turn Auto Layout on
        self.SetAutoLayout(True)
        # Lay out the form
        self.Layout()
        # Center on the Screen
        self.CenterOnScreen()
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

        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

        # If we are sending files from the SRB to the local File System...
        if direction == srb_DOWNLOAD:
            self.SetTitle(_('Downloading . . .'))

            if 'unicode' in wx.PlatformInfo:
                tmpCollectionName = collectionName.encode(TransanaGlobal.encoding)
                tmpFileName = fileName.encode(TransanaGlobal.encoding)
                
            # Open the proper File Object on the SRB
            FileResult = srb.srbDaiObjOpen(connectionID, tmpCollectionName, tmpFileName, 0)

            # Make sure the object opened correctly
            if FileResult < 0:
                # If the file object did not open correctly, display an Error message.
                self.parent.srbErrorMessage(FileResult)
            else:
                # Create and Open a local binary file for writing to accept the data being sent from the SRB
                outputFile = file(localDir + fileName, 'wb', self.bufferSize)

                # While there is data, read it from the SRB and write it to the local file system.
                # ("while 1" tells the program to just keep looping until a "break" command is triggered.)
                while True:
                    BufSize = self.bufferSize
                    # Initialize the buffer to empty spaces to prepare for the SRB Call
                    Buf = ' ' * BufSize
                    # Get a block of data from the SRB and put it in the Buffer
                    fileWrite = srb.srbDaiObjRead(connectionID, FileResult, Buf, BufSize)
                    # The SRB returns the number of bytes read or an error message (negative value)
                    if fileWrite < 0:
                        # A SRB Error has occurred
                        self.parent.srbErrorMessage(fileWrite)
                    elif (fileWrite == 0) or self.cancelled:
                        # The file has no more data, so we need to break out of the while loop, OR
                        # the user has requested to cancel the file transfer.
                        break
                    else:
                        # Data is read from the file
                        # We are tracking the number of bytes read for progress reporting purposes,
                        # and so that we know what size buffer we need for reading more data
                        BytesRead = BytesRead + fileWrite

                        # Reduce the size of the Buffer to match the data return size.  (We were getting
                        # incomplete downloads without this step!)
                        if fileWrite < self.bufferSize:
                            Buf = Buf[:fileWrite]

                        try:
                            # Let's write the data in the buffer to the local file
                            outputFile.write(Buf)
                        except exceptions.IOError, (errNum, errStr):
                            # The transfer must be interrupted
                            self.cancelled = True
                            # To display an error message, we need to import the error dialog
                            import Dialogs
                            # Create and display the error message
                            errDlg = Dialogs.ErrorDialog(self, errStr)
                            errDlg.ShowModal()
                            errDlg.Destroy()

                        # Let's provide the user with feedback as we read the file
                        self.UpdateDisplay(BytesRead)

                # When all the data has been read (or the transfer cancelled), we can close the local file.
                # Before we do, let's flush the file in case there is still data in the buffer.
                outputFile.flush()
                outputFile.close()

            # We can now close the SRB Data object (file)
            TempInt = srb.srbDaiObjClose(connectionID, FileResult)
            # srbDaiObjClose Error Checking
            if TempInt < 0:
                self.parent.srbErrorMessage(TempInt)


        # If we are sending data from the local file system to the SRB ...
        elif direction == srb_UPLOAD:
            self.SetTitle(_('Uploading . . .'))

            # For now, all files can be identified as "unknown" type.  TODO:  Add File Type specifications
            fileType = 'unknown'

            # Open the local binary file for reading
            inputFile = file(localDir + fileName, 'rb', self.bufferSize)

            # Create a file on the SRB to receive the data
            if fileSize <= sys.maxint:
                fs = ctypes.c_int(fileSize)
            else:
                fs = ctypes.c_longlong(fileSize)

            # Make Unicode conversions.
            if 'unicode' in wx.PlatformInfo:
                fileName = fileName.encode(TransanaGlobal.encoding)
                tmpResource = self.parent.tmpResource
                collectionName = collectionName.encode(TransanaGlobal.encoding)
            else:
                tmpResource = self.parent.srbResource
                
            FileResult = srb.srbDaiObjCreate(connectionID, 0, fileName, fileType, tmpResource, collectionName, fs)

            # See if the SRB File Object was successfully created
            if FileResult < 0:
                # A SRB Error has occurred
                self.parent.srbErrorMessage(FileResult)
            else:
                # The file was successfully created on the SRB.  Now transfer the data from the Local file system.
                # ("while 1" tells the program to just keep looping until a "break" command is triggered.)
                while 1:
                    # If there is more data left than self.bufferSize, use the Max Buffer Size.  Otherwise, use
                    # a buffer just big enough for the rest of the data.
                    if fileSize - BytesRead >= self.bufferSize:
                        BufSize = self.bufferSize
                    else:
                        BufSize = fileSize - BytesRead

                    # Read the appropriate number of bytes from the local file
                    Buf = inputFile.read(BufSize)

                    # If there is not more data in the file, or the user presses the "Cancel" button,
                    # break out of the while loop
                    if (not Buf) or self.cancelled:
                        break

                    # Send the data to the SRB File Object
                    fileWrite = srb.srbDaiObjWrite(connectionID, FileResult, Buf, ctypes.c_int(BufSize))

                    # The SRB returns the number of bytes written or an error message (negative value)
                    if fileWrite < 0:
                        # A SRB Error has occurred
                        self.parent.srbErrorMessage(fileWrite)
                    else:
                        # Data has been successfully written to the file on the SRB.
                        # We need to keep track of our progress for user notification and for determining
                        # the size of the buffer to use.
                        BytesRead = BytesRead + fileWrite
                        # provide User Feedback on the file transfer progress
                        self.UpdateDisplay(BytesRead)

            # Close the file on the Local File System
            inputFile.close()

            # Close the file object on the SRB
            TempInt = srb.srbDaiObjClose(connectionID, FileResult)
            # srbDaiObjClose Error Checking
            if TempInt < 0:
                self.parent.srbErrorMessage(TempInt)

        if self.cancelled:
            # If we are sending files from the SRB to the local File System...
            if direction == srb_DOWNLOAD:
                # Delete the file from the local file system
                os.remove(localDir + fileName)
            # If we are sending data from the local file system to the SRB ...
            elif direction == srb_UPLOAD:
                delResult = srb.srbDaiRemoveObj(connectionID, fileName, collectionName, 0)
                if delResult < 0:
                    self.srbErrorMessage(delResult)
                    
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

    def UpdateDisplay(self, bytesTransferred):
        """ Update the Transfer Dialog Box labels and progress bar """
        # Display Number of Bytes tranferred
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("%d bytes of %d transferred"), 'utf8')
        else:
            prompt = _("%d bytes of %d transferred")
        self.lblBytes.SetLabel(prompt % (bytesTransferred, self.fileSize))
        # Avoiding division by zero problems ...
        if bytesTransferred > 0:
            # ... display % of total bytes transferred
            self.lblPercent.SetLabel("%5.1f %%" % (float(bytesTransferred) / float(self.fileSize) * 100))
            # Update the Progress Bar
            self.progressBar.SetValue(int(float(bytesTransferred) / float(self.fileSize) * 100))
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
            timeRemaining = (((self.fileSize - bytesTransferred) / 1024.0) / (speed))
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
