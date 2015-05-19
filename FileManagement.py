# Copyright (C) 2003 - 2010 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the File Management window for Transana, but can be also run as a stand-alone utility.
   The File Management Tool is designed to facilitate local file management, such as copying and moving video
   files, and updating the Transana Database when video files have been moved.  It is also used for connecting
   to the Storage Resource Broker (SRB) and transfering files between the local file system and the SRB.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare <rasambhare@wisc.edu>'

import wx  # Import wxPython

if __name__ == '__main__':
    __builtins__._ = wx.GetTranslation
    wx.SetDefaultPyEncoding('utf_8')

import os                # used to determine what files are on the local computer
import sys               # Python sys module, used in exception reporting
import exceptions
import pickle            # Python pickle module
import string            # String manipulation
import shutil            # used to copy files on the local file system
import gettext           # localization module
import ctypes            # used for connecting to the srbClient DLL/Linked Library in Windows
import re                # used in parsing list of files in the list control
import SRBConnection     # SRB Connection Parameters dialog box
import SRBFileTransfer   # SRB File Transfer Progress Box and FTP Logic
import DBInterface       # Import Transana's Database Interface
import Dialogs           # Import Transana's Dialog Boxes
import Misc              # import Transana's Miscellaneous functions
import TransanaConstants # used for getting list of fileTypes
import TransanaGlobal    # get Transana's globals, used to get current encoding


class FMFileDropTarget(wx.FileDropTarget):
    """ This class allows the Source and Destination File Windows to become "Drop Targets" for
        wxFileDataObjects, enabling file name drag and drop functionality """
    def __init__(self, FMWindow, activeSide):
        """ Initialize a File Manager File Drop Target Object """
        # Instantiate a wxFileDropTarget
        wx.FileDropTarget.__init__(self)
        # Identify which widget should receive the File Drop
        self.activeSide = activeSide
        # This is a pointer to the File Management Window, used to implement the File Drop.
        # This is probably bad form.  The FileDropTarget Class should probably call
        # methods in the class which implements it.  However, this seems clearer in my warped
        # way of thinking.
        self.FileManagementWindow = FMWindow

    def OnDropFiles(self, x, y, filenames):
        """ Implements File Drop """
        # Determine the necessary values to implement the file drop
        # on either side of the File Management Screen
        if (self.activeSide == 'Left'):
            targetLbl = self.FileManagementWindow.lblLeft
            # Local and Remote File Systems have different methods for determing the selected Path
            if targetLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
                targetDir = self.FileManagementWindow.dirLeft.GetPath()
            else:
                targetDir = self.FileManagementWindow.GetFullPath(self.FileManagementWindow.srbDirLeft)
            target = self.FileManagementWindow.fileLeft
            filter = self.FileManagementWindow.filterLeft
            # We need to know the other side to, so it can also be updated if necessary
            otherLbl = self.FileManagementWindow.lblRight
            # Local and Remote File Systems have different methods for determing the selected Path
            if otherLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
                otherDir = self.FileManagementWindow.dirRight.GetPath()
            else:
                otherDir = self.FileManagementWindow.GetFullPath(self.FileManagementWindow.srbDirRight)
            othertarget = self.FileManagementWindow.fileRight
            otherfilter = self.FileManagementWindow.filterRight
        elif self.activeSide == 'Right':
            targetLbl = self.FileManagementWindow.lblRight
            # Local and Remote File Systems have different methods for determing the selected Path
            if targetLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
                targetDir = self.FileManagementWindow.dirRight.GetPath()
            else:
                targetDir = self.FileManagementWindow.GetFullPath(self.FileManagementWindow.srbDirRight)
            target = self.FileManagementWindow.fileRight
            filter = self.FileManagementWindow.filterRight
            # We need to know the other side to, so it can also be updated if necessary
            otherLbl = self.FileManagementWindow.lblLeft
            # Local and Remote File Systems have different methods for determing the selected Path
            if otherLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
                otherDir = self.FileManagementWindow.dirLeft.GetPath()
            else:
                otherDir = self.FileManagementWindow.GetFullPath(self.FileManagementWindow.srbDirLeft)
            othertarget = self.FileManagementWindow.fileLeft
            otherfilter = self.FileManagementWindow.filterLeft

        if os.path.split(filenames[0])[0] != targetDir:
            # If a Path has been defined for the Drop Target, copy the files.
            # First, determine if it's local or SRB receiving the files
            if targetLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
                if targetDir != '':
                    self.FileManagementWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
                    # for each file in the list...
                    for fileNm in filenames:
                        prompt = _('Copying %s to %s')
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(prompt, 'utf8')
                        self.FileManagementWindow.SetStatusText(prompt % (fileNm, targetDir))
                        # extract the file name...
                        (dir, fn) = os.path.split(fileNm)
                        # copy the file to the destination path
                        shutil.copyfile(fileNm, os.path.join(targetDir, fn))
                        # Update the File Window where the files were dropped.
                        self.FileManagementWindow.RefreshFileList(targetLbl.GetLabel(), targetDir, target, filter)
                        # Let's update the target control after every file so that the new files show up in the list ASAP.
                        # wxYield allows the Windows Message Queue to update the display.
                        wx.Yield()
                    self.FileManagementWindow.SetStatusText(_('Copy complete.'))
                    self.FileManagementWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
                    # If both sides of the File Manager are pointed to the same folder, we need to update the other
                    # side too!
                    if targetDir == otherDir:
                        self.FileManagementWindow.RefreshFileList(otherLbl.GetLabel(), otherDir, othertarget, otherfilter)
                    return(True)
                # If no Path is defined for the Drop Target, forget about the drop
                else:
                    dlg = Dialogs.ErrorDialog(self.FileManagementWindow, _('No destination path has been specified.  File copy cancelled.'))
                    dlg.ShowModal()
                    dlg.Destroy()
                    return(False)
            else:
                cancelPressed = False
                # For each file in the fileList ...
                for fileName in filenames:
                    if cancelPressed:
                        break
                    # Divide the fileName into directory and filename portions
                    (sourceDir, fileNm) = os.path.split(fileName)
                    # Get the File Size
                    fileSize = os.path.getsize(fileName)
                    prompt = _('Copying %s to %s')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(prompt, 'utf8')
                    self.FileManagementWindow.SetStatusText(prompt % (fileNm, targetDir))

                    # The SRBFileTransfer class handles file transfers and provides Progress Feedback
                    dlg = SRBFileTransfer.SRBFileTransfer(self.FileManagementWindow, _("SRB File Transfer"), fileNm, fileSize, sourceDir, self.FileManagementWindow.srbConnectionID, targetDir, SRBFileTransfer.srb_UPLOAD, self.FileManagementWindow.srbBuffer)
                    success = dlg.TransferSuccessful()
                    dlg.Destroy()

                    if success:
                        self.FileManagementWindow.SetStatusText(_('Copy complete.'))
                    else:
                        self.FileManagementWindow.SetStatusText(_('Copy cancelled.'))
                        cancelPressed = True
                    # Let's update the target control after every file so that the new files show up in the list ASAP.
                    self.FileManagementWindow.RefreshFileList(targetLbl, targetDir, target, filter)
                    # wxYield allows the Windows Message Queue to update the display.
                    wx.Yield()
                    # If both sides of the File Manager are pointed to the same folder, we need to update the other
                    # side too!
                    if targetDir == otherDir:
                        self.FileManagementWindow.RefreshFileList(otherLbl.GetLabel(), otherDir, othertarget, otherfilter)


class FileManagement(wx.Dialog):
    """ This displays the main File Management window. """
    def __init__(self,parent,id,title):
        """ Initialize the Main File Management Window.  (You must also call Setup) """

        # Because of Unicode issues, it's easiest to use Constants for the Local and Remote labels.
        # Otherwise, there are comparison problems.
        self.LOCAL_LABEL = _('Local:')
        if 'unicode' in wx.PlatformInfo:
            self.LOCAL_LABEL = unicode(self.LOCAL_LABEL, 'utf8')
        self.REMOTE_LABEL = _('Remote:')
        if 'unicode' in wx.PlatformInfo:
            self.REMOTE_LABEL = unicode(self.REMOTE_LABEL, 'utf8')

        # srbConnectionID indicates the Connection ID for the Storage Resource Broker connection.
        # It is initialized to None, as the connection is not established by default
        self.srbConnectionID = None
        # Initialize the SRB Buffer Size to 400000.  This size yeilds fast transfers on my office computer with a very fast connection.
        self.srbBuffer = '400000'

        # A wxDialog is used rather than a wxFrame.  While the frame has a Status Bar, which would have made some things
        # easier, a Dialog can be displayed modally, which this tool requires.

        # The creation of the Dialog follows the model from the wxPython 2.4.0.7 Demo.
        # This is necessary for this code to run as a stand-alone program.

        # First, create a "PreDialog" which will get displayed in a few steps.
        pre = wx.PreDialog()
        # If File Management is running as a stand-alone, it needs an additional style that indicates
        # that it is being created without a parent.
        if __name__ == '__main__':
            pre.SetExtraStyle(wx.DIALOG_NO_PARENT)
        # Now create the actual GUI Dialog using the Create method
        pre.Create(parent, -1, title, pos = wx.DefaultPosition, size = (800,650), style=wx.DEFAULT_DIALOG_STYLE | wx.CAPTION | wx.RESIZE_BORDER)
        # Convert the PreDialog into a REAL wrapper of the wxPython extension (whatever that means)
        self.this = pre.this

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the Connection Dialog
        self.SRBConnDlg = SRBConnection.SRBConnection(self)


    def Setup(self, showModal=False):
        """ Set up the form widgets for the File Management Window """
        # remember the showModal setting
        self.isModal = showModal

        # Create the form's Vertical Sizer
        formSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a main Horizontal Sizer
        mainSizer = wx.BoxSizer(wx.HORIZONTAL)

        # File Listings on the Left
        filesLeftSizer = wx.BoxSizer(wx.VERTICAL)

        # Connection Label
        self.lblLeft = wx.StaticText(self, -1, _("Local:"))
        # Add label to Sizer
        filesLeftSizer.Add(self.lblLeft, 0, wx.TOP | wx.LEFT | wx.RIGHT, 5)

        # Directory Listing
        self.dirLeft = wx.GenericDirCtrl(self, -1, style=wx.DIRCTRL_DIR_ONLY | wx.BORDER_DOUBLE)
        # Add directory listing to Sizer
        filesLeftSizer.Add(self.dirLeft, 3, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 5)

        # If we're NOT running stand-alone ...
        if __name__ != '__main__':
            # Set the initial directory
            self.dirLeft.SetPath(TransanaGlobal.configData.videoPath)
        else:
            if os.path.exists('E:\\Video'):
                self.dirLeft.SetPath('E:\\Video')

        # Source SRB Collections listing
        self.srbDirLeft = wx.TreeCtrl(self, -1, style=wx.TR_HAS_BUTTONS | wx.TR_SINGLE | wx.BORDER_DOUBLE)
        # This control is not visible initially, as Local Folders, not SRB Collections are shown
        self.srbDirLeft.Show(False)
        # Add directory listing to Sizer
        filesLeftSizer.Add(self.srbDirLeft, 3, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 5)

        # NOTE:  Although it would be possible to display files as well as folders in the wxGenericDirCtrls,
        #        we chose not to so that users could select multiple files at once for manipulation

        # File Listing
        self.fileLeft = wx.ListCtrl(self, -1, style=wx.LC_REPORT | wx.LC_NO_HEADER | wx.BORDER_DOUBLE)
        # Create a single column for the file names
        self.fileLeft.InsertColumn(0, "Files")
        self.fileLeft.SetColumnWidth(0, 100)
        # Add file listing to Sizer
        filesLeftSizer.Add(self.fileLeft, 5, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 5)

        # Make the File List a FileDropTarget
        dt = FMFileDropTarget(self, 'Left')
        self.fileLeft.SetDropTarget(dt)

        # We need to rebuild the file types list so the items get translated if a non-English language is being used!
        # Start with an empty list
        fileChoiceList = []
        # For each item in the file types list ...
        for item in TransanaConstants.fileTypesList:
            # ... add the translated item to the list
            fileChoiceList.append(_(item))
        # Create the File Types dropdown
        self.filterLeft = wx.Choice(self, -1, choices=fileChoiceList)
        # Set the default selection to the first item in the list
        self.filterLeft.SetSelection(0)
        # Add the file types dropdown to the Sizer
        filesLeftSizer.Add(self.filterLeft, 0, wx.EXPAND | wx.ALL, 5)


        # Buttons in the middle
        buttonSizer = wx.BoxSizer(wx.VERTICAL)

        # Get the BitMaps for the left and right arrow buttons
        bmpLeft = wx.ArtProvider_GetBitmap(wx.ART_GO_BACK, wx.ART_TOOLBAR, (16,16))
        bmpRight = wx.ArtProvider_GetBitmap(wx.ART_GO_FORWARD, wx.ART_TOOLBAR, (16,16))

        # Top Spacer
        buttonSizer.Add((1,1), 1)

        # Copy Buttons - Create Label
        txtCopy = wx.StaticText(self, -1, _("Copy"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(txtCopy, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Create left and right arrow bitmap buttons
        self.btnCopyToLeft = wx.BitmapButton(self, -1, bmpLeft, size=(36, 24))
        self.btnCopyToRight = wx.BitmapButton(self, -1, bmpRight, size=(36, 24))
        # Create a Horizontal Sizer for the buttons
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the arrow buttons to the button sizer with a spacer between
        hSizer.Add(self.btnCopyToLeft, 0)
        hSizer.Add((3, 0))
        hSizer.Add(self.btnCopyToRight, 0)
        # Add the horizontal button sizer to the vertical button area sizer
        buttonSizer.Add(hSizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Spacer
        buttonSizer.Add((1,1), 2)

        # Move Buttons - Create Label
        txtMove = wx.StaticText(self, -1, _("Move"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(txtMove, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Create left and right arrow bitmap buttons
        self.btnMoveToLeft = wx.BitmapButton(self, -1, bmpLeft, size=(36, 24))
        self.btnMoveToRight = wx.BitmapButton(self, -1, bmpRight, size=(36, 24))
        # Create a Horizontal Sizer for the buttons
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the arrow buttons to the button sizer with a spacer between
        hSizer.Add(self.btnMoveToLeft, 0)
        hSizer.Add((3, 0))
        hSizer.Add(self.btnMoveToRight, 0)
        # Add the horizontal button sizer to the vertical button area sizer
        buttonSizer.Add(hSizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Spacer
        buttonSizer.Add((1,1), 2)

        # Copy All New Buttons - Create Label
        txtSynch = wx.StaticText(self, -1, _("Copy All New"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(txtSynch, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Create left and right arrow bitmap buttons
        self.btnSynchToLeft = wx.BitmapButton(self, -1, bmpLeft, size=(36, 24))
        self.btnSynchToRight = wx.BitmapButton(self, -1, bmpRight, size=(36, 24))
        # Create a Horizontal Sizer for the buttons
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the arrow buttons to the button sizer with a spacer between
        hSizer.Add(self.btnSynchToLeft, 0)
        hSizer.Add((3, 0))
        hSizer.Add(self.btnSynchToRight, 0)
        # Add the horizontal button sizer to the vertical button area sizer
        buttonSizer.Add(hSizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Spacer
        buttonSizer.Add((1,1), 2)

        # Delete Buttons - Create Label
        txtDelete = wx.StaticText(self, -1, _("Delete"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(txtDelete, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Create left and right arrow bitmap buttons
        self.btnDeleteLeft = wx.BitmapButton(self, -1, bmpLeft, size=(36, 24))
        self.btnDeleteRight = wx.BitmapButton(self, -1, bmpRight, size=(36, 24))
        # Create a Horizontal Sizer for the buttons
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the arrow buttons to the button sizer with a spacer between
        hSizer.Add(self.btnDeleteLeft, 0)
        hSizer.Add((3, 0))
        hSizer.Add(self.btnDeleteRight, 0)
        # Add the horizontal button sizer to the vertical button area sizer
        buttonSizer.Add(hSizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Spacer
        buttonSizer.Add((1,1), 2)

        # New Folder Buttons - Create Label
        txtNewFolder = wx.StaticText(self, -1, _("New Folder"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(txtNewFolder, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Create left and right arrow bitmap buttons
        self.btnNewFolderLeft = wx.BitmapButton(self, -1, bmpLeft, size=(36, 24))
        self.btnNewFolderRight = wx.BitmapButton(self, -1, bmpRight, size=(36, 24))
        # Create a Horizontal Sizer for the buttons
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the arrow buttons to the button sizer with a spacer between
        hSizer.Add(self.btnNewFolderLeft, 0)
        hSizer.Add((3, 0))
        hSizer.Add(self.btnNewFolderRight, 0)
        # Add the horizontal button sizer to the vertical button area sizer
        buttonSizer.Add(hSizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Spacer
        buttonSizer.Add((1,1), 2)

        # Update DB Buttons - Create Label
        txtUpdateDB = wx.StaticText(self, -1, _("Update DB"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(txtUpdateDB, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Create left and right arrow bitmap buttons
        self.btnUpdateDBLeft = wx.BitmapButton(self, -1, bmpLeft, size=(36, 24))
        self.btnUpdateDBRight = wx.BitmapButton(self, -1, bmpRight, size=(36, 24))
        # Create a Horizontal Sizer for the buttons
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the arrow buttons to the button sizer with a spacer between
        hSizer.Add(self.btnUpdateDBLeft, 0)
        hSizer.Add((3, 0))
        hSizer.Add(self.btnUpdateDBRight, 0)
        # Add the horizontal button sizer to the vertical button area sizer
        buttonSizer.Add(hSizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Spacer
        buttonSizer.Add((1,1), 2)

        # If we're running stand-alone ...
        if __name__ == '__main__':
            # ... hide the UpdateDB Buttons
            txtUpdateDB.Show(False)
            self.btnUpdateDBLeft.Show(False)
            self.btnUpdateDBRight.Show(False)

        # Connect Buttons - Create Label
        txtConnect = wx.StaticText(self, -1, _("Connect"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(txtConnect, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Create left and right arrow bitmap buttons
        self.btnConnectLeft = wx.BitmapButton(self, -1, bmpLeft, size=(36, 24))
        self.btnConnectRight = wx.BitmapButton(self, -1, bmpRight, size=(36, 24))
        # Create a Horizontal Sizer for the buttons
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the arrow buttons to the button sizer with a spacer between
        hSizer.Add(self.btnConnectLeft, 0)
        hSizer.Add((3, 0))
        hSizer.Add(self.btnConnectRight, 0)
        # Add the horizontal button sizer to the vertical button area sizer
        buttonSizer.Add(hSizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Spacer
        buttonSizer.Add((1,1), 2)

        # Refresh Buttons - Create Label
        self.btnRefresh = wx.Button(self, -1, _("Refresh"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(self.btnRefresh, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Spacer
        buttonSizer.Add((1,1), 2)

        # Close Button - Create Label
        self.btnClose = wx.Button(self, -1, _("Close"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(self.btnClose, 0, wx.ALIGN_CENTER | wx.TOP | wx.LEFT | wx.RIGHT, 5)
        # Spacer
        buttonSizer.Add((1,1), 2)

        # Help Button - Create Label
        self.btnHelp = wx.Button(self, -1, _("Help"))
        # Add the button label to the vertical sizer
        buttonSizer.Add(self.btnHelp, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        # If we're running stand-alone ...
        if __name__ == '__main__':
            # ... hide the Help Button
            self.btnHelp.Show(False)

        # Bottom Spacer
        buttonSizer.Add((1,1), 1)


        # File Listings on the Right
        filesRightSizer = wx.BoxSizer(wx.VERTICAL)

        # Connection Label
        self.lblRight = wx.StaticText(self, -1, _("Local:"))
        # Add label to Sizer
        filesRightSizer.Add(self.lblRight, 0, wx.TOP | wx.LEFT | wx.RIGHT, 5)

        # Directory Listing
        self.dirRight = wx.GenericDirCtrl(self, -1, style=wx.DIRCTRL_DIR_ONLY | wx.BORDER_DOUBLE)
        # Add directory listing to Sizer
        filesRightSizer.Add(self.dirRight, 3, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 5)

        # If we're NOT running stand-alone ...
        if __name__ != '__main__':
            # ... set the initial directory to the video path
            self.dirRight.SetPath(TransanaGlobal.configData.videoPath)
        else:
            if os.path.exists('E:\\Video'):
                self.dirRight.SetPath('E:\\Video')

        # Destination SRB Collections listing
        self.srbDirRight = wx.TreeCtrl(self, -1, style=wx.TR_HAS_BUTTONS | wx.TR_SINGLE | wx.BORDER_DOUBLE)
        # This control is not visible initially, as Local Folders, not SRB Collections are shown
        self.srbDirRight.Show(False)
        # Add directory listing to Sizer
        filesRightSizer.Add(self.srbDirRight, 3, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 5)

        # NOTE:  Although it would be possible to display files as well as folders in the wxGenericDirCtrls,
        #        we chose not to so that users could select multiple files at once for manipulation

        # File Listing
        self.fileRight = wx.ListCtrl(self, -1, style=wx.LC_REPORT | wx.LC_NO_HEADER | wx.BORDER_DOUBLE)
        # Create a single column for the file names
        self.fileRight.InsertColumn(0, "Files")
        self.fileRight.SetColumnWidth(0, 100)
        # Add file listing to Sizer
        filesRightSizer.Add(self.fileRight, 5, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 5)

        # Make the File List a FileDropTarget
        dt = FMFileDropTarget(self, 'Right')
        self.fileRight.SetDropTarget(dt)

        # We need to rebuild the file types list so the items get translated if a non-English language is being used!
        # Start with an empty list
        fileChoiceList = []
        # For each item in the file types list ...
        for item in TransanaConstants.fileTypesList:
            # ... add the translated item to the list
            fileChoiceList.append(_(item))
        # Create the File Types dropdown
        self.filterRight = wx.Choice(self, -1, choices=fileChoiceList)
        # Set the default selection to the first item in the list
        self.filterRight.SetSelection(0)
        # Add the file types dropdown to the Sizer
        filesRightSizer.Add(self.filterRight, 0, wx.EXPAND | wx.ALL, 5)

        # Add the three vertical columns to the main sizer
        mainSizer.Add(filesLeftSizer, 2, wx.EXPAND, 0)
        mainSizer.Add(buttonSizer, 0, wx.EXPAND, 0)
        mainSizer.Add(filesRightSizer, 2, wx.EXPAND, 0)

        # Add the main Sizer to the form's Sizer
        formSizer.Add(mainSizer, 1, wx.EXPAND, 0)


        # wxFrames can have Status Bars.  wxDialogs can't.
        # The File Management Window needs to be a wxDialog, not a wxFrame.
        # Yet I want it to have a Status Bar.  So I'm going to fake a status bar by
        # creating a Panel at the bottom of the screen with text on it.

        # Place a Static Text Widget on the Status Bar Panel
        self.StatusText = wx.StaticText(self, -1, "", style=wx.ST_NO_AUTORESIZE)
        formSizer.Add(self.StatusText, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        # Set the mainSizer as the form's main sizer
        self.SetSizer(formSizer)


        # Define Events

        # Mouse Over Events to enable Button Descriptions in the Status Bar
        self.btnCopyToLeft.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnCopyToLeft.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnCopyToRight.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnCopyToRight.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnMoveToLeft.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnMoveToLeft.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnMoveToRight.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnMoveToRight.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnSynchToLeft.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnSynchToLeft.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnSynchToRight.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnSynchToRight.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnDeleteLeft.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnDeleteLeft.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnDeleteRight.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnDeleteRight.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnNewFolderLeft.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnNewFolderLeft.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnNewFolderRight.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnNewFolderRight.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnUpdateDBLeft.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnUpdateDBLeft.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnUpdateDBRight.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnUpdateDBRight.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnConnectLeft.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnConnectLeft.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnConnectRight.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnConnectRight.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnRefresh.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnRefresh.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnClose.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnClose.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.btnHelp.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseOver)
        self.btnHelp.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)

        # Define the Event to initiate the beginning of a File Drag operation
        self.fileLeft.Bind(wx.EVT_MOTION, self.OnFileStartDrag)
        self.fileRight.Bind(wx.EVT_MOTION, self.OnFileStartDrag)

        # Button Events (link functionality to button widgets)
        self.btnCopyToLeft.Bind(wx.EVT_BUTTON, self.OnCopyMove)
        self.btnCopyToRight.Bind(wx.EVT_BUTTON, self.OnCopyMove)
        self.btnMoveToLeft.Bind(wx.EVT_BUTTON, self.OnCopyMove)
        self.btnMoveToRight.Bind(wx.EVT_BUTTON, self.OnCopyMove)
        self.btnSynchToLeft.Bind(wx.EVT_BUTTON, self.OnSynch)
        self.btnSynchToRight.Bind(wx.EVT_BUTTON, self.OnSynch)
        self.btnDeleteLeft.Bind(wx.EVT_BUTTON, self.OnDelete)
        self.btnDeleteRight.Bind(wx.EVT_BUTTON, self.OnDelete)
        self.btnNewFolderLeft.Bind(wx.EVT_BUTTON, self.OnNewFolder)
        self.btnNewFolderRight.Bind(wx.EVT_BUTTON, self.OnNewFolder)
        self.btnUpdateDBLeft.Bind(wx.EVT_BUTTON, self.OnUpdateDB)
        self.btnUpdateDBRight.Bind(wx.EVT_BUTTON, self.OnUpdateDB)
        self.btnConnectLeft.Bind(wx.EVT_BUTTON, self.OnConnect)
        self.btnConnectRight.Bind(wx.EVT_BUTTON, self.OnConnect)
        self.btnRefresh.Bind(wx.EVT_BUTTON, self.OnRefresh)
        self.btnClose.Bind(wx.EVT_BUTTON, self.CloseWindow)
        self.btnHelp.Bind(wx.EVT_BUTTON, self.OnHelp)

        # Directory Tree and SRB Collection Selection Events
        self.dirLeft.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnDirSelect)
        self.dirRight.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnDirSelect)
        self.srbDirLeft.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnDirSelect)
        self.srbDirRight.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnDirSelect)

        # File Filter Events
        self.filterLeft.Bind(wx.EVT_CHOICE, self.OnDirSelect)
        self.filterRight.Bind(wx.EVT_CHOICE, self.OnDirSelect)

        # Form Resize Event
        self.Bind(wx.EVT_SIZE, self.OnSize)

        # File Manager Close Event
        self.Bind(wx.EVT_CLOSE, self.CloseWindow)

        # Define the minimum size of this dialog box
        self.SetSizeHints(600, 500)

        # lay out all the controls
        self.Layout()
        # Enable Auto Layout
        self.SetAutoLayout(True)
        self.CenterOnScreen()
        # Call OnSize to set the srb directory control size initially
        self.OnSize(None)

        # Populate File Lists for initial directory settings
        self.RefreshFileList(self.lblLeft.GetLabel(), self.dirLeft.GetPath(), self.fileLeft, self.filterLeft)
        self.RefreshFileList(self.lblRight.GetLabel(), self.dirRight.GetPath(), self.fileRight, self.filterRight)

        # Show the File Management Window
        if showModal:
            self.ShowModal()
        else:
            self.Show(True)

    def SetStatusText(self, txt):
        """ Update the Text in the fake Status Bar """
        self.StatusText.SetLabel(txt)

    def srbDisconnect(self):
        """ Disconnect from the SRB """
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
        # Break the connection to the SRB.
        srb.srb_disconnect(self.srbConnectionID.value)
        # Signal that the Connection has been broken by resetting the ConnectionID value to None
        self.srbConnectionID = None

    def OnSize(self, event):
        """ Form Resize Event """
        # If event != None (We're not on the initial sizing call) ...
        if event != None:
            # ... call the underlying Resize Event
            event.Skip()

        # If the LOCAL directory listing is showing on the LEFT ...
        if self.dirLeft.IsShown():
            # Get the Size and Position of the current Directory Controls
            rectLeft = self.dirLeft.GetRect()
            # Set the SRB Directory Controls to the same position and size as the regular Directory Controls
            self.srbDirLeft.SetDimensions(rectLeft[0], rectLeft[1], rectLeft[2], rectLeft[3])
        # If the REMOTE directory listing is showing on the LEFT ...
        else:
            # Get the Size and Position of the current Directory Controls
            rectLeft = self.srbDirLeft.GetRect()
            # Set the SRB Directory Controls to the same position and size as the regular Directory Controls
            self.dirLeft.SetDimensions(rectLeft[0], rectLeft[1], rectLeft[2], rectLeft[3])

        # If the LOCAL directory listing is showing on the RIGHT ...
        if self.dirRight.IsShown():
            # Get the Size and Position of the current Directory Controls
            rectRight = self.dirRight.GetRect()
            # Set the SRB Directory Controls to the same position and size as the regular Directory Controls
            self.srbDirRight.SetDimensions(rectRight[0], rectRight[1], rectRight[2], rectRight[3])
        # If the REMOTE directory listing is showing on the RIGHT ...
        else:
            # Get the Size and Position of the current Directory Controls
            rectRight = self.srbDirRight.GetRect()
            # Set the SRB Directory Controls to the same position and size as the regular Directory Controls
            self.dirRight.SetDimensions(rectRight[0], rectRight[1], rectRight[2], rectRight[3])

        # reset column widths for the File Lists
        if self.fileLeft.GetColumnCount() > 0:
            self.fileLeft.SetColumnWidth(0, self.fileLeft.GetSizeTuple()[0] - 24)
        if self.fileRight.GetColumnCount() > 0:
            self.fileRight.SetColumnWidth(0, self.fileRight.GetSizeTuple()[0] - 24)

    def CloseWindow(self, event):
        """ Clean up upon closing the File Management Window """
        try:
            # If we are connected to the SRB, we need to disconnect before closing the window
            if self.srbConnectionID != None:
                self.srbDisconnect()
        except:
            pass

        # Destroy the SRB Connection Dialog 
        self.SRBConnDlg.Destroy()
        # Hide the File Management Dialog.  (You can't destroy it here, like I used to, because that crashes on Mac as of
        # wxPython 2.8.6.1.
        self.Show(False)
        # As of wxPython 2.8.6.1, you can't destroy yourself during Close on the Mac if you're Modal, and let's delay it a bit
        # if you're not Modal.
        if not self.isModal:
            wx.CallAfter(self.Destroy)

    def OnMouseOver(self, event):
        """ Update the Status Bar when a button is moused-over """
        # Identify controls for the SOURCE data.  Copy, Move, and Synch RIGHT operate left to right, while all
        # other controls operate on the side they point to.
        if event.GetId() in [self.btnCopyToRight.GetId(), self.btnMoveToRight.GetId(), self.btnSynchToRight.GetId(),
                             self.btnDeleteLeft.GetId(), self.btnNewFolderLeft.GetId(), self.btnUpdateDBLeft.GetId(),
                             self.btnConnectLeft.GetId()]:
            # Get the Local / Remote label
            sourceLbl = self.lblLeft
            # Identify which side of the screen we're working on
            sideLbl = _("left")
            # The directory control and file path differ for local and remote file systems
            # If we are LOCAL ...
            if sourceLbl.GetLabel() == self.LOCAL_LABEL:
                # ... get the LOCAL path
                sourcePath = self.dirLeft.GetPath()
            # If we are REMOTE ...
            else:
                # ... get the REMOTE path
                sourcePath = self.GetFullPath(self.srbDirLeft)
            # Get the source File Control
            sourceFile = self.fileLeft
            # The directory control and file path differ for local and remote file systems
            # If the DESTINATION is LOCAL ...
            if self.lblRight.GetLabel() == self.LOCAL_LABEL:
                # ... get the LOCAL path
                destPath = self.dirRight.GetPath()
            # If the DESTINATION is REMOTE ...
            else:
                # ... get the REMOTE path
                destPath = self.GetFullPath(self.srbDirRight)
        # Copy, Move, and Synch LEFT operate right to left, while all other controls operate on the side they point to.
        elif event.GetId() in [self.btnCopyToLeft.GetId(), self.btnMoveToLeft.GetId(), self.btnSynchToLeft.GetId(), 
                               self.btnDeleteRight.GetId(), self.btnNewFolderRight.GetId(), self.btnUpdateDBRight.GetId(),
                               self.btnConnectRight.GetId()]:
            # Get the Local / Remote label
            sourceLbl = self.lblRight
            # Identify which side of the screen we're working on
            sideLbl = _("right")
            # The directory control and file path differ for local and remote file systems
            # If we are LOCAL ...
            if sourceLbl.GetLabel() == self.LOCAL_LABEL:
                # ... get the LOCAL path
                sourcePath = self.dirRight.GetPath()
            # If we are REMOTE ...
            else:
                # ... get the REMOTE path
                sourcePath = self.GetFullPath(self.srbDirRight)
            # Get the source File Control
            sourceFile = self.fileRight
            # The directory control and file path differ for local and remote file systems
            # If the DESTINATION is LOCAL ...
            if self.lblLeft.GetLabel() == self.LOCAL_LABEL:
                # ... get the LOCAL path
                destPath = self.dirLeft.GetPath()
            # If the DESTINATION is REMOTE ...
            else:
                # ... get the REMOTE path
                destPath = self.GetFullPath(self.srbDirLeft)

        # Determine which button is being moused-over and display an appropriate message in the
        # Status Bar

        # Copy buttons
        if event.GetId() in [self.btnCopyToLeft.GetId(), self.btnCopyToRight.GetId()]:
            # Indicate whether selected files or all files will be copied
            if sourceFile.GetSelectedItemCount() == 0:
                tempText = _('all')
            else:
                tempText = _('selected')
            # Create the appropriate Status Bar prompt
            prompt = _('Copy %s files from %s to %s')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                tempText = unicode(tempText, 'utf8')
                prompt = unicode(prompt, 'utf8')
            # Put the prompt in the Status Bar
            self.SetStatusText(prompt % (tempText, sourcePath, destPath))

        # Move buttons
        elif event.GetId() in [self.btnMoveToLeft.GetId(), self.btnMoveToRight.GetId()]:
            # Indicate whether selected files or all files will be copied
            if sourceFile.GetSelectedItemCount() == 0:
                tempText = _('all')
            else:
                tempText = _('selected')
            # Create the appropriate Status Bar prompt
            prompt = _('Move %s files from %s to %s')
            # Specify source and destination Folder/Collections in Status Bar
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                tempText = unicode(tempText, 'utf8')
                prompt = unicode(prompt, 'utf8')
            # Put the prompt in the Status Bar
            self.SetStatusText(prompt % (tempText, sourcePath, destPath))

        # CopyAllNew buttons
        elif event.GetId() in [self.btnSynchToLeft.GetId(), self.btnSynchToRight.GetId()]:
            # Create the appropriate Status Bar prompt
            prompt = _('Copy all new files files from %s to %s')
            # Specify source and destination Folder/Collections in Status Bar
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Put the prompt in the Status Bar
            self.SetStatusText(prompt % (sourcePath, destPath))

        # Delete buttons
        elif event.GetId() in [self.btnDeleteLeft.GetId(), self.btnDeleteRight.GetId()]:
            # Specify source Folder/Collection in the Status Bar
            # If there are files selected in the File List ...
            if (sourceFile.GetItemCount() > 0):
                # Create the appropriate Status Bar prompt
                prompt = _('Delete selected files from %s')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
            # If there are NO files selected in the File List ...
            else:
                # Create the appropriate Status Bar prompt
                prompt = _('Delete %s')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
            # Put the prompt in the Status Bar
            self.SetStatusText(prompt % sourcePath)

        # New Folder button
        elif event.GetId() in [self.btnNewFolderLeft.GetId(), self.btnNewFolderRight.GetId()]:
            # Create the appropriate Status Bar prompt
            prompt = _('Create new folder in %s')
            # Specify source Folder/Collection in the Status Bar
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Put the prompt in the Status Bar
            self.SetStatusText(prompt % sourcePath)

        # Update DB button
        elif event.GetId() in [self.btnUpdateDBLeft.GetId(), self.btnUpdateDBRight.GetId()]:
            if sourceLbl.GetLabel() == self.LOCAL_LABEL:
                # Put the prompt in the Status Bar
                self.SetStatusText(_('Update selected file locations in the Transana database to %s') % sourcePath)
            else:
                # Display error message in the Status Bar
                self.SetStatusText(_('You cannot update the file locations in the database to %s') % sourcePath)

        # Connect button
        elif event.GetId() in [self.btnConnectLeft.GetId(), self.btnConnectRight.GetId()]:
            # Status Bar Text needs to indicate whether this button will establish or break
            # the connection to the SRB
            if sourceLbl.GetLabel() == self.LOCAL_LABEL:
                # Put the prompt in the Status Bar
                self.SetStatusText(_('Connect %s side to SRB') % sideLbl)
            else:
                # Put the prompt in the Status Bar
                self.SetStatusText(_('Disconnect %s side from SRB') % sideLbl)

        # Refresh button
        elif event.GetId() == self.btnRefresh.GetId():
            # Put the prompt in the Status Bar
            self.SetStatusText(_('Refresh File Lists'))

        # Close button
        elif event.GetId() == self.btnClose.GetId():
            # Put the prompt in the Status Bar
            self.SetStatusText(_('Close File Management'))

        # Help button
        elif event.GetId() == self.btnHelp.GetId():
            # Put the prompt in the Status Bar
            self.SetStatusText(_('View File Management Help'))

    def OnMouseExit(self, event):
        """ Clear the Status Bar when the Mouse is no longer over a Button """
        # Set the Status Text to an empty string to clear it.
        self.SetStatusText('')

    def OnFileStartDrag(self, event):
        """ Mouse Left-button Down event initiates Drag for Drag-and-Drop file copying """
        # Event.Skip() calls the inherited OnStartDrag() method so everything will work correctly
        event.Skip()

        # NOTE:  Drag will only work from the Local file system.  This is because SRB files
        #        can't be treated as File objects by a FileDropTarget.

        # Determine which side is Source, which side is Target
        if event.GetId() == self.fileLeft.GetId():
            sourceLbl = self.lblLeft.GetLabel()
            # Local and Remote File Systems have different methods for determing the selected Path
            if sourceLbl == self.LOCAL_LABEL:
                sourceDir = self.dirLeft.GetPath()
            else:
                return       # If it's not local, we're not dragging anything!
            sourceFile = self.fileLeft
            targetDir = self.dirRight.GetPath()
        elif event.GetId() == self.fileRight.GetId():
            sourceLbl = self.lblRight.GetLabel()
            # Local and Remote File Systems have different methods for determing the selected Path
            if sourceLbl == self.LOCAL_LABEL:
                sourceDir = self.dirRight.GetPath()
            else:
                return       # If it's not local, we're not dragging anything!
            sourceFile = self.fileRight
            targetDir = self.dirLeft.GetPath()
        else:
            return

        # We only drag if File Names have been selected
        if event.LeftIsDown() and (sourceFile.GetSelectedItemCount() > 0):
            # Initialize the fileList to an empty list 
            fileList = []
            # Get all Selected files from the Source File List.  Starting with itemNum = -1 signals to search the whole list.
            # Store the data in fileList.
            itemNum = -1
            while True:
                # Get the Next Selected Item in the List Control
                itemNum = sourceFile.GetNextItem(itemNum, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
                # If the return value is -1, no more items in the list are selected
                if itemNum == -1:
                    break
                else:
                    # We build local filenames using os.join to get the right seperator character.
                    # Add the selected item to the fileList and continue looking for more items
                    fileList.append(os.path.join(sourceDir, sourceFile.GetItemText(itemNum)))
            # Create a File Data Object, which holds the File Names that are to be dragged
            tdo = wx.FileDataObject()
            # Add File names to the File Data object
            for fileNm in fileList:
                tdo.AddFile(fileNm)
            # Create a Drop Source Object, which enables the Drag operation
            tds = wx.DropSource(sourceFile)
            # Associate the Data to be dragged with the Drop Source Object
            tds.SetData(tdo)
            # Intiate the Drag Operation
            tds.DoDragDrop(True)
            # Refresh the File Windows in case something was moved
            self.OnRefresh(event)

    def RefreshDirCtrl(self, lbl, dirCtrl):
        """ Refresh a Directory Control, to update it if directories may have been added or removed """
        # If we are working with a LOCAL directory control ...
        if lbl == self.LOCAL_LABEL:
            # ... get the current LOCAL path
            path = dirCtrl.GetPath()
            # ... get the TREE from the control
            treeCtrl = dirCtrl.GetTreeCtrl()
            # Recreate the entire Directory Tree from scratch
            dirCtrl.ReCreateTree()
            # Re-Select the original PATH
            dirCtrl.SetPath(path)
            # Expand the current selection, which is the current path
            treeCtrl.Expand(treeCtrl.GetSelection())
        # If we are working with a REMOTE (SRB) directory control ...
        else:
            # ... get the current REMOTE path (before we start rebuilding the control)
            path = self.GetFullPath(dirCtrl)
            # Read the list of folders from the SRB and put them in the SRB Dir Tree
            # Start by clearing the list
            dirCtrl.DeleteAllItems()
            # Add a Root Item to the SRB Dir Tree
            srbRootItem = dirCtrl.AddRoot(self.srbCollection)
            # Initiate the recursive call to get all Collections from the SRB and add them to the SRB Dir Tree
            self.srbAddChildNode(dirCtrl, srbRootItem, self.tmpCollection)
            # Expand the previously selected item
            dirCtrl.Expand(srbRootItem)

            # Now look for the correct path, so we can expand it
            # Initialize a pointer to the position in the Node List
            nodeListPos = 0
            # Remove the slash that is probably at the end of the path
            if path[-1] == '/':
                path = path[:-1]
            # Create the list of nodes we need to move up.  Remove SRB basic info (/WCER/home/dwoods.digital-insight) 
            # from the start of the path.
            nodeList = path.split('/')[4:]
            # Start at the Root item
            currentNode = dirCtrl.GetRootItem()
            # For each node ...
            for node in nodeList:
                # ... initialize that we're not done yet
                notDone = True
                # Get the current node's first child
                (childNode, cookie) = dirCtrl.GetFirstChild(currentNode)
                # While we haven't run out of nodes and haven't found the correct node ...
                while notDone and (currentNode != None):
                    # ... see if THIS is the correct node by comparing node text to the correct position in
                    # the node list.  If so (climbing up) ....
                    if childNode.IsOk() and dirCtrl.GetItemText(childNode) == nodeList[nodeListPos]:
                        # ... increment the node counter ...
                        nodeListPos += 1
                        # ... set the current node to this node ...
                        currentNode = childNode
                        # ... and signal that we're done with this level
                        notDone = False
                    # Otherwise, see if we're at the end of the child nodes.  If so ...
                    elif childNode == dirCtrl.GetLastChild(currentNode):
                        # ... Set currentNode to None to signal failure ...
                        currentNode = None
                        # ... and signal that we're done with this level
                        notDone = False
                    # If we're not done yet ...
                    if notDone:
                        # ... then get the next child node for this level
                        (childNode, cookie) = dirCtrl.GetNextChild(currentNode, cookie)

            # If we have found the node we're looking for ...
            if currentNode != None:
                # ... select that node
                dirCtrl.SelectItem(currentNode)
                # ... make sure the user can see it on screen
                dirCtrl.EnsureVisible(currentNode)
                # ... and expand it to show its subdirectories, if there are any
                dirCtrl.Expand(currentNode)

    def GetSRBDirectoryList(self, startDir):
        """ Get a list of all directories within a SRB collection """
        #  srbDaiMcatQuerySubColls runs a query on the SRB server and returns a number of records.  Not the total
        #  number of records, perhaps, but the number it can read into a data buffer.  According to Bing Zhu at
        #  SDSC, it is about 200 records at a time.
        #
        #  Those records can then be accessed by using srbDaiMcatGetColl.
        #
        #  Once all the records in the buffer are read, srbDaiMcatGetMoreColls is called, telling the SRB server
        #  to fill the buffer with the next set of records.  Only when this function returns a value of 0 have all
        #  the records been returned.
        #
        #  Also, this function cannot be called recursively, because the SRB loses it's place when backing out of the
        #  recursion.  Therefore, I will use a List structure to pull all the Collection Names out at once, then call
        #  this method recursively based on that list's items.  In other words, we need to exhaust a given srbDaiMcatQuerySubColls
        #  call before we move on to the next recursion.

        # Set a temporary varaible for the starting collection, encoding it if needed
        tmpCollection = startDir

        if ('unicode' in wx.PlatformInfo) and isinstance(tmpCollection, unicode):
            tmpCollection = startDir.encode(TransanaGlobal.encoding)

        # Start with an empty Node List and an empty list of values to return
        nodeList = []
        returnList = []
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
        # Get the list of sub-Collections from the SRB
        srbCatalogCount = srb.srb_query_subcolls(self.srbConnectionID.value, 0, tmpCollection)
        # Check to see if Sub-Collections are returned or if an error code (negative result) is generated
        if srbCatalogCount < 0:
            # Report an error, if found
            self.srbErrorMessage(srbCatalogCount)
        # If no error occurred ...
        else:
            # Process the partial list and request more data from the SRB as long as additional data is forthcoming.
            # (srbCatalogCount > 0 indicates there is more data, while a value of 0 indicates all data has been sent.)
            while srbCatalogCount > 0:
                # Loop through all data in the partial list that has been returned from the SRB so far
                for loop in range(srbCatalogCount):
                    # Create a 1k buffer of spaces to receive data from the SRB
                    buf = ' ' * 1024
                    # Get data for the next Collection in the partial list (indicated by "loop")
                    srb.srb_get_coll_data(loop, buf, len(buf))

                    # Strip whitespace from the buffer
                    tempStr = string.strip(buf)
                    # Remove the #0 terminator character from the end of the string
                    if ord(tempStr[-1]) == 0:
                        tempStr = tempStr[:-1]
               
                    # Add the full path, not just the last part of it, to the nodeList for further processing
                    nodeList.append(tempStr)

                # When all elements of a partial list have been processed, request the next chunk of the list.
                # If the entire list of Collections has been sent, the SRB function will return 0
                srbCatalogCount = srb.srb_get_more_subcolls(self.srbConnectionID.value)

                # Check to see if Sub-Collections are returned or if an error code (negative result) is generated
                if srbCatalogCount < 0:
                    self.srbErrorMessage(srbCatalogCount)

            # For each item in the node list ...
            for item in nodeList:
                # ... add it to the Return List
                returnList.append(item)
                # Call this method recursively to get the subdirectories of each directory!
                returnList += self.GetSRBDirectoryList(item)
                
        # Return the results
        return returnList

    def GetSRBFileList(self, path):
        """ Get a list of all files within a SRB collection """
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

        #  srbDaiMcatQueryChildDataset runs a query on the SRB server and returns a number of records.  Not the total
        #  number of records, perhaps, but the number it can read into a data buffer.  According to Bing Zhu at
        #  SDSC, it is about 200 records at a time.
        #
        #  Those records can then be accessed by using srbDaiMcatGetData.
        #
        #  Once all the records in the buffer are read, srbDaiMcatGetMoreData is called, telling the SRB server
        #  to fill the buffer with the next set of records.  Only when this function returns a value of 0 have all
        #  the records been returned.
         
        # Create an empty File List
        filelist = []
        # If the calling path ends with a slash, remove that character (which interferes with SRB calls)
        if path[-1] == '/':
            path = path[:-1]
        # Encode the path
        if ('unicode' in wx.PlatformInfo) and isinstance(path, unicode):
            path = path.encode(TransanaGlobal.encoding)

        # Now we retrieve the list of files from the SRB.
        # Get the first batch of File Names from the SRB
        srbCatalogCount = srb.srb_query_child_dataset(self.srbConnectionID.value, 0, path)

        # Check for errors in getting the File Names
        if srbCatalogCount < 0:
            # Report the error, it present
            self.srbErrorMessage(srbCatalogCount)
        # If there's no error ...
        else:
            # As long as there are Data objects in the SRB's Data Buffer...
            while srbCatalogCount > 0:
                # Iterate through the items in the Buffer by Item Number
                for loop in range(srbCatalogCount):
                    # Create an string full of spaces to recieve data from the SRB
                    buf = ' ' * 1024
                    # Get the next Data Record (File Name) from the SRB (result stored in "buf")
                    srb.srb_get_dataset_system_metadata(0, loop, buf, len(buf))
                    # Strip extraneous whitespace (a chr(0) remains on the end of the file name)
                    tempStr = string.strip(buf)
                    if ord(tempStr[-1]) == 0:
                        tempStr = tempStr[:-1]
                    # Add the result to the File List
                    filelist.append(tempStr)

                # Make another call to the SRB to see if there are more records to read into the SRB's Data Buffer
                srbCatalogCount = srb.srb_get_more_subcolls(self.srbConnectionID.value)

        # Return the results
        return filelist

    def FileCopyMoveLocalToLocal(self, moveFlag, sourceDir, destDir, fileName):
        """ Copy or Move a FILE from LOCAL to LOCAL """
        # Build the full File Name by joining the source path with the file name from the list
        sourceFile = os.path.join(sourceDir, fileName)
        # If we want to MOVE the file ...
        if moveFlag:
            # Create the Status Prompt
            prompt = _('Moving %s to %s')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Display the prompt
            self.SetStatusText(prompt % (sourceFile, destDir))
            # Start Exception Handling
            try:
                # Check if the destination file already exists.  If NOT ...
                if not os.path.exists(os.path.join(destDir, fileName)):
                    # ... os.Rename accomplishes a fast LOCAL Move
                    os.rename(sourceFile, os.path.join(destDir, fileName))
                # If the file DOES exist ...
                else:
                    # Create an error message
                    errMsg = _('File "%s" could not be moved.') + '\n' + _('File "%s" already exists.')
                    if 'unicode' in wx.PlatformInfo:
                        errMsg = unicode(errMsg, 'utf8')
                    # Display the error message
                    dlg = Dialogs.ErrorDialog(self, errMsg % (sourceFile, os.path.join(destDir, fileName)))
                    dlg.Show()  # NOT Modal, and no Destroy prevents interruption
            # Process exceptions
            except:
                # Create an error message
                errMsg = _('File "%s" could not be moved.')
                if 'unicode' in wx.PlatformInfo:
                    errMsg = unicode(errMsg, 'utf8')
                # Display the error message
                dlg = Dialogs.ErrorDialog(self, errMsg % fileName)
                dlg.Show()  # NOT Modal, and no Destroy prevents interruption
            # Update the status to indicate the move is done
            self.SetStatusText(_('Move complete.'))

        # ... otherwise we want to copy the file
        else:
            # Create the Status Prompt
            prompt = _('Copying %s to %s')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Display the prompt
            self.SetStatusText(prompt % (sourceFile, destDir))
            # Check if the destination file already exists.  If NOT ...
            if not os.path.exists(os.path.join(destDir, fileName)):
                # shutil is the most efficient way I have found to copy a file on the local file system
                shutil.copyfile(sourceFile, os.path.join(destDir, fileName))
            # If the file DOES exist ...
            else:
                # Create an error message
                errMsg = _('File "%s" could not be copied.') + '\n' + _('File "%s" already exists.')
                if 'unicode' in wx.PlatformInfo:
                    errMsg = unicode(errMsg, 'utf8')
                # Display the error message
                dlg = Dialogs.ErrorDialog(self, errMsg % (sourceFile, os.path.join(destDir, fileName)))
                dlg.Show()  # NOT Modal, and no Destroy prevents interruption
            # Update the status to indicate the copy is done
            self.SetStatusText(_('Copy complete.'))

    def FileCopyMoveLocalToRemote(self, moveFlag, sourceDir, destDir, fileName):
        """ Copy or Move a FILE from LOCAL to REMOTE (SRB) """
        # Build the full File Name by joining the source path with the file name from the list
        sourceFile = os.path.join(sourceDir, fileName)
        # Get the File Size
        fileSize = os.path.getsize(sourceFile)
        # If we want to MOVE the file ...
        if moveFlag:
            # Create the Status Prompt
            prompt = _('Moving %s to %s')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Display the prompt
            self.SetStatusText(prompt % (sourceFile, destDir))
        # ... otherwise we want to copy the file
        else:
            # Create the Status Prompt
            prompt = _('Copying %s to %s')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Display the prompt
            self.SetStatusText(prompt % (sourceFile, destDir))
        # The SRBFileTransfer class handles file transfers and provides Progress Feedback
        dlg = SRBFileTransfer.SRBFileTransfer(self, _("SRB File Transfer"), fileName, fileSize, sourceDir, self.srbConnectionID, destDir, SRBFileTransfer.srb_UPLOAD, self.srbBuffer)
        success = dlg.TransferSuccessful()
        dlg.Destroy()
        # Report the outcome in the status bar.  If we MOVED ...
        if moveFlag:
            # ... and succeeded ...
            if success:
                # ... delete the SOURCE file to complete the MOVE ...
                self.DeleteFile(self.LOCAL_LABEL, sourceDir, fileName)
                # ... and report success
                self.SetStatusText(_('Move complete.'))
            # ... and failed ...
            else:
                # ... report the failure
                self.SetStatusText(_('Move cancelled.'))
        # If we COPIED ...
        else:
            # ... and succeeded ...
            if success:
                # ... report the success
                self.SetStatusText(_('Copy complete.'))
            # ... and failed ...
            else:
                # ... report the failure
                self.SetStatusText(_('Copy cancelled.'))
        # Return the result
        return success

    def FileCopyMoveRemoteToLocal(self, moveFlag, sourceDir, destDir, fileName, sourceFileIndex):
        """ Copy or Move a FILE from REMOTE (SRB) to LOCAL """
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
        # Build the full File Name by joining the source path with the file name from the list
        sourceFile = os.path.join(sourceDir, fileName)
        # If we want to MOVE the file ...
        if moveFlag:
            # Create the Status Prompt
            prompt = _('Moving %s to %s')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Display the prompt
            self.SetStatusText(prompt % (sourceFile, destDir))
        # ... otherwise we want to copy the file
        else:
            # Create the Status Prompt
            prompt = _('Copying %s to %s')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Display the prompt
            self.SetStatusText(prompt % (sourceFile, destDir))

        # Check if the destination file already exists.  If NOT ...
        if not os.path.exists(os.path.join(destDir, fileName)):
            # Find the File Size
            buf = ' ' * 50
            srb.srb_get_dataset_system_metadata(2, sourceFileIndex, buf, 50)
            # Strip whitespace and null character (c string terminator) from buf
            buf = string.strip(buf)
            # Strip the #0 null character off the buffer, if needed
            if ord(buf[-1]) == 0:
                buf = buf[:-1]
            # The SRBFileTransfer class handles file transfers and provides Progress Feedback
            dlg = SRBFileTransfer.SRBFileTransfer(self, _("SRB File Transfer"), fileName, int(buf), destDir, self.srbConnectionID, sourceDir, SRBFileTransfer.srb_DOWNLOAD, self.srbBuffer)
            success = dlg.TransferSuccessful()
            dlg.Destroy()
        # If the file DOES exist ...
        else:
            # ... indicate that the copy failed.
            success = False
            # Create an error message
            errMsg = _('File "%s" could not be copied.') + '\n' + _('File "%s" already exists.')
            if 'unicode' in wx.PlatformInfo:
                errMsg = unicode(errMsg, 'utf8')
            # Display the error message
            dlg = Dialogs.ErrorDialog(self, errMsg % (sourceFile, os.path.join(destDir, fileName)))
            dlg.Show()  # NOT Modal, and no Destroy prevents interruption
        # Report the outcome in the status bar.  If we MOVED ...
        if moveFlag:
            # ... and succeeded ...
            if success:
                # ... delete the SOURCE file to complete the MOVE ...
                self.DeleteFile(self.REMOTE_LABEL, sourceDir, fileName)
                # ... and report success
                self.SetStatusText(_('Move complete.'))
            # ... and failed ...
            else:
                # ... report the failure
                self.SetStatusText(_('Move cancelled.'))
        # If we COPIED ...
        else:
            # ... and succeeded ...
            if success:
                # ... report the success
                self.SetStatusText(_('Copy complete.'))
            # ... and failed ...
            else:
                # ... report the failure
                self.SetStatusText(_('Copy cancelled.'))
        # If the COPY/MOVE failed because the destination already exists ...
        if os.path.exists(os.path.join(destDir, fileName)):
            # ... we don't need to interrupt the whole copy/move process
            success = True
        # Return the result
        return success

    def FileCopyMoveRemoteToRemote(self, moveFlag, sourceDir, destDir, fileName):
        """ Copy or Move a FILE from REMOTE (SRB) to REMOTE (SRB) """
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
        # Build the full File Name by joining the source path with the file name from the list
        sourceFile = sourceDir + u'/' + fileName
        # If we want to MOVE the file ...
        if moveFlag:
            # Create the Status Prompt
            prompt = _('Moving %s to %s')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Display the prompt
            self.SetStatusText(prompt % (sourceFile, destDir))
        # ... otherwise we want to copy the file
        else:
            # Create the Status Prompt
            prompt = _('Copying %s to %s')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            if isinstance(sourceFile, str):
                sourceFile = sourceFile.decode(TransanaGlobal.encoding)
            if isinstance(destDir, str):
                destDir = destDir.decode(TransanaGlobal.encoding)
            # Display the prompt
            self.SetStatusText(prompt % (sourceFile, destDir))

        # Get temporary version of file name and directories
        tmpFile = fileName
        tmpSourceDir = sourceDir
        tmpTargetDir = destDir
        # Encode them if necessary
        if 'unicode' in wx.PlatformInfo:
            tmpFile = tmpFile.encode(TransanaGlobal.encoding)
            tmpSourceDir = tmpSourceDir.encode(TransanaGlobal.encoding)
            tmpTargetDir = tmpTargetDir.encode(TransanaGlobal.encoding)

        # Copy the SRB files
        Result = srb.srb_obj_to_coll_copy(  
                    self.srbConnectionID,    # Connection ID
                    0,                       # Catalog Type - Use 0
                    tmpFile,                 # Source File name
                    0,                       # "replNum" - Use 0. Transana does not support replications
                    tmpSourceDir,            # Source Collection
                    'unknown',               # File Type - "unknown" works
                    0,                       # Size - ??
                    tmpTargetDir,            # Destination Collection
                    self.tmpResource,        # Resource (could be different than old Resource!)  Transana does not support alternate Resource Specification at this time!
                    '');                     # Container.  Transana does not use Containers.
        # If there was an error ...
        if Result < 0:
            # ... report the error message
            self.srbErrorMessage(Result)
        # Report the outcome in the status bar.  If we MOVED ...
        if moveFlag:
            # If we succeeded ...
            if Result >= 0 :
                # ... delete the SOURCE file to complete the MOVE ...
                self.DeleteFile(self.REMOTE_LABEL, sourceDir, fileName)
                # ... and report success
                self.SetStatusText(_('Move complete.'))
            # ... and failed ...
            else:
                # ... report the failure
                self.SetStatusText(_('Move failed.'))
        # If we COPIED ...
        else:
            # ... and succeeded ...
            if Result >= 0:
                # ... report the success
                self.SetStatusText(_('Copy complete.'))
            # ... and failed ...
            else:
                # ... report the failure
                self.SetStatusText(_('Copy failed.'))

    def OnCopyMove(self, event):
        """ Copy or Move files between sides of the File Management Window """
        # If the Move button called this event, we have a Move request.  Otherwise we
        # have a Copy request.  This is signalled by moveFlag.
        if event.GetId() in [self.btnMoveToLeft.GetId(), self.btnMoveToRight.GetId()]:
            moveFlag = True
        else:
            moveFlag = False
        # Map Window widgets to the appropriate local objects for sending and receiving files.
        # If we're going from Left to Right ...
        if event.GetId() in [self.btnCopyToRight.GetId(), self.btnMoveToRight.GetId()]:
            # ... get the SOURCE from the LEFT side
            sourceLbl = self.lblLeft.GetLabel()
            # Local and Remote File Systems have different methods for determing the selected Path
            if sourceLbl == self.LOCAL_LABEL:
                sourceDir = self.dirLeft.GetPath()
            else:
                sourceDir = self.GetFullPath(self.srbDirLeft)
            sourceFile = self.fileLeft
            sourceFilter = self.filterLeft
            # ... and get the DEST from the RIGHT side
            destLbl = self.lblRight.GetLabel()
            # Local and Remote File Systems have different methods for determing the selected Path
            if destLbl == self.LOCAL_LABEL:
                destDir = self.dirRight.GetPath()
            else:
                destDir = self.GetFullPath(self.srbDirRight)
            destFile = self.fileRight
            destFilter = self.filterRight
        # If we're going from Right to Left ...
        elif event.GetId() in [self.btnCopyToLeft.GetId(), self.btnMoveToLeft.GetId()]:
            # ... get the SOURCE from the RIGHT side
            sourceLbl = self.lblRight.GetLabel()
            # Local and Remote File Systems have different methods for determing the selected Path
            if sourceLbl == self.LOCAL_LABEL:
                sourceDir = self.dirRight.GetPath()
            else:
                sourceDir = self.GetFullPath(self.srbDirRight)
            sourceFile = self.fileRight
            sourceFilter = self.filterRight
            # ... and get the DEST from the LEFT side
            destLbl = self.lblLeft.GetLabel()
            # Local and Remote File Systems have different methods for determing the selected Path
            if destLbl == self.LOCAL_LABEL:
                destDir = self.dirLeft.GetPath()
            else:
                destDir = self.GetFullPath(self.srbDirLeft)
            destFile = self.fileLeft
            destFilter = self.filterLeft

        # Build a list of files to be copied
        # Initialise the fileList
        fileList = []

        # If no items are selected in the Source File List, copy ALL FILES in the Source Directory.
        # This can be signalled by setting the sourceFile.GetNextItem "state" parameter to either
        # "wxLIST_STATE_DONTCARE" (all files) or "wxLIST_STATE_SELECTED" (only selected files).
        if sourceFile.GetSelectedItemCount() == 0:
            state = wx.LIST_STATE_DONTCARE
        else:
            state = wx.LIST_STATE_SELECTED
        # Get all Selected files from the File List.  Starting with itemNum = -1 signals to search the whole list.
        itemNum = -1
        # Keep looking until forced to stop
        while True:
            # Get the next Selected item in the List 
            itemNum = sourceFile.GetNextItem(itemNum, wx.LIST_NEXT_ALL, state)
            # If there are no more selected items in the list, itemNum is set to -1
            if itemNum == -1:
                # ... so we should stop looking
                break
            # If there are more items ...
            else:
                # add the Selected item to the fileList
                fileList.append(sourceFile.GetItemText(itemNum))

        # Set the Cursor to the Hourglass
        self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

        # Copy or Move from Local Drive to Local Drive
        if (sourceLbl == self.LOCAL_LABEL) and \
           (destLbl == self.LOCAL_LABEL):
            # For each file in the fileList ...
            for fileNm in fileList:
                # ... call the Local to Local File Copy routine
                self.FileCopyMoveLocalToLocal(moveFlag, sourceDir, destDir, fileNm)
                # Let's update the dest control after every file so that the new files show up in the list ASAP.
                self.RefreshFileList(destLbl, destDir, destFile, destFilter)
                # wxYield allows the Windows Message Queue to update the display.
                wx.Yield()

        # Copy or Move from Local Drive to SRB Collection
        elif (sourceLbl == self.LOCAL_LABEL) and \
             (destLbl == self.REMOTE_LABEL):
            # For each file in the fileList ...
            for fileNm in fileList:
                # See if the file is already in the destination list.  If NOT ...
                if destFile.FindItem(-1, fileNm) == -1:
                    # ... call the Local to Remote File Copy routine, capturing the result
                    success = self.FileCopyMoveLocalToRemote(moveFlag, sourceDir, destDir, fileNm)
                    # Let's update the dest control after every file so that the new files show up in the list ASAP.
                    self.RefreshFileList(destLbl, destDir, destFile, destFilter)
                    # wxYield allows the Windows Message Queue to update the display.
                    wx.Yield()
                    # If the Copy / Move failed or was cancelled by the user ...
                    if not success:
                        # ... we should stop trying to copy / move files
                        break

        # Copy or Move from SRB Collection to Local Drive
        elif (sourceLbl == self.REMOTE_LABEL) and \
             (destLbl == self.LOCAL_LABEL):
            # For each file in the fileList ...
            for fileNm in fileList:
                # ... call the Remote to Local File Copy routine, capturing the result
                success = self.FileCopyMoveRemoteToLocal(moveFlag, sourceDir, destDir, fileNm, sourceFile.FindItem(-1, fileNm))
                # Let's update the dest control after every file so that the new files show up in the list ASAP.
                self.RefreshFileList(destLbl, destDir, destFile, destFilter)
                # wxYield allows the Windows Message Queue to update the display.
                wx.Yield()
                # If the Copy / Move failed or was cancelled by the user ...
                if not success:
                    # ... we should stop trying to copy / move files
                    break

        # Copy or Move from SRB Collection to SRB Collection
        elif (sourceLbl == self.REMOTE_LABEL) and \
             (destLbl == self.REMOTE_LABEL):
            # For each file in the fileList ...
            for fileNm in fileList:
                # ... call the Remote to Remote File Copy routine
                self.FileCopyMoveRemoteToRemote(moveFlag, sourceDir, destDir, fileNm)
                # Let's update the dest control after every file so that the new files show up in the list ASAP.
                self.RefreshFileList(destLbl, destDir, destFile, destFilter)
                # wxYield allows the Windows Message Queue to update the display.
                wx.Yield()

        # Reset the cursor to the Arrow
        self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

        # If we have moved files, we need to refresh the Source file list so that the files that were moved
        # no longer appear in the list.
        if moveFlag:
            self.RefreshFileList(sourceLbl, sourceDir, sourceFile, sourceFilter)

    def OnSynch(self, event):
        """  Process the Copy All New buttons, which copy any files that don't exist in the SOURCE tree to the
             DEST tree, recursing through subdirectories  """
        # Set the Cursor to the Hourglass
        self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

        # If we are copying from Left to Right ...
        if event.GetId() == self.btnSynchToRight.GetId():
            # ... our SOURCE comes from the LEFT ...
            sourceLbl = self.lblLeft.GetLabel()
            # If our SOURCE is LOCAL ...
            if sourceLbl == self.LOCAL_LABEL:
                # ... we need  the Directory LIst, the Path, and the Tree
                sourceDir = self.dirLeft
                sourcePath = self.dirLeft.GetPath()
                sourceTree = sourceDir.GetTreeCtrl()
            # If our SOURCE is REMOTE ...
            else:
                # ... we need  the Directory LIst, the Path, and the Tree
                sourceDir = self.srbDirLeft
                sourcePath = self.GetFullPath(self.srbDirLeft)
                sourceTree = sourceDir
            # ... and our DEST goes to the RIGHT
            destLbl = self.lblRight.GetLabel()
            # If our DEST is LOCAL ...
            if destLbl == self.LOCAL_LABEL:
                # ... we need  the Directory LIst, the Path, and the Tree
                destDir = self.dirRight
                destPath = self.dirRight.GetPath()
                destTree = destDir.GetTreeCtrl()
            # If our DEST is REMOTE ...
            else:
                # ... we need  the Directory LIst, the Path, and the Tree
                destDir = self.srbDirRight
                destPath = self.GetFullPath(self.srbDirRight)
                destTree = destDir
            # We also need the DEST File Control and the DEST Filter Control
            destFile = self.fileRight
            destFilter = self.filterRight
        # If we are copying from Right to Left ...
        else:
            # ... our SOURCE comes from the RIGHT ...
            sourceLbl = self.lblRight.GetLabel()
            # If our SOURCE is LOCAL ...
            if sourceLbl == self.LOCAL_LABEL:
                # ... we need  the Directory LIst, the Path, and the Tree
                sourceDir = self.dirRight
                sourcePath = self.dirRight.GetPath()
                sourceTree = sourceDir.GetTreeCtrl()
            # If our SOURCE is REMOTE ...
            else:
                # ... we need  the Directory LIst, the Path, and the Tree
                sourceDir = self.srbDirRight
                sourcePath = self.GetFullPath(self.srbDirRight)
                sourceTree = sourceDir
            # ... and our DEST goes to the LEFT
            destLbl = self.lblLeft.GetLabel()
            # If our DEST is LOCAL ...
            if destLbl == self.LOCAL_LABEL:
                # ... we need  the Directory LIst, the Path, and the Tree
                destDir = self.dirLeft
                destPath = self.dirLeft.GetPath()
                destTree = destDir.GetTreeCtrl()
            # If our DEST is REMOTE ...
            else:
                # ... we need  the Directory LIst, the Path, and the Tree
                destDir = self.srbDirLeft
                destPath = self.GetFullPath(self.srbDirLeft)
                destTree = destDir
            # We also need the DEST File Control and the DEST Filter Control
            destFile = self.fileLeft
            destFilter = self.filterLeft

        # Now COPY the FILES.
        # If our SOURCE is LOCAL ...
        if sourceLbl == self.LOCAL_LABEL:
            # if our DEST is REMOTE ...
            if destLbl != self.LOCAL_LABEL:
                # ... then we need a list of the DEST Collections on the SRB
                srbCollList = self.GetSRBDirectoryList(destPath)

                print "Directory Comparison:"
                for td in srbCollList:
                    print td, type(td)
                print

                # Initialize the list of SRF Files
                srbFileList = []
                # Get a list of files from the DEST SRB 
                tmpSRBFileList = self.GetSRBFileList(destPath)
                # For each file in the file list ...
                for f in tmpSRBFileList:

                    tmpDestDir = destPath
                    if isinstance(tmpDestDir, str):
                        tmpDestDir = tmpDestDir.decode(TransanaGlobal.encoding)
                    tmpF = f
                    if isinstance(tmpF, str):
                        tmpF = tmpF.decode(TransanaGlobal.encoding)
                    # Create a temporary file name
                    tmpFile = tmpDestDir + '/' + tmpF
                    # Encode if needed
                    if 'unicode' in wx.PlatformInfo:
                        tmpFile = tmpFile.encode(TransanaGlobal.encoding)
                    # ... add the path and file name to the list of files on the SRB
                    srbFileList.append(tmpFile)

            # Initialize that the user has NOT pressed CANCEL
            cancelPressed = False
            # Get all directory and file names from the LOCAL SOURCE
            for (path, dirs, files) in os.walk(sourcePath):
                # If the user has pressed CANCEL ...
                if cancelPressed:
                    # ... then we need to stop iterating through the local directory / file list
                    break
                # For each LOCAL SOURCE Directory ...
                for d in dirs:
                    # If our DEST is LOCAL ...
                    if destLbl == self.LOCAL_LABEL:
                        # ... if the DEST Directory does not exist ...
                        if not os.path.exists(os.path.join(destPath, path[len(sourcePath) + 1:], d)):
                            # ... then we need to create the DEST Directory
                            os.mkdir(os.path.join(destPath, path[len(sourcePath) + 1:], d))
                    # If our DEST is REMOTE ....
                    else:
                        # Make sure our DEST Path ends with a slash
                        if destPath[-1] != '/':
                            destPath += '/'
                        # Make sure the path ends with a slash
                        if path[-1] != os.sep:
                            path += os.sep
                        # Create a temporary path from the Destination and the local path, replacing backslashes if needed
                        tmpPath = destPath + path[len(sourcePath) + 1:].replace(os.sep, '/')
                        if d[-1] == '/':
                            d = tmpPath[:-1]
                        tmpDir = tmpPath + d
                        if 'unicode' in wx.PlatformInfo:
                            tmpDir = tmpDir.encode(TransanaGlobal.encoding)

                        print tmpDir, type(tmpDir), tmpDir in srbCollList
                        print
                        print
                        for td in srbCollList:
                            if (td.decode('utf8') == tmpDir.decode('utf8')) and (not tmpDir in srbCollList):
                                print "DECODE WORKS!!!!!"
                            else:
                                print "Compare:"
                                print td
                                print tmpDir, td.decode('utf8') == tmpDir.decode('utf8')
                                print
                        
                        # ... if the DEST Directory does not exist (is not in the SRB Collection List) ...
                        if not tmpDir in srbCollList:
                            # ... Create temporary folder and path variables
                            tmpNewFolder = d
                            # Encode if needed
                            if 'unicode' in wx.PlatformInfo:
                                tmpNewFolder = tmpNewFolder.encode(TransanaGlobal.encoding)
                                tmpPath = tmpPath.encode(TransanaGlobal.encoding)
                            # We need to drop the ending slash from the path if there is one
                            if tmpPath[-1] == "/":
                                tmpPath = tmpPath[:-1]
                            # Create the folder on the SRB file system
                            newCollection = srb.srb_new_collection(self.srbConnectionID, 0, tmpPath, tmpNewFolder);
                            # If the SRB returns an error message ...
                            if newCollection < 0:
                               # ... display the error message
                               self.srbErrorMessage(newCollection)
                        # if the DEST SRB Collection exists ...
                        else:
                            # We need to drop the slash from the path we added above
                            if tmpPath[-1] != "/":
                                tmpPath += '/'
#                            if isinstance(tmpPath, unicode):
#                                tmpPath = tmpPath.decode(TransanaGlobal.encoding)
#                            if isinstance(d, unicode):
#                                d = d.decode(TransanaGlobal.encoding)
                            # Get the files from this SRB Collection, since it exists
                            tmpSRBFileList = self.GetSRBFileList(tmpPath + d)
                            # For each file in the file list ...
                            for f in tmpSRBFileList:

                                if isinstance(f, str):
                                    f = f.decode(TransanaGlobal.encoding)
                                
                                # Create a temporary file name
#                                tmpFile = tmpPath + d + '/' + f
                                # Encode if needed
                                if 'unicode' in wx.PlatformInfo:
                                    tmpFile = tmpPath + d + '/' + f



                                    
                                    tmpFile = tmpFile.encode(TransanaGlobal.encoding)
                                # ... add the path and file name to the list of files on the SRB
                                srbFileList.append(tmpFile)

                # For each LOCAL SOURCE file
                for f in files:
                    # If the user has pressed CANCEL ...
                    if cancelPressed:
                        # ... then we need to stop iterating through the local file list
                        break
                    # If our DEST is LOCAL ...
                    if destLbl == self.LOCAL_LABEL:
                        # ... if the DEST File does not exist ...
                        if not os.path.exists(os.path.join(destPath, path[len(sourcePath) + 1:], f)):
                            # ... then we need to copy the File
                            self.FileCopyMoveLocalToLocal(False, path, os.path.join(destPath, path[len(sourcePath) + 1:]), f)
                    # If our DEST is REMOTE ....
                    else:
                        # Create a temporary file name
                        tmpFile = os.path.join(destPath, path[len(sourcePath) + 1:], f).replace(os.sep, '/')
                        if 'unicode' in wx.PlatformInfo:
                            tmpFile = tmpFile.encode(TransanaGlobal.encoding)
                        # ... if the DEST File does not exist (is not in the SRB File List) ...
                        if not tmpFile in srbFileList:
                            # Build a temporary path from the DestPath and the local current directory, replacing backslashes if needed 
                            tmpDestPath = os.path.join(destPath, path[len(sourcePath) + 1:]).replace(os.sep, '/')
                            # We need to drop the ending slash from the path if there is one
                            if tmpDestPath[-1] == '/':
                                tmpDestPath = tmpDestPath[:-1]
                            # ... then we need to copy the File, capturing the result
                            success = self.FileCopyMoveLocalToRemote(False, path, tmpDestPath, f)
                            # If the result is not True, the user probably pressed Cancel.
                            if not success:
                                # Indicate that Cancel was pressed so outer loop processing will stop
                                cancelPressed = True
                                # ... and stop processing for the Files loop
                                break

        # if our SOURCE is REMOTE ...
        else:
            # ... then we need a list of the SOURCE Collections on the SRB
            sourceSRBCollList = self.GetSRBDirectoryList(sourcePath)
            # Add the CURRENT directory to the Collections List!  Otherwise, files from the root level don't get included
            # in the SRB File List
            tmpPath = sourcePath
            if 'unicode' in wx.PlatformInfo:
                tmpPath = tmpPath.encode(TransanaGlobal.encoding)
            sourceSRBCollList = [tmpPath] + sourceSRBCollList

            # If our DESTINATION is REMOTE ...
            if destLbl != self.LOCAL_LABEL:
                # ... then we need a list of the DEST Collections on the SRB
                destSRBCollList = self.GetSRBDirectoryList(destPath)
                tmpPath = destPath
                if 'unicode' in wx.PlatformInfo:
                    tmpPath = tmpPath.encode(TransanaGlobal.encoding)
                # Add the CURRENT directory to the Collections List!  Otherwise, files from the root level don't get included
                # in the SRB File List
                destSRBCollList = [tmpPath] + destSRBCollList

            # Initialize that the user has NOT pressed CANCEL
            cancelPressed = False
            # For each Collection in the SOURCE SRB Collection List ...
            for d in sourceSRBCollList:
                # If the user has pressed CANCEL ...
                if cancelPressed:
                    # ... then we need to stop iterating through the remote directory list
                    break
                # Get the SOURCE File list from the SRB
                sourceSRBFiles = self.GetSRBFileList(d)
                # If the DEST is LOCAL ...
                if destLbl == self.LOCAL_LABEL:
                    # ... build the appropriate directory entry for the current DEST directory
                    tmpDestDir = d[len(sourcePath) + 1:].replace('/', os.sep)
                    if len(tmpDestDir) > 0 and tmpDestDir[0] == os.sep:
                        tmpDestDir = tmpDestDir[1:]
                    if isinstance(tmpDestDir, str):
                        tmpDestDir = tmpDestDir.decode(TransanaGlobal.encoding)
                    # ... if the DEST Directory does not exist ...
                    if not os.path.exists(os.path.join(destPath, tmpDestDir)):
                        # ... then we need to create the DEST Directory
                        os.makedirs(os.path.join(destPath, tmpDestDir))
                # If the DEST is REMOTE ...
                else:
                    # ... intialize the DEST SRB File List
                    destSRBFileList = []
                    tmpDir = d
                    if isinstance(tmpDir, str):
                        tmpDir = tmpDir.decode(TransanaGlobal.encoding)
                    
                    # Build the appropriate DEST path
                    tmpDestDir = destPath + tmpDir[len(sourcePath):]
                    if 'unicode' in wx.PlatformInfo:
                        tmpDestDir = tmpDestDir.encode(TransanaGlobal.encoding)

                    # If the DEST Path does NOT exist (is not in the DEST SRB Collection List) ...
                    if not tmpDestDir in destSRBCollList:
                        # ... create temporary folder and path variables
                        tmpNewFolder = tmpDestDir[tmpDestDir.rfind('/') + 1:]
                        tmpPath = tmpDestDir[:tmpDestDir.rfind('/')]
                        # Encode, if needed
#                        if 'unicode' in wx.PlatformInfo:
#                            tmpNewFolder = tmpNewFolder.encode(TransanaGlobal.encoding)
#                            tmpPath = tmpPath.encode(TransanaGlobal.encoding)
                        # We need to drop the ending slash if there is one
                        if tmpPath[-1] == "/":
                            tmpPath = tmpPath[:-1]
                        # Create the folder on the SRB file system
                        newCollection = srb.srb_new_collection(self.srbConnectionID, 0, tmpPath, tmpNewFolder);
                        # If the SRB returns an error message ...
                        if newCollection < 0:
                            # ... display the error message
                            self.srbErrorMessage(newCollection)
                    # If the DEST path DOES exist ...
                    else:
                        # Get the files from this SRB Collection
                        tmpSRBFileList = self.GetSRBFileList(tmpDestDir)
                        # For each file in the file list ...
                        for f in tmpSRBFileList:

                            if isinstance(tmpDestDir, str):
                                tmpDestDir = tmpDestDir.decode(TransanaGlobal.encoding)
                            if isinstance(f, str):
                                f = f.decode(TransanaGlobal.encoding)
                            tmpF = tmpDestDir + u'/' + f
                            if ('unicode' in wx.PlatformInfo) and isinstance(tmpF, unicode):
                                tmpF.encode(TransanaGlobal.encoding)
                            # ... add the path and file name to the list of files on the SRB
                            destSRBFileList.append(tmpF)

                # For each file in the SOURCE SRB File List ...
                for f in sourceSRBFiles:
                    # If the user has pressed CANCEL ...
                    if cancelPressed:
                        # ... then we need to stop iterating through the remote source file list
                        break

                    # If the DEST is LOCAL ...
                    if destLbl == self.LOCAL_LABEL:

                        tmpPath = destPath
                        if isinstance(tmpPath, unicode):
                            tmpPath = tmpPath.encode(TransanaGlobal.encoding)

                        if isinstance(tmpDestDir, unicode):
                            tmpDestDir = tmpDestDir.encode(TransanaGlobal.encoding)
                        tmpFile = os.path.join(tmpPath, tmpDestDir, f[f.rfind('/') + 1:])
                        if isinstance(tmpFile, str):
                            tmpFile = tmpFile.decode('utf8')

                        # ... if the DEST File does not exist ...
                        if not os.path.exists(tmpFile):

                            if isinstance(tmpDestDir, str):
                                tmpDestDir = tmpDestDir.decode(TransanaGlobal.encoding)
                                
                            tmpD = d
                            if isinstance(tmpD, str):
                                tmpD = tmpD.decode(TransanaGlobal.encoding)
                            tmpF = f[f.rfind('/') + 1:]
                            if isinstance(tmpF, str):
                                tmpF = tmpF.decode(TransanaGlobal.encoding)

                            # ... then we need to copy the File, capturing the result
                            # FileCopyMoveRemoteToLocal(moveFlag, sourceDir, destDir, fileName, sourceFileIndex)
                            success = self.FileCopyMoveRemoteToLocal(False, tmpD, os.path.join(destPath, tmpDestDir), tmpF, sourceSRBFiles.index(f))
                            # If the result is not True, the user probably pressed Cancel.
                            if not success:
                                # Indicate that Cancel was pressed so outer loop processing will stop
                                cancelPressed = True
                                # ... and stop processing for the Files loop
                                break
                    # If the DEST is REMOTE ...
                    else:
                        if isinstance(tmpDestDir, str):
                            tmpDestDir = tmpDestDir.decode(TransanaGlobal.encoding)
                        tmpF = f
                        if isinstance(tmpF, str):
                            tmpF = tmpF.decode(TransanaGlobal.encoding)
                        tmpFile = tmpDestDir + u'/' + tmpF
                        # ... if the DEST File does not exist ...
                        if not tmpFile in destSRBFileList:

                            if isinstance(d, str):
                                d = d.decode(TransanaGlobal.encoding)
                            if isinstance(f, str):
                                f = f.decode(TransanaGlobal.encoding)
                            # ... then we need to copy the File
                            self.FileCopyMoveRemoteToRemote(False, d, tmpDestDir, f)

        # If the SOURCE and DEST are both REMOTE, SOURCE DIR may need to be updated, 
        if sourceLbl == destLbl:
            # ... so update the SOURCE Directory Control
            self.RefreshDirCtrl(sourceLbl, sourceDir)
        # Update the DEST Directory Control
        self.RefreshDirCtrl(destLbl, destDir)
        # Refresh the file list now that all files have been deleted
        self.RefreshFileList(destLbl, destPath, destFile, destFilter)
        # Reset the cursor to the Arrow
        self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
       
    def DeleteFile(self, lbl, path, filename):
        """ Delete the specified file """
        # Set cursor to hourglass
        self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
        # Create the appropriate Prompt
        prompt = _('Deleting %s')
        # Encode if needed
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(prompt, 'utf8')
        # if the file is LOCAL ...
        if lbl == self.LOCAL_LABEL:
            # Get the full name of the file to be deleted by adding the file path to it
            fileName = os.path.join(path, filename)
            # Set the Status Text
            self.SetStatusText(prompt % fileName)
            # Start Exception Handling
            try:
                # Delete the file
                os.remove(fileName)
            # Handle Exceptions
            except:
                # ... by simply reporting them
                dlg = Dialogs.ErrorDialog(self, "%s" % sys.exc_info()[1])
                dlg.ShowModal()
                dlg.Destroy()
        # if the file is REMOTE (SRB) ...
        else:
            # Get the full name of the file to be deleted by adding the file path to it
            fileName = path + '/' + filename
            # Set the Status Text
            self.SetStatusText(prompt % fileName)
            # Create temporary variables
            tmpFilename = filename
            tmpPath = path
            # Encode if needed
            if 'unicode' in wx.PlatformInfo:
                tmpFilename = tmpFilename.encode(TransanaGlobal.encoding)
                tmpPath = tmpPath.encode(TransanaGlobal.encoding)
            # Delete the File
            delResult = srb.srb_remove_obj(self.srbConnectionID, tmpFilename, tmpPath, 0)
            # If an error occurred ...
            if delResult < 0:
                # ... report the error
                self.srbErrorMessage(delResult)
        # Update the Status Text
        self.SetStatusText(_('Delete complete.'))
        # Set the cursor back to the arrow pointer
        self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
        # wxYield allows Windows Messages to process, thus allowing the controls to refresh
        wx.Yield()

    def OnDelete(self, event):
        """ Delete selected file(s) """
        # Determine which side is active and set local controls accordingly
        # If we are deleting on the LEFT side ...
        if event.GetId() == self.btnDeleteLeft.GetId():
            # ... get the LEFT Label, Directory Ctrl, Path, File Ctrl, and Filter Ctrl
            sourceLbl = self.lblLeft
            # ... Directory and Path are different if they're LOCAL or REMOTE (SRB) ...
            if sourceLbl.GetLabel() == self.LOCAL_LABEL:
                sourceDir = self.dirLeft
                sourcePath = sourceDir.GetPath()
            else:
                sourceDir = self.srbDirLeft
                sourcePath = self.GetFullPath(sourceDir)
            sourceFile = self.fileLeft
            sourceFilter = self.filterLeft
            # We also need to know the controls for the inactive side, so they can be updated too if needed
            # ... get the RIGHT Label, Directory Ctrl, Path, File Ctrl, and Filter Ctrl
            otherLbl = self.lblRight
            # ... Directory and Path are different if they're LOCAL or REMOTE (SRB) ...
            if otherLbl.GetLabel() == self.LOCAL_LABEL:
                other = self.dirRight
                otherPath = self.dirRight.GetPath()
            else:
                other = self.srbDirRight
                otherPath = self.GetFullPath(self.srbDirRight)
            otherFile = self.fileRight
            otherFilter = self.filterRight
        # If we are deleting on the RIGHT side ...
        elif event.GetId() == self.btnDeleteRight.GetId():
            # ... get the RIGHT Label, Directory Ctrl, Path, File Ctrl, and Filter Ctrl
            sourceLbl = self.lblRight
            # ... Directory and Path are different if they're LOCAL or REMOTE (SRB) ...
            if sourceLbl.GetLabel() == self.LOCAL_LABEL:
                sourceDir = self.dirRight
                sourcePath = sourceDir.GetPath()
            else:
                sourceDir = self.srbDirRight
                sourcePath = self.GetFullPath(sourceDir)
            sourceFile = self.fileRight
            sourceFilter = self.filterRight
            # We also need to know the controls for the inactive side, so they can be updated too if needed
            # ... get the LEFT Label, Directory Ctrl, Path, File Ctrl, and Filter Ctrl
            otherLbl = self.lblLeft
            # ... Directory and Path are different if they're LOCAL or REMOTE (SRB) ...
            if otherLbl.GetLabel() == self.LOCAL_LABEL:
                other = self.dirLeft
                otherPath = self.dirLeft.GetPath()
            else:
                other = self.srbDirLeft
                otherPath = self.GetFullPath(self.srbDirLeft)
            otherFile = self.fileLeft
            otherFilter = self.filterLeft
         
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

        # If files are listed, delete files.  Otherwise, remove the folder!
        # FILE ItemCount() > 0 indicates we need to delete FILES
        if sourceFile.GetItemCount() > 0:
            # Get all Selected files from the File List.  Starting with itemNum = -1 signals to search the whole list.
            itemNum = -1
            # Keep iterating until all files in the control have been examined
            while True:
                # Request the item number of the next selected item in the file list
                itemNum = sourceFile.GetNextItem(itemNum, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
                # If the return value is -1, there are no more files selected, so stop looping
                if itemNum == -1:
                    break
                else:
                    # If an itemNum is returned, delete the item.
                    self.DeleteFile(sourceLbl.GetLabel(), sourcePath, sourceFile.GetItemText(itemNum))
            # Refresh the file list now that all files have been deleted
            self.RefreshFileList(sourceLbl.GetLabel(), sourcePath, sourceFile, sourceFilter)
        # FILE ItemCount() = 0 indicates we need to delete DIRECTORIES
        else:
            # Create a prompt
            prompt = _('Removing %s')
            # Encode if needed
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Update the Status Bar
            self.SetStatusText(prompt % sourcePath)
            # If we have a LOCAL Directory ...
            if sourceLbl.GetLabel() == self.LOCAL_LABEL:
                # Get the Parent Item
                parentItem = sourceDir.GetTreeCtrl().GetItemParent(sourceDir.GetTreeCtrl().GetSelection())
                # Delete the Current (Selected) Item from the TreeCtrl (Does not remove it from the File System)
                sourceDir.GetTreeCtrl().Delete(sourceDir.GetTreeCtrl().GetSelection())
                # Update the Selection in the Tree to point to the Parent Item
                sourceDir.GetTreeCtrl().SelectItem(parentItem)
                # Remove the original path (folder) from the File System
                os.rmdir(sourcePath)
                # Update the other side's Folder Listing if it's pointing to the same location (Local vs. Remote) 
                if sourceLbl.GetLabel() == otherLbl.GetLabel():
                    # Temporarily match the OTHER SIDE's current selection with the one we just deleted from.
                    other.SetPath(sourceDir.GetPath())
                    # Get the Item ID for the current selection (the directory that has been changed)
                    itemId = other.GetTreeCtrl().GetSelection()
                    # From Robin Dunn -- Force refresh of the tree node where the change occurred
                    other.GetTreeCtrl().CollapseAndReset(itemId)
                    # If we changed the directory temporarily above ...
                    if otherPath != sourcePath:
                        # ... change it back to the directory it used to point to
                        other.SetPath(otherPath)
            # If we have a REMOTE (SRB) Directory ...
            else:
                # Get the Parent Item
                parentItem = sourceDir.GetItemParent(sourceDir.GetSelection())
                # Delete the Current (Selected) Item from the TreeCtrl (Does not remove it from the File System)
                sourceDir.Delete(sourceDir.GetSelection())
                # Update the Selection in the Tree to point to the Parent Item
                sourceDir.SelectItem(parentItem)
                # Get a temporary copy of the source path
                tmpSourcePath = sourcePath
                # Encode if needed
                if 'unicode' in wx.PlatformInfo:
                    tmpSourcePath = sourcePath.encode(TransanaGlobal.encoding)
                # Remove the directory from the SRB
                removeResult = srb.srb_remove_collection(self.srbConnectionID, 0, tmpSourcePath);
                # If an error was detected ...
                if removeResult < 0:
                    # ... display the error message
                    self.srbErrorMessage(removeResult)
                # If the Source and Other are showing the same location, we need to update the Other too!
                if sourceLbl.GetLabel() == otherLbl.GetLabel():
                    # Update the Status Bar
                    self.SetStatusText(_('Reading Collections from SRB.'))
                    # Read the list of folders from the SRB and put them in the SRB Dir Tree
                    # Start by clearing the list
                    other.DeleteAllItems()
                    # Add a Root Item to the SRB Dir Tree
                    srbRootItem = other.AddRoot(self.srbCollection)
                    # Initiate the recursive call to get all Collections from the SRB and add them to the SRB Dir Tree
                    self.srbAddChildNode(other, srbRootItem, self.tmpCollection)
                    # Expand the Root Item to show all the first-level collections
                    other.Expand(srbRootItem)
                    # Select the Root item in the tree
                    other.SelectItem(srbRootItem)

        # If both sides are pointed to the same folder, we need to update the other side's File List as well.  
        if sourcePath == otherPath:
            # Get the Other Path, in case the path saved above has been deleted.  How to get the path
            # depends on if we're local or remote.
            if otherLbl.GetLabel() == self.LOCAL_LABEL:
                otherPath = other.GetPath()
            else:
                otherPath = self.GetFullPath(other)
            # Refresh the file list now that all files have been deleted
            self.RefreshFileList(otherLbl.GetLabel(), otherPath, otherFile, otherFilter)

    def OnUpdateDB(self, event):
        """ Handle "Update DB" Button Press """
        # Determine the File Path Path and File Control
        # If we are updating on the LEFT side ...
        if event.GetId() == self.btnUpdateDBLeft.GetId():
            # ... get the LEFT Label, Path, and File Ctrl
            lblText = self.lblLeft.GetLabel()
            filePath = self.dirLeft.GetPath()
            control = self.fileLeft
        # If we are updating on the RIGHT side ...
        else:
            # ... get the LEFT Label, Path, and File Ctrl
            lblText = self.lblRight.GetLabel()
            filePath = self.dirRight.GetPath()
            control = self.fileRight
        # UpdateDB is only valid for LOCAL files!
        if lblText == self.LOCAL_LABEL:
            # Initialize "item" to -1 so that the first Selected Item will be returned
            item = -1
            # Initialize the File List as empty
            fileList = []
            # Find all selected Items in the appropriate File Control
            while True:
                # Get the next Selected Item in the File Control
                item = control.GetNextItem(item, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
                # If there is no next item...
                if item == -1:
                    # ... we can stop looking
                    break
                # Macs requires an odd conversion.
                if 'wxMac' in wx.PlatformInfo:
                    fileNm = Misc.convertMacFilename(control.GetItemText(item))
                else:
                    fileNm = control.GetItemText(item)
                # Add the item to the File List.
                fileList.append(fileNm)
               
            # Update all listed Files in the Database with the new File Path
            if not DBInterface.UpdateDBFilenames(self, filePath, fileList):
                # Display an error message if the update failed
                infodlg = Dialogs.InfoDialog(self, _('Update Failed.  Some records that would be affected may be locked by another user.'))
                infodlg.ShowModal()
                infodlg.Destroy()

    def srbAddChildNode(self, srbDir, treeNode, srbCollectionName):
        """ This method recursively populates a tree node in the SRB Directory Tree """

        #  srbDaiMcatQuerySubColls runs a query on the SRB server and returns a number of records.  Not the total
        #  number of records, perhaps, but the number it can read into a data buffer.  According to Bing Zhu at
        #  SDSC, it is about 200 records at a time.
        #
        #  Those records can then be accessed by using srbDaiMcatGetColl.
        #
        #  Once all the records in the buffer are read, srbDaiMcatGetMoreColls is called, telling the SRB server
        #  to fill the buffer with the next set of records.  Only when this function returns a value of 0 have all
        #  the records been returned.
        #
        #  Also, this function cannot be called recursively, because the SRB loses it's place when backing out of the
        #  recursion.  Therefore, I will use a List structure to pull all the Collection Names out at once, then call
        #  this method recursively based on that list's items.  In other words, we need to exhaust a given srbDaiMcatQuerySubColls
        #  call before we move on to the next recursion.

        # Start with an empty Node List
        nodeList = []
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
        # Get the list of sub-Collections from the SRB
        srbCatalogCount = srb.srb_query_subcolls(self.srbConnectionID.value, 0, srbCollectionName)
        # Check to see if an error code (negative result) is generated
        if srbCatalogCount < 0:
            # Display the error message
            self.srbErrorMessage(srbCatalogCount)
        # If Sub-Collections are returned 
        else:
            # Process the partial list and request more data from the SRB as long as additional data is forthcoming.
            # (srbCatalogCount > 0 indicates there is more data, while a value of 0 indicates all data has been sent.)
            while srbCatalogCount > 0:
                # Loop through all data in the partial list that has been returned from the SRB so far
                for loop in range(srbCatalogCount):
                    # Create a 1k buffer of spaces to receive data from the SRB
                    buf = ' ' * 1024
                    # Get data for the next Collection in the partial list (indicated by "loop")
                    srb.srb_get_coll_data(loop, buf, len(buf))
                    # Strip whitespace from the buffer
                    tempStr = string.strip(buf)
                    # Only include the last part of the full Collection string in the tree.
                    # (remove everything before the final "/" character)
                    tempStr = tempStr[string.rfind(tempStr, '/') + 1:]
                    # Add the last part of the Collecton String to the SRB Directory Control
                    newTreeNode = srbDir.AppendItem(treeNode, tempStr)
                    # Add the full path, not just the last part of it, to the nodeList for further processing
                    nodeList.append((newTreeNode, string.strip(buf)))
                # When all elements of a partial list have been processed, request the next chunk of the list.
                # If the entire list of Collections has been sent, the SRB function will return 0
                srbCatalogCount = srb.srb_get_more_subcolls(self.srbConnectionID.value)
                # Check to see if Sub-Collections are returned or if an error code (negative result) is generated
                if srbCatalogCount < 0:
                    self.srbErrorMessage(srbCatalogCount)
            # Now that the entire list of sub-Collections under the one that was initially called has been sent,
            # we can recursively request all the sub-Collections for each of the Collections returned, as stored
            # in the Node List.  (See the long comment at the beginning of this method.)
            for (node, collName) in nodeList:
                self.srbAddChildNode(srbDir, node, collName)

    def OnConnect(self, event):
        """ Process the "Connect" button """
        # Determine which side is active and which controls should be manipulated
        # If we are connecting on the LEFT side ...
        if event.GetId() == self.btnConnectLeft.GetId():
            # ... get the LEFT Label, Directory Ctrl, SRB Directory Ctrl, File Ctrl, and Filter Ctrl
            sourceLbl = self.lblLeft
            sourceDir = self.dirLeft
            sourceSrbDir = self.srbDirLeft
            sourceFile = self.fileLeft
            sourceFilter = self.filterLeft
        # If we are connecting on the RIGHT side ...
        else:
            # ... get the RIGHT Label, Directory Ctrl, SRB Directory Ctrl, File Ctrl, and Filter Ctrl
            sourceLbl = self.lblRight
            sourceDir = self.dirRight
            sourceSrbDir = self.srbDirRight
            sourceFile = self.fileRight
            sourceFilter = self.filterRight

        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

        # Determine whether we are connecting to the SRB or disconnecting from it.
        # If the label is "Remote", then we want to disconnect.
        if sourceLbl.GetLabel() == self.REMOTE_LABEL:
            # Change the label
            sourceLbl.SetLabel(self.LOCAL_LABEL)
            # Show the Local File System's Directories
            sourceDir.Show(True)
            # Hide the SRB's Directories
            sourceSrbDir.Show(False)
            # Refresh the File List to now display the appropriate files from either the Local File System or the SRB as appropriate
            if sourceDir.GetPath() != '':
                self.RefreshFileList(sourceLbl.GetLabel(), sourceDir.GetPath(), sourceFile, sourceFilter)
            else:
                sourceFile.ClearAll()
            # Check to see if we need to disconnect, and do so if needed.  (If both sides are "local", we should disconnect.)
            if (self.lblLeft.GetLabel() == self.LOCAL_LABEL) and \
               (self.lblRight.GetLabel() == self.LOCAL_LABEL):
                self.srbDisconnect()
        # If the label is not "Remote", we should connect to the SRB
        else:
            # If no connection exists, we should make a connection to the SRB
            if self.srbConnectionID == None:
                self.SetStatusText(_('Connecting to SRB...'))
                # The Status Bar won't be updated unless we allow the OS to Process Messages.
                wx.Yield()
                # Display the SRB Connection Dialog Box and see if the user fills it out and clicks OK
                if self.SRBConnDlg.ShowModal() == wx.ID_OK:
                    # Set Cursor to the HourGlass
                    self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
                    # Create the appropriate Collection Name from the information we've obtained from the user
                    if self.SRBConnDlg.rbSRBStorageSpace.GetSelection() == 1:
                        # Make sure the Collection Root does NOT end with the '/' character
                        if self.SRBConnDlg.editCollectionRoot.GetValue()[-1] == '/':
                            self.SRBConnDlg.editCollectionRoot.SetValue(self.SRBConnDlg.editCollectionRoot.GetValue()[:-1])
                        self.srbCollection = self.SRBConnDlg.editCollectionRoot.GetValue()
                    else:
                        # Make sure the Collection Root ends with the '/' character
                        if self.SRBConnDlg.editCollectionRoot.GetValue()[-1] != '/':
                            self.SRBConnDlg.editCollectionRoot.SetValue(self.SRBConnDlg.editCollectionRoot.GetValue() + '/')
                        self.srbCollection = self.SRBConnDlg.editCollectionRoot.GetValue() + self.SRBConnDlg.editUserName.GetValue() + '.' + self.SRBConnDlg.editDomain.GetValue()
                    # Assign the function to a Function Name
                    SRBConnectFx = srb.srbDaiConnect

                    # Initialize the ConnectionID variable for use within c_types
                    self.srbConnectionID = ctypes.c_int()

                    # Get temporary variable values for connecting to the SRB
                    self.tmpCollection = self.srbCollection
                    tmpDomain = self.SRBConnDlg.editDomain.GetValue()
                    tmpHost = self.SRBConnDlg.editSRBHost.GetValue()
                    tmpSEAOption = self.SRBConnDlg.editSRBSEAOption.GetValue()
                    tmpPort = self.SRBConnDlg.editSRBPort.GetValue()
                    self.tmpResource = self.SRBConnDlg.editSRBResource.GetValue()
                    tmpUserName = self.SRBConnDlg.editUserName.GetValue()
                    tmpPassword = self.SRBConnDlg.editPassword.GetValue()
                    tmpBuffer = self.SRBConnDlg.choiceBuffer.GetStringSelection()
                    # Encode if needed
                    if 'unicode' in wx.PlatformInfo:
                        self.tmpCollection = self.tmpCollection.encode(TransanaGlobal.encoding)
                        tmpDomain = tmpDomain.encode(TransanaGlobal.encoding)
                        self.tmpResource = self.tmpResource.encode(TransanaGlobal.encoding)
                        tmpSEAOption = tmpSEAOption.encode(TransanaGlobal.encoding)
                        tmpPort = tmpPort.encode(TransanaGlobal.encoding)
                        tmpHost = tmpHost.encode(TransanaGlobal.encoding)
                        tmpUserName = tmpUserName.encode(TransanaGlobal.encoding)
                        tmpPassword = tmpPassword.encode(TransanaGlobal.encoding)
                        tmpBuffer = tmpBuffer.encode(TransanaGlobal.encoding)

                    # Remember the buffer size setting, which is needed when transfers are initiated.
                    # (This parameter was added as different connection speeds work best with different buffer sizes.)
                    self.srbBuffer = tmpBuffer

                    # Call the Function
                    connResult = SRBConnectFx(ctypes.byref(self.srbConnectionID),
                                               self.tmpCollection,
                                               tmpDomain,
                                               self.tmpResource,
                                               tmpSEAOption,
                                               tmpPort,
                                               tmpHost,
                                               tmpUserName,
                                               tmpPassword)

                    # If connection is successful, save the Connection Data in the File Management configuration
                    if self.srbConnectionID.value >= 0:
                        self.SRBConnDlg.SaveConfiguration()
                    # If an error occurred ...
                    else:
                        # ... report the error message
                        self.srbErrorMessage(self.srbConnectionID.value)
                        self.srbConnectionID = None
                        connResult = -1
                   
                    # The File Management Object needs to remember what Resource the user connects to
                    self.srbResource = self.tmpResource  # self.SRBConnDlg.editSRBResource.GetValue()

                else:
                    # Cancelling the Connection Dialog fails to make the connection.  This signals that.
                    connResult = -1
               
            # If a connection to the SRB already exists, we need to update the display values, but don't need to establish the connection
            else:
                # Set Cursor to the HourGlass
                self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
                # The Connection Result from when the Connection was originally made is the srbConnectionID
                connResult = self.srbConnectionID.value

            # connResult of 0 or higher indicates successful connection to the SRB.
            if connResult >= 0:
                # Change the Label to indicate connection
                sourceLbl.SetLabel(self.REMOTE_LABEL)
                # hide the Local File System's Directory Structure
                sourceDir.Show(False)
                # Display the SRB's Directory Structure
                sourceSrbDir.Show(True)
                # Update Status Bar
                self.SetStatusText(_('Reading Collections from SRB.'))
                # Read the list of folders from the SRB and put them in the SRB Dir Tree
                # Start by clearing the list
                sourceSrbDir.DeleteAllItems()
                # Add a Root Item to the SRB Dir Tree
                srbRootItem = sourceSrbDir.AddRoot(self.srbCollection)
                # Initiate the recursive call to get all Collections from the SRB and add them to the SRB Dir Tree
                self.srbAddChildNode(sourceSrbDir, srbRootItem, self.tmpCollection)
                # Expand the Root Item to show all the first-level collections
                sourceSrbDir.Expand(srbRootItem)
                # Select the Root item in the tree
                sourceSrbDir.SelectItem(srbRootItem)
                # Add the Files from the selected Folder
                self.RefreshFileList(sourceLbl.GetLabel(), sourceSrbDir.GetItemText(sourceSrbDir.GetSelection()), sourceFile, sourceFilter)

            # Clear the Status Bar's text
            self.SetStatusText('')
            # Restore the Cursor to show the normal pointer
            self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
        # We're having some size issues.  Call this here to make sure everything is the right size!
        # NOTE:  I HATE THIS.  But I haven't found anything else that works.  I tried:
        #          manual calls to self.OnSize()
        #          calls to self.Refresh() and Self.Update()
        self.Freeze()
        self.SetSize((self.GetSize()[0] + 1, self.GetSize()[1] - 1))
        self.SetSize((self.GetSize()[0] - 1, self.GetSize()[1] + 1))
        self.Thaw()

    def OnNewFolder(self, event):
        """" Implements the "New Folder" button. """
        # Determine which side of the File Management Window we are acting on and identify the control to manipulate
        # If we are connecting on the LEFT side ...
        if event.GetId() == self.btnNewFolderLeft.GetId():
            # ... get the LEFT Label, Directory Ctrl, and path
            sourceLbl = self.lblLeft.GetLabel()
            # ... Directory and Path are different if they're LOCAL or REMOTE (SRB) ...
            if sourceLbl == self.LOCAL_LABEL:
                target = self.dirLeft
                path = target.GetPath()
            else:
                target = self.srbDirLeft
                path = self.GetFullPath(target)
            # We also need to know the controls for the inactive side, so they can be updated too if needed
            # ... get the RIGHT Label and Directory Ctrl
            otherLbl = self.lblRight.GetLabel()
            # ... Directory and Path are different if they're LOCAL or REMOTE (SRB) ...
            if otherLbl == self.LOCAL_LABEL:
                other = self.dirRight
            else:
                other = self.srbDirRight
        # If we are connecting on the RIGHT side ...
        elif event.GetId() == self.btnNewFolderRight.GetId():
            # ... get the RIGHT Label, Directory Ctrl, and path
            sourceLbl = self.lblRight.GetLabel()
            # ... Directory and Path are different if they're LOCAL or REMOTE (SRB) ...
            if sourceLbl == self.LOCAL_LABEL:
                target = self.dirRight
                path = target.GetPath()
            else:
                target = self.srbDirRight
                path = self.GetFullPath(target)
            # We also need to know the controls for the inactive side, so they can be updated too if needed
            # ... get the LEFT Label and Directory Ctrl
            otherLbl = self.lblLeft.GetLabel()
            # ... Directory and Path are different if they're LOCAL or REMOTE (SRB) ...
            if otherLbl == self.LOCAL_LABEL:
                other = self.dirLeft
            else:
                other = self.srbDirLeft
        # Create a prompt
        prompt = _('Create new folder for %s')
        # Encode if needed
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(prompt, 'utf8')
        # Display the prompt in the Status Bar
        self.SetStatusText(prompt % path)

        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
        # Create a TextEntry Dialog to get the name for the new Folder from the user
        newFolder = wx.TextEntryDialog(self, prompt % path, _('Create new folder'), style = wx.OK | wx.CANCEL | wx.CENTRE)
        # Display the Dialog and see if the user presses OK
        if newFolder.ShowModal() == wx.ID_OK:
            # If we're creating a LOCAL directory ...
            if sourceLbl == self.LOCAL_LABEL:
                # Build the full folder name by combining the current path and the text from the user
                folderName = os.path.join(path, newFolder.GetValue())
                if 'wxMSW' in wx.PlatformInfo:
                    OSErrorException = exceptions.WindowsError
                else:
                    OSErrorException = exceptions.OSError
                # Begin exception handling
                try:
                    # Create the folder on the local file system
                    os.mkdir(folderName)
                    # Update the Directory Tree Control.  Get the current selection
                    itemId = target.GetTreeCtrl().GetSelection()
                     # From Robin Dunn -- Forces refresh of the tree node to get it to update.
                    target.GetTreeCtrl().CollapseAndReset(itemId)
                    # If both sides are pointed to the same LOCAL / REMOTE (SRB) source, we need to update the other side too.
                    if sourceLbl == otherLbl:
                        # Get the current path for the opposite side
                        otherPath = other.GetPath()
                        # Set the path temporarily to the folder that was changed.
                        other.SetPath(path)
                        # Get the current selection
                        itemId = other.GetTreeCtrl().GetSelection()
                         # From Robin Dunn -- Forces refresh of the tree node to get it to update.
                        other.GetTreeCtrl().CollapseAndReset(itemId)
                        # Reset the path to the original location
                        other.SetPath(otherPath)
                except OSErrorException, e:
                    # Probably a duplicate directory name
                    prompt = _('Folder "%s" already exists.')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(prompt, 'utf8')
                    errordlg = Dialogs.ErrorDialog(None, prompt % folderName)
                    errordlg.ShowModal()
                    errordlg.Destroy()

                except:
                    prompt = "%s : %s"
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(prompt, 'utf8')
                    errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    
                # Automatically change to the newly created folder
                target.SetPath(folderName)
            # If we're creating a REMOTE (SRB) Directory (Collection) ...
            else:
                # Get a temporary copy of the new folder name
                tmpNewFolder = newFolder.GetValue()
                # Encode if needed
                if 'unicode' in wx.PlatformInfo:
                    tmpNewFolder = tmpNewFolder.encode(TransanaGlobal.encoding)
                    path = path.encode(TransanaGlobal.encoding)
                # Create the folder on the SRB file system
                newCollection = srb.srb_new_collection(self.srbConnectionID, 0, path, tmpNewFolder);
                # If no error occurred
                if newCollection >= 0:
                    # Update the Directory Tree Control
                    itemId = target.AppendItem(target.GetSelection(), newFolder.GetValue())
                    # If both sides are pointed to the same LOCAL / REMOTE (SRB) source, we need to update the other side too.
                    if sourceLbl == otherLbl:
                        # Update the Status Bar
                        self.SetStatusText(_('Reading Collections from SRB.'))
                        # Read the list of folders from the SRB and put them in the SRB Dir Tree
                        # Start by clearing the list
                        other.DeleteAllItems()
                        # Add a Root Item to the SRB Dir Tree
                        srbRootItem = other.AddRoot(self.srbCollection)
                        # Initiate the recursive call to get all Collections from the SRB and add them to the SRB Dir Tree
                        self.srbAddChildNode(other, srbRootItem, self.tmpCollection)
                        # Expand the Root Item to show all the first-level collections
                        other.Expand(srbRootItem)
                        # Select the Root item in the tree
                        other.SelectItem(srbRootItem)
                    # Automatically change to the newly created folder
                    target.SelectItem(itemId)
                # If an error did occur ...
                else:
                    # ... display the error message
                    self.srbErrorMessage(newCollection)
        # Destroy the TextEntry Dialog, now that we are done with it.
        newFolder.Destroy()

    def OnRefresh(self, event):
        """ Method to use when the Refresh Button is pressed.  """
        # Update Directory Listings
        # Get the LEFT label
        lblLeft = self.lblLeft.GetLabel()
        # If LEFT is LOCAL ...
        if lblLeft == self.LOCAL_LABEL:
            # ... refresh the LOCAL directory control
            self.RefreshDirCtrl(lblLeft, self.dirLeft)
        # If LEFT is REMOTE ...
        else:
            # ... refresh the REMOTE directory control
            self.RefreshDirCtrl(lblLeft, self.srbDirLeft)
        # Get the RIGHT label
        lblRight = self.lblRight.GetLabel()
        # If RIGHT is LOCAL ...
        if lblRight == self.LOCAL_LABEL:
            # ... refresh the LOCAL directory control
            self.RefreshDirCtrl(lblRight, self.dirRight)
        # If RIGHT is REMOTE ...
        else:
            # ... refresh the REMOTE directory control
            self.RefreshDirCtrl(lblRight, self.srbDirRight)

        # Update File Listings
        # Initialize a path variable
        path = ''
        # Get the appropriate path, depending on if the LEFT is LOCAL or REMOTE (SRB)
        if lblLeft == self.LOCAL_LABEL:
            path = self.dirLeft.GetPath()
        else:
            path = self.GetFullPath(self.srbDirLeft)
        # As long as there is a defined path ...
        if path != '':
            # ... refresh the left hand side controls
            self.RefreshFileList(lblLeft, path, self.fileLeft, self.filterLeft)

        # Initialize a path variable
        path = ''
        # Get the appropriate path, depending on if the LEFT is LOCAL or REMOTE (SRB)
        if lblRight == self.LOCAL_LABEL:
            path = self.dirRight.GetPath()
        else:
            path = self.GetFullPath(self.srbDirRight)
        # As long as there is a defined path ...
        if path != '':
            # refresh the right hand side controls
            self.RefreshFileList(lblRight, path, self.fileRight, self.filterRight)

    def OnHelp(self, event):
        """ Method to use when the Help Button is pressed """
        helpContext = 'File Management'
        # Getting this to work both from within Python and in the stand-alone executable
        # has been a little tricky.  To get it working right, we need the path to the
        # Transana executables, where Help.exe resides, and the file name, which tells us
        # if we're in Python or not.
        (path, fn) = os.path.split(sys.argv[0])
        # If the path is not blank, add the path seperator to the end if needed
        if (path != '') and (path[-1] != os.sep):
            path = path + os.sep
        if "__WXMAC__" in wx.PlatformInfo:
            # NOTE:  If we just call Help.Help(), you can't actually do the Tutorial because
            # the Help program's menus override Transana's, and there's no way to get them back.
            # instead of the old call:
            
            # Help.Help(helpContext)
            
            # NOTE:  I've tried a bunch of different things on the Mac without success.  It seems that
            #        the Mac doesn't allow command line parameters, and I have not been able to find
            #        a reasonable method for passing the information to the Help application to tell it
            #        what page to load.  What works is to save the string to the hard drive and 
            #        have the Help file read it that way.  If the user leave Help open, it won't get
            #        updated on subsequent calls, but for now that's okay by me.
            
            fileObj = open(os.getenv("HOME") + '/TransanaHelpContext.txt', 'w')
            pickle.dump(helpContext, fileObj)
            fileObj.flush()
            fileObj.close()

            # On OS X 10.4, when Transana is packed with py2app, the Help call stopped working.
            # It seems we have to remove certain environment variables to get it to work properly!
            # Let's investigate environment variables here!
            envirVars = os.environ
            if 'PYTHONHOME' in envirVars.keys():
                del(os.environ['PYTHONHOME'])
            if 'PYTHONPATH' in envirVars.keys():
                del(os.environ['PYTHONPATH'])
            if 'PYTHONEXECUTABLE' in envirVars.keys():
                del(os.environ['PYTHONEXECUTABLE'])

            os.system('open -a TransanaHelp.app')

        else:
            # Make the Help call differently from Python and the stand-alone executable.
            if fn.lower() == 'transana.py':
                # for within Python, we call python, then the Help code and the context
                os.spawnv(os.P_NOWAIT, 'python', [path + 'Help.py', helpContext])
            else:
                # The Standalone requires a "dummy" parameter here (Help), as sys.argv differs between the two versions.
                os.spawnv(os.P_NOWAIT, path + 'Help', ['Help', helpContext])

    def GetFullPath(self, control):
        """ Produce the full Path secification for the given srb Directory Tree Control """
        # Start with the current selection and work backwards up the tree
        item = control.GetSelection()
        # Start the Path List with the current folder name
        pathList = control.GetItemText(item).strip()
        # Move up the tree one node at a time until we get to the root
        while item != control.GetRootItem():
            item = control.GetItemParent(item)
            # Add the new folder name to the FRONT of the path list
            pathList = control.GetItemText(item).strip() + '/' + pathList  # '/', not os.sep, because the SRB ALWAYS uses unix-style paths.
        # On the Mac, there are null characters (#0) ending the strings.  We need to remove these.
        while pathList.find(chr(0)) > -1:
            pathList = pathList[:pathList.find(chr(0))] + pathList[pathList.find(chr(0)) + 1:]
        # Return the path list that has been built
        return pathList

    def IsEmpty(self, lbl, path):
        """ This method detects whether the specified file path is empty or not.  This is used to determine
            whether the "Delete" button should delete files or a folder.  """
        # Default to True, indicating that the target path IS empty.
        result = True
        # We need to know whether to look at the local file path or at the SRB.  If LOCAL ...
        if lbl == self.LOCAL_LABEL:
            # We need to trap the exception if the path is invalid (eg. pointing to a CD drive with no CD)
            try:
                # If local, see if any files are in the folder using the Python os module
                fileList = os.listdir(path)
            except:
                # If the exception is raised, return the result, which would still be True here
                return result
            # If any files are returned ...
            if len(fileList) != 0:
                # ... then we're not Empty.
                result = False
        # If REMOTE (SRB)
        else:
            # Load the SRB DLL / Dynamic Library
            if "__WXMSW__" in wx.PlatformInfo:
                srb = ctypes.cdll.srbClient
            else:
                srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
            # Encode the path, if needed
            if 'unicode' in wx.PlatformInfo:
                path = path.encode(TransanaGlobal.encoding)
            # See if there's anything in the SRB catalog
            srbCatalogCount = srb.srb_query_child_dataset(self.srbConnectionID.value, 0, path)
            # If any files are counted ...
            if srbCatalogCount != 0:
                # ... then we're not Empty.
                result = False
        # Return the result
        return result

    def OnDirSelect(self, event):
        """ This method is called when a directory/folder is selected in the tree, either locally or on the SRB """
        # First, determine whether we're on the source or destination side and set local variables lbl, path, target,
        # and filter accordingly.
        if (event.GetEventObject() == self.dirLeft.GetTreeCtrl()) or \
           (event.GetId() == self.srbDirLeft.GetId()) or \
           (event.GetId() == self.filterLeft.GetId()):
            lbl = self.lblLeft.GetLabel()
            # We need a different path depending on whether we are on the local file system or on the SRB
            if lbl == self.LOCAL_LABEL:
                path = self.dirLeft.GetPath()
            else:
                path = self.GetFullPath(self.srbDirLeft)
            target = self.fileLeft
            filter = self.filterLeft
        elif (event.GetEventObject() == self.dirRight.GetTreeCtrl()) or \
             (event.GetId() == self.srbDirRight.GetId()) or \
             (event.GetId() == self.filterRight.GetId()):
            lbl = self.lblRight.GetLabel()
            # We need a different path depending on whether we are on the local file system or on the SRB
            if lbl == self.LOCAL_LABEL:
                path = self.dirRight.GetPath()
            else:
                path = self.GetFullPath(self.srbDirRight)
            target = self.fileRight
            filter = self.filterRight
        prompt = _('Directory changed to %s')
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(prompt, 'utf8')
        self.SetStatusText(prompt % path)
        # We need to trap the exception if the path is invalid (eg. pointing to a CD drive with no CD)
        try:
            # Refresh the File List to reflect any changes made to path or filter
            self.RefreshFileList(lbl, path, target, filter)
        except:
            import traceback
            traceback.print_exc(file=sys.stdout)
            # Display the Exception Message, allow "continue" flag to remain true
            prompt = "%s : %s"
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
            errordlg.ShowModal()
            errordlg.Destroy()

            # If the exception is raised, see if we're looking at a local drive.
            if lbl == self.LOCAL_LABEL:
                # The File List has already been refreshed.  Since the path was bogus, it was refreshed with
                # data from the current working directory, which is totally NOT helpful.  Let's clear it.
                self.RefreshFileList(lbl, '', target, filter)

    def RefreshFileList(self, lbl, path, target, filter):
        """ Refresh the File List.  This is needed whenever something happens where the file list might change,
            such as when a Folder is selected.  It is also used for the Refresh button, which allows the user to
            update the screen to reflect changes that may have occurred outside of Transana or the File Management
            Utility."""
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

        # Clear the File List
        target.ClearAll()
        target.InsertColumn(0, "Files")
        target.SetColumnWidth(0, target.GetSizeTuple()[0] - 24)

        # If the path is blank, don't populate the list !  (Blank path signals that an invalid
        # drive was selected, such as a CD drive with no CD in it.)
        if path == '':
            return

        # If we are looking at the Local File System...
        if lbl == self.LOCAL_LABEL:
            # We need to trap the exception for if the path is bogus, such as to a CD drive with no CD in it.
            try:
                # Read the list of files from the Local File System
                filelist = os.listdir(path)
                # Get selection from dropdown list
                selection=filter.GetStringSelection()
                # Iterate through all files from the filelist and add to the ListCtrl target
                for item in self.itemsToShow(path, filelist, selection, isSrbList=False):
                    index = target.InsertStringItem(target.GetItemCount(), item)
            # The DirRefresh model causes a WindowsError to be raised when there is an empty CD drive
            # on the system.  This ignores that error.
            except WindowsError, e:
                pass
            except:
                import traceback
                traceback.print_exc(file=sys.stdout)
                # Display the Exception Message, allow "continue" flag to remain true
                prompt = "%s : %s"
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
                # Pass the exception on up the path
                raise

        # If we are not looking at the Local File System, we are looking at a SRB connection
        else:

            filelist = self.GetSRBFileList(path)

            # Get selection from dropdown list
            selection=filter.GetStringSelection()
            # Iterate through all files from the filelist and add to the ListCtrl target
            for item in self.itemsToShow(path, filelist, selection, isSrbList=True):
                target.InsertStringItem(target.GetItemCount(), item)

    def itemsToShow(self, path, filelist, selection, isSrbList):
        """Return a list containing filenames to be displayed depending upon the file extensions found in selection"""
        itemList = [] # Initialize return value to []

        # DO NOT SORT THE FILE LIST if it is on the SRB!!!!  This breaks the connection with the SRB so that files are not
        # given the proper File Size when downloading them from the SRB, thus making Transfer look bad.
        if not isSrbList:
            filelist.sort() # Sort the file list
        # Create a REfilter which can match '.' followed by alphanumerics and _
        regexp = re.compile(r'\.\w*') 
        # Get a list containing all extensions from selection
        srch = regexp.findall(selection)
        prompt = _('All files (*.*)')
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(prompt, 'utf8')
        # For each item in the file list
        for item in filelist:
            if isSrbList: 
                # files in the file list from the SRB have an invisible character [chr(0)] on the end that does not get stripped out.
                if ord(item[-1]) == 0:
                    item = item[:-1]
                # And if either *.* is selected or item's extension is present in the list srch
                if (selection == prompt) | (srch.count(string.lower(os.path.splitext(item)[1])) > 0):
                    itemList.append(item)            
            else:
                #If it is not a directory
                if os.path.isfile(os.path.join(path, item)): 
                    # And if either *.* is selected or item's extension is present in the list srch
                    if (selection == prompt) | (srch.count(string.lower(os.path.splitext(item)[1])) > 0):
                        itemList.append(item)
       
        return itemList

    def srbErrorMessage(self, errorCode):
        """ Translate a srb error message code into the relevant string and present the error message """
        errorMsg = " " * 200
        # Load the SRB DLL / Dynamic Library
        if "__WXMSW__" in wx.PlatformInfo:
            srb = ctypes.cdll.srbClient
        else:
            srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
        # Call the Function that gets the error message
        srb.srb_err_msg(errorCode, errorMsg, 200)
        # Strip whitespace and null character (c string terminator) from buf
        errorMsg = string.strip(errorMsg)[:-1]

        prompt = _('An error has occurred in working with the Storage Resource Broker.\nError Code: %s\nError Message: "%s"')
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(prompt, 'utf8')
          
        dlg = Dialogs.ErrorDialog(self, prompt % (errorCode, errorMsg))
        dlg.ShowModal()
        dlg.Destroy()


# Allow this module to function as a stand-alone application
if __name__ == '__main__':

    class MyApp(wx.App):
        def OnInit(self):
            gettext.install("Transana", './locale', False)
            frame = FileManagement(None, -1, _('Transana File Management'))
            frame.Setup()
            self.SetTopWindow(frame)
            return True
      

    app = MyApp(0)
    app.MainLoop()
