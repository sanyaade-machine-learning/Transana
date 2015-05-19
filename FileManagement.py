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

"""This module implements the File Management window for Transana, but can be also run as a stand-alone utility.
   The File Management Tool is designed to facilitate local file management, such as copying and moving video
   files, and updating the Transana Database when video files have been moved.  It is also used for connecting
   to the Storage Resource Broker (SRB) and transfering files between the local file system and the SRB.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare <rasambhare@wisc.edu>'

if __name__ == '__main__':
    import wxversion
    wxversion.select('2.6.1.0-unicode')
    
import wx  # Import wxPython

if __name__ == '__main__':
    __builtins__._ = wx.GetTranslation
    if ('unicode' in wx.PlatformInfo) and (wx.VERSION_STRING >= '2.6'):
        wx.SetDefaultPyEncoding('utf_8')

import os                # used to determine what files are on the local computer
import sys               # Python sys module, used in exception reporting
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

# Control Constants
ID_SOURCEDIR       = wx.NewId()    # ID for the wxGenericDirCtrl that displays drives and folders on the left side of the screen
ID_SRBSOURCEDIR    = wx.NewId()    # ID for the Tree Control that holds Collections from the SRB
ID_DESTDIR         = wx.NewId()    # ID for the wxGenericDirCtrl that displays drives and folders on the right side of the screen
ID_SRBDESTDIR      = wx.NewId()    # ID for the Tree Control that holds Collections from the SRB
                                   # NOTE:  Although it would be possible to display files as well as folders in the wxGenericDirCtrls,
                                   #        we chose not to so that users could select multiple files at once for manipulation
ID_SOURCEFILE      = wx.NewId()    # ID for the list of files on the left side of the screen
ID_DESTFILE        = wx.NewId()    # ID for the list of files on the right side of the screen
ID_SOURCEFILTER    = wx.NewId()    # ID for the File Filter control on the left side of the screen
ID_DESTFILTER      = wx.NewId()    # ID for the File Filter control on the right side of the screen
ID_BTNCOPY         = wx.NewId()    # ID for the Copy button
ID_BTNMOVE         = wx.NewId()    # ID for the Move button
ID_BTNDELETE       = wx.NewId()    # ID for the Delete button
ID_BTNNEWFOLDER    = wx.NewId()    # ID for the New Folder button
ID_BTNUPDATEDB     = wx.NewId()    # ID for the Update DB button
ID_BTNCONNECT      = wx.NewId()    # ID for the Connect button
ID_BTNREFRESH      = wx.NewId()    # ID for the Refresh button
ID_BTNCLOSE        = wx.NewId()    # ID for the Close button
ID_BTNHELP         = wx.NewId()    # ID for the Help button


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
      if (self.activeSide == 'Source'):
         targetLbl = self.FileManagementWindow.lblSource
         # Local and Remote File Systems have different methods for determing the selected Path
         if targetLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
            targetDir = self.FileManagementWindow.sourceDir.GetPath()
         else:
            targetDir = self.FileManagementWindow.GetFullPath(self.FileManagementWindow.srbSourceDir)
         target = self.FileManagementWindow.sourceFile
         filter = self.FileManagementWindow.sourceFilter
         # We need to know the other side to, so it can also be updated if necessary
         otherLbl = self.FileManagementWindow.lblDest
         # Local and Remote File Systems have different methods for determing the selected Path
         if otherLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
            otherDir = self.FileManagementWindow.destDir.GetPath()
         else:
            otherDir = self.FileManagementWindow.GetFullPath(self.FileManagementWindow.srbDestDir)
         othertarget = self.FileManagementWindow.destFile
         otherfilter = self.FileManagementWindow.destFilter
      elif self.activeSide == 'Dest':
         targetLbl = self.FileManagementWindow.lblDest
         # Local and Remote File Systems have different methods for determing the selected Path
         if targetLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
            targetDir = self.FileManagementWindow.destDir.GetPath()
         else:
            targetDir = self.FileManagementWindow.GetFullPath(self.FileManagementWindow.srbDestDir)
         target = self.FileManagementWindow.destFile
         filter = self.FileManagementWindow.destFilter
         # We need to know the other side to, so it can also be updated if necessary
         otherLbl = self.FileManagementWindow.lblSource
         # Local and Remote File Systems have different methods for determing the selected Path
         if otherLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
            otherDir = self.FileManagementWindow.sourceDir.GetPath()
         else:
            otherDir = self.FileManagementWindow.GetFullPath(self.FileManagementWindow.srbSourceDir)
         othertarget = self.FileManagementWindow.sourceFile
         otherfilter = self.FileManagementWindow.sourceFilter

      # If a Path has been defined for the Drop Target, copy the files.
      # First, determine if it's local or SRB receiving the files
      if targetLbl.GetLabel() == self.FileManagementWindow.LOCAL_LABEL:
          if targetDir != '':
             self.FileManagementWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
             # for each file in the list...
             for file in filenames:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Copying %s to %s'), 'utf8')
                else:
                    prompt = _('Copying %s to %s')
                self.FileManagementWindow.SetStatusText(prompt % (file, targetDir))
                # extract the file name...
                (dir, fn) = os.path.split(file)
                # copy the file to the destination path
                shutil.copyfile(file, os.path.join(targetDir, fn))
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
         # For each file in the fileList ...
         for fileName in filenames:
            # Divide the fileName into directory and filename portions
            (sourceDir, file) = os.path.split(fileName)
            # Get the File Size
            fileSize = os.path.getsize(fileName)
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Copying %s to %s'), 'utf8')
            else:
                prompt = _('Copying %s to %s')
            self.FileManagementWindow.SetStatusText(prompt % (file, targetDir))

            # The SRBFileTransfer class handles file transfers and provides Progress Feedback
            dlg = SRBFileTransfer.SRBFileTransfer(self.FileManagementWindow, _("SRB File Transfer"), file, fileSize, sourceDir, self.FileManagementWindow.srbConnectionID, targetDir, SRBFileTransfer.srb_UPLOAD)
            success = dlg.TransferSuccessful()
            dlg.Destroy()

            if success:
                self.FileManagementWindow.SetStatusText(_('Copy complete.'))
            else:
                self.FileManagementWindow.SetStatusText(_('Copy cancelled.'))
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

      # activeSide indicates which side of the screen should be acted upon, the left ('Source') or the right ('Dest').
      # It is initialized to None, as neither side is selected by default
      self.activeSide = None
      # srbConnectionID indicates the Connection ID for the Storage Resource Broker connection.
      # It is initialized to None, as the connection is not established by default
      self.srbConnectionID = None

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
      pre.Create(parent, -1, title, pos = wx.DefaultPosition, size = (700,550), style=wx.DEFAULT_DIALOG_STYLE | wx.CAPTION | wx.RESIZE_BORDER)
      # Convert the PreDialog into a REAL wrapper of the wxPython extension (whatever that means)
      self.this = pre.this

      # To look right, the Mac needs the Small Window Variant.
      if "__WXMAC__" in wx.PlatformInfo:
          self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

      # Create the Connection Dialog
      self.SRBConnDlg = SRBConnection.SRBConnection(self)


   def Setup(self, showModal=False):
      """ Set up the form widgets for the File Management Window """
      # Set the width of the center column buttons.
      # The buttons need to be slightly wider on the Mac to accommodate "Delete Folder >>"
      if 'wxMac' in wx.PlatformInfo:
          btnWidth = 110
      else:
          btnWidth = 100

      # Copy Button in the center of the screen
      lay = wx.LayoutConstraints()
      lay.top.PercentOf(self, wx.Height, 7)          
      lay.centreX.PercentOf(self, wx.Width, 50)       
      lay.width.Absolute(btnWidth)  
      lay.height.AsIs()
      self.btnCopy = wx.Button(self, ID_BTNCOPY, _("Copy"))
      self.btnCopy.SetConstraints(lay)

      # Move Button in the center of the screen
      lay = wx.LayoutConstraints()
      lay.top.PercentOf(self, wx.Height, 17)          
      lay.centreX.PercentOf(self, wx.Width, 50)       
      lay.width.SameAs(self.btnCopy, wx.Width)  
      lay.height.AsIs()
      self.btnMove = wx.Button(self, ID_BTNMOVE, _("Move"))
      self.btnMove.SetConstraints(lay)

      # Delete Button in the center of the screen
      lay = wx.LayoutConstraints()
      lay.top.PercentOf(self, wx.Height, 27)          
      lay.centreX.PercentOf(self, wx.Width, 50)       
      lay.width.SameAs(self.btnCopy, wx.Width)  
      lay.height.AsIs()
      self.btnDelete = wx.Button(self, ID_BTNDELETE, _("Delete"))
      self.btnDelete.SetConstraints(lay)

      # New Folder Button in the center of the screen
      lay = wx.LayoutConstraints()
      lay.top.PercentOf(self, wx.Height, 37)          
      lay.centreX.PercentOf(self, wx.Width, 50)       
      lay.width.SameAs(self.btnCopy, wx.Width)  
      lay.height.AsIs()
      self.btnNewFolder = wx.Button(self, ID_BTNNEWFOLDER, _("New Folder"))
      self.btnNewFolder.SetConstraints(lay)

      # Update DB Button in the center of the screen
      lay = wx.LayoutConstraints()
      lay.top.PercentOf(self, wx.Height, 47)          
      lay.centreX.PercentOf(self, wx.Width, 50)       
      lay.width.SameAs(self.btnCopy, wx.Width)  
      lay.height.AsIs()
      self.btnUpdateDB = wx.Button(self, ID_BTNUPDATEDB, _("Update DB"))
      self.btnUpdateDB.SetConstraints(lay)

      # Connect to SRB Button in the center of the screen
      lay = wx.LayoutConstraints()
      lay.top.PercentOf(self, wx.Height, 57)          
      lay.centreX.PercentOf(self, wx.Width, 50)       
      lay.width.SameAs(self.btnCopy, wx.Width)  
      lay.height.AsIs()
      self.btnConnect = wx.Button(self, ID_BTNCONNECT, _("Connect"))
      self.btnConnect.SetConstraints(lay)

      # Refresh Button in the center of the screen
      lay = wx.LayoutConstraints()
      lay.top.PercentOf(self, wx.Height, 67)          
      lay.centreX.PercentOf(self, wx.Width, 50)       
      lay.width.SameAs(self.btnCopy, wx.Width)  
      lay.height.AsIs()
      self.btnRefresh = wx.Button(self, ID_BTNREFRESH, _("Refresh"))
      self.btnRefresh.SetConstraints(lay)

      # Close Button in the center of the screen
      lay = wx.LayoutConstraints()
      lay.top.PercentOf(self, wx.Height, 77)          
      lay.centreX.PercentOf(self, wx.Width, 50)       
      lay.width.SameAs(self.btnCopy, wx.Width)  
      lay.height.AsIs()
      self.btnClose = wx.Button(self, ID_BTNCLOSE, _("Close"))
      self.btnClose.SetConstraints(lay)

      # Help Button in the center of the screen
      lay = wx.LayoutConstraints()
      lay.top.PercentOf(self, wx.Height, 87)          
      lay.centreX.PercentOf(self, wx.Width, 50)       
      lay.width.SameAs(self.btnCopy, wx.Width)  
      lay.height.AsIs()
      self.btnHelp = wx.Button(self, ID_BTNHELP, _("Help"))
      self.btnHelp.SetConstraints(lay)

      # Although all activities in this window are bidirectional, I have chosen to refer to the
      # Left-side screen widgets as "Source" widgets and the Right-side widgets as "Destination"
      # widgets

      # "Source" label
      lay = wx.LayoutConstraints()
      lay.top.SameAs(self, wx.Top, 10)
      lay.left.SameAs(self, wx.Left, 10)
      lay.width.Absolute(200)
      lay.height.AsIs()
      self.lblSource = wx.StaticText(self, -1, _("Local:"), style=wx.ST_NO_AUTORESIZE)
      self.lblSource.SetConstraints(lay)

      # Source Folders on Local Drive
      lay = wx.LayoutConstraints()
      lay.top.Below(self.lblSource, 5)
      lay.left.SameAs(self, wx.Left, 10)
      lay.right.LeftOf(self.btnCopy, 10)
      lay.bottom.PercentOf(self, wx.Height, 35)
      self.sourceDir = wx.GenericDirCtrl(self, ID_SOURCEDIR, style=wx.DIRCTRL_3D_INTERNAL | wx.DIRCTRL_DIR_ONLY)
      # To distinguish between the source and destinations controls in this tool, we have to reassign the
      # wxGenericDirCtrl's TreeCtrl's ID
      self.sourceDir.SetConstraints(lay)

      # Source SRB Collections listing
      lay = wx.LayoutConstraints()
      lay.top.Below(self.lblSource, 5)
      lay.left.SameAs(self, wx.Left, 10)
      lay.right.LeftOf(self.btnCopy, 10)
      lay.bottom.PercentOf(self, wx.Height, 35)
      self.srbSourceDir = wx.TreeCtrl(self, ID_SRBSOURCEDIR, style=wx.DIRCTRL_3D_INTERNAL | wx.TR_HAS_BUTTONS | wx.TR_SINGLE)
      self.srbSourceDir.SetConstraints(lay)
      # This control is not visible initially, as Local Folders, not SRB Collections are shown
      self.srbSourceDir.Show(False)
      
      # Source File List
      lay = wx.LayoutConstraints()
      lay.top.Below(self.sourceDir, 10)
      lay.left.SameAs(self, wx.Left, 10)
      lay.right.LeftOf(self.btnCopy, 10)
      lay.bottom.SameAs(self, wx.Bottom, 60)
      self.sourceFile = wx.ListCtrl(self, ID_SOURCEFILE, style=wx.LC_LIST | wx.LC_SORT_ASCENDING | wx.LC_NO_HEADER)
      self.sourceFile.SetConstraints(lay)

      # Make the Source File List a FileDropTarget
      dt = FMFileDropTarget(self, 'Source')
      self.sourceFile.SetDropTarget(dt)

      # Source File Filter
      lay = wx.LayoutConstraints()
      lay.top.Below(self.sourceFile, 10)
      lay.left.SameAs(self, wx.Left, 10)
      lay.right.LeftOf(self.btnCopy, 10)
      lay.bottom.SameAs(self, wx.Bottom, 30)
      # We need to rebuild the list so the items get translated if a non-English language is being used!
      fileChoiceList = []
      for item in TransanaConstants.fileTypesList:
          fileChoiceList.append(_(item))
      self.sourceFilter = wx.Choice(self, ID_SOURCEFILTER, choices=fileChoiceList)
      self.sourceFilter.SetSelection(0)
      self.sourceFilter.SetConstraints(lay)

      # "Destination" label
      lay = wx.LayoutConstraints()
      lay.top.SameAs(self, wx.Top, 10)
      lay.left.RightOf(self.btnCopy, 10)
      lay.width.Absolute(200)
      lay.height.AsIs()
      self.lblDest = wx.StaticText(self, -1, _("Local:"), style=wx.ST_NO_AUTORESIZE)
      self.lblDest.SetConstraints(lay)

      # Destination Folders for Local Drives
      lay = wx.LayoutConstraints()
      lay.top.Below(self.lblDest, 5)
      lay.left.SameAs(self.lblDest, 0)
      lay.right.SameAs(self, wx.Right, 10)  
      lay.bottom.PercentOf(self, wx.Height, 35)
      self.destDir = wx.GenericDirCtrl(self, ID_DESTDIR, style=wx.DIRCTRL_3D_INTERNAL | wx.DIRCTRL_DIR_ONLY)
      # To distinguish between the source and destinations controls, we have to reassign the wxGenericDirCtrl's TreeCtrl's ID
      self.destDir.SetConstraints(lay)

      # Destination SRB Collections listing
      lay = wx.LayoutConstraints()
      lay.top.Below(self.lblDest, 5)
      lay.left.SameAs(self.lblDest, 0)
      lay.right.SameAs(self, wx.Right, 10)  
      lay.bottom.PercentOf(self, wx.Height, 35)
      self.srbDestDir = wx.TreeCtrl(self, ID_SRBDESTDIR, style=wx.DIRCTRL_3D_INTERNAL | wx.TR_HAS_BUTTONS | wx.TR_SINGLE)
      self.srbDestDir.SetConstraints(lay)
      # This control is not visible initially, as Local Folders, not SRB Collections are shown
      self.srbDestDir.Show(False)

      # Destination File List
      lay = wx.LayoutConstraints()
      lay.top.Below(self.destDir, 10)
      lay.left.SameAs(self.lblDest, 0)
      lay.right.SameAs(self, wx.Right, 10)
      lay.bottom.SameAs(self, wx.Bottom, 60)
      self.destFile = wx.ListCtrl(self, ID_DESTFILE, style=wx.LC_LIST | wx.LC_SORT_ASCENDING | wx.LC_NO_HEADER) 
      self.destFile.SetConstraints(lay)

      # Make the Destination File List a FileDropTarget
      dt = FMFileDropTarget(self, 'Dest')
      self.destFile.SetDropTarget(dt)

      # Destination File Filter
      lay = wx.LayoutConstraints()
      lay.top.Below(self.destFile, 10)
      lay.left.SameAs(self.lblDest, 0)
      lay.right.SameAs(self, wx.Right, 10)
      lay.bottom.SameAs(self, wx.Bottom, 30)
      self.destFilter = wx.Choice(self, ID_DESTFILTER, choices=fileChoiceList)
      self.destFilter.SetSelection(0)
      self.destFilter.SetConstraints(lay)

      # wxFrames can have Status Bars.  wxDialogs can't.
      # The File Management Window needs to be a wxDialog, not a wxFrame.
      # Yet I want it to have a Status Bar.  So I'm going to fake a status bar by
      # creating a Panel at the bottom of the screen with text on it.
      
      # Status Bar Panel
      lay = wx.LayoutConstraints()
      lay.top.SameAs(self, wx.Bottom, -20)
      lay.left.SameAs(self, wx.Left, 0)
      lay.right.SameAs(self, wx.Right, 0)
      lay.bottom.SameAs(self, wx.Bottom, 0)
      StatusBar = wx.Panel(self, -1, style=wx.SUNKEN_BORDER, name='FileManagementPanel')
      StatusBar.SetConstraints(lay)

      # Place a Static Text Widget on the Status Bar Panel
      lay = wx.LayoutConstraints()
      lay.top.SameAs(StatusBar, wx.Top, 3)
      lay.left.SameAs(StatusBar, wx.Left, 6)
      lay.right.SameAs(StatusBar, wx.Right, 6)
      lay.bottom.SameAs(StatusBar, wx.Bottom, 2)
      self.StatusText = wx.StaticText(self, -1, "", style=wx.ST_NO_AUTORESIZE)
      self.StatusText.SetConstraints(lay)
      
      # Disable buttons that are not initially available to the user.
      # The user must select one or more screen locations before these operations make sense.
      self.btnConnect.Enable(False)
      self.btnNewFolder.Enable(False)
      self.btnCopy.Enable(False)
      self.btnMove.Enable(False)
      self.btnDelete.Enable(False)
      self.btnUpdateDB.Enable(False)

      # Define Events

      # Mouse Over Events to enable Button Descriptions in the Status Bar
      wx.EVT_ENTER_WINDOW(self.btnCopy, self.OnMouseOver)
      wx.EVT_LEAVE_WINDOW(self.btnCopy, self.OnMouseExit)
      wx.EVT_ENTER_WINDOW(self.btnMove, self.OnMouseOver)
      wx.EVT_LEAVE_WINDOW(self.btnMove, self.OnMouseExit)
      wx.EVT_ENTER_WINDOW(self.btnDelete, self.OnMouseOver)
      wx.EVT_LEAVE_WINDOW(self.btnDelete, self.OnMouseExit)
      wx.EVT_ENTER_WINDOW(self.btnNewFolder, self.OnMouseOver)
      wx.EVT_LEAVE_WINDOW(self.btnNewFolder, self.OnMouseExit)
      wx.EVT_ENTER_WINDOW(self.btnUpdateDB, self.OnMouseOver)
      wx.EVT_LEAVE_WINDOW(self.btnUpdateDB, self.OnMouseExit)
      wx.EVT_ENTER_WINDOW(self.btnConnect, self.OnMouseOver)
      wx.EVT_LEAVE_WINDOW(self.btnConnect, self.OnMouseExit)
      wx.EVT_ENTER_WINDOW(self.btnRefresh, self.OnMouseOver)
      wx.EVT_LEAVE_WINDOW(self.btnRefresh, self.OnMouseExit)
      wx.EVT_ENTER_WINDOW(self.btnClose, self.OnMouseOver)
      wx.EVT_LEAVE_WINDOW(self.btnClose, self.OnMouseExit)
      wx.EVT_ENTER_WINDOW(self.btnHelp, self.OnMouseOver)
      wx.EVT_LEAVE_WINDOW(self.btnHelp, self.OnMouseExit)

      # Define the Event to initiate the beginning of a File Drag operation
      wx.EVT_RIGHT_DOWN(self.sourceFile, self.OnFileRightDown)
      wx.EVT_RIGHT_DOWN(self.destFile, self.OnFileRightDown)

      # Button Events (link functionality to button widgets)
      wx.EVT_BUTTON(self, ID_BTNCOPY, self.OnCopyMove)
      wx.EVT_BUTTON(self, ID_BTNMOVE, self.OnCopyMove)
      wx.EVT_BUTTON(self, ID_BTNDELETE, self.OnDelete)
      wx.EVT_BUTTON(self, ID_BTNNEWFOLDER, self.OnNewFolder)
      wx.EVT_BUTTON(self, ID_BTNUPDATEDB, self.OnUpdateDB)
      wx.EVT_BUTTON(self, ID_BTNCONNECT, self.OnConnect)
      wx.EVT_BUTTON(self, ID_BTNREFRESH, self.OnRefresh)
      wx.EVT_BUTTON(self, ID_BTNCLOSE, self.CloseWindow)
      wx.EVT_BUTTON(self, ID_BTNHELP, self.OnHelp)

      # Directory Tree, SRB Collection and File List Focus Events
      wx.EVT_SET_FOCUS(self.sourceDir.GetTreeCtrl(), self.OnSetFocus)
      wx.EVT_SET_FOCUS(self.destDir.GetTreeCtrl(), self.OnSetFocus)
      wx.EVT_SET_FOCUS(self.srbSourceDir, self.OnSetFocus)
      wx.EVT_SET_FOCUS(self.srbDestDir, self.OnSetFocus)
      wx.EVT_SET_FOCUS(self.sourceFile, self.OnSetFocus)
      wx.EVT_SET_FOCUS(self.destFile, self.OnSetFocus)
      
      # Directory Tree and SRB Collection Selection Events
      wx.EVT_TREE_SEL_CHANGED(self, self.sourceDir.GetTreeCtrl().GetId(), self.OnDirSelect)
      wx.EVT_TREE_SEL_CHANGED(self, self.destDir.GetTreeCtrl().GetId(), self.OnDirSelect)
      wx.EVT_TREE_SEL_CHANGED(self, self.srbSourceDir.GetId(), self.OnDirSelect)
      wx.EVT_TREE_SEL_CHANGED(self, self.srbDestDir.GetId(), self.OnDirSelect)

      # File List Selection Events
      wx.EVT_LIST_ITEM_SELECTED(self, self.sourceFile.GetId(), self.OnListItem)
      wx.EVT_LIST_ITEM_SELECTED(self, self.destFile.GetId(), self.OnListItem)
      wx.EVT_LIST_ITEM_DESELECTED(self, self.sourceFile.GetId(), self.OnListItem)
      wx.EVT_LIST_ITEM_DESELECTED(self, self.destFile.GetId(), self.OnListItem)

      # File Filter Events
      wx.EVT_CHOICE(self, ID_SOURCEFILTER, self.OnDirSelect)
      wx.EVT_CHOICE(self, ID_DESTFILTER, self.OnDirSelect)

      # File Manager Close Event
      wx.EVT_CLOSE(self, self.CloseWindow)

      # lay out all the controls
      self.Layout()
      # Enable Auto Layout
      self.SetAutoLayout(True)
      self.CenterOnScreen()
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
      # Break the connection to the SRB.
      if "__WXMSW__" in wx.PlatformInfo:
          srb = ctypes.cdll.srbClient
      else:
          srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
      srb.srb_disconnect(self.srbConnectionID.value)
      
      # Signal that the Connection has been broken by resetting the ConnectionID value to None
      self.srbConnectionID = None
         
      # print "Disconnect\n\n"

   def CloseWindow(self, event):
      """ Clean up upon closing the File Management Window """
      try:
         # If we are connected to the SRB, we need to disconnect before closing the window
         if self.srbConnectionID != None:
            self.srbDisconnect()
      except:
         (exctype, excvalue) = sys.exc_info()[:2]
         print "File Management Exception in CloseWindow(): %s, %s" % (exctype, excvalue)
      # Destroy the SRB Connection Dialog 
      self.SRBConnDlg.Destroy()
      # Destroy the File Management Dialog
      self.Destroy()

   def OnMouseOver(self, event):
      """ Update the Status Bar when a button is moused-over """
      # Status Text can be customized based on the values of the current Active controls.
      # Therefore, we need to know which controls are active.
      if self.activeSide == 'Source':
         sourceLbl = self.lblSource
         # The path differs for local and remote file systems
         if sourceLbl.GetLabel() == self.LOCAL_LABEL:
            sourceDir = self.sourceDir.GetPath()
         else:
            sourceDir = self.GetFullPath(self.srbSourceDir)
         sourceFile = self.sourceFile
         # The path differs for local and remote file systems
         if self.lblDest.GetLabel() == self.LOCAL_LABEL:
            targetDir = self.destDir.GetPath()
         else:
            targetDir = self.GetFullPath(self.srbDestDir)
      elif self.activeSide == 'Dest':
         sourceLbl = self.lblDest
         # The path differs for local and remote file systems
         if sourceLbl.GetLabel() == self.LOCAL_LABEL:
            sourceDir = self.destDir.GetPath()
         else:
            sourceDir = self.GetFullPath(self.srbDestDir)
         sourceFile = self.destFile
         # The path differs for local and remote file systems
         if self.lblSource.GetLabel() == self.LOCAL_LABEL:
            targetDir = self.sourceDir.GetPath()
         else:
            targetDir = self.GetFullPath(self.srbSourceDir)

      # Determine which button is being moused-over and display an appropriate message in the
      # Status Bar

      # Copy button
      if event.GetId() == ID_BTNCOPY:
         # Indicate whether selected files or all files will be copied
         if sourceFile.GetSelectedItemCount() == 0:
            tempText = _('all')
         else:
            tempText = _('selected')
         # Specify source and destination Folder/Collections
         if 'unicode' in wx.PlatformInfo:
             # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
             tempText = unicode(tempText, 'utf8')
             prompt = unicode(_('Copy %s files from %s to %s'), 'utf8')
         else:
             prompt = _('Copy %s files from %s to %s')
         self.SetStatusText(prompt % (tempText, sourceDir, targetDir))

      # Move button
      elif event.GetId() == ID_BTNMOVE:
         # Indicate whether selected files or all files will be copied
         if sourceFile.GetSelectedItemCount() == 0:
            tempText = _('all')
         else:
            tempText = _('selected')
         # Specify source and destination Folder/Collections
         if 'unicode' in wx.PlatformInfo:
             # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
             tempText = unicode(tempText, 'utf8')
             prompt = unicode(_('Move %s files from %s to %s'), 'utf8')
         else:
             prompt = _('Move %s files from %s to %s')
         self.SetStatusText(prompt % (tempText, sourceDir, targetDir))

      # Delete button
      elif event.GetId() == ID_BTNDELETE:
         if 'unicode' in wx.PlatformInfo:
             # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
             prompt = unicode(_('Delete selected files from %s'), 'utf8')
         else:
             prompt = _('Delete selected files from %s')
         self.SetStatusText(prompt % sourceDir)

      # New Folder button
      elif event.GetId() == ID_BTNNEWFOLDER:
         if 'unicode' in wx.PlatformInfo:
             # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
             prompt = unicode(_('Create new folder in %s'), 'utf8')
         else:
             prompt = _('Create new folder in %s')
         self.SetStatusText(prompt % sourceDir)

      # Update DB button
      elif event.GetId() == ID_BTNUPDATEDB:
         self.SetStatusText(_('Update file locations in the Transana database'))

      # Connect button
      elif event.GetId() == ID_BTNCONNECT:
         # Status Bar Text needs to indicate whether this button will establish or break
         # the connection to the SRB
         if sourceLbl.GetLabel() == self.LOCAL_LABEL:
            self.SetStatusText(_('Connect to SRB'))
         else:
            self.SetStatusText(_('Disconnect from SRB'))

      # Refresh button
      elif event.GetId() == ID_BTNREFRESH:
         self.SetStatusText(_('Refresh File Lists'))

      # Close button
      elif event.GetId() == ID_BTNCLOSE:
         self.SetStatusText(_('Close File Management'))

      # Help button
      elif event.GetId() == ID_BTNHELP:
         self.SetStatusText(_('View File Management Help'))

   def OnMouseExit(self, event):
      """ Clear the Status Bar when the Mouse is no longer over a Button """
      self.SetStatusText('')

   def OnFileRightDown(self, event):
      """ Mouse Right-button Down event initiates Drag for Drag-and-Drop file copying """

      # NOTE:  Drag will only work from the Local file system.  This is because SRB files
      #        can't be treated as File objects by a FileDropTarget.

      # Determine which side is Source, which side is Target
      if event.GetId() == ID_SOURCEFILE:
         sourceLbl = self.lblSource.GetLabel()
         # Local and Remote File Systems have different methods for determing the selected Path
         if sourceLbl == self.LOCAL_LABEL:
            sourceDir = self.sourceDir.GetPath()
         else:
            return       # If it's not local, we're not dragging anything!
         sourceFile = self.sourceFile
         targetDir = self.destDir.GetPath()
      elif event.GetId() == ID_DESTFILE:
         sourceLbl = self.lblDest.GetLabel()
         # Local and Remote File Systems have different methods for determing the selected Path
         if sourceLbl == self.LOCAL_LABEL:
            sourceDir = self.destDir.GetPath()
         else:
            return       # If it's not local, we're not dragging anything!
         sourceFile = self.destFile
         targetDir = self.sourceDir.GetPath()
      else:
         return

      # We only drag if File Names have been selected
      if (sourceFile.GetSelectedItemCount() > 0):

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
         for file in fileList:
            tdo.AddFile(file)
         # Create a Drop Source Object, which enables the Drag operation
         tds = wx.DropSource(sourceFile)
         # Associate the Data to be dragged with the Drop Source Object
         tds.SetData(tdo)
         # Intiate the Drag Operation
         tds.DoDragDrop(True)
         # Refresh the File Windows in case something was moved
         self.OnRefresh(event)


   def OnCopyMove(self, event):
      """ Copy or Move files between sides of the File Management Window """
      # If the Move button called this event, we have a Move request.  Otherwise we
      # have a Copy request.  This is signalled by moveFlag.
      if event.GetId() == ID_BTNMOVE:
         moveFlag = True
      else:
         moveFlag = False
      # Map Window widgets to the appropriate local objects for sending and receiving files.
      if self.activeSide == 'Source':
         sourceLbl = self.lblSource.GetLabel()
         # Local and Remote File Systems have different methods for determing the selected Path
         if sourceLbl == self.LOCAL_LABEL:
            sourceDir = self.sourceDir.GetPath()
         else:
            sourceDir = self.GetFullPath(self.srbSourceDir)
         sourceFile = self.sourceFile
         sourceFilter = self.sourceFilter
         targetLbl = self.lblDest.GetLabel()
         # Local and Remote File Systems have different methods for determing the selected Path
         if targetLbl == self.LOCAL_LABEL:
            targetDir = self.destDir.GetPath()
         else:
            targetDir = self.GetFullPath(self.srbDestDir)
         targetFile = self.destFile
         targetFilter = self.destFilter
      elif self.activeSide == 'Dest':
         sourceLbl = self.lblDest.GetLabel()
         # Local and Remote File Systems have different methods for determing the selected Path
         if sourceLbl == self.LOCAL_LABEL:
            sourceDir = self.destDir.GetPath()
         else:
            sourceDir = self.GetFullPath(self.srbDestDir)
         sourceFile = self.destFile
         sourceFilter = self.destFilter
         targetLbl = self.lblSource.GetLabel()
         # Local and Remote File Systems have different methods for determing the selected Path
         if targetLbl == self.LOCAL_LABEL:
            targetDir = self.sourceDir.GetPath()
         else:
            targetDir = self.GetFullPath(self.srbSourceDir)
         targetFile = self.sourceFile
         targetFilter = self.sourceFilter

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
      while True:
         # Get the next Selected item in the List 
         itemNum = sourceFile.GetNextItem(itemNum, wx.LIST_NEXT_ALL, state)
         # If there are no more selected items in the list, itemNum is set to -1
         if itemNum == -1:
            break
         else:
            # Add the Selected item to the fileList
            fileList.append(sourceFile.GetItemText(itemNum))

      # Set the Cursor to the Hourglass
      self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))

      if "__WXMSW__" in wx.PlatformInfo:
          srb = ctypes.cdll.srbClient
      else:
          srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

      # Copy or Move from Local Drive to Local Drive
      if (sourceLbl == self.LOCAL_LABEL) and \
         (targetLbl == self.LOCAL_LABEL):

         # For each file in the fileList ...
         for file in fileList:
            # Build the full File Name by joining the source path with the file name from the list
            fileName = os.path.join(sourceDir, file)
            # If we want to MOVE the file ...
            if moveFlag:
               if 'unicode' in wx.PlatformInfo:
                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                   prompt = unicode(_('Moving %s to %s'), 'utf8')
               else:
                   prompt = _('Moving %s to %s')
               self.SetStatusText(prompt % (fileName, targetDir))
               # os.Rename accomplishes a fast Move (at least on Windows)  TODO:  Test this on Mac!
               os.rename(fileName, os.path.join(targetDir, file))
               self.SetStatusText(_('Move complete.'))
            # ... otherwise we want to copy the file
            else:
               if 'unicode' in wx.PlatformInfo:
                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                   prompt = unicode(_('Copying %s to %s'), 'utf8')
               else:
                   prompt = _('Copying %s to %s')
               self.SetStatusText(prompt % (fileName, targetDir))
               # shutil is the most efficient way I have found to copy a file on the local file system
               shutil.copyfile(fileName, os.path.join(targetDir, file))
               self.SetStatusText(_('Copy complete.'))
            # Let's update the target control after every file so that the new files show up in the list ASAP.
            self.RefreshFileList(targetLbl, targetDir, targetFile, targetFilter)
            # wxYield allows the Windows Message Queue to update the display.
            wx.Yield()

      # Copy or Move from Local Drive to SRB Collection
      elif (sourceLbl == self.LOCAL_LABEL) and \
           (targetLbl == self.REMOTE_LABEL):

         # For each file in the fileList ...
         for file in fileList:
            # Build the full File Name by joining the source path with the file name from the list
            fileName = os.path.join(sourceDir, file)
            # Get the File Size
            fileSize = os.path.getsize(fileName)
            # If we want to MOVE the file ...
            if moveFlag:
               if 'unicode' in wx.PlatformInfo:
                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                   prompt = unicode(_('Moving %s to %s'), 'utf8')
               else:
                   prompt = _('Moving %s to %s')
               self.SetStatusText(prompt % (fileName, targetDir))
            # ... otherwise we want to copy the file
            else:
               if 'unicode' in wx.PlatformInfo:
                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                   prompt = unicode(_('Copying %s to %s'), 'utf8')
               else:
                   prompt = _('Copying %s to %s')
               self.SetStatusText(prompt % (fileName, targetDir))

            # The SRBFileTransfer class handles file transfers and provides Progress Feedback
            dlg = SRBFileTransfer.SRBFileTransfer(self, _("SRB File Transfer"), file, fileSize, sourceDir, self.srbConnectionID, targetDir, SRBFileTransfer.srb_UPLOAD)
            success = dlg.TransferSuccessful()
            dlg.Destroy()

            if moveFlag:
                if success:
                    self.DeleteFile(sourceLbl, sourceDir, file)
                    self.SetStatusText(_('Move complete.'))
                else:
                    self.SetStatusText(_('Move cancelled.'))
            else:
                if success:
                    self.SetStatusText(_('Copy complete.'))
                else:
                    self.SetStatusText(_('Copy cancelled.'))
            # Let's update the target control after every file so that the new files show up in the list ASAP.
            self.RefreshFileList(targetLbl, targetDir, targetFile, targetFilter)
            # wxYield allows the Windows Message Queue to update the display.
            wx.Yield()

      # Copy or Move from SRB Collection to Local Drive
      elif (sourceLbl == self.REMOTE_LABEL) and \
           (targetLbl == self.LOCAL_LABEL):

         # For each file in the fileList ...
         for file in fileList:
            # Build the full File Name by joining the source path with the file name from the list
            fileName = os.path.join(sourceDir, file)
            # If we want to MOVE the file ...
            if moveFlag:
               if 'unicode' in wx.PlatformInfo:
                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                   prompt = unicode(_('Moving %s to %s'), 'utf8')
               else:
                   prompt = _('Moving %s to %s')
               self.SetStatusText(prompt % (fileName, targetDir))
            # ... otherwise we want to copy the file
            else:
               if 'unicode' in wx.PlatformInfo:
                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                   prompt = unicode(_('Copying %s to %s'), 'utf8')
               else:
                   prompt = _('Copying %s to %s')
               self.SetStatusText(prompt % (fileName, targetDir))

            # Find the File Size
            buf = ' ' * 25
            srb.srb_get_dataset_system_metadata(2, sourceFile.FindItem(-1, file), buf, 25)
      
            # Strip whitespace and null character (c string terminator) from buf
            buf = string.strip(buf)[:-1]
            # The SRBFileTransfer class handles file transfers and provides Progress Feedback
            dlg = SRBFileTransfer.SRBFileTransfer(self, _("SRB File Transfer"), file, int(buf), targetDir, self.srbConnectionID, sourceDir, SRBFileTransfer.srb_DOWNLOAD)
            success = dlg.TransferSuccessful()
            dlg.Destroy()

            if moveFlag:
                if success:
                    self.DeleteFile(sourceLbl, sourceDir, file)
                    self.SetStatusText(_('Move complete.'))
                else:
                    self.SetStatusText(_('Move cancelled.'))
            else:
                if success:
                    self.SetStatusText(_('Copy complete.'))
                else:
                    self.SetStatusText(_('Copy cancelled.'))
            # Let's update the target control after every file so that the new files show up in the list ASAP.
            self.RefreshFileList(targetLbl, targetDir, targetFile, targetFilter)
            # wxYield allows the Windows Message Queue to update the display.
            wx.Yield()

      # Copy or Move from SRB Collection to SRB Collection
      elif (sourceLbl == self.REMOTE_LABEL) and \
           (targetLbl == self.REMOTE_LABEL):

         # For each file in the fileList ...
         for file in fileList:
            # Build the full File Name by joining the source path with the file name from the list
            fileName = sourceDir + '/' + file
            # If we want to MOVE the file ...
            if moveFlag:
               if 'unicode' in wx.PlatformInfo:
                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                   prompt = unicode(_('Moving %s to %s'), 'utf8')
               else:
                   prompt = _('Moving %s to %s')
               self.SetStatusText(prompt % (fileName, targetDir))
            # ... otherwise we want to copy the file
            else:
               if 'unicode' in wx.PlatformInfo:
                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                   prompt = unicode(_('Copying %s to %s'), 'utf8')
               else:
                   prompt = _('Copying %s to %s')
               self.SetStatusText(prompt % (fileName, targetDir))

            if 'unicode' in wx.PlatformInfo:
                tmpFile = file.encode(TransanaGlobal.encoding)
                tmpSourceDir = sourceDir.encode(TransanaGlobal.encoding)
                tmpTargetDir = targetDir.encode(TransanaGlobal.encoding)
            else:
                tmpFile = file
                tmpSourceDir = sourceDir
                tmpTargetDir = targetDir

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
      
            if Result < 0:
               self.srbErrorMessage(Result)
               
            if moveFlag:
               self.DeleteFile(sourceLbl, sourceDir, file)
               self.SetStatusText(_('Move complete.'))
            else:
               self.SetStatusText(_('Copy complete.'))
            # Let's update the target control after every file so that the new files show up in the list ASAP.
            self.RefreshFileList(targetLbl, targetDir, targetFile, targetFilter)
            # wxYield allows the Windows Message Queue to update the display.
            wx.Yield()

      # Reset the cursor to the Arrow
      self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

      # If we have moved files, we need to refresh the Source file list so that the files that were moved
      # no longer appear in the list.
      if moveFlag:
         self.RefreshFileList(sourceLbl, sourceDir, sourceFile, sourceFilter)


   def DeleteFile(self, lbl, path, filename):
      # Set cursor to hourglass
      self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
      if "__WXMSW__" in wx.PlatformInfo:
          srb = ctypes.cdll.srbClient
      else:
          srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
      # Get the full name of the file to be deleted by adding the file path to it
      if lbl == self.LOCAL_LABEL:
         fileName = os.path.join(path, filename)
         if 'unicode' in wx.PlatformInfo:
             # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
             prompt = unicode(_('Deleting %s'), 'utf8')
         else:
             prompt = _('Deleting %s')
         self.SetStatusText(prompt % fileName)
         try:
             # Delete the file
             os.remove(fileName)
         except:
             dlg = Dialogs.ErrorDialog(self, "%s" % sys.exc_info()[1])
             dlg.ShowModal()
             dlg.Destroy()
      else:
         fileName = path + '/' + filename
         if 'unicode' in wx.PlatformInfo:
             # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
             prompt = unicode(_('Deleting %s'), 'utf8')
         else:
             prompt = _('Deleting %s')
         self.SetStatusText(prompt % fileName)

         if 'unicode' in wx.PlatformInfo:
             tmpFilename = filename.encode(TransanaGlobal.encoding)
             tmpPath = path.encode(TransanaGlobal.encoding)
         else:
             tmpFilename = filename
             tmpPath = path
         delResult = srb.srb_remove_obj(self.srbConnectionID, tmpFilename, tmpPath, 0)
      
         if delResult < 0:
            self.srbErrorMessage(delResult)
      self.SetStatusText(_('Delete complete.'))
      # Set the cursor back to the arrow pointer
      self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
      # wxYield allows Windows Messages to process, thus allowing the controls to refresh
      wx.Yield()

   def OnDelete(self, event):
      """ Delete selected file(s) """

      # Determine which side is active and set local controls accordingly
      if self.activeSide == 'Source':
         sourceLbl = self.lblSource
         if sourceLbl.GetLabel() == self.LOCAL_LABEL:
            sourceDir = self.sourceDir
            sourcePath = sourceDir.GetPath()
         else:
            sourceDir = self.srbSourceDir
            sourcePath = self.GetFullPath(sourceDir)
         sourceFile = self.sourceFile
         sourceFilter = self.sourceFilter
         # We also need to know the path for the inactive side, so it can be updated too if needed
         otherLbl = self.lblDest
         if otherLbl.GetLabel() == self.LOCAL_LABEL:
            other = self.destDir
            otherPath = self.destDir.GetPath()
         else:
            other = self.srbDestDir
            otherPath = self.GetFullPath(self.srbDestDir)
         otherFile = self.destFile
         otherFilter = self.destFilter
      elif self.activeSide == 'Dest':
         sourceLbl = self.lblDest
         if sourceLbl.GetLabel() == self.LOCAL_LABEL:
            sourceDir = self.destDir
            sourcePath = sourceDir.GetPath()
         else:
            sourceDir = self.srbDestDir
            sourcePath = self.GetFullPath(sourceDir)
         sourceFile = self.destFile
         sourceFilter = self.destFilter
         # We also need to know the path for the inactive side, so it can be updated too if needed
         otherLbl = self.lblSource
         if otherLbl.GetLabel() == self.LOCAL_LABEL:
            other = self.sourceDir
            otherPath = self.sourceDir.GetPath()
         else:
            other = self.srbSourceDir
            otherPath = self.GetFullPath(self.srbSourceDir)
         otherFile = self.sourceFile
         otherFilter = self.sourceFilter
         
      if "__WXMSW__" in wx.PlatformInfo:
          srb = ctypes.cdll.srbClient
      else:
          srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

      # If files are listed, delete files.  Otherwise, remove the folder!
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
      else:
         # We want to remove a folder, not a file, if we come here.
         if 'unicode' in wx.PlatformInfo:
             # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
             prompt = unicode(_('Removing %s'), 'utf8')
         else:
             prompt = _('Removing %s')
         self.SetStatusText(prompt % sourcePath)
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
                other.SetPath(sourceDir.GetPath())
                itemId = other.GetTreeCtrl().GetSelection()  # From Robin Dunn
                other.GetTreeCtrl().CollapseAndReset(itemId) # From Robin Dunn -- Forces refresh of the tree node.
                if otherPath != sourcePath:
                    other.SetPath(otherPath)
         else:
            # Get the Parent Item
            parentItem = sourceDir.GetItemParent(sourceDir.GetSelection())
            # Delete the Current (Selected) Item from the TreeCtrl (Does not remove it from the File System)
            sourceDir.Delete(sourceDir.GetSelection())
            # Update the Selection in the Tree to point to the Parent Item
            sourceDir.SelectItem(parentItem)
            if 'unicode' in wx.PlatformInfo:
                tmpSourcePath = sourcePath.encode(TransanaGlobal.encoding)
            else:
                tmpSourcePath = sourcePath
            RemoveResult = srb.srb_remove_collection(self.srbConnectionID, 0, tmpSourcePath);
      
            if RemoveResult < 0:
               self.srbErrorMessage(RemoveResult)
            if sourceLbl.GetLabel() == otherLbl.GetLabel():
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

      # If both sides are pointed to the same folder, we need to update the other side as well.  
      if sourcePath == otherPath:
          # That path no longer exists!
          if otherLbl.GetLabel() == self.LOCAL_LABEL:
              otherPath = other.GetPath()
          else:
              otherPath = self.GetFullPath(other)
          # Refresh the file list now that all files have been deleted
          self.RefreshFileList(otherLbl.GetLabel(), otherPath, otherFile, otherFilter)
          
      # When this is complete, the Delete button should be disabled, as there are no "selected" files,
      # unless there are no files left, in which case we can delete the folder

      # The Source path no longer exists!
      if sourceLbl.GetLabel() == self.LOCAL_LABEL:
          sourcePath = sourceDir.GetPath()
      else:
          sourcePath = self.GetFullPath(sourceDir)
      if self.IsEmpty(sourceLbl.GetLabel(), sourcePath):
         if self.activeSide == 'Source':
            self.btnDelete.SetLabel(_('<< Delete Folder'))
         else:
            self.btnDelete.SetLabel(_('Delete Folder >>'))
         self.btnDelete.Enable(True)
      else:
         self.btnDelete.Enable(False)

   def OnUpdateDB(self, event):
       """ Handle "Update DB" Button Press """
       # Determine the File Path Path and File Control
       if self.activeSide == 'Source':
           filePath = self.sourceDir.GetPath()
           control = self.sourceFile
       else:
           filePath = self.destDir.GetPath()
           control = self.destFile
       # Initialize "item" to -1 so that the first Selected Item will be returned
       item = -1
       # Initialize the File List as empty
       fileList = []
       # Find all selected Items in the appropriate File Control
       while 1:
           # Get the next Selected Item in the File Control
           item = control.GetNextItem(item, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
           # If there is no next item...
           if item == -1:
               # ... we can stop looking
               break
           # Macs requires an odd conversion.
           if 'wxMac' in wx.PlatformInfo:
               file = Misc.convertMacFilename(control.GetItemText(item))
           else:
               file = control.GetItemText(item)
           # Add the item to the File List.
           fileList.append(file)
           
       # Update all listed Files in the Database with the new File Path
       if not DBInterface.UpdateDBFilenames(self, filePath, fileList):
           infodlg = Dialogs.InfoDialog(self, _('Update Failed.  Some records that would be affected may be locked by another user.'))
           infodlg.ShowModal()
           infodlg.Destroy()

   def srbAddChildNode(self, srbDir, treeNode, srbCollectionName):
      """ This method populates a tree node in the SRB Directory Tree """

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
      # Get the list of sub-Collections from the SRB
      if "__WXMSW__" in wx.PlatformInfo:
          srb = ctypes.cdll.srbClient
      else:
          srb = ctypes.cdll.LoadLibrary("srbClient.dylib")
      srbCatalogCount = srb.srb_query_subcolls(self.srbConnectionID.value, 0, srbCollectionName)
      
      # Check to see if Sub-Collections are returned or if an error code (negative result) is generated
      if srbCatalogCount < 0:
         self.srbErrorMessage(srbCatalogCount)
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
      if self.activeSide == "Source":
         sourceLbl = self.lblSource
         sourceDir = self.sourceDir
         sourceSrbDir = self.srbSourceDir
         sourceFile = self.sourceFile
         sourceFilter = self.sourceFilter
      else:
         sourceLbl = self.lblDest
         sourceDir = self.destDir
         sourceSrbDir = self.srbDestDir
         sourceFile = self.destFile
         sourceFilter = self.destFilter

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
         # Change the button label to indicate that it will implement a "Connect" if pressed again
         if self.activeSide == "Source":
            self.btnConnect.SetLabel(_('<< Connect'))
         else:
            self.btnConnect.SetLabel(_('Connect >>'))
         # Refresh the File List to now display the appropriate files from either the Local File System or the SRB as appropriate
         if sourceDir.GetPath() != '':
             self.RefreshFileList(sourceLbl.GetLabel(), sourceDir.GetPath(), sourceFile, sourceFilter)
         else:
             sourceFile.ClearAll()
         # Check to see if we need to disconnect, and do so if needed.  (If both sides are "local", we should disconnect.)
         if (self.lblSource.GetLabel() == self.LOCAL_LABEL) and \
            (self.lblDest.GetLabel() == self.LOCAL_LABEL):
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

               if 'unicode' in wx.PlatformInfo:
                   self.tmpCollection = self.srbCollection.encode(TransanaGlobal.encoding)
                   tmpDomain = self.SRBConnDlg.editDomain.GetValue().encode(TransanaGlobal.encoding)
                   self.tmpResource = self.SRBConnDlg.editSRBResource.GetValue().encode(TransanaGlobal.encoding)
                   tmpSEAOption = self.SRBConnDlg.editSRBSEAOption.GetValue().encode(TransanaGlobal.encoding)
                   tmpPort = self.SRBConnDlg.editSRBPort.GetValue().encode(TransanaGlobal.encoding)
                   tmpHost = self.SRBConnDlg.editSRBHost.GetValue().encode(TransanaGlobal.encoding)
                   tmpUserName = self.SRBConnDlg.editUserName.GetValue().encode(TransanaGlobal.encoding)
                   tmpPassword = self.SRBConnDlg.editPassword.GetValue().encode(TransanaGlobal.encoding)
               else:
                   self.tmpCollection = self.srbCollection
                   tmpDomain = self.SRBConnDlg.editDomain.GetValue()
                   tmpHost = self.SRBConnDlg.editSRBHost.GetValue()
                   tmpSEAOption = self.SRBConnDlg.editSRBSEAOption.GetValue()
                   tmpPort = self.SRBConnDlg.editSRBPort.GetValue()
                   self.tmpResource = self.SRBConnDlg.editSRBResource.GetValue()
                   tmpUserName = self.SRBConnDlg.editUserName.GetValue()
                   tmpPassword = self.SRBConnDlg.editPassword.GetValue()

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
               else:
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
            # Change the button label to indicate that it will implement a "Disconnect" if pressed again         
            if self.activeSide == "Source":
               self.btnConnect.SetLabel(_('<< Disconnect'))
            else:
               self.btnConnect.SetLabel(_('Disconnect >>'))

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


   def OnNewFolder(self, event):
      """" Implements the "New Folder" button. """

      # Determine which side of the File Management Window we are acting on and identify the control to manipulate
      if self.activeSide == 'Source':
         sourceLbl = self.lblSource.GetLabel()
         if sourceLbl == self.LOCAL_LABEL:
            target = self.sourceDir
            path = target.GetPath()
         else:
            target = self.srbSourceDir
            path = self.GetFullPath(target)
         otherLbl = self.lblDest.GetLabel()
         if otherLbl == self.LOCAL_LABEL:
            other = self.destDir
         else:
            other = self.srbDestDir
      elif self.activeSide == 'Dest':
         sourceLbl = self.lblDest.GetLabel()
         if sourceLbl == self.LOCAL_LABEL:
            target = self.destDir
            path = target.GetPath()
         else:
            target = self.srbDestDir
            path = self.GetFullPath(target)
         otherLbl = self.lblSource.GetLabel()
         if otherLbl == self.LOCAL_LABEL:
            other = self.sourceDir
         else:
            other = self.srbSourceDir
      if 'unicode' in wx.PlatformInfo:
          # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
          prompt = unicode(_('Create new folder for %s'), 'utf8')
      else:
          prompt = _('Create new folder for %s')
      self.SetStatusText(prompt % path)
      
      if "__WXMSW__" in wx.PlatformInfo:
          srb = ctypes.cdll.srbClient
      else:
          srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

      # Create a TextEntry Dialog to get the name for the new Folder from the user
      newFolder = wx.TextEntryDialog(self, prompt % path, _('Create new folder'), style = wx.OK | wx.CANCEL | wx.CENTRE)
      # Display the Dialog and see if the user presses OK
      if newFolder.ShowModal() == wx.ID_OK:
         if sourceLbl == self.LOCAL_LABEL:
            # Build the full folder name by combining the current path and the text from the user
            folderName = os.path.join(path, newFolder.GetValue())
            # Create the folder on the local file system
            os.mkdir(folderName)
            # Update the Directory Tree Control
            itemId = target.GetTreeCtrl().GetSelection()  # From Robin Dunn
            target.GetTreeCtrl().CollapseAndReset(itemId) # From Robin Dunn -- Forces refresh of the tree node.
            if sourceLbl == otherLbl:
                otherPath = other.GetPath()
                other.SetPath(path)
                itemId = other.GetTreeCtrl().GetSelection()  # From Robin Dunn
                other.GetTreeCtrl().CollapseAndReset(itemId) # From Robin Dunn -- Forces refresh of the tree node.
                other.SetPath(otherPath)
                
            # Automatically change to the newly created folder
            target.SetPath(folderName)
         else:
            # Encode if needed
            if 'unicode' in wx.PlatformInfo:
                tmpNewFolder = newFolder.GetValue().encode(TransanaGlobal.encoding)
                path = path.encode(TransanaGlobal.encoding)
            else:
                tmpNewFolder = newFolder.GetValue()
            # Create the folder on the SRB file system
            NewCollection = srb.srb_new_collection(self.srbConnectionID, 0, path, tmpNewFolder);

            if NewCollection >= 0:
                # Update the Directory Tree Control
                itemId = target.AppendItem(target.GetSelection(), newFolder.GetValue())

                if sourceLbl == otherLbl:
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
            else:
                self.srbErrorMessage(NewCollection)
      # Destroy the TextEntry Dialog, now that we are done with it.
      newFolder.Destroy()

   def OnRefresh(self, event):
      """ Method to use when the Refresh Button is pressed.  """
      sourceLbl = self.lblSource.GetLabel()
      if sourceLbl == self.LOCAL_LABEL:
         path = self.sourceDir.GetPath()
      else:
         path = self.GetFullPath(self.srbSourceDir)
      if path != '':
         # if a path is defined on the left-hand side, refresh the left hand side controls
         self.RefreshFileList(sourceLbl, path, self.sourceFile, self.sourceFilter)
      destLbl = self.lblDest.GetLabel()
      if destLbl == self.LOCAL_LABEL:
         path = self.destDir.GetPath()
      else:
         path = self.GetFullPath(self.srbDestDir)
      if path != '':
         # if a path is defined on the right-hand side, refresh the right hand side controls
         self.RefreshFileList(destLbl, path, self.destFile, self.destFilter)

   def OnHelp(self, event):
        """ Method to use when the Help Button is pressed """
        helpContext = 'File Management Utility'
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
            
            file = open(os.getenv("HOME") + '/TransanaHelpContext.txt', 'w')
            pickle.dump(helpContext, file)
            file.flush()
            file.close()            
            
            os.system('open -a TransanaHelp.app')

        else:
            # Make the Help call differently from Python and the stand-alone executable.
            if fn.lower() == 'transana.py':
                # for within Python, we call python, then the Help code and the context
                os.spawnv(os.P_NOWAIT, 'python', [path + 'Help.py', helpContext])
            else:
                # The Standalone requires a "dummy" parameter here (Help), as sys.argv differs between the two versions.
                os.spawnv(os.P_NOWAIT, path + 'Help', ['Help', helpContext])

   def OnSetFocus(self, event):
      """ This event is triggered when any of several controls receives program focus.  It is primarily used
          for enabling and disabling control buttons. """
      # Assume UpdateDB is NOT valid
      updateFiles = False
      # If a control on the left side (the Source side) is selected, ...
      if (event.GetId() in [ID_SRBSOURCEDIR, ID_SOURCEFILE]) or \
         (event.GetEventObject() == self.sourceDir.GetTreeCtrl()):

         # If the user has selected an invalid device (eg. CD drive with no CD), this is signalled by
         # a blank path, in which case we should skip this event
         if self.sourceDir.GetPath() == '':
             return
            
         # Set the global indicator of which side is active
         self.activeSide = 'Source'
         # Change the button labels accordingly
         self.btnCopy.SetLabel(_('>> Copy >>'))
         self.btnMove.SetLabel(_('>> Move >>'))
         self.btnDelete.SetLabel(_('<< Delete'))
         self.btnNewFolder.SetLabel(_('<< New Folder'))
         self.btnUpdateDB.SetLabel(_('<< Update DB'))
         # We need to know the active Label
         lbl = self.lblSource.GetLabel()
         # We need to know the active File Control
         ctrl = self.sourceFile
         # We need to know if the source is local or the SRB.
         if lbl == self.LOCAL_LABEL:
            # If we are local, offer to connect to the SRB
            self.btnConnect.SetLabel(_('<< Connect'))
            # The source Path we need is the local path
            sourcePath = self.sourceDir.GetPath()
         else:
            # If we are connected to the SRB, offer to disconnect from the SRB
            self.btnConnect.SetLabel(_('<< Disconnect'))
            # The source Path we need is the SRB Path
            sourcePath = self.GetFullPath(self.srbSourceDir)
         # Determine if any items are selected, which enables the
         # Update DB and Delete buttons
         if self.sourceFile.GetSelectedItemCount() > 0:
            updateFiles = True
         # What destination path to use depends on whether the other side is local or SRB
         if self.lblDest.GetLabel() == self.LOCAL_LABEL:
            destPath = self.destDir.GetPath()
         else:
            destPath = self.GetFullPath(self.srbDestDir)
      # If a control on the right side (the Destination side) is selected, ...
      elif (event.GetId() in [ID_SRBDESTDIR, ID_DESTFILE]) or \
           (event.GetEventObject() == self.destDir.GetTreeCtrl()):

         # If the user has selected an invalid device (eg. CD drive with no CD), this is signalled by
         # a blank path, in which case we should skip this event
         if self.destDir.GetPath() == '':
             return
            
         # Set the global indicator of which side is active
         self.activeSide = 'Dest'
         # Change the button labels accordingly
         self.btnCopy.SetLabel(_('<< Copy <<'))
         self.btnMove.SetLabel(_('<< Move <<'))
         self.btnDelete.SetLabel(_('Delete >>'))
         self.btnNewFolder.SetLabel(_('New Folder >>'))
         self.btnUpdateDB.SetLabel(_('Update DB >>'))
         # We need to know the active Label
         lbl = self.lblDest.GetLabel()
         # We need to know the active File Control
         ctrl = self.destFile
         # We need to know if the destination is local or the SRB.
         if lbl == self.LOCAL_LABEL:
            # If we are local, offer to connect to the SRB
            self.btnConnect.SetLabel(_('Connect >>'))
            # The source Path we need is the local path
            sourcePath = self.destDir.GetPath()
         else:
            # If we are connected to the SRB, offer to disconnect from the SRB
            self.btnConnect.SetLabel(_('Disconnect >>'))
            # The source Path we need is the SRB Path
            sourcePath = self.GetFullPath(self.srbDestDir)
         # Determine if any items are selected, which enables the
         # Update DB and Delete buttons
         if self.destFile.GetSelectedItemCount() > 0:
            updateFiles = True
         # What destination path to use depends on whether the other side is local or SRB
         if self.lblSource.GetLabel() == self.LOCAL_LABEL:
            destPath = self.sourceDir.GetPath()
         else:
            destPath = self.GetFullPath(self.srbSourceDir)

      # Enable the folder button if a sourcePath is defined above
      if sourcePath == '':
         self.btnNewFolder.Enable(False)
      else:
         self.btnNewFolder.Enable(True)
      # Enable the Copy and Move buttons if source and destination paths are defined above
      if (sourcePath != '') and \
         (destPath != ''):
         self.btnCopy.Enable(True)
         self.btnMove.Enable(True)
      else:
         self.btnCopy.Enable(False)
         self.btnMove.Enable(False)
      # Enable the Delete button if appropriate
      # Looking at the correct controls, determine if there are any files in the current folder
      if self.activeSide == 'Source':
         fileCount = self.sourceFile.GetItemCount()
         lbl0 = _('<< Delete')
         lblF = _('<< Delete Folder')
      else:
         fileCount = self.destFile.GetItemCount()
         lbl0 = _('Delete >>')
         lblF = _('Delete Folder >>')
      # If no source path is defined above, disable the Delete Button
      if (sourcePath == ''):
         self.btnDelete.Enable(False)
      # If there is a source path, set the appropriate label and enable the Delete Button
      else:
         if self.IsEmpty(lbl, sourcePath):
            self.btnDelete.SetLabel(lblF)
            self.btnDelete.Enable(True)
         else:
            self.btnDelete.SetLabel(lbl0)
            if updateFiles:
                self.btnDelete.Enable(True)
            else:
                self.btnDelete.Enable(False)
      # Enable the Update DB button if a sourcePath is defined above and we're LOCAL
      if (lbl == self.LOCAL_LABEL):
          # on the Mac, we don't correctly detect that a file name has been selected immediately.  Therefore, 
          # we need to use CallAfter so that the check will be made later, after the file selection is complete.
          wx.CallAfter(self.EnableUpdateDBButton, ctrl)
      else:
          self.btnUpdateDB.Enable(False)
      # Enable the Connect/Disconnect button
      self.btnConnect.Enable(True)
      
   def EnableUpdateDBButton(self, ctrl):
       """ Determine if the Update DB Button should be enabled """
       # NOTE:  This must be done after the OnSetFocus event is complete on the Mac or it won't work because the 
       #        GetSelectedItemCount() value is not set until too late.
       if ctrl.GetSelectedItemCount() > 0:
           self.btnUpdateDB.Enable(True)
       else:
           self.btnUpdateDB.Enable(False)
       

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

       # We need to know whether to look at the local file path or at the SRB
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
       else:
           if "__WXMSW__" in wx.PlatformInfo:
               srb = ctypes.cdll.srbClient
           else:
               srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

           if 'unicode' in wx.PlatformInfo:
               path = path.encode(TransanaGlobal.encoding)
           
           # If SRB, see if there's anything in the catalog
           srbCatalogCount = srb.srb_query_child_dataset(self.srbConnectionID.value, 0, path)
      
           # If any files are counted ...
           if srbCatalogCount != 0:
               # ... then we're not Empty.
               result = False
               
       return result

   def OnDirSelect(self, event):
      """ This method is called when a directory/folder is selected in the tree, either locally or on the SRB """
      # First, determine whether we're on the source or destination side and set local variables lbl, path, target,
      # and filter accordingly.
      if (event.GetEventObject() == self.sourceDir.GetTreeCtrl()) or \
         (event.GetId() == self.srbSourceDir.GetId()) or \
         (event.GetId() == ID_SOURCEFILTER):
         self.activeSide = 'Source' 
         lbl = self.lblSource.GetLabel()
         # We need a different path depending on whether we are on the local file system or on the SRB
         if lbl == self.LOCAL_LABEL:
            path = self.sourceDir.GetPath()
         else:
            path = self.GetFullPath(self.srbSourceDir)
         target = self.sourceFile
         filter = self.sourceFilter
      elif (event.GetEventObject() == self.destDir.GetTreeCtrl()) or \
           (event.GetId() == self.srbDestDir.GetId()) or \
           (event.GetId() == ID_DESTFILTER):
         self.activeSide = 'Dest' 
         lbl = self.lblDest.GetLabel()
         # We need a different path depending on whether we are on the local file system or on the SRB
         if lbl == self.LOCAL_LABEL:
            path = self.destDir.GetPath()
         else:
            path = self.GetFullPath(self.srbDestDir)
         target = self.destFile
         filter = self.destFilter
      if 'unicode' in wx.PlatformInfo:
          # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
          prompt = unicode(_('Directory changed to %s'), 'utf8')
      else:
          prompt = _('Directory changed to %s')
      self.SetStatusText(prompt % path)
      # Enable the New Folder button if the appropriate folder is selected for the first time   
      if path != '':
         self.btnNewFolder.Enable(True)
      # We need to trap the exception if the path is invalid (eg. pointing to a CD drive with no CD)
      try:
          # Refresh the File List to reflect any changes made to path or filter
          self.RefreshFileList(lbl, path, target, filter)
          # Enable the Delete button if appropriate  (If a Directory pane is selected, we need to check the Delete button separate
          # from the OnSetFocus method above, as OnSetFocus is called BEFORE the File List is reset, so may give incorrect results
          # regarding when the Delete button should be disabled.
          if  (((self.activeSide == 'Source') and (self.sourceFile.GetSelectedItemCount() > 0)) or \
              ((self.activeSide == 'Dest') and (self.destFile.GetSelectedItemCount() > 0))):
             self.btnDelete.Enable(True)
          else:
             if self.IsEmpty(lbl, path):
                if self.activeSide == 'Source':
                   self.btnDelete.SetLabel(_('<< Delete Folder'))
                else:
                   self.btnDelete.SetLabel(_('Delete Folder >>'))
                self.btnDelete.Enable(True)
             else:
                if self.activeSide == 'Source':
                   self.btnDelete.SetLabel(_('<< Delete'))
                else:
                   self.btnDelete.SetLabel(_('Delete >>'))
                self.btnDelete.Enable(False)
      except:
          # If the exception is raised, see if we're looking at a local drive.
          if lbl == self.LOCAL_LABEL:
              # If so, disable the buttons
              self.btnConnect.Enable(False)
              self.btnNewFolder.Enable(False)
              self.btnCopy.Enable(False)
              self.btnMove.Enable(False)
              self.btnDelete.Enable(False)
              self.btnUpdateDB.Enable(False)
              # We also need to remove the bad selection from the GenericDirCtrl's Tree to prevent
              # the exception from being re-raised repeatedly
              if self.activeSide == 'Source':
                  # self.sourceDir.SetPath('') doesn't do anything.
                  self.sourceDir.GetTreeCtrl().Unselect()
              else:
                  # self.destDir.SetPath('')  doesn't do anything
                  self.destDir.GetTreeCtrl().Unselect()
              # The File List has already been refreshed.  Since the path was bogus, it was refreshed with
              # data from the current working directory, which is totally NOT helpful.  Let's clear it.
              self.RefreshFileList(lbl, '', target, filter)

   def OnListItem(self, event):
        if event.GetId() == self.sourceFile.GetId():
            sourceFile = self.sourceFile
        elif event.GetId() == self.destFile.GetId():
            sourceFile = self.destFile
        else:
            return
        
        if sourceFile.GetSelectedItemCount() > 0:
            self.btnDelete.Enable(True)
        else:
            self.btnDelete.Enable(False)
        
       
   def RefreshFileList(self, lbl, path, target, filter):
      """ Refresh the File List.  This is needed whenever something happens where the file list might change,
          such as when a Folder is selected.  It is also used for the Refresh button, which allows the user to
          update the screen to reflect changes that may have occurred outside of Transana or the File Management
          Utility."""
      if "__WXMSW__" in wx.PlatformInfo:
          srb = ctypes.cdll.srbClient
      else:
          srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

      # Clear the File List
      target.ClearAll()

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
                 target.InsertStringItem(target.GetItemCount(), item)
         except:
             if True:
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
         #  srbDaiMcatQueryChildDataset runs a query on the SRB server and returns a number of records.  Not the total
         #  number of records, perhaps, but the number it can read into a data buffer.  According to Bing Zhu at
         #  SDSC, it is about 200 records at a time.
         #
         #  Those records can then be accessed by using srbDaiMcatGetData.
         #
         #  Once all the records in the buffer are read, srbDaiMcatGetMoreData is called, telling the SRB server
         #  to fill the buffer with the next set of records.  Only when this function returns a value of 0 have all
         #  the records been returned.
         #
         # Create an empty File List
         filelist = []

         if 'unicode' in wx.PlatformInfo:
             path = path.encode(TransanaGlobal.encoding)

         # Now we retrieve the list of files from the SRB.
         # Get the first batch of File Names from the SRB
         srbCatalogCount = srb.srb_query_child_dataset(self.srbConnectionID.value, 0, path)

         # Check for errors in getting the File Names
         if srbCatalogCount < 0:
            self.srbErrorMessage(srbCatalogCount)
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
                  # Add the result to the File List
                  filelist.append(tempStr)
               # Make another call to the SRB to see if there are more records to read into the SRB's Data Buffer
               srbCatalogCount = srb.srb_get_more_subcolls(self.srbConnectionID.value)

         # Get selection from dropdown list
         selection=filter.GetStringSelection()
         # Iterate through all files from the filelist and add to the ListCtrl target
         for item in self.itemsToShow(path, filelist, selection, isSrbList=True):
             target.InsertStringItem(target.GetItemCount(), item)

   def itemsToShow(self, path, filelist, selection, isSrbList):
       """Return a list containing filenames to be displayed depending upon 
       the file extensions found in selection"""
       itemList = [] # Initialize return value to []

       # DO NOT SORT THE FILE LIST if it is on the SRB!!!!  This breaks the connection with the SRB so that files are not
       # given the proper File Size when downloading them from the SRB, thus making Transfer look bad.
       if not isSrbList:
           filelist.sort() # Sort the file list
       # Create a REfilter which can match '.' followed by alphanumerics and _
       regexp = re.compile(r'\.\w*') 
       # Get a list containing all extensions from selection
       srch = regexp.findall(selection)
       if 'unicode' in wx.PlatformInfo:
           # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
           prompt = unicode(_('All files (*.*)'), 'utf8')
       else:
           prompt = _('All files (*.*)')
       # For each item in the file list
       for item in filelist:
           if isSrbList: 
               # files in the file list from the SRB have an invisible character [chr(0)] on the end that does not get stripped out.
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
      # Call the Function
      if "__WXMSW__" in wx.PlatformInfo:
          srb = ctypes.cdll.srbClient
      else:
          srb = ctypes.cdll.LoadLibrary("srbClient.dylib")

      srb.srb_err_msg(errorCode, errorMsg, 200)
      # Strip whitespace and null character (c string terminator) from buf
      errorMsg = string.strip(errorMsg)[:-1]
      
      if 'unicode' in wx.PlatformInfo:
          # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
          prompt = unicode(_('An error has occurred in working with the Storage Resource Broker.\nError Code: %s\nError Message: "%s"'), 'utf8')
      else:
          prompt = _('An error has occurred in working with the Storage Resource Broker.\nError Code: %s\nError Message: "%s"')
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
