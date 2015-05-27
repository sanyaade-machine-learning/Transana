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

"""This module implements the DatabaseTreeTab class for the Data Display
Objects."""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

import DBInterface
import Dialogs
import KWManager
import Misc
from TransanaExceptions import *
# import Transana's Constants
import TransanaConstants
# Import Transana's Global Variables
import TransanaGlobal
# Import Transana's Images
import TransanaImages

# import wxPython
import wx
import CoreData
import CoreDataPropertiesForm
import Library
import Episode
import os
import sys
import string

# Define the maximum number of video files allowed.  (This could change!)
if TransanaConstants.proVersion:
    MEDIAFILEMAX = 4
else:
    MEDIAFILEMAX = 1

class EpisodePropertiesForm(Dialogs.GenForm):
    """Form containing Episode fields."""

    def __init__(self, parent, id, title, ep_object):
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, (550,435), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Episode Properties')
        # Remember the Parent Window
        self.parent = parent
        self.obj = ep_object

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Episode ID
        self.id_edit = self.new_edit_box(_("Episode ID"), v1, self.obj.id, maxLen=100)
        # Add the element to the sizer
        r1Sizer.Add(v1, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Library ID layout
        series_edit = self.new_edit_box(_("Library ID"), v2, self.obj.series_id)
        # Add the element to the sizer
        r2Sizer.Add(v2, 2, wx.EXPAND)
        series_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Dialogs.GenForm does not provide a Masked text control, so the Date
        # Field is handled differently than other fields.
        
        # Create a VERTICAL sizer for the next element
        v3 = wx.BoxSizer(wx.VERTICAL)
        # Date [label]
        date_lbl = wx.StaticText(self.panel, -1, _("Date (MM/DD/YYYY)"))
        v3.Add(date_lbl, 0, wx.BOTTOM, 3)

        # Date
        # Use the Masked Text Control (Requires wxPython 2.4.2.4 or later)
        # TODO:  Make Date autoformat localizable
        self.dt_edit = wx.lib.masked.TextCtrl(self.panel, -1, '', autoformat='USDATEMMDDYYYY/')
        v3.Add(self.dt_edit, 0, wx.EXPAND)
        # If a Date is know, load it into the control
        if (self.obj.tape_date != None) and (self.obj.tape_date != '') and (self.obj.tape_date != '01/01/0'):
            self.dt_edit.SetValue(self.obj.tape_date)
        # Add the element to the sizer
        r2Sizer.Add(v3, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Length
        self.len_edit = self.new_edit_box(_("Length"), v4, self.obj.tape_length_str())
        # Add the element to the sizer
        r2Sizer.Add(v4, 1, wx.EXPAND)
        self.len_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Media Filename(s) [label]
        txt = wx.StaticText(self.panel, -1, _("Media Filename(s)"))
        mainSizer.Add(txt, 0, wx.BOTTOM, 3)

        # Create a HORIZONTAL sizer for the next row
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Media Filename(s)
        # If the media filename path is not empty, we should normalize the path specification
        if self.obj.media_filename == '':
            filePath = self.obj.media_filename
        else:
            filePath = os.path.normpath(self.obj.media_filename)
        # Initialize the list of media filenames with the first one.
        self.filenames = [filePath]
        # For each additional Media file ...
        for vid in self.obj.additional_media_files:
            # ... add it to the filename list
            self.filenames.append(vid['filename'])
        self.fname_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.Size(180, 60), self.filenames)
        # Add the element to the sizer
        r3Sizer.Add(self.fname_lb, 5, wx.EXPAND)
        self.fname_lb.SetDropTarget(ListBoxFileDropTarget(self.fname_lb))
        
        # Add a horizontal spacer to the row sizer        
        r3Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Add File button layout
        addFile = wx.Button(self.panel, wx.ID_FILE1, _("Add File"), wx.DefaultPosition)
        v4.Add(addFile, 0, wx.EXPAND | wx.BOTTOM, 3)
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)

        # Remove File button layout
        removeFile = wx.Button(self.panel, -1, _("Remove File"), wx.DefaultPosition)
        v4.Add(removeFile, 0, wx.EXPAND | wx.BOTTOM, 3)
        wx.EVT_BUTTON(self, removeFile.GetId(), self.OnRemoveFile)

        if TransanaConstants.proVersion:
            # SynchronizeFiles button layout
            synchronize = wx.Button(self.panel, -1, _("Synchronize"), wx.DefaultPosition)
            v4.Add(synchronize, 0, wx.EXPAND)
            synchronize.Bind(wx.EVT_BUTTON, self.OnSynchronize)

        # Add the element to the sizer
        r3Sizer.Add(v4, 1, wx.EXPAND)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            r3Sizer.Add((2, 0))

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r3Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r4Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v5 = wx.BoxSizer(wx.VERTICAL)
        # Comment
        comment_edit = self.new_edit_box(_("Comment"), v5, self.obj.comment, maxLen=255)
        # Add the element to the sizer
        r4Sizer.Add(v5, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r4Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r5Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v6 = wx.BoxSizer(wx.VERTICAL)
        # Keyword Group [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword Group"))
        v6.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword Group [list box]

        kw_groups_id = wx.NewId()
        # Create an empty Keyword Group List for now.  We'll populate it later (for layout reasons)
        self.kw_groups = []
        self.kw_group_lb = wx.ListBox(self.panel, kw_groups_id, wx.DefaultPosition, wx.DefaultSize, self.kw_groups)
        v6.Add(self.kw_group_lb, 1, wx.EXPAND)

        # Add the element to the sizer
        r5Sizer.Add(v6, 1, wx.EXPAND)

        self.kw_list = []
        wx.EVT_LISTBOX(self, kw_groups_id, self.OnGroupSelect)

        # Add a horizontal spacer
        r5Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v7 = wx.BoxSizer(wx.VERTICAL)
        # Keyword [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword"))
        v7.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword [list box]
        self.kw_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.DefaultSize, self.kw_list, style=wx.LB_EXTENDED)
        v7.Add(self.kw_lb, 1, wx.EXPAND)

        wx.EVT_LISTBOX_DCLICK(self, self.kw_lb.GetId(), self.OnAddKW)

        # Add the element to the sizer
        r5Sizer.Add(v7, 1, wx.EXPAND)

        # Add a horizontal spacer
        r5Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v8 = wx.BoxSizer(wx.VERTICAL)
        # Keyword transfer buttons
        add_kw = wx.Button(self.panel, wx.ID_FILE2, ">>", wx.DefaultPosition)
        v8.Add(add_kw, 0, wx.EXPAND | wx.TOP, 20)
        wx.EVT_BUTTON(self.panel, wx.ID_FILE2, self.OnAddKW)

        rm_kw = wx.Button(self.panel, wx.ID_FILE3, "<<", wx.DefaultPosition)
        v8.Add(rm_kw, 0, wx.EXPAND | wx.TOP, 10)
        wx.EVT_BUTTON(self, wx.ID_FILE3, self.OnRemoveKW)

        kwm = wx.BitmapButton(self.panel, wx.ID_FILE4, TransanaImages.KWManage.GetBitmap())
        v8.Add(kwm, 0, wx.EXPAND | wx.TOP, 10)
        # Add a spacer to increase the height of the Keywords section
        v8.Add((0, 60))
        kwm.SetToolTipString(_("Keyword Management"))
        wx.EVT_BUTTON(self, wx.ID_FILE4, self.OnKWManage)

        # Add the element to the sizer
        r5Sizer.Add(v8, 0)

        # Add a horizontal spacer
        r5Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v9 = wx.BoxSizer(wx.VERTICAL)

        # Episode Keywords [label]
        txt = wx.StaticText(self.panel, -1, _("Episode Keywords"))
        v9.Add(txt, 0, wx.BOTTOM, 3)

        # Episode Keywords [list box]
        
        # Create an empty ListBox
        self.ekw_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.DefaultSize, style=wx.LB_EXTENDED)
        v9.Add(self.ekw_lb, 1, wx.EXPAND)
        
        self.ekw_lb.Bind(wx.EVT_KEY_DOWN, self.OnKeywordKeyDown)

        # Add the element to the sizer
        r5Sizer.Add(v9, 2, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r5Sizer, 5, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)

        # Core Data button layout
        CoreData = wx.Button(self.panel, -1, _("Core Data"))
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            btnSizer.Add((2, 0))
        btnSizer.Add(CoreData, 0)
        wx.EVT_BUTTON(self, CoreData.GetId(), self.OnCoreDataClick)

        # Add the buttons
        self.create_buttons(sizer=btnSizer)
        # Add the button sizer to the main sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            mainSizer.Add((0, 2))

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
        self.SetSize(wx.Size(max(550, width), max(435, height)))
        # Define the minimum size for this dialog as the current size
        self.SetSizeHints(max(550, width), max(435, height))
        # Center the form on screen
        TransanaGlobal.CenterOnPrimary(self)

        # We need to set some minimum sizes so the sizers will work right
        self.kw_group_lb.SetSizeHints(minW = 50, minH = 20)
        self.kw_lb.SetSizeHints(minW = 50, minH = 20)
        self.ekw_lb.SetSizeHints(minW = 50, minH = 20)

        # We populate the Keyword Groups, Keywords, and Clip Keywords lists AFTER we determine the Form Size.
        # Long Keywords in the list were making the form too big!

        self.kw_groups = DBInterface.list_of_keyword_groups()
        for keywordGroup in self.kw_groups:
            self.kw_group_lb.Append(keywordGroup)

        # Load the parent Library in order to determine the default Keyword Group
        tempLibrary = Library.Library(self.obj.series_id)
        # Select the Library Default Keyword Group in the Keyword Group list
        if (tempLibrary.keyword_group != '') and (self.kw_group_lb.FindString(tempLibrary.keyword_group) != wx.NOT_FOUND):
            self.kw_group_lb.SetStringSelection(tempLibrary.keyword_group)
        # If no Default Keyword Group is defined, select the first item in the list
        else:
            # but only if there IS a first item in the list.
            if len(self.kw_groups) > 0:
                self.kw_group_lb.SetSelection(0)
        if self.kw_group_lb.GetSelection() != wx.NOT_FOUND:
            self.kw_list = \
                DBInterface.list_of_keywords_by_group(self.kw_group_lb.GetStringSelection())
        else:
            self.kw_list = []
        for keyword in self.kw_list:
            self.kw_lb.Append(keyword)

        # Populate the ListBox
        for episodeKeyword in self.obj.keyword_list:
            self.ekw_lb.Append(episodeKeyword.keywordPair)

        self.id_edit.SetFocus()


    def refresh_keyword_groups(self):
        """Refresh the keyword groups listbox."""
        self.kw_groups = DBInterface.list_of_keyword_groups()
        self.kw_group_lb.Clear()
        self.kw_group_lb.InsertItems(self.kw_groups, 0)
        if len(self.kw_groups) > 0:
            self.kw_group_lb.SetSelection(0)

    def refresh_keywords(self):
        """Refresh the keywords listbox."""
        sel = self.kw_group_lb.GetStringSelection()
        if sel:
            self.kw_list = DBInterface.list_of_keywords_by_group(sel)
            self.kw_lb.Clear()
            if len(self.kw_list) > 0:
                self.kw_lb.InsertItems(self.kw_list, 0)
                self.kw_lb.EnsureVisible(0)

    def highlight_bad_keyword(self):
        """ Highlight the first bad keyword in the keyword list """
        # Get the Keyword Group name
        sel = self.kw_group_lb.GetStringSelection()
        # If there was a selected Keyword Group ...
        if sel:
            # ... initialize a list of keywords
            kwlist = []
            # Iterate through the current keyword group's keywords ...
            for item in range(self.kw_lb.GetCount()):
                # ... and add them to the list of keywords 
                kwlist.append("%s : %s" % (sel, self.kw_lb.GetString(item)))
            # Now iterate through the list of Episode Keywords
            for item in range(self.ekw_lb.GetCount()):
                # If the keyword is from the current Keyword Group AND the keyword is not in the keyword list ...
                if (self.ekw_lb.GetString(item)[:len(sel)] == sel) and (not self.ekw_lb.GetString(item) in kwlist):
                    # ... select the current item in the Episode Keywords control ...
                    self.ekw_lb.SetSelection(item)
                    # ... and stop looking for bad keywords.  (We just highlight the first!)
                    break

    def OnMediaFilenameEdit(self, event):
        """ Invoked when Media Filename is changed. """
        self.len_edit.SetValue('00:00:00')
        self.obj.tape_length = 0

    def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        # As long as we have fewer files than the max allowed ...
        if (self.fname_lb.GetCount() < MEDIAFILEMAX) or (self.fname_lb.GetString(0) == ''):
            # ... reset the object's MAIN filename to the top item in the list (if there is one)
            if self.fname_lb.GetCount() > 0:
                self.obj.media_filename = self.fname_lb.GetString(0)
            # Get the directory for that MAIN file name
            dirName = os.path.dirname(self.obj.media_filename)
            # If no path can be extracted from the file name, start with the Video Root Folder
            if dirName == '':
                dirName = TransanaGlobal.configData.videoPath
            # Get the File Name portion
            fileName = os.path.basename(self.obj.media_filename)
            # Determine the File's extension
            (fn, ext) = os.path.splitext(self.obj.media_filename)
            # If we have a known File Type or if blank, use "All Media Files".
            # If it's an unrecognized type, go to "All Files"
            if (TransanaGlobal.configData.LayoutDirection == wx.Layout_LeftToRight) and \
               (ext.lower() in ['.mpg', '.avi', '.mov', '.mp4', '.m4v', '.wmv', '.mp3', '.wav', '.wma', '.aac', '']):
                fileType =  '*.mpg;*.avi;*.mov;*.mp4;*.m4v;*.wmv;*.mp3;*.wav;*.wma;*.aac'
            elif (TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft) and \
               (ext.lower() in ['.mpg', '.avi', '.wmv', '.mp3', '.wav', '.wma', '.aac', '']):
                fileType =  '*.mpg;*.avi;*.wmv;*.mp3;*.wav;*.wma;*.aac'
            else:
                fileType = ''
            # Invoke the File Selector with the proper default directory, filename, file type, and style
            fs = wx.FileSelector(_("Select a media file"),
                            dirName,
                            fileName,
                            fileType, 
                            _(TransanaConstants.fileTypesString), 
                            wx.OPEN | wx.FILE_MUST_EXIST)
            # If user didn't cancel ..
            if fs != "":
                # Mac Filenames use a different encoding system.  We need to adjust the string returned by the FileSelector.
                # Surely there's an easier way, but I can't figure it out.
                if 'wxMac' in wx.PlatformInfo:
                    fs = Misc.convertMacFilename(fs)
                # If we have a non-blank line first line in the filenames list in the form ...
                if (self.fname_lb.GetCount() > 0) and (self.fname_lb.GetString(0).strip() != ''):
                    # ... add the file name to the end of the control
                    self.fname_lb.Append(fs)
                    # ... and add the file name to the Episode object as an ADDITIONAL media file
                    self.obj.additional_media_files = {'filename' : fs,
                                                       'length'   : 0,
                                                       'offset'   : 0,
                                                       'audio'    : 1 }
                # If we don't have a first media file defined ...
                else:
                    # ... set the filename as the MAIN media file for the Episode object
                    self.obj.media_filename = fs
                    # ... Clear the filename listbox
                    self.fname_lb.Clear()
                    # ... add the filename as the first line
                    self.fname_lb.Append(fs)
                    # If the form's ID field is empty ...
                    if self.id_edit.GetValue() == '':
                        # ... get the base filename
                        tempFilename = os.path.basename(fs)
                        # ... separate the filename root and extension
                        (self.obj.id, tempExt) = os.path.splitext(tempFilename)
                        # ... and set the ID to match the base file name
                        self.id_edit.SetValue(self.obj.id)
                self.OnMediaFilenameEdit(evt)
        # If we have the maximum number of media files already selected ...
        else:
            if MEDIAFILEMAX == 1:
                # ... Display an error message to the user.
                msg = _('Only one media file is allowed at a time in this version of Transana.')
            else:
                # ... Display an error message to the user.
                msg = _('A maximum of %d media files is allowed.') % MEDIAFILEMAX
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                 msg = unicode(msg, 'utf8')
            dlg = Dialogs.ErrorDialog(self, msg)
            dlg.ShowModal()
            dlg.Destroy()

    def OnRemoveFile(self, event):
        """ Remove File button event handler """
        # See if clips have been created from this episode
        clipCount = len(DBInterface.list_of_clips_by_episode(self.obj.number))
        # If clips have been created ...
        if clipCount > 0:
            # Build and display a warning message
            prompt = unicode(_('Removing media files from an Episode after clips have been created can\ncause problems with the clips.  Are you SURE you want to do this?'), 'utf8')
            msgDialog = Dialogs.QuestionDialog(self, prompt, noDefault=True)
            # Get the response from the user
            result = msgDialog.LocalShowModal()
            msgDialog.Destroy()
        # If no clips have been created ...
        else:
            # ... we just assume the user would say Yes
            result = wx.ID_YES
        # If the user says Yes ...
        if result == wx.ID_YES:
            # Determine the index of the selection in the file name list
            indx = self.fname_lb.GetSelection()
            # If there IS a selection in the file name list ...
            if indx > -1:
                # ... get rid of it!
                self.fname_lb.Delete(indx)
                # If this is an ADDITIONAL file ...
                if indx > 0:
                    # ... remove it from the underlying object
                    self.obj.remove_an_additional_vid(indx-1)
                # If it's not an ADDITIONAL file ...
                else:
                    # ... see if there ARE additional files.
                    if len(self.obj.additional_media_files) > 0:
                        # If so, move the top additional file's name to the main file name
                        self.obj.media_filename = self.obj.additional_media_files[0]['filename']
                        # Get the offset for the first additional file
                        tmpOffset = self.obj.additional_media_files[0]['offset']
                        # Delete the first additional file from the additional file list, as it's now the main file
                        self.obj.remove_an_additional_vid(0)
                        # For each remaining additional file ...
                        for vid in self.obj.additional_media_files:
                            # ... adjust the offset by the original offset of what is now the main file.
                            vid['offset'] -= tmpOffset
                    # If there are NO additional files ...
                    else:
                        # ... just clear the main file name value
                        self.obj.media_filename = ""
            self.OnMediaFilenameEdit(event)

    def OnSynchronize(self, event):
        """ Synchronize Media Files """
        # Determine the index of the selection in the file name list
        indx = self.fname_lb.GetSelection()
        # If there is no selection and there are exactly two files, synchronize them!
        if (indx == -1) and (self.fname_lb.GetCount() == 2):
            # We accomplish this by faking the selection of the second item in the list
            indx = 1
        # If there IS a selection in the file name list, and it's not the first item ...
        if indx > 0:
            # import Transana's Synchronize dialog  (Importing this earlier causes problems.)
            import Synchronize
            # Create a Synchronize Dialog
            synchDlg = Synchronize.Synchronize(self, self.fname_lb.GetString(0), self.fname_lb.GetString(indx), offset=self.obj.additional_media_files[indx - 1]['offset'])
            # See of the form was successfully built.
            if synchDlg.windowBuilt:
                # Show the form and get the results
                (offsetVal, lengthVal) = synchDlg.GetData()
                # Get the data from the Synchronize Dialog and assign it to the second media file's "offset" property
                self.obj.additional_media_files[indx - 1]['offset'] = offsetVal
                self.obj.additional_media_files[indx - 1]['length'] = lengthVal
                # Destroy the Synchronize Dialog
                synchDlg.Destroy()
        # If there is no selection ...
        else:
            # Create and display an error message
            prompt = _("Please select a file to synchronize with %s.")
            if 'unicode' in wx.PlatformInfo:
                prompt = unicode(prompt, 'utf8')
            tmpDlg = Dialogs.InfoDialog(self, prompt % self.fname_lb.GetString(0))
            tmpDlg.ShowModal()
            tmpDlg.Destroy()

    def OnCoreDataClick(self, event):
        """ Method for when Core Data button is clicked """
        # If no items are currently selected ...
        if (self.fname_lb.GetCount() > 0) and (len(self.fname_lb.GetSelections()) == 0):
            # ... then select the first item
            self.fname_lb.Select(0)
        
        # We can only edit Core Data once the Media Filename has been selected in the file list and the file must exist
        if (self.fname_lb.GetStringSelection() != '') and os.path.exists(self.fname_lb.GetStringSelection()):
            # Extract the path-less File Name from the full-path Media File Name for the selected item in the filename list
            (path, filename) = os.path.split(self.fname_lb.GetStringSelection())

            # Determine if the appropriate record exists in the database
            try:
                # Load the record if it exists, raises an exception if it does not
                coreData = CoreData.CoreData(filename)
                # If the record exists, lock it.
                coreData.lock_record()
            # If the record does not exist, a RecordNotFoundError exception is raised.
            except RecordNotFoundError:
                # Create an empty CoreData Object
                coreData = CoreData.CoreData()
                # Specify the path-less filename as the Identifier
                coreData.id = filename
            except RecordLockedError, e:
                ReportRecordLockedException(_('Core Data record'), coreData.id, e)
                # If we get a record lock error, we don't need to display the Core Data Properties form
                return
            # If a different exception is raised, report it and pass it on.
            except:
                (type, value, traceback) = sys.exc_info()
                # print "EpisodePropertiesForm.OnCoreDataClick:  Exception raised.\nType = %s\nValue = %s\nTraceback = %s" % (type, value, traceback)
                raise

            # Load the full-path filename of the selected item in the filenames list into the Core Data Object's Comment Field
            coreData.comment = self.fname_lb.GetStringSelection()
            # Create the Core Data Properites Form
            dlg = CoreDataPropertiesForm.EditCoreDataDialog(self, -1, coreData)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # While the "continue" flag is True ...
            while contin:
                try:
                    # Show the form and process the user input
                    if dlg.get_input() != None:
                        # Save the changes if the user presses OK
                        coreData.db_save()
                        # If the save goes through, we don't need to continue.
                        contin = False
                    # If the user presses "Cancel" ...
                    else:
                        # ... then we don't need to continue
                        contin = False
                except:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
            # Unlock the Record
            coreData.unlock_record()
        else:
            if (self.fname_lb.GetStringSelection() == ''):
                errordlg = Dialogs.ErrorDialog(self, _("You must specify a Media Filename before defining Core Data"))
                errordlg.ShowModal()
                errordlg.Destroy()
            else:
                errordlg = Dialogs.ErrorDialog(self, _("The Media Filename must point to a file that exists before defining Core Data"))
                errordlg.ShowModal()
                errordlg.Destroy()

    def OnAddKW(self, evt):
        """Invoked when the user activates the Add Keyword (>>) button."""
        # For each selected Keyword ...
        for item in self.kw_lb.GetSelections():
            # ... get the keyword group name ...
            kwg_name = self.kw_group_lb.GetStringSelection()
            # ... get the keyword name ...
            kw_name = self.kw_lb.GetString(item)
            # ... build the kwg : kw combination ...
            ep_kw = "%s : %s" % (kwg_name, kw_name)
            # ... and if it's NOT already in the Episode Keywords list ...
            if self.ekw_lb.FindString(ep_kw) == -1:
                # ... add the keyword to the Episode object ...
                self.obj.add_keyword(kwg_name, kw_name)
                # ... and add it to the Episode Keywords list box
                self.ekw_lb.Append(ep_kw)
        
    def OnRemoveKW(self, evt):
        """Invoked when the user activates the Remove Keyword (<<) button."""
        # Get the selection(s) from the Episode Keywords list box
        kwitems = self.ekw_lb.GetSelections()
        # The items are returned as an immutable tuple.  Convert this to a list.
        kwitems = list(kwitems)
        # Now sort the list.  For reasons that elude me, the list is arbitrarily ordered on the Mac, which causes
        # deletes to be done out of order so the wrong elements get deleted, which is BAD.
        kwitems.sort()
        # We have to go through the list items BACKWARDS so that item numbers don't change on us as we delete items!
        for item in range(len(kwitems), 0, -1):
            # Get the STRING of the keyword to delete
            sel = self.ekw_lb.GetString(kwitems[item - 1])
            # Separate out the Keyword Group and the Keyword
            kwlist = sel.split(':')
            kwg = kwlist[0].strip()
            # If the keyword contained a colon, we need to re-construct it!
            kw = ':'.join(kwlist[1:]).strip()
            # Try to delete the keyword
            delResult = self.obj.remove_keyword(kwg, kw)
            # If successful ...
            if delResult:
                # ... remove the item from the Episode Keywords list box.
                self.ekw_lb.Delete(kwitems[item - 1])
 
    def OnKWManage(self, evt):
        """Invoked when the user activates the Keyword Management button."""
        # find out if there is a default keyword group
        if self.kw_group_lb.IsEmpty():
            sel = None
        else:
            sel = self.kw_group_lb.GetStringSelection()
        # Create and display the Keyword Management Dialog
        kwm = KWManager.KWManager(self, sel, deleteEnabled=False)
        # Refresh the Keyword Groups list, in case it was changed.
        self.refresh_keyword_groups()
        # Make sure the last Keyword Group selected in the Keyword Management is selected when it gets closed.
        selPos = self.kw_group_lb.FindString(kwm.kw_group.GetStringSelection())
        if selPos == -1:
            selPos = 0
        if not self.kw_group_lb.IsEmpty():
            self.kw_group_lb.SetSelection(selPos)
        # Refresh the Keyword List, in case it was changed.
        self.refresh_keywords()
        # We must refresh the Keyword List in the DBTree to reflect changes made in the
        # Keyword Management.
        self.parent.tree.refresh_kwgroups_node()

    def OnGroupSelect(self, evt):
        """Invoked when the user selects a keyword group in the listbox."""
        self.refresh_keywords()
        
    def OnKeywordKeyDown(self, event):
        try:
            c = event.GetKeyCode()
            if c == wx.WXK_DELETE:
                if self.ekw_lb.GetSelection() != wx.NOT_FOUND:
                    self.OnRemoveKW(event)
        except:
            pass  # ignore non-ASCII keys

    def get_input(self):
        """Custom input routine."""
        # Inherit base input routine
        gen_input = Dialogs.GenForm.get_input(self)
        # If the user presses OK ...
        if gen_input:
            # ... get the ID
            self.obj.id = gen_input[_('Episode ID')]
            # Get all the file names from the form
            filenames = self.fname_lb.GetStrings()
            # If there are any file names in the form control ...
            if len(filenames) > 0:
                # ... use the first one for the MAIN media_filename
                self.obj.media_filename = filenames[0]
            # Additional Media Files are already in the self.obj object, so we don't need to worry about them here.
            # Get the taping data
            self.obj.tape_date = self.dt_edit.GetValue()
            # Get the comment
            self.obj.comment = gen_input[_('Comment')]
            # Keyword list is already updated via the OnAddKW() callback
        # If the user presses Cancel ...
        else:
            # ... clear the data object
            self.obj = None
        return self.obj
        

# This simple derrived class let's the user drop files onto a list box
class ListBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, listbox):
        wx.FileDropTarget.__init__(self)
        self.listbox = listbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        # If there are no files in the ListBox ...
        if (self.listbox.GetCount() == 0) or (self.listbox.GetString(0).strip() == ''):
            # ... clear it to prevent there being a blank line at the top of the list.
            self.listbox.Clear()
        # If we have not exceeded the maximum number of files allowed ...
        if self.listbox.GetCount() < MEDIAFILEMAX:
            # ... add the file name to the list box
            self.listbox.Append(files[0])
        # If we have the maximum number of media files already selected ...
        else:
            # ... Display an error message to the user.
             msg = _('A maximum of %d media files is allowed.')
             if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                 msg = unicode(msg, 'utf8')
             dlg = Dialogs.ErrorDialog(None, msg % MEDIAFILEMAX)
             dlg.ShowModal()
             dlg.Destroy()


class AddEpisodeDialog(EpisodePropertiesForm):
    """Dialog used when adding a new Episode."""

    def __init__(self, parent, id, library):
        obj = Episode.Episode()
        obj.owner = DBInterface.get_username()
        obj.series_num = library.number
        obj.series_id = library.id
        EpisodePropertiesForm.__init__(self, parent, id, _("Add Episode"), obj)

class EditEpisodeDialog(EpisodePropertiesForm):
    """Dialog used when editing Episode properties."""

    def __init__(self, parent, id, ep_object):
        EpisodePropertiesForm.__init__(self, parent, id, _("Episode Properties"), ep_object)
