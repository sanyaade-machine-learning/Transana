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

# import wxPython
import wx
import CoreData
import CoreDataPropertiesForm
import Series
import Episode
import os
import sys
import string

class EpisodePropertiesForm(Dialogs.GenForm):
    """Form containing Episode fields."""

    def __init__(self, parent, id, title, ep_object):
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, (550,385), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Episode Properties')
        # Define the minimum size for this dialog as the initial size
        self.SetSizeHints(550, 385)
        # Remember the Parent Window
        self.parent = parent
        self.obj = ep_object

        ######################################################
        # Tedious GUI layout code follows
        ######################################################

        # Episode ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 35)  # 35% width
        lay.height.AsIs()
        self.id_edit = self.new_edit_box(_("Episode ID"), lay, self.obj.id, maxLen=100)

        # Series ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.RightOf(self.id_edit, 6)            # 6 right of Episode ID
        lay.width.PercentOf(self.panel, wx.Width, 30)  # 35% width
        lay.height.AsIs()
        series_edit = self.new_edit_box(_("Series ID"), lay, self.obj.series_id)
        series_edit.Enable(False)

        # Dialogs.GenForm does not provide a Masked text control, so the Date
        # Field is handled differently than other fields.
        
        # Date layout [label]
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.RightOf(series_edit, 6)        # 6 right of Series ID
        lay.width.PercentOf(self.panel, wx.Width, 15)  # 15% width
        lay.height.AsIs()
        date_lbl = wx.StaticText(self.panel, -1, _("Date (MM/DD/YYYY)"))
        date_lbl.SetConstraints(lay)

        # Date layout
        lay = wx.LayoutConstraints()
        lay.top.Below(date_lbl, 3)         # 
        lay.left.SameAs(date_lbl, 0)        # 6 right of Series ID
        lay.width.PercentOf(self.panel, wx.Width, 15)  # 15% width
        lay.height.AsIs()
        # Use the Masked Text Control (Requires wxPython 2.4.2.4 or later)
        # TODO:  Make Date autoformat localizable
        self.dt_edit = wx.lib.masked.TextCtrl(self.panel, -1, '', autoformat='USDATEMMDDYYYY/')
        # If a Date is know, load it into the control
        if (self.obj.tape_date != None) and (self.obj.tape_date != '') and (self.obj.tape_date != '01/01/0'):
            self.dt_edit.SetValue(self.obj.tape_date)
        self.dt_edit.SetConstraints(lay)

        # Length layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.RightOf(self.dt_edit, 6)            # 6 right of Date Taped
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        self.len_edit = self.new_edit_box(_("Length"), lay, self.obj.tape_length_str())
        self.len_edit.Enable(False)

        # Media Filename Layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.id_edit, 10)              # 10 under ID
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 80)  # 80% width
        lay.height.AsIs()
        # If the media filename path is not empty, we should normalize the path specification
        if self.obj.media_filename == '':
            filePath = self.obj.media_filename
        else:
            filePath = os.path.normpath(self.obj.media_filename)
        self.fname_edit = self.new_edit_box(_("Media Filename"), lay, filePath, maxLen=255)
        self.fname_edit.SetDropTarget(EditBoxFileDropTarget(self.fname_edit))
        wx.EVT_TEXT(self, self.fname_edit.GetId(), self.OnMediaFilenameEdit)
        
        # Browse button layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.fname_edit, wx.Top)
        lay.left.RightOf(self.fname_edit, 10)
        lay.right.SameAs(self.panel, wx.Right, 10)
        lay.bottom.SameAs(self.fname_edit, wx.Bottom)
        browse = wx.Button(self.panel, wx.ID_FILE1, _("Browse"), wx.DefaultPosition)
        browse.SetConstraints(lay)
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)

        # Comment layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.fname_edit, 10)      # 10 under media filename
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        comment_edit = self.new_edit_box(_("Comment"), lay, self.obj.comment, maxLen=255)

        # Keyword Group layout [label]
        lay = wx.LayoutConstraints()
        lay.top.Below(comment_edit, 10)         # 10 under comment
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 22)  # 22% width
        lay.height.AsIs()
        txt = wx.StaticText(self.panel, -1, _("Keyword Group"))
        txt.SetConstraints(lay)

        # Keyword Group layout [list box]
        lay = wx.LayoutConstraints()
        lay.top.Below(txt, 3)                   # 3 under label
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.SameAs(txt, wx.Width)          # width same as label
        lay.bottom.SameAs(self.panel, wx.Height, 50)   # 50 from bottom

        # Load the parent Series in order to determine the default Keyword Group
        tempSeries = Series.Series(self.obj.series_id)

        kw_groups_id = wx.NewId()
        self.kw_groups = DBInterface.list_of_keyword_groups()
        self.kw_group_lb = wx.ListBox(self.panel, kw_groups_id, wx.DefaultPosition, wx.DefaultSize, self.kw_groups)
        self.kw_group_lb.SetConstraints(lay)
        # Select the Series Default Keyword Group in the Keyword Group list
        if (tempSeries.keyword_group != '') and (self.kw_group_lb.FindString(tempSeries.keyword_group) != wx.NOT_FOUND):
            self.kw_group_lb.SetStringSelection(tempSeries.keyword_group)
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
        wx.EVT_LISTBOX(self, kw_groups_id, self.OnGroupSelect)

        # Keyword layout [label]
        lay = wx.LayoutConstraints()
        lay.top.Below(comment_edit, 10)         # 10 under comment
        lay.left.RightOf(txt, 10)               # 10 right of KW Group
        lay.width.PercentOf(self.panel, wx.Width, 22)  # 22% width
        lay.height.AsIs()
        txt = wx.StaticText(self.panel, -1, _("Keyword"))
        txt.SetConstraints(lay)

        # Keyword layout [list box]
        lay = wx.LayoutConstraints()
        lay.top.Below(txt, 3)                   # 3 under label
        lay.left.SameAs(txt, wx.Left)            # left same as label
        lay.width.SameAs(txt, wx.Width)          # width same as label
        lay.bottom.SameAs(self.panel, wx.Height, 50)   # 50 from bottom
        
        self.kw_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.DefaultSize, self.kw_list)
        self.kw_lb.SetConstraints(lay)

        wx.EVT_LISTBOX_DCLICK(self, self.kw_lb.GetId(), self.OnAddKW)

        # Keyword transfer buttons
        lay = wx.LayoutConstraints()
        lay.top.Below(txt, 30)                  # 30 under label
        lay.left.RightOf(txt, 10)               # 10 right of label
        lay.width.PercentOf(self.panel, wx.Width, 6)   # 6% width
        lay.height.AsIs()
        add_kw = wx.Button(self.panel, wx.ID_FILE2, ">>", wx.DefaultPosition)
        add_kw.SetConstraints(lay)
        wx.EVT_BUTTON(self.panel, wx.ID_FILE2, self.OnAddKW)

        lay = wx.LayoutConstraints()
        lay.top.Below(add_kw, 10)
        lay.left.SameAs(add_kw, wx.Left)
        lay.width.SameAs(add_kw, wx.Width)
        lay.height.AsIs()
        rm_kw = wx.Button(self.panel, wx.ID_FILE3, "<<", wx.DefaultPosition)
        rm_kw.SetConstraints(lay)
        wx.EVT_BUTTON(self, wx.ID_FILE3, self.OnRemoveKW)

        lay = wx.LayoutConstraints()
        lay.top.Below(rm_kw, 10)
        lay.left.SameAs(rm_kw, wx.Left)
        lay.width.SameAs(rm_kw, wx.Width)
        lay.height.AsIs()
        bitmap = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "KWManage.xpm"), wx.BITMAP_TYPE_XPM)
        kwm = wx.BitmapButton(self.panel, wx.ID_FILE4, bitmap)
        kwm.SetConstraints(lay)
        kwm.SetToolTipString(_("Keyword Management"))
        wx.EVT_BUTTON(self, wx.ID_FILE4, self.OnKWManage)


        # Episode Keywords [label]
        lay = wx.LayoutConstraints()
        lay.top.Below(comment_edit, 10)         # 10 under comment
        lay.left.SameAs(add_kw, wx.Right, 10)   # 10 from Add keyword button
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        txt = wx.StaticText(self.panel, -1, _("Episode Keywords"))
        txt.SetConstraints(lay)

        # Episode Keywords [list box]
        lay = wx.LayoutConstraints()
        lay.top.Below(txt, 3)                   # 3 under label
        lay.left.SameAs(txt, wx.Left)            # left same as label
        lay.width.SameAs(txt, wx.Width)          # width same as label
        lay.bottom.SameAs(self.panel, wx.Height, 50)   # 50 from bottom
        
        # Create an empty ListBox
        self.ekw_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.DefaultSize)
        # Populate the ListBox
        for episodeKeyword in self.obj.keyword_list:
            self.ekw_lb.Append(episodeKeyword.keywordPair)

        self.ekw_lb.SetConstraints(lay)
        
        self.ekw_lb.Bind(wx.EVT_KEY_DOWN, self.OnKeywordKeyDown)

        # Core Data button layout
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self.panel, wx.Bottom, 10)
        lay.left.SameAs(self.id_edit, wx.Left, 0)
        lay.height.AsIs()
        lay.width.AsIs()
        CoreData = wx.Button(self.panel, -1, _("Core Data")) 
        CoreData.SetConstraints(lay)
        wx.EVT_BUTTON(self, CoreData.GetId(), self.OnCoreDataClick)

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

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
            self.kw_list = \
                DBInterface.list_of_keywords_by_group(sel)
            self.kw_lb.Clear()
            self.kw_lb.InsertItems(self.kw_list, 0)

    def OnMediaFilenameEdit(self, event):
        """ Invoked when Media Filename is changed. """
        self.len_edit.SetValue('00:00:00')
        self.obj.tape_length = 0

        # If we are using the ANSI version of wxPython OR
        # (if we're on the Mac AND are in the Single-User version of Transana)
        # then we need to block Unicode characters from media filenames.
        # Unicode characters still cause problems on the Mac for the Multi-User version of Transana,
        # but can be made to work if shared waveforming is done on a Windows computer.
        if ('ansi' in wx.PlatformInfo) or (('wxMac' in wx.PlatformInfo) and TransanaConstants.singleUserVersion):
            # Transana can't cope with non-English characters in the path or file name for media files.  (The ActiveX
            # control chokes on a string conversion on Windos and Python os.path.exists() can't cope on the Mac)
            # Create a string of legal characters for the file names
            allowedChars = TransanaConstants.legalFilenameCharacters
            # check each character in the file name string
            for char in self.fname_edit.GetValue():
                # If the character is illegal ...
                if allowedChars.find(char) == -1:
                    # ... Display an error message to the user.
                    msg = _('There is an unsupported character in the Media File Name.\n\n"%s" includes the "%s" character, \nwhich Transana on the Mac does not support at this time.  Please rename your folders \nand files so that they do not include characters that are not part of English.')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8')
                    dlg = Dialogs.ErrorDialog(self, msg % (self.fname_edit.GetValue(), char))
                    dlg.ShowModal()
                    dlg.Destroy()
                    # Clear the file name field
                    self.fname_edit.SetValue('')
                    # We only need to detect the first error
                    break
        
    def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        import TransanaConstants
        self.obj.media_filename = self.fname_edit.GetValue()
        dirName = os.path.dirname(self.obj.media_filename)
        # If no path can be extracted from the file file name, start with the Video Root Folder
        if dirName == '':
            dirName = TransanaGlobal.configData.videoPath
        # Get the File Name
        fileName = os.path.basename(self.obj.media_filename)
        # Determine the File's extension
        (fn, ext) = os.path.splitext(self.obj.media_filename)
        # Have the default File Type match the extension, but use "All Videos" if blank.
        # If it's an unrecognized type, go to "All Files"
        if ext.lower() in ['.mpg', '.avi', '.mov', '.mp4', '.wmv', '.mp3', '.wav', '.wma']:
            fileType =  '*.mpg;*.avi;*.mov;*.mp4;*.wmv;*.mp3;*.wav;*.wma'
        else:
            fileType = ''
            
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

            self.obj.media_filename = fs
            self.fname_edit.SetValue(fs)
            if self.id_edit.GetValue() == '':
                tempFilename = os.path.basename(fs)
                (self.obj.id, tempExt) = os.path.splitext(tempFilename)
                self.id_edit.SetValue(self.obj.id)
            # The OnMediaFilenameEdit method doesn't get automatically triggered on the Mac.
            if "__WXMAC__" in wx.PlatformInfo:
                self.OnMediaFilenameEdit(evt)

    def OnCoreDataClick(self, event):
        """ Method for when Core Data button is clicked """
        # We can only edit Core Data once the Media Filename has been specified
        if (self.fname_edit.GetValue() != '') and os.path.exists(self.fname_edit.GetValue()):
            # Extract the path-less File Name from the full-path Media File Name
            (path, filename) = os.path.split(self.fname_edit.GetValue())   # (self.obj.media_filename)

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

            # Load the full-path filename into the Core Data Object's Comment Field
            coreData.comment = self.fname_edit.GetValue()    # self.obj.media_filename
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
            if (self.fname_edit.GetValue() == ''):
                errordlg = Dialogs.ErrorDialog(self, _("You must specify a Media Filename before defining Core Data"))
                errordlg.ShowModal()
                errordlg.Destroy()
            else:
                errordlg = Dialogs.ErrorDialog(self, _("The Media Filename must point to a file that exists before defining Core Data"))
                errordlg.ShowModal()
                errordlg.Destroy()

    def OnAddKW(self, evt):
        """Invoked when the user activates the Add Keyword (>>) button."""
        kw_name = self.kw_lb.GetStringSelection()
        if not kw_name:
            return
        kwg_name = self.kw_group_lb.GetStringSelection()
        if not kwg_name:
            # This shouldn't really happen anymore though
            return
        ep_kw = "%s : %s" % (kwg_name, kw_name)
        if self.ekw_lb.FindString(ep_kw) == -1:
            self.obj.add_keyword(kwg_name, kw_name)
            self.ekw_lb.Append(ep_kw)
            
        
    def OnRemoveKW(self, evt):
        """Invoked when the user activates the Remove Keyword (<<) button."""
        sel = self.ekw_lb.GetSelection()
        if sel > -1:
            # Separate out the Keyword Group and the Keyword
            kwlist = string.split(self.ekw_lb.GetStringSelection(), ':')
            kwg = string.strip(kwlist[0])
            kw = ':'.join(kwlist[1:]).strip()
            delResult = self.obj.remove_keyword(kwg, kw)
            if delResult and (sel >= 0):
                self.ekw_lb.Delete(sel)

 
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
        if gen_input:
            self.obj.id = gen_input[_('Episode ID')]
            self.obj.media_filename = gen_input[_('Media Filename')]
            self.obj.tape_date = self.dt_edit.GetValue()
            self.obj.comment = gen_input[_('Comment')]
            # Keyword list is already updated via the OnAddKW() callback
        else:
            self.obj = None
        
        return self.obj
        

# This simple derrived class let's the user drop files onto an edit box
class EditBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, editbox):
        wx.FileDropTarget.__init__(self)
        self.editbox = editbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        self.editbox.SetValue(files[0])


class AddEpisodeDialog(EpisodePropertiesForm):
    """Dialog used when adding a new Episode."""

    def __init__(self, parent, id, series):
        obj = Episode.Episode()
        obj.owner = DBInterface.get_username()
        obj.series_num = series.number
        obj.series_id = series.id
        EpisodePropertiesForm.__init__(self, parent, id, _("Add Episode"), obj)

class EditEpisodeDialog(EpisodePropertiesForm):
    """Dialog used when editing Episode properties."""

    def __init__(self, parent, id, ep_object):
        EpisodePropertiesForm.__init__(self, parent, id, _("Episode Properties"), ep_object)
