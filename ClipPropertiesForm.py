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

"""  This module implements the Clip Properties form.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

import Episode
import Collection
import DBInterface
import Dialogs
import KWManager
import Misc
import TransanaGlobal

import wx
import os
import string
import sys
import TranscriptEditor

class ClipPropertiesForm(Dialogs.GenForm):
    """Form containing Clip fields."""

    def __init__(self, parent, id, title, clip_object):
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(600, 470), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Clip Properties')
        # Define the minimum size for this dialog as the initial size
        self.SetSizeHints(600, 470)
        # Remember the Parent Window
        self.parent = parent
        self.obj = clip_object
        # if Keywords that server as Keyword Examples are removed, we will need to remember them.
        # Then, when OK is pressed, the Keyword Example references in the Database Tree can be removed.
        # We can't remove them immediately in case the whole Clip Properties Edit process is cancelled.
        self.keywordExamplesToDelete = []

        ######################################################
        # Tedious GUI layout code follows
        ######################################################

        # Clip ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 23)  # 23% width
        lay.height.AsIs()
        self.id_edit = self.new_edit_box(_("Clip ID"), lay, self.obj.id)

        # Collection ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.RightOf(self.id_edit, 10)           # 10 right of Episode ID
        lay.width.PercentOf(self.panel, wx.Width, 23)  # 23% width
        lay.height.AsIs()
        collection_edit = self.new_edit_box(_("Collection ID"), lay, self.obj.collection_id)
        collection_edit.Enable(False)

        # Series ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.RightOf(collection_edit, 10)   # 10 right of Collection ID
        lay.width.PercentOf(self.panel, wx.Width, 23)  # 23% width
        lay.height.AsIs()
        series_edit = self.new_edit_box(_("Series ID"), lay, self.obj.series_id)
        series_edit.Enable(False)

        # Episode ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.RightOf(series_edit, 10)       # 10 right of Series ID
        lay.width.PercentOf(self.panel, wx.Width, 23)  # 23% width
        lay.height.AsIs()
        episode_edit = self.new_edit_box(_("Episode ID"), lay, self.obj.episode_id)
        episode_edit.Enable(False)

        # Media Filename Layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.id_edit, 10)              # 10 under ID
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 48)  # 48% width
        lay.height.AsIs()
        # If the media filename path is not empty, we should normalize the path specification
        if self.obj.media_filename == '':
            filePath = self.obj.media_filename
        else:
            filePath = os.path.normpath(self.obj.media_filename)
        self.fname_edit = self.new_edit_box(_("Media Filename"), lay, filePath)
        self.fname_edit.SetDropTarget(EditBoxFileDropTarget(self.fname_edit))
        
        # Clip Start layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.id_edit, 10)              # 10 under ID
        lay.left.RightOf(self.fname_edit, 10)   # 10 right of Media FileName
        lay.width.PercentOf(self.panel, wx.Width, 15)  # 15% width
        lay.height.AsIs()
        # Convert to HH:MM:SS.mm
        clip_start_edit = self.new_edit_box(_("Clip Start"), lay, Misc.time_in_ms_to_str(self.obj.clip_start))
        clip_start_edit.Enable(False)

        # Clip Stop layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.id_edit, 10)              # 10 under ID
        lay.left.RightOf(clip_start_edit, 10)   # 10 right of Clip start
        lay.width.PercentOf(self.panel, wx.Width, 15)  # 15% width
        lay.height.AsIs()
        # Convert to HH:MM:SS.mm
        clip_stop_edit = self.new_edit_box(_("Clip Stop"), lay, Misc.time_in_ms_to_str(self.obj.clip_stop))
        clip_stop_edit.Enable(False)

        # Clip Length layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.id_edit, 10)              # 10 under ID
        lay.left.RightOf(clip_stop_edit, 10)    # 10 right of Clip start
        lay.width.PercentOf(self.panel, wx.Width, 14)  # 14% width
        lay.height.AsIs()
        # Convert to HH:MM:SS.mm
        clip_length_edit = self.new_edit_box(_("Clip Length"), lay, Misc.time_in_ms_to_str(self.obj.clip_stop - self.obj.clip_start))
        clip_length_edit.Enable(False)

        # Title/Comment layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.fname_edit, 10)      # 10 under media filename
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        comment_edit = self.new_edit_box(_("Title/Comment"), lay, self.obj.comment)

        # Clip Text layout [label]
        lay = wx.LayoutConstraints()
        lay.top.Below(comment_edit, 10)         # 10 under comment
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 22)  # 22% width
        lay.height.AsIs()
        clip_text_lbl = wx.StaticText(self.panel, -1, _("Clip Text"))
        clip_text_lbl.SetConstraints(lay)

        # Clip Text layout
        lay = wx.LayoutConstraints()
        lay.top.Below(clip_text_lbl, 3)         # 3 under label
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.PercentOf(self.panel, wx.Height, 15) # 15% of frame height

        # Load the Transcript into an RTF Control so the RTF Encoding won't show
        self.text_edit = TranscriptEditor.TranscriptEditor(self.panel)
        self.text_edit.SetReadOnly(0)
        self.text_edit.Enable(False)
        self.text_edit.ProgressDlg = wx.ProgressDialog("Loading Transcript", \
                                               "Reading document stream", \
                                                maximum=100, \
                                                style=wx.PD_AUTO_HIDE)
        self.text_edit.LoadRTFData(self.obj.text)
        self.text_edit.ProgressDlg.Destroy()

        # This doesn't work!  Hidden text remains visible.
        self.text_edit.StyleSetVisible(self.text_edit.STYLE_HIDDEN, False)
        self.text_edit.codes_vis = 0

        self.text_edit.Enable(True)

        self.text_edit.SetConstraints(lay)

        # Keyword Group layout [label]
        lay = wx.LayoutConstraints()
        lay.top.Below(self.text_edit, 10)            # 10 under clip text
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

        # Load the parent Collection in order to determine the default Keyword Group
        tempCollection = Collection.Collection(self.obj.collection_num)
        self.kw_groups = DBInterface.list_of_keyword_groups()
        self.kw_group_lb = wx.ListBox(self.panel, 101, wx.DefaultPosition, wx.DefaultSize, self.kw_groups)
        self.kw_group_lb.SetConstraints(lay)
        # Select the Collection Default Keyword Group in the Keyword Group list
        if (tempCollection.keyword_group != '') and (self.kw_group_lb.FindString(tempCollection.keyword_group) != wx.NOT_FOUND):
            self.kw_group_lb.SetStringSelection(tempCollection.keyword_group)
        # If no Default Keyword Group is defined, select the first item in the list
        else:
            # but only if there IS a first item.
            if len(self.kw_groups) > 0:
                self.kw_group_lb.SetSelection(0)
        # If there's a selected keyword group ...
        if self.kw_group_lb.GetSelection() != wx.NOT_FOUND:
            # populate the Keywords list
            self.kw_list = \
                DBInterface.list_of_keywords_by_group(self.kw_group_lb.GetStringSelection())
        else:
            # If not, create a blank one
            self.kw_list = []
        wx.EVT_LISTBOX(self, 101, self.OnGroupSelect)

        # Keyword layout [label]
        lay = wx.LayoutConstraints()
        lay.top.Below(self.text_edit, 10)            # 10 under clip text
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
        wx.EVT_BUTTON(self, wx.ID_FILE2, self.OnAddKW)

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


        # Clip Keywords [label]
        lay = wx.LayoutConstraints()
        lay.top.Below(self.text_edit, 10)            # 10 under clip text
        lay.left.SameAs(add_kw, wx.Right, 10)        # 10 from Add keyword Button
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        txt = wx.StaticText(self.panel, -1, _("Clip Keywords"))
        txt.SetConstraints(lay)

        # Clip Keywords [list box]
        lay = wx.LayoutConstraints()
        lay.top.Below(txt, 3)                   # 3 under label
        lay.left.SameAs(txt, wx.Left)            # left same as label
        lay.width.SameAs(txt, wx.Width)          # width same as label
        lay.bottom.SameAs(self.panel, wx.Height, 50)   # 50 from bottom

        # Create an empty ListBox
        self.ekw_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.DefaultSize)
        # Populate the ListBox
        # If we are loading a defined Clips (Clip.number != 0), use the Clips's keywords
        if self.obj.number != 0:
            for clipKeyword in self.obj.keyword_list:
                self.ekw_lb.Append(clipKeyword.keywordPair)
        # If we are creating a NEW Clip (Clip.number == 0), use the Episode's Keywords as default Keywords
        else:
            if self.obj.episode_num != 0:
                tempEpisode = Episode.Episode(self.obj.episode_num)
                for clipKeyword in tempEpisode.keyword_list:
                    self.obj.add_keyword(clipKeyword.keywordGroup, clipKeyword.keyword)
                    self.ekw_lb.Append(clipKeyword.keywordPair)
                                
        self.ekw_lb.SetConstraints(lay)

        self.ekw_lb.Bind(wx.EVT_KEY_DOWN, self.OnKeywordKeyDown)

        # Because of the way Clips are created (with Drag&Drop / Cut&Paste functions), we have to trap the missing
        # ID error here.  Therefore, we need to override the EVT_BUTTON for the OK Button.
        # Since we don't have an object for the OK Button, we use FindWindowById to find it based on its ID.
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.FindWindowById(wx.ID_OK))

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

        self.id_edit.SetFocus()

    def OnOK(self, event):
        # Because of the way Clips are created (with Drag&Drop / Cut&Paste functions), we have to trap the missing
        # ID error here.  Duplicate ID Error is handled elsewhere.
        if self.id_edit.GetValue().strip() == '':
            # Display the error message
            dlg2 = Dialogs.ErrorDialog(self, _('Clip ID is required.'))
            dlg2.ShowModal()
            dlg2.Destroy()
            # Set the focus on the widget with the error.
            self.id_edit.SetFocus()
        else:
            # Continue on with the form's regular Button event
            event.Skip(True)
        

    def refresh_keyword_groups(self):
        """Refresh the keyword groups listbox."""
        self.kw_groups = DBInterface.list_of_keyword_groups()
        self.kw_group_lb.Clear()
        self.kw_group_lb.InsertItems(self.kw_groups, 0)
        self.kw_group_lb.SetSelection(0)

    def refresh_keywords(self):
        """Refresh the keywords listbox."""
        sel = self.kw_group_lb.GetStringSelection()
        if sel:
            self.kw_list = \
                DBInterface.list_of_keywords_by_group(sel)
            self.kw_lb.Clear()
            self.kw_lb.InsertItems(self.kw_list, 0)
        

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
            kw = string.strip(kwlist[1])
            # Remove the Keyword from the Clip Object (this CAN be overridden by the user!)
            delResult = self.obj.remove_keyword(kwg, kw)
            # Remove the Keyword from the Keywords list
            if (delResult != 0) and (sel >= 0):
                self.ekw_lb.Delete(sel)
                # If what we deleted was a Keyword Example, remember the crucial information
                if delResult == 2:
                    self.keywordExamplesToDelete.append((kwg, kw, self.obj.number))

 
    def OnKWManage(self, evt):
        """Invoked when the user activates the Keyword Management button."""
        # find out if there is a default keyword group
        if self.kw_group_lb.IsEmpty():
            sel = None
        else:
            sel = self.kw_group_lb.GetStringSelection()
        # Create and display the Keyword Management Dialog
        kwm = KWManager.KWManager(self, defaultKWGroup=sel, deleteEnabled=False)
        # Refresh the Keyword Groups list, in case it was changed.
        self.refresh_keyword_groups()
        # Make sure the last Keyword Group selected in the Keyword Management is selected when it gets closed.
        selPos = self.kw_group_lb.FindString(kwm.kw_group.GetStringSelection())
        if selPos == -1:
            selPos = 0
        self.kw_group_lb.SetSelection(selPos)
        # Refresh the Keyword List, in case it was changed.
        self.refresh_keywords()
        # We must refresh the Keyword List in the DBTree to reflect changes made in the
        # Keyword Management.

        #This doesn't work when we create a new clip, as the Parent isn't the DatabaseTreeTab!!
        #self.parent.tree.refresh_kwgroups_node()
        TransanaGlobal.menuWindow.ControlObject.DataWindow.DBTab.tree.refresh_kwgroups_node()

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
            self.obj.id = gen_input[_('Clip ID')]
            self.obj.comment = gen_input[_('Title/Comment')]
            self.obj.media_filename = gen_input[_('Media Filename')]
            # We use a different method of getting the Transcript Text!
            self.obj.text = self.text_edit.GetRTFBuffer()
            # Keyword list is already updated via the OnAddKW() callback
        else:
            self.obj = None
        
        return self.obj
        



# This simple derived class let's the user drop files onto an edit box
class EditBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, editbox):
        wx.FileDropTarget.__init__(self)
        self.editbox = editbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        self.editbox.SetValue(files[0])


class AddClipDialog(ClipPropertiesForm):
    """Dialog used when adding a new Clip."""
    # NOTE:  AddClipDialog is not like the AddDialog for other objects.  The Add for other objects
    #        creates the object and then opens the form.  Here, the Clip Object is created based on
    #        data specified in signalling the desire to create a Clip, so the object is created and
    #        mostly populated elsewhere and passed in here.  It's still an Add, though.

    def __init__(self, parent, id, clip_obj):
        ClipPropertiesForm.__init__(self, parent, id, _("Add Clip"), clip_obj)

class EditClipDialog(ClipPropertiesForm):
    """Dialog used when editing Clip properties."""

    def __init__(self, parent, id, clip_object):
        ClipPropertiesForm.__init__(self, parent, id, _("Clip Properties"), clip_object)


