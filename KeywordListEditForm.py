# Copyright (C) 2003 - 2005 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Edit Keyword List dialog box."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import Clip
import ClipKeywordObject
import Collection
import DBInterface
import Dialogs
import Episode
import Series
import KWManager
import Misc
import TransanaGlobal

import wx
import os
import string
import sys

class KeywordListEditForm(Dialogs.GenForm):
    """Form containing Keyword List Edit fields."""

    def __init__(self, parent, id, title, obj, keywords):
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, (600,385), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Edit Keywords')
        # Define the minimum size for this dialog as the initial size
        self.SetSizeHints(600, 385)

        # Remember the parent Window
        self.parent = parent
        self.obj = obj
        # if Keywords that server as Keyword Examples are removed, we will need to remember them.
        # Then, when OK is pressed, the Keyword Example references in the Database Tree can be removed.
        # We can't remove them immediately in case the whole Clip Properties Edit process is cancelled.
        self.keywordExamplesToDelete = []

        # COPY the keyword list, rather than just pointing to it, so that the list on this form will
        # be independent of the original list.  That way, pressing CANCEL does not cause the list to
        # be changed anyway, though it means that if OK is pressed, you must copy the list to update
        # it in the calling routine.
        self.keywords = []
        for kws in keywords:
            self.keywords.append(kws)

        ######################################################
        # Tedious GUI layout code follows
        ######################################################

        # Keyword Group layout [label]
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Left, 10)        # 10 under top
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
        lay.bottom.SameAs(self.panel, wx.Height, 45)   # 45 from bottom

        # Get the Parent Object, so we can know the Default Keyword Group
        if type(self.obj) == type(Episode.Episode()):
            objParent = Series.Series(self.obj.series_num)
        elif type(self.obj) == type(Clip.Clip()):
            objParent = Collection.Collection(self.obj.collection_num)
        # Obtain the list of Keyword Groups
        self.kw_groups = DBInterface.list_of_keyword_groups()
        self.kw_group_lb = wx.ListBox(self.panel, 101, wx.DefaultPosition, wx.DefaultSize, self.kw_groups)
        self.kw_group_lb.SetConstraints(lay)
        if len(self.kw_groups) > 0:
            # Set the Keyword Group to the Default keyword Group
            if self.kw_group_lb.FindString(objParent.keyword_group) != wx.NOT_FOUND:
                self.kw_group_lb.SetStringSelection(objParent.keyword_group)
            # Obtain the list of Keywords for the intial Keyword Group
            self.kw_list = DBInterface.list_of_keywords_by_group(self.kw_group_lb.GetStringSelection())
        else:
            self.kw_list = []
        wx.EVT_LISTBOX(self, 101, self.OnGroupSelect)

        # Keyword layout [label]
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Left, 10)         # 10 under comment
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
        lay.bottom.SameAs(self.panel, wx.Height, 45)   # 45 from bottom
        
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


        # Keywords [label]
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 under comment
        lay.left.SameAs(add_kw, wx.Right, 10)          # 10 from Add Keyword Button
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        txt = wx.StaticText(self.panel, -1, _("Keywords"))
        txt.SetConstraints(lay)

        # Keywords [list box]
        lay = wx.LayoutConstraints()
        lay.top.Below(txt, 3)                   # 3 under label
        lay.left.SameAs(txt, wx.Left)            # left same as label
        lay.width.SameAs(txt, wx.Width)          # width same as label
        lay.bottom.SameAs(self.panel, wx.Height, 45)   # 45 from bottom
        
        # Create an empty ListBox
        self.ekw_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.DefaultSize)
        # Populate the ListBox
        for clipKeyword in self.keywords:
            self.ekw_lb.Append(clipKeyword.keywordPair)
        self.ekw_lb.SetConstraints(lay)
        self.ekw_lb.Bind(wx.EVT_KEY_DOWN, self.OnKeywordKeyDown)

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()
        
        self.kw_group_lb.SetFocus()
        

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
            self.kw_list = DBInterface.list_of_keywords_by_group(sel)
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

        # We need to check to see if the keyword is already in the keyword list
        keywordFound = False
        # Iterate through the list
        for clipKeyword in self.keywords:
            # If we find a match, set the flag and quit looking.
            if (clipKeyword.keywordGroup == kwg_name) and (clipKeyword.keyword == kw_name):
                keywordFound = True
                break

        # If the keyword is not found, add it.  (If it's already there, we don't need to do anything!)
        if not keywordFound:
            # Create an appropriate ClipKeyword Object
            tempClipKeyword = ClipKeywordObject.ClipKeyword(kwg_name, kw_name)
            # Add it to the Keyword List
            self.keywords.append(tempClipKeyword)
            self.ekw_lb.Append(ep_kw)
            
        
    def OnRemoveKW(self, evt):
        """Invoked when the user activates the Remove Keyword (<<) button."""
        sel = self.ekw_lb.GetSelection()
        if sel > wx.NOT_FOUND:
            # Separate out the Keyword Group and the Keyword
            kwlist = string.split(self.ekw_lb.GetStringSelection(), ':')
            kwg = string.strip(kwlist[0])
            kw = string.strip(kwlist[1])
            for index in range(len(self.keywords)):
                # Look for the entry to be deleted
                if (self.keywords[index].keywordGroup == kwg) and (self.keywords[index].keyword == kw):
                    if self.keywords[index].example == 1:
                        dlg = wx.MessageDialog(self, _('Clip "%s" has been designated as an example of Keyword "%s : %s".\nRemoving this Keyword from the Clip will also remove the Clip as a Keyword Example.\n\nDo you want to remove Clip "%s" as an example of Keyword "%s : %s"?') % (self.obj.id, kwg, kw, self.obj.id, kwg, kw), _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
                        result = dlg.ShowModal()
                        dlg.Destroy()
                        if result == wx.ID_YES:
                            # If the entry is found, delete it and stop looking
                            del self.keywords[index]
                            if sel >= 0:
                                self.ekw_lb.Delete(sel)
                            # If what we deleted was a Keyword Example, remember the crucial information
                            self.keywordExamplesToDelete.append((kwg, kw, self.obj.number))
                    else:
                        # If the entry is found, delete it and stop looking
                        del self.keywords[index]
                        if sel >= 0:
                            self.ekw_lb.Delete(sel)
                    break

 
    def OnKWManage(self, evt):
        """Invoked when the user activates the Keyword Management button."""
        # Create and display the Keyword Management Dialog
        kwm = KWManager.KWManager(self, defaultKWGroup=self.kw_group_lb.GetStringSelection(), deleteEnabled=False)
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
        self.parent.ControlObject.DataWindow.DBTab.tree.refresh_kwgroups_node()

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
