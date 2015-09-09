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

"""This module implements the Edit Keyword List dialog box."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import Clip
import ClipKeywordObject
import Collection
import DBInterface
import Dialogs
import Document
import Episode
import Quote
import Library
import Snapshot
import KWManager
import Misc
import TransanaGlobal
# Import Transana's Images
import TransanaImages

import wx
import os
import string
import sys

class KeywordListEditForm(Dialogs.GenForm):
    """Form containing Keyword List Edit fields."""

    def __init__(self, parent, id, title, obj, keywords):
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=TransanaGlobal.configData.keywordListEditSize,
                                 style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, useSizers = True, HelpContext='Edit Keywords')

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

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Keyword Group [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword Group"))
        v1.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword Group layout [list box]

        # Create an empty Keyword Group List for now.  We'll populate it later (for layout reasons)
        self.kw_groups = []
        self.kw_group_lb = wx.ListBox(self.panel, 101, wx.DefaultPosition, wx.DefaultSize, self.kw_groups)
        v1.Add(self.kw_group_lb, 1, wx.EXPAND)

        # Add the element to the sizer
        r1Sizer.Add(v1, 1, wx.EXPAND)

        self.kw_list = []
        wx.EVT_LISTBOX(self, 101, self.OnGroupSelect)

        # Add a horizontal spacer
        r1Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Keyword [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword"))
        v2.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword [list box]
        self.kw_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.DefaultSize, self.kw_list, style=wx.LB_EXTENDED)
        v2.Add(self.kw_lb, 1, wx.EXPAND)

        wx.EVT_LISTBOX_DCLICK(self, self.kw_lb.GetId(), self.OnAddKW)

        # Add the element to the sizer
        r1Sizer.Add(v2, 1, wx.EXPAND)

        # Add a horizontal spacer
        r1Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v3 = wx.BoxSizer(wx.VERTICAL)
        # Keyword transfer buttons
        add_kw = wx.Button(self.panel, wx.ID_FILE2, ">>", wx.DefaultPosition)
        v3.Add(add_kw, 0, wx.EXPAND | wx.TOP, 20)
        wx.EVT_BUTTON(self, wx.ID_FILE2, self.OnAddKW)

        rm_kw = wx.Button(self.panel, wx.ID_FILE3, "<<", wx.DefaultPosition)
        v3.Add(rm_kw, 0, wx.EXPAND | wx.TOP, 10)
        wx.EVT_BUTTON(self, wx.ID_FILE3, self.OnRemoveKW)

        kwm = wx.BitmapButton(self.panel, wx.ID_FILE4, TransanaImages.KWManage.GetBitmap())
        v3.Add(kwm, 0, wx.EXPAND | wx.TOP, 10)
        # Add a spacer to increase the height of the Keywords section
        v3.Add((0, 60))
        kwm.SetToolTipString(_("Keyword Management"))
        wx.EVT_BUTTON(self, wx.ID_FILE4, self.OnKWManage)

        # Add the element to the sizer
        r1Sizer.Add(v3, 0)

        # Add a horizontal spacer
        r1Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Keywords [label]
        txt = wx.StaticText(self.panel, -1, _("Keywords"))
        v4.Add(txt, 0, wx.BOTTOM, 3)

        # Keywords [list box]
        # Create an empty ListBox
        self.ekw_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.DefaultSize, style=wx.LB_EXTENDED)
        v4.Add(self.ekw_lb, 1, wx.EXPAND)
        self.ekw_lb.Bind(wx.EVT_KEY_DOWN, self.OnKeywordKeyDown)

        # Add the element to the sizer
        r1Sizer.Add(v4, 2, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 1, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
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
        self.SetSize(wx.Size(max(600, width), max(385, height)))
        # Define the minimum size for this dialog as the current size
        self.SetSizeHints(max(600, width), max(385, height))
        # Center the form on screen
        TransanaGlobal.CenterOnPrimary(self)

        # We need to set some minimum sizes so the sizers will work right
        self.kw_group_lb.SetSizeHints(minW = 50, minH = 20)
        self.kw_lb.SetSizeHints(minW = 50, minH = 20)
        self.ekw_lb.SetSizeHints(minW = 50, minH = 20)

        # We need to capture the OK and Cancel button clicks locally.  We'll use FindWindowByID to locate the correct widgets.
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.FindWindowById(wx.ID_OK))
        self.Bind(wx.EVT_BUTTON, self.OnCancel, self.FindWindowById(wx.ID_CANCEL))
        
        # We populate the Keyword Groups, Keywords, and Clip Keywords lists AFTER we determine the Form Size.
        # Long Keywords in the list were making the form too big!

        # Obtain the list of Keyword Groups
        self.kw_groups = DBInterface.list_of_keyword_groups()
        for keywordGroup in self.kw_groups:
            self.kw_group_lb.Append(keywordGroup)
        if self.kw_group_lb.GetCount() > 0:
            self.kw_group_lb.EnsureVisible(0)

        # Get the Parent Object, so we can know the Default Keyword Group
        if isinstance(self.obj, Episode.Episode):
            objParent = Library.Library(self.obj.series_num)
        elif isinstance(self.obj, Document.Document):
            objParent = Library.Library(self.obj.library_num)
        elif isinstance(self.obj, Clip.Clip) or isinstance(self.obj, Snapshot.Snapshot) or isinstance(self.obj, Quote.Quote):
            objParent = Collection.Collection(self.obj.collection_num)
        if len(self.kw_groups) > 0:
            # Set the Keyword Group to the Default keyword Group
            if self.kw_group_lb.FindString(objParent.keyword_group) != wx.NOT_FOUND:
                self.kw_group_lb.SetStringSelection(objParent.keyword_group)
            else:
                self.kw_group_lb.SetSelection(0)
            # Obtain the list of Keywords for the intial Keyword Group
            self.kw_list = DBInterface.list_of_keywords_by_group(self.kw_group_lb.GetStringSelection())
        else:
            self.kw_list = []
        for keyword in self.kw_list:
            self.kw_lb.Append(keyword)
        if self.kw_lb.GetCount() > 0:
            self.kw_lb.EnsureVisible(0)

        # Populate the ListBox
        for clipKeyword in self.keywords:
            self.ekw_lb.Append(clipKeyword.keywordPair)
        if self.ekw_lb.GetCount() > 0:
            self.ekw_lb.EnsureVisible(0)

        self.kw_group_lb.SetFocus()

    def OnOK(self, event):
        """ Intercept the OK button click and process it """
        # Remember the SIZE of the dialog box for next time.
        TransanaGlobal.configData.keywordListEditSize = self.GetSize()
        # Now process the event normally
        event.Skip(True)

    def OnCancel(self, event):
        """ Intercept the Cancel button click and process it """
        # Remember the SIZE of the dialog box for next time.
        TransanaGlobal.configData.keywordListEditSize = self.GetSize()
        # now process the event normally
        event.Skip(True)

    def refresh_keyword_groups(self):
        """Refresh the keyword groups listbox."""
        self.kw_groups = DBInterface.list_of_keyword_groups()
        self.kw_group_lb.Clear()
        self.kw_group_lb.InsertItems(self.kw_groups, 0)
        if len(self.kw_groups) > 0:
            self.kw_group_lb.SetSelection(0)
        if self.kw_group_lb.GetCount() > 0:
            self.kw_group_lb.EnsureVisible(0)

    def refresh_keywords(self):
        """Refresh the keywords listbox."""
        sel = self.kw_group_lb.GetStringSelection()
        if sel:
            self.kw_list = DBInterface.list_of_keywords_by_group(sel)
            self.kw_lb.Clear()
            if len(self.kw_list) > 0:
                self.kw_lb.InsertItems(self.kw_list, 0)
                if self.kw_lb.GetCount() > 0:
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
        # Get the selection(s) from the Episode Keywords list box
        kwitems = self.ekw_lb.GetSelections()
        # The items are returned as an immutable tuple.  Convert this to a list.
        kwitems = list(kwitems)
        # Now sort the list.  For reasons that elude me, the list is arbitrarily ordered on the Mac, which causes
        # deletes to be done out of order so the wrong elements get deleted, which is BAD.
        kwitems.sort()
        # We have to go through the list items BACKWARDS so that item numbers don't change on us as we delete items!
        for item in range(len(kwitems), 0, -1):
            sel = kwitems[item - 1]
            # Separate out the Keyword Group and the Keyword
            kwlist = string.split(self.ekw_lb.GetString(sel), ':')
            kwg = string.strip(kwlist[0])
            kw = ':'.join(kwlist[1:]).strip()
            for index in range(len(self.keywords)):
                # Look for the entry to be deleted
                if (self.keywords[index].keywordGroup == kwg) and (self.keywords[index].keyword == kw):
                    if self.keywords[index].example == 1:
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('Clip "%s" has been designated as an example of Keyword "%s : %s".\nRemoving this Keyword from the Clip will also remove the Clip as a Keyword Example.\n\nDo you want to remove Clip "%s" as an example of Keyword "%s : %s"?'), 'utf8')
                        else:
                            prompt = _('Clip "%s" has been designated as an example of Keyword "%s : %s".\nRemoving this Keyword from the Clip will also remove the Clip as a Keyword Example.\n\nDo you want to remove Clip "%s" as an example of Keyword "%s : %s"?')
                        dlg = Dialogs.QuestionDialog(self, prompt % (self.obj.id, kwg, kw, self.obj.id, kwg, kw))
                        result = dlg.LocalShowModal()
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
        if not self.kw_group_lb.IsEmpty():
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
