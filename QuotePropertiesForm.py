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

"""  This module implements the Quote Properties form.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import Collection
import DBInterface
import Dialogs
import Document
import KWManager
import Misc
import PyXML_RTCImportParser
import Quote
import Library
import TransanaConstants
import TransanaExceptions
import TransanaGlobal
# Import Transana's Images
import TransanaImages
# Transana's RTF Editor Control
import TranscriptEditor_RTC

import wx
import os
import string
import sys
import xml.sax

class QuotePropertiesForm(Dialogs.GenForm):
    """Form containing Quote fields."""

    def __init__(self, parent, id, title, quote_object, mergeList=None):
        # If no Merge List is passed in ...
        if mergeList == None:
            # ... use the default Quote Properties Dialog size passed in from the config object.  (This does NOT get saved.)
            size = TransanaGlobal.configData.quotePropertiesSize
            # ... we can enable Quote Change Propagation
            propagateEnabled = True
            HelpContext='Quote Properties'
        # If we DO have a merge list ...
        else:
            # ... increase the size of the dialog to allow for the list of mergeable quotes.
            size = (TransanaGlobal.configData.quotePropertiesSize[0], TransanaGlobal.configData.quotePropertiesSize[1] + 130)
            # ... and Quote Change Propagation should be disabled
            propagateEnabled = False
            HelpContext='Quote Merge'

        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style.  Signal that Propogation is included.
        Dialogs.GenForm.__init__(self, parent, id, title, size=size, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 propagateEnabled=propagateEnabled, useSizers = True, HelpContext=HelpContext)
        if mergeList == None:
            # Define the minimum size for this dialog as the initial size
            minWidth = 750
            minHeight = 570
        else:
            # Define the minimum size for this dialog as the initial size
            minWidth = 750
            minHeight = 650
        # Remember the Parent Window
        self.parent = parent
        # Remember the original Quote Object passed in
        self.obj = quote_object
        # Add a placeholder to the quote object for the merge quote number
        self.obj.mergeNumber = 0
        # Remember the merge list, if one is passed in
        self.mergeList = mergeList
        # Initialize the merge item as unselected
        self.mergeItemIndex = -1
        # if Keywords that server as Keyword Examples are removed, we will need to remember them.
        # Then, when OK is pressed, the Keyword Example references in the Database Tree can be removed.
        # We can't remove them immediately in case the whole Quote Properties Edit process is cancelled.
        self.keywordExamplesToDelete = []
        # Initialize a variable to hold merged keyword examples.
        self.mergedKeywordExamples = []

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # If we're merging Quotes ...
        if self.mergeList != None:

            # ... display a label for the Merge Quotes ...
            lblMergeQuote = wx.StaticText(self.panel, -1, _("Quote to Merge"))
            mainSizer.Add(lblMergeQuote, 0)
            
            # Add a vertical spacer to the main sizer        
            mainSizer.Add((0, 3))

            # Create a HORIZONTAL sizer for the merge information
            mergeSizer = wx.BoxSizer(wx.HORIZONTAL)
            
            # ... display a ListCtrl for the Merge Quotes ...
            self.mergeQuotes = wx.ListCtrl(self.panel, -1, size=(300, 100), style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
            # Add the element to the sizer
            mergeSizer.Add(self.mergeQuotes, 1, wx.EXPAND)
            # ... bind the Item Selected event for the List Control ...
            self.mergeQuotes.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected)

            # ... define the columns for the Merge Quotes ...
            self.mergeQuotes.InsertColumn(0, _('Quote Name'))
            self.mergeQuotes.InsertColumn(1, _('Collection'))
            self.mergeQuotes.InsertColumn(2, _('Start Character'))
            self.mergeQuotes.InsertColumn(3, _('End Characer'))
            self.mergeQuotes.SetColumnWidth(0, 244)
            self.mergeQuotes.SetColumnWidth(1, 244)
            # ... and populate the Merge Quotes list from the mergeList data
            for (QuoteNum, QuoteID, DocumentNum, CollectNum, CollectID, StartChar, EndChar) in self.mergeList:
                index = self.mergeQuotes.InsertStringItem(sys.maxint, QuoteID)
                self.mergeQuotes.SetStringItem(index, 1, CollectID)
                self.mergeQuotes.SetStringItem(index, 2, u"%s" % StartChar)
                self.mergeQuotes.SetStringItem(index, 3, u"%s" % EndChar)

            # Add the row sizer to the main vertical sizer
            mainSizer.Add(mergeSizer, 3, wx.EXPAND)

            # Add a vertical spacer to the main sizer        
            mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Quote ID
        self.id_edit = self.new_edit_box(_("Quote ID"), v1, self.obj.id, maxLen=100)
        # Add the element to the sizer
        r1Sizer.Add(v1, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r1Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Collection ID
        collection_edit = self.new_edit_box(_("Collection ID"), v2, self.obj.GetNodeString(False))
        # Add the element to the sizer
        r1Sizer.Add(v2, 2, wx.EXPAND)
        collection_edit.Enable(False)
        
        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # If the source document is known ...
        if self.obj.source_document_num > 0:
            try:
                tmpSrcDoc = Document.Document(num = self.obj.source_document_num)
                srcDocID = tmpSrcDoc.id
                tmpLibrary = Library.Library(tmpSrcDoc.library_num)
                libraryID = tmpLibrary.id
            except TransanaExceptions.RecordNotFoundError:
                srcDocID = ''
                libraryID = ''
                self.obj.source_document_id = 0
            except:
                print
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                self.obj.source_document_id = 0
                srcDocID = ''
                libraryID = ''
        else:
            srcDocID = ''
            libraryID = ''            
        # Create a VERTICAL sizer for the next element
        v3 = wx.BoxSizer(wx.VERTICAL)
        # Library ID
        library_edit = self.new_edit_box(_("Library ID"), v3, libraryID)
        # Add the element to the sizer
        r2Sizer.Add(v3, 1, wx.EXPAND)
        library_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Document ID
        document_edit = self.new_edit_box(_("Document ID"), v4, srcDocID)
        # Add the element to the sizer
        r2Sizer.Add(v4, 1, wx.EXPAND)
        document_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r4Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v5 = wx.BoxSizer(wx.VERTICAL)
        # Quote Start Character
        self.quote_start_edit = self.new_edit_box(_("Quote Start Character"), v5, unicode(self.obj.start_char))
        # Add the element to the sizer
        r4Sizer.Add(v5, 1, wx.EXPAND)
        # For merging, we need to remember the merged value of Quote Start Char
        self.start_char = self.obj.start_char
        self.quote_start_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r4Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v6 = wx.BoxSizer(wx.VERTICAL)
        # Quote End Character
        self.quote_end_edit = self.new_edit_box(_("Quote End Character"), v6, unicode(self.obj.end_char))
        # Add the element to the sizer
        r4Sizer.Add(v6, 1, wx.EXPAND)
        # For merging, we need to remember the merged value of Quote End Char
        self.end_char = self.obj.end_char
        self.quote_end_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r4Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v7 = wx.BoxSizer(wx.VERTICAL)
        # Quote Length
        self.quote_length_edit = self.new_edit_box(_("Quote Length"), v7, unicode(self.obj.end_char - self.obj.start_char))
        # Add the element to the sizer
        r4Sizer.Add(v7, 1, wx.EXPAND)
        self.quote_length_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r4Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r5Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v8 = wx.BoxSizer(wx.VERTICAL)
        # Comment
        comment_edit = self.new_edit_box(_("Comment"), v8, self.obj.comment, maxLen=255)
        # Add the element to the sizer
        r5Sizer.Add(v8, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r5Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        txt = wx.StaticText(self.panel, -1, _("Quote Text"))
        mainSizer.Add(txt, 0, wx.BOTTOM, 3)
        self.text_edit = TranscriptEditor_RTC.TranscriptEditor(self.panel)
        self.text_edit.load_transcript(self.obj)
        self.text_edit.SetReadOnly(False)
        self.text_edit.Enable(True)
        mainSizer.Add(self.text_edit, 6, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a HORIZONTAL sizer for the next row
        r7Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v9 = wx.BoxSizer(wx.VERTICAL)
        # Keyword Group [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword Group"))
        v9.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword Group [list box]

        # Create an empty Keyword Group List for now.  We'll populate it later (for layout reasons)
        self.kw_groups = []
        self.kw_group_lb = wx.ListBox(self.panel, -1, choices = self.kw_groups)
        v9.Add(self.kw_group_lb, 1, wx.EXPAND)
        
        # Add the element to the sizer
        r7Sizer.Add(v9, 1, wx.EXPAND)

        # Create an empty Keyword List for now.  We'll populate it later (for layout reasons)
        self.kw_list = []
        wx.EVT_LISTBOX(self, self.kw_group_lb.GetId(), self.OnGroupSelect)

        # Add a horizontal spacer
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v10 = wx.BoxSizer(wx.VERTICAL)
        # Keyword [label]
        txt = wx.StaticText(self.panel, -1, _("Keyword"))
        v10.Add(txt, 0, wx.BOTTOM, 3)

        # Keyword [list box]
        self.kw_lb = wx.ListBox(self.panel, -1, choices = self.kw_list, style=wx.LB_EXTENDED)
        v10.Add(self.kw_lb, 1, wx.EXPAND)

        wx.EVT_LISTBOX_DCLICK(self, self.kw_lb.GetId(), self.OnAddKW)

        # Add the element to the sizer
        r7Sizer.Add(v10, 1, wx.EXPAND)

        # Add a horizontal spacer
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v11 = wx.BoxSizer(wx.VERTICAL)
        # Keyword transfer buttons
        add_kw = wx.Button(self.panel, wx.ID_FILE2, ">>", wx.DefaultPosition)
        v11.Add(add_kw, 0, wx.EXPAND | wx.TOP, 20)
        wx.EVT_BUTTON(self, wx.ID_FILE2, self.OnAddKW)

        rm_kw = wx.Button(self.panel, wx.ID_FILE3, "<<", wx.DefaultPosition)
        v11.Add(rm_kw, 0, wx.EXPAND | wx.TOP, 10)
        wx.EVT_BUTTON(self, wx.ID_FILE3, self.OnRemoveKW)

        kwm = wx.BitmapButton(self.panel, wx.ID_FILE4, TransanaImages.KWManage.GetBitmap())
        v11.Add(kwm, 0, wx.EXPAND | wx.TOP, 10)
        kwm.SetToolTipString(_("Keyword Management"))
        wx.EVT_BUTTON(self, wx.ID_FILE4, self.OnKWManage)

        # Add the element to the sizer
        r7Sizer.Add(v11, 0)

        # Add a horizontal spacer
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v12 = wx.BoxSizer(wx.VERTICAL)

        # Quote Keywords [label]
        txt = wx.StaticText(self.panel, -1, _("Quote Keywords"))
        v12.Add(txt, 0, wx.BOTTOM, 3)

        # Quote Keywords [list box]
        # Create an empty ListBox.  We'll populate it later for layout reasons.
        self.ekw_lb = wx.ListBox(self.panel, -1, style=wx.LB_EXTENDED)
        v12.Add(self.ekw_lb, 1, wx.EXPAND)

        self.ekw_lb.Bind(wx.EVT_KEY_DOWN, self.OnKeywordKeyDown)

        # Add the element to the sizer
        r7Sizer.Add(v12, 2, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r7Sizer, 8, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the buttons
        self.create_buttons(sizer=btnSizer)
        # Add the button sizer to the main sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)

        # Because of the way Quotes are created (with Drag&Drop / Cut&Paste functions), we have to trap the missing
        # ID error here.  Therefore, we need to override the EVT_BUTTON for the OK Button.
        # Since we don't have an object for the OK Button, we use FindWindowById to find it based on its ID.
        self.Bind(wx.EVT_BUTTON, self.OnOK, self.FindWindowById(wx.ID_OK))
        # We also need to intercept the Cancel button.
        self.Bind(wx.EVT_BUTTON, self.OnCancel, self.FindWindowById(wx.ID_CANCEL))

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
        # Determine which monitor to use and get its size and position
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            primaryScreen = TransanaGlobal.configData.primaryScreen
        else:
            primaryScreen = 0
        rect = wx.Display(primaryScreen).GetClientArea()

        # Reset the form's size to be at least the specified minimum width
        self.SetSize(wx.Size(max(minWidth, width), min(max(minHeight, height), rect[3])))
        # Define the minimum size for this dialog as the current size
        self.SetSizeHints(max(minWidth, width), min(max(minHeight, height), rect[3]))
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

        # Populate the Keywords ListBox
        # Load the parent Collection in order to determine the default Keyword Group
        tempCollection = Collection.Collection(self.obj.collection_num)
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
        for keyword in self.kw_list:
            self.kw_lb.Append(keyword)

        # Populate the Quote Keywords ListBox
        # If the quote object has keywords ... 
        for quoteKeyword in self.obj.keyword_list:
            # ... add them to the keyword list
            self.ekw_lb.Append(quoteKeyword.keywordPair)

        # Set initial focus to the Quote ID
        self.id_edit.SetFocus()

    def OnOK(self, event):
        """ Intercept the OK button click and process it """
        # Remember the SIZE of the dialog box for next time.  Remember, the size has been altered if
        # we are merging Quotes, and we need to remember the NON-MODIFIED size!
        if self.mergeList == None:
            TransanaGlobal.configData.quotePropertiesSize = self.GetSize()
        else:
            TransanaGlobal.configData.quotePropertiesSize = (self.GetSize()[0], self.GetSize()[1] - 130)
        # Because of the way Quotes are created (with Drag&Drop / Cut&Paste functions), we have to trap the missing
        # ID error here.  Duplicate ID Error is handled elsewhere.
        if self.id_edit.GetValue().strip() == '':
            # Display the error message
            dlg2 = Dialogs.ErrorDialog(self, _('Quote ID is required.'))
            dlg2.ShowModal()
            dlg2.Destroy()
            # Set the focus on the widget with the error.
            self.id_edit.SetFocus()
        else:
            # Continue on with the form's regular Button event
            event.Skip(True)
        
    def OnCancel(self, event):
        """ Intercept the Cancel button click and process it """
        # Remember the SIZE of the dialog box for next time.  Remember, the size has been altered if
        # we are merging Quotes, and we need to remember the NON-MODIFIED size!
        if self.mergeList == None:
            TransanaGlobal.configData.quotePropertiesSize = self.GetSize()
        else:
            TransanaGlobal.configData.quotePropertiesSize = (self.GetSize()[0], self.GetSize()[1] - 130)
        # Now process the event normally, as if it hadn't been intercepted
        event.Skip(True)
                
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
            # ... and if it's NOT already in the Quote Keywords list ...
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
            sel = kwitems[item - 1]
            # Separate out the Keyword Group and the Keyword
            kwlist = string.split(self.ekw_lb.GetString(sel), ':')
            kwg = string.strip(kwlist[0])
            kw = ':'.join(kwlist[1:]).strip()
            # If the selected keyword is in the current clip object ...
            if self.obj.has_keyword(kwg, kw):
                # Remove the Keyword from the Clip Object (this CAN be overridden by the user!)
                delResult = self.obj.remove_keyword(kwg, kw)
                # Remove the Keyword from the Keywords list
                if (delResult != 0) and (sel >= 0):
                    self.ekw_lb.Delete(sel)
                    # If what we deleted was a Keyword Example, remember the crucial information
                    if delResult == 2:
                        self.keywordExamplesToDelete.append((kwg, kw, self.obj.number))
            # If the selected keyword is NOT in the current object (and therefore has been merged into it) ...
            else:
                # See if it's a keyword example.
                if (kwg, kw) in self.mergedKeywordExamples:
                    # If so, prompt before removing it.
                    prompt = _('Clip "%s" has been defined as a Keyword Example for Keyword "%s : %s".')
                    data = (self.mergeList[self.mergeItemIndex][1], kwg, kw)
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(prompt, 'utf8') % data
                        prompt = prompt + unicode(_('\nAre you sure you want to delete it?'), 'utf8')
                    else:
                        prompt = prompt % data
                        prompt = prompt + _('\nAre you sure you want to delete it?')
                    dlg = Dialogs.QuestionDialog(None, prompt, _('Delete Clip'))
                    # If the user does NOT want to delete the keyword ...
                    if (dlg.LocalShowModal() == wx.ID_NO):
                        dlg.Destroy()
                        # ... then exit this method now before it gets removed.
                        return
                    else:
                        dlg.Destroy()
                        
               # ... merely remove it from the keyword list.  (Merged keywords are not stored in the original Clip object
                # in case the user changes the merge clip selection.)
                self.ekw_lb.Delete(sel)
 
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
        if not self.kw_group_lb.IsEmpty():
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
        """ Process the Key Down event for the Keyword field """
        # Begin exception handling, as some keys will raise an exception here
        try:
            # Get the key code of the pressed key
            c = event.GetKeyCode()
            # If it's the DELETE key ...
            if c == wx.WXK_DELETE:
                # ... see if there is a selection ...
                if self.ekw_lb.GetSelection() != wx.NOT_FOUND:
                    # ... and if so, delete this keyword!
                    self.OnRemoveKW(event)
        # If an exception occurs, ignore it!
        except:
            pass  # ignore non-ASCII keys

    def OnItemSelected(self, event):
        """ Process the selection of a Quote to be merged with the original Quote """
        # Identify the selected item
        self.mergeItemIndex = event.GetIndex()
        # Get the Quote Data for the selected item
        mergeQuote = Quote.Quote(self.mergeList[self.mergeItemIndex][0])

        # Merge the quotes.
        # First, determine whether the merging quote comes BEFORE or after the original
        if mergeQuote.start_char < self.obj.start_char:
            # The start value comes from the merge quote
            self.quote_start_edit.SetValue(u"%s" % mergeQuote.start_char)
            # Update the merged clip Start Character
            self.start_char = mergeQuote.start_char
            # The stop value comes from the original quote
            self.quote_end_edit.SetValue(u"%s" % self.obj.end_char)
            # Update the Quote Length
            self.quote_length_edit.SetValue("%s" % (self.obj.end_char - mergeQuote.start_char))
            # Update the merged Quote Stop Time
            self.end_char = self.obj.end_char
            # ... clear the Quote Text ...
            self.text_edit.ClearDoc(skipUnlock = True)
            # ... turn off read-only ...
            self.text_edit.SetReadOnly(0)
            # Create the Transana XML to RTC Import Parser.  This is needed so that we can
            # pull XML transcripts into the existing RTC.  Pass the RTC to be edited.
            handler = PyXML_RTCImportParser.XMLToRTCHandler(self.text_edit)
            # Parse the merge Quote text, adding it to the RTC
            xml.sax.parseString(mergeQuote.text, handler)
            # Add a blank line
            self.text_edit.Newline()
            # Parse the original Quote text, adding it to the RTC
            xml.sax.parseString(self.obj.text, handler)

        # If the merging Quote comes AFTER the original ...
        else:
            # The start value comes from the original Quote
            self.quote_start_edit.SetValue(u"%s" % self.obj.start_char)
            # Update the merged Quote Start Character
            self.start_char = self.obj.start_char
            # The end character value comes from the merge Quote
            self.quote_end_edit.SetValue(u"%s" % mergeQuote.end_char)
            # Update the Quote Length
            self.quote_length_edit.SetValue("%s" % (mergeQuote.end_char - self.obj.start_char))
            # Update the merged Quote End Character
            self.end_char = mergeQuote.end_char
            # ... clear the Quote Text ...
            self.text_edit.ClearDoc(skipUnlock = True)
            # ... turn off read-only ...
            self.text_edit.SetReadOnly(0)
            # Create the Transana XML to RTC Import Parser.  This is needed so that we can
            # pull XML Quote Text into the existing RTC.  Pass the RTC in.
            handler = PyXML_RTCImportParser.XMLToRTCHandler(self.text_edit)
            # Parse the original Quote text, adding it to the reportText RTC
            xml.sax.parseString(self.obj.text, handler)
            # Add a blank line
            self.text_edit.Newline()
            # ... now add the merge Quote's text ...
            # Parse the merge Quote text, adding it to the RTC
            xml.sax.parseString(mergeQuote.text, handler)

        # Remember the Merged Quote's Number
        self.obj.mergeNumber = mergeQuote.number
        # Create a list object for merging the keywords
        kwList = []
        # Add all the original keywords
        for kws in self.obj.keyword_list:
            kwList.append(kws.keywordPair)
        # Iterate through the merge Quote keywords.
        for kws in mergeQuote.keyword_list:
            # If they're not already in the list, add them.
            if not kws.keywordPair in kwList:
                kwList.append(kws.keywordPair)
        # Sort the keyword list
        kwList.sort()
        # Clear the keyword list box.
        self.ekw_lb.Clear()
        # Add the merged keywords to the list box.
        for kws in kwList:
            self.ekw_lb.Append(kws)

    def get_input(self):
        """Custom input routine."""
        # Inherit base input routine
        gen_input = Dialogs.GenForm.get_input(self)
        # If OK ...
        if gen_input:
            # Get main data values from the form, ID and Comment
            self.obj.id = gen_input[_('Quote ID')]
            self.obj.comment = gen_input[_('Comment')]
            # If we're merging Quotes, more data may have changed!
            if self.mergeList != None:
                # We need the post-merge Start Position
                self.obj.start_char = self.start_char
                # We need the post-merge Ending Position
                self.obj.end_char = self.end_char
                # When merging, keywords are not merged into the object automatically, so we need to process the Quote Keyword List
                for clkw in self.ekw_lb.GetStrings():
                    # Separate out the Keyword Group and the Keyword
                    kwlist = string.split(clkw, ':')
                    kwg = string.strip(kwlist[0])
                    kw = ':'.join(kwlist[1:]).strip()
                    # ... we need to add it to the list.  (Duplicates don't matter.)
                    self.obj.add_keyword(kwg, kw)

            self.obj.text = self.text_edit.GetFormattedSelection('XML')
            # Keyword list is already updated via the OnAddKW() callback
        else:
            self.obj = None

        return self.obj


class AddQuoteDialog(QuotePropertiesForm):
    """Dialog used when adding a new Quote."""
    # NOTE:  AddQuoteDialog is not like the AddDialog for other objects.  The Add for other objects
    #        creates the object and then opens the form.  Here, the Quote Object is created based on
    #        data specified in signalling the desire to create a Quote, so the object is created and
    #        mostly populated elsewhere and passed in here.  It's still an Add, though.

    def __init__(self, parent, id, quote_obj):
        QuotePropertiesForm.__init__(self, parent, id, _("Add Quote"), quote_obj)

class EditQuoteDialog(QuotePropertiesForm):
    """Dialog used when editing Quote properties."""

    def __init__(self, parent, id, quote_object):
        QuotePropertiesForm.__init__(self, parent, id, _("Quote Properties"), quote_object)
