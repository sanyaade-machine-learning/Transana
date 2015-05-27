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

""" This module implements the Transcript Properties form. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

# import wxPython
import wx
# import Transana's Database Interface
import DBInterface
# Import Transana's Dialogs
import Dialogs
# Import the Document Object
import Document
# import Transana's Keyword Management form
import KWManager
# import Transana's Library object
import Library
# import Transana's Globals
import TransanaGlobal
# import Transana's Images
import TransanaImages
# import Python's datetime module
import datetime
# import Python's os module
import os

class DocumentPropertiesForm(Dialogs.GenForm):
    """Form containing Document fields."""

    def __init__(self, parent, id, title, document_object):
        """ Create the Transcript Properties form """
        self.parent = parent
        self.width = 500
        self.height = 260
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(self.width, self.height), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Document Properties')

        # Define the form's main object
        self.obj = document_object
        # if a Library has been passed in ...
        if self.obj.library_num > 0:
            # ... get the Library's data
            library = Library.Library(self.obj.library_num)
        # If no Library has been passed in (ILLEGAL??) ...
        else:
            # ... create an empty Library
            library = Library.Library()

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Add the Transcript ID element
        self.id_edit = self.new_edit_box(_("Document ID"), v1, self.obj.id, maxLen=100)
        # Add the element to the sizer
        mainSizer.Add(v1, 1, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a VERTICAL sizer for the next element
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Add the Library ID element
        library_id_edit = self.new_edit_box(_("Library ID"), v2, library.id)
        # Add the element to the main sizer
        mainSizer.Add(v2, 1, wx.EXPAND)
        # Disable Library ID
        library_id_edit.Enable(False)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # If the Document has already been defined, we can skip this!
        if self.obj.number == 0:
            # Create a HORIZONTAL sizer for the next row
            r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

            # Create a VERTICAL sizer for the next element
            v3 = wx.BoxSizer(wx.VERTICAL)
            # Add the Import File element
            self.rtfname_edit = self.new_edit_box(_("RTF/XML/TXT File to import  (optional)"), v3, '')
            # Make this text box a File Drop Target
            self.rtfname_edit.SetDropTarget(EditBoxFileDropTarget(self.rtfname_edit))
            # Add the element to the row sizer
            r1Sizer.Add(v3, 1, wx.EXPAND)

            # Add a horizontal spacer to the row sizer        
            r1Sizer.Add((10, 0))

            # Add the Browse Button
            browse = wx.Button(self.panel, -1, _("Browse"))
            # Add the Browse Method to the Browse Button
            wx.EVT_BUTTON(self, browse.GetId(), self.OnBrowseClick)
            # Add the element to the sizer
            r1Sizer.Add(browse, 0, wx.ALIGN_BOTTOM)
            # If Mac ...
            if 'wxMac' in wx.PlatformInfo:
                # ... add a spacer to avoid control clipping
                r1Sizer.Add((2, 0))

            # Add the row sizer to the main vertical sizer
            mainSizer.Add(r1Sizer, 0, wx.EXPAND)

            # Add a vertical spacer to the main sizer        
            mainSizer.Add((0, 10))

        else:
            # Create a HORIZONTAL sizer for the next row
            r1Sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Create a VERTICAL sizer for the next element
            v3 = wx.BoxSizer(wx.VERTICAL)
            # Add the Import File element
            self.rtfname_edit = self.new_edit_box(_("Imported File"), v3, self.obj.imported_file)
            # Disable this field
            self.rtfname_edit.Enable(False)
            # Add the element to the row sizer
            r1Sizer.Add(v3, 4, wx.EXPAND)

            # Add a horizontal spacer to the row sizer        
            r1Sizer.Add((10, 0))

            # Create a VERTICAL sizer for the next element
            v35 = wx.BoxSizer(wx.VERTICAL)
            # If the import date is a datetime object ...
            if isinstance(self.obj.import_date, datetime.datetime):
                # ... get its YYYY-MM-DD representation
                dtStr = self.obj.import_date.strftime('%Y-%m-%d')
            # if the import date is None ...
            elif self.obj.import_date is None:
                # ... leave this field blank
                dtStr = ''
            # Otherwise ...
            else:
                # Use whatever the import date value is
                dtStr = self.obj.import_date
            # If no file was imported ...
            if self.obj.imported_file == '':
                # ... the date is the Document Creation Date
                prompt = _("Creation Date")
            # If a file was imported ...
            else:
                # ... the date is the Document Import Date
                prompt = _("Import Date")
            # Add the Import Date element
            self.import_date = self.new_edit_box(prompt, v35, dtStr)
            # Disable this field
            self.import_date.Enable(False)
            # Add the element to the row sizer
            r1Sizer.Add(v35, 1, wx.EXPAND)

            # Add the row sizer to the main vertical sizer
            mainSizer.Add(r1Sizer, 0, wx.EXPAND)

            # Add a vertical spacer to the main sizer        
            mainSizer.Add((0, 10))

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Add the Author element
        author_edit = self.new_edit_box(_("Author"), v4, self.obj.author, maxLen=100)
        # Add the element to the main sizer
        mainSizer.Add(v4, 1, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))
        
        # Create a VERTICAL sizer for the next element
        v5 = wx.BoxSizer(wx.VERTICAL)
        # Add the Comment element
        comment_edit = self.new_edit_box(_("Comment"), v5, self.obj.comment, maxLen=255)
        # Add the element to the main sizer
        mainSizer.Add(v5, 1, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

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
        r2Sizer.Add(v6, 1, wx.EXPAND)

        self.kw_list = []
        wx.EVT_LISTBOX(self, kw_groups_id, self.OnGroupSelect)

        # Add a horizontal spacer
        r2Sizer.Add((10, 0))

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
        r2Sizer.Add(v7, 1, wx.EXPAND)

        # Add a horizontal spacer
        r2Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v8 = wx.BoxSizer(wx.VERTICAL)
        # Keyword transfer buttons
        add_kw = wx.Button(self.panel, wx.ID_FILE2, ">>", wx.DefaultPosition)
        v8.Add(add_kw, 0, wx.EXPAND | wx.TOP, 20)
        wx.EVT_BUTTON(self, wx.ID_FILE2, self.OnAddKW)

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
        r2Sizer.Add(v8, 0)

        # Add a horizontal spacer
        r2Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v9 = wx.BoxSizer(wx.VERTICAL)

        # Episode Keywords [label]
        txt = wx.StaticText(self.panel, -1, _("Document Keywords"))
        v9.Add(txt, 0, wx.BOTTOM, 3)

        # Episode Keywords [list box]
        
        # Create an empty ListBox
        self.ekw_lb = wx.ListBox(self.panel, -1, wx.DefaultPosition, wx.DefaultSize, style=wx.LB_EXTENDED)
        v9.Add(self.ekw_lb, 1, wx.EXPAND)
        
        self.ekw_lb.Bind(wx.EVT_KEY_DOWN, self.OnKeywordKeyDown)

        # Add the element to the sizer
        r2Sizer.Add(v9, 2, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 5, wx.EXPAND)

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
        # Lay out the Form
        self.Layout()
        # Resize the form to fit the contents
        self.Fit()

        # Get the new size of the form
        (width, height) = self.GetSizeTuple()
        # Reset the form's size to be at least the specified minimum width
        self.SetSize(wx.Size(max(self.width, width), height))
        # Define the minimum size for this dialog as the current size, and define height as unchangeable
        self.SetSizeHints(max(self.width, width), height, -1, height)
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

        # Select the Library Default Keyword Group in the Keyword Group list
        if (library.keyword_group != '') and (self.kw_group_lb.FindString(library.keyword_group) != wx.NOT_FOUND):
            self.kw_group_lb.SetStringSelection(library.keyword_group)
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
        for documentKeyword in self.obj.keyword_list:
            self.ekw_lb.Append(documentKeyword.keywordPair)

        # Set focus to the Transcript ID
        self.id_edit.SetFocus()

    def refresh_keyword_groups(self):
        """Refresh the keyword groups listbox."""
        # Get the keyword groups from the database
        self.kw_groups = DBInterface.list_of_keyword_groups()
        # Clear the Keyword Groups list
        self.kw_group_lb.Clear()
        # Add the Keyword Groups to the Keyword Group list
        self.kw_group_lb.InsertItems(self.kw_groups, 0)
        # If there's at least one element in the Keyword Group list ...
        if len(self.kw_groups) > 0:
            # ... select the first element in the Keyword Groups list
            self.kw_group_lb.SetSelection(0)

    def refresh_keywords(self):
        """Refresh the keywords listbox."""
        # Get the current selection from the Keyword Groups list
        sel = self.kw_group_lb.GetStringSelection()
        # If there is a selection ...
        if sel:
            # ... get the Keywords for this Keyword Group
            self.kw_list = DBInterface.list_of_keywords_by_group(sel)
            # Clear the Keywords List
            self.kw_lb.Clear()
            if len(self.kw_list) > 0:
                # Add the Keywords to the Keyword list
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

    def OnBrowseClick(self, event):
        """ Method for when Browse button is clicked """
        # Get the default data directory
        dirName = TransanaGlobal.configData.videoPath
        # If we're using a Right-To-Left language ...
##        if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
            # ... we can only export to XML format
##            wildcard = _("Transcript Import Files (*.xml)|*.xml;|All Files (*.*)|*.*")
        # ... whereas with Left-to-Right languages
##        else:
            # ... we can import both RTF and XML formats
        wildcard = _("Transcript Import Formats (*.rtf, *.xml, *.txt)|*.rtf;*.xml;*.txt|Rich Text Format Files (*.rtf)|*.rtf|XML Files (*.xml)|*.xml|Text Files (*.txt)|*.txt|All Files (*.*)|*.*")
        # Allow for RTF, XML, TXT or *.* combinations
        dlg = wx.FileDialog(None, defaultDir=dirName, wildcard=wildcard, style=wx.OPEN)
        # Get a file selection from the user
        if dlg.ShowModal() == wx.ID_OK:
            # If the user clicks OK, set the file to import to the selected path.
            self.rtfname_edit.SetValue(dlg.GetPath())
            # If the ID field is blank ...
            if self.id_edit.GetValue() == '':
                # Get the base file name just selected
                tempFilename = os.path.basename(dlg.GetPath())
                # Split off the file extension
                (tempobjid, tempExt) = os.path.splitext(tempFilename)
                # Name the Transcript object after the imported Transcript
                self.id_edit.SetValue(tempobjid)
        # Destroy the File Dialog
        dlg.Destroy()

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
        # When a Keyword Group is selected, refresh the Keywords List to keywords for that Keyword Group
        self.refresh_keywords()
        
    def OnKeywordKeyDown(self, event):
        # Start Exception Handling
        try:
            # Get the Key Code
            c = event.GetKeyCode()
            # if the Delete Key was pressed ...
            if c == wx.WXK_DELETE:
                # ... and there is a Keyword selected in the Document Keywords list ...
                if self.ekw_lb.GetSelection() != wx.NOT_FOUND:
                    # ... delete that keyword
                    self.OnRemoveKW(event)
        # If an exception was raised ...
        except:
            # ... ignore it.
            pass

    def get_input(self):
        """Show the dialog and return the modified Document Object.  Result
        is None if user pressed the Cancel button."""
        # inherit parent method from Dialogs.Gen(eric)Form
        d = Dialogs.GenForm.get_input(self)
        # If the Form is created (not cancelled?) ...
        if d:
            # Set the Document ID
            self.obj.id = d[_('Document ID')]
            # Set the Author
            self.obj.author = d[_('Author')]
            # Set the Comment
            self.obj.comment = d[_('Comment')]
            # If this is a new Document ...
            if self.obj.number == 0:
                # Get the Media File to import
                fname = d[_('RTF/XML/TXT File to import  (optional)')]
                # Remember this as the imported file name
                self.obj.imported_file = fname
                # If a media file is entered ...
                if fname:
                    # ... start exception handling ...
                    try:
                        # Open the file
                        f = open(fname, "r")
                        # Read the file straight into the Transcript Text
                        self.obj.text = f.read()
                        # if the text does NOT have an RTF or XML header ...
                        if (self.obj.text[:5].lower() != '{\\rtf') and (self.obj.text[:5].lower() != '<?xml'):
                            # ... add "txt" to the start of the file to signal that it's probably a text file
                            self.obj.text = 'txt\n' + self.obj.text
                        # Close the file
                        f.close()
                    # If exceptions are raised ...
                    except:
                        # ... we don't need to do anything here.  (Error message??)
                        # The consequence is probably that the Document Text will be blank.
                        pass
##            else:
##                self.obj.imported_file = d[_('Imported File')]
                
        # If there's no input from the user ...
        else:
            # ... then we can set the Document Object to None to signal this.
            self.obj = None
        # Return the Document Object we've created / edited
        return self.obj


# This simple derrived class let's the user drop files onto an edit box
class EditBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, editbox):
        wx.FileDropTarget.__init__(self)
        self.editbox = editbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        self.editbox.SetValue(files[0])

        
class AddDocumentDialog(DocumentPropertiesForm):
    """Dialog used when adding a new Document."""

    def __init__(self, parent, id, library):
        obj = Document.Document()
        obj.library_num = library.number
        obj.author = DBInterface.get_username()
        DocumentPropertiesForm.__init__(self, parent, id, _("Add Document"), obj)


class EditDocumentDialog(DocumentPropertiesForm):
    """Dialog used when editing Document properties."""

    def __init__(self, parent, id, document_object):
        DocumentPropertiesForm.__init__(self, parent, id, _("Document Properties"), document_object)
