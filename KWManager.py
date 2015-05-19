# Copyright (C) 2003 - 2014 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Keyword Management interface."""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

import wx                           # import wxPython
import sys                          # import Python's sys module
from TransanaExceptions import *    # import Transana's exceptions
import KeywordObject as Keyword     # import Transana's Keyword Object definition
import KeywordPropertiesForm        # import Transana's Keyword Properties form
import DBInterface                  # import Transana's Database Interface
import Dialogs                      # import Transana's Dialog boxes
import TransanaConstants            # import Transana's Constants module
import TransanaGlobal               # import Transana's Global module

class KWManager(wx.Dialog):
    """User interface for managing keywords and keyword groups."""

    def __init__(self, parent, defaultKWGroup=None, deleteEnabled=True):
        """Initialize a KWManager object."""
        # Remember the value for deleteEnabled
        self.deleteEnabled = deleteEnabled
        wx.Dialog.__init__(self, parent, -1, _("Keyword Management"), wx.DefaultPosition, wx.Size(550,420), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Define the minimum size for this dialog as the initial size
        self.SetSizeHints(550, 420)

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Keyword Group
        txt = wx.StaticText(self, -1, _("Keyword Group"))
        mainSizer.Add(txt, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
        mainSizer.Add((0, 3))

        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.kw_group = wx.Choice(self, 101, wx.DefaultPosition, wx.DefaultSize, [])
        r1Sizer.Add(self.kw_group, 5, wx.EXPAND)
        wx.EVT_CHOICE(self, 101, self.OnGroupSelect)

        r1Sizer.Add((10, 0))

        # Create a new Keyword Group button
        new_kwg = wx.Button(self, wx.ID_FILE1, _("Create a New Keyword Group"))
        r1Sizer.Add(new_kwg, 3, wx.EXPAND)
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnNewKWG)

        mainSizer.Add(r1Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

        mainSizer.Add((0, 10))

        # Keywords label+listbox
        txt = wx.StaticText(self, -1, _("Keywords"))
        mainSizer.Add(txt, 0, wx.LEFT, 10)

        # Create a HORIZONTAL sizer for the first row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        v1 = wx.BoxSizer(wx.VERTICAL)
        if 'wxMac' in wx.PlatformInfo:
            style=wx.LB_SINGLE
        else:
            style=wx.LB_SINGLE | wx.LB_SORT
        self.kw_lb = wx.ListBox(self, 100, wx.DefaultPosition, wx.DefaultSize, [], style=style)
        v1.Add(self.kw_lb, 1, wx.EXPAND)
        wx.EVT_LISTBOX(self, 100, self.OnKeywordSelect)
        wx.EVT_LISTBOX_DCLICK(self, 100, self.OnKeywordDoubleClick)

        r2Sizer.Add(v1, 5, wx.EXPAND)
        r2Sizer.Add((10, 0))

        v2 = wx.BoxSizer(wx.VERTICAL)
        # Add Keyword to List button
        add_kw = wx.Button(self, wx.ID_FILE2, _("Add Keyword to List"))
        v2.Add(add_kw, 0, wx.EXPAND | wx.BOTTOM, 10)
        wx.EVT_BUTTON(self, wx.ID_FILE2, self.OnAddKW)

        # Edit Keyword button
        self.edit_kw = wx.Button(self, -1, _("Edit Keyword"))
        v2.Add(self.edit_kw, 0, wx.EXPAND | wx.BOTTOM, 10)
        wx.EVT_BUTTON(self, self.edit_kw.GetId(), self.OnEditKW)
        self.edit_kw.Enable(False)
        
        # Delete Keyword from List button
        self.del_kw = wx.Button(self, wx.ID_FILE3, _("Delete Keyword from List"))
        v2.Add(self.del_kw, 0, wx.EXPAND | wx.BOTTOM, 10)
        wx.EVT_BUTTON(self, wx.ID_FILE3, self.OnDelKW)
        self.del_kw.Enable(False)
        

        # Definition box
        def_txt = wx.StaticText(self, -1, _("Definition"))
        v2.Add(def_txt, 0, wx.BOTTOM, 3)
        
        self.definition = wx.TextCtrl(self, -1, '', style=wx.TE_MULTILINE)
        v2.Add(self.definition, 1, wx.EXPAND | wx.BOTTOM, 10)
        self.definition.Enable(False)

        btnSizer = wx.BoxSizer(wx.HORIZONTAL)        
        # Dialog Close button
        close = wx.Button(self, wx.ID_CLOSE, _("Close"))
        btnSizer.Add(close, 1, wx.EXPAND | wx.RIGHT, 10)
        close.SetDefault()
        wx.EVT_BUTTON(self, wx.ID_CLOSE, self.OnClose)

        # We don't want to use wx.ID_HELP here, as that causes the Help buttons to be replaced with little question
        # mark buttons on the Mac, which don't look good.
        ID_HELP = wx.NewId()

        # Dialog Help button
        helpBtn = wx.Button(self, ID_HELP, _("Help"))
        btnSizer.Add(helpBtn, 1, wx.EXPAND)
        wx.EVT_BUTTON(self, ID_HELP, self.OnHelp)

        v2.Add(btnSizer, 0, wx.EXPAND)

        r2Sizer.Add(v2, 3, wx.EXPAND)

        mainSizer.Add(r2Sizer, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()
        self.CenterOnScreen()

        self.kw_groups = DBInterface.list_of_keyword_groups()
        for kwg in self.kw_groups:
            self.kw_group.Append(kwg)

        if len(self.kw_groups) > 0:
            if defaultKWGroup != None:
                selPos = self.kw_group.FindString(defaultKWGroup)
            else:
                selPos = 0
            self.kw_group.SetSelection(selPos)
            self.kw_list = DBInterface.list_of_keywords_by_group(self.kw_group.GetString(selPos))
        else:
            self.kw_list = []

        for kw in self.kw_list:
            self.kw_lb.Append(kw)
            
        if len(self.kw_list) > 0:
            self.kw_lb.SetSelection(0, False)

        self.ShowModal()

    def refresh_keywords(self):
        """Refresh the keywords listbox."""
        sel = self.kw_group.GetStringSelection()
        if sel and (len(sel) > 0):
            self.kw_list = DBInterface.list_of_keywords_by_group(sel)
            self.kw_lb.Clear()
            if len(self.kw_list) > 0:
                self.kw_lb.AppendItems(self.kw_list)
        # Since there is no keyword selected following this operation, we need to
        # disable the Delete button and clear the Definition field.
        self.edit_kw.Enable(False)
        self.del_kw.Enable(False)
        self.definition.SetValue('')

    def OnClose(self, evt):
        """Invoked when dialog Close button is activated."""
        self.EndModal(0)

    def OnHelp(self, evt):
        """Invoked when dialog Help button is activated."""
        if TransanaGlobal.menuWindow != None:
            TransanaGlobal.menuWindow.ControlObject.Help('Keyword Management')

    def OnNewKWG(self, evt):
        """Invoked when the 'Create a new keyword group' button is activated."""
        kwg = Dialogs.add_kw_group_ui(self, self.kw_groups)
        if kwg:
            self.kw_groups.append(kwg)
            self.kw_group.Append(kwg)
            self.kw_group.SetStringSelection(kwg)
            self.OnGroupSelect(evt)

    def OnAddKW(self, evt):
        """Invoked when the 'Add Keyword to list' button is activated."""
        # Create the Keyword Properties Dialog Box to Add a Keyword
        dlg = KeywordPropertiesForm.AddKeywordDialog(self, -1, self.kw_group.GetStringSelection())
        # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
        contin = True
        # While the "continue" flag is True ...
        while contin:
            # Get a new Keyword object with properties as given
            kw = dlg.get_input()
            # If the user presses "OK" ...
            if kw != None:
                # Be ready to catch exceptions
                try:
                    # Try to save the keyword
                    kw.db_save()
                    # If the Keyword Group was not changed...
                    if kw.keywordGroup == self.kw_group.GetStringSelection():
                        # ...Add the new Keyword to the Keyword List
                        self.kw_lb.Append(kw.keyword)
                    # If the Keyword Group WAS changed ...
                    else:
                        # ... if the Keyword Group is NOT already in the Keyword Group List ...
                        if self.kw_group.FindString(kw.keywordGroup) == -1:
                            # ... add it to the list.
                            self.kw_group.Append(kw.keywordGroup)
                         # ... and select it.
                        self.kw_group.SetStringSelection(kw.keywordGroup)
                        self.refresh_keywords()
                    if not TransanaConstants.singleUserVersion:
                        if TransanaGlobal.chatWindow != None:
                            msgData = "%s >|< %s" % (kw.keywordGroup, kw.keyword)
                            TransanaGlobal.chatWindow.SendMessage("AK %s" % msgData)
                    # If we do all this, we don't need to continue any more.
                    contin = False
                # Handle "SaveError" exception
                except SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                # Handle other exceptions
                except:
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    # Display the Exception Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, "%s\n%s" % (sys.exc_info()[0], sys.exc_info()[1]))
                    errordlg.ShowModal()
                    errordlg.Destroy()
            # If the user pressed Cancel ...
            else:
                # ... then we don't need to continue any more.
                contin = False
                    

    def OnEditKW(self, evt):
        """Invoked when the 'Edit Keyword' button is activated."""
        kw_name = self.kw_lb.GetStringSelection()
        if kw_name == "":
            return
        # Load the selected keyword into a Keyword Object
        kw = Keyword.Keyword(self.kw_group.GetStringSelection(), kw_name)
        self.EditKeyword(kw, evt)

    def OnDelKW(self, evt):
        """Invoked when the 'Delete Keyword from list' button is activated."""
        kw_group = self.kw_group.GetStringSelection()
        kw_name = self.kw_lb.GetStringSelection()
        if kw_name == "":
            return
        sel = self.kw_lb.GetSelection()
        # If we found the keyword we're supposed to delete ...
        if TransanaGlobal.menuWindow.ControlObject.UpdateSnapshotWindows('Detect', evt, kw_group, kw_name):
            # ... present an error message to the user
            msg = _("Keyword %s : %s is contained in a Snapshot you are currently editing.\nYou cannot delete it at this time.")
            msg = unicode(msg, 'utf8')
            dlg = Dialogs.ErrorDialog(self, msg % (kw_group, kw_name))
            dlg.ShowModal()
            dlg.Destroy()
            # Signal that we do NOT want to delete the Keyword!
            result = wx.ID_NO
        else:
            msg = _('Are you sure you want to delete Keyword "%s" and all instances of it from the Clips and Snapshots?')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                msg = unicode(msg, 'utf8')
            dlg = Dialogs.QuestionDialog(self, msg % kw_name)
            result = dlg.LocalShowModal()
            dlg.Destroy()
        if result == wx.ID_YES:
            try:
                DBInterface.delete_keyword(self.kw_group.GetStringSelection(), kw_name)
                self.kw_lb.Delete(sel)
                self.definition.SetValue('')

                # We need to update all open Snapshot Windows based on the change in this Keyword
                TransanaGlobal.menuWindow.ControlObject.UpdateSnapshotWindows('Update', evt, kw_group, kw_name)

                if not TransanaConstants.singleUserVersion:
                    if TransanaGlobal.chatWindow != None:
                        # We need the UNTRANSLATED Root Node here
                        msgData = "%s >|< %s >|< %s >|< %s" % ('KeywordNode', 'Keywords', self.kw_group.GetStringSelection(), kw_name)
                        TransanaGlobal.chatWindow.SendMessage("DN %s" % msgData)
                        # We need to update the Keyword Visualization no matter what here, when deleting a keyword
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))
            except GeneralError, e:
                # Display the Exception Message, allow "continue" flag to remain true
                prompt = "%s"
                data = e.explanation
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % data)
                errordlg.ShowModal()
                errordlg.Destroy()

    def OnKeywordSelect(self, evt):
        """Invoked when a keyword is selected in the listbox."""
        # Check to see if there IS a keyword selected.  (an error was being raised!)
        if self.kw_lb.GetStringSelection() != '':
            kw = Keyword.Keyword(self.kw_group.GetStringSelection(), self.kw_lb.GetStringSelection())
            self.definition.SetValue(kw.definition)
            self.edit_kw.Enable(self.deleteEnabled)
            self.del_kw.Enable(self.deleteEnabled)

    def OnKeywordDoubleClick(self, event):
        """Double-clicking a keyword calls the Edit Properties screen!"""
        # Load the selected keyword into a Keyword Object
        kw = Keyword.Keyword(self.kw_group.GetStringSelection(), self.kw_lb.GetStringSelection())
        # Double-clicking should only work if Editing is enabled!!
        if self.edit_kw.IsEnabled():
            self.EditKeyword(kw, event)

    def EditKeyword(self, kw, evt):
        # We need to update all open Snapshot Windows based on the change in this Keyword
        # If we found the keyword we're supposed to delete ...
        if TransanaGlobal.menuWindow.ControlObject.UpdateSnapshotWindows('Detect', evt, kw.keywordGroup, kw.keyword):
            # ... present an error message to the user
            msg = _('Keyword "%s : %s" is contained in a Snapshot you are currently editing.\nYou cannot edit it at this time.')
            msg = unicode(msg, 'utf8')
            dlg = Dialogs.ErrorDialog(self, msg % (kw.keywordGroup, kw.keyword))
            dlg.ShowModal()
            dlg.Destroy()
        else:
            # use "try", as exceptions could occur
            try:
                # Try to get a Record Lock
                kw.lock_record()
            # Handle the exception if the record is locked
            except RecordLockedError, e:
                ReportRecordLockedException(_("Keyword"), kw.keywordGroup + ' : ' + kw.keyword, e)
            # If the record is not locked, keep going.
            else:
                if self.deleteEnabled:
                    # Create the Keyword Properties Dialog Box to edit the Keyword Properties
                    dlg = KeywordPropertiesForm.EditKeywordDialog(self, -1, kw)
                    # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
                    contin = True
                    # While the "continue" flag is True ...
                    while contin:
                        # Display the Keyword Properties Dialog Box and get the data from the user
                        if dlg.get_input() != None:
                            # if the user pressed "OK" ...
                            try:
                                # Try to save the Keyword Data
                                result = kw.db_save()
                                originalKeyword = self.kw_lb.GetStringSelection()
                                originalKeywordGroup = self.kw_group.GetStringSelection()
                                # See if the Keyword Group has been changed.  If it has, update the form.
                                if kw.keywordGroup != self.kw_group.GetStringSelection():
                                    # See if the new Keyword Group exists, and if not, create it
                                    if self.kw_group.FindString(kw.keywordGroup) == -1:
                                        self.kw_group.Append(kw.keywordGroup)
                                    # Remove the keyword from the current list
                                    self.kw_lb.Delete(self.kw_lb.GetSelection())
                                    if not TransanaConstants.singleUserVersion:
                                        if TransanaGlobal.chatWindow != None:
                                            # We need the UNTRANSLATED Root Node here
                                            msgData = "%s >|< %s >|< %s >|< %s" % ('KeywordNode', 'Keywords', originalKeywordGroup, originalKeyword)
                                            TransanaGlobal.chatWindow.SendMessage("DN %s" % msgData)
                                            msgData = "%s >|< %s" % (kw.keywordGroup, kw.keyword)
                                            TransanaGlobal.chatWindow.SendMessage("AK %s" % msgData)
                                    # If we've changed KW Groups, we need to disable the buttons.
                                    self.edit_kw.Enable(False)
                                    self.del_kw.Enable(False)
                                    # Clear the Definition field
                                    self.definition.SetValue('')
                                else:
                                    # If the Keyword has been changed, update it on the form.
                                    if kw.keyword != originalKeyword:
                                        self.kw_lb.SetString(self.kw_lb.GetSelection(), kw.keyword)
                                    # Update the Definition on the Form
                                    self.definition.SetValue(kw.definition)
                                    if not TransanaConstants.singleUserVersion:
                                        # We only rename the node if the chat window exists AND IF WE'RE NOT MERGING KEYWORDS
                                        if (TransanaGlobal.chatWindow != None) and result:
                                            # We need the UNTRANSLATED Root Node here
                                            msgData = "%s >|< %s >|< %s >|< %s >|< %s" % ('KeywordNode', 'Keywords', kw.keywordGroup, originalKeyword, kw.keyword)
                                            TransanaGlobal.chatWindow.SendMessage("RN %s" % msgData)
                                # if result if False, we have merged keywords!
                                if not result:
                                    # First remove the keyword from the list.  If the keyword was merged to a different keyword group,
                                    # there won't be one to delete, though.
                                    if self.kw_lb.GetSelection() > -1:
                                        self.kw_lb.Delete(self.kw_lb.GetSelection())
                                        # Then, if MU, send the message to others to remove the original keyword from the tree
                                        if not TransanaConstants.singleUserVersion:
                                            if TransanaGlobal.chatWindow != None:
                                                msgData = "KeywordNode >|< Keywords >|< %s >|< %s" % (originalKeywordGroup, originalKeyword)
                                                TransanaGlobal.chatWindow.SendMessage("DN %s" % msgData)

                                # We need to update all open Snapshot Windows based on the change in this Keyword
                                TransanaGlobal.menuWindow.ControlObject.UpdateSnapshotWindows('Update', evt, originalKeywordGroup, originalKeyword)

                                # This computer updates the keyword visualization later, but other computers might need to update it now.
                                if not TransanaConstants.singleUserVersion:
                                    # We need to update the Keyword Visualization no matter what here, when deleting a keyword group
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))
                                # If we do all this, we don't need to continue any more.
                                contin = False
                            # Handle "SaveError" exception
                            except SaveError:
                                # Display the Error Message, allow "continue" flag to remain true
                                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                                errordlg.ShowModal()
                                errordlg.Destroy()
                            # Handle other exceptions
                            except:
                                import traceback
                                traceback.print_exc(file=sys.stdout)
                                # Display the Exception Message, allow "continue" flag to remain true
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_("Exception %s: %s"), 'utf8')
                                else:
                                    prompt = _("Exception %s: %s")
                                errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                                errordlg.ShowModal()
                                errordlg.Destroy()
                        # If the user pressed Cancel ...
                        else:
                            # ... then we don't need to continue any more.
                            contin = False
            # Unlock the record regardless of what happens
            kw.unlock_record()


    def OnGroupSelect(self, evt):
        """Invoked when a keyword group is selected."""
        self.refresh_keywords()
