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

""" A Utility Program to deal with locked records for Transana-MU. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "RecordLock DEBUG is ON!"


# import wxPython
import wx

# import Python sys module
import sys
# import Transana's Dialogs
import Dialogs
# import Transana's Database Interface
import DBInterface
# import Transana's Global module
import TransanaGlobal

class RecordLock(wx.Dialog):
    """ This window displays the Record Lock Report form. """
    def __init__(self,parent,id,title):
        # Get the username from the TransanaGlobal module
        self.userName = TransanaGlobal.userName
        # Define the main Frame for the Chat Window
        wx.Dialog.__init__(self, parent, -1, title, size = (710,450), style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|wx.NO_FULL_REPAINT_ON_RESIZE)
        # Set the background to White
        self.SetBackgroundColour(wx.WHITE)
        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create a Sizer for the form
        box = wx.BoxSizer(wx.VERTICAL)

        # Create a sizer for the Buttons
        boxButtons = wx.BoxSizer(wx.HORIZONTAL)

        # Add an "Update Report" button
        self.btnUpdate = wx.Button(self, -1, _("Update Report"))
        # Add the Update button to the Buttons sizer
        boxButtons.Add(self.btnUpdate, 0, wx.LEFT, 6)
        # Bind the OnUpdate event to the Update button's press event
        self.btnUpdate.Bind(wx.EVT_BUTTON, self.OnUpdate)
        
        # Add a "Clear" button
        self.btnClear = wx.Button(self, -1, _("Clear"))
        # Add the Clear button to the Buttons sizer
        boxButtons.Add(self.btnClear, 0, wx.LEFT, 6)
        # Bind the OnClear event to the Clear button's press event
        self.btnClear.Bind(wx.EVT_BUTTON, self.OnClear)

        # Add a spacer
        boxButtons.Add((1, 1), 1, wx.EXPAND)

        # Add an "Automatic Updates" checkbox
        self.AutoUpdate = wx.CheckBox(self, -1, _("Update Automatically"))
        # Add the checkbox to the Button sizer
        boxButtons.Add(self.AutoUpdate, 0, wx.RIGHT, 6)
        # Bind the AutoUpdate Checkbox to an event
        self.AutoUpdate.Bind(wx.EVT_CHECKBOX, self.OnAutoUpdate)

        # Add a Timer for automatic updating
        self.AutoUpdateTimer = wx.Timer()
        # Bind the Timer to the Update event
        self.AutoUpdateTimer.Bind(wx.EVT_TIMER, self.OnUpdate)
        
        # Create a Sizer for Report data
        boxReport = wx.BoxSizer(wx.HORIZONTAL)

        # Create a Sizer for the Memo section
        boxMemo = wx.BoxSizer(wx.VERTICAL)
        # Create a label for the Memo section
        txtMemo = wx.StaticText(self, -1, _("Record Lock Information"))
        # Put the label in the Memo Sizer, with a little padding below
        boxMemo.Add(txtMemo, 0, wx.BOTTOM, 3)
        # Add a TextCtrl for the Report text.  This is read only, as it is filled programmatically.
        self.memo = wx.TextCtrl(self, -1, style = wx.TE_MULTILINE | wx.TE_WORDWRAP | wx.TE_READONLY)
        # Put the Memo control in the Memo Sizer
        boxMemo.Add(self.memo, 1, wx.EXPAND)
        # Add the Memo display to the Report Sizer
        boxReport.Add(boxMemo, 6, wx.EXPAND)

        # Create a Sizer for the User section
        boxUser = wx.BoxSizer(wx.VERTICAL)
        # Create a label for the User section
        txtUser = wx.StaticText(self, -1, _("Current Users"))
        # Put the label in the User Sizer with a little padding below
        boxUser.Add(txtUser, 0, wx.BOTTOM, 3)

        # Add a ListBox to hold the names of active users
        self.userList = wx.ListBox(self, -1, choices=[], style=wx.LB_SINGLE)
        # Add the Current User List to the User Sizer
        boxUser.Add(self.userList, 1, wx.EXPAND)

        # Create a label for the User section
        txtUser = wx.StaticText(self, -1, _("Users who have records\nthat can be unlocked"))
        # Put the label in the User Sizer with a little padding above and below.  (Also needs Right on Mac)
        boxUser.Add(txtUser, 0, wx.TOP | wx.BOTTOM | wx.RIGHT, 3)

        # Add a ListBox to hold the names of users with locks that can be removed
        self.usersWithLocksThatCanBeCleared = wx.ListBox(self, -1, choices=[], style=wx.LB_SINGLE)
        # Add the Locks that can be Cleared list to the User Sizer
        boxUser.Add(self.usersWithLocksThatCanBeCleared, 1, wx.EXPAND)

        # Add an "Unlock Records" button
        self.btnUnlock = wx.Button(self, -1, _("Unlock Records"))
        # Disable the Unlock button initially
        self.btnUnlock.Enable(False)
        # Add the Unlock button to the User Sizer
        boxUser.Add(self.btnUnlock, 0, wx.EXPAND | wx.BOTTOM, 3)
        # Bind the OnUnlock event to the Unlock button's press event
        self.btnUnlock.Bind(wx.EVT_BUTTON, self.OnUnlock)
        
        # Add the User List Sizer to the Report Sizer
        boxReport.Add(boxUser, 2, wx.EXPAND | wx.LEFT | wx.RIGHT, 6)

        # Add the Buttons Sizer to the Form Sizer
        box.Add(boxButtons, 1, wx.EXPAND | wx.ALL, 4)
        # Put the Report Sizer in the form sizer
        box.Add(boxReport, 13, wx.EXPAND | wx.ALL, 4)

        # Attach the Form's Main Sizer to the form
        self.SetSizer(box)
        # Set AutoLayout on
        self.SetAutoLayout(True)
        # Lay out the form
        self.Layout()
        # Set the minimum size for the form.
        self.SetSizeHints(minW = 600, minH = 440)
        # Center the form on the screen
        self.CentreOnScreen()
        # Populate the report
        self.OnUpdate(None)

    def OnUpdate(self, event):
        """ Update Report handler """
        # rebuild the list of current users.  Start by clearing it.
        self.userList.Clear()
        # We need to count the number of times THIS USER's UserName is in the list.  Initialize a counter.
        userNameCount = 0
        # initialize the list of current users
        currentUsers = []
        # If there's a Chat Window defined, we'll get the current users from there.
        if TransanaGlobal.chatWindow != None:
            # Iterate through the items in the ChatWindow's User List control
            for n in range(TransanaGlobal.chatWindow.userList.GetCount()):
                # Assign the userName to a variable for easier manipulation
                userName = TransanaGlobal.chatWindow.userList.GetString(n)
                # Add the UserName to the currentUsers list
                currentUsers.append(userName)
                # Add the UserName to the RLR's Current Users control
                self.userList.Append(userName)
                # Strip off the (#) extension added to multiple copies of the same username.
                # We have to do this in case DKW has dropped off leaving only DKW(2) on line, and
                # also to allow us to differentiate Alex from Alexandra, which my first try couldn't.
                #
                # Start by looking in the last character position for a closing paren and confirming
                # that there's an open paren.  If this isn't good enough, we may have to mess with
                # regular expressions here, but I'd rather not.
                if (userName[-1] == ')') and (userName.find('(') > -1):
                    # Strip the username, starting with the right-most open paren
                    userName = userName[:userName.rfind('(')]
                # If the username matches the current user ...
                if self.userName == userName:
                    # ... count the number of occurrences of the current user's name
                    # (to check for duplicates, which is possible.)
                    userNameCount += 1
        # Clear the old Report data
        self.memo.Clear()
        # Get the new Record Lock Text and Record Lock Users List from the Database
        (rlText, rlUsers) = DBInterface.ReportRecordLocks(self)
        # Put the Text in the Recort Lock Report Memo
        self.memo.AppendText(rlText)
        
        # Now that we know who the current users are and who the users with locks are,
        # we can determine if there are any locks that can be cleared.
        # Start by clearing the list of users with locks that can be cleared.
        self.usersWithLocksThatCanBeCleared.Clear()
        # Determine if the user currently has a Transcript in Edit Mode.
        inReadOnlyMode = TransanaGlobal.menuWindow.ControlObject.TranscriptWindow[TransanaGlobal.menuWindow.ControlObject.activeTranscript].dlg.editor.get_read_only()
        # If the current user has records locked and if there is only one instance of the current user 
        # in the UserName list, then allow the user to unlock his/her own records.  If there are multiple
        # instances of the current user's account in use, we better not clear ANY locks for this user.
        # If the user is not in Read Only mode for the Transcript, we'd better not clear the locks either.
        if (self.userName in rlUsers) and (userNameCount == 1) and inReadOnlyMode:
            self.usersWithLocksThatCanBeCleared.Append(self.userName)
        # Add users with locks who are NOT logged on to the list of users whose locks can be cleared.
        # Iterate through the list of users with locks...
        for user in rlUsers:
            # If the user is not in the list of current users ...
            # (... also, we've already added the current user if eligible, so don't add it again ...)
            if not (user in currentUsers) and (user != self.userName):
                # ... then add the name to the list of users whose locks can be cleared.
                self.usersWithLocksThatCanBeCleared.Append(user)
        # If there are any names in the list of users whose locks can be cleared ...
        if self.usersWithLocksThatCanBeCleared.GetCount() > 0:
            # ... select the first name in the list to ensure that there's always a selection ...
            self.usersWithLocksThatCanBeCleared.SetSelection(0)
            # ... and enable the "Unlock" button.
            self.btnUnlock.Enable(True)
        # If there are NO users in the list, which means there's nothing to unlock ...
        else:
            # ... we'd better disable the "Unlock" button.
            self.btnUnlock.Enable(False)
        
    def OnAutoUpdate(self, event):
        """ Handle Auto Update Checkbox """
        # If the checkbox is checked ...
        if self.AutoUpdate.IsChecked():
            # Call the Update event immediately
            self.OnUpdate(event)
            # Start the Timer so the Update will be called every 2.5 seconds
            self.AutoUpdateTimer.Start(2500)
        # if the checkbox is un-checked ...
        else:
            # ... stop the timer
            self.AutoUpdateTimer.Stop()

    def OnUnlock(self, event):
        """ Unlock handler """
        # In general, Transana Record Locking employs a pessimistic locking strategy.
        # However, since the Record Lock Utility is used so rarely, and because the probability 
        # is so small that a user with a dead Lock will log on while another user is in the process 
        # of trying to unlock that record, I've employed an optimistic strategy here.  The Report
        # assumes things won't change, and doesn't look for or respond to change.  But when the
        # unlock is requested, we'd better check to make sure that the selected user is STILL inactive.
        #
        # We need to count the number of times THIS USER's UserName is in the list.  Initialize a counter.
        userNameCount = 0
        # As long as a Chat Window is defined, which it always should be, we'll use it to get
        # info on what users are currently connected.
        if TransanaGlobal.chatWindow != None:
            # Initialize a list of current users.
            currentUsers = []
            # Iterate through the list of users registered in the Chat window.
            for n in range(TransanaGlobal.chatWindow.userList.GetCount()):
                # Capture the user name to a variable for manipulation
                userName = TransanaGlobal.chatWindow.userList.GetString(n)
                # Add the user to the Current User list
                currentUsers.append(userName)
                # Strip off the (#) extension added to multiple copies of the same username.
                # We have to do this in case DKW has dropped off leaving only DKW(2) on line, and
                # also to allow us to differentiate Alex from Alexandra, which my first try couldn't.
                #
                # Start by looking in the last character position for a closing paren and confirming
                # that there's an open paren.  If this isn't good enough, we may have to mess with
                # regular expressions here, but I'd rather not.
                if (userName[-1] == ')') and (userName.find('(') > -1):
                    # Strip the username, starting with the right-most open paren
                    userName = userName[:userName.rfind('(')]
                # If the username matches the current user ...
                if self.userName == userName:
                    # ... count the number of occurrences of the current user's name
                    # (to check for duplicates, which is possible.)
                    userNameCount += 1

        # If the current user selected his/her own username, unlock can only proceed if this
        # is the only instance of this user being connected.  If the current user selected
        # another user name, confirm that the selected name is not in the refreshed list of
        # current users.
        if ((self.usersWithLocksThatCanBeCleared.GetStringSelection() == self.userName) and \
            (userNameCount == 1)) or \
           (not self.usersWithLocksThatCanBeCleared.GetStringSelection() in currentUsers):
            # Only then can we proceed with the record unlock request
            DBInterface.UnlockRecords(self, self.usersWithLocksThatCanBeCleared.GetStringSelection())
        # If the requested user HAS logged back in, display an error message.
        else:
            # Create an error message
            msg = _('Records locked by %s cannot be unlocked at this time because of a change in current users.')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                msg = unicode(msg, 'utf8')
            # Create an Error Dialog
            dlg = Dialogs.ErrorDialog(self, msg % self.usersWithLocksThatCanBeCleared.GetStringSelection())
            # Display the Error Message
            dlg.ShowModal()
            # Destroy the Dialog
            dlg.Destroy()
        # Update the Record Lock Report data
        self.OnUpdate(event)

    def OnClear(self, event):
        """ Clear button handler """
        # If the AutoUpdate Timer is running, ...
        if self.AutoUpdateTimer.IsRunning():
            # ... STOP it!
            self.AutoUpdateTimer.Stop()
            # Also, uncheck Auto Update!
            self.AutoUpdate.SetValue(False)
        # Clear the list of Current Users
        self.userList.Clear()
        # If there's a Chat Window defined, we'll get the current users from there.
        if TransanaGlobal.chatWindow != None:
            # Iterate through the items in the ChatWindow's User List control
            for n in range(TransanaGlobal.chatWindow.userList.GetCount()):
                # Add the UserName to the RLR's Current Users control
                self.userList.Append(TransanaGlobal.chatWindow.userList.GetString(n))
        # Clear the previous Record Lock Report text
        self.memo.Clear()
        # Clear the list of users whose locks can be cleared
        self.usersWithLocksThatCanBeCleared.Clear()
        # Disable the Unlock button
        self.btnUnlock.Enable(False)
