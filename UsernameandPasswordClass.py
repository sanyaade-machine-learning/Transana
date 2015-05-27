#Copyright (C) 2003 - 2015  The Board of Regents of the University of Wisconsin System
#
#This program is free software; you can redistribute it and/or
#modify it under the terms of the GNU General Public License
#as published by the Free Software Foundation; either version 2
#of the License, or (at your option) any later version.
#
#This program is distributed in the hope that it will be useful,
#but WITHOUT ANY WARRANTY; without even the implied warranty of
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#GNU General Public License for more details.
#
#You should have received a copy of the GNU General Public License
#along with this program; if not, write to the Free Software
#Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

""" For single-user Transana, this module requests the Database Name from the user.
    For multi-user Transana, this module requests UserName, Password, Database Server,
    and Database Name from the user. """

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx

if __name__ == '__main__':
    __builtins__._ = wx.GetTranslation

# Import Python's os and sys modules
import os, sys
# import Transana's Database Interface
import DBInterface
# import Transana's Dialogs
import Dialogs
# import Transana's Constants
import TransanaConstants
# import Transana's Global Variables
import TransanaGlobal

class UsernameandPassword(wx.Dialog):
    """ Username, Password, Database Server, and Database Name Dialog """
    def __init__(self, parent):
        """ This is the Username, Password, Database Server, and Database Name Dialog """
        # For single-user Transana, all we need it the Database Name.  For multi-user Transana,
        # we need UserName, Password, Database Server, and Database Name.

        # First, let's set up some variables depending on whether we are in single-user or multi-user
        # mode.
        if TransanaConstants.singleUserVersion:
            # Dialog Title
            dlgTitle = _("Select Database")
            # Dialog Size
            dlgSize=(350, 40)
            # Instructions Text
            instructions = _("Please enter the name of the database you wish to use.\nTo create a new database, type in a new database name.\n(Database names may contain only letters and numbers in a single word.)")
            if TransanaConstants.demoVersion:
                instructions = _("Database selection has been disabled for the Demonstration Version.\n\nIn the full version of Transana, you can create as many separate\ndatabases as you want.\n")
            # Sizer Proportion for the instructions
            instProportion = 4
        else:
            # Dialog Title
            dlgTitle = _("Username and Password")
            # Dialog Size
            dlgSize=(450, 320)  # (350, 310)
            # Instructions Text
            instructions = _("Please enter your MySQL username and password, as well \nas the names of the server and database you wish to use.\nTo create a new database, type in a new database name\n(if you have appropriate permissions.)\n(Database names may contain only letters and numbers in\na single word.)")
            # Macs don't need as much space for instructions as Windows machines do.
            if 'wxMac' in wx.PlatformInfo:
                # Sizer Proportion for the instructions
                instProportion = 4
            else:
                # Sizer Proportion for the instructions
                instProportion = 6

        # Define the main Dialog Box
        wx.Dialog.__init__(self, parent, -1, dlgTitle, size=dlgSize, style=wx.CAPTION | wx.RESIZE_BORDER)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create a BoxSizer for the main elements to go onto the Panel
        userPanelSizer = wx.BoxSizer(wx.VERTICAL)

        # If we are in the Multi-user Version ...
        if not TransanaConstants.singleUserVersion:
            notebook = wx.Notebook(self, -1, size=self.GetSizeTuple())
            panelParent = notebook
            # Adding this line prevents an odd visual distortion in Arabic!
            notebook.SetBackgroundColour(wx.WHITE)
        else:
            panelParent = self

        # Place a Panel on the Dialog Box.  All Controls will go on the Panel.
        # (Panels can have DefaultItems, enabling the desired "ENTER" key functionality.
        userPanel = wx.Panel(panelParent, -1, name='UserNamePanel')

        # Instructions Text
        lblIntro = wx.StaticText(userPanel, -1, instructions)
        # Add the Instructions to the Main Sizer
#        userPanelSizer.Add(lblIntro, instProportion, wx.EXPAND | wx.ALL, 10)
        userPanelSizer.Add(lblIntro, 0, wx.ALL, 10) # instProportion, wx.EXPAND | wx.ALL, 10)

        # Get the dictionary of defined database hosts and databases from the Configuration module
        self.Databases = TransanaGlobal.configData.databaseList

        # With release 2.30, we add the Port parameter for MU.  This requires a change to the format of the
        # configuration data.  Therefore, we need to examine the structure of the configuration data.
        # If it exists and the data type is a list, we need to update the structure of the existing data.
        if (len(self.Databases) >0) and (type(self.Databases[self.Databases.keys()[0]]) == type([])):
            # Create a new Dictionary object
            newDatabases = {}
            # Iterate through the old dictionary object.  For each database server ...
            for dbServer in self.Databases.keys():
                # ... we need the database list AND the port.  Port 3306 was the only option until this change
                #, so makes an adequate default.
                newDatabases[dbServer] = {'dbList' : self.Databases[dbServer],
                                          'port' : '3306'}
            # Replace the old config data with this modified version.
            self.Databases = newDatabases

        # The multi-user version has more fields than the single-user version.
        # If not the single-user version, put these fields on the form.
        if not TransanaConstants.singleUserVersion:

            # Let's use a FlexGridSizer for the data entry fields.
            # for MU, we want 5 rows with 2 columns
#            box2 = wx.FlexGridSizer(5, 2, 6, 0)
            # We want to be flexible horizontally
#            box2.SetFlexibleDirection(wx.HORIZONTAL)
            # We want the data entry fields to expand
#            box2.AddGrowableCol(1)

            # The proportion for the data entry portion of the screen should be 6
#            box2Proportion = 0

            # Use a BoxSizer instead of a FlexGridSizer.  The FlexGridSizer isn't handling alternate font sizes
            # on Windows correctly.
#            box2 = wx.BoxSizer(wx.VERTICAL)

            # Create a Row Sizer for Username
            r1Sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Username Label        
            lblUsername = wx.StaticText(userPanel, -1, _("Username:"))
            r1Sizer.Add(lblUsername, 1, wx.ALL, 5)

            # User Name TextCtrl
            self.txtUsername = wx.TextCtrl(userPanel, -1, style=wx.TE_LEFT)
            if DBInterface.get_username() != '':
                self.txtUsername.SetValue(DBInterface.get_username())
            r1Sizer.Add(self.txtUsername, 4, wx.EXPAND | wx.ALL, 5)
#            box2.Add(r1Sizer, 0, wx.EXPAND)
            userPanelSizer.Add(r1Sizer, 0, wx.EXPAND)

            # Create a Row Sizer for Password
            r2Sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Password Label
            lblPassword = wx.StaticText(userPanel, -1, _("Password:"))
            r2Sizer.Add(lblPassword, 1, wx.ALL, 5)

            # Password TextCtrl (with PASSWORD style)
            self.txtPassword = wx.TextCtrl(userPanel, -1, style=wx.TE_LEFT|wx.TE_PASSWORD)
            r2Sizer.Add(self.txtPassword, 4, wx.EXPAND | wx.ALL, 5)
#            box2.Add(r2Sizer, 0, wx.EXPAND)
            userPanelSizer.Add(r2Sizer, 0, wx.EXPAND)

            # Create a Row Sizer for Host / Server
            r3Sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Host / Server Label
            lblDBServer = wx.StaticText(userPanel, -1, _("Host / Server:"))
            r3Sizer.Add(lblDBServer, 1, wx.ALL, 5)

            # If Databases has entries, use that to create the Choice List for the Database Servers list.
            if self.Databases.keys() != []:
               choicelist = self.Databases.keys()
               choicelist.sort()
            # If not, create a list with a blank entry
            else:
               choicelist = ['']

            # As of wxPython 2.9.5.0, Mac doesn't support wx.CB_SORT and gives an ugly message about it!
            if 'wxMac' in wx.PlatformInfo:
                style = wx.CB_DROPDOWN
            else:
                style = wx.CB_DROPDOWN | wx.CB_SORT
            # Host / Server Combo Box, with a list of servers from the Databases Object if appropriate
            self.chDBServer = wx.ComboBox(userPanel, -1, choices=choicelist, style = style)

            # Set the value to the default value provided by the Configuration Data
            self.chDBServer.SetValue(TransanaGlobal.configData.host)
            r3Sizer.Add(self.chDBServer, 4, wx.EXPAND | wx.ALL, 5)
#            box2.Add(r3Sizer, 0, wx.EXPAND)
            userPanelSizer.Add(r3Sizer, 0, wx.EXPAND)

            # Define the Selection, SetFocus and KillFocus events for the Host / Server Combo Box
            wx.EVT_COMBOBOX(self, self.chDBServer.GetId(), self.OnServerSelect)

            # NOTE:  These events don't work on the MAC!  There appears to be a wxPython bug.  See wxPython ticket # 9862
            wx.EVT_KILL_FOCUS(self.chDBServer, self.OnServerKillFocus)

            # Create a Row Sizer for Port
            r4Sizer = wx.BoxSizer(wx.HORIZONTAL)
            # Define the Port TextCtrl and its KillFocus event
            lblPort = wx.StaticText(userPanel, -1, _("Port:"))
            r4Sizer.Add(lblPort, 1, wx.ALL, 5)
            self.txtPort = wx.TextCtrl(userPanel, -1, TransanaGlobal.configData.dbport, style=wx.TE_LEFT)
            r4Sizer.Add(self.txtPort, 4, wx.EXPAND | wx.ALL, 5)
#            box2.Add(r4Sizer, 0, wx.EXPAND)
            userPanelSizer.Add(r4Sizer, 0, wx.EXPAND)

            # This wx.EVT_SET_FOCUS is a poor attempt to compensate for wxPython bug # 9862
            if 'wxMac' in wx.PlatformInfo:
                self.txtPort.Bind(wx.EVT_SET_FOCUS, self.OnServerKillFocus)
            self.txtPort.Bind(wx.EVT_KILL_FOCUS, self.OnPortKillFocus)

#            # Let's add the MU controls we've created to the Data Entry Sizer
#            box2.AddMany([(lblUsername, 1, wx.RIGHT, 10),
#                          (self.txtUsername, 2, wx.EXPAND),
#                          (lblPassword, 1, wx.RIGHT, 10),
#                          (self.txtPassword, 2, wx.EXPAND),
#                          (lblDBServer, 1, wx.RIGHT, 10),
#                          (self.chDBServer, 2, wx.EXPAND),
#                          (lblPort, 1, wx.RIGHT, 10),
#                          (self.txtPort, 2, wx.EXPAND)
#                         ])
#        else:
#            # For single-user Transana, we only need one row with two columns
#            box2 = wx.FlexGridSizer(1, 2, 0, 0)
#            # We want the grid to grow horizontally
#            box2.SetFlexibleDirection(wx.HORIZONTAL)
#            # We want the data entry field to grow
#            box2.AddGrowableCol(1)
#            # Since there's only one row, the sizer proportion can be small.
#            box2Proportion = 2

            # The proportion for the data entry portion of the screen should be 6
#            box2Proportion = 0

            # Use a BoxSizer instead of a FlexGridSizer.  The FlexGridSizer isn't handling alternate font sizes
            # on Windows correctly.
#            box2 = wx.BoxSizer(wx.VERTICAL)

        # The rest of the controls are needed for both single- and multi-user versions.
        # Create a Row Sizer for Port
        r5Sizer = wx.BoxSizer(wx.HORIZONTAL)
        # Databases Label
        lblDBName = wx.StaticText(userPanel, -1, _("Database:"))
        r5Sizer.Add(lblDBName, 1, wx.ALL, 5)

        # If a Host is defined, get the list of Databases defined for that host.
        # The single-user version ...
        if TransanaConstants.singleUserVersion:
            # ... uses "localhost" as the server.
            DBServerName = 'localhost'
        # The multi-user version ...
        else:
            # ... uses whatever is selected in the DBServer Choice Box.
            DBServerName = self.chDBServer.GetValue()

        # There's a problem with wxPython's wxComboBox.  It sets the drop-down height to the number of options
        # in the initial choicelist, and keeps that even if the choicelist is changed.
        # To get around this, we give it a fake choice list with enough entries, then populate it later.
        choicelist = ['', '', '', '', '']

        # As of wxPython 2.9.5.0, Mac doesn't support wx.CB_SORT and gives an ugly message about it!
        if 'wxMac' in wx.PlatformInfo:
            style = wx.CB_DROPDOWN
        else:
            style = wx.CB_DROPDOWN | wx.CB_SORT
        # Database Combo Box
        self.chDBName = wx.ComboBox(userPanel, -1, choices=choicelist, style = style)

        # If we're in the Demo version ...
        if TransanaConstants.demoVersion:
            # ... then "Demonstration" is the only allowable Database Name
            self.chDBName.Clear()
            self.chDBName.Append("Demonstration")
            self.chDBName.SetValue("Demonstration")
            self.chDBName.Enable(False)
        # If we're NOT in the Demo version ...
        else:
            # If some Database have been defined...
            if len(self.Databases) >= 1:
                # ... Use the Databases object to get the list of Databases for the
                # identified Server
                if DBServerName != '':
                    choicelist = self.Databases[DBServerName]['dbList']
                    # Sort the Database List
                    choicelist.sort()
                else:
                    choicelist = ['']

                # Clear out the control's Choices ...
                self.chDBName.Clear()
                # ... and populate them with the appropriate values
                for choice in choicelist:
                    if ('unicode' in wx.PlatformInfo) and (isinstance(choice, str)):
                        choice = unicode(choice, 'utf8')  # TransanaGlobal.encoding)
                    self.chDBName.Append(choice)

            # if the configured database isn't in the database list, don't show it!
            # This can happen if you have a Russian database name but change away from Russian encoding
            # by changing languages during the session.

##            print
##            print "UsernameandPasswordClass.__init__():"
##            print TransanaGlobal.configData.database.encode('utf8'), type(TransanaGlobal.configData.database)
##            for x in range(len(choicelist)):
##                print choicelist[x], type(choicelist[x]), TransanaGlobal.configData.database.encode('utf8') == choicelist[x]
##            print TransanaGlobal.configData.database.encode('utf8') in choicelist, choicelist.index(TransanaGlobal.configData.database.encode('utf8'))
##            print
##
##            # If we're on the Mac ...
##            if 'wxMac' in wx.PlatformInfo:
##                # ... SetStringSelection is broken, so we locate the string's item number and use SetSelection!
##                if TransanaGlobal.configData.database in choicelist:
##                    self.chDBName.SetSelection(choicelist.index(TransanaGlobal.configData.database.encode('utf8')))  # (self.chDBName.GetItems().index(TransanaGlobal.configData.database.encode('utf8')))
##            else:
            if self.chDBName.FindString(TransanaGlobal.configData.database) != wx.NOT_FOUND:
                # Set the value to the default value provided by the Configuration Data
                self.chDBName.SetStringSelection(TransanaGlobal.configData.database)

        r5Sizer.Add(self.chDBName, 4, wx.EXPAND | wx.ALL, 5)
#        box2.Add(r5Sizer, 0, wx.EXPAND)
        userPanelSizer.Add(r5Sizer, 0, wx.EXPAND)

        # Define the SetFocus and KillFocus events for the Database Combo Box
        wx.EVT_KILL_FOCUS(self.chDBName, self.OnNameKillFocus)

        # Weird Mac bug ... Can't select all Database Names from the list sometimes.  So, if Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a Combo Box Event handler for the Database Name selection
            self.chDBName.Bind(wx.EVT_COMBOBOX, self.OnNameSelect)

#        # Add the Database name fields to the Data Entry sizer
#        box2.AddMany([(lblDBName, 1, wx.RIGHT, 10),
#                      (self.chDBName, 2, wx.EXPAND)
#                     ])
        # Now add the Data Entry sizer to the Main sizer
#        userPanelSizer.Add(box2, box2Proportion, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
#        userPanelSizer.Add(box2, 0, wx.LEFT | wx.RIGHT, 10)

        # If we are in the Multi-user Version ...
        if not TransanaConstants.singleUserVersion:

            # Add an SSL checkbox to the Username panel
            self.sslCheck = wx.CheckBox(userPanel, -1, _("Use SSL") + "  ", style=wx.CHK_2STATE | wx.ALIGN_RIGHT )
            # If SSL is True in the configuration file ...
            if TransanaGlobal.configData.ssl:
                # ... check the box
                self.sslCheck.SetValue(True)
            # Add a Spacer
            userPanelSizer.Add((0, 5))
            # Add the checkbox to the Username panel
            userPanelSizer.Add(self.sslCheck, 0, wx.LEFT | wx.ALIGN_LEFT, 5)
            # Add a Spacer
            userPanelSizer.Add((0, 10))

            # Add the User Panel to the first Notebook tab
            notebook.AddPage(userPanel, _("Database"), True)

            # Import the Transana OptionsSettings module, needed for the Message Server panel
            import OptionsSettings

            # Create teh Message Server Panel (which parents the MessageServer and MessageServerPort TextCtrls)
            self.messageServerPanel = OptionsSettings.MessageServerPanel(notebook, name='Username.MessageServerPanel')
            # Add the Message Server Panel to the second Notebook 
            notebook.AddPage(self.messageServerPanel, _("Message Server"), False)

            # Create a Panel for the SSL information
            sslPanel = wx.Panel(notebook, -1, size=notebook.GetSizeTuple(), name='UsernameandPasswordClass.sslPanel')
            # Create a Sizer for the SSL Panel
            sslSizer = wx.BoxSizer(wx.VERTICAL)

            # Get wxPython's Standard Paths
            sp = wx.StandardPaths.Get()
            # Set the initial SSL Directory to the user's Document Directory
            self.sslDir = sp.GetDocumentsDir()

            # Add the Client Certificate File to the SSL Tab
            lblClientCert = wx.StaticText(sslPanel, -1, _("Database Server SSL Client Certificate File"), style=wx.ST_NO_AUTORESIZE)
            # Add the label to the Panel Sizer
            sslSizer.Add(lblClientCert, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
            
            # Add a spacer
            sslSizer.Add((0, 3))
            
            # Add the Client Certificate File TextCtrl to the SSL Tab
            self.sslClientCert = wx.TextCtrl(sslPanel, -1, TransanaGlobal.configData.sslClientCert)
            # Add the element to the Panel Sizer
            sslSizer.Add(self.sslClientCert, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            # Add a Browse button for the Client Certificate File
            self.sslClientCertBrowse = wx.Button(sslPanel, -1, _("Browse"))
            # Add the button to the Sizer
            sslSizer.Add(self.sslClientCertBrowse, 0, wx.LEFT | wx.BOTTOM, 10)
            # Bind the button to the event processor
            self.sslClientCertBrowse.Bind(wx.EVT_BUTTON, self.OnSSLButton)
            # Add a spacer
            sslSizer.Add((0, 5))

            # Add the Client Key File to the SSL Tab
            lblClientKey = wx.StaticText(sslPanel, -1, _("Database Server SSL Client Key File"), style=wx.ST_NO_AUTORESIZE)
            # Add the label to the Panel Sizer
            sslSizer.Add(lblClientKey, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
            # Add a spacer
            sslSizer.Add((0, 3))
            
            # Add the Client Key File TextCtrl to the SSL Tab
            self.sslClientKey = wx.TextCtrl(sslPanel, -1, TransanaGlobal.configData.sslClientKey)
            # Add the element to the Panel Sizer
            sslSizer.Add(self.sslClientKey, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            # Add a Browse button for the Client Key File
            self.sslClientKeyBrowse = wx.Button(sslPanel, -1, _("Browse"))
            # Add the button to the Sizer
            sslSizer.Add(self.sslClientKeyBrowse, 0, wx.LEFT | wx.BOTTOM, 10)
            # Bind the button to the event processor
            self.sslClientKeyBrowse.Bind(wx.EVT_BUTTON, self.OnSSLButton)

            # Add the Message Server Certificate File to the SSL Tab
            lblMsgSrvCert = wx.StaticText(sslPanel, -1, _("Message Server SSL Certificate File"), style=wx.ST_NO_AUTORESIZE)
            # Add the label to the Panel Sizer
            sslSizer.Add(lblMsgSrvCert, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
            # Add a spacer
            sslSizer.Add((0, 3))
            
            # Add the Message Server Certificate File TextCtrl to the SSL Tab
            self.sslMsgSrvCert = wx.TextCtrl(sslPanel, -1, TransanaGlobal.configData.sslMsgSrvCert)
            # Add the element to the Panel Sizer
            sslSizer.Add(self.sslMsgSrvCert, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            # Add a Browse button for the Message Server Certificate File
            self.sslMsgSrvCertBrowse = wx.Button(sslPanel, -1, _("Browse"))
            # Add the button to the Sizer
            sslSizer.Add(self.sslMsgSrvCertBrowse, 0, wx.LEFT | wx.BOTTOM, 10)
            # Bind the button to the event processor
            self.sslMsgSrvCertBrowse.Bind(wx.EVT_BUTTON, self.OnSSLButton)

            # Set the SSL Panel's Sizer
            sslPanel.SetSizer(sslSizer)
            # Turn on AutoLayout
            sslPanel.SetAutoLayout(True)
            # Lay out the panel
            sslPanel.Layout()

            # Add the SSL Panel as the third Notebook tab
            notebook.AddPage(sslPanel, _("SSL"), False)

        # Create another sizer for the buttons, with a horizontal orientation
        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)

        # If this is NOT the Demo version ...
        if not TransanaConstants.demoVersion:
            # Define the "Delete Database" Button
            btnDeleteDatabase = wx.Button(self, -1, _("Delete Database"))
            self.Bind(wx.EVT_BUTTON, self.OnDeleteDatabase, btnDeleteDatabase)

        # Define the "OK" button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Make the OK button the default.  (This one works on Linux, while the next line doesn't!)
        btnOK.SetDefault()
        # Define the Default Button for the dialog.  This allows the "ENTER" key
        # to fire the OK button regardless of which widget has focus on the form.
        self.SetDefaultItem(btnOK)
        
        # Define the Cancel Button
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))

        # Define the "OK" button
        if not TransanaConstants.demoVersion:
            # Add the Delete Database button to the lower left corner
            buttonSizer.Add(btnDeleteDatabase, 3, wx.ALIGN_LEFT | wx.ALIGN_BOTTOM | wx.LEFT | wx.BOTTOM, 10)
        # Lets have some space between this button and  the others.
        buttonSizer.Add((30, 1), 1, wx.EXPAND)
        # Add the OK button to the lower right corner
        buttonSizer.Add(btnOK, 2, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT | wx.BOTTOM, 10)
        # Add the Cancel button to the lower right corner, bumping the OK button to the left
        buttonSizer.Add(btnCancel, 2, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT | wx.BOTTOM, 10)

        # Lay out the panel and the form, request AutoLayout for both.
        userPanel.Layout()
        userPanel.SetAutoLayout(True)

        # Set the main Panel's sizer to the main sizer and "fit" it.
        userPanel.SetSizer(userPanelSizer)
        userPanel.Fit()

        # Remember the size of the UserPanel at this point, to determine what DISPLAY TEXT size is selected in Windows
        userPanelSize = userPanel.GetSize()

        # Let's stick the panel on a sizer to fit in the Dialog
        panSizer = wx.BoxSizer(wx.VERTICAL)
        if TransanaConstants.singleUserVersion:
            panSizer.Add(userPanel, 1, wx.EXPAND)
        else:
            panSizer.Add(notebook, 1, wx.EXPAND)
        panSizer.Add((0, 3))
            
        # Add the Button sizer to the main sizer
        panSizer.Add(buttonSizer, 0, wx.EXPAND)

        self.SetSizer(panSizer)
        self.Layout()

        self.SetAutoLayout(True)
        self.Fit()

        # Lay out the dialog box, and tell it to resize automatically
        self.Layout()

        # If we are in the Multi-user Version ...
        if not TransanaConstants.singleUserVersion:
            # ... if we're using Small or Normal fonts ...
            if userPanelSize[1] < 350:
                # ... this size should work for the dialog
                self.SetSize((self.GetSize()[0], 410))
            # ... but if we're using Medium fonts ...
            else:
                # ... we need a larger dialog size
                self.SetSize((520, 500))

        # Set minimum window size
        dlgSize = self.GetSizeTuple()
        self.SetSizeHints(dlgSize[0], dlgSize[1])

        # Center the dialog on the screen
        self.CentreOnScreen()

        # The Mac version needs the focus to be set explicitly.
        # If we're running single-user Transana ...
        if TransanaConstants.singleUserVersion:
            # ... focus on the Database name, which is the only field
            self.chDBName.SetFocus()
        # If we're running multi-user Transana ...
        else:
            # ... and we don't know the username ...
            if self.txtUsername.GetValue() == '':
                # ... then focus on the username ...
                self.txtUsername.SetFocus()
            # ... but if we DO know the username ...
            else:
                # ... then let's focus on the password instead.
                self.txtPassword.SetFocus()

        # Show the Form modally, process if OK selected        
        if self.ShowModal() == wx.ID_OK:
          # Save the Values
          if TransanaConstants.singleUserVersion:
              self.Username = ''
              self.Password = ''
              self.DBServer = 'localhost'
              self.Port     = '3306'
              self.SSL      = False
              self.MessageServer = ''
              self.MessageServerPort = ''
              self.SSLClientCert = ''
              self.SSLClientKey = ''
              self.SSLMsgSrvCert = ''
          else:
              self.Username = self.txtUsername.GetValue()
              self.Password = self.txtPassword.GetValue()
              self.DBServer = self.chDBServer.GetValue()
              self.Port     = self.txtPort.GetValue()
              self.SSL      = self.sslCheck.IsChecked()
              self.MessageServer = self.messageServerPanel.messageServer.GetValue()
              self.MessageServerPort = self.messageServerPanel.messageServerPort.GetValue()
              self.SSLClientCert = self.sslClientCert.GetValue()
              self.SSLClientKey = self.sslClientKey.GetValue()
              self.SSLMsgSrvCert = self.sslMsgSrvCert.GetValue()
          self.DBName   = self.chDBName.GetValue()

          # the EVT_KILL_FOCUS for the Combo Boxes isn't getting called on the Mac.  Let's call it manually here
          if "__WXMAC__" in wx.PlatformInfo:
              self.OnNameKillFocus(None)

          # Since the list could have been changed here, pass it back to the Configuration module so the appropriate
          # data will be saved during program shutdown
          TransanaGlobal.configData.databaseList = self.Databases
        else:   
          self.Username = ''
          self.Password = ''
          self.DBServer = ''
          self.DBName   = ''
          self.Port     = ''
          self.SSL      = False
          self.MessageServer = ''
          self.MessageServerPort = ''
          self.SSLClientCert = ''
          self.SSLClientKey = ''
          self.SSLMsgSrvCert = ''
        return None

    def OnCloseWindow(self, event):
        event.Veto()

    def OnNameSelect(self, event):
        """ Process Database Name Selection (only used on Mac due to weird Mac bug!) """
        # Determine the string of the item number just selected by the user.
        selectStr = event.GetEventObject().GetItems()[event.GetEventObject().GetSelection()]
        # Normally, this string should ALWAYS be the same as the return of the GetValue() method, right?
        # Well, if it's not (this is the Mac bug) ...
        if event.GetEventObject().GetValue() != selectStr:
            # ... set the Database Name to the selected string
            self.DBName = selectStr
            # ... and set the Combo Box's selection to the correct selection number.
            event.GetEventObject().SetSelection(event.GetEventObject().GetSelection())
        
    def OnServerSelect(self, event):
        """ Process the Selection of a Database Host """
        # If we're on the Mac, there's a weird Combo Box bug on this form.
        if 'wxMac' in wx.PlatformInfo:
            # Determine the string of the item number just selected by the user.
            selectStr = event.GetEventObject().GetItems()[event.GetEventObject().GetSelection()]
            # Normally, this string should ALWAYS be the same as the return of the GetValue() method, right?
            # Well, if it's not (this is the Mac bug) ...
            if event.GetEventObject().GetValue() != selectStr:
                # ... set the Database Server to the selected string
                self.DBServer = selectStr
                # ... and set the Combo Box's selection to the correct selection number.
                event.GetEventObject().SetSelection(event.GetEventObject().GetSelection())

        # Clear the list of Databases
        self.chDBName.Clear()
        # If there is a database list defined for the Host selected ...
        if (event.GetString().strip() != '') and (self.Databases[event.GetString()]['dbList'] != []):
            # Get the list of databases for this server
            choiceList = self.Databases[event.GetString()]['dbList']
            # Sort the database list
            choiceList.sort()
            # ... iterate through the list of databases for the host ...
            for db in choiceList:
                # ... and put the databases in the Database Combo Box
                self.chDBName.Append(db)
        # If not, show an empty list
        else:
            self.chDBName.Append('')
        # Select the first entry in the list
        self.chDBName.SetSelection(0)
        # If we're in the multi-user version ...
        if (event.GetString().strip() != '') and (not TransanaConstants.singleUserVersion):
            # ... set the value for the server's Port based on the config data
            self.txtPort.SetValue(self.Databases[event.GetString()]['port'])

    def OnServerKillFocus(self, event):
        """ KillFocus event for Host/Server Combo Box """
        # See if the Host Name has not yet been used.
        if (self.chDBServer.GetValue() != '') and (not self.Databases.has_key(self.chDBServer.GetValue())):
            # If single-user, use port 3306
            if TransanaConstants.singleUserVersion:
                portVal = '3306'
            # if multi-user, get the Port value on screen
            else:
                portVal = self.txtPort.GetValue()
            # Add the new Server to the Databases Dictionary
            self.Databases[self.chDBServer.GetValue()] = {'dbList' : [],
                                                          'port' : portVal}
            # Add the new value to the control's dropdown
            self.chDBServer.Append(self.chDBServer.GetValue())
            # Update the DBName Control based on the new Server
            self.chDBName.Clear()
            self.chDBName.Append('')
            self.chDBName.SetSelection(0)
        self.chDBServer.SetInsertionPointEnd()
        event.Skip()

    def OnNameKillFocus(self, event):
        """ KillFocus event for Database Combo Box """
        # Determine the Database Server name and port
        if TransanaConstants.singleUserVersion:
            DBServerName = 'localhost'
            portVal = '3306'
        else:
            DBServerName = self.chDBServer.GetValue()
            portVal = self.txtPort.GetValue()
        # If we have a NEW Database server ...
        if not self.Databases.has_key(DBServerName):
            # ... create a new entry with a blank database list and the correct port
            self.Databases[DBServerName] = {'dbList' : [],
                                            'port' : portVal}
        if 'unicode' in wx.PlatformInfo:
            try:
                dbName = self.chDBName.GetValue().encode('utf8')  #(TransanaGlobal.encoding)
            except UnicodeEncodeError:
                # If you've changed languages in the single-user version of Transana, this COULD cause a change in encodings.
                # In this case, you might not be able to decode dbName because it's in the wrong encoding, causing Transana to
                # crash and burn.  This attempts to fix that problem by forgetting the un-decodable database name.
                dbName = ''
        else:
            dbName = self.chDBName.GetValue()
        
        # See if the Database Name has not yet been added to the database list.  If not ...
        if (self.chDBName.GetValue() != '') and \
           (self.Databases.has_key(DBServerName)) and \
           (not dbName in self.Databases[DBServerName]['dbList']):
            # Add the new Database Name to the Database List
            self.Databases[DBServerName]['dbList'].append(dbName)
            # Add the new value to the control's dropdown
            self.chDBName.Append(self.chDBName.GetValue())
        self.chDBName.SetInsertionPointEnd()
        # On the Mac, this method doesn't get called because EVT_KILL_FOCUS is broken.  Therefore, we call it manually
        # with the event set to None.  To avoid an exception here, we need to test for that.
        if event != None:
            event.Skip()

    def OnPortKillFocus(self, event):
        """ Lose Focus event for Port Entry """
        # If this server has not yet been used, it won't have the correct config information.  Check.
        if not self.Databases.has_key(self.chDBServer.GetValue()):
            # If it doesn't exist, create a dictionary for saving config values for this server.
            self.Databases[self.chDBServer.GetValue()] = {'dbList' : [],
                                                          'port' : self.txtPort.GetValue()}
        else:
            # When leaving the port field, we need to update the configuration data object
            self.Databases[self.chDBServer.GetValue()]['port'] = self.txtPort.GetValue()

    def OnSSLButton(self, event):
        """ Handle the Browse buttons for the SSL Client Certificate and the SSL CLient Key file fields """
        # Define the File Type as *.pem files
        fileType = '*.pem'
        fileTypesString = _("SSL Certificate Files (*.pem)|*.pem|All files (*.*)|*.*")
        if event.GetId() == self.sslClientCertBrowse.GetId():
            prompt = _("Select the SSL Client Certificate file")
            fileName = self.sslClientCert.GetValue()
        elif event.GetId() == self.sslClientKeyBrowse.GetId():
            prompt = _("Select the SSL Client Key file")
            fileName = self.sslClientKey.GetValue()
        elif event.GetId() == self.sslMsgSrvCertBrowse.GetId():
            prompt = _("Select the SSL Message Server Certificate file")
            fileName = self.sslMsgSrvCert.GetValue()
        (path, flnm) = os.path.split(fileName)
        if path != '':
            self.sslDir = path
        # Invoke the File Selector with the proper default directory, filename, file type, and style
        fs = wx.FileSelector(prompt, self.sslDir, fileName, fileType, fileTypesString, wx.OPEN | wx.FILE_MUST_EXIST)
        # If user didn't cancel ..
        if fs != "":
            # Mac Filenames use a different encoding system.  We need to adjust the string returned by the FileSelector.
            # Surely there's an easier way, but I can't figure it out.
            if 'wxMac' in wx.PlatformInfo:
                import Misc
                fs = Misc.convertMacFilename(fs)
            if event.GetId() == self.sslClientCertBrowse.GetId():
                self.sslClientCert.SetValue(fs)
            elif event.GetId() == self.sslClientKeyBrowse.GetId():
                self.sslClientKey.SetValue(fs)
            elif event.GetId() == self.sslMsgSrvCertBrowse.GetId():
                self.sslMsgSrvCert.SetValue(fs)
            (path, flnm) = os.path.split(fs)
            if path != '':
                self.sslDir = path

    def GetValues(self):
        """ Get the Data Values the user entered on the UserName panel """
        return (self.Username, self.Password, self.DBServer, self.DBName, self.Port)

    def GetMultiUserValues(self):
        """ Get all Data Values needed for Multi-user version """
        return (self.SSL, self.MessageServer, self.MessageServerPort, self.SSLClientCert, self.SSLClientKey, self.SSLMsgSrvCert)

    def GetUsername(self):
        """ Get the User Name Entry """
        return self.Username

    def GetPassword(self):
        """ Get the Password Entry """
        return self.Password

    def GetDBServer(self):
        """ Get the Host / Server Selection """
        return self.DBServer

    def GetDBName(self):
        """ Get the Database Selection """
        return self.DBName

    def GetPort(self):
        """ Get the Port Selection """
        return self.Port

    def GetSSL(self):
        """ Is SSL Required? """
        return self.SSL

    def GetMessageServer(self):
        """ Get the Message Server Entry """
        return self.messageServer

    def GetMessageServerPort(self):
        """ Get the Message Server Port Entry """
        return self.messageServerPort

    def GetSSLClientCert(self):
        """ Get the SSL Client Certificate Entry """
        return self.SSLClientCert

    def GetSSLClientKey(self):
        """ Get the SSL Client Key Entry """
        return self.SSLClientKey

    def GetMsgSrvCert(self):
        """ Get the SSL Message Server Certificate """
        return self.SSLMsgSrvCert

    def OnDeleteDatabase(self, event):
        """ Delete a database """
        # There are two reasons we only want the full Delete Database function on the single-user
        # version.  First, we don't want everyone to have the capacity to delete the communal
        # database.  That should be limited to the DBA.  Second, there appears to be a bug
        # somewhere in MySQL for Python (maybe?) or my code (maybe?) or MySQL's BDB tables (maybe?)
        # that makes it so that if someone
        # creates a database and then deletes it, the server can get trashed if it gets shut down.
        # Literally, you won't be able to restart the server.
        #
        # Therefore, the multi-user version of this routine will only remove the database name from
        # the user's dropdown list of databases.  It won't actually delete any data.

        if TransanaConstants.singleUserVersion:
            errormsg = ''

            if self.chDBName.GetValue() == '':
                errormsg += _('You must specify a Database.\n')

            if errormsg == '':
                username = ''
                password = ''
                server = 'localhost'
                database = self.chDBName.GetValue()
                port = '3306'

                # Get the name of the database to delete
                if 'unicode' in wx.PlatformInfo:
                    dbName = self.chDBName.GetValue().encode('utf8')
                else:
                    dbName = self.chDBName.GetValue()
                delresult = DBInterface.DeleteDatabase(username, password, server, database, port)
                if delresult == 0:
                    msg = _('Transana could not delete database "%s".\nHowever, it has been removed from the list of databases you have used.')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8')
                    dlg = Dialogs.InfoDialog(None, msg % database)
                    dlg.ShowModal()
                    dlg.Destroy()
                if delresult in [0, 1]:
                    # If we've deleted the database, remove that database name from the list of existing databases
                    self.Databases[server]['dbList'].remove(dbName)
                    # Clear the database name from the screen control
                    self.chDBName.SetValue('')
                    # Clear out the Database name control's Choices ...
                    self.chDBName.Clear()
                    # ... and populate them with the appropriate values from the database list
                    for choice in self.Databases[server]['dbList']:
                        self.chDBName.Append(choice)
                    # Remove the Database Name from the Configuration record of existing databases
                    TransanaGlobal.configData.databaseList = self.Databases
                    # Remove the Database Name from the Configuration Record for the current database
                    TransanaGlobal.configData.database = ''
                    # Start exception handling
                    try:
                        # Now remove the database Paths for this database
                        del TransanaGlobal.configData.pathsByDB[(username, server, database)]
                    # If an exception is raised ...
                    except:
                        # ... we can ignore it.
                        pass
                    # Save the Config changes
                    TransanaGlobal.configData.SaveConfiguration()
            else:
                dlg = Dialogs.ErrorDialog(None, errormsg)
                dlg.ShowModal()
                dlg.Destroy()
        else:
            server = self.chDBServer.GetValue()
            database = self.chDBName.GetValue()

            if (server != '') and (database != ''):
                # Get the name of the database to delete
                if 'unicode' in wx.PlatformInfo:
                    dbName = database.encode('utf8')
                else:
                    dbName = database
                # If we've deleted the database, remove that database name from the list of existing databases
                self.Databases[server]['dbList'].remove(dbName)
                # Clear the database name from the screen control
                self.chDBName.SetValue('')
                # Clear out the Database name control's Choices ...
                self.chDBName.Clear()
                # ... and populate them with the appropriate values from the database list
                for choice in self.Databases[server]['dbList']:
                    self.chDBName.Append(choice)
                # Remove the Database Name from the Configuration record of existing databases
                TransanaGlobal.configData.databaseList = self.Databases
                # Remove the Database Name from the Configuration Record for the current database
                TransanaGlobal.configData.database = ''
                msg = _("The multi-user version of Transana does not allow users to delete databases.\nPlease ask your project's Database Administrator to delete database '%s'.\nHowever, it has been removed from the list of databases you have used.")
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(msg, 'utf8')
                dlg = Dialogs.InfoDialog(None, msg % database)
                dlg.ShowModal()
                dlg.Destroy()

            else:
                errormsg = ''
                if self.chDBServer.GetValue() == '':
                    errormsg += _('You must specify a Database Server.\n')
                if self.chDBName.GetValue() == '':
                    errormsg += _('You must specify a Database.\n')
                dlg = Dialogs.ErrorDialog(None, errormsg)
                dlg.ShowModal()
                dlg.Destroy()


if __name__ == '__main__':

    __builtins__._ = wx.GetTranslation

    class MyApp(wx.App):
        def OnInit(self):
            self.frame = UsernameandPassword(None)
            self.frame.Destroy()
            return(True)

    app = MyApp(0)
    app.MainLoop()
