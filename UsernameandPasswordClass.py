#Copyright (C) 2003 - 2010  The Board of Regents of the University of Wisconsin System
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
            dlgSize=(350, 250)
            # Instructions Text
            instructions = _("Please enter your MySQL username and password, as well \nas the names of the server and database you wish to use.\nTo create a new database, type in a new database name\n(if you have appropriate permissions.)\n(Database names may contain only letters and numbers in\na single word.)")
            # Macs don't need as much space for instructions as Windows machines do.
            if 'wxMac' in wx.PlatformInfo:
                # Sizer Proportion for the instructions
                instProportion = 4
            else:
                # Sizer Proportion for the instructions
                instProportion = 5

        # Define the main Dialog Box
        wx.Dialog.__init__(self, parent, -1, dlgTitle, size=dlgSize, style=wx.CAPTION | wx.RESIZE_BORDER)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create a BoxSizer for the main elements to go onto the Panel
        box = wx.BoxSizer(wx.VERTICAL)

        # Place a Panel on the Dialog Box.  All Controls will go on the Panel.
        # (Panels can have DefaultItems, enabling the desired "ENTER" key functionality.
        panel = wx.Panel(self, -1, name='UserNamePanel')

        # Instructions Text
        lblIntro = wx.StaticText(panel, -1, instructions)
        # Add the Instructions to the Main Sizer
        box.Add(lblIntro, instProportion, wx.EXPAND | wx.ALL, 10)

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
            box2 = wx.FlexGridSizer(5, 2, 6, 0)
            # We want to be flexible horizontally
            box2.SetFlexibleDirection(wx.HORIZONTAL)
            # We want the data entry fields to expand
            box2.AddGrowableCol(1)
            # The proportion for the data entry portion of the screen should be 6
            box2Proportion = 6

            # Username Label        
            lblUsername = wx.StaticText(panel, -1, _("Username:"))

            # User Name TextCtrl
            self.txtUsername = wx.TextCtrl(panel, -1, style=wx.TE_LEFT)
            if DBInterface.get_username() != '':
                self.txtUsername.SetValue(DBInterface.get_username())

            # Password Label
            lblPassword = wx.StaticText(panel, -1, _("Password:"))

            # Password TextCtrl (with PASSWORD style)
            self.txtPassword = wx.TextCtrl(panel, -1, style=wx.TE_LEFT|wx.TE_PASSWORD)

            # Host / Server Label
            lblDBServer = wx.StaticText(panel, -1, _("Host / Server:"))

            # If Databases has entries, use that to create the Choice List for the Database Servers list.
            if self.Databases.keys() != []:
               choicelist = self.Databases.keys()
            # If not, create a list with a blank entry
            else:
               choicelist = ['']
               
            # Host / Server Combo Box, with a list of servers from the Databases Object if appropriate
            self.chDBServer = wx.ComboBox(panel, -1, choices=choicelist, style = wx.CB_DROPDOWN | wx.CB_SORT)

            # Set the value to the default value provided by the Configuration Data
            self.chDBServer.SetValue(TransanaGlobal.configData.host)

            # Define the Selection, SetFocus and KillFocus events for the Host / Server Combo Box
            wx.EVT_COMBOBOX(self, self.chDBServer.GetId(), self.OnServerSelect)

            # NOTE:  These events don't work on the MAC!  There appears to be a wxPython bug.  See wxPython ticket # 9862
            wx.EVT_SET_FOCUS(self.chDBServer, self.OnCBSetFocus)
            wx.EVT_KILL_FOCUS(self.chDBServer, self.OnServerKillFocus)

            # Define the Port TextCtrl and its KillFocus event
            lblPort = wx.StaticText(panel, -1, _("Port:"))
            self.txtPort = wx.TextCtrl(panel, -1, TransanaGlobal.configData.dbport, style=wx.TE_LEFT)

            # This wx.EVT_SET_FOCUS is a poor attempt to compensate for wxPython bug # 9862
            if 'wxMac' in wx.PlatformInfo:
                self.txtPort.Bind(wx.EVT_SET_FOCUS, self.OnServerKillFocus)
            self.txtPort.Bind(wx.EVT_KILL_FOCUS, self.OnPortKillFocus)

            # Let's add the MU controls we've created to the Data Entry Sizer
            box2.AddMany([(lblUsername, 1, wx.RIGHT, 10),
                          (self.txtUsername, 2, wx.EXPAND),
                          (lblPassword, 1, wx.RIGHT, 10),
                          (self.txtPassword, 2, wx.EXPAND),
                          (lblDBServer, 1, wx.RIGHT, 10),
                          (self.chDBServer, 2, wx.EXPAND),
                          (lblPort, 1, wx.RIGHT, 10),
                          (self.txtPort, 2, wx.EXPAND)
                         ])
        else:
            # For single-user Transana, we only need one row with two columns
            box2 = wx.FlexGridSizer(1, 2, 0, 0)
            # We want the grid to grow horizontally
            box2.SetFlexibleDirection(wx.HORIZONTAL)
            # We want the data entry field to grow
            box2.AddGrowableCol(1)
            # Since there's only one row, the sizer proportion can be small.
            box2Proportion = 2

        # The rest of the controls are needed for both single- and multi-user versions.
        # Databases Label
        lblDBName = wx.StaticText(panel, -1, _("Database:"))

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

        # Database Combo Box
        self.chDBName = wx.ComboBox(panel, -1, choices=choicelist, style = wx.CB_DROPDOWN | wx.CB_SORT)

        if TransanaConstants.demoVersion:
            self.chDBName.Clear()
            self.chDBName.Append("Demonstration")
            self.chDBName.SetValue("Demonstration")
            self.chDBName.Enable(False)

        else:
            # If some Database have been defined...
            if len(self.Databases) >= 1:
                # ... Use the Databases object to get the list of Databases for the
                # identified Server
                if DBServerName != '':
                    choicelist = self.Databases[DBServerName]['dbList']
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
            if self.chDBName.FindString(TransanaGlobal.configData.database) != wx.NOT_FOUND:
                # Set the value to the default value provided by the Configuration Data
                self.chDBName.SetValue(TransanaGlobal.configData.database)

        # Define the SetFocus and KillFocus events for the Database Combo Box
        wx.EVT_SET_FOCUS(self.chDBName, self.OnCBSetFocus)
        wx.EVT_KILL_FOCUS(self.chDBName, self.OnNameKillFocus)

        # Add the Database name fields to the Data Entry sizer
        box2.AddMany([(lblDBName, 1, wx.RIGHT, 10),
                      (self.chDBName, 2, wx.EXPAND)
                     ])
        # Now add the Data Entry sizer to the Main sizer
        box.Add(box2, box2Proportion, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

        # Create another sizer for the buttons, with a horizontal orientation
        box4 = wx.BoxSizer(wx.HORIZONTAL)

        if not TransanaConstants.demoVersion:
            # Define the "Delete Database" Button
            btnDeleteDatabase = wx.Button(panel, -1, _("Delete Database"))
            self.Bind(wx.EVT_BUTTON, self.OnDeleteDatabase, btnDeleteDatabase)

        # Define the "OK" button
        btnOK = wx.Button(panel, wx.ID_OK, _("OK"))

        # Define the Default Button for the dialog.  This allows the "ENTER" key
        # to fire the OK button regardless of which widget has focus on the form.
        self.SetDefaultItem(btnOK)
        
        # Define the Cancel Button
        btnCancel = wx.Button(panel, wx.ID_CANCEL, _("Cancel"))

        if not TransanaConstants.demoVersion:
            # Add the Delete Database button to the lower left corner
            box4.Add(btnDeleteDatabase, 3, wx.ALIGN_LEFT | wx.ALIGN_BOTTOM | wx.LEFT | wx.BOTTOM, 10)
        # Lets have some space between this button and  the others.
        box4.Add((30, 1), 1, wx.EXPAND)
        # Add the OK button to the lower right corner
        box4.Add(btnOK, 2, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT | wx.BOTTOM, 10)
        # Add the Cancel button to the lower right corner, bumping the OK button to the left
        box4.Add(btnCancel, 2, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT | wx.BOTTOM, 10)
        # Add the Button sizer to the main sizer
        box.Add(box4, 2, wx.EXPAND)

        # Lay out the panel and the form, request AutoLayout for both.
        panel.Layout()
        panel.SetAutoLayout(True)

        # Set the main Panel's sizer to the main sizer and "fit" it.
        panel.SetSizer(box)
        panel.Fit()

        # Let's stick the panel on a sizer to fit in the Dialog
        panSizer = wx.BoxSizer(wx.VERTICAL)
        panSizer.Add(panel, 1, wx.EXPAND)
        self.SetSizer(panSizer)
        self.Fit()

        # Lay out the dialog box, and tell it to resize automatically
        self.Layout()
        self.SetAutoLayout(True)

        # Set minimum window size
        dlgSize = self.GetSizeTuple()
        self.SetSizeHints(dlgSize[0], dlgSize[1], -1, round(dlgSize[1]))

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
          else:
              self.Username = self.txtUsername.GetValue()
              self.Password = self.txtPassword.GetValue()
              self.DBServer = self.chDBServer.GetValue()
              self.Port     = self.txtPort.GetValue()
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
          self.Port = ''
        return None

    def OnServerSelect(self, event):
        """ Process the Selection of a Database Host """
        # Clear the list of Databases
        self.chDBName.Clear()
        # If there is a database list defined for the Host selected ...
        if (event.GetString().strip() != '') and (self.Databases[event.GetString()]['dbList'] != []):
            # ... iterate through the list of databases for the host ...
            for db in self.Databases[event.GetString()]['dbList']:
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

    def OnCBSetFocus(self, event):
        """ Combo Box Set Focus Event """
        # Do nothing
        event.Skip()
        
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
            

    def GetValues(self):
        """ Get all Data Values the user entered into this Dialog Box """
        return (self.Username, self.Password, self.DBServer, self.DBName, self.Port)

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
                elif delresult == 1:
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
