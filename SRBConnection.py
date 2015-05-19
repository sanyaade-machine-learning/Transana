# Copyright (C) 2003 - 2007  The Board of Regents of the University of Wisconsin System 
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

"""This module implements the logon screen used in connecting to the Storage Resource Broker (SRB).  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import wx

RB_STORAGESPACE   =   wx.NewId()

class SRBConnection(wx.Dialog):
    """ Implements the logon screen used in connecting to the Storage Resource Broker (SRB). """
    def __init__(self, parent):
        wx.Dialog.__init__(self, parent, -1, _("SRB Connection Parameters"), size=(250, 520), style=wx.CAPTION|wx.RESIZE_BORDER|wx.NO_FULL_REPAINT_ON_RESIZE)

        self.LoadConfiguration()
        
        # Create a Sizer
        box = wx.BoxSizer(wx.VERTICAL)

        lblUserName = wx.StaticText(self, -1, _("User Name:"))
        box.Add(lblUserName, 0, wx.LEFT, 10)

        self.editUserName = wx.TextCtrl(self, -1, self.srbUserName)
        box.Add(self.editUserName, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 10)
        
        lblPassword = wx.StaticText(self, -1, _("Password:"))
        box.Add(lblPassword, 0, wx.LEFT, 10)

        self.editPassword = wx.TextCtrl(self, -1, style=wx.TE_PASSWORD)
        box.Add(self.editPassword, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)
        
        lblDomain = wx.StaticText(self, -1, _("Domain:"))
        box.Add(lblDomain, 0, wx.LEFT, 10)

        self.editDomain = wx.TextCtrl(self, -1, self.srbDomain)
        box.Add(self.editDomain, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        lblCollectionRoot = wx.StaticText(self, -1, _("Collection Root:"))
        box.Add(lblCollectionRoot, 0, wx.LEFT, 10)

        self.editCollectionRoot = wx.TextCtrl(self, -1, self.srbCollectionRoot)
        box.Add(self.editCollectionRoot, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        lblSRBHost = wx.StaticText(self, -1, _("SRB Host:"))
        box.Add(lblSRBHost, 0, wx.LEFT, 10)

        self.editSRBHost = wx.TextCtrl(self, -1, self.srbHost)
        box.Add(self.editSRBHost, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        lblSRBPort = wx.StaticText(self, -1, _("SRB Port:"))
        box.Add(lblSRBPort, 0, wx.LEFT, 10)

        self.editSRBPort = wx.TextCtrl(self, -1, self.srbPort)
        box.Add(self.editSRBPort, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        lblResource = wx.StaticText(self, -1, _("Resource:"))
        box.Add(lblResource, 0, wx.LEFT, 10)

        self.editSRBResource = wx.TextCtrl(self, -1, self.srbResource)
        box.Add(self.editSRBResource, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        lblSRBSEAOption = wx.StaticText(self, -1, _("SEA Option:"))
        box.Add(lblSRBSEAOption, 0, wx.LEFT, 10)

        self.editSRBSEAOption = wx.TextCtrl(self, -1, self.srbSEAOption)
        box.Add(self.editSRBSEAOption, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        lblBuffer = wx.StaticText(self, -1, _("Buffer Size:"))
        box.Add(lblBuffer, 0, wx.LEFT, 10)

        self.choiceBuffer = wx.Choice(self, -1, size=wx.Size(100, 24), choices=['4096', '8192', '16384', '32767', '64000', '100000', '200000','300000','400000','500000'])
        self.choiceBuffer.SetStringSelection(self.srbBuffer)
        box.Add(self.choiceBuffer, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        self.rbSRBStorageSpace = wx.RadioBox(self, RB_STORAGESPACE, _("Connect to:"), choices=[" " + _("My own storage space"), " " + _("All user's storage spaces") + "                  "], majorDimension=2, style=wx.RA_SPECIFY_ROWS)
        box.Add(self.rbSRBStorageSpace, 5, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        btnBox = wx.BoxSizer(wx.HORIZONTAL)

        btnBox.Add((20, 0))

        btnConnect = wx.Button(self, wx.ID_OK, _("Connect"))
        btnBox.Add(btnConnect, 1, wx.ALIGN_CENTER_HORIZONTAL)

        # Make the Connect button the default
        self.SetDefaultItem(btnConnect)
       
        btnBox.Add((20, 0))

        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        btnBox.Add(btnCancel, 1, wx.ALIGN_CENTER_HORIZONTAL)

        btnBox.Add((20, 0))
        box.Add(btnBox, 0, wx.ALIGN_CENTER)
        box.Add((0, 10))

        self.SetSizer(box)
        self.Fit()

        self.Layout()
        self.SetAutoLayout(True)
        # Center on the Screen
        self.CenterOnScreen()

        # Set Focus
        self.editUserName.SetFocus()

    def LoadConfiguration(self):
        """ Load Configuration Data from the Registry or Config File """
        # Set Default Values
        defaultsrbDomain         = 'digital-insight'
        defaultsrbCollectionRoot = '/WCER/home/'       # '/home/'
        defaultsrbHost           = 'storage.wcer.wisc.edu'    # 'srb.wcer.wisc.edu'
        defaultsrbPort           = '5544'              # '5823'
        defaultsrbResource       = 'WCERSRBV1'         # 'nt-wcersrb-1'
        defaultsrbSEAOption      = 'ENCRYPT1'
        defaultsrbBuffer         = '400000'
        # Load the Config Data.  wxConfig automatically uses the Registry on Windows and the appropriate file on Mac.
        # Program Name is Transana, Vendor Name is Verception to remain compatible with Transana 1.0.
        config = wx.Config('Transana', 'Verception')
        # See if a version 2.0 Configuration exists, and use it if so
        if config.Exists('/2.0/srb'):
            self.srbUserName       = config.Read('/2.0/srb/srbUserName', '')
            self.srbDomain         = config.Read('/2.0/srb/srbDomain', defaultsrbDomain)
            self.srbCollectionRoot = config.Read('/2.0/srb/srbCollectionRoot', defaultsrbCollectionRoot)
            self.srbHost           = config.Read('/2.0/srb/srbHost', defaultsrbHost)
            self.srbPort           = config.Read('/2.0/srb/srbPort', defaultsrbPort)
            self.srbResource       = config.Read('/2.0/srb/srbResource', defaultsrbResource)
            self.srbSEAOption      = config.Read('/2.0/srb/srbSEAOptions', defaultsrbSEAOption)

        # If no version 2.0 Config File exists, ...
        else:

            # See if a verion 1 config file exists, and use it's data if it does.
            if config.Exists('/1.0/'):
                self.srbUserName       = config.Read('/1.0/srbUserName', '')
                self.srbDomain         = config.Read('/1.0/srbDomain', defaultsrbDomain)
                self.srbCollectionRoot = defaultsrbCollectionRoot
                self.srbHost           = config.Read('/1.0/srbHost', defaultsrbHost)
                self.srbPort           = config.Read('/1.0/srbPort', defaultsrbPort)
                self.srbResource       = config.Read('/1.0/srbResource', defaultsrbResource)
                self.srbSEAOption      = config.Read('/1.0/srbSEAOptions', defaultsrbSEAOption)

            # If no Config File exists, use default settings
            else:
                self.srbUserName       = ''
                self.srbDomain         = defaultsrbDomain
                self.srbCollectionRoot = defaultsrbCollectionRoot
                self.srbHost           = defaultsrbHost
                self.srbPort           = defaultsrbPort
                self.srbResource       = defaultsrbResource
                self.srbSEAOption      = defaultsrbSEAOption

        self.srbBuffer         = config.Read('/2.0/srb/srbBuffer', defaultsrbBuffer)
            

    def SaveConfiguration(self):
        """ Save Configuration Data to the Registry or a Config File. """
        # Save the Config Data.  wxConfig automatically uses the Registry on Windows and the appropriate file on Mac.
        # Program Name is Transana, Vendor Name is Verception to remain compatible with Transana 1.0.
        config = wx.Config('Transana', 'Verception')
        self.srbUserName = self.editUserName.GetValue()
        config.Write('/2.0/srb/srbUserName', self.srbUserName)
        self.srbDomain = self.editDomain.GetValue()
        config.Write('/2.0/srb/srbDomain', self.srbDomain)
        self.srbCollectionRoot = self.editCollectionRoot.GetValue()
        config.Write('/2.0/srb/srbCollectionRoot', self.srbCollectionRoot)
        self.srbHost = self.editSRBHost.GetValue()
        config.Write('/2.0/srb/srbHost', self.srbHost)
        self.srbPort = self.editSRBPort.GetValue()
        config.Write('/2.0/srb/srbPort', self.srbPort)
        self.srbResource = self.editSRBResource.GetValue()
        config.Write('/2.0/srb/srbResource', self.srbResource)
        self.srbSEAOption = self.editSRBSEAOption.GetValue()
        config.Write('/2.0/srb/srbSEAOptions', self.srbSEAOption)
        self.srbBuffer = self.choiceBuffer.GetStringSelection()
        config.Write('/2.0/srb/srbBuffer', self.srbBuffer)
