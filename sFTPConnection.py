# Copyright (C) 2003 - 2012  The Board of Regents of the University of Wisconsin System 
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

"""This module implements the logon screen used in connecting to an sFTP Server.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import wx

# RB_STORAGESPACE   =   wx.NewId()

class sFTPConnection(wx.Dialog):
    """ Implements the logon screen used in connecting to an sFTP Server). """
    def __init__(self, parent):
        wx.Dialog.__init__(self, parent, -1, _("sFTP Connection Parameters"), size=(250, 520), style=wx.CAPTION|wx.RESIZE_BORDER|wx.NO_FULL_REPAINT_ON_RESIZE)

        self.LoadConfiguration()
        
        # Create a Sizer
        box = wx.BoxSizer(wx.VERTICAL)

        lblUserName = wx.StaticText(self, -1, _("User Name:"))
        box.Add(lblUserName, 0, wx.LEFT, 10)

        self.editUserName = wx.TextCtrl(self, -1)
        box.Add(self.editUserName, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 10)
        
        lblPassword = wx.StaticText(self, -1, _("Password:"))
        box.Add(lblPassword, 0, wx.LEFT, 10)

        self.editPassword = wx.TextCtrl(self, -1, style=wx.TE_PASSWORD)
        box.Add(self.editPassword, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)
        
        lblsFTPServer = wx.StaticText(self, -1, _("sFTP Server:"))
        box.Add(lblsFTPServer, 0, wx.LEFT, 10)

        self.editsFTPServer = wx.TextCtrl(self, -1, self.sFTPServer)
        box.Add(self.editsFTPServer, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        lblsFTPPort = wx.StaticText(self, -1, _("sFTP Port:"))
        box.Add(lblsFTPPort, 0, wx.LEFT, 10)

        self.editsFTPPort = wx.TextCtrl(self, -1, self.sFTPPort)
        box.Add(self.editsFTPPort, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        lblsFTPPublicKeyType = wx.StaticText(self, -1, _("sFTP Server Public Key Type:"))
        box.Add(lblsFTPPublicKeyType, 0, wx.LEFT, 10)

        self.choicesFTPPublicKeyType = wx.Choice(self, -1, choices=[_('None'), 'ssh-rsa', 'ssh-dss'])
        self.choicesFTPPublicKeyType.SetStringSelection(self.sFTPPublicKeyType)
        self.choicesFTPPublicKeyType.Bind(wx.EVT_CHOICE, self.OnPublicKeyTypeSelect)
        box.Add(self.choicesFTPPublicKeyType, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

        lblsFTPPublicKey = wx.StaticText(self, -1, _("sFTP Server Public Key:"))
        box.Add(lblsFTPPublicKey, 0, wx.LEFT, 10)

        self.editsFTPPublicKey = wx.TextCtrl(self, -1, self.sFTPPublicKey)
        if self.sFTPPublicKeyType == _("None"):
            self.editsFTPPublicKey.Enable(False)
        box.Add(self.editsFTPPublicKey, 2, wx.LEFT | wx.RIGHT | wx.BOTTOM| wx.EXPAND, 10)

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

    def OnPublicKeyTypeSelect(self, event):
        if self.choicesFTPPublicKeyType.GetStringSelection() == _('None'):
            self.editsFTPPublicKey.Enable(False)
        else:
            self.editsFTPPublicKey.Enable(True)

    def LoadConfiguration(self):
        """ Load Configuration Data from the Registry or Config File """
        # Load the Config Data.  wxConfig automatically uses the Registry on Windows and the appropriate file on Mac.
        # Program Name is Transana, Vendor Name is Verception to remain compatible with Transana 1.0.
        config = wx.Config('Transana', 'Verception')
#        self.sFTPUserName       = config.Read('/2.0/sFTP/sFTPUserName', '')
        self.sFTPServer         = config.Read('/2.0/sFTP/sFTPServer', 'ftp.wcer.wisc.edu')
        self.sFTPPort           = config.Read('/2.0/sFTP/sFTPPort', '22')
        self.sFTPPublicKeyType  = config.Read('/2.0/sFTP/sFTPPublicKeyType', 'ssh-rsa')
        self.sFTPPublicKey      = config.Read('/2.0/sFTP/sFTPPublicKey', 'AAAAB3NzaC1yc2EAAAABIwAAAQEAzi+1CH0qa4UwRgxf2J4cQE6HNJYSTYaDlionweNWmRBBXG5mu2L7bAXYx32QnRwukbZNJAF/APm80aLbQo/m6nn8Do5eSWDkel0J1GeXOFd/yfINrcNMN2UD2r7J0o6PMBtZVQq2logm1Ckbu2UDhW8HWDMKC1YsfU4y5HTd09qvduEGMCW7YsKECiqlBBX8/Pg0GuThi0h5IOuyufpCpTPaUxL0tBoIo7pRH2Dax5ivtAaxO+xWP8mMnOCLjzxHknD0z3h/bmr8QJ4o9vR8xjXab/Skrtmd1FSqei4cdWFYW8jDdbPMS14sOTj0pk58/8I7h838kQ7WkYwNcC4fEQ==')

    def SaveConfiguration(self):
        """ Save Configuration Data to the Registry or a Config File. """
        # Save the Config Data.  wxConfig automatically uses the Registry on Windows and the appropriate file on Mac.
        # Program Name is Transana, Vendor Name is Verception to remain compatible with Transana 1.0.
        config = wx.Config('Transana', 'Verception')
#        self.sFTPUserName = self.editUserName.GetValue()
#        config.Write('/2.0/sFTP/sFTPUserName', self.sFTPUserName)
        self.sFTPServer = self.editsFTPServer.GetValue()
        config.Write('/2.0/sFTP/sFTPServer', self.sFTPServer)
        self.sFTPPort = self.editsFTPPort.GetValue()
        config.Write('/2.0/sFTP/sFTPPort', self.sFTPPort)
        self.sFTPPublicKeyType = self.choicesFTPPublicKeyType.GetStringSelection().strip()
        config.Write('/2.0/sFTP/sFTPPublicKeyType', self.sFTPPublicKeyType)
        self.sFTPPublicKey = self.editsFTPPublicKey.GetValue()
        config.Write('/2.0/sFTP/sFTPPublicKey', self.sFTPPublicKey)
