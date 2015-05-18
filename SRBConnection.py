#
# Copyright (C) 2003 The Board of Regents of the University of Wisconsin System 
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
        wx.Dialog.__init__(self, parent, -1, _("SRB Connection Parameters"), size=(250, 520), style=wx.CAPTION|wx.NO_FULL_REPAINT_ON_RESIZE)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        self.LoadConfiguration()

        # Create a BoxSizer for layout
        box = wx.BoxSizer(wx.VERTICAL)
        
        #lay = wx.LayoutConstraints()
        #lay.top.SameAs(self, wx.Top, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        lblUserName = wx.StaticText(self, -1, _("User Name:"))
        #lblUserName.SetConstraints(lay)
        box.Add(lblUserName, 1, wx.ALIGN_LEFT | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(lblUserName, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editUserName = wx.TextCtrl(self, -1, self.srbUserName)
        #self.editUserName.SetConstraints(lay)
        box.Add(self.editUserName, 1, wx.ALIGN_LEFT | wx.EXPAND | wx.ALL, 3)
        
        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.editUserName, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        lblPassword = wx.StaticText(self, -1, _("Password:"))
        #lblPassword.SetConstraints(lay)
        box.Add(lblPassword, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(lblPassword, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editPassword = wx.TextCtrl(self, -1, style=wx.TE_PASSWORD)
        #self.editPassword.SetConstraints(lay)
        box.Add(self.editPassword, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)
        
        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.editPassword, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        lblDomain = wx.StaticText(self, -1, _("Domain:"))
        #lblDomain.SetConstraints(lay)
        box.Add(lblDomain, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(lblDomain, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editDomain = wx.TextCtrl(self, -1, self.srbDomain)
        #self.editDomain.SetConstraints(lay)
        box.Add(self.editDomain, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)
        
        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.editDomain, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        lblCollectionRoot = wx.StaticText(self, -1, _("Collection Root:"))
        #lblCollectionRoot.SetConstraints(lay)
        box.Add(lblCollectionRoot, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(lblCollectionRoot, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editCollectionRoot = wx.TextCtrl(self, -1, self.srbCollectionRoot)
        #self.editCollectionRoot.SetConstraints(lay)
        box.Add(self.editCollectionRoot, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)
        
        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.editCollectionRoot, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        lblSRBHost = wx.StaticText(self, -1, _("SRB Host:"))
        #lblSRBHost.SetConstraints(lay)
        box.Add(lblSRBHost, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(lblSRBHost, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editSRBHost = wx.TextCtrl(self, -1, self.srbHost)
        #self.editSRBHost.SetConstraints(lay)
        box.Add(self.editSRBHost, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)
        
        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.editSRBHost, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        lblSRBPort = wx.StaticText(self, -1, _("SRB Port:"))
        #lblSRBPort.SetConstraints(lay)
        box.Add(lblSRBPort, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(lblSRBPort, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editSRBPort = wx.TextCtrl(self, -1, self.srbPort)
        #self.editSRBPort.SetConstraints(lay)
        box.Add(self.editSRBPort, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)
        
        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.editSRBPort, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        lblResource = wx.StaticText(self, -1, _("Resource:"))
        #lblResource.SetConstraints(lay)
        box.Add(lblResource, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(lblResource, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editSRBResource = wx.TextCtrl(self, -1, self.srbResource)
        #self.editSRBResource.SetConstraints(lay)
        box.Add(self.editSRBResource, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)
        
        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.editSRBResource, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        lblSRBSEAOption = wx.StaticText(self, -1, _("SEA Option:"))
        #lblSRBSEAOption.SetConstraints(lay)
        box.Add(lblSRBSEAOption, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(lblSRBSEAOption, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editSRBSEAOption = wx.TextCtrl(self, -1, self.srbSEAOption)
        #self.editSRBSEAOption.SetConstraints(lay)
        box.Add(self.editSRBSEAOption, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)
        
        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.editSRBSEAOption, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        self.rbSRBStorageSpace = wx.RadioBox(self, RB_STORAGESPACE, _("Connect to:"), choices=[" " + _("My own storage space"), " " + _("Another user's storage space") + "                  "], majorDimension=2, style=wx.RA_SPECIFY_ROWS)
        #self.rbSRBStorageSpace.SetConstraints(lay)
        box.Add(self.rbSRBStorageSpace, 2, wx.ALIGN_LEFT |  wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.rbSRBStorageSpace, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        self.lblOtherUserName = wx.StaticText(self, -1, _("User Name to connect to:"))
        #self.lblOtherUserName.SetConstraints(lay)
        self.lblOtherUserName.Show(False)
        box.Add(self.lblOtherUserName, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.lblOtherUserName, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editOtherUserName = wx.TextCtrl(self, -1)
        #self.editOtherUserName.SetConstraints(lay)
        self.editOtherUserName.Show(False)
        box.Add(self.editOtherUserName, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)
        
        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.editOtherUserName, 10)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        self.lblOtherDomain = wx.StaticText(self, -1, _("Domain to connect to:"))
        #self.lblOtherDomain.SetConstraints(lay)
        self.lblOtherDomain.Show(False)
        box.Add(self.lblOtherDomain, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.Below(self.lblOtherDomain, 5)
        #lay.left.SameAs(self, wx.Left, 10)
        #lay.height.AsIs()
        #lay.right.SameAs(self, wx.Right, 10)
        self.editOtherDomain = wx.TextCtrl(self, -1, self.srbDomain)
        #self.editOtherDomain.SetConstraints(lay)
        self.editOtherDomain.Show(False)
        box.Add(self.editOtherDomain, 1, wx.ALIGN_LEFT |  wx.EXPAND | wx.ALL, 3)

        box2 = wx.BoxSizer(wx.HORIZONTAL)
        
        # Add a spacer to the left side of box2
        box2.Add((100, 1))
                
        #lay = wx.LayoutConstraints()
        #lay.top.SameAs(self, wx.Bottom, -25)
        #lay.right.SameAs(self, wx.CentreX, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        btnConnect = wx.Button(self, wx.ID_OK, _("Connect"))
        #btnConnect.SetConstraints(lay)
        box2.Add(btnConnect, 1, wx.ALIGN_RIGHT | wx.ALL, 3)

        #lay = wx.LayoutConstraints()
        #lay.top.SameAs(self, wx.Bottom, -25)
        #lay.left.SameAs(self, wx.CentreX, 10)
        #lay.height.AsIs()
        #lay.width.AsIs()
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        #btnCancel.SetConstraints(lay)
        box2.Add(btnCancel, 1, wx.ALIGN_RIGHT | wx.ALL, 3)

        box.Add(box2, 1, wx.ALIGN_RIGHT | wx.ALL, 3)

        self.SetSizer(box)
        self.Fit()
        self.Layout()
        self.SetAutoLayout(True)

        # Set Focus
        self.editUserName.SetFocus()

        wx.EVT_RADIOBOX(self, RB_STORAGESPACE, self.onStorageSpaceSelect)

    def onStorageSpaceSelect(self, event):
        (x, y, width, height) = self.GetRect()
        if self.rbSRBStorageSpace.GetSelection() == 1:
            self.SetDimensions(x, y - 49, width, height + 98)
            self.lblOtherUserName.Show(True)
            self.editOtherUserName.Show(True)
            self.lblOtherDomain.Show(True)
            self.editOtherDomain.Show(True)
        else:
            self.SetDimensions(x, y + 49, width, height - 98)
            self.lblOtherUserName.Show(False)
            self.editOtherUserName.Show(False)
            self.lblOtherDomain.Show(False)
            self.editOtherDomain.Show(False)
        self.Fit()

    def LoadConfiguration(self):
        """ Load Configuration Data from the Registry or Config File """
        # Set Default Values
        defaultsrbDomain         = 'digital-insight'
        defaultsrbCollectionRoot = '/WCER/home/'       # '/home/'
        defaultsrbHost           = 'storage.wcer.wisc.edu'    # 'srb.wcer.wisc.edu'
        defaultsrbPort           = '5544'              # '5823'
        defaultsrbResource       = 'WCERSRBV1'         # 'nt-wcersrb-1'
        defaultsrbSEAOption      = 'ENCRYPT1'
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
