# Copyright (C) 2003 - 2005 The Board of Regents of the University of Wisconsin System 
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

"""This module contains miscellaneous general-purpose dialog classes.  These
are mostly intended to be sub-classed for specific uses."""

__author__ = 'Nathaniel Case <nacase@wisc.edu>, David Woods <dwoods@wcer.wisc.edu>'

import wx
from TransanaExceptions import *
import os
import copy
import string
import DBInterface
import TransanaGlobal

class ErrorDialog(wx.Dialog):
    """Error message dialog to the user."""

    def __init__(self, parent, errmsg):
        # This should be easy, right?  Just use the OS MessageDialog like so:
        # wx.MessageDialog.__init__(self, parent, errmsg, _("Transana Error"), wx.OK | wx.CENTRE | wx.ICON_ERROR)
        # That's all there is to it, right?
        #
        # Yeah, right.  Unfortunately, on Windows, this dialog isn't TRULY modal.  It's modal to the parent window
        # it's called from, but you can still select one of the other Transana Windows and do stuff.  This message
        # can even get hidden behind other windows, and cause all kinds of problems.  According to Robin Dunn,
        # writing my own class to do this is the only solution.  Here goes.

        # print "ErrorDialog", errmsg

        wx.Dialog.__init__(self, parent, -1, _("Transana Error"), size=(350, 150), style=wx.CAPTION | wx.CLOSE_BOX | wx.STAY_ON_TOP)
        # Set "Window Variant" to small only for Mac to make fonts match better
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        box = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)

        bitmap = wx.EmptyBitmap(32, 32)
        bitmap = wx.ArtProvider_GetBitmap(wx.ART_ERROR, wx.ART_MESSAGE_BOX, (32, 32))
        graphic = wx.StaticBitmap(self, -1, bitmap)

        box2.Add(graphic, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 10)
        
        message = wx.StaticText(self, -1, errmsg)

        box2.Add(message, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)
        box.Add(box2, 0, wx.EXPAND)

        btnOK = wx.Button(self, wx.ID_OK, _("OK"))

        box.Add(btnOK, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)

        self.SetAutoLayout(True)

        self.SetSizer(box)
        self.Fit()
        self.Layout()

        self.CentreOnScreen()



class InfoDialog(wx.MessageDialog):
    """Information message dialog to the user."""

    def __init__(self, parent, msg):
        # This should be easy, right?  Just use the OS MessageDialog like so:
        # wx.MessageDialog.__init__(self, parent, msg, _("Transana Information"), \
        #                     wx.OK | wx.CENTRE | wx.ICON_INFORMATION)
        # That's all there is to it, right?
        #
        # Yeah, right.  Unfortunately, on Windows, this dialog isn't TRULY modal.  It's modal to the parent window
        # it's called from, but you can still select one of the other Transana Windows and do stuff.  This message
        # can even get hidden behind other windows, and cause all kinds of problems.  According to Robin Dunn,
        # writing my own class to do this is the only solution.  Here goes.

        # print "InfoDialog", msg

        wx.Dialog.__init__(self, parent, -1, _("Transana Information"), size=(350, 150), style=wx.CAPTION | wx.CLOSE_BOX | wx.STAY_ON_TOP)

        box = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)

        bitmap = wx.EmptyBitmap(32, 32)
        bitmap = wx.ArtProvider_GetBitmap(wx.ART_INFORMATION, wx.ART_MESSAGE_BOX, (32, 32))
        graphic = wx.StaticBitmap(self, -1, bitmap)

        box2.Add(graphic, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 10)
        
        message = wx.StaticText(self, -1, msg)

        box2.Add(message, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)
        box.Add(box2, 0, wx.EXPAND)

        btnOK = wx.Button(self, wx.ID_OK, _("OK"))

        box.Add(btnOK, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)

        self.SetAutoLayout(True)

        self.SetSizer(box)
        self.Fit()
        self.Layout()

        self.CentreOnScreen()


class GenForm(wx.Dialog):
    """General dialog form used for getting basic field input."""

    def __init__(self, parent, id, title, size=(400,230), style=wx.DEFAULT_DIALOG_STYLE, HelpContext='Welcome'):
        self.width = size[0]
        self.height = size[1]
        self.HelpContext = HelpContext                                 # The HELPID for the Help Button is passed in as a parameter
        wx.Dialog.__init__(self, parent, id, title, wx.DefaultPosition,
                            wx.Size(self.width, self.height), style=style)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Due to a problem with wxWindows, Tab Order is not correct if wxTE_MULTILINE style is used in wxTextCtrls.
        # To fix this, all controls must be placed on a Panel.  This creates that Panel.
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self, wx.Top, 0)
        lay.bottom.SameAs(self, wx.Bottom, 0)
        lay.left.SameAs(self, wx.Left, 0)
        lay.right.SameAs(self, wx.Right, 0)
        self.panel = wx.Panel(self, -1, name='Dialog.Genform.Panel')
        self.panel.SetConstraints(lay)

        self.create_buttons()
        # list of (label, widget) tuples
        self.edits = []
        self.choices = []
        self.combos = []
        
    def new_edit_box(self, label, layout, def_text, style=0, ctrlHeight=40):
        """Create a edit box with label above it."""
        # The ctrlHeight parameter comes into play only when the wxTE_MULTILINE comes into play.
        # The default value of 40 is 2 text lines in height (or a little bit more.)
        txt = wx.StaticText(self.panel, -1, label)
        txt.SetConstraints(layout)

        edit = wx.TextCtrl(self.panel, -1, def_text, style=style)

        lay = wx.LayoutConstraints()
        lay.top.Below(txt, 3)
        lay.left.SameAs(txt, wx.Left)
        lay.right.SameAs(txt, wx.Right)
        if style == 0:
            lay.height.AsIs()
        else:
            lay.height.Absolute(ctrlHeight)
        edit.SetConstraints(lay)

        self.edits.append((label, edit))

        return edit
 
    def new_choice_box(self, label, layout, choices, default=0):
        """Create a choice box with label above it."""
        txt = wx.StaticText(self.panel, -1, label)
        txt.SetConstraints(layout)

        choice = wx.Choice(self.panel, -1, pos=wx.DefaultPosition, size=wx.DefaultSize, choices=choices)
        choice.SetSelection(default)
        lay = wx.LayoutConstraints()
        lay.top.Below(txt, 3)
        lay.left.SameAs(txt, wx.Left)
        lay.right.SameAs(txt, wx.Right)
        lay.height.AsIs()
        choice.SetConstraints(lay)

        self.choices.append((label, choice))

        return choice
 
    def new_combo_box(self, label, layout, choices, default='', style=wx.CB_DROPDOWN | wx.CB_SORT):
        """Create a combo box with label above it."""
        txt = wx.StaticText(self.panel, -1, label)
        txt.SetConstraints(layout)

        combo = wx.ComboBox(self.panel, -1, pos=wx.DefaultPosition, size=wx.DefaultSize, choices=choices, style=style)
        combo.SetValue(default)
        lay = wx.LayoutConstraints()
        lay.top.Below(txt, 3)
        lay.left.SameAs(txt, wx.Left)
        lay.right.SameAs(txt, wx.Right)
        lay.height.AsIs()
        combo.SetConstraints(lay)

        self.combos.append((label, combo))

        return combo
    
    def create_buttons(self, gap=10):
        """Create the dialog buttons."""
        # FIXME: Redo this using layout constraints (not urgent)
        
        # We don't want to use wx.ID_HELP here, as that causes the Help buttons to be replaced with little question
        # mark buttons on the Mac, which don't look good.
        ID_HELP = wx.NewId()

        # Tuple format = (label, id, make default?, indent from right)
        buttons = ( (_("OK"), wx.ID_OK, 1, 180),
                    (_("Cancel"), wx.ID_CANCEL, 0, 95),
                    (_("Help"), ID_HELP, 0, 10) )

       
        for b in buttons:
            lay = wx.LayoutConstraints()
            lay.height.AsIs()
            lay.bottom.SameAs(self.panel, wx.Bottom, 10)
            lay.right.SameAs(self.panel, wx.Right, b[3])
            lay.width.AsIs()
            button = wx.Button(self.panel, b[1], b[0])
            button.SetConstraints(lay)
            if b[2]:
                button.SetDefault()

        wx.EVT_BUTTON(self, ID_HELP, self.OnHelp)


    def get_input(self):
        """Show the dialog and return the inputed data.  Result is None if
        user pressed the Cancel button."""
        # Do Layout for the Panel
        self.panel.Layout()
        # Tell the Panel to AutoLayout
        self.panel.SetAutoLayout(True)
        # Do Layout for the Dialog Box (sizes the Panel)
        self.Layout()
        # Tell the Dialog Box to AutoLayout
        self.SetAutoLayout(True)
        
        val = self.ShowModal()
        if val == wx.ID_OK:
            data = {}
            for edit in self.edits:
                # Convert to string in case it's Unicode, since we use
                # some modules that aren't Unicode-aware
                data[edit[0]] = str(edit[1].GetValue())

            for choice in self.choices:
                data[choice[0]] = choice[1].GetSelection()

            for combo in self.combos:
                data[combo[0]] = combo[1].GetValue()

            return data
        else:
            return None     # Cancel

    def OnHelp(self, evt):
        """ Method to use when the Help Button is pressed """
        if TransanaGlobal.menuWindow != None:
            TransanaGlobal.menuWindow.ControlObject.Help(self.HelpContext)

    def layout_clone(self, layout):
        """Return a copy of a LayoutConstraints."""
        new = copy.deepcopy(layout)
        #new = wx.LayoutConstraints()
        
        #new.left = None
        
        #(new.bottom, new.centreX, new.centreY, new.height,
        #new.left, new.right, new.top, new.width) = \
        #(layout.bottom, layout.centreX, layout.centreY, layout.height,
        #layout.left, layout.right, layout.top, layout.width)

        return new

        

########################################
# Miscellaneous dialog functions
########################################

def add_kw_group_ui(parent, kw_groups):
    """User interface dialog and logic for adding a new keyword group.
    Return the name of the new keyword group to add, or None if cancelled."""
    s = ""
    ok = 0
    while not ok:
        s = string.strip(wx.GetTextFromUser(_("New Keyword Group:"), _("Add Keyword Group"), s))
   
        if string.find(s, ":") > -1:
            msg = _('You may not use a colon (":") in the Keyword Group name.')
            dlg = ErrorDialog(parent, msg)
            dlg.ShowModal()
            dlg.Destroy()
        elif kw_groups.count(s) != 0:
            msg = _('A Keyword Group by that name already exists.')
            dlg = ErrorDialog(parent, msg)
            dlg.ShowModal()
            dlg.Destroy()
        else:
            ok = 1
             
    # If the user cancelled
    if s == "":
        return None
    else:
        return s
