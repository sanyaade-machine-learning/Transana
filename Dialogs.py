# Copyright (C) 2003 - 2009 The Board of Regents of the University of Wisconsin System 
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

        # Make the OK button the default
        self.SetDefaultItem(btnOK)

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

        # Make the OK button the default
        self.SetDefaultItem(btnOK)

        self.SetAutoLayout(True)

        self.SetSizer(box)
        self.Fit()
        self.Layout()

        self.CentreOnScreen()

class QuestionDialog(wx.MessageDialog):
    """Replacement for wxMessageDialog with style=wx.YES_NO | wx.ICON_QUESTION."""

    def __init__(self, parent, msg, header=_("Transana Confirmation"), noDefault=False, useOkCancel=False):
        """ QuestionDialog Parameters:
                parent        Parent Window
                msg           Message to display
                header        Dialog box header, "Transana Confirmation by default
                noDefault     Set the No or Cancel button as the default, instead of Yes or OK
                useOkCancel   Use OK / Cancel as the button labels rather than Yes / No """
        
        # This should be easy, right?  Just use the OS MessageDialog like so:
        # wx.MessageDialog.__init__(self, parent, msg, _("Transana Information"), \
        #                     wx.OK | wx.CENTRE | wx.ICON_INFORMATION)
        # That's all there is to it, right?
        #
        # Yeah, right.  Unfortunately, on Windows, the MessageDialog isn't TRULY modal.  It's modal to the parent window
        # it's called from, but you can still select one of the other Transana Windows and do stuff.  This message
        # can even get hidden behind other windows, and cause all kinds of problems.  According to Robin Dunn,
        # writing my own class to do this is the only solution.

        # Set the default result to indicate failure
        self.result = -1
        # Remember the noDefault setting
        self.noDefault = noDefault
        # Define the default Window style
        dlgStyle = wx.CAPTION | wx.CLOSE_BOX | wx.STAY_ON_TOP
        
        # Create a small dialog box
        wx.Dialog.__init__(self, parent, -1, header, size=(350, 150), style=dlgStyle)
        # Create a main vertical sizer
        box = wx.BoxSizer(wx.VERTICAL)
        # Create a horizontal sizer for the first row
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        # Create a horizontal sizer for the buttons
        boxButtons = wx.BoxSizer(wx.HORIZONTAL)

        # Create an empty bitmap for the question mark graphic
        bitmap = wx.EmptyBitmap(32, 32)
        # Get the Question mark graphic and put it in the bitmap
        bitmap = wx.ArtProvider_GetBitmap(wx.ART_QUESTION, wx.ART_MESSAGE_BOX, (32, 32))
        # Create a bitmap screen object for the graphic
        graphic = wx.StaticBitmap(self, -1, bitmap)
        # Add the graphic to the first row horizontal sizer
        box2.Add(graphic, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 10)

        # Create a text screen object for the dialog text
        message = wx.StaticText(self, -1, msg)
        # Add it to the first row sizer
        box2.Add(message, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)
        # Add the first row to the main sizer
        box.Add(box2, 0, wx.EXPAND)

        # Determine the appropriate text and ID values for the buttons.
        # if useOkCancel is True ...
        if useOkCancel:
            # ... set the buttons to OK and Cancel
            btnYesText = _("OK")
            btnYesID = wx.ID_OK
            btnNoText = _("Cancel")
            btnNoID = wx.ID_CANCEL
        # If useOkCancel is False (the default) ...
        else:
            # ... set the buttons to Yes and No
            btnYesText = _("&Yes")
            btnYesID = wx.ID_YES
            btnNoText = _("&No")
            btnNoID = wx.ID_NO
        # Create the first button, which is Yes or OK
        btnYes = wx.Button(self, btnYesID, btnYesText)
        # Bind the button event to its method
        btnYes.Bind(wx.EVT_BUTTON, self.OnButton)
        # Create the second button, which is No or Cancel
        self.btnNo = wx.Button(self, btnNoID, btnNoText)
        # Bind the button event to its method
        self.btnNo.Bind(wx.EVT_BUTTON, self.OnButton)
        # Add an expandable spacer to the button sizer
        boxButtons.Add((20,1), 1)
        # If we're on the Mac, we want No/Cancel then Yes/OK
        if "__WXMAC__" in wx.PlatformInfo:
            # Add No first
            boxButtons.Add(self.btnNo, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
            # Add a spacer
            boxButtons.Add((20,1))
            # Then add Yes
            boxButtons.Add(btnYes, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        # If we're not on the Mac, we want Yes/OK then No/Cancel
        else:
            # Add Yes first
            boxButtons.Add(btnYes, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
            # Add a spacer
            boxButtons.Add((20,1))
            # Then add No
            boxButtons.Add(self.btnNo, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        # Add a final expandable spacer
        boxButtons.Add((20,1), 1)
        # Add the button bar to the main sizer
        box.Add(boxButtons, 0, wx.ALIGN_CENTER | wx.EXPAND)
        # Turn AutoLayout On
        self.SetAutoLayout(True)
        # Set the form's main sizer
        self.SetSizer(box)
        # Fit the form
        self.Fit()
        # Lay the form out
        self.Layout()
        # Center the form on screen
        self.CentreOnScreen()

    def OnButton(self, event):
        """ Button Event Handler """
        # Set the result variable to the ID of the button that was pressed
        self.result = event.GetId()
        # Close the form
        self.Close()

    def LocalShowModal(self):
        """ ShowModal wasn't working right.  This allows some modification. """
        # If No/Cancel should be the default button ...
        if self.noDefault:
            # ... set the focus to the No button
            self.btnNo.SetFocus()
        # Show the form modally
        self.ShowModal()
        # Return the result to indicate what button was pressed
        return self.result


class GenForm(wx.Dialog):
    """General dialog form used for getting basic field input."""

    def __init__(self, parent, id, title, size=(400,230), style=wx.DEFAULT_DIALOG_STYLE, propagateEnabled=False, HelpContext='Welcome'):
        self.width = size[0]
        self.height = size[1]
        # Remember if the Propagate Button should be displayed
        self.propagateEnabled = propagateEnabled
        # The HELPID for the Help Button is passed in as a parameter
        self.HelpContext = HelpContext
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

    def new_edit_box(self, label, layout, def_text, style=0, ctrlHeight=40, maxLen=0):
        """Create a edit box with label above it."""
        # The ctrlHeight parameter comes into play only when the wxTE_MULTILINE comes into play.
        # The default value of 40 is 2 text lines in height (or a little bit more.)
        txt = wx.StaticText(self.panel, -1, label)
        txt.SetConstraints(layout)

        edit = wx.TextCtrl(self.panel, -1, def_text, style=style)
        # If a maximum length is specified, apply it.
        if maxLen > 0:
            edit.SetMaxLength(maxLen)

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
        # We don't want to use wx.ID_HELP here, as that causes the Help buttons to be replaced with little question
        # mark buttons on the Mac, which don't look good.
        ID_HELP = wx.NewId()
        # If the Propagation feature is enabled ...
        if self.propagateEnabled:
            # ... load the propagate button graphic
            propagateBMP = wx.Bitmap(os.path.join("images", "Propagate.xpm"), wx.BITMAP_TYPE_XPM)
            # Create a bitmap button for propagation
            self.propagate = wx.BitmapButton(self.panel, -1, propagateBMP)
            # Set the tool tip for the graphic button
            self.propagate.SetToolTipString(_("Propagate Changes"))
            # Bind the propogate button to its event
            self.propagate.Bind(wx.EVT_BUTTON, self.OnPropagate)
            # Position the button on screen
            lay = wx.LayoutConstraints()
            lay.height.AsIs()
            lay.bottom.SameAs(self.panel, wx.Bottom, 10)
            lay.right.SameAs(self.panel, wx.Right, 265)
            lay.width.AsIs()
            self.propagate.SetConstraints(lay)

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
        # Initialize the propagation flag for False
        self.propagatePressed = False
        # Show the Properties Dialog modally and get user feedback
        val = self.ShowModal()
        # If the user presses OK OR presses the propagate button, we save the entered data to Form properties.
        if (val == wx.ID_OK) or self.propagatePressed:
            data = {}
            for edit in self.edits:
                if 'ansi' in wx.PlatformInfo:
                    # Convert to string in case it's Unicode, since we use
                    # some modules that aren't Unicode-aware
                    data[edit[0]] = str(edit[1].GetValue())
                else:
                    data[edit[0]] = edit[1].GetValue()

            for choice in self.choices:
                data[choice[0]] = choice[1].GetSelection()

            for combo in self.combos:
                data[combo[0]] = combo[1].GetValue()
            # While we return the data dictionary, that data isn't actually readable from the calling
            # routine.  I don't know exactly what Nate did here.
            return data
        else:
            return None     # Cancel

    def OnPropagate(self, event):
        """ Handle Clip Change Propagation """
        # Figuring out when and how to initiate the propagation process proved more challenging than I'd expected.
        # The current clip needs to get saved before propagation begins, and that save happens in the DatabaseTreeTab
        # module.  So if the user presses the Propagate button, all we do here is set the propagation flag and close
        # the dialog box.
        
        # If the user presses the propagation button, set the propagation flag to true
        self.propagatePressed = True
        # Propagation mimics OK, so close the dialog
        self.Close()
        
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

    # ALSO SEE Keyword._set_keywordGroup().  The same errors are caught there.

    # initialize local variables
    s = ""
    ok = False

    # Let's get a copy of kw_groups that's all upper case
    kw_groups_upper = []
    for kwg in kw_groups:
        kw_groups_upper.append(kwg.upper())

    # Repeat until no error is found
    while not ok:
        # Get a Keyword Group name from the user
        s = string.strip(wx.GetTextFromUser(_("New Keyword Group:"), _("Add Keyword Group"), s))

        # Check (case-insensitively) whether the Keyword Group already exists.
        if kw_groups_upper.count(s.upper()) != 0:
            msg = _('A Keyword Group by that name already exists.')
            dlg = ErrorDialog(parent, msg)
            dlg.ShowModal()
            dlg.Destroy()
        else:
            # Make sure parenthesis characters are not allowed in Keyword Group.  Remove them if necessary.
            if (s.find('(') > -1) or (s.find(')') > -1):
                s = s.replace('(', '')
                s = s.replace(')', '')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Keyword Groups cannot contain parenthesis characters.\nYour Keyword Group has been renamed to "%s".'), 'utf8')
                else:
                    prompt = _('Keyword Groups cannot contain parenthesis characters.\nYour Keyword Group has been renamed to "%s".')
                dlg = ErrorDialog(None, prompt % s)
                dlg.ShowModal()
                dlg.Destroy()
                
            # Colons are not allowed in Keyword Groups.  Remove them if necessary.
            if s.find(":") > -1:
                s = s.replace(':', '')
                if 'unicode' in wx.PlatformInfo:
                    msg = unicode(_('You may not use a colon (":") in the Keyword Group name.  Your Keyword Group has been changed to\n"%s"'), 'utf8')
                else:
                    msg = _('You may not use a colon (":") in the Keyword Group name.  Your Keyword Group has been changed to\n"%s"')
                dlg = ErrorDialog(parent, msg % s)
                dlg.ShowModal()
                dlg.Destroy()
                
            # Let's make sure we don't exceed the maximum allowed length for a Keyword Group.
            # First, let's see what the max length is.
            maxLen = TransanaGlobal.maxKWGLength
            # Check to see if we've exceeded the max length
            if len(s) > maxLen:
                # If so, truncate the Keyword Group
                s = s[:maxLen]
                # Display a message to the user describing the trunctions
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(_('Keyword Group is limited to %d characters.  Your Keyword Group has been changed to\n"%s"'), 'utf8')
                else:
                    msg = _('Keyword Group is limited to %d characters.  Your Keyword Group has been changed to\n"%s"')
                dlg = ErrorDialog(parent, msg % (maxLen, s))
                dlg.ShowModal()
                dlg.Destroy()
            # If we hit here, there's no reason to block the closing of the dialog box.  (Parens, Colons, and Length violation should not block close.)
            ok = True
             
    # If the user cancelled
    if s == "":
        return None
    else:
        return s
