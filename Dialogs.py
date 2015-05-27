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

"""This module contains miscellaneous general-purpose dialog classes.  These
are mostly intended to be sub-classed for specific uses."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case <nacase@wisc.edu>'

import wx
from TransanaExceptions import *
import os
import copy
import string
import DBInterface
import TransanaGlobal
# Import Transana's Images
import TransanaImages

class ErrorDialog(wx.Dialog):
    """Error message dialog to the user."""

    def __init__(self, parent, errmsg, includeSkipCheck=False):
        # This should be easy, right?  Just use the OS MessageDialog like so:
        # wx.MessageDialog.__init__(self, parent, errmsg, _("Transana Error"), wx.OK | wx.CENTRE | wx.ICON_ERROR)
        # That's all there is to it, right?
        #
        # Yeah, right.  Unfortunately, on Windows, this dialog isn't TRULY modal.  It's modal to the parent window
        # it's called from, but you can still select one of the other Transana Windows and do stuff.  This message
        # can even get hidden behind other windows, and cause all kinds of problems.  According to Robin Dunn,
        # writing my own class to do this is the only solution.  Here goes.

        # Remember if we're supposed to include the checkbox to skip additional messages
        self.includeSkipCheck = includeSkipCheck

        # Create a Dialog box
        wx.Dialog.__init__(self, parent, -1, _("Transana Error"), size=(350, 150), style=wx.CAPTION | wx.CLOSE_BOX | wx.STAY_ON_TOP)
        # Set "Window Variant" to small only for Mac to make fonts match better
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create Vertical and Horizontal Sizers
        box = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)

        # Display the Error graphic in the dialog box
        graphic = wx.StaticBitmap(self, -1, TransanaImages.ArtProv_ERROR.GetBitmap())
        # Add the graphic to the Sizers
        box2.Add(graphic, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 10)

        # Display the error message in the dialog box
        message = wx.StaticText(self, -1, errmsg)
        # Add the error message to the Sizers
        box2.Add(message, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)
        box.Add(box2, 0, wx.EXPAND)

        # If we should add the "Skip Further messages" checkbox ...
        if self.includeSkipCheck:
            # ... then add the checkbox to the dialog
            self.skipCheck = wx.CheckBox(self, -1, _('Do not show this message again'))
            # ... and add the checkbox to the Sizers
            box.Add(self.skipCheck, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        
        # Add an OK button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Add the OK button to the Sizers
        box.Add(btnOK, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        # Make the OK button the default
        self.SetDefaultItem(btnOK)

        # Turn on Auto Layout
        self.SetAutoLayout(True)
        # Set the form sizer
        self.SetSizer(box)
        # Set the form size
        self.Fit()
        # Perform the Layout
        self.Layout()
        # Center the dialog on the screen
#        self.CentreOnScreen()
        # That's not working.  Let's try this ...
        TransanaGlobal.CenterOnPrimary(self)

    def GetSkipCheck(self):
        """ Provide the value of the Skip Checkbox """
        # If the checkbox is displayed ...
        if self.includeSkipCheck:
            # ... return the value of the checkbox
            return self.skipCheck.GetValue()
        # If the checkbox is NOT displayed ...
        else:
            # ... then indicate that it is NOT checked
            return False


class InfoDialog(wx.MessageDialog):
    """ Information message dialog to the user. """

    def __init__(self, parent, msg, dlgTitle = _("Transana Information")):
        # This should be easy, right?  Just use the OS MessageDialog like so:
        # wx.MessageDialog.__init__(self, parent, msg, _("Transana Information"), \
        #                     wx.OK | wx.CENTRE | wx.ICON_INFORMATION)
        # That's all there is to it, right?
        #
        # Yeah, right.  Unfortunately, on Windows, this dialog isn't TRULY modal.  It's modal to the parent window
        # it's called from, but you can still select one of the other Transana Windows and do stuff.  This message
        # can even get hidden behind other windows, and cause all kinds of problems.  According to Robin Dunn,
        # writing my own class to do this is the only solution.  Here goes.

        # Create a wxDialog
        wx.Dialog.__init__(self, parent, -1, dlgTitle, size=(350, 150), style=wx.CAPTION | wx.CLOSE_BOX | wx.STAY_ON_TOP)

        # Create Vertical and Horizontal Sizers
        box = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)

        # Display the Information graphic in the dialog box
        graphic = wx.StaticBitmap(self, -1, TransanaImages.ArtProv_INFORMATION.GetBitmap())
        # Add the graphic to the Sizers
        box2.Add(graphic, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 10)
        
        # Display the error message in the dialog box
        message = wx.StaticText(self, -1, msg)

        # Add the information message to the Sizers
        box2.Add(message, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)
        box.Add(box2, 0, wx.EXPAND)

        # Add an OK button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Add the OK button to the Sizers
        box.Add(btnOK, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        # Make the OK button the default
        self.SetDefaultItem(btnOK)

        # Turn on Auto Layout
        self.SetAutoLayout(True)
        # Set the form sizer
        self.SetSizer(box)
        # Set the form size
        self.Fit()
        # Perform the Layout
        self.Layout()
        # Center the dialog on the screen
#        self.CentreOnScreen()
        # That's not working.  Let's try this ...
        TransanaGlobal.CenterOnPrimary(self)


class QuestionDialog(wx.MessageDialog):
    """Replacement for wxMessageDialog with style=wx.YES_NO | wx.ICON_QUESTION."""

    def __init__(self, parent, msg, header=_("Transana Confirmation"), noDefault=False, useOkCancel=False, yesToAll=False, includeEncoding=False):
        """ QuestionDialog Parameters:
                parent           Parent Window
                msg              Message to display
                header           Dialog box header, "Transana Confirmation" by default
                noDefault        Set the No or Cancel button as the default, instead of Yes or OK
                useOkCancel      Use OK / Cancel as the button labels rather than Yes / No
                yesToAll         Include the "Yes to All" option
                includeEncoding  Include Encoding selection for upgrading the single-user Windows database for 2.50 """
        
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
        # Remember if Encoding is included
        self.includeEncoding = includeEncoding
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
        # Create a bitmap screen object for the graphic
        graphic = wx.StaticBitmap(self, -1, TransanaImages.ArtProv_QUESTION.GetBitmap())
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

        # If the "Yes to All" option is enabled
        if yesToAll:
            # ... create a YesToAll ID
            self.YESTOALLID = wx.NewId()
            # Create a Yes To All button
            self.btnYesToAll = wx.Button(self, self.YESTOALLID, _("Yes to All"))
            # Bind the button handler for the Yes To All button
            self.btnYesToAll.Bind(wx.EVT_BUTTON, self.OnButton)
        
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
            # If the "Yes to All" option is enabled
            if yesToAll:
                # Add No first
                boxButtons.Add(self.btnYesToAll, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
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
            # If the "Yes to All" option is enabled
            if yesToAll:
                # Add No first
                boxButtons.Add(self.btnYesToAll, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
                # Add a spacer
                boxButtons.Add((20,1))
            # Then add No
            boxButtons.Add(self.btnNo, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        # Add a final expandable spacer
        boxButtons.Add((20,1), 1)
        # Add the button bar to the main sizer
        box.Add(boxButtons, 0, wx.ALIGN_CENTER | wx.EXPAND)

        # If we're supposed to include the Encoding prompt ...
        if includeEncoding:
            # Create a horizontal sizer for the first row
            box3 = wx.BoxSizer(wx.HORIZONTAL)
            # Define the options for the Encoding choice box
            choices = ['',
                       _('Most languages from single-user Transana 2.1 - 2.42 on Windows'),
                       _('Chinese from single-user Transana 2.1 - 2.42 on Windows'),
                       _('Russian from single-user Transana 2.1 - 2.42 on Windows'),
                       _('Eastern European data from single-user Transana 2.1 - 2.42 on Windows'),
                       _('Greek data from single-user Transana 2.1 - 2.42 on Windows'),
                       _('Japanese data from single-user Transana 2.1 - 2.42 on Windows'),
                       _("All languages from OS X or MU database files")]
            # Use a dictionary to define the encodings that go with each of the Encoding options
            self.encodingOptions = {'' : None,
                                    _('Most languages from single-user Transana 2.1 - 2.42 on Windows') : 'latin1',
                                    _('Chinese from single-user Transana 2.1 - 2.42 on Windows') : 'gbk',
                                    _('Russian from single-user Transana 2.1 - 2.42 on Windows') : 'koi8_r',
                                    _('Eastern European data from single-user Transana 2.1 - 2.42 on Windows') : 'iso8859_2',
                                    _('Greek data from single-user Transana 2.1 - 2.42 on Windows') : 'iso8859_7',
                                    _('Japanese data from single-user Transana 2.1 - 2.42 on Windows') : 'cp932',
                                    _("All languages from OS X or MU database files") : 'utf8'}
            # Create a Choice Box where the user can select an import encoding, based on information about how the
            # Transana-XML file in question was created.  This adds it to the Vertical Sizer created above.
            chLbl = wx.StaticText(self, -1, _('Language Option used:'))
            box3.Add(chLbl, 0, wx.LEFT | wx.TOP | wx.BOTTOM, 10)
            self.chImportEncoding = wx.Choice(self, -1, choices=choices)
            self.chImportEncoding.SetSelection(0)
            box3.Add(self.chImportEncoding, 1, wx.EXPAND| wx.ALIGN_CENTER | wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)
            box.Add(box3, 0, wx.EXPAND)

        # Turn AutoLayout On
        self.SetAutoLayout(True)
        # Set the form's main sizer
        self.SetSizer(box)
        # Fit the form
        self.Fit()
        # Lay the form out
        self.Layout()
        # Center the form on screen
#        self.CentreOnScreen()
        # That's not working.  Let's try this ...
        TransanaGlobal.CenterOnPrimary(self)

    def OnButton(self, event):
        """ Button Event Handler """
        # Set the result variable to the ID of the button that was pressed
        self.result = event.GetId()
        # If we need an encoding and none has been selected and the user has selected Yes ...
        if self.includeEncoding and (self.chImportEncoding.GetSelection() == 0) and (self.result == wx.ID_YES):
            # then we need to over-ride the button press!  First, report the problem to the user.
            tmpDlg = ErrorDialog(self, _('You must select a language option to proceed.'))
            tmpDlg.ShowModal()
            tmpDlg.Destroy()
        # If we have a valid form ...
        else:
            # ... close the form
            self.EndModal(self.result)

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

class PopupDialog(wx.Dialog):
    """ A popup dialog for temporary user messages """

    def __init__(self, parent, title, msg):
        # Create a dialog
        wx.Dialog.__init__(self, parent, -1, title, size=(350, 150), style=wx.CAPTION | wx.STAY_ON_TOP)
        # Add sizers
        box = wx.BoxSizer(wx.VERTICAL)
        box2 = wx.BoxSizer(wx.HORIZONTAL)
        # Add an Info graphic
        graphic = wx.StaticBitmap(self, -1, TransanaImages.ArtProv_INFORMATION.GetBitmap())
        box2.Add(graphic, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 10)
        # Add the message
        message = wx.StaticText(self, -1, msg)
        box2.Add(message, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)
        box.Add(box2, 0, wx.EXPAND)
        # Handle layout
        self.SetAutoLayout(True)
        self.SetSizer(box)
        self.Fit()
        self.Layout()
#        self.CentreOnScreen()
        TransanaGlobal.CenterOnPrimary(self)
        # Display the Dialog
        self.Show()
        # Make sure the screen gets fully drawn before continuing.  (Needed for SAVE)
        try:
            wx.Yield()
        except:
            pass

class GenForm(wx.Dialog):
    """General dialog form used for getting basic field input."""

    def __init__(self, parent, id, title, size=(400,230), style=wx.DEFAULT_DIALOG_STYLE, propagateEnabled=False, useSizers=False, HelpContext='Welcome'):
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

        # Create a panel to hold the form contents.  This ensures that Tab Order works properly.
        self.panel = wx.Panel(self, -1, name='Dialog.Genform.Panel')
        # If we are using Sizers (and we are transitioning to Sizers)
        if useSizers:
            # Create a vertical sizer
            vSizer = wx.BoxSizer(wx.VERTICAL)
            # Create a horizontal sizer
            hSizer = wx.BoxSizer(wx.HORIZONTAL)
            # Add the panel to the horizontal sizer to expand horizontally
            hSizer.Add(self.panel, 1, wx.EXPAND)
            # Add the horizontal sizer to the vertical sizer to expand vertically.  Give it a 10 point margin all around.
            vSizer.Add(hSizer, 1, wx.EXPAND | wx.ALL, 10)
            # Set the vertical sizer as the form's main sizer.
            self.SetSizer(vSizer)
        # If we're using LayoutConstraints (which we're transitioning away from) 
        else:
            # Create the Layout Constraint
            lay = wx.LayoutConstraints()
            # Set all margins to 0
            lay.top.SameAs(self, wx.Top, 0)
            lay.bottom.SameAs(self, wx.Bottom, 0)
            lay.left.SameAs(self, wx.Left, 0)
            lay.right.SameAs(self, wx.Right, 0)
            # Add these layout constraints to the panel.  (The items on the panel implement form margins under Layout Constaints.)
            self.panel.SetConstraints(lay)

            # We can create the buttons here with Layout Constraints, but not Sizers.
            self.create_buttons()
            
        # list of (label, widget) tuples
        self.edits = []
        self.choices = []
        self.combos = []

    def new_edit_box(self, label, layout, def_text, style=0, ctrlHeight=40, maxLen=0):
        """Create a edit box with label above it."""
        # The ctrlHeight parameter comes into play only when the wxTE_MULTILINE is used.
        # The default value of 40 is 2 text lines in height (or a little bit more.)

        # Create the text box's label
        txt = wx.StaticText(self.panel, -1, label)
        # If we're using LayoutContraints ...
        if isinstance(layout, wx.LayoutConstraints):
            # ... apply the layout constraints that were passed in.
            txt.SetConstraints(layout)
        # If we're using Sizers ...
        elif isinstance(layout, wx.BoxSizer):
            # ... add the text label to the sizer that was passed in.
            layout.Add(txt, 0, wx.BOTTOM, 3)

        # Create the Text Control
        edit = wx.TextCtrl(self.panel, -1, def_text, style=style)
        # If a maximum length is specified, apply it.  (I don't think this is supported under unicode!!)
        if maxLen > 0:
            edit.SetMaxLength(maxLen)

        # If we're using LayoutContraints ...
        if isinstance(layout, wx.LayoutConstraints):
            # ... create a new layout constraint just below the one passed in using its values
            lay = wx.LayoutConstraints()
            lay.top.Below(txt, 3)
            lay.left.SameAs(txt, wx.Left)
            lay.right.SameAs(txt, wx.Right)
            if style == 0:
                lay.height.AsIs()
            else:
                lay.height.Absolute(ctrlHeight)
            # Apply the layout constraint to the text control.
            edit.SetConstraints(lay)
        # If we're using Sizers ...
        elif isinstance(layout, wx.BoxSizer):
            # ... add the text control to the sizer.
            layout.Add(edit, 1, wx.EXPAND)

        # Add this text control to the list of edit controls
        self.edits.append((label, edit))
        # Return the text control
        return edit
 
    def new_choice_box(self, label, layout, choices, default=0):
        """Create a choice box with label above it."""
        txt = wx.StaticText(self.panel, -1, label)

        # If we're using LayoutContraints ...
        if isinstance(layout, wx.LayoutConstraints):
            # ... apply the layout constraints that were passed in.
            txt.SetConstraints(layout)
        # If we're using Sizers ...
        elif isinstance(layout, wx.BoxSizer):
            # ... add the text label to the sizer that was passed in.
            layout.Add(txt, 0, wx.BOTTOM, 3)

        choice = wx.Choice(self.panel, -1, pos=wx.DefaultPosition, size=wx.DefaultSize, choices=choices)
        choice.SetSelection(default)

        # If we're using LayoutContraints ...
        if isinstance(layout, wx.LayoutConstraints):
            # ... create a new layout constraint just below the one passed in using its values
            lay = wx.LayoutConstraints()
            lay.top.Below(txt, 3)
            lay.left.SameAs(txt, wx.Left)
            lay.right.SameAs(txt, wx.Right)
            lay.height.AsIs()
            # Apply the layout constraint to the choice control.
            choice.SetConstraints(lay)
        # If we're using Sizers ...
        elif isinstance(layout, wx.BoxSizer):
            # ... add the choice control to the sizer.
            layout.Add(choice, 1, wx.EXPAND)

        # Add this choice control to the list of choice controls
        self.choices.append((label, choice))
        # Return the choice control
        return choice
 
    def new_combo_box(self, label, layout, choices, default='', style=None):
        """Create a combo box with label above it."""
        if style == None:
            # As of wxPython 2.9.5.0, Mac doesn't support wx.CB_SORT and gives an ugly message about it!
            if 'wxMac' in wx.PlatformInfo:
                style = wx.CB_DROPDOWN
                choices.sort()
            else:
                style = wx.CB_DROPDOWN | wx.CB_SORT

        txt = wx.StaticText(self.panel, -1, label)

        # If we're using LayoutContraints ...
        if isinstance(layout, wx.LayoutConstraints):
            # ... apply the layout constraints that were passed in.
            txt.SetConstraints(layout)
        # If we're using Sizers ...
        elif isinstance(layout, wx.BoxSizer):
            # ... add the text label to the sizer that was passed in.
            layout.Add(txt, 0, wx.BOTTOM, 3)

        combo = wx.ComboBox(self.panel, -1, pos=wx.DefaultPosition, size=wx.DefaultSize, choices=choices, style=style)
        combo.SetValue(default)

        # If we're using LayoutContraints ...
        if isinstance(layout, wx.LayoutConstraints):
            # ... create a new layout constraint just below the one passed in using its values
            lay = wx.LayoutConstraints()
            lay.top.Below(txt, 3)
            lay.left.SameAs(txt, wx.Left)
            lay.right.SameAs(txt, wx.Right)
            lay.height.AsIs()
            # Apply the layout constraint to the combo control.
            combo.SetConstraints(lay)
        # If we're using Sizers ...
        elif isinstance(layout, wx.BoxSizer):
            # ... add the combo control to the sizer.
            layout.Add(combo, 1, wx.EXPAND)

        # Add this combo control to the list of combo controls
        self.combos.append((label, combo))
        # Return the combo control
        return combo
    
    def create_buttons(self, gap=10, sizer=None):
        """Create the dialog buttons."""
        # We don't want to use wx.ID_HELP here, as that causes the Help buttons to be replaced with little question
        # mark buttons on the Mac, which don't look good.
        ID_HELP = wx.NewId()

        # If we're using a Sizer (not LayoutConstraints)
        if sizer:
            # ... add an expandable spacer here so buttons will be right-justified
            sizer.Add((1, 0), 1, wx.EXPAND)

        # If the Propagation feature is enabled ...
        if self.propagateEnabled:
            # ... load the propagate button graphic
            propagateBMP = TransanaImages.Propagate.GetBitmap()
            # Create a bitmap button for propagation
            self.propagate = wx.BitmapButton(self.panel, -1, propagateBMP)
            # Set the tool tip for the graphic button
            self.propagate.SetToolTipString(_("Propagate Changes"))
            # Bind the propogate button to its event
            self.propagate.Bind(wx.EVT_BUTTON, self.OnPropagate)

            # If we're using Layout Constraints ...
            if sizer == None:
                # Position the Propagate button on screen
                lay = wx.LayoutConstraints()
                lay.height.AsIs()
                lay.bottom.SameAs(self.panel, wx.Bottom, 10)
                lay.right.SameAs(self.panel, wx.Right, 265)
                lay.width.AsIs()
                self.propagate.SetConstraints(lay)
            # If we're using Sizers ...
            else:
                # ... add the propagate button to the sizer
                sizer.Add(self.propagate, 0, wx.ALIGN_RIGHT | wx.EXPAND)

        # Define the standard buttons we want to create.
        # Tuple format = (label, id, make default?, indent from right)
        buttons = ( (_("OK"), wx.ID_OK, 1, 180),
                    (_("Cancel"), wx.ID_CANCEL, 0, 95),
                    (_("Help"), ID_HELP, 0, gap) )

       
        for b in buttons:
            # If we're using Layout Constraints:
            if sizer == None:
                lay = wx.LayoutConstraints()
                lay.height.AsIs()
                lay.bottom.SameAs(self.panel, wx.Bottom, 10)
                lay.right.SameAs(self.panel, wx.Right, b[3])
                lay.width.AsIs()
            button = wx.Button(self.panel, b[1], b[0])
            # If we're using Layout Constraints:
            if sizer == None:
                button.SetConstraints(lay)
            else:
                sizer.Add(button, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.EXPAND, gap)

            if b[2]:
                button.SetDefault()
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            sizer.Add((2, 0))

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

class add_kw_group_ui(wx.Dialog):
    """ User interface dialog and logic for adding a new keyword group. """

    def __init__(self, parent, kw_groups):

        # ALSO SEE Keyword._set_keywordGroup().  The same errors are caught there.

        # Let's get a copy of kw_groups that's all upper case
        self.kw_groups_upper = []
        for kwg in kw_groups:
            self.kw_groups_upper.append(kwg.upper())

        # Define the default Window style
        dlgStyle = wx.CAPTION | wx.CLOSE_BOX | wx.STAY_ON_TOP | wx.DIALOG_NO_PARENT
        # Create a small dialog box
        wx.Dialog.__init__(self, parent, -1, _("Add Keyword Group"), size=(350, 145), style=dlgStyle)
        # Create a main vertical sizer
        box = wx.BoxSizer(wx.VERTICAL)
        # Create a horizontal sizer for the buttons
        boxButtons = wx.BoxSizer(wx.HORIZONTAL)

        # Create a text screen object for the dialog text
        message = wx.StaticText(self, -1, _("New Keyword Group:"))
        # Add the first row to the main sizer
        box.Add(message, 0, wx.EXPAND | wx.ALL, 10)

        # Create a TextCtrl for the Keyword Group name
        self.kwGroup = wx.TextCtrl(self, -1, "")
        box.Add(self.kwGroup, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        # Do error-checking in the EVT_TEXT event
        self.kwGroup.Bind(wx.EVT_TEXT, self.OnText)

        # Create the first button, which is OK
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Set as the Default to handle ENTER keypress
        btnOK.SetDefault()
        # Bind the button event to its method
        btnOK.Bind(wx.EVT_BUTTON, self.OnButton)

        # Create the second button, which is Cancel
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        # Add an expandable spacer to the button sizer
        boxButtons.Add((20,1), 1)
        # If we're on the Mac, we want Cancel then OK
        if "__WXMAC__" in wx.PlatformInfo:
            # Add No first
            boxButtons.Add(btnCancel, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
            # Add a spacer
            boxButtons.Add((20,1))
            # Then add Yes
            boxButtons.Add(btnOK, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        # If we're not on the Mac, we want OK then Cancel
        else:
            # Add Yes first
            boxButtons.Add(btnOK, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
            # Add a spacer
            boxButtons.Add((20,1))
            # Then add No
            boxButtons.Add(btnCancel, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        # Add a final expandable spacer
        boxButtons.Add((20,1), 1)
        # Add the button bar to the main sizer
        box.Add(boxButtons, 0, wx.ALIGN_RIGHT | wx.EXPAND)

        # Turn AutoLayout On
        self.SetAutoLayout(True)
        # Set the form's main sizer
        self.SetSizer(box)
        # Fit the form
        self.Fit()
        # Lay the form out
        self.Layout()
        # Center the form on screen
    #        self.CentreOnScreen()
        # That's not working.  Let's try this ...
        TransanaGlobal.CenterOnPrimary(self)


    def OnText(self, event):

        text = self.kwGroup.GetValue()
        # Parentheses are not allowed
        if '(' in text or ')' in text:
            text = text.replace('(', '')
            text = text.replace(')', '')
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Keyword Groups cannot contain parenthesis characters.'), 'utf8')
            dlg = ErrorDialog(None, prompt)
            dlg.ShowModal()
            dlg.Destroy()
            self.kwGroup.SetValue(text)
            self.kwGroup.SetInsertionPointEnd()
            self.kwGroup.SetFocus()

        # Colons are not allowed in Keyword Groups.  Remove them if necessary.
        if ':' in text:
            text = text.replace(':', '')
            msg = unicode(_('You may not use a colon (":") in the Keyword Group name.'), 'utf8')
            dlg = ErrorDialog(None, msg)
            dlg.ShowModal()
            dlg.Destroy()
            self.kwGroup.SetValue(text)
            self.kwGroup.SetInsertionPointEnd()
            self.kwGroup.SetFocus()

        # Let's make sure we don't exceed the maximum allowed length for a Keyword Group.
        # First, let's see what the max length is.
        maxLen = TransanaGlobal.maxKWGLength
        # Check to see if we've exceeded the max length
        if len(text) > maxLen:
            ip = self.kwGroup.GetInsertionPoint()
            # If so, truncate the Keyword Group
            text = text[:maxLen]
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            msg = unicode(_('Keyword Group is limited to %d characters.'), 'utf8')
            dlg = ErrorDialog(None, msg % maxLen)
            dlg.ShowModal()
            dlg.Destroy()
            self.kwGroup.SetValue(text)
            self.kwGroup.SetInsertionPoint(ip)
            self.kwGroup.SetFocus()

    def OnButton(self, event):
        text = self.kwGroup.GetValue()
        # Check (case-insensitively) whether the Keyword Group already exists.
        if self.kw_groups_upper.count(text.upper()) != 0:
            msg = _('A Keyword Group by that name already exists.')
            dlg = ErrorDialog(None, msg)
            dlg.ShowModal()
            dlg.Destroy()
            self.kwGroup.SetFocus()
        else:
            event.Skip()
            

