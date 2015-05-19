# Copyright (C) 2004 - 2012 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Keyword Properties Form."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case <nacase@wisc.edu>'

import DBInterface
import Dialogs
import TransanaGlobal

import wx
import KeywordObject as Keyword

class KeywordPropertiesForm(Dialogs.GenForm):
    """Form containing Keyword fields."""

    def __init__(self, parent, id, title, keyword_object):
        self.width = 400
        self.height = 250
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(self.width, self.height), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Keyword Properties')

        self.obj = keyword_object

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Keyword Group
        self.kwg_choice = self.new_combo_box(_("Keyword Group"), v1, [""])
        # Add the element to the sizer
        r1Sizer.Add(v1, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r1Sizer.Add((10, 0))

        # wxComboBox doesn't have a Max Length method.  Let's emulate one here using the Combo Box's EVT_TEXT method
        self.kwg_choice.Bind(wx.EVT_TEXT, self.OnKWGText)
        
        # Create a VERTICAL sizer for the next element
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Keyword
        keyword_edit = self.new_edit_box(_("Keyword"), v2, self.obj.keyword, maxLen=85)
        # Add the element to the sizer
        r1Sizer.Add(v2, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v3 = wx.BoxSizer(wx.VERTICAL)
        # Definition [label]
        definition_lbl = wx.StaticText(self.panel, -1, _("Definition"))
        v3.Add(definition_lbl, 0, wx.BOTTOM, 3)

        # Definition layout [control]
        self.definition_edit = wx.TextCtrl(self.panel, -1, self.obj.definition, style=wx.TE_MULTILINE)
        v3.Add(self.definition_edit, 1, wx.EXPAND)

        # Add the element to the sizer
        r2Sizer.Add(v3, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 1, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the buttons
        self.create_buttons(sizer=btnSizer)
        # Add the button sizer to the main sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            mainSizer.Add((0, 2))

        # Set the PANEL's main sizer
        self.panel.SetSizer(mainSizer)
        # Tell the PANEL to auto-layout
        self.panel.SetAutoLayout(True)
        # Lay out the Panel
        self.panel.Layout()
        # Lay out the panel on the form
        self.Layout()
        # Resize the form to fit the contents
        self.Fit()

        # Get the new size of the form
        (width, height) = self.GetSizeTuple()
        # Reset the form's size to be at least the specified minimum width
        self.SetSize(wx.Size(max(self.width, width), max(self.height, height)))
        # Define the minimum size for this dialog as the current size
        self.SetSizeHints(max(self.width, width), max(self.height, height))
        # Center the form on screen
        self.CenterOnScreen()

        # We populate the Keyword Groups, Keywords, and Clip Keywords lists AFTER we determine the Form Size.
        # Long Keywords in the list were making the form too big!

        self.kw_groups = DBInterface.list_of_keyword_groups()
        for keywordGroup in self.kw_groups:
            self.kwg_choice.Append(keywordGroup)

        if self.obj.keywordGroup:
            # If the Keyword Group of the passed-in Keyword object does not exist, add it to the list.
            # Otherwise, there's no way to ever add a first keyword to a group!
            if self.kwg_choice.FindString(self.obj.keywordGroup) == -1:
                self.kwg_choice.Append(self.obj.keywordGroup)
            self.kwg_choice.SetStringSelection(self.obj.keywordGroup)
        else:
            self.kwg_choice.SetSelection(0)

        if self.obj.keywordGroup == '':
            self.kwg_choice.SetFocus()
        else:
            keyword_edit.SetFocus()


    def get_input(self):
        """Show the dialog and return the modified Keyword Object.  Result
        is None if user pressed the Cancel button."""
        d = Dialogs.GenForm.get_input(self)     # inherit parent method
        if d:
            self.obj.keywordGroup = d[_('Keyword Group')]
            self.obj.keyword = d[_('Keyword')]
            self.obj.definition = self.definition_edit.GetValue()
        else:
            self.obj = None

        return self.obj

    def OnKWGText(self, event):
        """ Emulate the SetMaxLength() method of the wxTextCtrl for the Keyword Group combo """
        # Current maximum length for a Keyword Group is 50 characters
        maxLen = TransanaGlobal.maxKWGLength
        # Check to see if we've exceeded out max length
        if len(self.kwg_choice.GetValue()) > maxLen:
            # If so, remove excess text
            self.kwg_choice.SetValue(self.kwg_choice.GetValue()[:maxLen])
            # Place the Insertion Point at the end of the text in the control
            self.kwg_choice.SetInsertionPoint(maxLen)
            # On Windows, emulate the sound that the TextCtrl makes.
            if 'wxMSW' in wx.PlatformInfo:
                import winsound
                winsound.PlaySound("SystemQuestion", winsound.SND_ALIAS)
        
class AddKeywordDialog(KeywordPropertiesForm):
    """Dialog used when adding a new Keyword."""

    def __init__(self, parent, id, keywordGroup=''):
        obj = Keyword.Keyword()
        obj.keywordGroup = keywordGroup
        KeywordPropertiesForm.__init__(self, parent, id, _("Add Keyword"), obj)


class EditKeywordDialog(KeywordPropertiesForm):
    """Dialog used when editing Keyword properties."""

    def __init__(self, parent, id, keyword_object):
        KeywordPropertiesForm.__init__(self, parent, id, _("Keyword Properties"), keyword_object)
