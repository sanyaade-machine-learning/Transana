# Copyright (C) 2004 - 2014 The Board of Regents of the University of Wisconsin System 
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
        self.width = 450
        self.height = 400
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

        # Create a HORIZONTAL sizer for the next row
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Definition [label]
        definition_lbl = wx.StaticText(self.panel, -1, _("Default Code Color"))
        v4.Add(definition_lbl, 0, wx.BOTTOM, 3)

        # Get the list of colors for populating the control
        choices = ['']
        # Make a dictionary for looking up the UNTRANSLATED color names and color definintions that match the TRANSLATED color names
        self.colorList = {}
        # Iterate through the global Graphics Colors (which can be user-defined!)
        for x in TransanaGlobal.transana_graphicsColorList:
            # We need to exclude WHITE
            if x[1] != (255, 255, 255):
                # Get the TRANSLATED color name
                tmpColorName = _(x[0])
                # If the color name is a string ...
                if isinstance(tmpColorName, str):
                    # ... convert it to unicode
                    tmpColorName = unicode(tmpColorName, 'utf8')
                # Add the translated color name to the choice box
                choices.append(tmpColorName)
                # Add the color definition to the dictionary, using the translated name as the key
                self.colorList[tmpColorName] = x
            
        # Add the Choice Control
        self.line_color_cb = wx.Choice(self.panel, wx.ID_ANY, choices=choices)
        # If the current Keyword Object's color is in the Color List ...
        if self.obj.lineColorName in TransanaGlobal.keywordMapColourSet:
            # ... select the correct color in the choice box
            self.line_color_cb.SetStringSelection(_(self.obj.lineColorName))
        # If the current Keyword Object's color is not blank and is not in the list ...
        else:
            if self.obj.lineColorName != '':
                # ... add the color to the choice box
                self.line_color_cb.SetStringSelection(_(self.obj.lineColorName))
                # ... and add the color definition to the color list
                self.colorList[self.obj.lineColorName] = (self.obj.lineColorName, (int(self.obj.lineColorDef[1:3], 16), int(self.obj.lineColorDef[3:5], 16), int(self.obj.lineColorDef[5:7], 16)))
            else:
                self.obj.lineColorDef = ''
            
        v4.Add(self.line_color_cb, 0, wx.EXPAND | wx.BOTTOM, 3)

        # Add the element to the sizer
        r3Sizer.Add(v4, 1, wx.EXPAND)
        
        # Add a horizontal spacer to the row sizer        
        r3Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v5 = wx.BoxSizer(wx.VERTICAL)
        # Definition [label]
        definition_lbl = wx.StaticText(self.panel, -1, _("Default Snapshot Code Tool"))
        v5.Add(definition_lbl, 0, wx.BOTTOM, 3)
        self.codeShape = wx.Choice(self.panel, wx.ID_ANY, choices=['', _('Rectangle'), _('Ellipse'), _('Line'), _('Arrow')])
        # If the Draw Mode is Rectangle ...
        if self.obj.drawMode == 'Rectangle':
            # ... set the Shape to Rectangle
            self.codeShape.SetStringSelection(_('Rectangle'))
        # If the Draw Mode is Ellipse ...
        elif self.obj.drawMode == 'Ellipse':
            # ... set the Shape to Ellipse
            self.codeShape.SetStringSelection(_('Ellipse'))
        # If the Draw Mode is Line ...
        elif self.obj.drawMode == 'Line':
            # ... set the Shape to Line
            self.codeShape.SetStringSelection(_('Line'))
        # If the Draw Mode is Arrow ...
        elif self.obj.drawMode == 'Arrow':
            # ... set the Shape to Arrow
            self.codeShape.SetStringSelection(_('Arrow'))
        # If it's none of these ...
        else:
            # ... set it to blank!
            self.codeShape.SetStringSelection('')
        v5.Add(self.codeShape, 0, wx.EXPAND | wx.BOTTOM, 3)

        # Add the element to the sizer
        r3Sizer.Add(v5, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r3Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r4Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v6 = wx.BoxSizer(wx.VERTICAL)
        # Definition [label]
        definition_lbl = wx.StaticText(self.panel, -1, _("Default Snapshot Line Width"))
        v6.Add(definition_lbl, 0, wx.BOTTOM, 3)

        self.lineSize = wx.Choice(self.panel, wx.ID_ANY, choices=['', '1', '2', '3', '4', '5', '6'])
        # Set Line Width
        if self.obj.lineWidth > 0:
            self.lineSize.SetStringSelection("%d" % self.obj.lineWidth)
        else:
            self.lineSize.SetStringSelection('')
        v6.Add(self.lineSize, 0, wx.EXPAND | wx.BOTTOM, 3)

        # Add the element to the sizer
        r4Sizer.Add(v6, 1, wx.EXPAND)
        
        # Add a horizontal spacer to the row sizer        
        r4Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v7 = wx.BoxSizer(wx.VERTICAL)
        # Definition [label]
        definition_lbl = wx.StaticText(self.panel, -1, _("Default Snapshot Line Style"))
        v7.Add(definition_lbl, 0, wx.BOTTOM, 3)
        # ShortDash is indistinguishable from LongDash, at least on Windows, so I've left it out of the options here.
        self.line_style_cb = wx.Choice(self.panel, wx.ID_ANY, choices=['', _('Solid'), _('Dot'), _('Dash'), _('Dot Dash')])
        # Set the Line Style
        if self.obj.lineStyle == 'Solid':
            self.line_style_cb.SetStringSelection(_('Solid'))
        elif self.obj.lineStyle == 'Dot':
            self.line_style_cb.SetStringSelection(_('Dot'))
        elif self.obj.lineStyle == 'LongDash':
            self.line_style_cb.SetStringSelection(_('Dash'))
        elif self.obj.lineStyle == 'DotDash':
            self.line_style_cb.SetStringSelection(_('Dot Dash'))
        else:
            self.line_style_cb.SetStringSelection('')
        v7.Add(self.line_style_cb, 0, wx.EXPAND | wx.BOTTOM, 3)

        # Add the element to the sizer
        r4Sizer.Add(v7, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r4Sizer, 0, wx.EXPAND)

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

        # We need to set some minimum sizes so the sizers will work right
        self.kwg_choice.SetSizeHints(minW = 50, minH = 20)
        keyword_edit.SetSizeHints(minW = 50, minH = 20)

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
                
#            self.kwg_choice.SetStringSelection(self.obj.keywordGroup)
            # If we're on the Mac ...
            if 'wxMac' in wx.PlatformInfo:
                # ... SetStringSelection is broken, so we locate the string's item number and use SetSelection!
                self.kwg_choice.SetSelection(self.kwg_choice.GetItems().index(self.obj.keywordGroup))
            else:
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
            if self.line_color_cb.GetStringSelection() != u'':
                # This needs to be the UNTRANSLATED color name
                self.obj.lineColorName = self.colorList[self.line_color_cb.GetStringSelection()][0]  # self.line_color_cb.GetStringSelection()
                if self.obj.lineColorName == '':
                    self.obj.lineColorDef = ''
                else:
                    self.obj.lineColorDef = "#%02x%02x%02x" % self.colorList[self.line_color_cb.GetStringSelection()][1]
            else:
                self.obj.lineColorName = u''
                self.obj.lineColorDef = ''
            # The Draw Mode needs to be the UNTRANSLATED value
            choices=['', 'Rectangle', 'Ellipse', 'Line', 'Arrow']
            self.obj.drawMode = choices[self.codeShape.GetSelection()]
            self.obj.lineWidth = self.lineSize.GetStringSelection()
            # The Line Style needs to be the UNTRANSLATED value
            choices=['', 'Solid', 'Dot', 'LongDash', 'DotDash']
            self.obj.lineStyle = choices[self.line_style_cb.GetSelection()]
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
