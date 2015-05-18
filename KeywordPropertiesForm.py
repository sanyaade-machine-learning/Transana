# Copyright (C) 2004 The Board of Regents of the University of Wisconsin System 
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

import wx
import Keyword

class KeywordPropertiesForm(Dialogs.GenForm):
    """Form containing Keyword fields."""

    def __init__(self, parent, id, title, keyword_object):
        self.width = 400
        self.height = 250
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(self.width, self.height), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Keyword Properties')
        # Define the minimum size for this dialog as the initial size
        self.SetSizeHints(self.width, self.height)

        self.obj = keyword_object

        self.kw_groups = DBInterface.list_of_keyword_groups()
        # Keyword Group Layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 46)  # 46% width
        lay.height.AsIs()
        self.kwg_choice = self.new_combo_box(_("Keyword Group"), lay, [""] + self.kw_groups)
        if self.obj.keywordGroup:
            # If the Keyword Group of the passed-in Keyword object does not exist, add it to the list.
            # Otherwise, there's no way to ever add a first keyword to a group!
            if self.kwg_choice.FindString(self.obj.keywordGroup) == -1:
                self.kwg_choice.Append(self.obj.keywordGroup)
            self.kwg_choice.SetStringSelection(self.obj.keywordGroup)
        else:
            self.kwg_choice.SetSelection(0)

        # Keyword layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from Keyword Group
        lay.left.RightOf(self.kwg_choice, 10)   # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 46)  # 46% width
        lay.height.AsIs()
        keyword_edit = self.new_edit_box(_("Keyword"), lay, self.obj.keyword)

        # Definition layout [label]
        lay = wx.LayoutConstraints()
        lay.top.Below(self.kwg_choice, 10)      # 10 under Keyword Group
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        definition_lbl = wx.StaticText(self.panel, -1, _("Definition"))
        definition_lbl.SetConstraints(lay)

        # Definition layout [control]
        lay = wx.LayoutConstraints()
        lay.top.Below(definition_lbl, 3)        # 10 under label
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.bottom.SameAs(self.panel, wx.Bottom, 40)   # 40 from bottom
        self.definition_edit = wx.TextCtrl(self.panel, -1, self.obj.definition, style=wx.TE_MULTILINE)
        self.definition_edit.SetConstraints(lay)
        # Definitions have a maximum length of 255 characters
        self.definition_edit.SetMaxLength(255)

        self.Layout()
        self.SetAutoLayout(True)

        # Center on the screen (for Mac)
        self.CentreOnScreen()

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

