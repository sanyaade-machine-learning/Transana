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

"""This module implements the DatabaseTreeTab class for the Data Display
Objects."""

__author__ = 'Nathaniel Case <nacase@wisc.edu>, David Woods <dwoods@wcer.wisc.edu>'

import DBInterface
import Dialogs

import wx
import Collection

class CollectionPropertiesForm(Dialogs.GenForm):
    """Form containing Collection fields."""

    def __init__(self, parent, id, title, coll_object):
        self.width = 400
        self.height = 210
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, (self.width, self.height), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Collection Properties')
        # Define the minimum size for this dialog as the initial size, and define height as unchangeable
        self.SetSizeHints(self.width, self.height, -1, self.height)

        self.obj = coll_object

        # Collection ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 40)  # 40% width
        lay.height.AsIs()
        id_edit = self.new_edit_box(_("Collection ID"), lay, self.obj.id)

        # Owner layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.RightOf(id_edit, 10)           # 10 right of series ID
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        owner_edit = self.new_edit_box(_("Owner"), lay, self.obj.owner)

        # Title/Comment layout
        lay = wx.LayoutConstraints()
        lay.top.Below(id_edit, 10)              # 10 under ID
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
        lay.height.AsIs()
        comment_edit = self.new_edit_box(_("Title/Comment"), lay, self.obj.comment)

        # Default KW Group Layout
        lay = wx.LayoutConstraints()
        lay.top.Below(comment_edit, 10)         # 10 under comment
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 40)  # 40% width
        lay.height.AsIs()
        self.kw_groups = DBInterface.list_of_keyword_groups()
        self.kwg_choice = self.new_choice_box(_("Default Keyword Group"), lay, [""] + self.kw_groups, 0)
        if (self.obj.keyword_group) and (self.kwg_choice.FindString(self.obj.keyword_group) != wx.NOT_FOUND):
            self.kwg_choice.SetStringSelection(self.obj.keyword_group)
        else:
            self.kwg_choice.SetSelection(0)

        if self.obj.parent <> 0:
            # Parent Collection Name layout
            lay = wx.LayoutConstraints()
            lay.top.Below(comment_edit, 10)         # 10 under Comment
            lay.left.RightOf(self.kwg_choice, 10)   # 10 from KW Group
            lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from right
            lay.height.AsIs()
            coll_parent_edit = self.new_edit_box(_("Parent Collection"), lay, self.obj.parentName)
            coll_parent_edit.Enable(False)

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

        id_edit.SetFocus()
        

    def get_input(self):
        """Show the dialog and return the modified Collection Object.  Result
        is None if user pressed the Cancel button."""
        d = Dialogs.GenForm.get_input(self)     # inherit parent method
        if d:
            self.obj.id = d[_('Collection ID')]
            self.obj.owner = d[_('Owner')]
            self.obj.comment = d[_('Title/Comment')]
            self.obj.keyword_group = self.kwg_choice.GetStringSelection()
        else:
            self.obj = None

        return self.obj


class AddCollectionDialog(CollectionPropertiesForm):
    """Dialog used when adding a new Collection."""

    def __init__(self, parent, id, ParentNum=0):
        obj = Collection.Collection()
        obj.owner = DBInterface.get_username()
        obj.parent = ParentNum
        CollectionPropertiesForm.__init__(self, parent, id, _("Add Collection"), obj)


class EditCollectionDialog(CollectionPropertiesForm):
    """Dialog used when editing Collection properties."""

    def __init__(self, parent, id, coll_object):
        CollectionPropertiesForm.__init__(self, parent, id, _("Collection Properties"), coll_object)

