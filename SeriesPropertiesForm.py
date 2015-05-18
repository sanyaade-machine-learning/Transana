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
from Series import *

class SeriesPropertiesForm(Dialogs.GenForm):
    """Form containing Series fields."""

    def __init__(self, parent, id, title, series_object):
        self.width = 400
        self.height = 210
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(self.width, self.height), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Series Properties')
        # Define the minimum size for this dialog as the initial size, and define height as unchangeable
        self.SetSizeHints(self.width, self.height, -1, self.height)

        self.obj = series_object

        # Series ID layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from top
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 40)  # 40% width
        lay.height.AsIs()
        id_edit = self.new_edit_box(_("Series ID"), lay, self.obj.id)

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

        self.kw_groups = DBInterface.list_of_keyword_groups()
        # Default KW Group Layout
        lay = wx.LayoutConstraints()
        lay.top.Below(comment_edit, 10)         # 10 under comment
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from left
        lay.width.PercentOf(self.panel, wx.Width, 40)  # 40% width
        lay.height.AsIs()
        self.kwg_choice = self.new_choice_box(_("Default Keyword Group"), lay, [""] + self.kw_groups, 0)
        if (self.obj.keyword_group) and (self.kwg_choice.FindString(self.obj.keyword_group) != wx.NOT_FOUND):
            self.kwg_choice.SetStringSelection(self.obj.keyword_group)
        else:
            self.kwg_choice.SetSelection(0)

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

        id_edit.SetFocus()

    def get_input(self):
        """Show the dialog and return the modified Series Object.  Result
        is None if user pressed the Cancel button."""
        d = Dialogs.GenForm.get_input(self)     # inherit parent method
        if d:
            self.obj.id = d[_('Series ID')]
            self.obj.owner = d[_('Owner')]
            self.obj.comment = d[_('Title/Comment')]
            self.obj.keyword_group = self.kwg_choice.GetStringSelection()
        else:
            self.obj = None

        return self.obj
        
class AddSeriesDialog(SeriesPropertiesForm):
    """Dialog used when adding a new Series."""

    def __init__(self, parent, id):
        obj = Series()
        obj.owner = DBInterface.get_username()
        SeriesPropertiesForm.__init__(self, parent, id, _("Add Series"), obj)


class EditSeriesDialog(SeriesPropertiesForm):
    """Dialog used when editing Series properties."""

    def __init__(self, parent, id, series_object):
        SeriesPropertiesForm.__init__(self, parent, id, _("Series Properties"), series_object)

