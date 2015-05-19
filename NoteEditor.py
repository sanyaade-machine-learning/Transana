# Copyright (C) 2003 - 2006 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the TranscriptEditor class as part of the Editors
component.
"""

__author__ = 'Nathaniel Case, David K. Woods <dwoods@wcer.wisc.edu>'


import wx


class NoteEditor(object):
    """This class provides a simple editor for notes that are attached to
    various Transana objects.  It intends to be based on a simple text
    editing control such as wxTextCtrl or wxEditor."""

    def __init__(self, parent, default_text=""):
        """Initialize an NoteEditor object."""
        self.dlg = _NoteDialog(parent, -1, default_text)

# Public methods
    def get_text(self):
        """Run the note editor and return the note string."""
        return self.dlg.get_text()


class _NoteDialog(wx.Dialog):

    def __init__(self, parent, id, default_text=""):
        # Due to an odd behavior on the part of wxTextCtrl (it selects all text upon initialization and ignores
        # SetSelection requests), we need to track to see if we have done the initial SetSelection.  We haven't.
        self.initialized = False
        rect = wx.ClientDisplayRect()
        self.width = rect[2] * .60
        self.height = rect[3] * .60
        wx.Dialog.__init__(self, parent, id, _("Note"), wx.DefaultPosition,
                            wx.Size(self.width, self.height), wx.CAPTION | wx.CLOSE_BOX | wx.SYSTEM_MENU | wx.RESIZE_BORDER)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Specify the minimium size of this window.  It's initial size should be an adequate minimum.
        self.SetSizeHints(self.width, self.height)

        # add the note editing widget to the panel
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self, wx.Top)
        lay.left.SameAs(self, wx.Left)
        lay.right.SameAs(self, wx.Right)
        lay.bottom.SameAs(self, wx.Bottom)
        self.txt = wx.TextCtrl(self, -1, default_text, style=wx.TE_MULTILINE)
        
        # Define the SetFocus Event for the Text Control.  This is where the SetSelection can get called so it will work.
        wx.EVT_SET_FOCUS(self.txt, self.OnSetFocus)
        
        # Put the focus on the Text Control (needed for Mac)
        self.txt.SetFocus()
        
        self.txt.SetConstraints(lay)
        self.Layout()
        self.CenterOnScreen()


    def get_text(self):
        """Show the modal note editing dialog and return the note string."""
        # For reasons beyond my comprehension, this does not work.
        # self.txt.SetSelection(0, 0)
        self.ShowModal()
        return self.txt.GetValue()

    def OnSetFocus(self, event):
        # This method was suggested by Robin Dunn via the wxPython-users mailing list.
        # Essentially, to get the SetSelection to work, we have to call it after the control receives focus.

        # If we have not done this before ...
        if not self.initialized:
            # Set up the SetSelection command for the initial entry into the control
            wx.CallAfter(self.txt.SetSelection, 0, 0)
            # Then indicate that we've done this so we won't do it again.  Otherwise, the
            # cursor moves to the beginning of the text every time we focus on this control
            # instead of only the first.
            self.initialized = True
