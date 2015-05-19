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

"""This module implements the DataWindow class for the Data Display Objects."""

__author__ = 'Nathaniel Case <nacase@wisc.edu>, David Woods <dwoods@wcer.wisc.edu>'

import wx
from DatabaseTreeTab import *
from EpisodeClipsTab import *
from KeywordsTab import *
import TransanaGlobal

class DataWindow(wx.Dialog):
    """This class implements the window containing all data display tabs."""

    def __init__(self, parent, id=-1):
        """Initialize a DataWindow object."""
        # Start with a Dialog Box (wxPython)
        wx.Dialog.__init__(self, parent, id, _("Data"), self.__pos(),
                            self.__size(),
                            style=wx.CAPTION | wx.RESIZE_BORDER)
        # Set "Window Variant" to small only for Mac to make fonts match better
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # print "DataWindow:", self.__pos(), self.__size()

        # add a Notebook Control to the Dialog Box
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self, wx.Top, 0)
        lay.left.SameAs(self, wx.Left, 0)
        lay.bottom.SameAs(self, wx.Bottom, 0)
        lay.right.SameAs(self, wx.Right, 0)
        # The wxCLIP_CHILDREN style allegedly reduces flicker.
        self.nb = wx.Notebook(self, -1, style=wx.CLIP_CHILDREN)
        # In order to 
        self.nb.parent = self
        self.nb.SetConstraints(lay)

        # Create tabs for the Notebook Control.  These tabs are complex enough that they are
        # instantiated as separate objects.

        self.DBTab = DatabaseTreeTab(self.nb)

        # Initialize the remaining Tabs to None for the moment
        self.EpisodeClipsTab = None
        self.SelectedEpisodeClipsTab = None
        self.KeywordsTab = None
        

        # Add the tabs to the Notebook Control
        self.nb.AddPage(self.DBTab, _("Database"), True)
        

        # OSX requires this for some reason or else it won't have a default
        # page selected.
        self.nb.SetSelection(0)

        self.Layout()
        self.SetAutoLayout(True)

        self.ControlObject = None            # The ControlObject handles all inter-object communication, initialized to None

        # Capture Size Changes
        wx.EVT_SIZE(self, self.OnSize)

        # Capture the selection of different Notebook Pages
        wx.EVT_NOTEBOOK_PAGE_CHANGED(self, self.nb.GetId(), self.OnNotebookPageSelect)


    def AddEpisodeClipsTab(self, seriesObj=None, episodeObj=None):
        if self.EpisodeClipsTab == None:
            self.EpisodeClipsTab = EpisodeClipsTab(self.nb, seriesObj, episodeObj)
            # If a ControlObject is defined, propagate it to the EpisodeClipsTab so that Clips can be loaded via double-clicking
            if self.ControlObject != None:
                self.EpisodeClipsTab.Register(self.ControlObject)
            self.nb.AddPage(self.EpisodeClipsTab, _("Episode Clips"), False)
            # Allow the Database Tab to redraw, as adding the Episode Clips tab can interfere with the appearance of the Data Window
            self.DBTab.Refresh()

    def AddSelectedEpisodeClipsTab(self, seriesObj=None, episodeObj=None, TimeCode=0):
        if self.SelectedEpisodeClipsTab == None:
            self.SelectedEpisodeClipsTab = EpisodeClipsTab(self.nb, seriesObj, episodeObj, TimeCode)
            # If a ControlObject is defined, propagate it to the EpisodeClipsTab so that Clips can be loaded via double-clicking
            if self.ControlObject != None:
                self.SelectedEpisodeClipsTab.Register(self.ControlObject)
            self.nb.AddPage(self.SelectedEpisodeClipsTab, _("Selected Clips"), False)
            # Allow the Database Tab to redraw, as adding the Episode Clips tab can interfere with the appearance of the Data Window
            self.DBTab.Refresh()

    def AddKeywordsTab(self, seriesObj=None, episodeObj=None, collectionObj=None, clipObj=None):
        if self.KeywordsTab == None:
            self.KeywordsTab = KeywordsTab(self.nb, seriesObj, episodeObj, collectionObj, clipObj)
            self.nb.AddPage(self.KeywordsTab, _("Keywords"), False)
            # Allow the Database Tab to redraw, as adding the Keywords tab can interfere with the appearance of the Data Window
            self.DBTab.Refresh()
        else:
            dlg = Dialogs.ErrorDialog(self, _('Problem creating the Keywords Tab!'))
            dlg.ShowModal()
            dlg.Destroy()

    def DeleteTabs(self):
        """ Delete all tabs in the Data Window except the DatabaseTreeTab, which should always be retained. """
        # On the Mac, double-clicking a Clip from the Episode Clips Tab of the Selected Clips Tab was causing
        # Transana to crash with either a Segmentation Fault or a Bus Error.  This is due to deleting 
        # the tab before it's done processing it's double-click events.  Therefore, let's detect the active tab
        # and put off deleting it until later.
        currentTab = self.nb.GetSelection()
        # Set the selection to the DatabaseTreeTab tab, which won't be deleted
        self.nb.SetSelection(0)

        # Delete all tabs but the DatabaseTreeTab.
        # Start with the Keywords Tab.  Check to see if it exists ...
        if self.KeywordsTab != None:
            # ... and if it does, delete it ...
            self.nb.DeletePage(self.nb.GetPageCount() - 1)
            # ... and remove the reference to it.
            self.KeywordsTab = None

        # Do the Selected Episode Clips Tab next.  Check to see if it exists ...
        if self.SelectedEpisodeClipsTab != None:
            # ... and if it does, delete it.  If it's the current tab, delete it later.  Otherwise, just delete it now ...
            if currentTab == 2:
                wx.CallAfter(self.DeleteTabLater)
            else:
                self.nb.DeletePage(2)
            # ... and remove the reference to it.
            self.SelectedEpisodeClipsTab = None

        # Do the Episode Clips Tab next.  Check to see if it exists ...
        if self.EpisodeClipsTab != None:
            # ... and if it does, delete it.  If it's the current tab, delete it later.  Otherwise, just delete it now ...
            if currentTab == 1:
                wx.CallAfter(self.DeleteTabLater)
            else:
                self.nb.DeletePage(1)
            # ... and remove the reference to it.
            self.EpisodeClipsTab = None

    def DeleteTabLater(self):
        """ On the Mac, selecting a clip from the Episode Clips tab or the Selected Clips tab caused a
            Segment Fault or a Bus Error.  That was because the EpisodeClipsTab tab was being deleted by
            the DeleteTabs() call above before its double-click method was done.  This method allows us
            to delay the deletion of the appropriate tab until after the tab's method is done.  """
        # The tab to be deleted is always tab 1, as the other tabs have been already deleted and the 
        # DatabaseTreeTab isn't supposed to be deleted.
        self.nb.DeletePage(1)

    def OnSize(self, event):
        """ Data Window Resize Event """
        # If we're not resizing ALL the Transana Windows ...  (avoid recursive calls!)
        if not TransanaGlobal.resizingAll:
            # Get the size of the Data Window
            (left, top) = self.GetPositionTuple()
            # Ask the Control Object to resize all other windows
            self.ControlObject.UpdateWindowPositions('Data', left - 4, YLower = top - 4)
        # Call to Layout() is required so that the Notebook Control resizes properly
        self.Layout()

    def OnNotebookPageSelect(self, event):
        """ Detect which tab in the Notebook is selected and prepare that tab for display. """
        if 'unicode' in wx.PlatformInfo:
            # If Episode Clips Tab is selected ...
            if self.nb.GetPageText(event.GetSelection()) == unicode(_("Episode Clips"), 'utf8'):
                # ... get the latest Data for the Episode Clips
                self.EpisodeClipsTab.Refresh()
            # If the Selected Clips Tab is selected ...
            elif self.nb.GetPageText(event.GetSelection()) == unicode(_("Selected Clips"), 'utf8'):
                # ... get the latest Data based on the current Video Position
                self.SelectedEpisodeClipsTab.Refresh(self.ControlObject.GetVideoPosition())
            # If the Keyword Tab is selected ...
            elif self.nb.GetPageText(event.GetSelection()) == unicode(_('Keywords'), 'utf8'):
                # update the Keywords Tab in case something has changed since the objects were first loaded.
                self.KeywordsTab.Refresh()
        else:
            # If Episode Clips Tab is selected ...
            if self.nb.GetPageText(event.GetSelection()) == _("Episode Clips"):
                # ... get the latest Data for the Episode Clips
                self.EpisodeClipsTab.Refresh()
            # If the Selected Clips Tab is selected ...
            elif self.nb.GetPageText(event.GetSelection()) == _("Selected Clips"):
                # ... get the latest Data based on the current Video Position
                self.SelectedEpisodeClipsTab.Refresh(self.ControlObject.GetVideoPosition())
            # If the Keyword Tab is selected ...
            elif self.nb.GetPageText(event.GetSelection()) == _('Keywords'):
                # update the Keywords Tab in case something has changed since the objects were first loaded.
                self.KeywordsTab.Refresh()

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.ControlObject=ControlObject
        self.DBTab.Register(ControlObject=self.ControlObject)  # Propagate the Control Object registration
        
    def ClearData(self):
        """Clear the display."""
        # Remove any extra tabs that are displayed
        self.DeleteTabs()

    def GetDimensions(self):
        (left, top) = self.GetPositionTuple()
        (width, height) = self.GetSizeTuple()
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
        self.SetDimensions(left, top, width, height)

    def ChangeLanguages(self):
        self.SetTitle(_("Data"))
        self.nb.SetPageText(0, _("Database"))
        self.DBTab.tree.refresh_tree()
        self.DBTab.tree.create_menus()

        # print "DataWindow.ChangeLanguages()  (%s)" % _("Data")

    def __size(self):
        """Determine default size of Data Frame."""
        rect = wx.ClientDisplayRect()
        width = rect[2] * .28
        height = (rect[3] - TransanaGlobal.menuHeight) * .64
        return wx.Size(width, height)

    def __pos(self):
        """Determine default position of Data Frame."""
        rect = wx.ClientDisplayRect()
        (width, height) = self.__size()
        # rect[0] compensates if the Start menu is on the Left
        x = rect[0] + rect[2] - width - 3
        # rect[1] compensates if the Start menu is on the Top
        y = rect[1] + rect[3] - height - 3
        return wx.Point(x, y)
