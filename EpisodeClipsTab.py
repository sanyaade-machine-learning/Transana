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

""" This module implements the KeywordsTab class for the Data Window. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Import wxPython
import wx
from wx import grid

import Collection
# import Database Calls
import DBInterface
# import Miscellaneous Routines
import Misc

class EpisodeClipsTab(wx.Panel):
    """ Display all clips created from a given episode. """

    def __init__(self, parent, seriesObj=None, episodeObj=None, TimeCode=None):
        """Initialize an Episode Clips Tab object.  If the TimeCode parameter is included, you get a Selected Clips Tab. """
        self.parent = parent
        # Initialize the ControlObject
        self.ControlObject = None
        # Make the initial data objects which are passed in available to the entire KeywordsTab object
        self.seriesObj = seriesObj
        self.episodeObj = episodeObj

        # Get the size of the parent window
        psize = parent.GetSizeTuple()
        # Determine the size of the panel to be placed on the dialog, slightly smaller than the parent window
        width = psize[0] - 13 
        height = psize[1] - 45

        # Create a Panel to put stuff on.  Use WANTS_CHARS style so the panel doesn't eat the Enter key.
        # (This panel implements the Episode Clips Tab!  All of the window and Notebook structure is provided by DataWindow.py.)
        wx.Panel.__init__(self, parent, -1, style=wx.WANTS_CHARS, size=(width, height), name='EpisodeClipsTabPanel')

        # Create a Grid Control on the panel where the clip data will be displayed
        lay = wx.LayoutConstraints()
        lay.left.SameAs(self, wx.Left, 1)
        lay.top.SameAs(self, wx.Top, 1)
        lay.right.SameAs(self, wx.Right, 1)
        lay.bottom.SameAs(self, wx.Bottom, 1)
        self.gridClips = grid.Grid(self, -1)
        self.gridClips.SetConstraints(lay)

        # Populate the Grid with initial values
        #   Grid is read-only, not editable
        self.gridClips.EnableEditing(False)
        #   Grid has 5 columns (one of which will not be visible), no rows initially
        self.gridClips.CreateGrid(1, 5)
        # Set the height of the Column Headings
        self.gridClips.SetColLabelSize(20)
        #Set the minimum acceptable column width to 0 to allow for the hidden column
        self.gridClips.SetColMinimalAcceptableWidth(0)
        # Set minimum acceptable widths for each column
        self.gridClips.SetColMinimalWidth(0, 30)
        self.gridClips.SetColMinimalWidth(1, 30)
        self.gridClips.SetColMinimalWidth(2, 30)
        self.gridClips.SetColMinimalWidth(3, 30)
        self.gridClips.SetColMinimalWidth(4, 0)
        # Set Default Text in the Header row
        self.gridClips.SetColLabelValue(0, _("Clip Start"))
        self.gridClips.SetColLabelValue(1, _("Clip End"))
        self.gridClips.SetColLabelValue(2, _("Collection ID"))
        self.gridClips.SetColLabelValue(3, _("Clip ID"))
        # We don't need Row labels
        self.gridClips.SetRowLabelSize(0)
        # Set the column widths
        self.gridClips.SetColSize(0, 65)
        self.gridClips.SetColSize(1, 65)
        self.gridClips.SetColSize(2, 100)
        if width > 330:
            self.gridClips.SetColSize(3, width - 230)
        else:
            self.gridClips.SetColSize(3, 100)
        # The 4th column should not be visible.  The data is necessary for loading clips, but should not be displayed.
        self.gridClips.SetColSize(4, 0)

        # Display Cell Data
        self.DisplayCells(TimeCode)
        
        # Define the Grid Double-Click event
        grid.EVT_GRID_CELL_LEFT_DCLICK(self.gridClips, self.OnCellLeftDClick)
        
        # Perform GUI Layout
        self.Layout()
        self.SetAutoLayout(True)

    def DisplayCells(self, TimeCode):
        """ Get data from the database and populate the Clips Grid """
        # Get clip data from the database
        clipData = DBInterface.list_of_clips_by_episode(self.episodeObj.number, TimeCode)

        # Add rows to the Grid to accomodate the amount of data returned, or delete rows if we have too many
        if len(clipData) > self.gridClips.GetNumberRows():
            self.gridClips.AppendRows(len(clipData) - self.gridClips.GetNumberRows(), False)
        elif len(clipData) < self.gridClips.GetNumberRows():
            self.gridClips.DeleteRows(numRows = self.gridClips.GetNumberRows() - len(clipData))

        # Add the data to the Grid
        for loop in range(len(clipData)):
            self.gridClips.SetCellValue(loop, 0, Misc.time_in_ms_to_str(clipData[loop]['ClipStart']))
            self.gridClips.SetCellValue(loop, 1, Misc.time_in_ms_to_str(clipData[loop]['ClipStop']))
            self.gridClips.SetCellValue(loop, 2, clipData[loop]['CollectID'])
            self.gridClips.SetCellValue(loop, 3, clipData[loop]['ClipID'])
            # Convert value to a string
            self.gridClips.SetCellValue(loop, 4, "%s" % clipData[loop]['ClipNum'])

    def Refresh(self, TimeCode=None):
        """ Redraw the contents of this tab to reflect possible changes in the data since the tab was created. """
        # To refresh the window, all we need to do is re-call the DisplayCells method!
        self.DisplayCells(TimeCode)
        

    def OnCellLeftDClick(self, event):
        if self.ControlObject != None:
            # Load the Clip
            self.ControlObject.LoadClipByNumber(int(self.gridClips.GetCellValue(event.GetRow(), 4)))
            # Get a pointer to the Clip
            tempClip = self.ControlObject.currentObj
            # Get the Clip's Collection
            tempCollection = Collection.Collection(tempClip.collection_num)
            # Initialize the Collection List
            collectionList = [tempCollection.id]
            # Seek Collection Parents up to the root collection ...
            while tempCollection.parent != 0:
                # Load the parent collection
                tempCollection = Collection.Collection(tempCollection.parent)
                # ... and add them to the front of the Collection List
                collectionList.insert(0, tempCollection.id)
            # Add the Collections Root and the Clip name to either end of the node list
            nodeList = [_('Collections')] + collectionList + [tempClip.id]
            # Now point the DBTree (the notebook's parent window's DBTab's tree) to the loaded Clip
            self.parent.parent.DBTab.tree.select_Node(nodeList, 'ClipNode')

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.ControlObject=ControlObject
