# Copyright (C) 2003 - 2012 The Board of Regents of the University of Wisconsin System 
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

import Clip
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

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Create a Grid Control on the panel where the clip data will be displayed
        self.gridClips = grid.Grid(self, -1)
        mainSizer.Add(self.gridClips, 1, wx.EXPAND)

        # Populate the Grid with initial values
        #   Grid is read-only, not editable
        self.gridClips.EnableEditing(False)
        #   Grid has 4 columns (one of which will not be visible), no rows initially
        self.gridClips.CreateGrid(1, 4)
        # Set the height of the Column Headings
        self.gridClips.SetColLabelSize(20)
        #Set the minimum acceptable column width to 0 to allow for the hidden column
        self.gridClips.SetColMinimalAcceptableWidth(0)
        # Set minimum acceptable widths for each column
        self.gridClips.SetColMinimalWidth(0, 30)
        self.gridClips.SetColMinimalWidth(1, 30)
        self.gridClips.SetColMinimalWidth(2, 30)
        self.gridClips.SetColMinimalWidth(3, 0)
        # Set Default Text in the Header row
        self.gridClips.SetColLabelValue(0, _("Clip Time"))
        self.gridClips.SetColLabelValue(1, _("Clip ID"))
        self.gridClips.SetColLabelValue(2, _("Keywords"))
        # We don't need Row labels
        self.gridClips.SetRowLabelSize(0)
        # Set the column widths
        self.gridClips.SetColSize(0, 65)
        if width > 330:
            self.gridClips.SetColSize(1, width - 185)
        else:
            self.gridClips.SetColSize(1, 140)
        self.gridClips.SetColSize(2, 250)
        # The 3rd column should not be visible.  The data is necessary for loading clips, but should not be displayed.
        self.gridClips.SetColSize(3, 0)

        # Display Cell Data
        self.DisplayCells(TimeCode)
        
        # Define the Grid Double-Click event
        grid.EVT_GRID_CELL_LEFT_DCLICK(self.gridClips, self.OnCellLeftDClick)

        # Define the Key Down Event Handler
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        
        # Perform GUI Layout
        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()

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
            # load the Clip
            tmpClip = Clip.Clip(clipData[loop]['ClipNum'])
            # Initialize the string for all the Keywords
            kwString = ''
            # Initialize the prompt for building the keyword string
            kwPrompt = '%s'
            # For each Keyword in the Keyword List ...
            for kws in tmpClip.keyword_list:
                # ... add the Keyword to the Keyword List
                kwString += kwPrompt % kws.keywordPair
                # After the first keyword, we need a NewLine in front of the Keywords.  This accompishes that!
                kwPrompt = '\n%s'

            # Insert the data values into the Grid Row
            self.gridClips.SetCellValue(loop, 0, "%s -\n %s" % (Misc.time_in_ms_to_str(clipData[loop]['ClipStart']), Misc.time_in_ms_to_str(clipData[loop]['ClipStop'])))
            self.gridClips.SetCellValue(loop, 1, tmpClip.GetNodeString(includeClip=True))
            # make the Collection / Clip ID line auto-word-wrap
            self.gridClips.SetCellRenderer(loop, 1, grid.GridCellAutoWrapStringRenderer())
            self.gridClips.SetCellValue(loop, 2, kwString)
            # Convert value to a string
            self.gridClips.SetCellValue(loop, 3, "%s" % clipData[loop]['ClipNum'])

            # Auto-size THIS row
            self.gridClips.AutoSizeRow(loop, True)
                                        

    def Refresh(self, TimeCode=None):
        """ Redraw the contents of this tab to reflect possible changes in the data since the tab was created. """
        # To refresh the window, all we need to do is re-call the DisplayCells method!
        self.DisplayCells(TimeCode)
        

    def OnCellLeftDClick(self, event):
        if self.ControlObject != None:
            # Load the Clip
            # Switched to CallAfter because of crashes on the Mac.
            wx.CallAfter(self.ControlObject.LoadClipByNumber, int(self.gridClips.GetCellValue(event.GetRow(), 3)))

            # NOTE:  LoadClipByNumber eliminates the EpisodeClipsTab, so no further processing can occur!!
            #        There used to be code here to select the Clip in the Database window, but it stopped
            #        working when I added multiple transcripts, so I moved it to the ControlObject.LoadClipByNumber()
            #        method.

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.ControlObject=ControlObject

    def OnKeyDown(self, event):
        """ Handle Key Down Events """
        # See if the ControlObject wants to handle the key that was pressed.
        if self.ControlObject.ProcessCommonKeyCommands(event):
            # If so, we're done here.  (Actually, we're done anyway.)
            return

