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

""" This module implements the Episode Items and Selected Items tab class for the Data Window. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Import wxPython
import wx
from wx import grid

# import Transana's Clip Object
import Clip
# import Transana's Collection Object
import Collection
# import Database Calls
import DBInterface
# import Miscellaneous Routines
import Misc
# import Transana's Snapshot Object
import Snapshot
# import Transana's Constants
import TransanaConstants

class DataItemsTab(wx.Panel):
    """ Display all clips created from a given episode. """

    def __init__(self, parent, seriesObj=None, episodeObj=None, TimeCode=None):
        """Initialize an Episode Clips Tab object.  If the TimeCode parameter is included, you get a Selected Clips Tab. """
        self.parent = parent
        # Initialize the ControlObject
        self.ControlObject = None
        # Make the initial data objects which are passed in available to the entire EpisodeClips object
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
        #   Grid has 5 columns (two of which will not be visible), no rows initially
        self.gridClips.CreateGrid(1, 5)
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
        self.gridClips.SetColLabelValue(0, _("Time"))
        self.gridClips.SetColLabelValue(1, _("Item ID"))
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
        self.gridClips.SetColSize(4, 0)

        # Define the Grid Double-Click event
        grid.EVT_GRID_CELL_LEFT_DCLICK(self.gridClips, self.OnCellLeftDClick)

        # Define the Key Down Event Handler
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        
        # Perform GUI Layout
        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()

    def DisplayCells(self, TimeCode):
        """ Get data from the database and populate the Episode Clips / Selected Clips Grid """
        # Get clip data from the database
        clipData = DBInterface.list_of_clips_by_episode(self.episodeObj.number, TimeCode)
        if TransanaConstants.proVersion:
            # Get the snapshot data from the database
            snapshotData = DBInterface.list_of_snapshots_by_episode(self.episodeObj.number, TimeCode)

        # Combine the two lists
        cellData = {}
        # For each Clip ...
        for clip in clipData:
            # load the Clip
            tmpObj = Clip.Clip(clip['ClipNum'])
            # add the Clip to the cellData
            cellData[(clip['ClipStart'], clip['ClipStop'], tmpObj.GetNodeString(True).upper())] = tmpObj
        if TransanaConstants.proVersion:
            # for each Snapshot ...
            for snapshot in snapshotData:
                # load the Snapshot
                tmpObj = Snapshot.Snapshot(snapshot['SnapshotNum'])
                # add the Snapshot to the cellData
                cellData[(snapshot['SnapshotStart'], snapshot['SnapshotStop'], tmpObj.GetNodeString(True).upper())] = tmpObj

        # Get the Keys for the cellData
        sortedKeys = cellData.keys()
        # Sort the keys for the cellData into the right order
        sortedKeys.sort()
        
        # Add rows to the Grid to accomodate the amount of data returned, or delete rows if we have too many
        if len(cellData) > self.gridClips.GetNumberRows():
            self.gridClips.AppendRows(len(cellData) - self.gridClips.GetNumberRows(), False)
        elif len(cellData) < self.gridClips.GetNumberRows():
            self.gridClips.DeleteRows(numRows = self.gridClips.GetNumberRows() - len(cellData))

        # Initialize the Row Counter
        loop = 0
        # Add the data to the Grid
        for keyVals in sortedKeys:
            # If we have a Clip ...
            if isinstance(cellData[keyVals], Clip.Clip):
                # ... get the start, stop times and the object type
                startTime = cellData[keyVals].clip_start
                stopTime = cellData[keyVals].clip_stop
                objType = 'Clip'
                # Initialize the string for all the Keywords to blank
                kwString = unicode('', 'utf8')
                # Initialize the prompt for building the keyword string
                kwPrompt = '%s'
            # If we have a Snapshot ...
            elif isinstance(cellData[keyVals], Snapshot.Snapshot):
                # ... get the start, stop times and the object type
                startTime = cellData[keyVals].episode_start
                stopTime = cellData[keyVals].episode_start + cellData[keyVals].episode_duration
                objType = 'Snapshot'
                # if there are whole snapshot keywords ...
                if len(cellData[keyVals].keyword_list) > 0:
                    # ... initialize the string for all the Keywords to indicate this
                    kwString = unicode(_('Whole:'), 'utf8') + '\n'
                # If there are NOT whole snapshot keywords ...
                else:
                    # ... initialize the string for all the Keywords to blank
                    kwString = unicode('', 'utf8')
                # Initialize the prompt for building the keyword string
                kwPrompt = '  %s'
            # For each Keyword in the Keyword List ...
            for kws in cellData[keyVals].keyword_list:
                # ... add the Keyword to the Keyword List
                kwString += kwPrompt % kws.keywordPair
                # If we have a Clip ...
                if isinstance(cellData[keyVals], Clip.Clip):
                    # After the first keyword, we need a NewLine in front of the Keywords.  This accompishes that!
                    kwPrompt = '\n%s'
                # If we have a Snapshot ...
                elif isinstance(cellData[keyVals], Snapshot.Snapshot):
                    # After the first keyword, we need a NewLine in front of the Keywords.  This accompishes that!
                    kwPrompt = '\n  %s'

            # If we have a Snapshot, we also want to display CODED Keywords in addition to the WHOLE Snapshot keywords
            # we've already included
            if isinstance(cellData[keyVals], Snapshot.Snapshot):
                # Keep a list of the coded keywords we've already displayed
                codedKeywords = []
                # Modify the template for additional keywords
                kwPrompt = '\n  %s : %s'
                # For each of the Snapshot's Coding Objects ...
                for x in range(len(cellData[keyVals].codingObjects)):
                    # ... if the Coding Object is visible and if it is not already in the codedKeywords list ...
                    if (cellData[keyVals].codingObjects[x]['visible']) and \
                      (not (cellData[keyVals].codingObjects[x]['keywordGroup'], cellData[keyVals].codingObjects[x]['keyword']) in codedKeywords):
                        # ... if this is the FIRST Coded Keyword ...
                        if len(codedKeywords) == 0:
                            # ... and if there WERE Whole Snapshot Keywords ...
                            if len(kwString) > 0:
                                # ... then add a line break to the Keywords String ...
                                kwString += '\n'
                            # ... add the indicator to the Keywords String that we're starting to show Coded Keywords
                            kwString += unicode(_('Coded:'), 'utf8')
                        # ... add the coded keyword to the Keywords String ...
                        kwString += kwPrompt % (cellData[keyVals].codingObjects[x]['keywordGroup'], cellData[keyVals].codingObjects[x]['keyword'])
                        # ... add the keyword to the Coded Keywords list
                        codedKeywords.append((cellData[keyVals].codingObjects[x]['keywordGroup'], cellData[keyVals].codingObjects[x]['keyword']))

            # Insert the data values into the Grid Row
            # Start and Stop time in column 0
            self.gridClips.SetCellValue(loop, 0, "%s -\n %s" % (Misc.time_in_ms_to_str(startTime), Misc.time_in_ms_to_str(stopTime)))
            # Node String (including Item name) in column 1
            self.gridClips.SetCellValue(loop, 1, cellData[keyVals].GetNodeString(True))
            # make the Collection / Item ID line auto-word-wrap
            self.gridClips.SetCellRenderer(loop, 1, grid.GridCellAutoWrapStringRenderer())
            # Keywords in column 2
            self.gridClips.SetCellValue(loop, 2, kwString)
            # Item Number (hidden) in column 3.  Convert value to a string
            self.gridClips.SetCellValue(loop, 3, "%s" % cellData[keyVals].number)
            # Item Type (hidden) in column 4
            self.gridClips.SetCellValue(loop, 4, "%s" % objType)
            # Auto-size THIS row
            self.gridClips.AutoSizeRow(loop, True)
            # Increment the Row Counter
            loop += 1
        # Select the first cell
        self.gridClips.SetGridCursor(0, 0)

    def Refresh(self, TimeCode=None):
        """ Redraw the contents of this tab to reflect possible changes in the data since the tab was created. """
        # To refresh the window, all we need to do is re-call the DisplayCells method!
        self.DisplayCells(TimeCode)
        

    def OnCellLeftDClick(self, event):
        """ Handle Double-Click of a Grid Cell """
        # If there is a defined Control Object (which there always should be) ...
        if self.ControlObject != None:
            # ... "activate" the selected cell
            self.SelectCell(event.GetCol(), event.GetRow())

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.ControlObject=ControlObject

    def OnKeyDown(self, event):
        """ Handle Key Down Events """
        # Determine what key was pressed
        key = event.GetKeyCode()
        
        # Cursor Down
        if key == wx.WXK_DOWN:
            self.gridClips.MoveCursorDown(False)
        # Cursor Up
        elif key == wx.WXK_UP:
            self.gridClips.MoveCursorUp(False)
        # Cursor Left
        elif key == wx.WXK_LEFT:
            self.gridClips.MoveCursorLeft(False)
        # Cursor Right
        elif key == wx.WXK_RIGHT:
            self.gridClips.MoveCursorRight(False)
        # Return / Enter
        elif key == wx.WXK_RETURN:
            # ... "activate" the selected cell
            self.SelectCell(self.gridClips.GetGridCursorCol(), self.gridClips.GetGridCursorRow())
        # Otherwise, see if the ControlObject wants to handle the key that was pressed.
        elif self.ControlObject.ProcessCommonKeyCommands(event):
            # If so, we're done here.  (Actually, we're done anyway.)
            return

    def SelectCell(self, col, row):
        """ Handle the selection of an individual cell in the Grid """
        # If we're in the "Time" column ...
        if col == 0:
            # ... determine the time of the selected item
            times = self.gridClips.GetCellValue(row, col).split('-')
            # ... and move to that time
            self.ControlObject.SetVideoStartPoint(Misc.time_in_str_to_ms(times[0]))
            # Retain the program focus in the Data Window
            self.gridClips.SetFocus()
        # If we're in the "ID" or "Keywords" column ...
        else:
            # ... if we have a Clip ...
            if self.gridClips.GetCellValue(row, 4) == 'Clip':
                # Load the Clip
                # Switched to CallAfter because of crashes on the Mac.
                wx.CallAfter(self.ControlObject.LoadClipByNumber, int(self.gridClips.GetCellValue(row, 3)))

                # NOTE:  LoadClipByNumber eliminates the DataItemsTab, so no further processing can occur!!
                #        There used to be code here to select the Clip in the Database window, but it stopped
                #        working when I added multiple transcripts, so I moved it to the ControlObject.LoadClipByNumber()
                #        method.

            # ... if we have a Snapshot ...
            elif self.gridClips.GetCellValue(row, 4) == 'Snapshot':
                tmpSnapshot = Snapshot.Snapshot(int(self.gridClips.GetCellValue(row, 3)))
                # ... load the Snapshot
                self.ControlObject.LoadSnapshot(tmpSnapshot)
                # ... determine the time of the selected item
                times = self.gridClips.GetCellValue(row, 0).split('-')
                # ... and move to that time
                self.ControlObject.SetVideoStartPoint(Misc.time_in_str_to_ms(times[0]))
                # ... and set the program focus on the Snapshot
                self.ControlObject.SelectSnapshotWindow(tmpSnapshot.id, tmpSnapshot.number, selectInDataWindow=True)
