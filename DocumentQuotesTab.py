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

""" This module implements the Document Items tab class for the Data Window. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Import wxPython
import wx
from wx import grid

# Import Python's sys module
import sys

# import Transana's Collection Object
import Collection
# import Database Calls
import DBInterface
### import Miscellaneous Routines
##import Misc
# import Transana's Quote Object
import Quote
### import Transana's Snapshot Object
##import Snapshot
### import Transana's Constants
##import TransanaConstants

class DocumentQuotesTab(wx.Panel):
    """ Display all Quotes created from a given Document. """

    def __init__(self, parent, libraryObj=None, documentObj=None, textPos=-1, textSel=(-2, -2)):
        """Initialize a Document Quotes Tab object."""
        self.parent = parent
        # Initialize the ControlObject
        self.ControlObject = None
        # Make the initial data objects which are passed in available to the entire DocumentQuotesTab object
        self.libraryObj = libraryObj
        self.documentObj = documentObj
        self.textPos = textPos
        self.textSel = textSel

        # Get the size of the parent window
        psize = parent.GetSizeTuple()
        # Determine the size of the panel to be placed on the dialog, slightly smaller than the parent window
        width = psize[0] - 13 
        height = psize[1] - 45

        # Create a Panel to put stuff on.  Use WANTS_CHARS style so the panel doesn't eat the Enter key.
        # (This panel implements the Document Quotes Tab!  All of the window and Notebook structure is provided by DataWindow.py.)
        wx.Panel.__init__(self, parent, -1, style=wx.WANTS_CHARS, size=(width, height), name='DocumentQuotesTabPanel')

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Create a Grid Control on the panel where the quote data will be displayed
        self.gridQuotes = grid.Grid(self, -1)
        mainSizer.Add(self.gridQuotes, 1, wx.EXPAND)

        # Populate the Grid with initial values
        #   Grid is read-only, not editable
        self.gridQuotes.EnableEditing(False)
        #   Grid has 5 columns (two of which will not be visible), no rows initially
        self.gridQuotes.CreateGrid(1, 5)
        # Set the height of the Column Headings
        self.gridQuotes.SetColLabelSize(20)
        #Set the minimum acceptable column width to 0 to allow for the hidden column
        self.gridQuotes.SetColMinimalAcceptableWidth(0)
        # Set minimum acceptable widths for each column
        self.gridQuotes.SetColMinimalWidth(0, 30)
        self.gridQuotes.SetColMinimalWidth(1, 30)
        self.gridQuotes.SetColMinimalWidth(2, 30)
        self.gridQuotes.SetColMinimalWidth(3, 0)
        # Set Default Text in the Header row
        self.gridQuotes.SetColLabelValue(0, _("Position"))
        self.gridQuotes.SetColLabelValue(1, _("Item ID"))
        self.gridQuotes.SetColLabelValue(2, _("Keywords"))
        # We don't need Row labels
        self.gridQuotes.SetRowLabelSize(0)
        # Set the column widths
        self.gridQuotes.SetColSize(0, 65)
        if width > 330:
            self.gridQuotes.SetColSize(1, width - 185)
        else:
            self.gridQuotes.SetColSize(1, 140)
        self.gridQuotes.SetColSize(2, 250)
        # The 3rd column should not be visible.  The data is necessary for loading clips, but should not be displayed.
        self.gridQuotes.SetColSize(3, 0)
        self.gridQuotes.SetColSize(4, 0)

        # Define the Grid Double-Click event
        grid.EVT_GRID_CELL_LEFT_DCLICK(self.gridQuotes, self.OnCellLeftDClick)

        # Create a variable to track whether the screen needs to be re-drawn
        self.redraw = False
        self.redrawComplete = True
        # Define the Key Down Event Handler
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        # Define the OnIdle event handler
        self.Bind(wx.EVT_IDLE, self.OnIdle)
        # Define the OnClose event handler
##        self.Bind(wx.EVT_CLOSE, self.OnClose)
        
        # Perform GUI Layout
        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()

    def DisplayCells(self, textPos=-1, textSel=(-2, -2)):
        """ Get data from the database and populate the Document Quotes Grid """
        # Update the Position and Selection values
        self.textPos = textPos
        self.textSel = textSel
        # Signal the need to redraw the data grid
        self.redraw = True

    def OnIdle(self, event):
        """ Update the contents of this DataWindow tab during IDLE time, as it could be slow, especially during
            the editing of a document """

        # NOTE:  When there are a large number of Quotes, this routine can, not surprisingly, be pretty
        #        slow.  In order to avoid this being disruptive to program activity, I've taken a couple
        #        of steps.  First, this is getting called in the OnIdle event, so will only happen when
        #        there's time.  Second, I've added "wx.YieldIfNeeded()" calls at several places so that
        #        Transana won't lock up while this is being processed.  Third, this processing will only
        #        occur while the page is showing.  Fourth, there are a number of places where this routine
        #        will bail out if another call to redraw it is made before the last redraw is complete.

        try:
            # If the DocumentQuotes tab needs to be re-drawn AND it is currently showing ...
            if self.redraw and self.IsShown():
                # Signal that the Redraw has begun.  Do it here so we can detect if ANOTHER redraw is
                # requested even before this one is complete.
                self.redraw = False
                self.redrawComplete = False

                # Get quote data from the CONTROL OBJECT.  This way, we will get the LIVE data if this
                # document is currently loaded, and will get the data from the database otherwise
                quoteData = self.ControlObject.GetQuoteDataForDocument(self.documentObj.number, self.textPos, self.textSel)
        ##        if TransanaConstants.proVersion:
        ##            # Get the snapshot data from the database
        ##            snapshotData = DBInterface.list_of_snapshots_by_episode(self.documentObj.number, TimeCode)
        ##

                # Combine the two lists
                cellData = {}
                # For each Quote ...
                for quote in quoteData:
                    # load the Quote
                    tmpObj = Quote.Quote(quote['QuoteNum'])
                    # Update the Quote's Start and End times based on editing changes
                    tmpObj.start_char = quote['StartChar']
                    tmpObj.end_char = quote['EndChar']
                    # add the Quote to the cellData
                    cellData[(quote['StartChar'], quote['EndChar'], tmpObj.GetNodeString(True).upper())] = tmpObj

                    # See if other Transana actions that take priority should be allowed
                    wx.YieldIfNeeded()
                    if self.redraw:
                        break

        ##        if TransanaConstants.proVersion:
        ##            # for each Snapshot ...
        ##            for snapshot in snapshotData:
        ##                # load the Snapshot
        ##                tmpObj = Snapshot.Snapshot(snapshot['SnapshotNum'])
        ##                # add the Snapshot to the cellData
        ##                cellData[(snapshot['SnapshotStart'], snapshot['SnapshotStop'], tmpObj.GetNodeString(True).upper())] = tmpObj

                # See if other Transana actions that take priority should be allowed
                wx.YieldIfNeeded()

                if self.redraw:
                    return
                
                # Get the Keys for the cellData
                sortedKeys = cellData.keys()
                # Sort the keys for the cellData into the right order
                sortedKeys.sort()

                # Add rows to the Grid to accomodate the amount of data returned, or delete rows if we have too many
                if len(cellData) > self.gridQuotes.GetNumberRows():
                    self.gridQuotes.AppendRows(len(cellData) - self.gridQuotes.GetNumberRows(), False)
                elif len(cellData) < self.gridQuotes.GetNumberRows():
                    self.gridQuotes.DeleteRows(numRows = self.gridQuotes.GetNumberRows() - len(cellData))

                # Initialize the Row Counter
                loop = 0
                # Add the data to the Grid
                for keyVals in sortedKeys:
                    # If we have a Quote ...
                    if isinstance(cellData[keyVals], Quote.Quote):
                        # ... get the start, end characters and the object type
                        startChar = cellData[keyVals].start_char
                        endChar = cellData[keyVals].end_char
                        objType = 'Quote'
                        # Initialize the string for all the Keywords to blank
                        kwString = unicode('', 'utf8')
                        # Initialize the prompt for building the keyword string
                        kwPrompt = '%s'
        ##            # If we have a Snapshot ...
        ##            elif isinstance(cellData[keyVals], Snapshot.Snapshot):
        ##                # ... get the start, stop times and the object type
        ##                startTime = cellData[keyVals].episode_start
        ##                stopTime = cellData[keyVals].episode_start + cellData[keyVals].episode_duration
        ##                objType = 'Snapshot'
        ##                # if there are whole snapshot keywords ...
        ##                if len(cellData[keyVals].keyword_list) > 0:
        ##                    # ... initialize the string for all the Keywords to indicate this
        ##                    kwString = unicode(_('Whole:'), 'utf8') + '\n'
        ##                # If there are NOT whole snapshot keywords ...
        ##                else:
        ##                    # ... initialize the string for all the Keywords to blank
        ##                    kwString = unicode('', 'utf8')
        ##                # Initialize the prompt for building the keyword string
        ##                kwPrompt = '  %s'


                    # See if other Transana actions that take priority should be allowed
                    wx.YieldIfNeeded()

                    if self.redraw:
                        break
                
                    # For each Keyword in the Keyword List ...
                    for kws in cellData[keyVals].keyword_list:
                        # ... add the Keyword to the Keyword List
                        kwString += kwPrompt % kws.keywordPair
                        # If we have a Quote ...
                        if isinstance(cellData[keyVals], Quote.Quote):
                            # After the first keyword, we need a NewLine in front of the Keywords.  This accompishes that!
                            kwPrompt = '\n%s'
        ##                # If we have a Snapshot ...
        ##                elif isinstance(cellData[keyVals], Snapshot.Snapshot):
        ##                    # After the first keyword, we need a NewLine in front of the Keywords.  This accompishes that!
        ##                    kwPrompt = '\n  %s'

        ##            # If we have a Snapshot, we also want to display CODED Keywords in addition to the WHOLE Snapshot keywords
        ##            # we've already included
        ##            if isinstance(cellData[keyVals], Snapshot.Snapshot):
        ##                # Keep a list of the coded keywords we've already displayed
        ##                codedKeywords = []
        ##                # Modify the template for additional keywords
        ##                kwPrompt = '\n  %s : %s'
        ##                # For each of the Snapshot's Coding Objects ...
        ##                for x in range(len(cellData[keyVals].codingObjects)):
        ##                    # ... if the Coding Object is visible and if it is not already in the codedKeywords list ...
        ##                    if (cellData[keyVals].codingObjects[x]['visible']) and \
        ##                      (not (cellData[keyVals].codingObjects[x]['keywordGroup'], cellData[keyVals].codingObjects[x]['keyword']) in codedKeywords):
        ##                        # ... if this is the FIRST Coded Keyword ...
        ##                        if len(codedKeywords) == 0:
        ##                            # ... and if there WERE Whole Snapshot Keywords ...
        ##                            if len(kwString) > 0:
        ##                                # ... then add a line break to the Keywords String ...
        ##                                kwString += '\n'
        ##                            # ... add the indicator to the Keywords String that we're starting to show Coded Keywords
        ##                            kwString += unicode(_('Coded:'), 'utf8')
        ##                        # ... add the coded keyword to the Keywords String ...
        ##                        kwString += kwPrompt % (cellData[keyVals].codingObjects[x]['keywordGroup'], cellData[keyVals].codingObjects[x]['keyword'])
        ##                        # ... add the keyword to the Coded Keywords list
        ##                        codedKeywords.append((cellData[keyVals].codingObjects[x]['keywordGroup'], cellData[keyVals].codingObjects[x]['keyword']))

                    # Insert the data values into the Grid Row
                    # Start and End Characters  in column 0
                    self.gridQuotes.SetCellValue(loop, 0, "%s -\n %s" % (startChar, endChar))
                    # Node String (including Item name) in column 1
                    self.gridQuotes.SetCellValue(loop, 1, cellData[keyVals].GetNodeString(True))
                    # make the Collection / Item ID line auto-word-wrap
                    self.gridQuotes.SetCellRenderer(loop, 1, grid.GridCellAutoWrapStringRenderer())
                    # Keywords in column 2
                    self.gridQuotes.SetCellValue(loop, 2, kwString)
                    # Item Number (hidden) in column 3.  Convert value to a string
                    self.gridQuotes.SetCellValue(loop, 3, "%s" % cellData[keyVals].number)
                    # Item Type (hidden) in column 4
                    self.gridQuotes.SetCellValue(loop, 4, "%s" % objType)
                    # Auto-size THIS row
                    self.gridQuotes.AutoSizeRow(loop, True)
                    # Increment the Row Counter
                    loop += 1

                # See if other Transana actions that take priority should be allowed
                wx.YieldIfNeeded()

                if not self.redraw:
                    # Select the first cell
                    self.gridQuotes.SetGridCursor(0, 0)
                    self.gridQuotes.Refresh()
                    self.redrawComplete = True

        # We get PyDeadObjectError exceptions when a tab has been deleted because we're changing Objects.
        except wx._core.PyDeadObjectError, e:
            # We can safely ignore this!
            pass
        
        except:

            print
            print "DocumentQuoteTab.OnIdle EXCEPTION:"
            print sys.exc_info()[0]
            print sys.exc_info()[1]
            print
            import traceback
            print traceback.print_exc(file=sys.stdout)
            

    def Refresh(self, textPos=-1, textSel=(-2, -2)):
        """ Redraw the contents of this tab to reflect possible changes in the data since the tab was created. """
        # To refresh the window, all we need to do is re-call the DisplayCells method!
        self.DisplayCells(textPos, textSel)

    def OnCellLeftDClick(self, event):
        """ Handle Double-Click of a Grid Cell """
        # If there is a defined Control Object (which there always should be) ...
        if (self.ControlObject != None):
            if not self.redraw and self.redrawComplete:
                # ... "activate" the selected cell
                self.SelectCell(event.GetCol(), event.GetRow())
            else:

                print "DocumentQuotesTab.OnCellLeftDClick():  Wait for it ...."

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.ControlObject=ControlObject

    def OnKeyDown(self, event):
        """ Handle Key Down Events """
        # Determine what key was pressed
        key = event.GetKeyCode()
        
        # Cursor Down
        if key == wx.WXK_DOWN:
            self.gridQuotes.MoveCursorDown(False)
        # Cursor Up
        elif key == wx.WXK_UP:
            self.gridQuotes.MoveCursorUp(False)
        # Cursor Left
        elif key == wx.WXK_LEFT:
            self.gridQuotes.MoveCursorLeft(False)
        # Cursor Right
        elif key == wx.WXK_RIGHT:
            self.gridQuotes.MoveCursorRight(False)
        # Return / Enter
        elif key == wx.WXK_RETURN:
            # ... "activate" the selected cell
            self.SelectCell(self.gridQuotes.GetGridCursorCol(), self.gridQuotes.GetGridCursorRow())
        # Otherwise, see if the ControlObject wants to handle the key that was pressed.
        elif self.ControlObject.ProcessCommonKeyCommands(event):
            # If so, we're done here.  (Actually, we're done anyway.)
            return

    def SelectCell(self, col, row):
        """ Handle the selection of an individual cell in the Grid """
        # If we're in the Position column ...
        if col == 0:
            try:
                # ... determine the position of the selected item
                positions = self.gridQuotes.GetCellValue(row, col).split('-')
                # ... and move to that time
                self.ControlObject.TranscriptWindow.dlg.editor.SetSelection(long(positions[0]), long(positions[1]))
            
                # Make sure the selection is visible
                # Show the END of the selection
                self.ControlObject.TranscriptWindow.dlg.editor.ShowPosition(long(positions[1]))
                # Show the START of the selection
                self.ControlObject.TranscriptWindow.dlg.editor.ShowPosition(long(positions[0]))

                # Retain the program focus in the Data Window
                self.gridQuotes.SetFocus()
            except ValueError, e:
                self.Refresh()
        # If we're in the "ID" or "Keywords" column ...
        else:
            # ... if we have a Quote ...
            if self.gridQuotes.GetCellValue(row, 4) == 'Quote':
                # Load the Quote
                # Switched to CallAfter because of crashes on the Mac.
                wx.CallAfter(self.ControlObject.LoadQuote, int(self.gridQuotes.GetCellValue(row, 3)))

##                # NOTE:  LoadClipByNumber eliminates the DataItemsTab, so no further processing can occur!!
##                #        There used to be code here to select the Clip in the Database window, but it stopped
##                #        working when I added multiple transcripts, so I moved it to the ControlObject.LoadClipByNumber()
##                #        method.

##            # ... if we have a Snapshot ...
##            elif self.gridQuotes.GetCellValue(row, 4) == 'Snapshot':
##                tmpSnapshot = Snapshot.Snapshot(int(self.gridQuotes.GetCellValue(row, 3)))
##                # ... load the Snapshot
##                self.ControlObject.LoadSnapshot(tmpSnapshot)
##                # ... determine the time of the selected item
##                times = self.gridQuotes.GetCellValue(row, 0).split('-')
##                # ... and move to that time
##                self.ControlObject.SetVideoStartPoint(Misc.time_in_str_to_ms(times[0]))
##                # ... and set the program focus on the Snapshot
##                self.ControlObject.SelectSnapshotWindow(tmpSnapshot.id, tmpSnapshot.number, selectInDataWindow=True)
