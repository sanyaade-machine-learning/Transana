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

""" This module implements the TranscriptionUI class for the wx.RichTextCtrl as part of the Transcript Editor component. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "TranscriptionUI_RTC DEBUG is ON!!"

# Show the Formatting Panel (used for debugging)
SHOWFORMATTINGPANEL = False

# For testing purposes, there can be a small input box that allows the entry of Unicode codes
# for characters.
ALLOW_UNICODE_ENTRY = False
if ALLOW_UNICODE_ENTRY:
    print "TranscriptionUI ALLOW_UNICODE_ENTRY is ON!"

# import wxPython
import wx

# import Python's gettext module
import gettext
# import Python's os module
import os
# Import Transana's Global variables
import TransanaGlobal
# Import Transana's Dialogs
import Dialogs
# import Transana's Transcript Toolbar
from TranscriptToolbar import TranscriptToolbar
# Import Transana's Transcript Editor for the wx.RichTextCtrl
import TranscriptEditor_RTC

import time
#print "TranscriptionUI_RTC -- import time"

class TranscriptionUI(wx.Dialog):
    """This class manages the graphical user interface for the transcription
    editors component.  It creates the transcript window containing a
    TranscriptToolbar and a TranscriptEditor object."""

    def __init__(self, parent, includeClose=False):
        """Initialize an TranscriptionUI object."""
        # import the _TransanaDialog object (wrapper for TranscriptEditor_RTC) as the dlg object
        self.dlg = _TranscriptDialog(parent, -1, includeClose=includeClose, showLineNumbers=True)
        # Disable the Dialog's toolbar
        self.dlg.toolbar.Enable(0)
        # Initialize the TranscriptWindowNumber to -1
        self.transcriptWindowNumber = -1
        # Set the FOCUS in the Editor.  (Required on Mac so that CommonKeys work.)
        wx.CallAfter(self.dlg.editor.SetFocus)
        
# Public methods
    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        # Assign the passed-in control object to the dlg's ControlObject property
        self.dlg.ControlObject=ControlObject
        # Determine the appropriate Transcript window number
        self.transcriptWindowNumber = len(ControlObject.TranscriptWindow) - 1
        # Share the Transcript Window Number with the dlg
        self.dlg.transcriptWindowNumber = self.transcriptWindowNumber

    def Show(self, value=True):
        """Show the window."""
        # "value" parameter added to allow for both showing AND hiding of this window.
        # It needs to be hidden when "Video Only" Presentation Mode is invoked.
        self.dlg.Show(value)
        
    def get_editor(self):
        """Get a reference to its TranscriptEditor object."""
        return self.dlg.editor
    
    def get_toolbar(self):
        """Get a reference to its TranscriptToolbar object."""
        return self.dlg.toolbar

    def ClearDoc(self):
        """Clear the Transcript window, unload any transcript."""
        # Clear the Transcription Editor
        self.dlg.editor.ClearDoc()
        # Clear the Line Numbers
        self.dlg.ClearLineNum()
        self.dlg.EditorPaint(None)
        # Reset the Toolbar
        self.dlg.toolbar.ClearToolbar()
        # Disable the toolbar
        self.dlg.toolbar.Enable(False)
        # Disable the Search
        self.dlg.EnableSearch(False)
        # If Line Numbers have been disabled ...
        if not self.dlg.lineNum.IsShown():
            # ... then make them re-appear!
            self.dlg.lineNum.Show(True)
            self.dlg.Layout()
            self.dlg.showLineNumbers = True
        
    def LoadTranscript(self, transcriptObj):
        """Load a transcript object."""
        # Transcripts should always be loaded in a Read-Only editor.  Set the Editor to Read Only.
        # This triggers a save prompt if the current transcript needs to be saved.
        self.dlg.editor.set_read_only(True)
        # Clear the toolbar
        self.dlg.toolbar.ClearToolbar()

        # Load the transcript
        self.dlg.editor.load_transcript(transcriptObj)

        # Enable the Toolbar
        self.dlg.toolbar.Enable(True)
        # Enable the Search
        self.dlg.EnableSearch(True)

    def GetCurrentTranscriptObject(self):
        """ Return the current Transcript Object, with the edited text even if it hasn't been saved. """
        # Make a copy of the Transcript Object, since we're going to be changing it.
        tempTranscriptObj = self.dlg.editor.TranscriptObj.duplicate()
        # Update the Transcript Object's text to reflect the edited state
        # STORE XML IN THE TEXT FIELD  (This shouldn't be necessary, but the time codes don't show up without it!)
        tempTranscriptObj.text = self.dlg.editor.GetFormattedSelection('XML')
        # Now return the copy of the Transcript Object
        return tempTranscriptObj
        
    def GetDimensions(self):
        """ Get the position and size of the current TranscriptionUI window """
        # Get current Position information
        (left, top) = self.dlg.GetPositionTuple()
        # Get current Size information
        (width, height) = self.dlg.GetSizeTuple()
        # Return the values in a tuple
        return (left, top, width, height)

    def GetTranscriptDims(self):
        """Return dimensions of transcript editor component."""
        # Get current Position information
        (left, top) = self.dlg.editor.GetPositionTuple()
        # Get current Size information
        (width, height) = self.dlg.editor.GetSizeTuple()
        # Return the values in a tuple
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
        """ Set the position and size of the current TranscriptionUI window """
        self.dlg.SetDimensions(left, top, width, height)

    def UpdatePosition(self, positionMS):
        """Update the Transcript position given a time in milliseconds."""
        # Don't do anything unless auto word-tracking is enabled AND
        # edit mode is disabled.
        if TransanaGlobal.configData.wordTracking and self.dlg.editor.get_read_only():
            return self.dlg.editor.scroll_to_time(positionMS)
        else:
            # Return true so that database tree tab updating that depends
            # on this return value still works (see ControlObject)
            return 1

    def InsertText(self, text):
        """Insert text at the current cursor position."""
        # Pass through to the Editor
        self.dlg.editor.InsertStyledText(text)
        
    def InsertTimeCode(self):
        """Insert a timecode at the current transcript position."""
        # Pass through to the Editor
        self.dlg.editor.insert_timecode()

    def InsertSelectionTimeCode(self, start_ms, end_ms):
        """Insert a timecode for the currently selected period in the
        Waveform.  start_ms and end_ms should contain the start and
        end time positions of the selected time period, in milliseconds."""
        # Pass through to the Editor
        self.dlg.editor.insert_timed_pause(start_ms, end_ms)

    def TranscriptModified(self):
        """Return TRUE if transcript was modified since last save."""
        # Pass through to the Editor
        return self.dlg.editor.modified()
        
    def SaveTranscript(self):
        """Save the Transcript to the database."""
        # Pass through to the Editor
        self.dlg.editor.save_transcript()

    def SaveTranscriptAs(self, fname):
        """Export the Transcript to an RTF or XML file."""
        # Pass through to the Editor
        self.dlg.editor.export_transcript(fname)
        # If saving an RTF file on a Mac, print a user warning
        if False and (fname[-4:].lower() == '.rtf') and ("__WXMAC__" in wx.PlatformInfo):
            msg = _('If you load this RTF file into Word on the Macintosh, you need to select "Format" > "AutoFormat...",\nmake sure the "AutoFormat now" option is selected, and press "OK".  Otherwise you will\nlose some Font formatting information from the file when you save it.\n(Courier New font will be changed to Times font anyway.)')
            msg = msg + '\n\n' + _('Also, Word on the Macintosh appears to handle the Whisper (Open Dot) Character for Jeffersonian \nNotation improperly.  You will need to convert this character to Symbol font within Word, but \nconvert it back to Courier New font prior to re-import into Transana.')
            dlg = Dialogs.InfoDialog(self.dlg, msg)
            dlg.ShowModal()
            dlg.Destroy()
        
    def TranscriptUndo(self, event):
        """ Make Transcript Undo available outside the TranscriptWindow """
        # Implement this by emulating an Undo press in the tool bar
        self.get_toolbar().OnUndo(event)

    def TranscriptCut(self, event):
        """  Pass-through for the Cut() method """
        # Pass through to the Editor
        self.get_editor().OnCutCopy(event)

    def TranscriptCopy(self, event):
        """  Pass-through for the Copy() method """
        # Pass through to the Editor
        self.get_editor().OnCutCopy(event)

    def TranscriptPaste(self, event):
        """  Pass-through for the Paste() method """
        # Pass through to the Editor
        self.get_editor().OnPaste(event)

    def CallFormatDialog(self, tabToOpen=0):
        """  Pass-through for the CallFormatDialog() method """
        # Pass through to the Editor
        self.dlg.editor.CallFormatDialog(tabToOpen)

    def InsertImage(self, fileName = None):
        """ Insert an Image into the Transcript """
        # If the transcript is NOT read-only ... (needed for Snapshot auto-insertion)
        if not self.dlg.editor.get_read_only():
            # ... then signal the Editor to insert an image
            self.dlg.editor.InsertImage(fileName)

    def AdjustIndexes(self, adjustmentAmount):
        """ Adjust Transcript Time Codes by the specified amount """
        # Check to see if Human Readable Time Code Values are displayed
        tcHRValues = self.dlg.editor.timeCodeDataVisible
        # If Human Readable Time Code Values are visible ...
        if tcHRValues:
            # ... hide them!
            self.dlg.editor.changeTimeCodeValueStatus(False)
        # If the transcript is "Read-Only", it must be put into "Edit" mode.  Let's do
        # this automatically.
        if self.dlg.editor.get_read_only():
            # First, we "push" the Edit Mode button ...
            self.dlg.toolbar.ToggleTool(self.dlg.toolbar.CMD_READONLY_ID, True)
            # ... then we call the button's method as if it really had been pushed.
            self.dlg.toolbar.OnReadOnlySelect(None)
        # Now adjust the indexes
        self.dlg.editor.AdjustIndexes(adjustmentAmount)
        # If Human Readable Time Code Values were visible ...
        if tcHRValues:
            # ... show them again!
            self.dlg.editor.changeTimeCodeValueStatus(True)

    def UpdateSelectionText(self, text):
        """ Update the text indicating the start and end points of the current selection """
        # Pass through to the Dialog
        self.dlg.UpdateSelectionText(text)

    def ChangeLanguages(self):
        """ Change all on-screen prompts to the new language. """
        # Instruct the Toolbar to change languages
        self.dlg.toolbar.ChangeLanguages()
        # Instruct the Dialog (Editor) to change languages
        self.dlg.ChangeLanguages()
        

# import the wxRTC-based RichTextEditCtrl
import RichTextEditCtrl_RTC

class _TranscriptDialog(wx.Dialog):
    """ Implement a wx.Dialog-based control for the private use of the TranscriptionUI object """

    def __init__(self, parent, id=-1, includeClose=False, showLineNumbers=False):
        # If we're including an optional Close button ...
        if includeClose:
            # ... define a style that includes the Close Box.  (System_Menu is required for Close to show on Windows in wxPython.)
            style = wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS | wx.SYSTEM_MENU | wx.CLOSE_BOX
        # If we don't need the close box ...
        else:
            # ... then we don't need that defined in the style
            style = wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS
        # Remember whether we're including line numbers
        self.showLineNumbers = showLineNumbers
        # Create the Dialog with the appropriate style
        wx.Dialog.__init__(self, parent, id, _("Transcript"), self.__pos(), self.__size(), style=style)

        # Set "Window Variant" to small only for Mac to make fonts match better
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # The ControlObject handles all inter-object communication, initialized to None
        self.ControlObject = None
        # Initialize the Transcript Window Number to -1
        self.transcriptWindowNumber = -1

        # add the widgets to the panel
        sizer = wx.BoxSizer(wx.VERTICAL)

        hsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.toolbar = TranscriptToolbar(self)
        hsizer.Add(self.toolbar, 0, wx.ALIGN_TOP, 10)

        # On Windows, we need to display a "Transcript Health" indicator related to GDI resources.
        # I tried to handle this through changing the "Save" button graphic in the Toolbar, but that
        # caused the Search box to be disabled when the graphic was updated.  Weird.
        if 'wxMSW' in wx.PlatformInfo:
            # Create Bitmap objects for the four images used to populate the health indicator
            self.GDIBmpAll = wx.Bitmap(os.path.join(TransanaGlobal.programDir, 'images', "StoplightAll.xpm"), wx.BITMAP_TYPE_XPM)
            self.GDIBmpRed = wx.Bitmap(os.path.join(TransanaGlobal.programDir, 'images', "StoplightRed.xpm"), wx.BITMAP_TYPE_XPM)
            self.GDIBmpYellow = wx.Bitmap(os.path.join(TransanaGlobal.programDir, 'images', "StoplightYellow.xpm"), wx.BITMAP_TYPE_XPM)
            self.GDIBmpGreen = wx.Bitmap(os.path.join(TransanaGlobal.programDir, 'images', "StoplightGreen.xpm"), wx.BITMAP_TYPE_XPM)
            # Create a Static Bitmap to display the Transcript Health indicator.  Start it with the neutral graphic
            self.GDIBmp = wx.StaticBitmap(self, -1, self.GDIBmpAll)
            # Set the proper Tool Tip
            self.GDIBmp.SetToolTip(wx.ToolTip(_("Transcript Healthy")))
            # Add the Health Indicator graphic to the horizontal sizer just after the Toolbar.
            hsizer.Add(self.GDIBmp, 0, wx.ALL, 4)

        # Add Quick Search tools
        # Start with the Search Backwards button
        self.CMD_SEARCH_BACK_ID = wx.NewId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_BACK, wx.ART_TOOLBAR, (16,16))
        self.searchBack = wx.BitmapButton(self, self.CMD_SEARCH_BACK_ID, bmp, style=wx.NO_BORDER)
        hsizer.Add(self.searchBack, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        wx.EVT_BUTTON(self, self.CMD_SEARCH_BACK_ID, self.OnSearch)
        hsizer.Add((10, 1), 0)
        self.searchBackToolTip = wx.ToolTip(_("Search backwards"))
        self.searchBack.SetToolTip(self.searchBackToolTip)

        # Add the Search Text box
        self.searchText = wx.TextCtrl(self, -1, size=(100, 20), style=wx.TE_PROCESS_ENTER)
        hsizer.Add(self.searchText, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnSearch, self.searchText)
        hsizer.Add((10, 1), 0)

        # Add the Search Forwards button
        self.CMD_SEARCH_NEXT_ID = wx.NewId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_FORWARD, wx.ART_TOOLBAR, (16,16))
        self.searchNext = wx.BitmapButton(self, self.CMD_SEARCH_NEXT_ID, bmp, style=wx.NO_BORDER)
        hsizer.Add(self.searchNext, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        wx.EVT_BUTTON(self, self.CMD_SEARCH_NEXT_ID, self.OnSearch)
        self.searchNextToolTip = wx.ToolTip(_("Search forwards"))
        self.searchNext.SetToolTip(self.searchNextToolTip)

        # If Unicode Entry is enabled ...
        if ALLOW_UNICODE_ENTRY:
            # ... add the Unicode Entry box
            self.UnicodeEntry = wx.TextCtrl(self, -1, size=(40, 16), style=wx.TE_PROCESS_ENTER)
            self.UnicodeEntry.SetMaxLength(4)
            hsizer.Add((10,1), 0, wx.ALIGN_CENTER | wx.GROW)
            hsizer.Add(self.UnicodeEntry, 0, wx.ALIGN_RIGHT | wx.TOP | wx.RIGHT, 8)
            self.UnicodeEntry.Bind(wx.EVT_TEXT, self.OnUnicodeText)
            self.UnicodeEntry.Bind(wx.EVT_TEXT_ENTER, self.OnUnicodeEnter)

        # Add a text label that will indicate the start and end points of the current transcript selection
        self.selectionText = wx.StaticText(self, -1, "")
        hsizer.Add(self.selectionText, 1, wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 20)
        sizer.Add(hsizer, 0, wx.ALIGN_TOP | wx.EXPAND, 10)

        # Call toolbar.Realize() to initialize the toolbar            
        self.toolbar.Realize()

        # If the Format Panel is enabled ...
        if SHOWFORMATTINGPANEL:
            # ... add the Format Panel
            self.formatPanel = FormatPanel(self, -1, size=(500, 28))
            sizer.Add(self.formatPanel, 0, wx.EXPAND)

        # Add a blank Bitmap for creating Line Numbers
        self.lineNumBmp = wx.EmptyBitmap(100, 100)

        # Create a horizontal Sizer to hold the Line Numbers and the Edit Control
        hsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        
        # Create a Panel to hold Line Numbers
        self.lineNum = wx.Panel(self, -1, style = wx.BORDER_DOUBLE)
        # Set the background color for the Line Numbers panel to a pale gray
        self.lineNum.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNFACE))
        # Set the minimum and maximum width of the Line Number Panel to fix its size
        self.lineNum.SetSizeHints(minW = 50, minH = 0, maxW = 50)
        # Add the Line Number Panel to the Horizontal Sizer
        hsizer2.Add(self.lineNum, 0, wx.EXPAND | wx.TOP | wx.LEFT | wx.BOTTOM, 3)
        # Bind the PAINT event used to display line numbers
        self.lineNum.Bind(wx.EVT_PAINT, self.LineNumPaint)

        self.lineNum.Show(self.showLineNumbers)

        # Create the Rich Text Edit Control itself, using Transana's TranscriptEditor object
        self.editor = TranscriptEditor_RTC.TranscriptEditor(self, id, self.toolbar.OnStyleChange, updateSelectionText=True)
        # Place the editor on the horizontal sizer
        hsizer2.Add(self.editor, 1, wx.EXPAND | wx.ALL, 3)
        # Create a timer to update the Line Numbers
        self.LineNumTimer = wx.Timer()
        # Bind the method to the Timer event
        self.LineNumTimer.Bind(wx.EVT_TIMER, self.EditorPaint)

        # Disable the Search components initially
        self.EnableSearch(False)

        # Put the horizontal sizer in the main vertical sizer
        sizer.Add(hsizer2, 1, wx.EXPAND)
        if "__WXMAC__" in wx.PlatformInfo:
            # This adds a space at the bottom of the frame on Mac, so that the scroll bar will get the down-scroll arrow.
            sizer.Add((0, 15))
        # Finish the Size implementation
        self.SetSizer(sizer)
        self.SetAutoLayout(True)
        self.Layout()
        # Set the focus in the Editor
        self.editor.SetFocus()

        # Call Line Number Paint just once so the screen looks right
        self.EditorPaint(None)

        # Capture Size Changes
        wx.EVT_SIZE(self, self.OnSize)

        # If we're on Windows ...
        if 'wxMSW' in wx.PlatformInfo:
            # Create a timer to update the GDI Resource label
            self.GDITimer = wx.Timer()
            # Bind the method to the Timer event
            self.GDITimer.Bind(wx.EVT_TIMER, self.OnGDITimer)
            # update the GDI Resources label every few seconds
            self.GDITimer.Start(5000)

        try:
            # Defind the Activate event (for setting the active window)
            self.Bind(wx.EVT_ACTIVATE, self.OnActivate)
            # Define the Close event (for THIS Transcript Window)
            self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)

        except:
            import sys
            print "TranscriptionUI._TranscriptDialog.__init__():"
            print sys.exc_info()[0]
            print sys.exc_info()[1]
            print

    def OnActivate(self, event):
        """ Activate Event for a Transcript Window """
        # If the Control Object is defined ...
        # There's a weird Mac bug.  If you click on the Visualization Window, then close one of the
        # multiple transcript windows, then right-click the Visualization, Transana crashes.
        # That's because, the act of closing the Transcript window calls that window's
        # TranscriptionUI.OnActivate() method AFTER its OnCloseWindow() method,
        # reseting self.ControlObject.activeTranscript to the window's transcriptWindowNumber, so the subsequent
        # Visualization call tried to activate a closed transcript window.  So here we will avoid that
        # by detecting that the number of Transcript Windows is smaller than this window's transcriptWindowNumber
        # and NOT resetting activeTranscript in that circumstance.
        if (self.ControlObject != None) and (self.transcriptWindowNumber < len(self.ControlObject.TranscriptWindow)):
            # ... make this transcript the Control Object's active trasncript
            self.ControlObject.activeTranscript = self.transcriptWindowNumber
        elif (self.ControlObject != None) and (self.ControlObject.activeTranscript >= len(self.ControlObject.TranscriptWindow)):
            self.ControlObject.activeTranscript = 0
        # Let the event fall through to parent windows.
        event.Skip()

    def OnCloseWindow(self, event):
        """ Event for the Transcript Window Close button, which should only exist on secondary transcript windows """
        # Stop the Line Numbers timer
        self.LineNumTimer.Stop()
        # Clear the Line Numbers
        self.ClearLineNum()
        self.EditorPaint(None)
        # As long as there are multiple transcript windows ...
        if (self.transcriptWindowNumber > 0) or (len(self.ControlObject.TranscriptWindow) > 1):
            # Use the Control Object to close THIS transcript window
            self.ControlObject.CloseAdditionalTranscript(self.transcriptWindowNumber)
            # Allow the wxDialog's normal Close processing
            event.Skip()

    def OnSize(self, event):
        """ Transcription Window Resize Method """
        # If we are not doing global resizing of all windows ...
        if not TransanaGlobal.resizingAll:
            # Get the position of the Transcript window
            (left, top) = self.GetPositionTuple()
            # Get the size of the Transcript window
            (width, height) = self.GetSize()
            # Call the ControlObject's routine for adjusting all windows
            self.ControlObject.UpdateWindowPositions('Transcript', width + left, YUpper = top - 4)
        # Call the Transcript Window's Layout.
        self.Layout()
        # We may need to scroll to keep the current selection in the visible part of the window.
        # Find the start of the selection.
        start = self.editor.GetSelectionStart()
        # Make sure the current selection is visible
        self.editor.ShowCurrentSelection()

    def LineNumPaint(self, event):
        """ Paint Event Handler for the Line Number control.  This is used to display Line Numbers. """
        # Call the parent event handler
        event.Skip()
        # Create a Buffered Paint Device Context for the Line Number panel which will draw the Line Number bitmap
        # created in the EditorPaint() method
        dc = wx.BufferedPaintDC(self.lineNum, self.lineNumBmp, style=wx.BUFFER_VIRTUAL_AREA)

    def ClearLineNum(self):
        """ Clear Line Numbers Control """
        # Get the size of the Line Number control
        (w, h) = self.lineNum.GetSize()
        # Create an empty bitmap the same size as the lineNum control
        self.lineNumBmp = wx.EmptyBitmap(w, h)
        # Create a Buffered Device Context for manipulating the Bitmap
        dc = wx.BufferedDC(None, self.lineNumBmp)
        # Paint the background of the Bitmap the color the Control's Background is supposed to be
        dc.SetBackground(wx.Brush(self.lineNum.GetBackgroundColour()))
        # Clear the Bitmap
        dc.Clear()
        # Call Refresh so the Control will re-paint
        self.lineNum.Refresh()

    def AddLineNum(self, num, yPos):
        """ Add Line Numbers to the Line Number Control """
        # First line correction!  (This compensates for the RTC's control border)
        if yPos == 0:
            if 'wxMac' in wx.PlatformInfo:
                yPos = 6
            else:
                yPos = 4

        # Get a buffered Device Context based on the line number bitmap
        dc = wx.BufferedDC(None, self.lineNumBmp)
        # Specify the default font
        font = wx.Font(10, wx.ROMAN, wx.NORMAL, wx.NORMAL)
        # Set the font for the device context
        dc.SetFont(font)
        # Set the text color for the device context
        dc.SetTextForeground(wx.BLACK)
        # Get the size of the line number text to be drawn on the device context
        (w, h) = dc.GetTextExtent(str(num))
        # Make the Device Context editable
        dc.BeginDrawing()
        # Place the Line Number on the Device Context, right-justified
        dc.DrawText(str(num), self.lineNum.GetSize()[0] - w - 5, yPos)
        # The Device Context won't be edited any more here.
        dc.EndDrawing()

    def EditorPaint(self, event):
        """ Paint Event Handler for the Editor control.  This is used to display Line Numbers. """
        # If we're not showing line numbers, we can skip this!
        if not self.showLineNumbers:
            return

        # We need to know how long this is taking.  Get the start time.
        start = time.time()
        
        # Call the parent event handler, if there is one
        if event != None:
            event.Skip()
        # Clear all line numbers
        self.ClearLineNum()
        # Initialize the current document position to before the start of the document
        curPos = -1
        # Initialize the current line to 0, before the first line of the document
        curLine = 0
        # Initialize the positioning offset to zero
        offset = 0
        # Depending on platform, set the positioning adjustment factor for line spacing
        if 'wxMac' in wx.PlatformInfo:
            adjustmentFactor = 1.5
            stepSize = 6
        else:
            adjustmentFactor = 1.75
            stepSize = 1

        # For each stepSize vertical pixels in the Transcript display ...
        # (Using a START of 10 prevents overlapping first line numbers.  Using a STEP of stepSize prevents typing lag on the Mac.)
        for y in range(10, self.editor.GetSize()[1], stepSize):
            # ... get the character position (pos) of the character at the (10, y) pixel
            (result, pos) = self.editor.HitTest((10, y))
            # if we have move down on screen to a new character position ...
            if pos > curPos:
                # ... get the formatting for the current character
                textAttr = self.editor.GetStyleAt(pos)
                # Adjust the positioning offset for Paragraph Space Before.
                # (Dividing by three approximately translates centimeters to pixels!)
                offset += int(textAttr.GetParagraphSpacingBefore() / 3)
                # Use PositionToXY to translate character position into text row/column, then
                # see if we've moved on to the next LINE.  (Line is paragraph number, not physical
                # screen line!)  If so ...
                if curLine < self.editor.PositionToXY(pos + 1)[1] + 1:
                    # ... add the line number to to the Line Number display at the proper (adjusted) vertical value
                    self.AddLineNum(self.editor.PositionToXY(pos + 1)[1] + 1, y + offset)
                    # Since we've added the line number, update the current line number value
                    curLine = self.editor.PositionToXY(pos + 1)[1] + 1
                # Update the current character position value
                curPos = pos
                # We need to adjust the NEXT line number position for Paragraph Space After.  Note that this
                # re-initialized the offset value.
                offset = int(textAttr.GetParagraphSpacingAfter() / 3)
                # We also need to adjust the NEXT line number position for Line Spacing.
                offset += int((textAttr.GetLineSpacing() - 10) * adjustmentFactor)
        # Now that the line numbers have been determined, cause the Line Number control to be updated
        self.lineNum.Refresh()

        # note the End Time
        end = time.time()

        # If updating line numbers takes more than an acceptable amount of time, adjust the frequency of updates.
        if (end - start > 0.36) and (self.LineNumTimer.GetInterval() < 10000):
            # Reset to 10 seconds
            self.LineNumTimer.Stop()
            self.LineNumTimer.Start(10000)
        elif (end - start > 0.30) and (self.LineNumTimer.GetInterval() < 7000):
            # Reset to 7 seconds
            self.LineNumTimer.Stop()
            self.LineNumTimer.Start(7000)
        elif (end - start > 0.22) and (self.LineNumTimer.GetInterval() < 4000):
            # Reset to 4 seconds
            self.LineNumTimer.Stop()
            self.LineNumTimer.Start(4000)
        elif (end - start < 0.22) and (self.LineNumTimer.GetInterval() > 1000):
            # Reset to 1 second
            self.LineNumTimer.Stop()
            self.LineNumTimer.Start(1000)
            
            # .. then we have to make them go away.
            # First, flag them as disabled
#            self.showLineNumbers = False
            # Stop the Line Number Timer
#            self.LineNumTimer.Stop()
            # Hide the Line Number control
#            self.lineNum.Show(False)
            # Re-do the dialog layout without the line number control
#            self.Layout()
            # And inform the user.
#            prompt = _("Line Numbers have been disabled.  They were starting to interfere with Transana's performance.")
#            dlg = Dialogs.InfoDialog(None, prompt)
#            dlg.ShowModal()
#            dlg.Destroy()

    def OnSearch(self, event):
        """ Implement the Toolbar's QuickSearch """
        # Get the text for the search
        txt = self.searchText.GetValue()
        # If there is text ...
        if txt != '':
            # Determine whether we're searching forward or backward
            if event.GetId() == self.CMD_SEARCH_BACK_ID:
                direction = "back"
            # Either CMD_SEARCH_FORWARD_ID or ENTER in the text box indicate forward!
            else:
                direction = "next"
            # Perform the search in the Editor
            self.editor.find_text(txt, direction)
            # Set the focus back on the editor component, rather than the button, so Paste or typing work.
            self.editor.SetFocus()

    def EnableSearch(self, enable):
        """ Change the "Enabled" status of the Search controls """
        # Enable / Disable the Back Button
        self.searchBack.Enable(enable)
        # Enable / Disable the Search Text Box
        self.searchText.Enable(enable)
        # Enable / Disable the Forward Button
        self.searchNext.Enable(enable)
        # If we're enabling Search and are showing Line Numbers ...
        if enable and self.showLineNumbers:
            # update the Line Numbers every second.  (This isn't Search specific, but it's convenient!)
            self.LineNumTimer.Start(1000)
        else:
            # Stop the Line Numbers timer
            self.LineNumTimer.Stop()

    def ClearSearch(self):
        """ Clear the Search Box """
        # Clear the Search Text box
        self.searchText.SetValue('')
        
    def OnUnicodeText(self, event):
        """ The Unicode Entry box should only accept HEX character values """
        
        # NOTE:  This only works on characters added to the END of the string.
        #        This routine is NOT INTENDED FOR GENERAL USE, and is for developer testing only.
        
        # Get the string currently in the text box
        s = event.GetString()
        # If something has been typed ...
        if len(s) > 0:
            # If the value of the latest character isn't part of the HEX character set ...
            if not s[len(s)-1].upper() in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F']:
                # ... remove the last character from the string
                self.UnicodeEntry.SetValue(s[:-1])
                # ... and move the insertion point to the end
                self.UnicodeEntry.SetInsertionPoint(len(s)-1)

    def OnUnicodeEnter(self, event):
        """ Process the ENTER key for Unicode Entry """
        # Start exception handling
        try:
            # Get the Unicode character
            c = unichr(int(self.UnicodeEntry.GetValue(), 16))
            # Convert it to UTF8 encoding
            c = c.encode('utf8')
            # Get the current cursor position
            curpos = self.editor.GetCurrentPos()

            # This is from wxSTC.  I don't know if it will work with wxRTC.
            len = self.editor.GetTextLength()
            self.editor.InsertText(curpos, c)
            len = self.editor.GetTextLength() - len
            self.editor.GotoPos(curpos + len)
            self.UnicodeEntry.SetValue('')
            self.editor.SetFocus()
            
        # Handle exceptions
        except:
            # If we don't have a proper unicode character, we can swallow the exception
            pass

    def UpdateSelectionText(self, text):
        """ Update the text indicating the start and end points of the current selection """
        # Update the selectionText label with the supplied text
        self.selectionText.SetLabel(text)

    def ChangeLanguages(self):
        """ Change Languages """
        # update the Tool Tips for the Search Buttons
        self.searchBackToolTip.SetTip(_("Search backwards"))
        self.searchNextToolTip.SetTip(_("Search forwards"))

    def FormatUpdate(self, textAttr):
        """ Update the Format Panel """
        # If we are showing the format panel ...
        if SHOWFORMATTINGPANEL and self.formatPanel:
            # ... update it based on the style passed in
            self.formatPanel.Update(textAttr)

    def OnGDITimer(self, event):
        """ A timer method to update the GDI Resources Label on Windows """
        # If we are in Read Only mode ...
        if self.editor.get_read_only():
            # ... then skip this!
            return
        
        # Get the number of GDI Resources in use
        GDI = RichTextEditCtrl_RTC.GDIReport()
        # If the GDI usage is greater than 7,500 ...
        if GDI > 7500:
            # ... show the Trancript Health indicator as RED
            self.GDIBmp.SetBitmap(self.GDIBmpRed)
            # ... and set the appropriate Tool Tip
            self.GDIBmp.SetToolTip(wx.ToolTip(_("GDI Resources very low.\nSave immediately.")))

        # If the GDI usage is greater than 5000 ...
        elif GDI > 5000:
            # ... show the Trancript Health indicator as YELLOW
            self.GDIBmp.SetBitmap(self.GDIBmpYellow)
            # ... and set the appropriate Tool Tip
            self.GDIBmp.SetToolTip(wx.ToolTip(_("GDI Resources low.\nSave soon.")))

        # Otherwise ...
        else:
            # ... show the Trancript Health indicator as GREEN
            self.GDIBmp.SetBitmap(self.GDIBmpGreen)
            # ... and set the appropriate Tool Tip
            self.GDIBmp.SetToolTip(wx.ToolTip(_("Transcript Healthy")))


    def __size(self):
        """Determine the default size for the Transcript frame."""
        # Get the size of the first monitor
        rect = wx.Display(0).GetClientArea()

        # If we're on Linux, we may not be getting the right screen size value
        if 'wxGTK' in wx.PlatformInfo:

            print "TranscriptionUI_RTC.__size():", rect, wx.ClientDisplayRect()
            
            width = min(rect[2], 1440) * .715
        # If we're on Windows or OS X ...
        else:
            # Transcript Compontent should be 71.5% of the WIDTH
            width = rect[2] * .715

        # Transcript Compontent should be 74% of the HEIGHT, adjusted for the menu height
        height = (rect[3] - TransanaGlobal.menuHeight) * .74

        # Compensate in Linux.  I'm not sure why this is necessary, but it seems to be.
        # SEE IF CHANGINGING TransanaGlobal.menuHeight fixes this.
        if 'wxGTK' in wx.PlatformInfo:
            height -= 60

        # Return the SIZE values
        return wx.Size(width, height)

    def __pos(self):
        """Determine default position of Transcript Frame."""
        # Get the size of the first monitor
        rect = wx.Display(0).GetClientArea()
        # Get the adjusted default SIZE of the Transcription UI area of the screen
        (width, height) = self.__size()
        # rect[0] compensates if Start menu is on Left
        x = rect[0]
        # rect[1] compensates if Start menu is on Top
        y = rect[1] + rect[3] - height - 3
        # Return the POSITION values
        return (x, y)    

import wx.richtext as rt

class FormatPanel(wx.Panel):
    """ Display a panel with formatting information (Currently DEVELOPMENT ONLY) """
    def __init__(self, parent, id, size):
        wx.Panel.__init__(self, parent, id, size=size, style=wx.BORDER_SIMPLE)
        self.SetBackgroundColour('#ECE9D8')

        sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.font = wx.TextCtrl(self, -1, size=(100, 16))
        sizer.Add(self.font, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.fontSize = wx.TextCtrl(self, -1, size=(24, 16))
        sizer.Add(self.fontSize, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Bold.xpm"))
        self.bold = wx.StaticBitmap(self, -1, bmp)
        sizer.Add(self.bold, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Italic.xpm"))
        self.italic = wx.StaticBitmap(self, -1, bmp)
        sizer.Add(self.italic, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Underline.xpm"))
        self.underline = wx.StaticBitmap(self, -1, bmp)
        sizer.Add(self.underline, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        bmp = self.GetColorBitmap(wx.Colour(0, 0, 0))
        self.color = wx.StaticBitmap(self, -1, bmp)
        sizer.Add(self.color, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        bmp = self.GetColorBitmap(wx.Colour(255, 255, 255))
        self.background = wx.StaticBitmap(self, -1, bmp)
        sizer.Add(self.background, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.parAlign = wx.TextCtrl(self, -1, size=(60, 16))
        sizer.Add(self.parAlign, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.parLeftIndent = wx.TextCtrl(self, -1, size=(48, 16))
        sizer.Add(self.parLeftIndent, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.parFirstLineIndent = wx.TextCtrl(self, -1, size=(48, 16))
        sizer.Add(self.parFirstLineIndent, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.parRightIndent = wx.TextCtrl(self, -1, size=(48, 16))
        sizer.Add(self.parRightIndent, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.parLineSpacing = wx.TextCtrl(self, -1, size=(48, 16))
        sizer.Add(self.parLineSpacing, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.parBefore = wx.TextCtrl(self, -1, size=(48, 16))
        sizer.Add(self.parBefore, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.parAfter = wx.TextCtrl(self, -1, size=(48, 16))
        sizer.Add(self.parAfter, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.tabs = wx.TextCtrl(self, -1, size=(100, 16))
        sizer.Add(self.tabs, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.txt = wx.StaticText(self, -1, "", size=(150, 16))
        sizer.Add(self.txt, 1,  wx.LEFT | wx.RIGHT | wx.TOP | wx.BOTTOM, 3)

        self.SetSizer(sizer)
        self.SetAutoLayout(True)
        self.Layout()

    def Update(self, textAttr):

        if TransanaGlobal.configData.formatUnits == 'cm':
            convertVal = 100.0
        else:
            convertVal = 254.0

        font = textAttr.GetFont()
        if font.IsOk():
            self.font.SetValue(font.GetFaceName())
            self.fontSize.SetValue(str(font.GetPointSize()))

            if font.GetWeight() == wx.FONTWEIGHT_NORMAL:
                bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Bold.xpm"))
            else:
                bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "BoldOn.xpm"))
            self.bold.SetBitmap(bmp)

            if font.GetStyle() == wx.FONTSTYLE_NORMAL:
                bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Italic.xpm"))
            else:
                bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "ItalicOn.xpm"))
            self.italic.SetBitmap(bmp)

            if not font.GetUnderlined():
                bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Underline.xpm"))
            else:
                bmp = wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "UnderlineOn.xpm"))
            self.underline.SetBitmap(bmp)

        bmp = self.GetColorBitmap(textAttr.GetTextColour())
        self.color.SetBitmap(bmp)

        bmp = self.GetColorBitmap(textAttr.GetBackgroundColour())
        self.background.SetBitmap(bmp)

        valAlign = textAttr.GetAlignment()
        if valAlign == wx.TEXT_ALIGNMENT_LEFT:
            strAlign = 'Left'
        elif valAlign == wx.TEXT_ALIGNMENT_CENTRE:
            strAlign = 'Center'
        elif valAlign == wx.TEXT_ALIGNMENT_RIGHT:
            strAlign = 'Right'
        else:
            strAlign = 'Unknown'
        self.parAlign.SetValue(strAlign)

        self.parLeftIndent.SetValue("%0.2f" % (textAttr.GetLeftSubIndent() / convertVal))
        self.parFirstLineIndent.SetValue("%0.2f" % ((textAttr.GetLeftIndent() - textAttr.GetLeftSubIndent()) / convertVal))
        self.parRightIndent.SetValue("%0.2f" % (textAttr.GetRightIndent() / convertVal))

        lineSpacing = textAttr.GetLineSpacing()
        if lineSpacing == rt.TEXT_ATTR_LINE_SPACING_NORMAL:
            strLineSpacing = 'Single'
        elif lineSpacing == rt.TEXT_ATTR_LINE_SPACING_HALF:
            strLineSpacing = '1.5'
        elif lineSpacing == rt.TEXT_ATTR_LINE_SPACING_TWICE:
            strLineSpacing = 'Double'
        elif lineSpacing == 25:
            strLineSpacing = '2.5'
        elif lineSpacing == 30:
            strLineSpacing = 'Triple'
        else:
            strLineSpacing = 'Unknown'
        self.parLineSpacing.SetValue(strLineSpacing)

        self.parBefore.SetValue("%0.2f" % (textAttr.GetParagraphSpacingBefore() / convertVal))
        self.parAfter.SetValue("%0.2f" % (textAttr.GetParagraphSpacingAfter() / convertVal))

        tabs = ''
        for tab in textAttr.GetTabs():
            if tabs != '':
                tabs += ', '
            tabs += "%0.2f" % (tab / convertVal)
        self.tabs.SetValue("%s" % tabs)

    def GetColorBitmap(self, color):
        bmp = wx.EmptyBitmap(16, 16)
        # Create a Device Context for manipulating the bitmap
        dc = wx.BufferedDC(None, bmp)
        # Begin the drawing process
        dc.BeginDrawing()
        # Paint the bitmap white
        dc.SetBackground(wx.Brush(wx.Colour(255, 255, 255)))
        # Clear the device context
        dc.Clear()
        # Define the pen to draw with
        pen = wx.Pen(wx.Colour(0, 0, 0), 1, wx.SOLID)
        # Set the Pen for the Device Context
        dc.SetPen(pen)
        # Define the brush to paint with in the defined color
        brush = wx.Brush(color)
        # Set the Brush for the Device Context
        dc.SetBrush(brush)
        # Draw a black border around the color graphic
        dc.DrawRectangle(0, 0, 15, 15)
        # End the drawing process
        dc.EndDrawing()
        # Select a different object into the Device Context, which allows the bmp to be used.
        dc.SelectObject(wx.EmptyBitmap(5,5))
        
        return bmp
