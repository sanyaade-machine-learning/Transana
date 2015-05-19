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

""" This module implements the TranscriptionUI class as part of the Editors
component. """

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>, Jonathan Beavers <jonathan.beavers@gmail.com>'

DEBUG = False
if DEBUG:
    print "TranscriptionUI DEBUG is ON!!"

# For testing purposes, there can be a small input box that allows the entry of Unicode codes
# for characters.
ALLOW_UNICODE_ENTRY = False
if ALLOW_UNICODE_ENTRY:
    print "TranscriptionUI ALLOW_UNICODE_ENTRY is ON!"

import wx

import gettext
import os
import pickle
from TranscriptToolbar import TranscriptToolbar
from TranscriptEditor import TranscriptEditor
import Dialogs
# Import the Transana Font Dialog
import TransanaFontDialog
import TransanaGlobal


class TranscriptionUI(wx.Dialog):
    """This class manages the graphical user interface for the transcription
    editors component.  It creates the transcript window containing a
    TranscriptToolbar and a TranscriptEditor object."""

    def __init__(self, parent, includeClose=False):
        """Initialize an TranscriptionUI object."""
        self.dlg = _TranscriptDialog(parent, -1, includeClose=includeClose)
        
        self.dlg.toolbar.Enable(0)

        self.transcriptWindowNumber = -1

        # We need to adjust the screen position on the Mac.  I don't know why.
        if "__WXMAC__" in wx.PlatformInfo:
            pos = self.dlg.GetPosition()
            self.dlg.SetPosition((pos[0]-20, pos[1]))
        
# Public methods
    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.dlg.ControlObject=ControlObject
        self.transcriptWindowNumber = len(ControlObject.TranscriptWindow) - 1
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
        self.dlg.editor.ClearDoc()
        self.dlg.toolbar.ClearToolbar()
        self.dlg.toolbar.Enable(0)
        self.dlg.EnableSearch(False)
        
    def LoadTranscript(self, transcriptObj):
        """Load a transcript object."""
        self.dlg.editor.set_read_only(True)
        self.dlg.toolbar.ClearToolbar()
        
        # let's figure out what format the desired transcript was saved as
        if transcriptObj.text != None:
            try:
                # was it plain text?
                if transcriptObj.text[:4] == 'txt\n':
                    self.dlg.editor.load_transcript(transcriptObj, 'text')
                # was it RTF?
                elif transcriptObj.text[2:5] == u'rtf':
                    self.dlg.editor.load_transcript(transcriptObj)
                # or was it pickled?
                else:
                    self.dlg.editor.load_transcript(transcriptObj, 'pickle')
            except UnicodeDecodeError:
                if DEBUG:
                    print "TranscriptionUI.LoadTranscript():"
                    import sys, traceback
                    print sys.exc_info()[0], sys.exc_info()[1]
                    traceback.print_exc()

                # any unicode decoding errors are likely coming from attempting
                # to decode pickled data, so the transcript is most probably
                # pickled.
                self.dlg.editor.load_transcript(transcriptObj, 'pickle')

        self.dlg.toolbar.Enable(1)
        self.dlg.EnableSearch(True)

    def GetCurrentTranscriptObject(self):
        """ Return the current Transcript Object, with the edited text even if it hasn't been saved. """
        # Make a copy of the Transcript Object, since we're going to be changing it.
        tempTranscriptObj = self.dlg.editor.TranscriptObj.duplicate()
        # Update the Transcript Object's text to reflect the edited state
        tempTranscriptObj.text = self.dlg.editor.GetRTFBuffer()
        # Now return the copy of the Transcript Object
        return tempTranscriptObj
        
    def GetDimensions(self):
        (left, top) = self.dlg.GetPositionTuple()
        (width, height) = self.dlg.GetSizeTuple()
        # For some reason, the Mac started displaying the Transcript window at -20 by default.
        if left < 0:
            left = 0
        return (left, top, width, height)

    def GetTranscriptDims(self):
        """Return dimensions of transcript editor component."""
        (left, top) = self.dlg.editor.GetPositionTuple()
        (width, height) = self.dlg.editor.GetSizeTuple()
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
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

    def InsertText(self, text, wxfont=None):
        """Insert text at the current cursor position."""
        self.dlg.editor.InsertStyledText(text)
        
    def InsertTimeCode(self):
        """Insert a timecode at the current transcript position."""
        self.dlg.editor.insert_timecode()

    def InsertSelectionTimeCode(self, start_ms, end_ms):
        """Insert a timecode for the currently selected period in the
        Waveform.  start_ms and end_ms should contain the start and
        end time positions of the selected time period, in milliseconds."""
        self.dlg.editor.insert_timed_pause(start_ms, end_ms)

    def TranscriptModified(self):
        """Return TRUE if transcript was modified since last save."""
        return self.dlg.editor.modified()
        
    def SaveTranscript(self):
        """Save the Transcript to the database."""
        self.dlg.editor.save_transcript()

    def SaveTranscriptAs(self, rtf_fname):
        """Export the Transcript to an RTF file."""
        self.dlg.editor.export_transcript(rtf_fname)
        if "__WXMAC__" in wx.PlatformInfo:
            msg = _('If you load this RTF file into Word on the Macintosh, you need to select "Format" > "AutoFormat...",\nmake sure the "AutoFormat now" option is selected, and press "OK".  Otherwise you will\nlose some Font formatting information from the file when you save it.\n(Courier New font will be changed to Times font anyway.)')
            msg = msg + '\n\n' + _('Also, Word on the Macintosh appears to handle the Whisper (Open Dot) Character for Jeffersonian \nNotation improperly.  You will need to convert this character to Symbol font within Word, but \nconvert it back to Courier New font prior to re-import into Transana.')
            dlg = Dialogs.InfoDialog(self.dlg, msg)
            dlg.ShowModal()
            dlg.Destroy()
        
    def TranscriptUndo(self, event):
        self.get_toolbar().OnUndo(event)

    def TranscriptCut(self):
        """  Pass-through for the Cut() method """
        self.get_editor().Cut()

    def TranscriptCopy(self):
        """  Pass-through for the Copy() method """
        self.get_editor().Copy()

    def TranscriptPaste(self):
        """  Pass-through for the Paste() method """
        self.get_editor().Paste()

    def CallFontDialog(self):
        """  Pass-through for the CallFontDialog() method """
        self.dlg.editor.CallFontDialog()

    def AdjustIndexes(self, adjustmentAmount):
        """ Adjust Transcript Time Codes by the specified amount """
        # If the transcript is "Read-Only", it must be put into "Edit" mode.  Let's do
        # this automatically.
        if self.dlg.editor.get_read_only():
            # First, we "push" the Edit Mode button ...
            self.dlg.toolbar.ToggleTool(self.dlg.toolbar.CMD_READONLY_ID, True)
            # ... then we call the button's method as if it really had been pushed.
            self.dlg.toolbar.OnReadOnlySelect(None)
        # Now adjust the indexes
        self.dlg.editor.AdjustIndexes(adjustmentAmount)

    def UpdateSelectionText(self, text):
        """ Update the text indicating the start and end points of the current selection """
        self.dlg.UpdateSelectionText(text)

    def ChangeLanguages(self):
        """ Change all on-screen prompts to the new language. """
        self.dlg.toolbar.ChangeLanguages()
        self.dlg.ChangeLanguages()
        

# Private methods    

# Public properties

# dependent classes

class _TranscriptDialog(wx.Dialog):

    def __init__(self, parent, id=-1, includeClose=False):
        # If we're including an optional Close button ...
        if includeClose:
            # ... define a style that includes the Close Box.  (System_Menu is required for Close to show on Windows in wxPython.)
            style = wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS | wx.SYSTEM_MENU | wx.CLOSE_BOX
        # If we don't need the close box ...
        else:
            # ... then we don't need that defined in the style
            style = wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS
        # Create the Dialog with the appropriate style
        wx.Dialog.__init__(self, parent, id, _("Transcript"), self.__pos(), self.__size(), style=style)

        # Set "Window Variant" to small only for Mac to make fonts match better
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        self.ControlObject = None            # The ControlObject handles all inter-object communication, initialized to None
        self.transcriptWindowNumber = -1

        # add the widgets to the panel
        sizer = wx.BoxSizer(wx.VERTICAL)

        hsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.toolbar = TranscriptToolbar(self)
        hsizer.Add(self.toolbar, 0, wx.ALIGN_TOP, 10)

        # Add Quick Search tools
        self.CMD_SEARCH_BACK_ID = wx.NewId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_BACK, wx.ART_TOOLBAR, (16,16))
        self.searchBack = wx.BitmapButton(self, self.CMD_SEARCH_BACK_ID, bmp, style=wx.NO_BORDER)
        hsizer.Add(self.searchBack, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        wx.EVT_BUTTON(self, self.CMD_SEARCH_BACK_ID, self.OnSearch)
        hsizer.Add((10, 1), 0)
        self.searchBackToolTip = wx.ToolTip(_("Search backwards"))
        self.searchBack.SetToolTip(self.searchBackToolTip)

        self.searchText = wx.TextCtrl(self, -1, size=(100, 20), style=wx.TE_PROCESS_ENTER)
        hsizer.Add(self.searchText, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnSearch, self.searchText)
        hsizer.Add((10, 1), 0)
        
        self.CMD_SEARCH_NEXT_ID = wx.NewId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_FORWARD, wx.ART_TOOLBAR, (16,16))
        self.searchNext = wx.BitmapButton(self, self.CMD_SEARCH_NEXT_ID, bmp, style=wx.NO_BORDER)
        hsizer.Add(self.searchNext, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        wx.EVT_BUTTON(self, self.CMD_SEARCH_NEXT_ID, self.OnSearch)
        self.searchNextToolTip = wx.ToolTip(_("Search forwards"))
        self.searchNext.SetToolTip(self.searchNextToolTip)

        self.EnableSearch(False)

        if ALLOW_UNICODE_ENTRY:
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
            
        self.toolbar.Realize()
        
        self.editor = TranscriptEditor(self, id, self.toolbar.OnStyleChange, updateSelectionText=True)
        sizer.Add(self.editor, 1, wx.EXPAND, 10)
        if "__WXMAC__" in wx.PlatformInfo:
            # This adds a space at the bottom of the frame on Mac, so that the scroll bar will get the down-scroll arrow.
            sizer.Add((0, 15))
        self.SetSizer(sizer)
        self.SetAutoLayout(True)
        self.Layout()
        # Set the focus in the Editor
        self.editor.SetFocus()

        # Capture Size Changes
        wx.EVT_SIZE(self, self.OnSize)

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
        # Determine the visible line from the starting position's document line, and scroll so that the highlight
        # is 2 lines down, if possible.  (In wxSTC, the position in a document has a Document line, which does not
        # take line wrapping into account, and a visible line, which does.)
        self.editor.ScrollToLine(max(self.editor.VisibleFromDocLine(self.editor.LineFromPosition(start) - 2), 0))

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
        self.searchBack.Enable(enable)
        self.searchText.Enable(enable)
        self.searchNext.Enable(enable)

    def ClearSearch(self):
        """ Clear the Search Box """
        self.searchText.SetValue('')
        
    def OnUnicodeText(self, event):
        s = event.GetString()
        if len(s) > 0:
            if not s[len(s)-1].upper() in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F']:
                self.UnicodeEntry.SetValue(s[:-1])
                self.UnicodeEntry.SetInsertionPoint(len(s)-1)

    def OnUnicodeEnter(self, event):
        try:
            c = unichr(int(self.UnicodeEntry.GetValue(), 16))
            c = c.encode('utf8')  # (TransanaGlobal.encoding)
            curpos = self.editor.GetCurrentPos()
            len = self.editor.GetTextLength()
            self.editor.InsertText(curpos, c)
            len = self.editor.GetTextLength() - len
            self.editor.StartStyling(curpos, 0x7f)
            self.editor.SetStyling(len, self.editor.style)
            self.editor.GotoPos(curpos + len)
            self.UnicodeEntry.SetValue('')
            self.editor.SetFocus()
        except:
            pass

    def UpdateSelectionText(self, text):
        """ Update the text indicating the start and end points of the current selection """
        self.selectionText.SetLabel(text)

    def ChangeLanguages(self):
        """ Change Languages """
        self.searchBackToolTip.SetTip(_("Search backwards"))
        self.searchNextToolTip.SetTip(_("Search forwards"))

    def __size(self):
        """Determine the default size for the Transcript frame."""
        rect = wx.ClientDisplayRect()
        width = rect[2] * .715
        height = (rect[3] - TransanaGlobal.menuHeight) * .74
        return wx.Size(width, height)

    def __pos(self):
        """Determine default position of Transcript Frame."""
        rect = wx.ClientDisplayRect()
        (width, height) = self.__size()
        # rect[0] compensates if Start menu is on Left
        x = rect[0]
        # rect[1] compensates if Start menu is on Top
        y = rect[1] + rect[3] - height - 3
        return (x, y)    
