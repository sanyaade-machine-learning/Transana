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

"""This module implements the TranscriptEditor class as part of the Editors
component.
"""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "TranscriptEditor DEBUG is ON."

import wx
from RichTextEditCtrl import RichTextEditCtrl
import Transcript
import DragAndDropObjects
import Dialogs
import Episode, Clip
import TransanaGlobal
import Misc
import re
import cPickle
import types

# This character is interpreted as a timecode marker in transcripts
TIMECODE_CHAR = '\xA4'
UNI_TIMECODE_CHAR = u'\xA4'

# Nate's original REGEXP, "\xA4<[^<]*>", was not working correctly.
# Given the string "this \xA4<1234> is a string > with many characters",
# Nate's REGEXP returned "\xA4<1234> is a string >"
# rather than the desired "\xA4<1234>".
# My REGEXP "\xA4<[\d]*>" appears to do that.
TIMECODE_REGEXP = "\xA4<[\d]*>"             # "\xA4<[^<]*>"

class TranscriptEditor(RichTextEditCtrl):
    """This class is a word processor for transcribing and editing.  It
    provides only the actual text editing control, without any external GUI
    components to aid editing (such as a toolbar)."""

    def __init__(self, parent, id=-1, stylechange_cb=None):
        """Initialize an TranscriptEditor object."""
        RichTextEditCtrl.__init__(self, parent)

        self.parent = parent
        self.StyleChanged = stylechange_cb

        # There are times related to right-click play control when we need to remember the cursor position.
        # Create a variable to store that information, initialized to 0
        self.cursorPosition = 0

        # These ASCII characters are treated as codes and hidden
        self.HIDDEN_CHARS = [TIMECODE_CHAR,]
        self.HIDDEN_REGEXPS = [re.compile(TIMECODE_REGEXP),]
        self.codes_vis = 0
        self.TranscriptObj = None
        self.timecodes = []
        self.current_timecode = -1
        self.set_read_only(1)
        
        # Remove Drag-and-Drop reference on the mac due to the Quicktime Drag-Drop bug
        if not '__WXMAC__' in wx.PlatformInfo:
            dt = TranscriptEditorDropTarget(self)
            self.SetDropTarget(dt)

            self.SetDragEvent(id, self.OnStartDrag)
            
        # We need to trap both the EVT_KEY_DOWN and the EVT_CHAR event.
        # EVT_KEY_DOWN processes NON-ASCII keys, such as cursor keys and Ctrl-key combinations.
        # All characters are reported as upper-case.
        wx.EVT_KEY_DOWN(self, self.OnKeyPress)
        # EVT_CHAR is used to detect normal typing.  Characters are case sensitive here.
        wx.EVT_CHAR(self, self.OnChar)
        # EVT_LEFT_UP is used to detect the left click positioning in the Transcript.
        wx.EVT_LEFT_UP(self, self.OnLeftUp)
        # This causes the Transana Transcript Window to override the default
        # RichTextEditCtrl right-click menu.  Transana needs the right-click
        # for play control rather than an editing menu.
        wx.EVT_RIGHT_UP(self, self.OnRightClick)


# Public methods
    def load_transcript(self, transcript):
        """Load the given transcript object into the editor.  Pass a RTF
        filename as a string or pass a Transcript object."""
        
        # The Transcript should already have been saved or cleared by this point.
        # This code should never get activated.  It's just here for safety's sake.
        if self.TranscriptObj:
            self.parent.ControlObject.SaveTranscript(1)
      
        self.ClearDoc()
        self.set_read_only(0)

        # Disable widget while loading transcript
        self.Enable(False)

        self.ProgressDlg = wx.ProgressDialog(_("Loading Transcript"), \
                                               _("Reading document stream"), \
                                                maximum=100, \
                                                style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)
        if isinstance(transcript, types.StringTypes):
            self.LoadDocument(transcript)
            self.TranscriptObj = None
        else:
            # Assume a Transcript object was passed 
            self.LoadRTFData(transcript.text)
            self.TranscriptObj = transcript
            self.TranscriptObj.has_changed = 0
        self.ProgressDlg.Destroy()
        
        # Scan transcript for timecodes
        self.load_timecodes()

        # Re-enable widget
        self.Enable(True)
        self.set_read_only(1)
        
        self.GotoPos(0)
        # Clear the Undo Buffer, so you can't Undo past this point!  (BUG FIX!!)
        self.EmptyUndoBuffer()
        # Display Time Codes by default
        self.show_codes()
        self.parent.toolbar.ToggleTool(self.parent.toolbar.CMD_SHOWHIDE_ID, True)
        # Set save point
        self.SetSavePoint()
    
    def load_timecodes(self):
        """Scan the document for timecodes and add to internal list."""
        txt = self.GetText()
        if type(txt) is unicode:
            findstr = UNI_TIMECODE_CHAR + "<"
        else:
            findstr = TIMECODE_CHAR + "<"
        i = txt.find(findstr, 0)
        while i >= 0:
            endi = txt.find(">", i)
            timestr = txt[i+2:endi]
            try:
                self.timecodes.append(int(timestr))
            except:
                pass
            i = txt.find(findstr, i+1)
        
    def save_transcript(self):
        """Save the transcript to the database."""
        # Let's try to remember the cursor position
        self.cursorPosition = (self.GetCurrentPos(), self.GetSelection())
        
        # We can't save with Time Codes showing!  Remember the initial status, and hide them
        # if they are showing.
        initCodesVis = self.codes_vis
        if initCodesVis:
            self.hide_codes()
        
        if self.TranscriptObj:
            self.TranscriptObj.has_changed = self.modified()
            self.TranscriptObj.text = self.GetRTFBuffer()
            self.TranscriptObj.lock_record()
            self.TranscriptObj.db_save()
            self.TranscriptObj.unlock_record()
            # Tell wxSTC that we saved it so it can keep track of
            # modifications.
            self.SetSavePoint()

        # If time codes were showing, show them again.
        if initCodesVis:
            self.show_codes()
        
        # Let's try restoring the Cursor Position when all is said and done.
        self.RestoreCursor()

    def get_transcript(self):
        """Get in memory transcript object."""
        # Do we really want this?

    def get_transcript_doc_data(self):
        return self.GetRTFBuffer()

    def export_transcript(self, rtf_fname):
        """Export the transcript to an RTF file."""
        self.SaveRTFDocument(rtf_fname)
        
    def set_default_font(self, font):
        """Change the default font."""

    def set_font(self, font_face, font_size, font_fg=0x000000, font_bg=0xffffff):
        """Change the current font or the font for the selected text."""
        self.SetFont(font_face, font_size, font_fg, font_bg)
    
    def get_font(self):
        """Get the current font."""
        return self.GetFont()
    
    def set_bold(self, enable=-1):
        """Set bold state for current font or for the selected text.
        If enable is not specified as 0 or 1, then it will toggle the
        current bold state."""
        if self.get_read_only():
            return
        if enable == -1:
            enable = not self.GetBold()
        self.SetBold(enable)
    
    def get_bold(self):
        return self.GetBold()
    
    def set_italic(self, enable=-1):
        """Set italic state for current font or for the selected text."""
        if self.get_read_only():
            return
        if enable == -1:
            enable = not self.GetItalic()
        self.SetItalic(enable)

    def get_italic(self):
        return self.GetItalic()

    def set_underline(self, enable=-1):
        """Set underline state for current font or for the selected text."""
        if self.get_read_only():
            return
        if enable == -1:
            enable = not self.GetUnderline()
        self.SetUnderline(enable)

    def get_underline(self):
        return self.GetUnderline()
        
    def cut_selected_text(self):
        """Delete selected text and place in clipboard."""
    def copy_seleted_text(self):
        """Copy selected text to clipboard."""
    def paste_text(self):
        """Paste text from clipboard."""
    def select_all(self):
        """Select all document text."""
        
    def show_codes(self):
        """Make encoded text in document visible."""
        self.changeTimeCodeHiddenStatus(False)
        self.codes_vis = 1
        
    def hide_codes(self):
        """Make encoded text in document visible."""
        self.changeTimeCodeHiddenStatus(True)
        self.codes_vis = 0

    def codes_visible(self):
        """Return 1 if encoded text is visible."""
        return self.codes_vis

    def changeTimeCodeHiddenStatus(self, hiddenVal):
        """ Changes the Time Code marks (but not the time codes themselves) between visible and invisble styles. """
        # We don't want the screen to move, so let's remember the current position
        topLine = self.GetFirstVisibleLine()
        
        # Note whether the document has had a style change yet.
        initStyleChange = self.stylechange
        
        # Let's try to remember the cursor position
        self.cursorPosition = (self.GetCurrentPos(), self.GetSelection())
        
        # Move the cursor to the beginning of the document
        self.GotoPos(0)

        # Let's show all the hidden text of the time codes.  This doesn't work without it!
        if not self.codes_vis:
            wereHidden = True
            self.show_all_hidden()
        else:
            wereHidden = False

        # Let's find each time code mark and update it with the new style.  
        for loop in range(0, len(self.timecodes)):
            # Find the Timecode
            self.cursor_find('%s' % TIMECODE_CHAR)
            # Note the Cursor's Current Position
            curpos = self.GetCurrentPos()
            # Start Styling from the Cursor Position

            self.StartStyling(curpos, 255)
            # If the TimeCodes should be hidden ...
            if hiddenVal:
                # ... set their style to STYLE_HIDDEN
                self.SetStyling(1, self.STYLE_HIDDEN)
            # If the TimeCodes should be displayed ...
            else:
                # ... set their style to STYLE_TIMECODE
                self.SetStyling(1, self.STYLE_TIMECODE)
            # We then need to move the cursor past the current TimeCode so the next one will be
            # found by the "cursor_find" call.
            self.SetCurrentPos(curpos + 1)

        # We better hide all the hidden text for the time codes again
        if wereHidden:
            self.hide_all_hidden()
            
        # now reset the position of the document
        self.ScrollToLine(topLine)

        # Okay, this might not work because of changes we've made to the transcript, but let's
        # try restoring the Cursor Position when all is said and done.
        self.RestoreCursor()
        self.Update()

        # This event should NOT cause the Style Change indicator to suggest the document has been changed.
        self.stylechange = initStyleChange

        
    def show_all_hidden(self):
        """Make encoded text in document visible."""
        self.StyleSetVisible(self.STYLE_HIDDEN, True)
        self.codes_vis = 1

    def hide_all_hidden(self):
        """Make encoded text in document visible."""
        self.StyleSetVisible(self.STYLE_HIDDEN, False)
        self.codes_vis = 0

    def find_text(self, text, matchcase, wraparound):
        """Find text in document."""
    def insert_text(self, text):
        """Insert text at current cursor position."""
   
    def insert_timecode(self, time_ms=-1):
        """Insert a timecode in the current cursor position of the
        Transcript.  The parameter time_ms is optional and will default
        to the current Video Position if not used."""
        if self.get_read_only():
            # Don't do it in read-only mode
            return
        
        (prevTimeCode, nextTimeCode) = self.get_selected_time_range()
        
        if time_ms >= 0:
            timepos = time_ms
        else:
            timepos = self.parent.ControlObject.GetVideoPosition()

        # The second line here allows you to insert a timecode at 0.0 if there isn't already one there.
        if (len(self.timecodes) == 0) or (prevTimeCode < timepos) and ((timepos < nextTimeCode) or (nextTimeCode == -1)) \
           or ((timepos == 0) and (prevTimeCode == 0.0) and (self.timecodes[0] > 0)):
            if self.codes_vis:
                tempStyle = self.style
                self.style = self.STYLE_TIMECODE
                self.InsertStyledText("%s" % TIMECODE_CHAR)
                self.InsertHiddenText("<%d>" % timepos)
                self.style = tempStyle
            else:
                self.InsertHiddenText("%s<%d>" % (TIMECODE_CHAR, timepos))

            self.Refresh()
            # Update the 'timecodes' list, putting it in the right spot
            i = 0
            if len(self.timecodes) > 0:
                # the index variable (i) must be less than the number of elements in
                # self.timecodes to avoid an index error when inserting a selection timecode
                # at the end of the Transcript.
                while (i < len(self.timecodes) and (self.timecodes[i] < timepos)):
                    i = i + 1
            self.timecodes.insert(i, timepos)
        else:
            msg = _('Time Code Sequence error.\nYou are trying to insert a Time Code at %s\nbetween time codes at %s and %s.') % (Misc.time_in_ms_to_str(timepos), Misc.time_in_ms_to_str(prevTimeCode), Misc.time_in_ms_to_str(nextTimeCode))
            errordlg = Dialogs.ErrorDialog(self, msg)
            errordlg.ShowModal()
            errordlg.Destroy()

    def insert_timed_pause(self, start_ms, end_ms):
        if self.get_read_only():
            # Don't do it in read-only mode
            return
        (prevTimeCode, nextTimeCode) = self.get_selected_time_range()
        # Issue 231 -- Selection Insert fails if it is after the last known time code.
        # If the last time code is undefined, use the media length for an Episode or the
        # Clip Stop Point for a Clip.
        if nextTimeCode == -1:
            if type(self.parent.ControlObject.currentObj) == type(Episode.Episode()):
                nextTimeCode = self.parent.ControlObject.currentObj.tape_length
            else:
                nextTimeCode = self.parent.ControlObject.currentObj.clip_stop
        
        if prevTimeCode > start_ms:
            msg = _('Time Code Sequence error.\nYou are trying to insert a Time Code at %s\nbetween time codes at %s and %s.') % (Misc.time_in_ms_to_str(start_ms), Misc.time_in_ms_to_str(prevTimeCode), Misc.time_in_ms_to_str(nextTimeCode))
            errordlg = Dialogs.ErrorDialog(self, msg)
            errordlg.ShowModal()
            errordlg.Destroy()
        elif nextTimeCode < end_ms:
            msg = _('Time Code Sequence error.\nYou are trying to insert a Time Code at %s\nbetween time codes at %s and %s.') % (Misc.time_in_ms_to_str(end_ms), Misc.time_in_ms_to_str(prevTimeCode), Misc.time_in_ms_to_str(nextTimeCode))
            errordlg = Dialogs.ErrorDialog(self, msg)
            errordlg.ShowModal()
            errordlg.Destroy()
        else:
            time_span_secs = (end_ms - start_ms) / 1000.0
            time_span_secs = round(time_span_secs, 1)
            self.insert_timecode(start_ms)
            timespan = "(%.1f)" % time_span_secs
            self.InsertText(self.GetCurrentPos(), timespan)
            self.GotoPos(self.GetCurrentPos() + len(timespan))
            self.insert_timecode(end_ms)


    def scroll_to_time(self, ms):
        """Scroll to the nearest timecode matching 'ms' that isn't
        greater than 'ms'.  Return TRUE if position in document is
        actually changed."""
        # Bump up the time a little bit to account for rounding errors,
        # since sometimes I've noticed the numbers that get passed here
        # are off by one.  Otherwise it gets stuck, for example, when
        # the next timecode is X and ms is X-1, and other code depends
        # on it actually updating (which happened with the CTRL-N
        # behavior to jump to the next segment)
        ms = ms + 5

        # If no timecodes exist, ignore this call
        if len(self.timecodes) == 0:
            return False
        
        # Find the timecode that's closest
        closest_time = self.timecodes[0]
        for timecode in self.timecodes:
            if (timecode <= ms) and (ms - timecode < ms - closest_time):
                closest_time = timecode
            #if abs(timecode - ms) < abs(closest_time - ms):
            #    closest_time = timecode
        
        # Check if ALL timecodes in document are higher than given time.
        # In this case, we scroll to 0
        if (closest_time > ms):
            closest_time = 0

        # Get start and end points of current selection
        (start, end) = self.GetSelection()
        # If timecode position hasn't changed and a selection is highlighted rather than a cursor being displayed,
        # we don't need to change the current selection.
        if (closest_time == self.current_timecode) and (start != end):
            # Do nothing if the closest timecode hasn't changed
            return False
      
        self.current_timecode = closest_time

        # Update cursor position and select text up until the next timecode
        pos = self.GetCurrentPos()
        self.cursor_find("%s<%d>" % (TIMECODE_CHAR, closest_time))
        self.select_find(TIMECODE_CHAR + "<")
        if self.GetCurrentPos() != pos:
            return True  # return TRUE since position changed
        else:
            return False
        
    def cursor_find(self, text):
        """Move the cursor to the next occurrence of given text in the
        transcript (for word tracking)."""
        # We first try searching from the current cursor position
        # for efficiency reasons (most of the time you're jumping just
        # ahead of the current cursor position).
        pos = self.FindText(self.GetCurrentPos(), self.GetLength()-1, text) 
        if pos >= 0:
            self.GotoPos(pos)
        else:
            # Try searching in reverse
            pos = self.FindText(self.GetCurrentPos(), 0, text)
            if pos >= 0:
                self.GotoPos(pos)

    def select_find(self, text):
        """Select the text from the current cursor position to the next
        occurrence of given text (for word tracking)."""
        # In some cases, the text is a defined time code.  In others, it's a time indicator not in the timecode list.
        # We have to make sure it's something in the list.

        # First, not all "text" refers to time codes.  Some is actual text, including the time code symbol.
        # The try ... except structure catches this (as the text will not convert to an integer) and leaves the text unchanged.
        # If the text IS a time code, this code will locate the appropriate time code AFTER the number sent.  This is needed
        # when a selection is made in the Visualization Window which may not align with a know ending time code.
        try:
            timecodePos = 0
            while (timecodePos < len(self.timecodes)) and (int(self.timecodes[timecodePos]) < int(float(text))):
                timecodePos += 1
            if text != str(self.timecodes[timecodePos]):
                text = str(self.timecodes[timecodePos])
        except:
            pass
        
        curpos = self.GetCurrentPos()
        endpos = self.FindText(curpos + 1, self.GetLength()-1, text)
        
        # If not found, select until the end of the document
        if endpos ==  -1:
            endpos = self.GetLength()-1

        # The selection process should not include the initial time code.
        # Set a flag to indicate whether we are in the midst of a time code indicator, which is intially false
        inTimeCode = False
        # Set a flag to signal that we are not yet done, which initially is true
        contin = True
        # While we're either not done or we're inside a time code ...
        while contin or inTimeCode:
            # Assume that we are finished unless it is proven otherwise
            contin = False
            # If the current character is the Time Code itself ...
            if self.GetCharAt(curpos) == ord(TIMECODE_CHAR):
                # ... move on to the next character ...
                curpos += 1
                # ... and indicate that we're not done yet.
                contin = True
            # if we are at the character signalling the start of a time code ...
            elif self.GetCharAt(curpos) == ord('<'):
                # ... move on to the next character ...
                curpos += 1
                # ... and indicate that we are within a time code
                inTimeCode = True
            # If we are at the character signalling the end of a time code ...
            elif self.GetCharAt(curpos) == ord('>'):
                # ... move on to the next character ...
                curpos += 1
                # ... and only if we are currently within the time code ...
                if inTimeCode:
                    # ... indicate that the time code is over.
                    inTimeCode = False
            # If we're inside a time code but not at the end of it ...
            elif inTimeCode:
                # ... move on to the next character.
                curpos += 1
                
        # When searching for time codes for positioning of the selection, we end up selecting
        # the time code symbol and "<" that starts the time code.  We don't want to do that.
        while (endpos > 1) and (self.GetCharAt(endpos-1) in [ord('<'), ord(TIMECODE_CHAR)]):
            endpos -= 1

        self.SetSelection(endpos, curpos)
        # NOTE:  STC seems to make a distinction between VISIBLE lines and DOCUMENT lines.
        # we need Visible Lines for scrolling the current position to match the video.
        startline = self.VisibleFromDocLine(self.LineFromPosition(curpos))
        endline = self.VisibleFromDocLine(self.LineFromPosition(endpos))
        
        # Attempt to make the selection visible by scrolling to make
        # the end line visible, and then the start line visible.
        # (so in the worst case, the start is still always visible)
        # We make sure endline+1 is visible because with long lines
        # that span multiple visible lines, it will only ensure that
        # the beginning is visible.
        if startline < self.GetFirstVisibleLine():
            self.ScrollToLine(startline)
        if endline + 1 > self.GetFirstVisibleLine() + self.LinesOnScreen():
            self.ScrollToLine(endline - self.LinesOnScreen() + 2)
        

    def spell_check(self):
        """Interactively spell-check document."""
    
    def undo(self):
        """Undo last operation(s)."""
        self.Undo()
        
    def redo(self):
        """Redo last undone operation(s)."""
        self.Redo()

    def modified(self):
        """Return TRUE if transcript was modified since last save.
        If no transcript is loaded, this will always return FALSE."""
        return self.TranscriptObj and self.GetModify()

    def set_read_only(self, state=1):
        """Enable or disable read-only mode, to prevent the transcript
        from being modified."""
        self.SetReadOnly(state)

    def get_read_only(self):
        return self.GetReadOnly()

    def find_timecode_before_cursor(self, pos):
        """Return the position of the first timecode before the given
        cursor position."""
        return self.FindText(pos, 0, "%s<" % TIMECODE_CHAR)
        
    def get_selected_time_range(self):
        """Get the time range of the currently selected text.  Return a tuple
        with the start and end times in milliseconds."""
        
        # Default start/end time is 0
        start_timecode = 0
        end_timecode = 0

        # Determine current selection indices
        selstart = self.GetSelectionStart()
        selend = self.GetSelectionEnd()
        if selstart > selend:
            # Need to swap start/end
            temp = selstart
            selstart = selend
            selend = selstart

        # Setup for searching transcript for timecodes
        txt = self.GetText()
        if type(txt) is unicode:
            findstr = UNI_TIMECODE_CHAR + "<"
        else:
            findstr = TIMECODE_CHAR + "<"

        # Extract the start timecode
        # Finds first timecode BEFORE selection start
        pos = txt.rfind(findstr, 0, selstart)
        if pos >= 0:
            endi = txt.find(">", pos)
            timestr = txt[pos+2:endi]
            
            try:
                start_timecode = int(timestr)
            except:
                pass
            
        # Extract the end timecode
        # Finds first timecode AFTER selection end
        pos = txt.find(findstr, selend)
        if pos >= 0:
            endi = txt.find(">", pos)
            timestr = txt[pos+2:endi]
            
            try:
                end_timecode = int(timestr)
            except:
                pass
        # If no later time code is found return -1 
        else:
            end_timecode = -1
        
        return (start_timecode, end_timecode)

    def ClearDoc(self):
        # I think we want to be in read-only always at the end of ClearDoc()
        #state = self.get_read_only()
        self.set_read_only(0)
        RichTextEditCtrl.ClearDoc(self)
        self.TimePosition = 0
        self.TranscriptObj = None
        self.timecodes = []
        self.current_timecode = -1
        #self.set_read_only(state)
        self.set_read_only(1)
        self.codes_vis = 0

    def PrevTimeCode(self, tc=None):
        """Return the timecode immediately before the current one."""
        if tc == None:
            tc = self.current_timecode
        i = self.timecodes.index(tc) - 1
        while i >= 0:
            # Sometimes we might have multiple timecodes with the same value,
            # so we have to do this to ensure we really get a lower timecode
            if self.timecodes[i] < tc:
                return self.timecodes[i]
            i = i - 1
        return self.timecodes[i]

    def NextTimeCode(self, tc=None):
        """Return the timecode immediately after the current one."""
        if tc == None:
            tc = self.current_timecode
        i = self.timecodes.index(tc) + 1
        while i < len(self.timecodes):
            # Sometimes we might have multiple timecodes with the same value,
            # so we have to do this to ensure we really get a higher timecode
            if self.timecodes[i] > tc:
                return self.timecodes[i]
            i = i + 1
        return self.timecodes[i]

    def InsertSymbol(self, ch):
        # Save the current font
        f = self.get_font()
        # Use symbol font
        self.set_font("Symbol", TransanaGlobal.configData.defaultFontSize, 0x000000)
        self.InsertStyledText(ch)
        # restore previous font
        self.set_font(f[0], f[1], f[2])

    def OnKeyPress(self, event):
        """ Called when a key is pressed down.  All characters are upper case.  """

        # Let's try to remember the cursor position.  (We've had problems wiht the cursor moving during transcription on the Mac.)
        self.cursorPosition = (self.GetCurrentPos(), self.GetSelection())

        if DEBUG:
            print "Start:", self.GetCurrentPos(), self.GetSelection()

        # if not event.ControlDown():
        #     dlg = wx.MessageDialog(self, 'Start of OnKeyPress')
        #     dlg.ShowModal()
        #     dlg.Destroy()

        # It might be necessary to block the "event.Skip()" call.  Assume for the moment that it is not.
        blockSkip = False
        if event.ControlDown():
            try:
                c = event.GetKeyCode()
                
                # NOTE:  NON-ASCII keys must be processed first, as the chr(c) call raises an exception!
                if c == wx.WXK_UP:
                    # Ctrl-Cursor-Up inserts the Up Arrow / Rising Intonation symbol
                    ch = '\xAD'
                    self.InsertSymbol(ch)
                    return
                elif c == wx.WXK_DOWN:
                    # Ctrl-Cursor-Down inserts the Down Arrow / Falling Intonation symbol
                    ch = '\xAF'
                    self.InsertSymbol(ch)
                    return
                elif chr(c) == "O":
                    # Ctrl-O inserts the Open Dot / Whispered Speech symbol
                    ch = chr(176)
                    self.InsertSymbol(ch)
                    return
                elif chr(c) == "H":
                    # Ctrl-H inserts the High Dot / Inbreath symbol
                    ch = '\xB7'
                    self.InsertSymbol(ch)
                    return
                elif chr(c) == "B":
                    self.set_bold()
                    self.StyleChanged(self)
                elif chr(c) == "U":
                    self.set_underline()
                    self.StyleChanged(self)
                elif chr(c) == "I":
                    self.set_italic()
                    self.StyleChanged(self)
                elif chr(c) == "T":
                    # CTRL-T pressed
                    self.insert_timecode()
                    return
                elif chr(c) == "S":
                    if not self.parent.ControlObject.IsPlaying():
                        # Explicitly tell Transana to play to the end of the Episode/Clip
                        self.parent.ControlObject.SetVideoEndPoint(-1)
                    # CTRL-S: Play/Pause with auto-rewind
                    self.parent.ControlObject.PlayPause(1)
                    return
                elif chr(c) == "D":
                    if not self.parent.ControlObject.IsPlaying():
                        # Explicitly tell Transana to play to the end of the Episode/Clip
                        self.parent.ControlObject.SetVideoEndPoint(-1)
                    # CTRL-D: Play/Pause without rewinding
                    self.parent.ControlObject.PlayPause(0)
                    return
                elif chr(c) == "A":
                    # CTRL-A: Rewind the video by 10 seconds
                    vpos = self.parent.ControlObject.GetVideoPosition()
                    self.parent.ControlObject.SetVideoStartPoint(vpos-10000)
                    # Explicitly tell Transana to play to the end of the Episode/Clip
                    self.parent.ControlObject.SetVideoEndPoint(-1)
                    # Play should always be initiated on Ctrl-A
                    self.parent.ControlObject.Play(0)
                    return
                elif chr(c) == "F":
                    # CTRL-F: Advance video by 10 seconds
                    vpos = self.parent.ControlObject.GetVideoPosition()
                    self.parent.ControlObject.SetVideoStartPoint(vpos+10000)
                    # Explicitly tell Transana to play to the end of the Episode/Clip
                    self.parent.ControlObject.SetVideoEndPoint(-1)
                    # Play should always be initiated on Ctrl-F
                    self.parent.ControlObject.Play(0)
                    return
                elif chr(c) == "P":
                    # CTRL-P: Previous segment
                    start_timecode = self.PrevTimeCode()
                    self.parent.ControlObject.SetVideoStartPoint(start_timecode)
                    # Explicitly tell Transana to play to the end of the Episode/Clip
                    self.parent.ControlObject.SetVideoEndPoint(-1)
                    # Play should always be initiated on Ctrl-P
                    self.parent.ControlObject.Play(0)
                    return
                elif chr(c) == "N":
                    # CTRL-N: Next segment
                    start_timecode = self.NextTimeCode()
                    self.parent.ControlObject.SetVideoStartPoint(start_timecode)
                    # Explicitly tell Transana to play to the end of the Episode/Clip
                    self.parent.ControlObject.SetVideoEndPoint(-1)
                    # Play should always be initiated on Ctrl-P
                    self.parent.ControlObject.Play(0)
                    return
            except:
                pass    # Non-ASCII value key pressed
        else:
            # Because of time codes and hidden text, we need a bit of extra code here to make sure the cursor is not left in
            # the middle of hidden text and to prevent accidental deletion of hidden time codes.
            c = event.GetKeyCode()
            curpos = self.GetCurrentPos()
            cursel = self.GetSelection()

            # If the are moving to the LEFT with the cursor ...
            if (c == wx.WXK_LEFT):
                # ... and we come to a TIMECODE Character ...
                if (curpos > 0) and (chr(self.GetCharAt(curpos - 1)) == '>') and (self.GetStyleAt(curpos - 1) == self.STYLE_HIDDEN):
                    # ... then we need to find the start of the time code data, signalled by the TIMECODE character ...
                    while (chr(self.GetCharAt(curpos - 1)) != TIMECODE_CHAR):
                        curpos -= 1
                    # If Time Codes are not visible, we need one more character here.  It doesn't make sense
                    # to me, as we should be at the end of the time code data, but we DO need this.
                    if not(self.codes_vis):
                        curpos -= 1

                    # If you cursor over a time code while making a selection, the selection was getting lost with
                    # the original code.  Instead, determine if a selection is being made, and if so, make a new
                    # selection appropriately.
                    
                    # If these values differ, we're selecting rather than merely moving.
                    if cursel[0] == cursel[1]:
                        self.GotoPos(curpos)
                    else:
                        # The selection must be made in this order, or the cursor is moved to the END rather than being
                        # left at the beginning of the selection where it belongs!
                        self.SetSelection(cursel[1], curpos)
                    
            # If the are moving to the RIGHT with the cursor ...
            elif (c == wx.WXK_RIGHT):
                # ... and we come to a TIMECODE Character ...
                if chr(self.GetCharAt(curpos)) == TIMECODE_CHAR:
                    # ... then we need to find the end of the time code data, signalled by the '>' character ...
                    while chr(self.GetCharAt(curpos)) != '>':
                        curpos += 1
                    # If Time Codes are not visible, we need one more character here.  It doesn't make sense
                    # to me, as we should be at the end of the time code data, but we DO need this.
                    if not(self.codes_vis):
                        curpos += 1

                    # If you cursor over a time code while making a selection, the selection was getting lost with
                    # the original code.  Instead, determine if a selection is being made, and if so, make a new
                    # selection appropriately.
                    
                    # If these values differ, we're selecting rather than merely moving.
                    if cursel[0] == cursel[1]:
                        # Position the cursor after the hidden timecode data
                        self.GotoPos(curpos)
                    else:
                        self.SetSelection(cursel[0], curpos)
                    
            # DELETE KEY pressed
            elif (c == wx.WXK_DELETE):
                # First, we need to determine if we are deleting a single character or a selection in the transcript.
                (selStart, selEnd) = self.GetSelection()
                # If selStart and selEnd are the same, we are deleting a character.
                if (selStart == selEnd):
                    # Are we in Edit Mode?  Are we trying to delete a Time Code?
                    if not(self.get_read_only()) and (chr(self.GetCharAt(curpos)) == TIMECODE_CHAR):
                        # If deleting a Time Code, first we determine the time code data for the current position
                        (prevTimeCode, nextTimeCode) = self.get_selected_time_range()
                        # Prompt the user about deleting the Time Code
                        msg = _('Do you want to delete the Time Code at %s?') % Misc.time_in_ms_to_str(nextTimeCode)
                        dlg = wx.MessageDialog(self, msg, _('Transana Confirmation'), wx.YES_NO | wx.ICON_QUESTION | wx.CENTRE)
                        # If the user really does want to delete the time code ...
                        if dlg.ShowModal() == wx.ID_YES:
                            # We need to remember the current state of self.codes_vis, whether time codes are visible or not.
                            # (This variable gets updated when we call show_all_hidden() and may no longer be accurate.)
                            codes_vis = self.codes_vis
                            # Let's show all the hidden text of the time codes.  This doesn't work without it!
                            self.show_all_hidden()

                            # curpos is the start of our time code, but we need to find the end.
                            # we'll start looking just to the right of the time code.
                            selEnd = curpos + 1
                            # We'll keep looking until we find the ">" character, which closes the time code data.
                            while chr(self.GetCharAt(selEnd)) != '>':
                                selEnd += 1
                            # Set the RichTextCtrl (wxSTC) selection to encompass the full time code
                            self.SetSelection(curpos, selEnd)
                            # Replace the Time Code with nothing to delete it.
                            self.ReplaceSelection('')
                            # We need to remove the time code from the self.timecodes List too.
                            # First we locate that entry ...
                            index = self.timecodes.index(nextTimeCode)
                            # ... then we remove it from the list.
                            self.timecodes = self.timecodes[:index] + self.timecodes[index+1:]
                            # Now we reset the Text Cursor
                            self.SetCurrentPos(curpos)
                            # We better hide all the hidden text for the time code data again                           
                            self.hide_all_hidden()
                            # and we need to reset self.codes_vis to its original state.  (This variable gets updated
                            # when we call hide_all_hidden() and may no longer be accurate.)
                            self.codes_vis = codes_vis

                        else:
                            # We need to block the Skip call so that the key event is not passed up to this control's parent for
                            # processing if the user decides not to delete the time code.
                            blockSkip = True
                        # Now we need to destroy the dialog box
                        dlg.Destroy()
                        
                # Otherwise, we're deleting a Selection
                else:
                    # First, let's set the wxSTC "Target" area to the selection
                    self.SetTargetStart(selStart)
                    self.SetTargetEnd(selEnd)
                    # Let's determine if there is a Time Code in the selection
                    timeCodeSearch = self.SearchInTarget(TIMECODE_CHAR)
                    # A value of -1 indicates no Time Code.  Otherwise there is one (or more).
                    if timeCodeSearch != -1:
                        # Prompt the user about deleting the Time Code
                        msg = _('Your current selection contains at least one Time Code.\nAre you sure you want to delete it?')
                        dlg = wx.MessageDialog(self, msg, _('Transana Confirmation'), wx.YES_NO | wx.ICON_QUESTION | wx.CENTRE)
                        # If the user really does want to delete the time code ...
                        if dlg.ShowModal() == wx.ID_YES:
                            # We need to remember the current state of self.codes_vis, whether time codes are visible or not.
                            # (This variable gets updated when we call show_all_hidden() and may no longer be accurate.)
                            codes_vis = self.codes_vis
                            # Let's show all the hidden text of the time codes.  This doesn't work without it!
                            self.show_all_hidden()

                            # The SearchInTarget command appears to have changed our Traget Text, so let's set the wxSTC
                            # "Target" area to the selection again.
                            self.SetTargetStart(selStart)
                            self.SetTargetEnd(selEnd)
                            # Replace the Target Text with nothing to delete it.
                            self.ReplaceTarget('')
                            # We need to remove the time code from the self.timecodes List too.
                            # First we determine the time code data for the current position, which should
                            # give us the time code before the deleted segment and the time code after the
                            # deleted segment.
                            (prevTimeCode, nextTimeCode) = self.get_selected_time_range()
                            # First we locate the indexes for the previous and next time codes
                            startIndex = self.timecodes.index(prevTimeCode)
                            endIndex= self.timecodes.index(nextTimeCode)
                            # ... then we remove items from the list that fall between them.
                            self.timecodes = self.timecodes[:startIndex + 1] + self.timecodes[endIndex:]
                            # Now we reset the Text Cursor
                            self.SetCurrentPos(selStart)
                            # We better hide all the hidden text for the time code data again                           
                            self.hide_all_hidden()
                            # and we need to reset self.codes_vis to its original state.  (This variable gets updated
                            # when we call hide_all_hidden() and may no longer be accurate.)
                            self.codes_vis = codes_vis

                        # We need to block the Skip call so that the key event is not passed up to this control's parent for
                        # processing.  The delete is handled locally or is declined by the user.
                        blockSkip = True
                        
                        # Now we need to destroy the dialog box
                        dlg.Destroy()

            # BACKSPACE KEY pressed
            elif (c == wx.WXK_BACK):
                # First, we need to determine if we are backspacing over a single character or a selection in the transcript.
                (selStart, selEnd) = self.GetSelection()
                # If selStart and selEnd are the same, we are backspacing over a character.
                if (selStart == selEnd):
                    # If we're in Edit Mode and the cursor is not at 0 (where you can't backspace) and
                    # you are backspacing over a hidden '>' character, indicating the end of Time Code data ...
                    if not(self.get_read_only()) and (curpos > 0) and \
                       (chr(self.GetCharAt(curpos - 1)) == '>') and (self.GetStyleAt(curpos - 1) == self.STYLE_HIDDEN):

                        # Under some odd set of circumstances, the characters "BS" appear in the transcript upon backspacing!
                        # We can't detect the BS here.  Therefore, we have to let it appear, and then remove it later.
                        # BSCurpos = self.GetCurrentPos()
                        # If deleting a Time Code, first we determine the time code data for the current position
                        (prevTimeCode, nextTimeCode) = self.get_selected_time_range()
                        # Prompt the user about deleting the Time Code
                        msg = _('Do you want to delete the Time Code at %s?') % Misc.time_in_ms_to_str(prevTimeCode)
                        dlg = wx.MessageDialog(self, msg, _('Transana Confirmation'), wx.YES_NO | wx.ICON_QUESTION | wx.CENTRE)
                        # If the user really does want to delete the time code ...
                        if dlg.ShowModal() == wx.ID_YES:
                            # We need to remember the current state of self.codes_vis, whether time codes are visible or not.
                            # (This variable gets updated when we call show_all_hidden() and may no longer be accurate.)
                            codes_vis = self.codes_vis
                            # Let's show all the hidden text of the time codes.  This doesn't work without it!
                            self.show_all_hidden()

                            # curpos is the end of our time code, but we need to find the start.
                            # we'll start looking where we are.
                            selStart = curpos
                            selEnd = curpos 
                            # keep moving to the left until we find the Time Code character
                            while chr(self.GetCharAt(selStart)) != TIMECODE_CHAR:
                                selStart -= 1
                            # Set the RichTextCtrl (wxSTC) selection to encompass the full time code
                            self.SetSelection(selStart, selEnd)
                            # Replace the Time Code with nothing to delete it.
                            self.ReplaceSelection('')
                            # We need to remove the time code from the self.timecodes List too.
                            # First we locate that entry ...
                            index = self.timecodes.index(prevTimeCode)
                            # ... then we remove it from the list.
                            self.timecodes = self.timecodes[:index] + self.timecodes[index+1:]
                            # We better hide all the hidden text for the time codes again
                            self.hide_all_hidden()
                            # and we need to reset self.codes_vis to its original state.  (This variable gets updated
                            # when we call hide_all_hidden() and may no longer be accurate.)
                            self.codes_vis = codes_vis

                        # Okay, this is weird.  I suspect a bug in wx.STC.
                        # If you insert a Time Code, then Backspace over it, the letters "BS" get added to the Transcript.  This is
                        # obviously not acceptable.  I've added some code here to try to detect and prevent this from showing up.
                        # I'd rather this code appeared above the MessageDialog, but it can't because it won't detect the Backspace
                        # character yet up there.  Bummer.

                        BSCurpos = self.GetCurrentPos()
                        # This detects the BS when the user elects NOT to remove the Time Code
                        if self.GetCharAt(BSCurpos-1) == 8:
                            self.SetSelection(BSCurpos-1, BSCurpos)
                            self.ReplaceSelection('')
                        # This detects the BS when the user elects to remove the Time Code
                        elif self.GetCharAt(BSCurpos) == 8:
                            self.SetSelection(BSCurpos, BSCurpos+1)
                            self.ReplaceSelection('')

                        # We need to block the Skip call so that the key event is not passed up to this control's parent for
                        # processing.  We do this regardless of the user response in the dialog, as the deletion is handled locally.
                        blockSkip = True
                            
                        dlg.Destroy()
                        
                # Otherwise, we're backspacing over a Selection
                else:
                    # First, let's set the wxSTC "Target" area to the selection
                    self.SetTargetStart(selStart)
                    self.SetTargetEnd(selEnd)
                    # Let's determine if there is a Time Code in the selection
                    timeCodeSearch = self.SearchInTarget(TIMECODE_CHAR)
                    # A value of -1 indicates no Time Code.  Otherwise there is one (or more).
                    if timeCodeSearch != -1:
                        # Prompt the user about deleting the Time Code
                        msg = _('Your current selection contains at least one Time Code.\nAre you sure you want to delete it?')
                        dlg = wx.MessageDialog(self, msg, _('Transana Confirmation'), wx.YES_NO | wx.ICON_QUESTION | wx.CENTRE)
                        # If the user really does want to delete the time code ...
                        if dlg.ShowModal() == wx.ID_YES:
                            # We need to remember the current state of self.codes_vis, whether time codes are visible or not.
                            # (This variable gets updated when we call show_all_hidden() and may no longer be accurate.)
                            codes_vis = self.codes_vis
                            # Let's show all the hidden text of the time codes.  This doesn't work without it!
                            self.show_all_hidden()

                            # The SearchInTarget command appears to have changed our Traget Text, so let's set the wxSTC
                            # "Target" area to the selection again.
                            self.SetTargetStart(selStart)
                            self.SetTargetEnd(selEnd)
                            # Replace the Target Text with nothing to delete it.
                            self.ReplaceTarget('')
                            # We need to remove the time code from the self.timecodes List too.
                            # First we determine the time code data for the current position, which should
                            # give us the time code before the deleted segment and the time code after the
                            # deleted segment.
                            (prevTimeCode, nextTimeCode) = self.get_selected_time_range()
                            # First we locate the indexes for the previous and next time codes
                            startIndex = self.timecodes.index(prevTimeCode)
                            endIndex= self.timecodes.index(nextTimeCode)
                            # ... then we remove items from the list that fall between them.
                            self.timecodes = self.timecodes[:startIndex + 1] + self.timecodes[endIndex:]
                            # Now we reset the Text Cursor
                            self.SetCurrentPos(selStart)
                            # We better hide all the hidden text for the time code data again                           
                            self.hide_all_hidden()
                            # and we need to reset self.codes_vis to its original state.  (This variable gets updated
                            # when we call hide_all_hidden() and may no longer be accurate.)
                            self.codes_vis = codes_vis

                        # We need to block the Skip call so that the key event is not passed up to this control's parent for
                        # processing.  We do this regardless of the user response in the dialog, as the deletion is handled locally.
                        blockSkip = True
                        
                        # Now we need to destroy the dialog box
                        dlg.Destroy()

        if not(blockSkip):
            event.Skip()
            
        if DEBUG:
            print "End:", self.GetCurrentPos(), self.GetSelection()
            print

    def OnChar(self, event):
        """Called when a character key is pressed.  Works with case-sensitive characters.  """
        # It might be necessary to block the "event.Skip()" call.  Assume for the moment that it is not.
        blockSkip = False
        # We only need to deal with characters here.
        try:
            ch = chr(event.GetKeyCode())
            # First, we need to determine if we are deleting a single character or a selection in the transcript.
            (selStart, selEnd) = self.GetSelection()
            # If selStart and selEnd are different, we have a selection.  We also need to be in Edit mode.
            if not(self.get_read_only()) and (selStart != selEnd):
                # First, let's set the wxSTC "Target" area to the selection
                self.SetTargetStart(selStart)
                self.SetTargetEnd(selEnd)
                # Let's determine if there is a Time Code in the selection
                timeCodeSearch = self.SearchInTarget(TIMECODE_CHAR)
                # A value of -1 indicates no Time Code.  Otherwise there is one (or more).
                if timeCodeSearch != -1:
                    # Prompt the user about deleting the Time Code
                    msg = _('Your current selection contains at least one Time Code.\nAre you sure you want to delete it?')
                    dlg = wx.MessageDialog(self, msg, _('Transana Confirmation'), wx.YES_NO | wx.ICON_QUESTION | wx.CENTRE)
                    # If the user really does want to delete the time code ...
                    if dlg.ShowModal() == wx.ID_YES:
                        self.SetTargetStart(selStart)
                        self.SetTargetEnd(selEnd)
                        # Replace the Target Text with nothing to delete it.
                        self.ReplaceTarget(ch)
                        self.GotoPos(selStart + 1)
                    blockSkip = True
        except:
            pass

        if not(blockSkip):
            event.Skip()
        
    def OnStartDrag(self, event):
        """Called on the initiation of a Drag within the Transcript."""
        if not self.TranscriptObj:
            # No transcript loaded, abort
            return

        # We need to make sure the cursor is not positioned between a time code symbol and the time code data, which unfortunately
        # can happen.  Preventing this is the sole function of this section of this method.

        # First, see if we have a click Position or a click-drag Selection.
        if (self.GetSelectionStart() != self.GetSelectionEnd()):
            
            # We have a click-drag Selection.  We need to check the start and the end points.
            selStart = self.GetSelectionStart()
            selEnd = self.GetSelectionEnd()

            # Let's see if any change is made.
            selChanged = False
            # Let's see if the start of the selection falls between a Time Code and its data.
            # We can also check to see if the first character is a time code, in which case it should be excluded too!
            if (selStart > 0) and (((chr(self.GetCharAt(selStart - 1)) == TIMECODE_CHAR) and (chr(self.GetCharAt(selStart)) == '<')) or (chr(self.GetCharAt(selStart)) == TIMECODE_CHAR)):
                # If so, we have a change
                selChanged = True
                # Let's find the position of the end of the Time Code
                while chr(self.GetCharAt(selStart - 1)) != '>':
                    selStart += 1

            # Let's see if the end of the selection falls between a Time Code and its data
            if (selEnd > 0) and (((chr(self.GetCharAt(selEnd - 1)) == TIMECODE_CHAR) and (chr(self.GetCharAt(selEnd)) == '<')) or ((chr(self.GetCharAt(selEnd - 1)) == '>') and (self.STYLE_HIDDEN == self.GetStyleAt(selEnd - 1)))):
                # If so, we have a change
                selChanged = True
                # Let's find the position before the Time Code
                while chr(self.GetCharAt(selEnd)) != TIMECODE_CHAR:
                    selEnd -= 1

            # self.SetSelection(selStart, selEnd) does not work if selEnd is between the timecode and the data.  We must use
            # CallAfter to position this correctly.  I'm not sure WHY, but it works.
            # We only need to change the Selection if one of the ends was between a Time Code and its data.
            if selChanged:
                # Remember the new Selection points
                self.selection = (selStart, selEnd)

                self.SetSelection(self.selection[0], self.selection[1])

        (start_time, end_time) = self.get_selected_time_range()

        if end_time == -1:
            end_time = self.parent.ControlObject.GetVideoEndPoint()

        data = DragAndDropObjects.ClipDragDropData(self.TranscriptObj.number, self.TranscriptObj.episode_num, \
                start_time, end_time, self.GetRTFBuffer(select_only=1))

        pdata = cPickle.dumps(data, 1)
        cdo = wx.CustomDataObject(wx.CustomDataFormat("ClipDragDropData"))
        # Put the pickled data object in the wxCustomDataObject
        cdo.SetData(pdata)
        
        if event.GetId() == self.parent.toolbar.CMD_CLIP_ID:
            wx.TheClipboard.SetData(cdo)
        else:
            # Put the data in the DropSource object
            tds = TranscriptDropSource(self.parent)
            tds.SetData(cdo)

            # Initiate the drag operation.
            # NOTE:  Trying to use a value of wx.Drag_CopyOnly to resolve a Mac bug (which it didn't) caused
            #        Windows to stop allowing Clip Creation!
            dragResult = tds.DoDragDrop(wx.Drag_AllowMove)

            # This is a HORRIBLE, EVIL HACK and I hang my head in shame.  I've also spent two days on this one and have not been able to 
            # find any other way to handle it.
            # On the Mac, if you are in Edit Mode and create a Clip, the text in your Transcript gets cut.  There doesn't seem to be anything
            # you can do to prevent it.  Telling the STC it can CopyOnly has no effect.  Changing the dragResult to any legal value
            # at any point in the process has no effect.  Trying to invoke the STC Event using event.Skip() has no effect.
            # The only thing I've found to do is to detect the circumstances, and tell the RichTextEditCtrl to undo the removal of the text.
            # As mentioned above, I hang my head in shame.  Hopefully the next wxPython (this is 2.5.3.1) will fix this.
            if ("__WXMAC__" in wx.PlatformInfo) and (not self.get_read_only()):
                wx.CallAfter(self.undo)

        # Reset the cursor following the Drag/Drop event
        self.parent.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

    def PositionAfter(self):
        self.GotoPos(self.pos)

    def SetSelectionAfter(self):
        self.SetSelection(self.selection[0], self.selection[1])

    def OnLeftUp(self, event):
        """ Left Mouse Button Up event """
        # If we do not already have a cursor position saved, save it
        #if self.cursorPosition == 0:
        # No, just save it anyway.  Otherwise, we occasionally can't make a new selection.
        self.cursorPosition = (self.GetCurrentPos(), self.GetSelection())
        
        # Get the Start and End times from the time codes on either side of the cursor
        (segmentStartTime, segmentEndTime) = self.get_selected_time_range()

        # The code that is now indented was causing an error if you tried to edit the Transcript from the
        # Clip Properties form, as the parent didn't have a ControlObject property.
        # Therefore, let's test that we're coming from the Transcript Window before we do this test.
        if type(self.parent).__name__ == '_TranscriptDialog':

            # If we have a Clip loaded, the StartTime should be the beginning of the Clip, not 0!
            if type(self.parent.ControlObject.currentObj).__name__ == 'Clip' and (segmentStartTime == 0):
                segmentStartTime = self.parent.ControlObject.currentObj.clip_start

            # If we're in read_only mode and the video is not currently playing, position video immediately with left-click.
            # If we are not in read_only mode, we need to delay the selection until right-click.
            if self.get_read_only() and not self.parent.ControlObject.IsPlaying():
                # First, clear the current selection in the visualization window, if there is one.
                self.parent.ControlObject.ClearVisualizationSelection()
                # Set the start and end points to match the current segment
                self.parent.ControlObject.SetVideoSelection(segmentStartTime, segmentEndTime)

        # We need to make sure the cursor is not positioned between a time code symbol and the time code data, which unfortunately
        # can happen.  Preventing this is the sole function of this method.

        # First, see if we have a click Position or a click-drag Selection.
        if (self.GetSelectionStart() == self.GetSelectionEnd()):
            # We have a click Position.  We just need to check the current position.
            curPos = self.GetCurrentPos()
            # Let's see if we are between a Time code and its Data
            if (curPos > 0) and (chr(self.GetCharAt(curPos - 1)) == TIMECODE_CHAR) and (chr(self.GetCharAt(curPos)) == '<'):
                # Let's find the position of the end of the Time Code
                while chr(self.GetCharAt(curPos - 1)) != '>':
                    curPos += 1

                # THE FOLLOWING DON'T WORK!
                # self.GotoPos(curPos) # Cursor right fails once, and backspace doesn't get trapped properly.
                # self.SetAnchor(curPos) and self.SetCurrentPos(curPos)
                # We apparently can't change the Anchor here, so we need to use a wx.CallAfter method to
                # position the cursor later.  I'm not sure WHY, but it works.

                # Remember the current position
                self.pos = curPos
                # use CallAfter to set the cursor to the designated position
                wx.CallAfter(self.PositionAfter)

        else:
            # We have a click-drag Selection.  We need to check the start and the end points.
            selStart = self.GetSelectionStart()
            selEnd = self.GetSelectionEnd()
            # Let's see if any change is made.
            selChanged = False
            # Let's see if the start of the selection falls between a Time Code and its data
            if (selStart > 0) and (chr(self.GetCharAt(selStart - 1)) == TIMECODE_CHAR) and (chr(self.GetCharAt(selStart)) == '<'):
                # If so, we have a change
                selChanged = True
                # Let's find the position of the end of the Time Code
                while chr(self.GetCharAt(selStart - 1)) != '>':
                    selStart += 1
            # Let's see if the end of the selection falls between a Time Code and its data
            if (selEnd > 0) and (chr(self.GetCharAt(selEnd - 1)) == TIMECODE_CHAR) and (chr(self.GetCharAt(selEnd)) == '<'):
                # If so, we have a change
                selChanged = True
                # Let's find the position of the end of the Time Code
                while chr(self.GetCharAt(selEnd - 1)) != '>':
                    selEnd += 1

            # self.SetSelection(selStart, selEnd) does not work if selEnd is between the timecode and the data.  We must use
            # CallAfter to position this correctly.  I'm not sure WHY, but it works.
            # We only need to change the Selection if one of the ends was between a Time Code and its data.
            if selChanged:
                # Remember the new Selection points
                self.selection = (selStart, selEnd)
                # Use CallAfter to set the Selection to the designated positions
                wx.CallAfter(self.SetSelectionAfter)

            else:

                # The transcript has a nasty tendency to display a selection from the point clicked to the
                # following time code marker in read-only mode if the following code is removed.
                if self.get_read_only():
                    wx.CallAfter(self.RestoreCursor)

        # Okay, now let the RichTextEditCtrl have the LeftUp event
        event.Skip()

    def OnRightClick(self, event):
        """ Right-clicking should handle Video Play Control rather than providing the
            traditional right-click editing control menu """
        # If we do not already have a cursor position saved, save it
        if self.cursorPosition == 0:
            self.cursorPosition = (self.GetCurrentPos(), self.GetSelection())

        # Get the Start and End times from the time codes on either side of the cursor
        (segmentStartTime, segmentEndTime) = self.get_selected_time_range()

        # The code that is now indented was causing an error if you tried to edit the Transcript from the
        # Clip Properties form, as the parent didn't have a ControlObject property.
        # Therefore, let's test that we're coming from the Transcript Window before we do this test.
        if type(self.parent).__name__ == '_TranscriptDialog':

            # If we have a Clip loaded, the StartTime should be the beginning of the Clip, not 0!
            if type(self.parent.ControlObject.currentObj).__name__ == 'Clip' and (segmentStartTime == 0):
                segmentStartTime = self.parent.ControlObject.currentObj.clip_start

            # If we're not in read_only mode and the video is not currently playing ...
            if not self.get_read_only() and not self.parent.ControlObject.IsPlaying():
                # First, clear the current selection in the visualization window, if there is one.
                self.parent.ControlObject.ClearVisualizationSelection()
                # Set the start and end points to match the current segment
                self.parent.ControlObject.SetVideoSelection(segmentStartTime, segmentEndTime)

        
        # FOR NON-MAC PLATFORMS
        # Ctrl-Right-Click should play from current position to the end of the video for non-Mac machines!
        if (not '__WXMAC__' in wx.PlatformInfo) and event.ControlDown():
            # Setting End Time to 0 instructs the video player to play to the end of the video!
            segmentEndTime = 0

        # FOR MAC PLATFORM
        # The Mac with a one-button mouse requires the Ctrl-key to emulate a right-click, so we
        # have to add the Meta (Open Apple) key to the mix here.
        # Meta-Right-Click should play from current position to the end of the video on Mac!
        if ('__WXMAC__' in wx.PlatformInfo) and event.MetaDown():
            # Setting End Time to 0 instructs the video player to play to the end of the video!
            segmentEndTime = 0
            
        # If the video is not currently playing ...
        if not self.parent.ControlObject.IsPlaying():
            # First, clear the current selection in the visualization window, if there is one.
            self.parent.ControlObject.ClearVisualizationSelection()
            # Set the start and end points to match the current segment
            self.parent.ControlObject.SetVideoSelection(segmentStartTime, segmentEndTime)
        # Play or Pause the video, depending on its current state
        self.parent.ControlObject.PlayStop()
            
    def RestoreCursor(self):
        """ Restore the Cursor position following right-click play control operation """
        # If we have a stored Cursor Position ...
        if self.cursorPosition != 0:
            # Reset the Cursor Position
            self.SetCurrentPos(self.cursorPosition[0])
            # And reset the Selection, if there was one.
            self.SetSelection(self.cursorPosition[1][0], self.cursorPosition[1][1])
            # Once the cursor position has been reset, we need to clear out the Cursor Position Data
            self.cursorPosition = 0

    def AdjustIndexes(self, adjustmentAmount):
        """ Adjust Transcript Time Codes by the specified amount """
        # Let's try to remember the cursor position
        self.cursorPosition = (self.GetCurrentPos(), self.GetSelection())
        
        # Move the cursor to the beginning of the document
        self.GotoPos(0)

        # Let's remember the self.codes_vis setting, as show_all_hidden() changes it.
        codes_vis = self.codes_vis
        # Let's show all the hidden text of the time codes.  This doesn't work without it!
        self.show_all_hidden()

        # Let's find each time code mark and update it.  This will be easier if we use the
        # POSITION rather than the VALUE of the "timecodes" list, as we need to change that
        # list as we go too!
        for loop in range(0, len(self.timecodes)):
            self.cursor_find('%s' % TIMECODE_CHAR)
            self.select_find('<%s>' % self.timecodes[loop])

            # For some reason, the time code character doesn't get selected at the beginning of the document.
            # Therefore, we need to check for it and exclude it if it is not present, include it if it is.
            if self.GetSelectedText()[0] == TIMECODE_CHAR:
                self.ReplaceSelection('')
                if codes_vis:
                    tempStyle = self.style
                    self.style = self.STYLE_TIMECODE
                    self.InsertStyledText("%s" % TIMECODE_CHAR)
                    self.InsertHiddenText("<%d>" % (self.timecodes[loop] + int(adjustmentAmount* 1000)))
                    self.style = tempStyle
                else:
                    self.InsertHiddenText("%s<%d>" % (TIMECODE_CHAR, self.timecodes[loop] + int(adjustmentAmount* 1000)))
            else:
                self.ReplaceSelection('')
                if codes_vis:
                    tempStyle = self.style
                    self.style = self.STYLE_TIMECODE
                    self.InsertHiddenText("<%d>" % (self.timecodes[loop] + int(adjustmentAmount * 1000)))
                    self.style = tempStyle
                else:
                    self.InsertHiddenText("<%d>" % (self.timecodes[loop] + int(adjustmentAmount* 1000)))
            self.timecodes[loop] = self.timecodes[loop] + int(adjustmentAmount * 1000)
        

        # We better hide all the hidden text for the time codes again
        self.hide_all_hidden()
        # We also need to reset self.codes_vis, which was incorrectly changed by hide_all_hidden()
        self.codes_vis = codes_vis
            
        # Okay, this might not work because of changes we've made to the transcript, but let's
        # try restoring the Cursor Position when all is said and done.
        self.RestoreCursor()
        self.Update()
        
# Events    
    def EVT_DOC_CHANGED(self, win, id, func):
        """Set function to be called when document is modified."""


class TranscriptEditorDropTarget(wx.PyDropTarget):
    
    def __init__(self, editor):
        wx.PyDropTarget.__init__(self)
        self.editor = editor

        # specify the type of data we will accept
        # DataTreeDragData is the format used by the tree control,
        # which is a cPickle.dumps() of DataTreeDragDropData()

        self.df = wx.CustomDataFormat("DataTreeDragData")
        self.data = wx.CustomDataObject(self.df)
        self.SetDataObject(self.data)

    # some virtual methods that track the progress of the drag
    def OnEnter(self, x, y, dragResult):
        return dragResult

    def OnLeave(self):
        pass

    def OnDrop(self, x, y):
        # Drop location isn't important for this target, proceed.
        return True

    def OnDragOver(self, x, y, d):
        #self.log.WriteText("OnDragOver: %d, %d, %d\n" % (x, y, d))

        # The value returned here tells the source what kind of visual
        # feedback to give.  For example, if wxDragCopy is returned then
        # only the copy cursor will be shown, even if the source allows
        # moves.  You can use the passed in (x,y) to determine what kind
        # of feedback to give.  In this case we return the suggested value
        # which is based on whether the Ctrl key is pressed.
        return d

    # Called when OnDrop returns True.  We need to get the data and
    # do something with it.
    def OnData(self, x, y, d):
        # copy the data from the drag source to our data object
        if self.GetData():
            # Extract actual data passed by DataTreeDropSource
            sourceData = cPickle.loads(self.data.GetData())
            # Now you can do sourceData.recNum, sourceData.text,
            # sourceData.nodetype should be 'KeywordNode'
            # Determine where the Transcript was loaded from
            if self.editor.TranscriptObj:
                if self.editor.TranscriptObj.clip_num != 0:
                    targetType = 'Clip'
                    targetRecNum = self.editor.TranscriptObj.clip_num
                    clipObj = Clip.Clip(targetRecNum)
                    targetName = clipObj.id
                else:
                    targetType = 'Episode'
                    targetRecNum = self.editor.TranscriptObj.episode_num
                    epObj = Episode.Episode(targetRecNum)
                    targetName = epObj.id
                DragAndDropObjects.DropKeyword(self.editor, sourceData, \
                    targetType, targetName, targetRecNum, 0)
            else:
                # No transcript Object loaded, do nothing
                pass
        return d  # what is returned signals the source what to do
                  # with the original data (move, copy, etc.)  In this
                  # case we just return the suggested value given to us.


# TODO: Be a Drag/Drop source for text and for creating Clips.
# To create Clips, you need to populate and send a ClipDragDropData object.

class TranscriptDropSource(wx.DropSource):
    """This is a custom DropSource object to drag text from the Transcript
    onto the Data Tree tab."""

    def __init__(self, parent):
        wx.DropSource.__init__(self, parent)
        self.parent = parent

    def SetData(self, obj):
        wx.DropSource.SetData(self, obj)
        self.data = cPickle.loads(obj.GetData())

    def InDatabase(self, windowx, windowy):
       """Determine if the given X/Y position is within the Database Window."""
       (transLeft, transTop, transWidth, transHeight) = self.parent.ControlObject.GetDatabaseDims()
       transRight = transLeft + transWidth
       transBot = transTop + transHeight
       return (windowx >= transLeft and windowx <= transRight and windowy >= transTop and windowy <= transBot)

    def GiveFeedback(self, effect):
        # This method does not provide the x, y coordinates of the mouse within the control, so we
        # have to figure that out the hard way. (Contrast with DropTarget's OnDrop and OnDragOver methods)
        # Get the Mouse Position on the Screen
        (windowx, windowy) = wx.GetMousePosition()

        # Determine if we are over the Database Tree Tab
        if self.InDatabase(windowx, windowy):
            # We need the Database Tree to scroll up or down if we reach the top or bottom of the Tree Control.
            # I KNOW this is poor form, but can't figure out a better way to do it.
            (x, y) = self.parent.ControlObject.DataWindow.DBTab.tree.ScreenToClientXY(windowx, windowy)
            (w, h) = self.parent.ControlObject.DataWindow.DBTab.tree.GetClientSizeTuple()
            # If we are dragging at the top of the window, scroll down
            if y < 8:
                # The wxWindow.ScrollLines() method is only implemented on Windows.  We must use something different on the Mac.
                if "wxMSW" in wx.PlatformInfo:
                   self.parent.ControlObject.DataWindow.DBTab.tree.ScrollLines(-2)
                else:
                   # Suggested by Robin Dunn
                   first = self.parent.ControlObject.DataWindow.DBTab.tree.GetFirstVisibleItem()
                   prev = self.parent.ControlObject.DataWindow.DBTab.tree.GetPrevSibling(first)
                   if prev:
                      # drill down to find last expanded child
                      while self.parent.ControlObject.DataWindow.DBTab.tree.IsExpanded(prev):
                         prev = self.parent.ControlObject.DataWindow.DBTab.tree.GetLastChild(prev)
                   else:
                      # if no previous sub then try the parent
                      prev = self.parent.ControlObject.DataWindow.DBTab.tree.GetItemParent(first)

                   if prev:
                      self.parent.ControlObject.DataWindow.DBTab.tree.ScrollTo(prev)
                   else:
                      self.parent.ControlObject.DataWindow.DBTab.tree.EnsureVisible(first)
            # If we are dragging at the bottom of the window, scroll up
            elif y > h - 8:
                # The wxWindow.ScrollLines() method is only implemented on Windows.  We must use something different on the Mac.
                if "wxMSW" in wx.PlatformInfo:
                   self.parent.ControlObject.DataWindow.DBTab.tree.ScrollLines(2)
                else:
                   # Suggested by Robin Dunn
                   # first find last visible item by starting with the first
                   next = None
                   last = None
                   item = self.parent.ControlObject.DataWindow.DBTab.tree.GetFirstVisibleItem()
                   while item:
                      if not self.parent.ControlObject.DataWindow.DBTab.tree.IsVisible(item): break
                      last = item
                      item = self.parent.ControlObject.DataWindow.DBTab.tree.GetNextVisible(item)

                   # figure out what the next visible item should be,
                   # either the first child, the next sibling, or the
                   # parent's sibling
                   if last:
                       if self.parent.ControlObject.DataWindow.DBTab.tree.IsExpanded(last):
                          next = self.parent.ControlObject.DataWindow.DBTab.tree.GetFirstChild(last)[0]
                       else:
                          next = self.parent.ControlObject.DataWindow.DBTab.tree.GetNextSibling(last)
                          if not next:
                             prnt = self.parent.ControlObject.DataWindow.DBTab.tree.GetItemParent(last)
                             if prnt:
                                next = self.parent.ControlObject.DataWindow.DBTab.tree.GetNextSibling(prnt)

                   if next:
                      self.parent.ControlObject.DataWindow.DBTab.tree.ScrollTo(next)
                   elif last:
                      self.parent.ControlObject.DataWindow.DBTab.tree.EnsureVisible(last)

            if (self.parent.ControlObject.GetDatabaseTreeTabObjectNodeType() == 'CollectionNode') or \
               (self.parent.ControlObject.GetDatabaseTreeTabObjectNodeType() == 'ClipNode'):
                # Make sure the cursor reflects an acceptable drop.  (This resets it if it was previously changed
                # to indicate a bad drop.)
                self.parent.ControlObject.SetDatabaseTreeTabCursor(wx.CURSOR_ARROW)
                # FALSE indicates that feedback is NOT being overridden, and thus that the drop is GOOD!
                return False
            else:
                # Set the cursor to give visual feedback that the drop will fail.
                self.parent.ControlObject.SetDatabaseTreeTabCursor(wx.CURSOR_NO_ENTRY)
                # Setting the Effect to wxDragNone has absolutely no effect on the drop, if I understand this correctly.
                effect = wx.DragNone
                # returning TRUE indicates that the default feedback IS being overridden, thus that the drop is BAD!
                return True
        else:
            # Set the cursor to give visual feedback that the drop will fail.
            # NOTE:  We do NOT want to enable text drag within the Transcript.  This would cause problems with keeping the
            #        timecodes ordered correctly.
            self.parent.SetCursor(wx.StockCursor(wx.CURSOR_NO_ENTRY))
            # Setting the Effect to wxDragNone has absolutely no effect on the drop, if I understand this correctly.
            effect = wx.DragNone
            # returning TRUE indicates that the default feedback IS being overridden, thus that the drop is BAD!
            return True
