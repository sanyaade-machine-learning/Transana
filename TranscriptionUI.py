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

if __name__ == '__main__':
    import wxversion
    wxversion.select('2.6-unicode')
    
import wx

if __name__ == '__main__':
    __builtins__._ = wx.GetTranslation
    wx.SetDefaultPyEncoding('utf_8')

import gettext
import pickle
from TranscriptToolbar import TranscriptToolbar
from TranscriptEditor import TranscriptEditor
import Dialogs
# Import the Transana Font Dialog
import TransanaFontDialog
import TransanaGlobal


class TranscriptionUI(object):
    """This class manages the graphical user interface for the transcription
    editors component.  It creates the transcript window containing a
    TranscriptToolbar and a TranscriptEditor object."""

    def __init__(self, parent):
        """Initialize an TranscriptionUI object."""
        self.dlg = _TranscriptDialog(parent, -1)
        
        self.dlg.toolbar.Enable(0)

        # We need to adjust the screen position on the Mac.  I don't know why.
        if "__WXMAC__" in wx.PlatformInfo:
            pos = self.dlg.GetPosition()
            self.dlg.SetPosition((pos[0]-20, pos[1]))
        
        
# Public methods
    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.dlg.ControlObject=ControlObject

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
        
    def LoadTranscript(self, transcriptObj):
        """Load a transcript object."""
        self.dlg.editor.set_read_only(True)
        self.dlg.toolbar.ClearToolbar()
        
        # let's figure out what format the desired transcript was saved as
        if transcriptObj.text != None:
            temp = transcriptObj.text[2:5]
            
            if DEBUG:
                print "TranscriptionUI.LoadTranscript():  temp == 'rtf'?? ", temp
                
            try:
                # was it RTF?
                if temp != u'rtf':
                    
                    if DEBUG:
                        print "TranscriptionUI.LoadTranscript():  loading with pickle"
                        
                    self.dlg.editor.load_transcript(transcriptObj, 'pickle')
                # or was it pickled?
                else:
                    
                    if DEBUG:
                        print "TranscriptionUI.LoadTranscript():  loading without pickle"
                        
                    self.dlg.editor.load_transcript(transcriptObj) # flies off to transcripteditor.py
            except UnicodeDecodeError:
                if DEBUG:
                    import sys, traceback
                    print sys.exc_info()[0], sys.exc_info()[1]
                    traceback.print_exc()

                # any unicode decoding errors are likely coming from attempting
                # to decode pickled data, so the transcript is most probably
                # pickled.
                self.dlg.editor.load_transcript(transcriptObj, 'pickle')

        self.dlg.toolbar.Enable(1)

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
        # On the Mac, we need to leave the Clipboard open to handle complex objects, which get lost when the Clipboard
        # is closed.  However, this interferes with the internal Paste operation.  So we'll explicitly close the Clipboard here
        if "__WXMAC__" in wx.PlatformInfo:
            wx.TheClipboard.Close()
        self.get_editor().Cut()
        # On the Mac, we need to leave the Clipboard open to handle complex objects, which get lost when the Clipboard
        # is closed.  However, this interferes with the internal Paste operation.  So we'll explicitly re-open the Clipboard here
        if "__WXMAC__" in wx.PlatformInfo:
            wx.TheClipboard.Open()

    def TranscriptCopy(self):
        # On the Mac, we need to leave the Clipboard open to handle complex objects, which get lost when the Clipboard
        # is closed.  However, this interferes with the internal Paste operation.  So we'll explicitly close the Clipboard here
        if "__WXMAC__" in wx.PlatformInfo:
            wx.TheClipboard.Close()
        self.get_editor().Copy()
        # On the Mac, we need to leave the Clipboard open to handle complex objects, which get lost when the Clipboard
        # is closed.  However, this interferes with the internal Paste operation.  So we'll explicitly re-open the Clipboard here
        if "__WXMAC__" in wx.PlatformInfo:
            wx.TheClipboard.Open()

    def TranscriptPaste(self):
        # On the Mac, we need to leave the Clipboard open to handle complex objects, which get lost when the Clipboard
        # is closed.  However, this interferes with the internal Paste operation.  So we'll explicitly close the Clipboard here
        if "__WXMAC__" in wx.PlatformInfo:
            wx.TheClipboard.Close()
        self.get_editor().Paste()
        # On the Mac, we need to leave the Clipboard open to handle complex objects, which get lost when the Clipboard
        # is closed.  However, this interferes with the internal Paste operation.  So we'll explicitly re-open the Clipboard here
        if "__WXMAC__" in wx.PlatformInfo:
            wx.TheClipboard.Open()

    def CallFontDialog(self):
        """ Trigger the TransanaFontDialog, either updating the font settings for the selected text or
            changing the the font settingss for the current cursor position. """
        # Let's try to remember the cursor position
        self.dlg.editor.cursorPosition = (self.dlg.editor.GetCurrentPos(), self.dlg.editor.GetSelection())
        # Get current Font information from the Editor
        editorFont = self.dlg.editor.get_font()

        # If we don't have a text selection, we can just get the wxFontData and go.
        if self.dlg.editor.GetSelection()[0] == self.dlg.editor.GetSelection()[1]:
            # Create and populate a wxFont object
            font = wx.Font(editorFont[1], wx.FONTFAMILY_DEFAULT, wx.NORMAL, wx.NORMAL, faceName=editorFont[0])
            # font.SetFaceName(editorFont[0])
            #font.SetPointSize(editorFont[1])
            if self.dlg.editor.get_bold():
                font.SetWeight(wx.BOLD)
            if self.dlg.editor.get_italic():
                font.SetStyle(wx.ITALIC)
            if self.dlg.editor.get_underline():
                font.SetUnderlined(True)

            # Create and populate a wxFontData object
            fontData = wx.FontData()
            fontData.EnableEffects(True)
            fontData.SetInitialFont(font)

            # There is a bug in wxPython.  wx.ColourRGB() transposes Red and Blue.  This hack fixes it!
            color = wx.ColourRGB(editorFont[2])
            rgbValue = (color.Red() << 16) | (color.Green() << 8) | color.Blue()
            fontData.SetColour(wx.ColourRGB(rgbValue))
            
        # If we DO have a selection, we need to check, for mixed font specs in the selection
        else:
            # Set the Wait cursor (This doesn't appear to show up.)
            self.dlg.editor.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
            
            # First, get the initial values for the Font Dialog.  This will match the
            # formatting of the LAST character in the selection.
            fontData = TransanaFontDialog.TransanaFontDef()
            fontData.fontFace = editorFont[0]
            fontData.fontSize = editorFont[1]
            if self.dlg.editor.get_bold():
                fontData.fontWeight = TransanaFontDialog.tfd_BOLD
            else:
                fontData.fontWeight = TransanaFontDialog.tfd_OFF
            if self.dlg.editor.get_italic():
                fontData.fontStyle = TransanaFontDialog.tfd_ITALIC
            else:
                fontData.fontStyle = TransanaFontDialog.tfd_OFF
            if self.dlg.editor.get_underline():
                fontData.fontUnderline = TransanaFontDialog.tfd_UNDERLINE
            else:
                fontData.fontUnderline = TransanaFontDialog.tfd_OFF
            # There is a bug in wxPython.  wx.ColourRGB() transposes Red and Blue.  This hack fixes it!
            color = wx.ColourRGB(editorFont[2])
            rgbValue = (color.Red() << 16) | (color.Green() << 8) | color.Blue()
            fontData.fontColorDef = wx.ColourRGB(rgbValue)

            # Now we need to iterate through the selection and look for any characters with different font values.
            for selPos in range(self.dlg.editor.GetSelection()[0], self.dlg.editor.GetSelection()[1]):
                # We don't touch the settings for TimeCodes or Hidden TimeCode Data, so these characters can be ignored.                
                if not (self.dlg.editor.GetStyleAt(selPos) in [self.dlg.editor.STYLE_TIMECODE, self.dlg.editor.STYLE_HIDDEN]):
                    # Get the Font Attributes of the current Character
                    attrs = self.dlg.editor.style_attrs[self.dlg.editor.GetStyleAt(selPos)]

                    # Now look for specs that are different, and flag the TransanaFontDef object if one is found.
                    # If the the Symbol Font is used, we ignore this.  (We don't want to change the Font Face of Special Characters.)
                    if (fontData.fontFace != None) and (attrs.font_face != 'Symbol') and (attrs.font_face != fontData.fontFace):
                        del(fontData.fontFace)

                    if (fontData.fontSize != None) and (attrs.font_size != fontData.fontSize):
                        del(fontData.fontSize)

                    if (fontData.fontWeight != TransanaFontDialog.tfd_AMBIGUOUS) and \
                       ((attrs.bold == True) and (fontData.fontWeight == TransanaFontDialog.tfd_OFF)) or \
                       ((attrs.bold == False) and (fontData.fontWeight == TransanaFontDialog.tfd_BOLD)):
                        fontData.fontWeight = TransanaFontDialog.tfd_AMBIGUOUS

                    if (fontData.fontStyle != TransanaFontDialog.tfd_AMBIGUOUS) and \
                       ((attrs.italic == True) and (fontData.fontStyle == TransanaFontDialog.tfd_OFF)) or \
                       ((attrs.italic == False) and (fontData.fontStyle == TransanaFontDialog.tfd_ITALIC)):
                        fontData.fontStyle = TransanaFontDialog.tfd_AMBIGUOUS

                    if (fontData.fontUnderline != TransanaFontDialog.tfd_AMBIGUOUS) and \
                       ((attrs.underline == True) and (fontData.fontUnderline == TransanaFontDialog.tfd_OFF)) or \
                       ((attrs.underline == False) and (fontData.fontUnderline == TransanaFontDialog.tfd_UNDERLINE)):
                        fontData.fontUnderline = TransanaFontDialog.tfd_AMBIGUOUS

                    color = wx.ColourRGB(attrs.font_fg)
                    rgbValue = (color.Red() << 16) | (color.Green() << 8) | color.Blue()
                    if (fontData.fontColorDef != None) and (fontData.fontColorDef != wx.ColourRGB(rgbValue)):
                        del(fontData.fontColorDef)
            # Set the cursor back to normal
            self.dlg.editor.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

        # Create the TransanaFontDialog.
        # Note:  We used to use the wx.FontDialog, but this proved inadequate for a number of reasons.
        #        It offered very few font choices on the Mac, and it couldn't handle font ambiguity.
        fontDialog = TransanaFontDialog.TransanaFontDialog(self.dlg, fontData)
        # Display the FontDialog and get the user feedback
        if fontDialog.ShowModal() == wx.ID_OK:
            # If we don't have a text selection, we can just update the current font settings.
            if self.dlg.editor.GetSelection()[0] == self.dlg.editor.GetSelection()[1]:
                # OLD MODEL -- All characters formatted with all attributes -- no ambiguity allowed.
                # This still applies if there is no selection!
                # Get the wxFontData from the Font Dialog
                newFontData = fontDialog.GetFontData()
                # Extract the Font and Font Color from the FontData
                newFont = newFontData.GetChosenFont()
                newColor = newFontData.GetColour()
                # Set the appropriate Font Attributes.  (Remember, there can be no font ambiguity if there's no selection.)
                if newFont.GetWeight() == wx.BOLD:
                    self.dlg.editor.set_bold(True)
                else:
                    self.dlg.editor.set_bold(False)
                if newFont.GetStyle() == wx.NORMAL:
                    self.dlg.editor.set_italic(False)
                else:
                    self.dlg.editor.set_italic(True)
                if newFont.GetUnderlined():
                    self.dlg.editor.set_underline(True)
                else:
                    self.dlg.editor.set_underline(False)
                # Build a RGB value from newColor.  For some reason the
                # GetRGB() method was returning values in BGR format instead of
                # RGB. -- Nate
                rgbValue = (newColor.Red() << 16) | (newColor.Green() << 8) | newColor.Blue()
                self.dlg.editor.set_font(newFont.GetFaceName(), newFont.GetPointSize(), rgbValue, 0xffffff)

            else:
                # Set the Wait cursor
                self.dlg.editor.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
                # NEW MODEL -- Only update those attributes not flagged as ambiguous.  This is necessary
                # when processing a selection
                # Get the TransanaFontDef data from the Font Dialog.
                newFontData = fontDialog.GetFontDef()
                
                # print
                # print "newFontData =", newFontData

                # Now we need to iterate through the selection and update the font information.
                # It doesn't work to try to apply formatting to the whole block, as ambiguous attributes
                # lose their values.
                for selPos in range(self.dlg.editor.GetSelection()[0], self.dlg.editor.GetSelection()[1]):
                    # We don't want to update the formatting of Time Codes or of hidden Time Code Data.  
                    if not (self.dlg.editor.GetStyleAt(selPos) in [self.dlg.editor.STYLE_TIMECODE, self.dlg.editor.STYLE_HIDDEN]):
                        # Select the character we want to work on from the larger selection
                        self.dlg.editor.SetSelection(selPos, selPos + 1)
                        # Get the previous font attributes for this character
                        attrs = self.dlg.editor.style_attrs[self.dlg.editor.GetStyleAt(selPos)]

                        # Now alter those characteristics that are not ambiguous in the newFontData.
                        # Where the specification is ambiguous, use the old value from attrs.
                        
                        # We don't want to change the font of special symbols!  Therefore, we don't change
                        # the font name for anything in Symbol font.
                        if (newFontData.fontFace != None) and \
                           (attrs.font_face != 'Symbol'):
                            fontFace = newFontData.fontFace
                        else:
                            fontFace = attrs.font_face
                        
                        # print chr(self.dlg.editor.GetCharAt(selPos)), "fontFace = ", fontFace
                        
                        if newFontData.fontSize != None:
                            fontSize = newFontData.fontSize
                        else:
                            fontSize = attrs.font_size

                        if newFontData.fontWeight == TransanaFontDialog.tfd_BOLD:
                            self.dlg.editor.set_bold(True)
                        elif newFontData.fontWeight == TransanaFontDialog.tfd_OFF:
                            self.dlg.editor.set_bold(False)
                        else:
                            # if fontWeight is ambiguous, use the old value
                            if attrs.bold:
                                self.dlg.editor.set_bold(True)
                            else:
                                self.dlg.editor.set_bold(False)

                        if newFontData.fontStyle == TransanaFontDialog.tfd_OFF:
                            self.dlg.editor.set_italic(False)
                        elif newFontData.fontStyle == TransanaFontDialog.tfd_ITALIC:
                            self.dlg.editor.set_italic(True)
                        else:
                            # if fontStyle is ambiguous, use the old value
                            if attrs.italic:
                                self.dlg.editor.set_italic(True)
                            else:
                                self.dlg.editor.set_italic(False)

                        if newFontData.fontUnderline == TransanaFontDialog.tfd_UNDERLINE:
                            self.dlg.editor.set_underline(True)
                        elif newFontData.fontUnderline == TransanaFontDialog.tfd_OFF:
                            self.dlg.editor.set_underline(False)
                        else:
                            # if fontUnderline is ambiguous, use the old value
                            if attrs.underline:
                                self.dlg.editor.set_underline(True)
                            else:
                                self.dlg.editor.set_underline(False)

                        if newFontData.fontColorDef != None:
                            color = newFontData.fontColorDef
                            rgbValue = (color.Red() << 16) | (color.Green() << 8) | color.Blue()
                        else:
                            # There is a bug in wxPython.  wx.ColourRGB() transposes Red and Blue.  This hack fixes it!
                            color = wx.ColourRGB(attrs.font_fg)
                            rgbValue = (color.Blue() << 16) | (color.Green() << 8) | color.Red()
                        # Now apply the font settings for the current character
                        self.dlg.editor.set_font(fontFace, fontSize, rgbValue, 0xffffff)
                # Set the cursor back to normal
                self.dlg.editor.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

        # Destroy the Font Dialog Box, now that we're done with it.
        fontDialog.Destroy()
        # Let's try restoring the Cursor Position when all is said and done.
        self.dlg.editor.RestoreCursor()
        # We've probably taken the focus from the editor.  Let's return it.
        self.dlg.editor.SetFocus()
        
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

    def ChangeLanguages(self):
        """ Change all on-screen prompts to the new language. """
        self.dlg.toolbar.ChangeLanguages()
        

# Private methods    

# Public properties

# dependent classes

class _TranscriptDialog(wx.Dialog):

    def __init__(self, parent, id=-1):
        #print "TranscriptUI Dialog @ (%d, %d) " % (x, y)
        wx.Dialog.__init__(self, parent, id, _("Transcript"),
                            self.__pos(),
                            self.__size(),
                            style=wx.CAPTION | \
                                    wx.RESIZE_BORDER | wx.WANTS_CHARS)

        # Set "Window Variant" to small only for Mac to make fonts match better
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # print "TranscriptWindow:", self.__pos(), self.__size()

        self.ControlObject = None            # The ControlObject handles all inter-object communication, initialized to None

        # add the widgets to the panel
        sizer = wx.BoxSizer(wx.VERTICAL)

        if ALLOW_UNICODE_ENTRY:
            hsizer = wx.BoxSizer(wx.HORIZONTAL)
            self.toolbar = TranscriptToolbar(self)
            hsizer.Add(self.toolbar, 0, wx.ALIGN_TOP, 10)
            self.UnicodeEntry = wx.TextCtrl(self, -1, size=(40, 16), style=wx.TE_PROCESS_ENTER)
            self.UnicodeEntry.SetMaxLength(4)
            hsizer.Add((10,1), 0, wx.ALIGN_CENTER | wx.GROW)
            hsizer.Add(self.UnicodeEntry, 0, wx.ALIGN_RIGHT | wx.TOP | wx.RIGHT, 8)
            self.UnicodeEntry.Bind(wx.EVT_TEXT, self.OnUnicodeText)
            self.UnicodeEntry.Bind(wx.EVT_TEXT_ENTER, self.OnUnicodeEnter)
            sizer.Add(hsizer, 0, wx.ALIGN_TOP, 10)
        else:
            self.toolbar = TranscriptToolbar(self)
            sizer.Add(self.toolbar, 0, wx.ALIGN_TOP, 10)
            
        self.toolbar.Realize()
        
        self.editor = TranscriptEditor(self, id, self.toolbar.OnStyleChange)
        sizer.Add(self.editor, 1, wx.EXPAND, 10)
        if "__WXMAC__" in wx.PlatformInfo:
            # This adds a space at the bottom of the frame on Mac, so that the scroll bar will get the down-scroll arrow.
            sizer.Add((0, 15))
        self.SetSizer(sizer)
        self.SetAutoLayout(True)
        self.Layout()

        # Capture Size Changes
        wx.EVT_SIZE(self, self.OnSize)

        # For debugging purposes so we can load a transcript without
        # database access
        #self.editor.load_transcript("SampleTranscript.rtf")

    def OnCharHook(self, event):
        print "OnCharHook()"

    def OnSize(self, event):
        (left, top) = self.GetPositionTuple()
        (width, height) = self.GetSize()
        self.ControlObject.UpdateWindowPositions('Transcript', width + left, YUpper = top - 4)
        self.Layout()
        # We may need to scroll to keep the current selection in the visible part of the window.
        # Find the start of the selection.
        start = self.editor.GetSelectionStart()
        # Determine the visible line from the starting position's document line, and scroll so that the highlight
        # is 2 lines down, if possible.  (In wxSTC, the position in a document has a Document line, which does not
        # take line wrapping into account, and a visible line, which does.)
        self.editor.ScrollToLine(max(self.editor.VisibleFromDocLine(self.editor.LineFromPosition(start) - 2), 0))

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


class TranscriptionTestApp(wx.App):
    """Test class container for a TranscriptionUI."""

    def OnInit(self):
        gettext.install('Transana', './locale', False)
        # Define supported languages
        self.presLan_en = gettext.translation('Transana', './locale', \
                languages=['en']) # English
        self.presLan_en.install()
        
        self.transcriptWindow = TranscriptionUI(None)
        self.SetTopWindow(self.transcriptWindow.dlg)
        self.transcriptWindow.dlg.editor.load_transcript("SampleTranscript.rtf")
        self.transcriptWindow.dlg.editor.set_read_only()
        self.transcriptWindow.dlg.toolbar.Enable(True)
        self.transcriptWindow.Show()
        self.transcriptWindow.dlg.editor.SaveRTFDocument('test.rtf')
        return True
        
def main():
    """Stand-alone test for Transcription UI.  Does not require database
    connection or other Transana components."""
    
    app = TranscriptionTestApp(0)
    app.MainLoop()
    
if __name__ == '__main__':
    main()
