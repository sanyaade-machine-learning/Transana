# Copyright (C) 2010-2014 The Board of Regents of the University of Wisconsin System 
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

""" This module implements the RichTextEditCtrl class, a Rich Text editor based on wxRichTextCtrl. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "RichTextEditCtrl_RTC DEBUG is ON."

# import wxPython
import wx
# import wxPython's RichTextCtrl
import wx.richtext as richtext

SHOWHIDDEN = False
if SHOWHIDDEN:
    print "RichTextEditCtrl_RTC SHOWHIDDEN is ON."
    HIDDENSIZE = 8
    HIDDENCOLOR = wx.Colour(0, 255, 0)
else:
    HIDDENSIZE = 1
    HIDDENCOLOR = wx.Colour(255, 255, 255)

# import Transana's Constants
import TransanaConstants
# import Transana's Globals
import TransanaGlobal
# import Transana's Dialogs for the ErrorDialog
import Dialogs
# import Transana's Snapshot object
import Snapshot
# import the TranscriptionUI module for parent comparisons
import TranscriptionUI_RTC
# import the TranscriptEditor object
import TranscriptEditor_RTC
# Import the Python Rich Text Parser I wrote
import PyRTFParser

# import Python modules sys, os, string, and re (regular expressions)
import sys, os, string, re
# import Python Exceptions
import exceptions
# the pickle module enables us to fast-save
import pickle                   
# import Python's cStringIO for faster string processing
import cStringIO
# Import Python's traceback module for error reporting in DEBUG mode
import traceback
# import python's webbrowser module
import webbrowser

# Define the Time Code Character
TIMECODE_CHAR = unicode('\xc2\xa4', 'utf-8')


### On Windows, there is a problem with the wxWidgets' wxRichTextCtrl.  It uses up Windows GDI Resources and Transana will crash
### if the GDI Resource Usage exceeds 10,000.  This happens primarily during report generation and bulk formatting.  See wxWidgets
### ticket #13381.
###
### NOTE:  True for wxPython 2.8.12.x, but not for wxPython 2.9.4.0
###
### The following Windows-only function determines the number of GDI Resources in use by Transana.
##
### If we're on Windows ...
##if 'wxMSW' in wx.PlatformInfo:
##
##    # import the win32 methods and constants we need to determine the number of GDI Resources in use.
##    # This is based on code found in wxWidgets ticket #4451
##    from win32api import GetCurrentProcess
##    from win32con import GR_GDIOBJECTS
##    from win32process import GetGuiResources
##
##    def GDIReport():
##        """ Report the number of GDI Resources in use by Transana """
##        # Get the number of GDI Resources in use from the Operating System
##        numGDI = GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS)
##        # Return the number of GDI Resources in use
##        return numGDI
    

class RichTextEditCtrl(richtext.RichTextCtrl):
    """ This class is a Rich Text edit control implemented for Transana """

    def __init__(self, parent, id=-1, pos = None, suppressGDIWarning = False):
        """Initialize a RichTextCtrl object."""
        # Initialize a RichTextCtrl object
        richtext.RichTextCtrl.__init__(self, parent, id, pos = pos, style=wx.VSCROLL | wx.HSCROLL | wx.WANTS_CHARS)

        ## Unfortunately, SetScale is still pretty broken.  It changes the size of the transcript display,
        ## but event.GetPosition() isn't Scale sensitive and GetScaleX() and GetScaleY() don't return the
        ## correct values (on OSX for sure, I didn't check Windows.)  I tried adjusting GetPosition for a
        ## KNOWN scale of 1.25, and the results were just not right.  As a result, Drag-and-Drop in the
        ## Transcript stopped working, so I couldn't make a selection and create a Clip or an inVivo code.

        ## # Get the Transcript Scale from the Configuration Data
        ## scaleFactor = TransanaGlobal.configData.transcriptScale
        ## # ... Set the Scale so text looks bigger
        ## self.SetScale(scaleFactor, scaleFactor)

        # Bind key press handlers
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        self.Bind(wx.EVT_CHAR, self.OnKey)
        self.Bind(wx.EVT_KEY_UP, self.OnKeyUp)
        # Bind mouse click handlers
        self.Bind(wx.EVT_LEFT_DOWN, self.OnLeftDown)
        self.Bind(wx.EVT_LEFT_UP, self.OnLeftUp)
        self.Bind(wx.EVT_MOTION, self.OnMotion)
        # Bind a hyperlink event handler
        self.Bind(wx.EVT_TEXT_URL, self.OnURL)

        # The wx.richtext.RichTextCtrl does some things that aren't Transana-friendly with default behaviors.
        # This section, and the accompanying methods, clean that up by replacing the standard Cut, Copy, Paste,
        # Undo, and Redo methods
        
        # Replace the Accelerator Table for the RichTextCtrl.
        # This removes the Ctrl-A Select All accelerator completely
        # As of wxPython 2.9.5, this is AUTOMATICALLY converted to CMD for Mac
        accelList = [(wx.ACCEL_CTRL,  ord('C'), wx.ID_COPY),
                     (wx.ACCEL_CTRL,  ord('V'), wx.ID_PASTE),
                     (wx.ACCEL_CTRL,  ord('X'), wx.ID_CUT),
                     (wx.ACCEL_CTRL,  ord('Y'), wx.ID_REDO),
                     (wx.ACCEL_CTRL,  ord('Z'), wx.ID_UNDO)]

        aTable = wx.AcceleratorTable(accelList)
        # Assign the modified accelerator table to the control
        self.SetAcceleratorTable(aTable)

        # We need to define our own Cut, Copy, and Paste methods rather than using the default ones from the RichTextCtrl.
        self.Bind(wx.EVT_MENU, self.OnCutCopy, id=wx.ID_CUT)
        self.Bind(wx.EVT_MENU, self.OnCutCopy, id=wx.ID_COPY)
        self.Bind(wx.EVT_MENU, self.OnPaste, id=wx.ID_PASTE)
        # However, we can leave the Undo and Redo commands alone and use the default ones from the RichTextCtrl.

        # Initialize current style to None
        self.txtAttr = None

        # Define the Default Text style
        self.txtOriginalAttr = richtext.RichTextAttr()
        self.txtOriginalAttr.SetTextColour(wx.Colour(0,0,0))
        self.txtOriginalAttr.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.txtOriginalAttr.SetFontFaceName(TransanaGlobal.configData.defaultFontFace)
        self.txtOriginalAttr.SetFontSize(TransanaGlobal.configData.defaultFontSize)
        self.txtOriginalAttr.SetFontWeight(wx.FONTWEIGHT_NORMAL)
        self.txtOriginalAttr.SetFontStyle(wx.FONTSTYLE_NORMAL)
        self.txtOriginalAttr.SetFontUnderlined(False)

        # Define the style for Time Codes
        self.txtTimeCodeAttr = richtext.RichTextAttr()
        self.txtTimeCodeAttr.SetTextColour(wx.Colour(255,0,0))
        self.txtTimeCodeAttr.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.txtTimeCodeAttr.SetFontFaceName('Courier New')
        self.txtTimeCodeAttr.SetFontSize(14)
        self.txtTimeCodeAttr.SetFontWeight(wx.FONTWEIGHT_NORMAL)
        self.txtTimeCodeAttr.SetFontStyle(wx.FONTSTYLE_NORMAL)
        self.txtTimeCodeAttr.SetFontUnderlined(False)

        self.STYLE_TIMECODE = 0

        # Define the style for Hidden Text (used for Time Code Data)
        self.txtHiddenAttr = richtext.RichTextAttr()
        self.txtHiddenAttr.SetTextColour(HIDDENCOLOR)
        self.txtHiddenAttr.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.txtHiddenAttr.SetFontFaceName('Times New Roman')
        self.txtHiddenAttr.SetFontSize(HIDDENSIZE)
        self.txtHiddenAttr.SetFontWeight(wx.FONTWEIGHT_NORMAL)
        self.txtHiddenAttr.SetFontStyle(wx.FONTSTYLE_NORMAL)
        self.txtHiddenAttr.SetFontUnderlined(False)

        self.STYLE_HIDDEN = 0

        # Define the style for Human Readable Time Code data
        self.txtTimeCodeHRFAttr = richtext.RichTextAttr()
        self.txtTimeCodeHRFAttr.SetTextColour(wx.Colour(0,0,255))
        self.txtTimeCodeHRFAttr.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.txtTimeCodeHRFAttr.SetFontFaceName('Courier New')
        self.txtTimeCodeHRFAttr.SetFontSize(10)
        self.txtTimeCodeHRFAttr.SetFontWeight(wx.FONTWEIGHT_NORMAL)
        self.txtTimeCodeHRFAttr.SetFontStyle(wx.FONTSTYLE_NORMAL)
        self.txtTimeCodeHRFAttr.SetFontUnderlined(False)

        # If set to point to a wxProgressDialog, this dialog will be updated as a document is loaded.
        # Initialize to None.
        self.ProgressDlg = None

        self.lineNumberWidth = 0
        # Initialize the String Selection used in keystroke processing
        self.keyStringSelection = ''

        # We occasionally need to adjust styles based on proximity to a time code.  We need a flag to indicate that.
        self.timeCodeFormatAdjustment = False

        # This document should only display the GDI Warning once.  Create a flag.
        if 'wxMSW' in wx.PlatformInfo:
            # The GDIWarningShown flag should be False initially, unless this segment is supposed to suppress the GDI warning
            self.GDIWarningShown = suppressGDIWarning

    def SetReadOnly(self, state):
        """ Change Read Only status, implemented for wxSTC compatibility """
        # RTC uses SetEditable(), the opposite of STC's SetReadOnly()
        self.SetEditable(not state)

    def GetReadOnly(self):
        """ Report Read Only status, implemented for wxSTC compatibility """
        # RTC uses IsEditable(), the opposite of STC's GetReadOnly()
        return not self.IsEditable()

    def GetModify(self):
        """ Report Modified Status, implemented for wxSTC compatibility """
        return self.IsModified()

##    def PutEditedSelectionInClipboard(self):
##        """ Put the TEXT for the current selection into the Clipboard """
##        tempTxt = self.GetStringSelection()
##        # Initialize an empty string for the modified data
##        newSt = ''
##        # Created a TextDataObject to hold the text data from the clipboard
##        tempDataObject = wx.TextDataObject()
##        # Track whether we're skipping characters or not.  Start out NOT skipping
##        skipChars = False
##        # Now let's iterate through the characters in the text
##        for ch in tempTxt:
##            # Detect the time code character
##            if ch == TransanaConstants.TIMECODE_CHAR:
##                # If Time Code, start skipping characters
##                skipChars = True
##            # if we're skipping characters and we hit the ">" end of time code data symbol ...
##            elif (ch == '>') and skipChars:
##                # ... we can stop skipping characters.
##                skipChars = False
##            # If we're not skipping characters ...
##            elif not skipChars:
##                # ... add the character to the new string.
##                newSt += ch
##        # Save the new string in the Text Data Object
##        tempDataObject.SetText(newSt)
##        # Place the Text Data Object in the Clipboard
##        wx.TheClipboard.SetData(tempDataObject)
##
    def GetFormattedSelection(self, format, selectionOnly=False, stripTimeCodes=False):
        """ Return a string with the formatted contents of the RichText control, or just the
            current selection if specified.  The format parameter can either be 'XML' or 'RTF'. """
        # If a valid format is NOT passed ...
        if not format in ['XML', 'RTF']:
            # ... return a blank string
            return ''
        
        # NOTE:  The wx.RichTextCtrl doesn't provide an easy way to get a formatted selection that I can figure out.
        #        This is WAY harder than I expected.  This is a hack, but it works.

        # If we are getting a selection or are stripping time codes ...
        if selectionOnly or stripTimeCodes:
            # Remember the Edited Status of the RTC
            isModified = self.IsModified()
            # Recursively get the current contents of the entire buffer (as it is NOW, potentially modified above)
            originalText = self.GetFormattedSelection('XML')

        # If we're just getting a selection ...
        if selectionOnly:
            # Freeze the control so things will work faster
            self.Freeze()
            # Get the start and end of the current selection
            sel = self.GetSelection()
            # Our last paragraph loses formatting!  Let's fix that!  First, grab the style info
            tmpStyle = self.GetStyleAt(sel[1] - 1)

           # Delete everything AFTER the end of the current selection
            self.Delete((sel[1], self.GetLastPosition()))
            # Delete everything BEFORE the start of the current selection
            self.Delete((0, sel[0]))
            # This leaves us with JUST the selection in the control!

            # Now apply the formatting to the last paragraph
            self.SetSelection(self.GetLastPosition() - 1, self.GetLastPosition())
            self.SetStyleEx(self.GetSelectionRange(), tmpStyle, richtext.RICHTEXT_SETSTYLE_PARAGRAPHS_ONLY)

            # Now reset the selection
            self.SelectAll()

        # If we want to strip time codes ...
        if stripTimeCodes:
            # Recursively get the contents of the entire buffer (as it is NOW, potentially modified above)
            text = self.GetFormattedSelection('XML')
            # Strip the Time Codes from the buffer contents
            text = self.StripTimeCodes(text)

            # Now put the altered contents of the buffer back into the control!
            try:
                stream = cStringIO.StringIO(text)
                # Create an XML Handler
                handler = richtext.RichTextXMLHandler()
                # Load the XML text via the XML Handler.
                # Note that for XML, the BUFFER is passed.
                handler.LoadStream(self.GetBuffer(), stream)
            # exception handling
            except:
                print "XML Handler Load failed"
                print
                print sys.exc_info()[0], sys.exc_info()[1]
                print traceback.print_exc()
                print
                pass

        # Now get the remaining contents of the form, and translate them to the desired format.
        # If XML format is requested ...
        if format == 'XML':
            # Create a Stream
            stream = cStringIO.StringIO()
            # Get an XML Handler
            handler = richtext.RichTextXMLHandler()
            # Save the contents of the control to the stream
            handler.SaveStream(self.GetBuffer(), stream)
            # Convert the stream to a usable string
            tmpBuffer = stream.getvalue()
        # If RTF format is requested ....
        elif format == 'RTF':
            # Get an RTF Handler
            handler = PyRTFParser.PyRichTextRTFHandler()
            # Get the string representation by leaving off the filename parameter
            tmpBuffer = handler.SaveFile(self.GetBuffer())

        # If we are getting a selection or are stripping time codes ...
        if selectionOnly or stripTimeCodes:
            # Now put the ORIGINAL contents of the buffer back into the control!
            try:
                stream = cStringIO.StringIO(originalText)
                # Create an XML Handler
                handler = richtext.RichTextXMLHandler()
                # Load the XML text via the XML Handler.
                # Note that for XML, the BUFFER is passed.
                handler.LoadStream(self.GetBuffer(), stream)
            # exception handling
            except:
                print "XML Handler Load failed"
                print
                print sys.exc_info()[0], sys.exc_info()[1]
                print traceback.print_exc()
                print
                pass

            # If the transcript had been modified ...
            if isModified:
                # ... then mark it as modified
                self.MarkDirty()
            # If the transcript had NOT been modified ...
            else:
                # ... mark the transcript as unchanged.  (Yeah, one of these is probably not necessary.)
                self.DiscardEdits()

        # If we're just getting a selection ...
        if selectionOnly:
            # Restore the original selection
            self.SetSelection(sel[0], sel[1])
            # Now thaw the control so that updates will be displayed again
            self.Thaw()
            
        # Return the buffer's XML string
        return tmpBuffer

    def OnCutCopy(self, event):
        """ Handle Cut and Copy events, over-riding the RichTextCtrl versions.
            This implementation supports Rich Text Formatted text, and at least on Windows and OS X 
            can share formatted text with other programs. """
        # Note the original selection in the text
        origSelection = self.GetSelection()
        # Get the current selection in RTF format
        rtfSelection = self.GetFormattedSelection('RTF', selectionOnly=True, stripTimeCodes=True)

        # Create a Composite Data Object for the Clipboard
        compositeDataObject = wx.DataObjectComposite()
        # If we're on OS X ...
        if 'wxMac' in wx.PlatformInfo:
                # ... then RTF Format is called "public.rtf"
            rtfFormat = wx.CustomDataFormat('public.rtf')
        # If we're on Windows ...
        else:
            # ... then RTF Format is called "Rich Text Format"
            rtfFormat = wx.CustomDataFormat('Rich Text Format')
        # Create a Custom Data Object for the RTF format
        rtfDataObject = wx.CustomDataObject(rtfFormat)
        # Save the RTF version of the control selection to the RTF Custom Data Object
        rtfDataObject.SetData(rtfSelection)
        # Add the RTF Custom Data Object to the Composite Data Object
        compositeDataObject.Add(rtfDataObject)

        # Get the current selection in Plain Text
        txtSelection = self.GetStringSelection()
        # Create a Text Data Object
        txtDataObject = wx.TextDataObject()
        # Save the Plain Text version of the control selection to the Text Data Object
        txtDataObject.SetText(txtSelection)
        # Add the Plain Text Data Object to the Composite Data object
        compositeDataObject.Add(txtDataObject)
        # Open the Clipboard
        wx.TheClipboard.Open()
        # Place the Composite Data Object (with RTF and Plain Text) on the Clipboard
        wx.TheClipboard.SetData(compositeDataObject)
        # Close the Clipboard
        wx.TheClipboard.Close()
        # If we are CUTting (rather than COPYing) ...
        # (NOTE:  On OS X, the event object isn't what we expect, it's a MENU, so we have to get the menu item
        #         text and do a comparison!!!)
        if self.IsEditable() and \
           ((event.GetId() == wx.ID_CUT) or \
            ((sys.platform == 'darwin') and \
             (event.GetEventObject().GetLabel(event.GetId()) == _("Cu&t\tCtrl-X").decode('utf8')))):
            # Reset the selection, which was mangled by the GetFormattedSelection call
            self.SetSelection(origSelection[0], origSelection[1])
            # ... delete the selection from the Rich Text Ctrl.
            self.DeleteSelection()

    def OnPaste(self, event=None):
        """ Handle Paste events, over-riding the RichTextCtrl version.
            This implementation supports Rich Text Formatted text, and at least on Windows can
            share formatted text with other programs. """
        # If the Clipboard isn't Open ...
        if not wx.TheClipboard.IsOpened():
            # ... open it!
            wx.TheClipboard.Open()
        # If the transcript is in EDIT mode ...
        if self.IsEditable():
            # Start a Batch Undo
            self.BeginBatchUndo('Paste')
            # If there's a selection ...
            if self.GetSelection() != (-2, -2):
                # ... delete it.
                self.DeleteSelection()

            # If we're on OS X ...
            if 'wxMac' in wx.PlatformInfo:
                # ... then RTF Format is called "public.rtf"
                rtfFormat = wx.CustomDataFormat('public.rtf')
            # if we're on Windows ...
            else:
                # ... then RTF Format is called "Rich Text Format"
                rtfFormat = wx.CustomDataFormat('Rich Text Format')
            # See if the RTF Format is supported by the current clipboard data object
            if wx.TheClipboard.IsSupported(rtfFormat):
                # Specify that the data object accepts data in RTF format
                customDataObject = wx.CustomDataObject(rtfFormat)
                # Try to get data from the Clipboard
                success = wx.TheClipboard.GetData(customDataObject)
                # If the data in the clipboard is in an appropriate format ...
                if success:
                    # ... get the data from the clipboard
                    formattedText = customDataObject.GetData()

                    if DEBUG:
                        print
                        print "RTF data:"
                        print formattedText
                        print

                    # If the RTF Text ends with a Carriage Return (and it always does!) ...
                    if formattedText[-6:] == '\\par\n}':
                        # ... then remove that carriage return!
                        formattedText = formattedText[:-6] + formattedText[-1]
                        
                    # Prepare the control for data
                    self.Freeze()
                    # Start exception handling
                    try:
                        # Use the custom RTF Handler
                        handler = PyRTFParser.PyRichTextRTFHandler()
                        # Load the RTF data into the Rich Text Ctrl via the RTF Handler.
                        # Note that for RTF, the wxRichTextCtrl CONTROL is passed with the RTF string.
                        handler.LoadString(self, formattedText, insertionPoint=self.GetInsertionPoint(), displayProgress=False)
                    # exception handling
                    except:
                        print "Custom RTF Handler Load failed"
                        print
                        print sys.exc_info()[0], sys.exc_info()[1]
                        print traceback.print_exc()
                        print
                        pass

                    # Signal the end of changing the control
                    self.Thaw()
            # If there's not RTF data, see if there's Plain Text data
            elif wx.TheClipboard.IsSupported(wx.DataFormat(wx.DF_TEXT)):
                # Create a Text Data Object
                textDataObject = wx.TextDataObject()
                # Get the Data from the Clipboard
                wx.TheClipboard.GetData(textDataObject)
                # Write the plain text into the Rich Text Ctrl
                self.WriteText(textDataObject.GetText())
            # End the Batch Undo
            self.EndBatchUndo()
            # Close the Clipboard
            wx.TheClipboard.Close()

##    def SetSavePoint(self):

##        if DEBUG or True:
##            print "RichTextEditCtrl.SetSavePoint() -- Which does NOTHING!!!!"

    def GetSelectionStart(self):
        """ Get the starting point of the current selection, implemented for wxSTC compatibility """
        # If there is NO current selection ...
        if self.GetSelection() == (-2, -2):
            # ... return the current Insertion Point
            return self.GetInsertionPoint()
        # If there IS a selection ...
        else:
            # ... return the Selection's starting point
            return self.GetSelection()[0]

    def GetSelectionEnd(self):
        """ Get the ending point of the current selection, implemented for wxSTC compatibility """
        # If there is NO current selection ...
        if self.GetSelection() == (-2, -2):
            # ... return the current Insertion Point
            return self.GetInsertionPoint()
        # If there IS a selection ...
        else:
            # ... return the Selection's ending point
            return self.GetSelection()[1]

    def ShowCurrentSelection(self):
        """ Ensure that ALL of the current selection is displayed on screen, if possible """
        # Get the current selection
        newSel = self.GetSelection()
        # If there is NO selection ...
        if newSel == (-2, -2):
            # ... then we can exit!
            return
        # Show the END of the selection
        self.ShowPosition(newSel[1])
        # Show the START of the selection
        self.ShowPosition(newSel[0])

        # Don't you wish it were that simple?  If even a little bit of the selection's row is visible, wxPython
        # claims that the selection is visible.  We want to do some additional correction of positioning.

        # Get the position for the First Visible Line on the control
        vlffl = self.GetVisibleLineForCaretPosition(self.GetFirstVisiblePosition())
        # If this value returns None, there's nothing loaded in the control ...
        if vlffl == None:
            # ... so we can exit!
            return
        # Now get the position for the current Caret Position, which should be the start of the selection
        vlfcp = self.GetVisibleLineForCaretPosition(self.GetCaretPosition())
        # Check to see if the selection is ABOVE the TOP of the control, and there's room to scroll down
        if (vlffl.GetAbsolutePosition()[1] >= vlfcp.GetAbsolutePosition()[1]) and (vlffl.GetAbsolutePosition()[1] > 22):
            while (vlffl.GetAbsolutePosition()[1] >= vlfcp.GetAbsolutePosition()[1]) and (vlffl.GetAbsolutePosition()[1] > 22):
                # note the initial position
                before = vlffl.GetAbsolutePosition()[1]

                # Scroll DOWN a bit
                self.ScrollLines(-5)
                # Get the position for the First Visible Line on the control
                vlffl = self.GetVisibleLineForCaretPosition(self.GetFirstVisiblePosition())

                # Note the corrected position
                after = vlffl.GetAbsolutePosition()[1]
                # If the position is unchanged, (this was causing problems because ScrollLines(-5) wasn't always
                # having an affect) ...
                if before == after:
                    # ... then exit the WHILE loop because it's not changing anything
                    break

        # Check to see if the selection is BELOW THE BOTTOM OF THE CONTROL, adjusted a bit to make it work
        elif vlfcp.GetAbsolutePosition()[1] - vlffl.GetAbsolutePosition()[1] > self.GetSize()[1] - 45:
            # If so, scroll UP a bit
            self.ScrollLines(5)

    def GotoPos(self, pos):
        """ Go to the specified position in the transcript, implemented for wxSTC compatibility """
        # Move the Insertion Point to the desired position
        self.SetInsertionPoint(pos)

    def GetCharAt(self, pos):
        """ Get the character at the specified position, implemented for wxSTC compatibility """
        # If there IS a character at the requested position ...
        if len(self.GetRange(pos, pos + 1)) == 1:
            # ... return that character
            return ord(self.GetRange(pos, pos + 1))
        # If there is no valid character at the requested position ...
        else:
            # ... return -1 to signal a failure.
            return -1

    def GetText(self):
        """ Get the text from the control, implemented for wxSTC compatibility """
        # Return the control's contents, obtained from its GetValue() method
        return self.GetValue()
        
    def SetCurrentPos(self, pos):
        """ Set the control's Current position, implemented for wxSTC compatibility """
        # Set the Insertion Point
        self.SetInsertionPoint(pos)
        
    def GetCurrentPos(self):
        """ Get the control's Current position, implemented for wxSTC compatibility """
        # The Insertion Point is the Current Position
        return self.GetInsertionPoint()

    def SetAnchor(self, pos):
        """ Mimic the wxSTC's SetAnchor() method, which makes a selection from the current cursor position to the
            specified point """
        # If a position is passed ...
        if pos >= 0:
            # ... select from the current cursor position to the specified position
            self.SetSelection(self.GetCurrentPos(), pos)
        # If a negative position is passed ...
        else:
            # ... just set the insertion point to 0
            self.SetInsertionPoint(0)

    def GetLength(self):
        """ Report the length of the current document, implemented for wxSTC compatibility """
        # RTC's GetLastPosition() method provides the information
        return self.GetLastPosition()

    def GetTextLength(self):
        """ Report the length of the current document, implemented for wxSTC compatibility """
        # RTC's GetLastPosition() method provides the information
        return self.GetLastPosition()

    def FindText(self, startPos, endPos, text):
        """ Locate the specified text in the RTC """
        # Okay, here's the problem.  RTC doesn't provide a good, FAST text find capacity.  And finding time codes is proving
        # TOO SLOW, especially as transcripts start to get very large.  This attempts to speed that process up.

        # Remember the original selection in the control
        (originalSelStart, originalSelEnd) = self.GetSelection()

        # First, find the text in the STRING representation of the control's data using Python's string.find(),
        # which is very fast.
        textPos = self.GetValue().find(text, startPos, endPos)

        # textPos represents the position of the desired text in the STRING representation.
        # If there are IMAGES in the RTC control, this isn't the correct position, but it's pretty close.
        # Each image only alters the position by 1 place!  So here, we adjust slightly if there are images.

        # If the text was found in the STRING representation ...
        if (textPos > -1):
            # ... search, starting at that location, to the end of the area to be searched
            for pos in range(textPos, endPos - len(text)):
                # Set the Selection based on the for loop variable and the length of the text
                self.SetSelection(pos, pos + len(text))
                # If this is our desired text ...
                if self.GetStringSelection() == text:
                    # ... then this is the TRUE position of the search text.
                    textPos = pos
                    # We can stop looping
                    break

        # If there was a selection highlight before we started mucking about with the control selection ...
        if originalSelStart != originalSelEnd:
            # ... then reset the control selection
            self.SetSelection(originalSelStart, originalSelEnd)
        # If we only had a cursor position, not a selection ...
        else:
            # ... reset the cursor position
            self.SetInsertionPoint(originalSelStart)
        # Return the position of the search text
        return textPos

    def StripTimeCodes(self, XMLText):
        """ This method will take the contents of an RTC buffer in XML format and remove the Time Codes and
            Time Code Data """

        # This deletes based on FORMAT, deleting everything in TIME CODE FORMAT and HIDDEN FORMAT.

        # First, let's create the XML represenation of TIME CODE FORMATTING
        st = '<text textcolor="#FF0000" bgcolor="#FFFFFF" fontsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">'

        # While there is TIME CODE FORMATTING in the text ...
        while st in XMLText:
            # ... identify the starting position of the TIME CODE FORMATTING
            startPos = XMLText.find(st)
            # ... identify the ending position of the TIME CODE FORMATTING
            endPos = XMLText.find('</text>', startPos)
            # Remove the time code with all formatting
            XMLText = XMLText[ : startPos] + XMLText[endPos + 7 : ]

        # First, let's create the XML represenation of TIME CODE FORMATTING FROM RTF, which has different formatting.
        # I've found two variations so far.
        stList = ['<text textcolor="#FF0000" bgcolor="#FFFFFF" fontsize="11" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Arial">',
                  '<text textcolor="#FF0000" bgcolor="#FFFFFF" fontsize="11" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">']

        # For each XML variation ...
        for st in stList:
            # While there is TIME CODE FORMATTING in the text ...
            while st in XMLText:

                # ... identify the starting position of the TIME CODE FORMATTING
                startPos = XMLText.find(st)
                # ... identify the ending position of the TIME CODE FORMATTING
                endPos = XMLText.find('</text>', startPos)
                # Remove the time code with all formatting
                XMLText = XMLText[ : startPos] + XMLText[endPos + 7 : ]

                # ... identify the starting position of the TIME CODE FORMATTING
                startPos = XMLText.find('&lt;', startPos)
                # ... identify the ending position of the TIME CODE FORMATTING
                endPos = XMLText.find('&gt;', startPos)

                # If time code data was found ...
                if (startPos > -1) and (endPos > -1):
                    # ... remove the time code with all formatting
                    XMLText = XMLText[ : startPos] + XMLText[endPos + 4 : ]
            

        # First, let's create the XML represenation of HIDDEN FORMATTING
        st = '<text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">'

        # While there is HIDDEN FORMATTING in the text ...
        while st in XMLText:
            # ... identify the starting position of the HIDDEN FORMATTING
            startPos = XMLText.find(st)
            # ... identify the ending position of the HIDDEN FORMATTING
            endPos = XMLText.find('</text>', startPos)
            # Remove the hidden time code data with all formatting
            XMLText = XMLText[ : startPos] + XMLText[endPos + 7 : ]

        # Some RTF transcripts won't have the formatting right on the time codes, so they won't be found by the code above.
        # This will try to find and remove these additional time codes.

        # First, let's create the XML represenation of A TIME CODE
        st = '&#164;'

        # While there are TIME CODES in the text ...
        while st in XMLText:
            # ... identify the starting position of the TIME CODE
            startPos = XMLText.find(st)
            # ... identify the ending position of the TIME CODE DATA
            endPos = XMLText.find('&gt;', startPos + 4)
            # Remove the hidden time code data
            XMLText = XMLText[ : startPos] + XMLText[endPos + 4 : ]

        return XMLText

    def SetDefaultStyle(self, tmpStyle):
        """ Over-ride the RichTextEditCtrl's SetDefaultStyle() to fix problems with setting the style at the
            end of a line / paragraph.  The font change doesn't always stick. """
        # Attempt to apply the specified font to the document
        richtext.RichTextCtrl.SetDefaultStyle(self, tmpStyle)
        # Get the current default style
        currentStyle = self.GetDefaultStyle()
        # If the font isn't OK, the formatting didn't take!
        if not currentStyle.GetFont().IsOk():
            # In this case, let's add a SPACE character
            self.WriteText(' ')
            # Now select that space character, so it will be over-written by the next typing the user does
            self.SetSelection(self.GetInsertionPoint() - 1, self.GetInsertionPoint())
            # Now FORMAT that space character.
            richtext.RichTextCtrl.SetDefaultStyle(self, tmpStyle)

    def SetTxtStyle(self, fontColor = None, fontBgColor = None, fontFace = None, fontSize = None,
                          fontBold = None, fontItalic = None, fontUnderline = None,
                          parAlign = None, parLeftIndent = None, parRightIndent = None,
                          parTabs = None, parLineSpacing = None, parSpacingBefore = None, parSpacingAfter = None,
                          overrideGDIWarning = False):
        """ I find some of the RichTextCtrl method names to be misleading.  Some character styles are stacked in the RichTextCtrl,
            and they are removed in the reverse order from how they are added, regardless of the method called.

            For example, starting with plain text, BeginBold() makes it bold, and BeginItalic() makes it bold-italic. EndBold()
            should make it italic but instead makes it bold. EndItalic() takes us back to plain text by removing the bold.

            According to Julian, this functions "as expected" because of the way the RichTextCtrl is written.

            The SetTxtStyle() method handles overlapping styles in a way that avoids this problem.  """

        # We need to determine if font conditions are being altered.
        # Remember, (setting is None) is different than (setting == False) for Bold, Italic, and Underline!
        if fontColor or fontBgColor or fontFace or fontSize or \
           (fontBold is not None) or (fontItalic is not None) or (fontUnderline is not None):
            formattingFont = True
        else:
            formattingFont = False
        # If we are formatting ANY paragraph characteristic ...
        if parAlign or parLeftIndent or parRightIndent or parTabs or parLineSpacing or parSpacingBefore or parSpacingAfter:
            # ... note that we are doing paragraph formatting
            formattingParagraph = True
        # If we're NOT doing paragraph formatting ...
        else:
            # ... note that we're NOT doing paragraph formatting
            formattingParagraph = False

##        # If we're on Windows, have reached the GDI threshhold, and are not over-riding the GDI limit ...
##        if ('wxMSW' in wx.PlatformInfo) and (GDIReport() > 8800) and (not overrideGDIWarning):
##            # If we haven't shown the GDI Warning yet ...
##            if not self.GDIWarningShown:
##                # ... create and display the GDI warning
##                prompt = _("This report has reached the upper limit for Windows GDI Resources.") + '\n' + \
##                         _("This report can display formatting in only part of the document.")
##                dlg = Dialogs.InfoDialog(self, prompt)
##                dlg.ShowModal()
##                dlg.Destroy()
##                # Signal that we HAVE shown the GDI Warning
##                self.GDIWarningShown = True
##
##            # Update the current text to use the Original (default) text for the rest of the document
##            self.txtAttr = self.txtOriginalAttr
##            # Set the Basic Style.
##            self.SetBasicStyle(self.txtOriginalAttr)
##            
##        # If we don't have to worry about Windows GDI issues ...
##        else:

        # If there's no SELECTION in the text, paragraph formatting changes get lost.  They don't happen.
        # This code corrects for that.

        # We do Paragraph Changes before font changes.  In theory, this can help prevent the situation where
        # changing both font and paragraph characteristics causes the font changes to be applied to the
        # whole paragraph instead of just the insertion point or selection intended.
        
        # If we have paragraph formatting ...
        if formattingParagraph:

            # If Paragraph Alignment is specified, set the alignment
            if parAlign != None:
                self.txtAttr.SetAlignment(parAlign)
            # If Left Indent is specified, set the left indent
            if parLeftIndent != None:
                # Left Indent can be an integer for left margin only, or a 2-element tuple for left indent and left subindent.
                if type(parLeftIndent) == int:
                    self.txtAttr.SetLeftIndent(parLeftIndent)
                elif (type(parLeftIndent) == tuple) and (len(parLeftIndent) > 1):
                    self.txtAttr.SetLeftIndent(parLeftIndent[0], parLeftIndent[1])
            # If Right Indent is specified, set the right indent
            if parRightIndent != None:
                self.txtAttr.SetRightIndent(parRightIndent)
            # If Tabs are specified, set the tabs
            if parTabs != None:
                self.txtAttr.SetTabs(parTabs)
            # If Line Spacing is specified, set Line Spacing
            if parLineSpacing != None:
                self.txtAttr.SetLineSpacing(parLineSpacing)
            # If Paragraph Spacing Before is set, set spacing before
            if parSpacingBefore != None:
                self.txtAttr.SetParagraphSpacingBefore(parSpacingBefore)
            # If Paragraph Spacing After is set, set spacing after
            if parSpacingAfter != None:
                self.txtAttr.SetParagraphSpacingAfter(parSpacingAfter)

            # Get the current range
            tmpRange = self.GetSelectionRange()
            # Apply the paragraph formatting change
            self.SetStyle(tmpRange, self.txtAttr)

            # Apply the modified paragraph formatting to the document
            self.SetDefaultStyle(self.txtAttr)
            # Set the Basic Style too.
            self.SetBasicStyle(self.txtAttr)
            
        if formattingFont:

            # Create a Font object
            tmpFont = self.txtAttr.GetFont()

            # If the font face (font name) is specified, set the font face
            if fontFace:
                # If we have a valid font and need to change the name ...
                if not tmpFont.IsOk() or (tmpFont.GetFaceName() != fontFace):
                    # ... then change the name
                    self.txtAttr.SetFontFaceName(fontFace)

            # If the font size is specified, set the font size
            if fontSize:
                # If we have a valid font and need to change the size ...
                if not tmpFont.IsOk() or (tmpFont.GetPointSize() != fontSize):
                    # ... then change the size
                    self.txtAttr.SetFontSize(fontSize)

            # If a color is specified, set text color
            if fontColor:
                # If we have a valid font and need to change the color ...
                if self.txtAttr.GetTextColour() != fontColor:
                    # ... then change the color
                    self.txtAttr.SetTextColour(fontColor)

            # If a background color is specified, set the background color
            if fontBgColor:
                # If we have a valid font and need to change the background color ...
                if self.txtAttr.GetBackgroundColour() != fontBgColor:
                    # ... then change the background color
                    self.txtAttr.SetBackgroundColour(fontBgColor)

            # If bold is specified, set or remove bold as requested
            if fontBold != None:
                # If Bold is being turned on ...
                if fontBold:
                    # If we have a valid font and need to change to bold ...
                    if not tmpFont.IsOk() or (tmpFont.GetWeight() != wx.FONTWEIGHT_BOLD):
                        # ... turn on bold
                        self.txtAttr.SetFontWeight(wx.FONTWEIGHT_BOLD)
                # if Bold is being turned off ...
                else:
                    # If we have a valid font and need to change to not-bold ...
                    if not tmpFont.IsOk() or (tmpFont.GetWeight() != wx.FONTWEIGHT_NORMAL):
                        # ... turn off bold
                        self.txtAttr.SetFontWeight(wx.FONTWEIGHT_NORMAL)
            # If italics is specified, set or remove italics as requested
            if fontItalic != None:
                # If Italics are being turned on ...
                if fontItalic:
                    # If we have a valid font and need to change to italics ...
                    if not tmpFont.IsOk() or (tmpFont.GetStyle() != wx.FONTSTYLE_ITALIC):
                        # ... turn on italics
                        self.txtAttr.SetFontStyle(wx.FONTSTYLE_ITALIC)
                # if Italics are being turned off ...
                else:
                    # If we have a valid font and need to change to not-italics ...
                    if not tmpFont.IsOk() or (tmpFont.GetStyle() != wx.FONTSTYLE_NORMAL):
                        # ... turn off italics
                        self.txtAttr.SetFontStyle(wx.FONTSTYLE_NORMAL)
            # If underline is specified, set or remove underline as requested
            if fontUnderline != None:
                # If Underline is being turned on ...
                if fontUnderline:
                    # If we have a valid font and need to change to underline ...
                    if not tmpFont.IsOk() or (not tmpFont.GetUnderlined()):
                        # ... turn on underline
                        self.txtAttr.SetFontUnderlined(True)
                # if Underline is being turned off ...
                else:
                    # If we have a valid font and need to change to not-underline ...
                    if not tmpFont.IsOk() or (tmpFont.GetUnderlined()):
                        # ... turn off underline
                        self.txtAttr.SetFontUnderlined(False)

            # Apply the modified font to the document
            self.SetDefaultStyle(self.txtAttr)

        # If we're in Read Only Mode ...
        if self.GetReadOnly():
            # ... then we DON'T want the control marked as changed
            self.DiscardEdits()

    def SetBold(self, setting):
        """ Change the BOLD attribute """
        # This doesn't properly ignore Time Codes and Hidden Text!!
        # Apply bold to the current selection, or toggle bold if no selection
        # self.ApplyBoldToSelection()

        # If we have a text selection, not an insertion point, ...
        if self.GetSelection() != (-2, -2):

            # Begin an Undo Batch for formatting
            self.BeginBatchUndo('Format')
            
##            # For each character in the current selection ...
##            for selPos in range(self.GetSelection()[0], self.GetSelection()[1]):
##                # Get the Font Attributes of the current Character
##                tmpStyle = self.GetStyleAt(selPos)
##                # We don't want to update the formatting of Time Codes or of hidden Time Code Data.  
##                if (not self.CompareFormatting(tmpStyle, self.txtTimeCodeAttr, fullCompare=False)) and (not self.IsStyleHiddenAt(selPos)):
##                    if setting:
##                        tmpStyle.SetFontWeight(wx.FONTWEIGHT_BOLD)
##                    else:
##                        tmpStyle.SetFontWeight(wx.FONTWEIGHT_NORMAL)
##                    # ... format the selected text
##                    tmpRange = richtext.RichTextRange(selPos, selPos + 1)
##                    self.SetStyle(tmpRange, tmpStyle)
##
##                    print "Applying Bold to ", selPos, selPos + 1

            # Create a blank RichTextAttributes object
            tmpAttr = richtext.RichTextAttr()
            # If we're adding Bold ...
            if setting:
                # ... set the Font Weight to Bold
                tmpAttr.SetFontWeight(wx.FONTWEIGHT_BOLD)
            # If we're removing Bold ...
            else:
                # ... set the Font Weight to Normal
                tmpAttr.SetFontWeight(wx.FONTWEIGHT_NORMAL)
            # Apply the style to the selected block
            self.SetStyle(self.GetSelection(), tmpAttr)

            # End the Formatting Undo Batch
            self.EndBatchUndo()
        else:
            # Set the Style
            self.SetTxtStyle(fontBold = setting)

    def GetBold(self):
        """ Determine the current value of the BOLD attribute """
        # If there is a selection ...
        if self.GetSelection() != (-2, -2):
            # ... get the style of the selection
            tmpStyle = self.GetStyleAt(self.GetSelection()[0])
        # If there is NO selection ...
        else:
            # Determine the insertion point
            ip = self.GetInsertionPoint()
            # If we're not at the first character in the document ...
            if ip > 0:
                # ... get the style BEFORE the current insertion point, which is what the next character typed will use
                tmpStyle = self.GetStyleAt(ip - 1)
            # If we ARE at the first character ...
            else:
                # ... get that character's formatting.  It'll have to do.
                tmpStyle = self.GetStyleAt(ip)
        # If the current style has a valid font ...
        if tmpStyle.GetFont().IsOk():
            # ... return True if bold, False if not
            return tmpStyle.GetFont().GetWeight() == wx.FONTWEIGHT_BOLD
        # If the current style lacks a valid font ...
        else:
            # ... we certainly don't have bold.
            return False
      
    def SetItalic(self, setting):
        """ Change the ITALIC attribute """

# This doesn't properly ignore Time Codes and Hidden Text!!
        # Apply italics to the current selection, or toggle italics if no selection
#        self.ApplyItalicToSelection()

        # If we have a text selection, not an insertion point, ...
        if self.GetSelection() != (-2, -2):
            # Begin an Undo Batch for formatting
            self.BeginBatchUndo('Format')
##            # For each character in the current selection ...
##            for selPos in range(self.GetSelection()[0], self.GetSelection()[1]):
##                # Get the Font Attributes of the current Character
##                tmpStyle = self.GetStyleAt(selPos)
##                # We don't want to update the formatting of Time Codes or of hidden Time Code Data.  
##                if (not self.CompareFormatting(tmpStyle, self.txtTimeCodeAttr, fullCompare=False)) and (not self.IsStyleHiddenAt(selPos)):
##                    if setting:
##                        tmpStyle.SetFontStyle(wx.FONTSTYLE_ITALIC)
##                    else:
##                        tmpStyle.SetFontStyle(wx.FONTSTYLE_NORMAL)
##                    # ... format the selected text
##                    tmpRange = richtext.RichTextRange(selPos, selPos + 1)
##                    self.SetStyle(tmpRange, tmpStyle)

            # Create a blank RichTextAttributes object
            tmpAttr = richtext.RichTextAttr()
            # If we're adding Italic ...
            if setting:
                # ... set the Font Style to Italic
                tmpAttr.SetFontStyle(wx.FONTSTYLE_ITALIC)
            # If we're removing Italic ...
            else:
                # ... set the Font style to Normal
                tmpAttr.SetFontStyle(wx.FONTSTYLE_NORMAL)
            # Apply the style to the selected block
            self.SetStyle(self.GetSelection(), tmpAttr)

            # End the Formatting Undo Batch
            self.EndBatchUndo()
        else:
            # Set the Style
            self.SetTxtStyle(fontItalic = setting)

        
    def GetItalic(self):
        """ Determine the current value of the ITALICS attribute """
        # If there is a selection ...
        if self.GetSelection() != (-2, -2):
            # ... get the style of the selection
            tmpStyle = self.GetStyleAt(self.GetSelection()[0])
        # If there is NO selection ...
        else:
            # Determine the insertion point
            ip = self.GetInsertionPoint()
            # If we're not at the first character in the document ...
            if ip > 0:
                # ... get the style BEFORE the current insertion point, which is what the next character typed will use
                tmpStyle = self.GetStyleAt(ip - 1)
            # If we ARE at the first character ...
            else:
                # ... get that character's formatting.  It'll have to do.
                tmpStyle = self.GetStyleAt(ip)
        # If the current style has a valid font ...
        if tmpStyle.GetFont().IsOk():
            # ... return True if italics, False if not
            return tmpStyle.GetFont().GetStyle() == wx.FONTSTYLE_ITALIC
        # If the current style lacks a valid font ...
        else:
            # ... we certainly don't have italics.
            return False
      
    def SetUnderline(self, setting):
        """ Change the UNDERLINE attribute """

# This doesn't properly ignore Time Codes and Hidden Text!!
        # Apply underline to the current selection, or toggle underline if no selection
#        self.ApplyUnderlineToSelection()

        # If we have a text selection, not an insertion point, ...
        if self.GetSelection() != (-2, -2):

            # There's a bug in wxWidgets on Windows and OS X such that under some circumstances, when you try to apply underlining
            # to part of a line, the whole line gets underlined.

##            # If we're setting a selection to underlined on Windows or OS X ...
##            if setting:
##                # ... assume the surrounding text is NOT underlined
##                surroundingUnderline = False
##                # Find the current selection
##                selection = self.GetSelection()
##                # If there's a character to the left of the selection ...
##                if selection[0] - 1 > 0:
##                    # ... get the style of that character to the left
##                    tmpStyle = self.GetStyleAt(selection[0] - 1)
##                    # If that character IS underlined ...
##                    if tmpStyle.GetFont().GetUnderlined():
##                        # ... signal that the surroundings ARE underlined BEFORE we apply formatting
##                        surroundingUnderline = True
##                # If there's a character to the right of the selection ...
##                if selection[1] + 1 < self.GetLastPosition():
##                    # ... get the style of that character to the left
##                    tmpStyle = self.GetStyleAt(selection[1] + 1)
##                    # If that character IS underlined ...
##                    if tmpStyle.GetFont().GetUnderlined():
##                        # ... signal that the surroundings ARE underlined BEFORE we apply formatting
##                        surroundingUnderline = True
                
            # Begin an Undo Batch for formatting
            self.BeginBatchUndo('Format')
##            # For each character in the current selection ...
##            for selPos in range(self.GetSelection()[0], self.GetSelection()[1]):
##                # Get the Font Attributes of the current Character
##                tmpStyle = self.GetStyleAt(selPos)
##                # We don't want to update the formatting of Time Codes or of hidden Time Code Data.  
##                if (not self.CompareFormatting(tmpStyle, self.txtTimeCodeAttr, fullCompare=False)) and (not self.IsStyleHiddenAt(selPos)):
##                    tmpStyle.SetFontUnderlined(setting)
##                    # ... format the selected text
##                    tmpRange = richtext.RichTextRange(selPos, selPos + 1)
##                    self.SetStyle(tmpRange, tmpStyle)

            # Create a blank RichTextAttributes object
            tmpAttr = richtext.RichTextAttr()
            # If we're adding Underline ...
            if setting:
                # ... set the Font Underline on
                tmpAttr.SetFontUnderlined(True)
            # If we're removing Underline ...
            else:
                # ... set the Font Underline off
                tmpAttr.SetFontUnderlined(False)
            # Apply the style to the selected block
            self.SetStyle(self.GetSelection(), tmpAttr)

            # End the Formatting Undo Batch
            self.EndBatchUndo()

            # There's a bug in wxWidgets on Windows such that underlining doesn't always show up correctly.
            # The following code changes the window size slightly so that a re-draw is forced, thus
            # causing the underlining to show up.  Update() and Refresh() don't work.

            # if we're turning Underline ON ...
            if setting:
                # If we're on Windows, resize the parent window to force the redraw ...
                if ('wxMSW' in wx.PlatformInfo):

                    # ... find the size of the parent window
                    size = self.parent.GetSizeTuple()
                    # Move the Insertion Point to the end of the selection (or this won't work!)
                    self.SetInsertionPoint(self.GetSelection()[1])
                    # Shrink the parent window slightly
                    self.parent.SetSize((size[0], size[1] - 5))
                    # Set the Parent Window back to the original size
                    self.parent.SetSize(size)

##                # If there was NOT surrounding underlining BEFORE ...
##                if not surroundingUnderline:
##                    # ... note that we have not made any adjustment yet
##                    done = False
##                    # If there's a character to the left of the selection ...
##                    if selection[0] - 1 > 0:
##                        # ... get the style of that character to the left
##                        tmpStyle = self.GetStyleAt(selection[0] - 1)
##                        # If that character IS underlined, WE HAVE THE PROBLEM of the whole line having been underlined.
##                        # (This can happen on Windows or OS X!)
##                        if tmpStyle.GetFont().GetUnderlined():
##                            # Select the character before the original selection
##                            self.SetSelection(selection[0] - 1, selection[0])
##                            # Turn underlining OFF for that character.  That usually removes ALL the incorrect underlining!
##                            self.SetUnderline(False)
##                            # Signal that we've already made the underlining adjustment
##                            done = True
##
##                    # If we haven't already made the adjustment and there's a character to the right of the selection ...
##                    if not done and (selection[1] + 1 < self.GetLastPosition()):
##                        # ... get the style of that character to the left
##                        tmpStyle = self.GetStyleAt(selection[1] + 1)
##                        # If that character IS underlined ...
##                        if tmpStyle.GetFont().GetUnderlined():
##                            # Select the character after the original selection
##                            self.SetSelection(selection[1], selection[1] + 1)
##                            # Turn underlining OFF for that character.  That usually removes ALL the incorrect underlining!
##                            self.SetUnderline(False)
##
##                            print "RichTextEditCtrl_RTC.SetUnderline():", selection[1]
##                            
##                    # This will leave our cursor in the wrong place.  Reset it.
##                    self.SetCurrentPos(selection[1])
##
####                # In wxPython 2.9.5.0, this seems to hide the underlining!!  We'll need to let the selection disappear!
##                            
####                # After a moment (or it won't work), reset the cursor selection
####                wx.CallLater(10, self.SetSelection, selection[0], selection[1])
                
        else:
            # Set the Style
            self.SetTxtStyle(fontUnderline = setting)


    def GetUnderline(self):
        """ Determine the current value of the UNDERLINED attribute """
        # If there is a selection ...
        if self.GetSelection() != (-2, -2):
            # ... get the style of the selection
            tmpStyle = self.GetStyleAt(self.GetSelection()[0])
        # If there is NO selection ...
        else:
            # Determine the insertion point
            ip = self.GetInsertionPoint()
            # If we're not at the first character in the document ...
            if ip > 0:
                # ... get the style BEFORE the current insertion point, which is what the next character typed will use
                tmpStyle = self.GetStyleAt(ip - 1)
            # If we ARE at the first character ...
            else:
                # ... get that character's formatting.  It'll have to do.
                tmpStyle = self.GetStyleAt(ip)
        # If the current style has a valid font ...
        if tmpStyle.GetFont().IsOk():
            # ... return True if uderlined, False if not
            return tmpStyle.GetFont().GetUnderlined()
        # If the current style lacks a valid font ...
        else:
            # ... we certainly don't have underline.
            return False
      
    def CompareFormatting(self, fmt1, fmt2, fullCompare=True):
        """ Compare two styles to see if they match.  If fullCompare is True, paragraph formatting and
            tabs are included in the comparison """
        # if either format is None ...
        if (fmt1 is None) or (fmt2 is None):
            # ... we don't have a match!
            return False

        # Start Exception Handline
        try:
            # Get the font specifications from the two formats
            font1 = fmt1.GetFont()
            font2 = fmt2.GetFont()
        # If an exception is raised ...
        except:
            # ... we don't have a match!
            return False

        # Perform the comparison
        if (not font1.IsOk()) or \
           (not font2.IsOk()) or \
           (font1.GetFaceName() != font2.GetFaceName()) or \
           (font1.GetPointSize() != font2.GetPointSize()) or \
           (font1.GetWeight() != font2.GetWeight()) or \
           (font1.GetStyle() != font2.GetStyle()) or \
           (font1.GetUnderlined() != font2.GetUnderlined()) or \
           (fmt1.GetTextColour() != fmt2.GetTextColour()) or \
           (fmt1.GetBackgroundColour() != fmt2.GetBackgroundColour()) or \
           (fullCompare and fmt1.GetAlignment() != fmt2.GetAlignment()) or \
           (fullCompare and fmt1.GetLeftIndent() != fmt2.GetLeftIndent()) or \
           (fullCompare and fmt1.GetLeftSubIndent() != fmt2.GetLeftSubIndent()) or \
           (fullCompare and fmt1.GetRightIndent() != fmt2.GetRightIndent()) or \
           (fullCompare and fmt1.GetLineSpacing() != fmt2.GetLineSpacing()) or \
           (fullCompare and fmt1.GetParagraphSpacingBefore() != fmt2.GetParagraphSpacingBefore()) or \
           (fullCompare and fmt1.GetParagraphSpacingAfter() != fmt2.GetParagraphSpacingAfter()) or \
           (fullCompare and fmt1.GetTabs() != fmt2.GetTabs()):
            # If not the same, return False
            return False
        # If the same, return True
        else:
            return True
        
    def InsertTimeCode(self, timecode):
        """ Insert a time code """
        # If not in Edit Mode ...
        if not self.IsEditable():
            # ... don't edit this!
            return
        # Get the current Style
        tmpStyle = self.GetDefaultStyle()
        # Check to see if tmpStyle is GOOD.
        if not tmpStyle.GetFont().IsOk():
            # If not, use the current font
            tmpStyle = self.txtAttr
        # ... batch the undo
        self.BeginBatchUndo('InsertTimeCode')
        # Set the Style to TimeCode Style
        self.SetDefaultStyle(self.txtTimeCodeAttr)
        # Insert the Time Code character
        self.WriteText(TIMECODE_CHAR)
        # Set the Style to Hidden Time Code Data Style
        self.SetDefaultStyle(self.txtHiddenAttr)
        # Insert the hidden time code data
        self.WriteText('<%d> ' % timecode)
        # Return the Style to whatever it was before
        self.SetDefaultStyle(tmpStyle)
        # End the Undo batch
        self.EndBatchUndo()
        # Let's ALWAYS leave the program focus in the Transcript after adding a time code.
        wx.CallAfter(self.SetFocus)

    def InsertRisingIntonation(self):
        """ Insert the Rising Intonation (Up Arrow) symbol """
        # If not in Edit Mode ...
        if not self.IsEditable():
            # ... don't edit this!
            return
        # Get the current Style
        tmpStyle = self.GetDefaultStyle()
        # Check to see if tmpStyle is GOOD.
        if not tmpStyle.GetFont().IsOk():
            # If not, use the current font
            tmpStyle = self.txtAttr
        # Get the Font Name and Size
        if tmpStyle.GetFont().IsOk():
            fontName = tmpStyle.GetFont().GetFaceName()
            fontSize = tmpStyle.GetFont().GetPointSize()
        # If the font is still having problems, just get the defaults!
        else:
            fontName = TransanaGlobal.configData.defaultFontFace
            fontSize = TransanaGlobal.configData.defaultFontSize
        # Change the Style to Special Symbol Font Fact, and adjust the current size as configured
        self.SetTxtStyle(fontFace=TransanaGlobal.configData.specialFontFace, fontSize=TransanaGlobal.configData.specialFontSize)
        # Define the Up Arrow character
        ch = unicode('\xe2\x86\x91', 'utf8')  # \u2191
        # Insert the character
        self.WriteText(ch)
        # Restore the formatting
        self.SetTxtStyle(fontFace=fontName, fontSize=fontSize)
        # Add a space here (to anchor the formatting!) and then backspace over it
#        self.WriteText(' ' + wx.WXK_BACK)

    def InsertFallingIntonation(self):
        """ Insert the Falling Intonation (Down Arrow) symbol """
        # If not in Edit Mode ...
        if not self.IsEditable():
            # ... don't edit this!
            return
        # Get the current Style
        tmpStyle = self.GetDefaultStyle()
        # Check to see if tmpStyle is GOOD.
        if not tmpStyle.GetFont().IsOk():
            # If not, use the current font
            tmpStyle = self.txtAttr
        # Get the Font Name and Size
        if tmpStyle.GetFont().IsOk():
            fontName = tmpStyle.GetFont().GetFaceName()
            fontSize = tmpStyle.GetFont().GetPointSize()
        # If the font is still having problems, just get the defaults!
        else:
            fontName = TransanaGlobal.configData.defaultFontFace
            fontSize = TransanaGlobal.configData.defaultFontSize
        # Change the Style to Special Symbol Font Fact, and adjust the current size as configured
        self.SetTxtStyle(fontFace=TransanaGlobal.configData.specialFontFace, fontSize=TransanaGlobal.configData.specialFontSize)
        # Define the Down Arrow character
        ch = unicode('\xe2\x86\x93', 'utf8')
        # Insert the character
        self.WriteText(ch)
        # Restore the formatting
        self.SetTxtStyle(fontFace=fontName, fontSize=fontSize)
        # Add a space here (to anchor the formatting!) and then backspace over it
#        self.WriteText(' ' + wx.WXK_BACK)

    def InsertInBreath(self):
        """ Insert the In Breath (Closed Dot) symbol """
        # If not in Edit Mode ...
        if not self.IsEditable():
            # ... don't edit this!
            return
        # Get the current Style
        tmpStyle = self.GetDefaultStyle()
        # Check to see if tmpStyle is GOOD.
        if not tmpStyle.GetFont().IsOk():
            # If not, use the current font
            tmpStyle = self.txtAttr
        # Get the Font Name and Size
        if tmpStyle.GetFont().IsOk():
            fontName = tmpStyle.GetFont().GetFaceName()
            fontSize = tmpStyle.GetFont().GetPointSize()
        # If the font is still having problems, just get the defaults!
        else:
            fontName = TransanaGlobal.configData.defaultFontFace
            fontSize = TransanaGlobal.configData.defaultFontSize
        # Change the Style to Special Symbol Font Fact, and adjust the current size as configured
        self.SetTxtStyle(fontFace=TransanaGlobal.configData.specialFontFace, fontSize=TransanaGlobal.configData.specialFontSize)
        # Define the Closed Dot character
        ch = unicode('\xe2\x80\xa2', 'utf8')
        # Insert the character
        self.WriteText(ch)
        # Restore the formatting
        self.SetTxtStyle(fontFace=fontName, fontSize=fontSize)
        # Add a space here (to anchor the formatting!) and then backspace over it
#        self.WriteText(' ' + wx.WXK_BACK)

    def InsertWhisper(self):
        """ Insert the Whisper (Open Dot) symbol """
        # If not in Edit Mode ...
        if not self.IsEditable():
            # ... don't edit this!
            return
        # Get the current Style
        tmpStyle = self.GetDefaultStyle()
        # Check to see if tmpStyle is GOOD.
        if not tmpStyle.GetFont().IsOk():
            # If not, use the current font
            tmpStyle = self.txtAttr
        # Get the Font Name and Size
        if tmpStyle.GetFont().IsOk():
            fontName = tmpStyle.GetFont().GetFaceName()
            fontSize = tmpStyle.GetFont().GetPointSize()
        # If the font is still having problems, just get the defaults!
        else:
            fontName = TransanaGlobal.configData.defaultFontFace
            fontSize = TransanaGlobal.configData.defaultFontSize
        # Change the Style to Special Symbol Font Fact, and adjust the current size as configured
        self.SetTxtStyle(fontFace=TransanaGlobal.configData.specialFontFace, fontSize=TransanaGlobal.configData.specialFontSize)
        # Define the Open Dot character
        ch = unicode('\xc2\xb0', 'utf8')
        # Insert the character
        self.WriteText(ch)
        # Restore the formatting
        self.SetTxtStyle(fontFace=fontName, fontSize=fontSize)
        # Add a space here (to anchor the formatting!) and then backspace over it
#        self.WriteText(' ' + wx.WXK_BACK)

    def GetStyleAt(self, pos):
        """ Determine the style of the character at the given position """
        # Create a RichTextAttr object to hold the style
        tmpStyle = richtext.RichTextAttr()
        # Get the style for the specified position
        style = self.GetStyle(pos, tmpStyle)
        # Return that style to the user
        return tmpStyle

    def ClearDoc(self):
        """ Clear the wxRichTextCtrl """
        # Create default font specifications
        fontColor = wx.Colour(0, 0, 0)
        fontBgColor = wx.Colour(255, 255, 255)
        fontFace = TransanaGlobal.configData.defaultFontFace
        fontSize = TransanaGlobal.configData.defaultFontSize

        # If the current font is defined ...
        if self.txtAttr != None:
            # ... delete it to ensure it is cleared
            del(self.txtAttr)
        # Create a default font object
        self.txtAttr = richtext.RichTextAttr()

        # Stop adding to the UNDO stack
        self.BeginSuppressUndo()
        # Prepare the control for massive change
        self.Freeze()

        # Populate the Default Font object with the default font specifications
        self.SetTxtStyle(fontColor = fontColor, fontBgColor = fontBgColor, fontFace = fontFace, fontSize = fontSize,
                         fontBold = False, fontItalic = False, fontUnderline = False)

        # Clear the control.  This must occur AFTER the Style is set, or old styles could infect the new transcript.
        self.Clear()
        # Finish up change handling
        self.Thaw()

        if DEBUG:
            print "RichTextEditCtrl.ClearDoc():"
            print "  Font Face:", self.txtAttr.GetFont().GetFaceName()
            print "  Font Size:", self.txtAttr.GetFont().GetPointSize()
            print "  Font Color:", self.txtAttr.GetTextColour()
            print "  Font Background:", self.txtAttr.GetBackgroundColour()
            print "  Font Weight:", self.txtAttr.GetFont().GetWeight(), wx.FONTWEIGHT_NORMAL, wx.FONTWEIGHT_BOLD
            print "  Font Style:", self.txtAttr.GetFont().GetStyle(), wx.FONTSTYLE_NORMAL, wx.FONTSTYLE_ITALIC
            print "  Font Underline:", self.txtAttr.GetFont().GetUnderlined()
            print "  Alignment:", self.txtAttr.GetAlignment(), wx.TEXT_ALIGNMENT_LEFT, wx.TEXT_ALIGNMENT_CENTER, wx.TEXT_ALIGNMENT_RIGHT
            print "  Left Indent:", self.txtAttr.GetLeftIndent(), self.txtAttr.GetLeftSubIndent()
            print "  Right Indent:", self.txtAttr.GetRightIndent()
            print "  Line Spacing:", self.txtAttr.GetLineSpacing(), wx.TEXT_ATTR_LINE_SPACING_NORMAL, wx.TEXT_ATTR_LINE_SPACING_HALF, wx.TEXT_ATTR_LINE_SPACING_TWICE
            print "  Spacing Before:", self.txtAttr.GetParagraphSpacingBefore()
            print "  Spacing After:", self.txtAttr.GetParagraphSpacingAfter()
            print "  Tabs:", self.txtAttr.GetTabs()
            print

        # Allow undo again
        self.EndSuppressUndo()

    def LoadRTFData(self, text, clearDoc=True):
        """ Load Rich Text Format data into the RichTextEditCtrl """
        # Prepare the control for data
        self.Freeze()
        self.BeginSuppressUndo()
        # If clearing the document first is requested ...
        if clearDoc:
            # Clear the Control AND the default text attributes
            self.ClearDoc()
        # Start exception handling
        try:
            # Use the custom RTF Handler
            handler = PyRTFParser.PyRichTextRTFHandler()
            # Load the RTF text string via the XML Handler.
            # Note that for RTF, the wxRichTextCtrl CONTROL is passed.
            handler.LoadString(self, buf=text)
        # exception handling for Memory Errors
        except exceptions.MemoryError:
            # Create the error message
            prompt = _("Memory Error.  This RTF file is too large to import.\nTry removing some images from the document.")
            # Display the Error Message
            errDlg = Dialogs.ErrorDialog(self, prompt)
            errDlg.ShowModal()
            errDlg.Destroy()
        # exception handling
        except:
            print "Custom RTF Handler Load failed"
            print
            print sys.exc_info()[0], sys.exc_info()[1]
            print traceback.print_exc()
            print
            pass
        # Signal the end of changing the control
        self.EndSuppressUndo()
        self.Thaw()

    def LoadXMLData(self, text, clearDoc=True):
        """ Load XML data into the RichTextEditCtrl """
        # Prepare the control for data
        self.Freeze()
        self.BeginSuppressUndo()
        # If clearing the document first is requested ...
        if clearDoc:
            # Clear the Control AND the default text attributes
            self.ClearDoc()
        # Start exception handling
        try:
            # We need to setup a StringIO object to emulate a file
            stream = cStringIO.StringIO(text)
            # Create an XML Handler
            handler = richtext.RichTextXMLHandler()
            # Load the XML text via the XML Handler.
            # Note that for XML, the BUFFER is passed.
            handler.LoadStream(self.GetBuffer(), stream)
        # exception handling
        except:
            print "XML Handler Load failed"
            print
            print sys.exc_info()[0], sys.exc_info()[1]
            print traceback.print_exc()
            print
            pass
        # Signal the end of changing the control
        self.EndSuppressUndo()
        self.Thaw()

    def SaveRTFDocument(self, filepath):
        """ Save the control's contents as a Rich Text Format document """
        # If there's text in the wxRichTextCtrl ...
        if len(self.GetValue()) > 0:
            # Begin exception handling
            try:
                # If an existing file is selected ...
                if (filepath != ""):
                    # Use the custom RTF Handler
                    handler = PyRTFParser.PyRichTextRTFHandler()
                    # Save the file with the custom RTF Handler.
                    # The custom RTF Handler can take either a wxRichTextCtrl or a wxRichTextBuffer argument.
                    handler.SaveFile(self.GetBuffer(), filepath)
            # Exception Handling
            except:
                print "Custom RTF Handler Save failed"
                print
                print sys.exc_info()[0], sys.exc_info()[1]
                print traceback.print_exc()
                print
                pass

    def SaveXMLDocument(self, filepath):
        """ Save the control's contents as an XML Format document """
        # If there's text in the wxRichTextCtrl ...
        if len(self.GetValue()) > 0:
            # Begin exception handling
            try:
                # If an existing file is selected ...
                if (filepath != ""):
                    # Use the standard XML Handler
                    handler = richtext.RichTextXMLHandler()
                    # Save the file with the standard XML Handler.
                    handler.SaveFile(self.GetBuffer(), filepath)
            # Exception Handling
            except:
                print "XML Handler Save failed"
                print
                print sys.exc_info()[0], sys.exc_info()[1]
                print traceback.print_exc()
                print
                pass

##    def AddText(self, text):
##        """ Add text to the RichTextEditCtrl, implemented for wxSTC compatibility """
##        # Append the text at the end of the RTC buffer
##        self.AppendText(text)

##    def InsertHiddenText(self, text):
##        self.AppendText(text)

    def OnKeyDown(self, event):
        """ Handler for EVT_KEY_DOWN events for use with Transana.
            This handles deletion of time codes. """
        # Assume that event.Skip() should be called unless proven otherwise
        shouldSkip = True
        # Create some variables to make this code a little simpler to read
        ctrl = event.GetEventObject()
        ip = ctrl.GetInsertionPoint()
        sel = (ctrl.GetSelectionStart(), ctrl.GetSelectionEnd())

        # Capture the current text style
        tmpStyle = self.GetStyleAt(ip)

        # If the Delete key is pressed ...
        if event.GetKeyCode() == wx.WXK_DELETE:
            # See if we have an insertion point rather than a selection.
            if sel[0] == sel[1]:
                # If so, look at the character to the right, the one to be deleted
                ctrl.SetSelection(ip, ip+1)
                # If that character is a time code ...
                if ctrl.GetStringSelection() == TIMECODE_CHAR:
                    # ... batch the undo
                    self.undoString = ''
                    ctrl.BeginBatchUndo(self.undoString)

                # Delete the character to be deleted
                ctrl.DeleteSelection()

                # If we're deleting a Time Code, we need to remove it from the Time Codes list ...
                if self.IsStyleHiddenAt(ip):
                    # Initialize a string to store the deleted Time Code value
                    tcVal = ''
                    # Make a temporary copy of the Insertion Point, adding one to skip the "<" in the Time Code
                    ipTemp = ip + 1
                    # Select the next character
                    ctrl.SetSelection(ipTemp,  ipTemp + 1)
                    # While the next character is part of the TIME CODE DATA ...
                    while ctrl.GetStringSelection() in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']:
                        # Add the digit to the deleted time code value string
                        tcVal += ctrl.GetStringSelection()
                        # Increment the temporary insertion point
                        ipTemp += 1
                        # Select the next character
                        ctrl.SetSelection(ipTemp,  ipTemp + 1)

                    try:
                        # Remove the temporary time code value from the Time Codes List
                        self.timecodes.remove(int(tcVal))
                    except:
                        pass

                # if the style shows a point size of 1, that signals "hidden" text which needs to be removed
                while self.IsStyleHiddenAt(ip):
                    # As long as the style is "hidden", keep deleting characters
                    ctrl.SetSelection(ip, ip+1)
                    ctrl.DeleteSelection()
                    ip = ctrl.GetInsertionPoint()
                # We are handling the delete manually, so event.Skip() should be skipped!
                shouldSkip = False
            if self.IsEditable():
                # Remember the current selection
                selection = self.GetSelection()
                # If the selection start differs from the selection end (i.e. we have text selected for over-type...)
                if selection[0] != selection[1]:
                    # ... remove any time codes in that selection from the self.timecodes list
                    self.RemoveTimeCodeData(self.GetStringSelection())

        # If the backspace key is pressed ...
        elif event.GetKeyCode() == wx.WXK_BACK:
            # See if we have an insertion point rather than a selection.
            if sel[0] == sel[1]:
                # Capture the current text style of the character BEFORE the insertion point
                # if the style shows a point size of 1, that signals "hidden" text which needs to be removed
                if self.IsStyleHiddenAt(ip - 1):
                    # ... batch the undo
                    self.undoString = ''
                    ctrl.BeginBatchUndo(self.undoString)
                # Go ahead and call Skip to implement the backspace
                event.Skip()
                # If we're deleting a Time Code, we need to remove it from the Time Codes list ...
                if self.IsStyleHiddenAt(ip - 1):
                    # Initialize a string to store the deleted Time Code value
                    tcVal = ''
                    # Make a temporary copy of the Insertion Point, reduced by two to skip the ">" in the Time Code
                    ipTemp = ip - 2
                    # Select the next character
                    ctrl.SetSelection(ipTemp-1,  ipTemp)
                    # While the next character is part of the TIME CODE DATA ...
                    while ctrl.GetStringSelection() in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']:
                        # Add the digit to the deleted time code value string
                        tcVal = ctrl.GetStringSelection() + tcVal
                        # Increment the temporary insertion point
                        ipTemp -= 1
                        # Select the next character
                        ctrl.SetSelection(ipTemp - 1,  ipTemp)

                    try:
                        # Remove the temporary time code value from the Time Codes List
                        self.timecodes.remove(int(tcVal))
                    except:
                        pass

                # if the style shows a point size of 1, that signals "hidden" text which needs to be removed
                while self.IsStyleHiddenAt(ip - 1):
                    # As long as the style is "hidden", keep deleting characters
                    ip -= 1
                    ctrl.SetSelection(ip, ip+1)
                    ctrl.DeleteSelection()
                    ctrl.SetInsertionPoint(ip)
                # We have already called event.Skip(), so the next event.Skip() should be skipped!
                shouldSkip = False
            elif self.IsEditable():
                # Remember the current selection
                selection = self.GetSelection()
                # If the selection start differs from the selection end (i.e. we have text selected for over-type...)
                if selection[0] != selection[1]:
                    # ... remove any time codes in that selection from the self.timecodes list
                    self.RemoveTimeCodeData(self.GetStringSelection())

##        # If Alt-F1 is pressed ...
##        elif event.AltDown() and event.GetKeyCode() == wx.WXK_F1:
##            # Iterate through 10 characters
##            for x in range(0, 10):
##                if ip + x <= self.GetLastPosition():
##                    # Move the insertion point
##                    ctrl.SetInsertionPoint(ip + x)
##                    # Select the next character
##                    ctrl.SetSelection(ip + x, ip + x + 1)
##                    # Get that character's style
##                    tmpStyle = self.GetStyleAt(ip + x)
##                    # If it's NOT a time code ...
##                    if ctrl.GetStringSelection() != TIMECODE_CHAR:
##                        # print basic formatting information
####                        self.PrintTextAttr("%05d  %s" % (ip + x, ctrl.GetStringSelection()), tmpStyle)
##                        print "%05d  %s" % (ip + x, ctrl.GetStringSelection())
##                    # Otherwise
##                    else:
##                        # don't print the character, as that raises an exception
####                        self.PrintTextAttr("%05d  %s" % (ip + x, 'TC'), tmpStyle)
##                        print "%05d  %s" % (ip + x, 'TC')
##
####            self.PrintTextAttr("Current:", self.txtAttr)
####            self.PrintTextAttr("Default:", self.GetDefaultStyle())
####            self.PrintTextAttr("Basic:", self.GetBasicStyle())
##                        
##            # Reset the insertion point back to where we were
##            ctrl.SetInsertionPoint(ip)
##
##            print self.timecodes
##            print

        # Otherwise, if we're in EDIT mode and have not just received a MOVEMENT key ...
        elif self.IsEditable() and \
             not (event.GetKeyCode() in [wx.WXK_LEFT, wx.WXK_RIGHT, wx.WXK_DOWN, wx.WXK_UP,
                                         wx.WXK_PAGEUP, wx.WXK_PAGEDOWN, wx.WXK_HOME, wx.WXK_END,
                                         wx.WXK_SHIFT, wx.WXK_ALT, wx.WXK_CONTROL, wx.WXK_MENU,
                                         wx.WXK_WINDOWS_LEFT, wx.WXK_WINDOWS_MENU, wx.WXK_WINDOWS_RIGHT]):

            # Remember the current selection
            selection = self.GetSelection()
            # If the selection start differs from the selection end (i.e. we have text selected for over-type...)
            if selection[0] != selection[1]:
                # ... remove any time codes in that selection from the self.timecodes list
                self.RemoveTimeCodeData(self.GetStringSelection())

        # If the last character is a Time Code (with hidden data), then a newly loaded transcript
        # could not be edited beyond that point, as all new typing was hidden.  This fixes that.
        
        # Get the Default Style, which is the style that will be used for the current character
        defStyle = self.GetDefaultStyle()
        # If our character is in Time Code or Hidden Style ...
        while (ip > 0) and \
              (self.CompareFormatting(defStyle, self.txtTimeCodeAttr, fullCompare=False) or \
              self.CompareFormatting(defStyle, self.txtHiddenAttr, fullCompare=False) or \
              self.CompareFormatting(defStyle, self.txtTimeCodeHRFAttr, fullCompare=False)):
            # Get the current position's style
            tmpStyle = self.GetStyleAt(ip)
            # move the pointer to the previous position
            ip -= 1
            # Set the Default Style
            self.SetDefaultStyle(tmpStyle)

            # Get the new Default Style for the next comparison
            defStyle = self.GetDefaultStyle()
            # If we get to the beginning of the document ...
            if ip == 0:
                # Call CheckFormatting(), which should set the formatting correctly
                self.CheckFormatting()
                # ... and we can get out of this while loop.
                break
            # Signal that we have adjusted the formatting for a Time Code
            self.timeCodeFormatAdjustment = True
            
        # If we should call event.Skip()...
        if shouldSkip:
            # ... then call event.Skip()!
            event.Skip()

    def RemoveTimeCodeData(self, txt):
        """  Remove any time code values contained in "txt" from the self.timecodes list """
        # Get the position of the first "<" character in the string
        minPos = txt.find('<', 0)
        # Determine the position of the selection within the RichTextCtrl as a whole
        sel = self.GetSelection()
        # While there is a "<" character ...
        while (minPos > -1):
            # ... check to see if it is HIDDEN, i.e. part of a time code.  Use the position in the RichTextCtrl,
            #     not merely the position within the string being evaluated.
            if self.IsStyleHiddenAt(minPos + sel[0]):
                # Identify the start and end of the time code data
                st = minPos
                # Get the ">" character that follows the current "<" character
                end = txt.find(">", minPos + 1)
                # Capture the time code value
                tcval = int(txt[st+1 : end])
                # Remove this value from the self.timecodes list
                self.timecodes.remove(tcval)
                # Remove it (including the angle brackets) from the string to allow the next one to be found
                txt = txt[:st] + txt[end + 1:]
            # Determin the position of the next "<" character.
            # (This allows for "<" characters in the text that are not part of time codes!)
            minPos = txt.find('<', minPos + 1)

    def OnKey(self, event):
        """ Handler for EVT_CHAR events for use with Transana """
        # Create some variables to make this code a little simpler to read
        ctrl = event.GetEventObject()
        # Get the current String Selection and remember it for later
        self.keyStringSelection = ctrl.GetStringSelection()

        # At the moment, there's nothing special to do.
        event.Skip()

    def OnKeyUp(self, event):
        """ Handler for EVT_KEY_UP events for use with Transana.
            This handles cursor-move over time codes. """
        # Create some variables to make this code a little simpler to read
        ctrl = event.GetEventObject()
        # Get the current insertion point
        ip = ctrl.GetInsertionPoint()
        # Get the current String Selection, noted earlier in the OnKey method
        st = self.keyStringSelection
        # if Cursor Left is pressed ...
        if event.GetKeyCode() == wx.WXK_LEFT:
            # If the current style is "hidden" text, it needs to be skipped ...
            while self.IsStyleHiddenAt(ip) and (ip > 0):

                if ctrl.HasSelection():
                    (start, end) = ctrl.GetSelection()
                    # Setting from END to START ensures cursor presses move the correct end of the selection!!
                    ctrl.SetSelection(start - 1, end)
                    ip = ip - 1
                else:
                    # ... so move a character to the left
                    ctrl.MoveLeft()
                    # Get the new insertion point
                    ip = ctrl.GetInsertionPoint()

        # if Cursor Right, Cursor Down, or Cursor Up is pressed ...
        elif event.GetKeyCode() in [wx.WXK_RIGHT, wx.WXK_DOWN, wx.WXK_UP, wx.WXK_END, wx.WXK_PAGEDOWN, wx.WXK_PAGEUP]:
            # If the current style is "hidden" text, it needs to be skipped ...
            while self.IsStyleHiddenAt(ip) and (ip < self.GetLastPosition()):
                # if we're at the end of a LINE ...
                if (self.GetCharAt(ip) == 10) and not self.GetReadOnly():
                    # ... we don't want to go to the next line.  Instead, insert a SPACE ...
                    self.WriteText(' ')
                    # If the current style is Time Code formatting ...
                    if self.CompareFormatting(self.txtAttr, self.txtTimeCodeAttr, fullCompare=False):
                        # ... use the default text style
                        self.SetStyle((ip, ip + 1), self.txtOriginalAttr)
                    # If the current style is NOT time code formatting ...
                    else:
                        # ... use the current style
                        self.SetStyle((ip, ip + 1), self.txtAttr)
                    # ... and move the insertion point before the space
                    self.SetInsertionPoint(ip)
                else:

                    if ctrl.HasSelection():
                        tmpSel = ctrl.GetSelection()
                        ctrl.SetSelection(tmpSel[0], tmpSel[1] + 1)
                        ip = ctrl.GetInsertionPoint()
                    else:
                        # ... so move a character to the right
                        ctrl.MoveRight()
                        # Get the new insertion point
                        ip = ctrl.GetInsertionPoint()

        # Call event.Skip() was removed to prevent MULTIPLE CALLS to this method
        # event.Skip()

        # If the transcript is editable and there WAS a selection and there's no longer a selection ...
        if self.IsEditable() and (len(st) > 0) and (ctrl.GetSelectionStart() == ctrl.GetSelectionEnd()):
            # Look for Time Codes in the selection
            while TIMECODE_CHAR in st:
                # Determine the Time Code Data value for the first time code
                tcVal = st[st.find(TIMECODE_CHAR) + 2 : st.find('>', st.find(TIMECODE_CHAR))]
                try:
                    # Remove the temporary time code value from the Time Codes List
                    self.timecodes.remove(int(tcVal))
                except:
                    pass
                # Remove the first time code and its data from the selection string
                st = st[st.find('>', st.find(TIMECODE_CHAR)) + 1:]
        # We're done, so clear the Selection String variable
        self.keyStringSelection = ''

        # if either Backspace or Delete was pressed ...
        if event.GetKeyCode() in [wx.WXK_BACK, wx.WXK_DELETE]:
            # ... and if we are in the midst of a batch undo specification ...
            if ctrl.BatchingUndo():
                # ... then signal that the undo batch is complete.
                 ctrl.EndBatchUndo()

        # If we adjusted away from Time Code Style above ...
        if self.timeCodeFormatAdjustment:
            # ... we need to return to Time Code Style, or the Time Code will be displayed incorrectly
            self.SetDefaultStyle(self.txtTimeCodeAttr)

            # Signal that Time Code Style adjustment is complete.
            self.timeCodeFormatAdjustment = False

        # If we have a key that changes the cursor position ...
        if event.GetKeyCode() in [wx.WXK_LEFT, wx.WXK_RIGHT, wx.WXK_UP, wx.WXK_DOWN,
                                  wx.WXK_HOME, wx.WXK_END, wx.WXK_PAGEUP, wx.WXK_PAGEDOWN,
                                  wx.WXK_NUMPAD_LEFT, wx.WXK_NUMPAD_RIGHT, wx.WXK_NUMPAD_UP, wx.WXK_NUMPAD_DOWN,
                                  wx.WXK_NUMPAD_HOME, wx.WXK_NUMPAD_END, wx.WXK_NUMPAD_PAGEUP, wx.WXK_NUMPAD_PAGEDOWN]:
            # ... update the style to the style at the new position
            self.txtAttr = self.GetStyleAt(ip)

##        print "RichTextEditCtrl.OnKeyUp()", ip
##        self.PrintTextAttr("RichTextEditCtrl.OnKeyUp()", self.txtAttr)
##        print

    def OnLeftDown(self, event):
        """ Handles the Left Mouse Down event """
        
        # NOTE:  This does NOT get called in Transana!
        
        # If there is currently a selection ...
        if event.GetEventObject().HasSelection():
            # Determine the start and end character numbers of the current selection
            textSelection = event.GetEventObject().GetSelection()
            # Determine the character number of the current mouse position (!)
#            mousePos = event.GetEventObject().HitTest(event.GetPosition())[1]
            # If we're using wxPython 2.8.x.x ...
            if wx.VERSION[:2] == (2, 8):
                # ... use HitTest()
                mousePos = event.GetEventObject().HitTest(event.GetPosition())[1]
            # If we're using a later wxPython version ...
            else:
                # ... use HitTestPos()
                mousePos = event.GetEventObject().HitTestPos(event.GetPosition())[1]
         
            # If the Mouse Character is inside the selection ...
            if (textSelection[0] <= mousePos) and (mousePos < textSelection[1]):
                # Create a Text Data Object, which holds the text that is to be dragged
                tdo = wx.PyTextDataObject(event.GetEventObject().GetStringSelection())
                # Create a Drop Source Object, which enables the Drag operation
                tds = wx.DropSource(event.GetEventObject())
                # Associate the Data to be dragged with the Drop Source Object
                tds.SetData(tdo)
                # Remember the original selection
                originalSelection = event.GetEventObject().GetSelection()
                # Remember the STRING in the original selection
                originalStringSelection = event.GetEventObject().GetStringSelection()
                # Intiate the Drag Operation
                result = tds.DoDragDrop(True)
                # If we have a MOVE (instead of a COPY) ...
                if result == wx.DragMove:
                    # ... we need to delete the selection.
                    if originalStringSelection == event.GetEventObject().GetRange(originalSelection[0], originalSelection[1]):
                        event.GetEventObject().Delete(originalSelection)
                    else:
                        deleteSelection = (originalSelection[0] + len(originalStringSelection), originalSelection[1] + len(originalStringSelection))
                        event.GetEventObject().Delete(deleteSelection)
            # If the mouse is outside the selection ...
            else:
                # Skip here allows new selection to be made
                event.Skip()
        # If there is NOT currently a selection ...
        else:
            # Cursor getting lost on OS X.  This is an attempt to make it visible
            self.GetCaret().Show()
            # Skip here allows selection to be made
            event.Skip()

    def OnLeftUp(self, event):
        """ Handler for EVT_LEFT_UP mouse events for use with Transana.
            This prevents the cursor from being placed between a time code and its hidden data. """
        # Create some variables to make this code a little simpler to read
        # Get the Control
        ctrl = event.GetEventObject()
        # Get the current insertion point
        ip = ctrl.GetInsertionPoint()
        # Get the current selection
        sel = ctrl.GetSelection()
        # If there's no SELECTION ...
        if sel == (-2, -2):
            # Check to see if we're in the middle of "hidden" text which needs to be skipped ...
            while self.IsStyleHiddenAt(ip) and (ip < self.GetLastPosition()):
                # if we're at the end of a LINE ...
                if (self.GetCharAt(ip) == 10) and not self.GetReadOnly():
                    # ... we don't want to go to the next line.  Instead, insert a SPACE ...
                    self.WriteText(' ')
                    # Apply the paragraph formatting change
                    self.SetStyle((ip, ip + 1), self.txtAttr)
                    # ... and move the insertion point before the space
                    self.SetInsertionPoint(ip)
                # If we're NOT at the end of a line ...
                else:
                    # ... move a character to the right
                    ctrl.MoveRight()
                    # Get the new insertion point
                    ip = ctrl.GetInsertionPoint()
        # If there IS a selection ...
        else:
            # Get the start and end points
            (sp, ep) = self.GetSelection()
            # Check to see if we START in the middle of "hidden" text which needs to be skipped ...
            while self.IsStyleHiddenAt(sp) and (sp < self.GetLastPosition()):
                # ... and if so, increase the selection by one
                ctrl.SetSelection(sp + 1, ep)
                # Get the start and end points
                sp += 1
            # Check to see if we're in the middle of "hidden" text which needs to be skipped ...
            while self.IsStyleHiddenAt(ep) and (ep < self.GetLastPosition()):
                # ... and if so, increase the selection by one
                ctrl.SetSelection(sp, ep + 1)
                # Get the start and end points
                ep += 1
        # Call event.Skip()
        event.Skip()

        # If we're not at the first character in the document AND
        # we're not at the first character following a time code ...
        if (ip > 1) and not self.IsStyleHiddenAt(ip - 1):
            # ... update the style to the style for the character PRECEEDING the cursor
            self.txtAttr = self.GetStyleAt(ip - 1)
        # Otherwise ...
        else:
            # update the style to the style for the character FOLLOWING the cursor
            self.txtAttr = self.GetStyleAt(ip)

##        print "RichTextEditCtrl.OnLeftUp()", ip
##        self.PrintTextAttr("RichTextEditCtrl.OnLeftUp()", self.txtAttr)
##        print

    def OnMotion(self, event):
        """ Handle Mouse Movement Events """
        # If there's a SELECTION in the transcript, we need to change the cursor when we mouse over it
        # to indicate that a DRAG is possible
        
        # Process the underlying OnMotion events
        event.Skip()
        # Determine the current Mouse Position
        (x, y) = wx.GetMousePosition()
        # If the control has a selection ...
        if event.GetEventObject().HasSelection():
            # Determine the start and end character numbers of the current selection
            textSelection = event.GetEventObject().GetSelection()
            # Determine the character number of the current mouse position (!)
            # If we're using wxPython 2.8.x.x ...
            if wx.VERSION[:2] == (2, 8):
                # ... use HitTest()
                mousePos = event.GetEventObject().HitTest(event.GetPosition())[1]
            # If we're using a later wxPython version ...
            else:
                # ... use HitTestPos()
                mousePos = event.GetEventObject().HitTestPos(event.GetPosition())[1]
            # If the Mouse Character is inside the selection ...
            if (textSelection[0] <= mousePos) and (mousePos < textSelection[1]):
                # ... change the cursor of the PARENT object (not sure why Parent is necessary, but it is.)
                event.GetEventObject().parent.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

    def OnURL(self, event):
        """ Handle EVT_TEXT_URL events """

        # Get the URL for the hyperlink
        hyperlink = event.GetString()

        # Split the hyperlink at the colon to determine what type of link it is and what data it holds
        (linkType, data) = hyperlink.split(':')
        # If we have a Transana Hyperlink ...
        if linkType.lower() == 'transana':
            # Break the data apart at the equal sign to determine what type of object we have and its object number
            (objType, objNum) = data.split('=')

            # NOTE:  We cannot support EPISODE links, as we require a Transcript Number!
            
            # If we have a CLIP link ...
            if objType.lower() == 'clip':
                # ... and if there's a defined Control Object ...
                if self.parent.controlObject != None:
                    # ... load the Clip
                    self.parent.controlObject.LoadClipByNumber(int(objNum))
                # ... if there's NO defined Control Object ...
                else:
                    # ... print an error message
                    print "RichTextEditCtrl_RTC.OnURL():  ControlObject is None!!"
            # if we have a Snapshot Link ...
            elif objType.lower() == 'snapshot':
                # ... and if there's a defined Control Object ...
                if self.parent.controlObject != None:
                    # ... get the Snapshot Object ...
                    tmpSnapshot = Snapshot.Snapshot(int(objNum))
                    # ... and load the Snapshot Window
                    self.parent.controlObject.LoadSnapshot(tmpSnapshot)
                # ... if there's NO defined Control Object ...
                else:
                    # ... print an error message
                    print "RichTextEditCtrl_RTC.OnURL():  ControlObject is None!!"
                
        # If we have an HTTP link ...
        elif linkType.lower() == 'http':
            # ... use Python's webbrowser module to open a web browser!
            webbrowser.open(hyperlink, new=True)
        # Otherwise ...
        else:
            # ... display an error message.
            prompt = unicode(_('Transana does not support "%s" hyperlinks at this time.'), 'utf8')
            tmpDlg = Dialogs.ErrorDialog(self, hyperlink + u'\n\n' + prompt % linkType)
            tmpDlg.ShowModal()
            tmpDlg.Destroy()

    def CheckFormatting(self):
        """ We need to make sure we don't get into trouble with formatting.  There are a couple of odd situations
            that need to be handled.   We'll do that here.  """

        if DEBUG:
            print
            print "RichTextEditCtrl_RTC.CheckFormatting()", self.GetInsertionPoint(), self.GetLastPosition(), self.GetSelection()

            print
            print "I need to completely re-write this.  Can we just iterate through styles (current, minus one, default, basic, Original) until we hit one that's not TimeCode or Hidden?"
            print

        # If self.txtAttr is not defined, there's no transcript yet.
        if (self.txtAttr is None) or (self.GetSelection() != (-2, -2)) or self.GetReadOnly():

            if DEBUG:
                if (self.txtAttr is None):
                    print "txtAttr is None.  Exiting."

                if (self.GetSelection() != (-2, -2)):
                    print "(self.GetSelection() != (-2, -2)) -- Exiting"

                if self.GetReadOnly():
                    print "self.GetReadOnly() -- exiting"
            
            # ... so just exit
            return

        # Note if the control has been edited
        isModified = self.IsModified()

        # Get the current position
        pos = self.GetInsertionPoint()
        # Get the style for the current character        
        textAttr = self.GetStyleAt(pos)

        # If we're at the beginning of the file, we need to INITIALIZE the default and basic styles.
        if (pos == 0):
            # If we're at a time code, skip the time code and hidden text
            while (self.CompareFormatting(textAttr, self.txtHiddenAttr, fullCompare=False) or \
                   self.CompareFormatting(textAttr, self.txtTimeCodeAttr, fullCompare=False) or \
                   self.CompareFormatting(textAttr, self.txtTimeCodeHRFAttr, fullCompare=False)) and \
                  ((pos == 0) or (self.GetCharAt(pos-1) != 32)):
                # Look one character to the right
                pos += 1
                # Get the styleof the new character
                textAttr = self.GetStyleAt(pos)

            # Set the Default Style
            self.SetDefaultStyle(textAttr)
            # Set the Basic Style
            self.SetBasicStyle(textAttr)
            # Get the Default style for comparison purposes
            defTextAttr = textAttr
            # Get the Basic Style for comparison purposes
            basicTextAttr = textAttr

        # If we're NOT at the beginning of the file ...
        else:
            # Get the Default style for comparison purposes
            defTextAttr = self.GetDefaultStyle()
            # Get the Basic Style for comparison purposes
            basicTextAttr = self.GetBasicStyle()

        # If the character position we're looking at has changed ...
        if pos != self.GetInsertionPoint():
            # ... we need to update the style for the new position
            textAttr = self.GetStyleAt(pos)
            # ... and we need to move the insertion point!
            self.SetInsertionPoint(pos)

        # Let's find the last character with GOOD formatting
        # start at the current position
        tmpPos = pos
        # Initialize the formatting as NOT GOOD
        textAttrMinusOne = None

        if DEBUG :
            self.PrintTextAttr("Character at pos %d:" % pos, textAttr)
            self.PrintTextAttr("Current:", self.txtAttr)
            self.PrintTextAttr("Default:", defTextAttr)

        # Create an ordered list of styles to compare.
        stylesToCheck = [(textAttr, False), (self.txtAttr, False), (defTextAttr, False), (textAttrMinusOne, True), (basicTextAttr, False), (self.txtOriginalAttr, False)]
        # For each style to compare ...
        for (style, needToPopulate) in stylesToCheck:
            # If we need to FIND the appropriate style information ...
            if needToPopulate:
                # While there are characters to check and the formatting is NOT GOOD ...
                while (tmpPos > 0) and (textAttrMinusOne == None):
                    # Move backwards one character
                    tmpPos -= 1
                    # Get the character's formatting
                    textAttrMinusOne = self.GetStyleAt(tmpPos - 1)
                    # Check to see if the character style is Time Code or hidden ...
                    if (self.CompareFormatting(textAttrMinusOne, self.txtHiddenAttr, fullCompare=False) or \
                        self.CompareFormatting(textAttrMinusOne, self.txtTimeCodeAttr, fullCompare=False) or \
                        self.CompareFormatting(textAttrMinusOne, self.txtTimeCodeHRFAttr, fullCompare=False)):
                        # If so, label it NOT GOOD
                        textAttrMinusOne = None
                    # If NOT ...
                    else:
                        # ... signal that we can stop checking styles!
                        tmpPos = 0

                # If we didn't find any GOOD styles BEFORE the current position ...
                if textAttrMinusOne is None:
                    # Let's find the next character with GOOD formatting
                    # start at the current position
                    tmpPos = pos
                    # Initialize the formatting as NOT GOOD
                    textAttrMinusOne = None
                    # While there are characters to check and the formatting is NOT GOOD ...
                    while (tmpPos < self.GetLastPosition()) and (textAttrMinusOne == None):
                        # Move forwards one character
                        tmpPos += 1
                        # Get the character's formatting
                        textAttrMinusOne = self.GetStyleAt(tmpPos - 1)
                        # Check to see if the character style is Time Code or hidden ...
                        if (self.CompareFormatting(textAttrMinusOne, self.txtHiddenAttr, fullCompare=False) or \
                            self.CompareFormatting(textAttrMinusOne, self.txtTimeCodeAttr, fullCompare=False) or \
                            self.CompareFormatting(textAttrMinusOne, self.txtTimeCodeHRFAttr, fullCompare=False)):
                            # If so, label it NOT GOOD
                            textAttrMinusOne = None
                        # If NOT ...
                        else:
                            # ... signal that we can stop checking styles!
                            tmpPos = self.GetLastPosition() + 1

                if DEBUG:
                    if textAttrMinusOne != None:
                        self.PrintTextAttr("Minus One:", textAttrMinusOne)
                    else:
                        print "Minus One:"
                        print "  None"
                        print
                
            # Check to see if the character style is Time Code or hidden ...
            if not (self.CompareFormatting(style, self.txtHiddenAttr, fullCompare=False) or \
                    self.CompareFormatting(style, self.txtTimeCodeAttr, fullCompare=False) or \
                    self.CompareFormatting(style, self.txtTimeCodeHRFAttr, fullCompare=False)):

                textAttr = style
                
                break

        if DEBUG:

            self.PrintTextAttr("Basic:", basicTextAttr)
            self.PrintTextAttr("Original:", self.txtOriginalAttr)

            print
            print
            self.PrintTextAttr("Recommended:", textAttr)
            print
            print

        # If the font that should be used isn't the Default Font ...
        if not self.CompareFormatting(textAttr, defTextAttr):

            if DEBUG:
                print "Updating Formatting", pos

            # If we're at the Beginning, where we need to update the formatting, or the END of the document, where we lose formatting ....
            if (pos == 0) or (pos == self.GetLastPosition()):

                # We can't go to the Default or the Basic style here.  They might not be what we want in
                # a transcript that has multiple styles in it.

                if pos == 0:
                    increment = 1
                    endpoint = self.GetLastPosition()
                else:
                    increment = -1
                    endpoint = 0

                    lastChar = self.GetRange(pos - 1, pos)

                    while (pos > 0) and (len(lastChar) == 1) and (ord(lastChar) == 10):
                        pos -= 1
                        lastChar = self.GetRange(pos - 1, pos)

                while (pos != endpoint) and \
                      (self.CompareFormatting(self.GetStyleAt(pos), self.txtTimeCodeAttr, fullCompare=False) or \
                       self.CompareFormatting(self.GetStyleAt(pos), self.txtHiddenAttr, fullCompare=False) or \
                       self.CompareFormatting(self.GetStyleAt(pos), self.txtTimeCodeHRFAttr, fullCompare=False)):
                    pos += increment
                
                # ... use the default style
                desiredStyle = self.GetStyleAt(pos)

            # If we are in the middle of the document ...
            else:
                # ... use the current character's style
                desiredStyle = textAttr
                

                if DEBUG:
                    print "Style set to Current Text Position"

            # If we failed to get a valid style ...
            if desiredStyle is None:
                # ... set the style to the document's basic/default style
                desiredStyle = self.txtOriginalAttr

                # This problem seems to occur at the end of lines preceding time codes.  The cursor
                # gets placed on the next line in front of the time code, but has time code formatting.
                # Instead, let's insert a SPACE and a line break.
                self.WriteText(' \n')
                # Apply the paragraph formatting change
                self.SetStyle((pos, pos + 2), desiredStyle)
                # And move the cursor to BEFORE the space
                self.SetInsertionPoint(pos)
                # Now, if we're not at the beginning of the document ...
                if pos > 0:
                    # ... see if the last character was a line break.
                    # (It almost always is.  If it is, we end up with an extra line.)
                    if self.GetCharAt(pos - 1) == 10:
                        # Select that line break.
                        self.SetSelection(pos - 1, pos)
                        # Delete the line break.
                        self.DeleteSelection()

            # Get the font definition associated with the current style
            font1 = desiredStyle.GetFont()                

            # If we don't have a valid Font definition ...  (This is probably unnecessary)
            if not font1.IsOk():

                if DEBUG:
                    print "RichTextEditCtrl_RTC.CheckFormatting():  FONT is BAD:", font1.IsOk()

                # If we're already on the Default Style ...
                if desiredStyle == defTextAttr:
                    # ... then try the Basic Style
                    font1 = self.GetBasicStyle().GetFont()
                # If we aren't on the Default Style ...
                else:
                    # ... then try the Default Style
                    font1 = defTextAttr.GetFont()

            # If we now have a valid font ...
            if font1.IsOk():
                # Set the current text style definition to match our desired font
                self.SetTxtStyle(fontFace=font1.GetFaceName(), fontSize = font1.GetPointSize(),
                                 fontBold = (font1.GetWeight() == wx.FONTWEIGHT_BOLD),
                                 fontItalic = (font1.GetStyle() == wx.FONTSTYLE_ITALIC),
                                 fontUnderline = font1.GetUnderlined())

            else:
                if DEBUG:
                    print "RichTextEditCtrl_RTC.CheckFormatting():  FONT is REALLY BAD:", font1.IsOk()
                    print "  Need to set it to Transana Defaults!"

            # Apply the desired paragraph formatting
            self.SetTxtStyle(fontColor = desiredStyle.GetTextColour(),
                             fontBgColor = desiredStyle.GetBackgroundColour(),
                             parAlign = desiredStyle.GetAlignment(),
                             parLeftIndent = (desiredStyle.GetLeftIndent(), desiredStyle.GetLeftSubIndent()),
                             parRightIndent = desiredStyle.GetRightIndent(),
                             parLineSpacing = desiredStyle.GetLineSpacing(),
                             parSpacingBefore = desiredStyle.GetParagraphSpacingBefore(),
                             parSpacingAfter = desiredStyle.GetParagraphSpacingAfter(),
                             parTabs = desiredStyle.GetTabs())

        if DEBUG:
            self.PrintTextAttr("Final:", self.txtAttr)

        # If we have a Transcript Window ...
        if isinstance(self.parent, TranscriptionUI_RTC._TranscriptDialog):
            # ... update the Formatting Bar
            self.parent.FormatUpdate(self.txtAttr)
        
            if DEBUG and TranscriptionUI_RTC.SHOWFORMATTINGPANEL:
                self.parent.formatPanel.txt.SetLabel("%d %d" % (self.GetLastPosition(), len(self.GetFormattedSelection('XML'))))

        # Check to see if the transcript has exceeded the transcript object maximum size.
        if not self.get_read_only() and (len(self.GetFormattedSelection('XML')) > TransanaGlobal.max_allowed_packet - 80000):   # 8300000
            # If so, generate an error message.  (There's actually room for more text here.)
            prompt = _("You have reached Transana's maximum transcript size.") + "  " + _("You should save your transcript immediately.")
            prompt += '\n' + _("You may need to remove or shrink some of your images or break your transcript into multiple parts.")
            # Display the Error Message
            errDlg = Dialogs.ErrorDialog(self, prompt)
            errDlg.ShowModal()
            errDlg.Destroy()

        # If the contents had not been modified ...
        if not isModified:
            # ... this method should not mark them as having been modified.
            self.DiscardEdits()

        if DEBUG:
            print
            print ' - - - - - - - - - - - - - - - - - - - '
            print
            print

    def PrintTextAttr(self, label, textAttr):
        """ Print a description of the current style, for debugging purposes """
        print label
        if textAttr.GetFont().IsOk():
            print "  Font:               ", textAttr.GetFont().GetFaceName(), textAttr.GetFont().GetPointSize(),
            if textAttr.GetFont().GetWeight() == wx.FONTWEIGHT_BOLD:
                print "Bold ",
            if textAttr.GetFont().GetStyle() == wx.FONTSTYLE_ITALIC:
                print "Italic ",
            if textAttr.GetFont().GetUnderlined():
                print "Underlined ",
            print
        else:
            print "  Invalid Font"
        print "  Colors:             ", textAttr.GetTextColour(), textAttr.GetBackgroundColour()
        print "  Alignment:          ", textAttr.GetAlignment(),
        print "     Indent:             ", textAttr.GetLeftIndent(), textAttr.GetLeftSubIndent(), textAttr.GetRightIndent()
        print "  Line Spacing:       ", textAttr.GetLineSpacing(),
        print "    Paragraph Spacing:  ", textAttr.GetParagraphSpacingBefore(), textAttr.GetParagraphSpacingAfter()
        print "  Tabs:               ", textAttr.GetTabs()
        if self.CompareFormatting(textAttr, self.txtTimeCodeAttr, fullCompare=False):
            print "***  TIME CODE STYLE  ***"
        if self.CompareFormatting(textAttr, self.txtHiddenAttr, fullCompare=False):
            print "***  HIDDEN FONT STYLE  ***"
        if self.CompareFormatting(textAttr, self.txtTimeCodeHRFAttr, fullCompare=False):
            print "***  TIME CODE HRF STYLE  ***"
        print
