# Copyright (C) 2003 - 2014 The Board of Regents of the University of Wisconsin System 

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

"""This module implements the TranscriptEditor class using the wxRichTextCtrl as part of the Transcript Editor
component.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "TranscriptEditor_RTC DEBUG is ON."
SHOWHIDDEN = False
if SHOWHIDDEN:
    print "TranscriptEditor_RTC SHOWHIDDEN is ON."

# Import wxPython
import wx
# Import the Rich Text Control
from RichTextEditCtrl_RTC import RichTextEditCtrl
### If we're on Windows ...
##if 'wxMSW' in wx.PlatformInfo:
##    # import the GDI Report class from RichTextEditCtrl_RTC
##    from RichTextEditCtrl_RTC import GDIReport
# import wxPython's RichTextCtrl
import wx.richtext as richtext
# Import the Format Dialog
import FormatDialog
# Import the Transana Transcript Object
import Transcript
# Import the Transana Drag and Drop infrastructure
import DragAndDropObjects
# Import Transana's Dialogs
import Dialogs
# Import Transana's Episode and Clip Objects 
import Episode, Clip
# Import Transana's Constants
import TransanaConstants
# Import Transana's Global variables
import TransanaGlobal
# import the RTC version of the Transcript User Interface module
import TranscriptionUI_RTC
# import TextReport module
import TextReport
# Import Transana's Miscellaneous functions
import Misc
# Import Python's Regular Expression handler
import re
# Import Python's cPickle module
import cPickle
# Import Python's pickle module
import pickle
# import fast string IO handling
import cStringIO
# Import Python's os module
import os
# Import Python's string module
import string
# Import Python's sys module
import sys
# Import Python's types module
import types

# This character is interpreted as a timecode marker in transcripts
TIMECODE_CHAR = TransanaConstants.TIMECODE_CHAR

# Nate's original REGEXP, "\xA4<[^<]*>", was not working correctly.
# Given the string "this \xA4<1234> is a string > with many characters",
# Nate's REGEXP returned "\xA4<1234> is a string >"
# rather than the desired "\xA4<1234>".
# My REGEXP "\xA4<[\d]*>" appears to do that.
TIMECODE_REGEXP = "%s<[\d]*>" % TIMECODE_CHAR            # "\xA4<[^<]*>"

class TranscriptEditor(RichTextEditCtrl):
    """This class is a word processor for transcribing and editing.  It
    provides only the actual text editing control, without any external GUI
    components to aid editing (such as a toolbar)."""

    def __init__(self, parent, id=-1, stylechange_cb=None, updateSelectionText=False, pos=None, suppressGDIWarning = False):
        """Initialize an TranscriptEditor object."""
        # Initialize with the RTC RichTextEditCtrl object
        RichTextEditCtrl.__init__(self, parent, pos = pos, suppressGDIWarning = suppressGDIWarning)

        # Remember initialization parameters
        self.parent = parent
        self.StyleChanged = stylechange_cb
        self.updateSelectionText = updateSelectionText

        # There are times related to right-click play control when we need to remember the cursor position.
        # Create a variable to store that information, initialized to 0
        self.cursorPosition = 0

        # Initialize CanDrag to False.  If Drag is allowed, it will be enabled later
        self.canDrag = False

        # Define the Regular Expression that can be used to find Time Codes
        self.HIDDEN_REGEXPS = [re.compile(TIMECODE_REGEXP),]
        # Indicate whether Time Code Symbols are shown, default to NOT
        self.codes_vis = 0
        # Indicate whether Time Code Data is shown, default to NOT
        self.timeCodeDataVisible = False
        # Initialize the Transcript Object to be held in this Transcript Editor
        self.TranscriptObj = None
        # For Partial Transcript Loading, we need to track the number of lines loaded in the Text Control
        self.LinesLoaded = 0
        # Initialize the Time Codes array to empty
        self.timecodes = []
        # Initialize the current time code to DOES NOT EXIST
        self.current_timecode = -1

        # Create the AutoSave Timer
        self.autoSaveTimer = wx.Timer()
        # Define the Time Event
        self.autoSaveTimer.Bind(wx.EVT_TIMER, self.OnAutoSave)

        # We should start out in Read Only mode so that we get Word Tracking
        self.set_read_only(True)

        # Initialize Mouse Position to avoid problems later
        self.mousePosition = None
        
        # Remove Drag-and-Drop reference on the mac due to the Quicktime Drag-Drop bug.
        # NOTE:  This bug has been fixed, so the macDragDrop constant is TRUE!
        if TransanaConstants.macDragDrop or (not '__WXMAC__' in wx.PlatformInfo):
            dt = TranscriptEditorDropTarget(self)
            self.SetDropTarget(dt)

# NOTE:  These Bindings have been removed to prevent DUPLICATE CALLS to the methods!!!
        # We need to trap both the EVT_KEY_DOWN and the EVT_CHAR event.
        # EVT_KEY_DOWN processes NON-ASCII keys, such as cursor keys and Ctrl-key combinations.
        # All characters are reported as upper-case.
#        wx.EVT_KEY_DOWN(self, self.OnKeyDown)
        # EVT_CHAR is used to detect normal typing.  Characters are case sensitive here.
#        wx.EVT_CHAR(self, self.OnKey)
        # We need to catch EVT_KEY_UP as well
        wx.EVT_KEY_UP(self, self.OnKeyUp)
        # EVT_LEFT_DOWN is used to detect the left mouse button going down.  Needed (with Left_Up) for unselecting selections.
#        wx.EVT_LEFT_DOWN(self, self.OnLeftDown)
        # EVT_LEFT_UP is used to detect the left click positioning in the Transcript.
#        wx.EVT_LEFT_UP(self, self.OnLeftUp)
        # EVT_MOTION is used to detect mouse motion
#        self.Bind(wx.EVT_MOTION, self.OnMotion)

        # This causes the Transana Transcript Window to override the default
        # RichTextEditCtrl right-click menu.  Transana needs the right-click
        # for play control rather than an editing menu.
        self.Bind(wx.EVT_RIGHT_DOWN, self.OnRightDown)
        self.Bind(wx.EVT_RIGHT_UP, self.OnRightUp)

    # Public methods
    def load_transcript(self, transcript):
        """ Load the given transcript object or RTF file name into the editor. """
        # Remember Partial Transcript Editing status
        tmpPartialTranscriptEdit = TransanaConstants.partialTranscriptEdit
        # Temporarily turn partial transcript editing off
        TransanaConstants.partialTranscriptEdit = False
        
        # Create a popup telling the user about the load (needed for large files)
        loadDlg = Dialogs.PopupDialog(None, _("Loading..."), _("Loading your transcript.\nPlease wait...."))
        # Freeze the control to speed transcript load / RTF Import times
	self.Freeze()
	# Suppress Undo tracking
	self.BeginSuppressUndo()

        # prepare the buffer for the incoming data.
        self.ClearDoc()
        # The control needs to be editable to insert a transcript!
        self.set_read_only(False)

        # The Transcript should already have been saved or cleared by this point.
        # This code should never get activated. It's just here for safety's sake.
        if self.TranscriptObj:
            # Save the transcript
            if TransanaConstants.partialTranscriptEdit:
                self.parent.ControlObject.SaveTranscript(1, continueEditing=False)
            else:
                self.parent.ControlObject.SaveTranscript(1)
            # If you have the Transcript locked, then load something else without leaving
            # Edit Mode, you need to unlock the record!
            if self.TranscriptObj.isLocked:
                self.TranscriptObj.unlock_record()

        # Disable widget while loading transcript
        self.Enable(False)

        # Let's figure out what sort of transcript we have
        # If a STRING is passed in, we probably have a file name, not a Transcript Object.
        if isinstance(transcript, types.StringTypes):
            dataType = 'filename'
        # If we have an empty transcript or a TEXT file ...
        elif transcript.text[:4] == 'txt\n':
            dataType = 'text'
        # If we have a transcript in XML format ...
        elif transcript.text[:5] == '<?xml':
            dataType = 'xml'
        # If we have a transcript in Rich Text Format ...
        elif transcript.text[2:5] == u'rtf':
            dataType = 'rtf'
        # If we are creating a Transcript-less Clip ...
        elif (transcript.text == '') or transcript.text[0:24] == u'<(transcript-less clip)>':
            dataType = 'transcript-less clip'
        # Otherwise, we probably have a Styled Text Ctrl object that's been pickled (Transana 2.42 and earlier)
        else:
            dataType = 'pickle'
        # dataType should only ever be "pickle", "text", "rtf", "xml" or "filename"

        # If we are dealing with a Plain Text document ...
        if dataType == 'text':
            # If the transcript has the TXT indicator ...
            if transcript.text[:4] == 'txt\n':
                # ... we need to remove that!
                text = transcript.text[4:]
            else:
                text = transcript.text

            # Let's scan the file for characters we need to handle.  Start at the beginning of the file.
            # NOTE:  This is a very preliminary implementation.  It only deals with English, and only with ASCII or UTF-8
            #        encoding of time codes (chr(164) or chr(194) + chr(164)).
            pos = 0
            # Keep working until we get to the end of the file.
            while pos < len(text):
                # if we have a non-English character, one the ASCII encoding can't handle ...
                if (ord(text[pos]) > 127):
                    # If we have a Time Code character (chr(164)) ...
                    if (ord(text[pos]) == 164):
                        # ... we can let this pass.  We know how to handle this.
                        pass
                    # In UTF-8 Encoding, the time code is a PAIR, chr(194) + chr(164).  If we see this ...
                    elif (ord(text[pos]) == 194) and (ord(text[pos + 1]) == 164):
                        # ... then let's drop the chr(194) part of things.  At the moment, we're just handling ASCII.
                        text = text[:pos] + text[pos + 1:]
                    # If it's any other non-ASCII character (> 127) ...
                    else:
                        # ... replace it with a question mark for the moment
                        text = text[:pos] + '?' + text[pos + 1:]
                # Increment the position indicator to move on to the next character.
                pos += 1

            # As long as there is text to process ...
            while len(text) > 0:
                # Look for Time Codes
                if text.find(chr(164)) > -1:
                    # Take the chunk of text before the next time code and isolate it
                    chunk = text[:text.find(chr(164))]
                    # remove that chuck of text from the rest of the text
                    # skip the time code character and the opening bracket ("<")
                    text = text[text.find(chr(164)) + 2:]
                    # Grab the text up to the closing bracket (">"), which will be the time code data.
                    timeval = text[:text.find('>')]
                    # Remove the time code data and the closing bracket from the remaining text
                    text = text[text.find('>')+1:]
                    # Add the text chunk to the Transcript
                    self.WriteText(chunk)

                    # Add the time code (with data) to the Transcript 
                    self.insert_timecode(int(timeval))

                # if there are no more time codes in the text ...
                else:
                    # ... add the rest of the text to the Transcript ...
                    self.WriteText(text)
                    # ... and clear the text variable to signal that we're done.
                    text = ''

            # Set the Transcript to the Editor's TranscriptObj
            self.TranscriptObj = transcript
            # Indicate the Transcript hasn't been edited yet.
            self.TranscriptObj.has_changed = 0
	    # If we have an Episode Transcript in TXT Form, save it in FastSave format upon loading
	    # to convert it.  
            self.TranscriptObj.lock_record()

            if TransanaConstants.partialTranscriptEdit:
                self.save_transcript(continueEditing=False)
            else:
                self.save_transcript()
            self.TranscriptObj.unlock_record()

        # If we have an XML document, let's assume it's XML from Transana
        elif dataType == 'xml':
            # Load the XML Data held in the transcript's text field
            self.LoadXMLData(transcript.text)
            # The transcript that was passed in is our Transcript Object
            self.TranscriptObj = transcript
            # Initialize that the transcript has not yet changed.
            self.TranscriptObj.has_changed = 0

        # if we have a wxSTC-style pickled transcript object ...
        elif dataType == 'pickle':
            # import the STC Transcript Editor
            import TranscriptEditor_STC
            # Create an invisible STC-based Transcript Editor
            invisibleSTC = TranscriptEditor_STC.TranscriptEditor(self.parent)
            # HIDE the invisible editor
            invisibleSTC.Show(False)
            # Load the STC-style picked data into the STC Editor
            invisibleSTC.load_transcript(transcript, 'pickle')
            # Convert the STC contents to Rich Text Format
            transcript.text = invisibleSTC.GetRTFBuffer()
            # Destroy the STC-based Editor
            invisibleSTC.Destroy()
        # If we have a transcript-less Clip ...
        elif dataType == 'transcript-less clip':
            # ... then it should have not Transcript Text !!!
            transcript.text = ''
            # The transcript that was passed in is our Transcript Object
            self.TranscriptObj = transcript
            # Initialize that the transcript has not yet changed.
            self.TranscriptObj.has_changed = 0

        # THIS SHOULD NOT BE AN ELIF.
        # If we have a filename, RichTextFormat, or a pickle that has just been converted to RTF ...
        if dataType in ['filename', 'rtf', 'pickle']:
            # looks like this is an RTF text file or an rtf transcript.

            # was the given transcript object simply a filename?
            # NOTE:  This should NEVER occur within Transana!
            if isinstance(transcript, types.StringTypes):
                # Load the document
                self.LoadDocument(transcript)
                # We don't have a Transana Transcript Object in this case.
                self.TranscriptObj = None
            # Is the given transcript a Transcript Object?
            else:
                # Destroy the Load Popup Dialog
                loadDlg.Destroy()
                # Load the Transcript Text using the RTF Data processor
                self.LoadRTFData(transcript.text)
                # The transcript that was passed in is our Transcript Object
                self.TranscriptObj = transcript
                # Initialize that the transcript has not yet changed.
                self.TranscriptObj.has_changed = 0
                # Create a popup telling the user about the load (needed for large files)
                loadDlg = Dialogs.PopupDialog(None, _("Loading..."), _("Loading your transcript.\nPlease wait...."))

            # Hide the Time Code Data, which may be visible for RTF and TXT data
            self.HideTimeCodeData()

	    # If we have a Transcript in RTF Form, save it in FastSave format upon loading
	    # to convert it.  

            try:
                # this was added in to automatically convert an RTF document into
                # the fastsave format.
                self.TranscriptObj.lock_record()
                # If Partial Transcript Editing is enabled ...
                if TransanaConstants.partialTranscriptEdit:
                    # ... save the transcript
                    self.save_transcript(continueEditing=False)
                # If Partial Transcript Editing is NOT enabled ...
                else:
                    # ... save the transcript
                    self.save_transcript()
                # Unlock the transcript
                self.TranscriptObj.unlock_record()
            except:
                # Note the failure in the Error Log
                print "TranscriptEditor_RTC.load_transcript():  SAVE AFTER CONVERSION FAILED."
            
        # Restore Partial Transcript Editing status
        TransanaConstants.partialTranscriptEdit = tmpPartialTranscriptEdit
        
        # Scan transcript for timecodes
        self.load_timecodes()
        # Re-enable widget
        self.Enable(True)
        # Set the Transcript to Read Only initially so that the highlight will scroll as the media plays
        self.set_read_only(True)
        # If the transcript contains time codes ...
        if len(self.timecodes) > 0:
            # ... go to the first time code
            if self.scroll_to_time(self.timecodes[0]):
                # Get the style of the time code
                tmpStyle1 = self.GetStyleAt(self.GetInsertionPoint())
                # if the Time Code's style is HIDDEN ...
                if self.CompareFormatting(tmpStyle1, self.txtHiddenAttr, False):
                    # ... display Time Codes
                    self.show_codes()
        # Time codse should be showing now.
        self.codes_vis = 1

        # Go to the start of the Transcript
        self.GotoPos(0)
        # If we are in the Transcript Dialog, which HAS a toolbar ...
        # (When this is called from the Clip Properties form, there is no tool bar!)
        if isinstance(self.parent, TranscriptionUI_RTC._TranscriptDialog):
            # ... make sure the Toolbar's buttons reflect the current display
            self.parent.toolbar.ToggleTool(self.parent.toolbar.CMD_SHOWHIDE_ID, True)

        # Check Formatting to set initial Default and Basic Style info
	self.CheckFormatting()

        # Now that the transcript is loaded / imported, we can thaw the control
        self.EndSuppressUndo()
	self.Thaw()
        # Mark the Edit Control as unmodified.
	self.DiscardEdits()

        # Implement Minimum Transcript Width by setting size hints for the TranscriptionUI dialog
        self.parent.SetSizeHints(minH = 0, minW = self.TranscriptObj.minTranscriptWidth)
        # Destroy the Load Popup Dialog
        loadDlg.Destroy()

        # if Partial Transcript Editing is enabled ...
        if TransanaConstants.partialTranscriptEdit:
            # ... we need to track the lines that are loaded
            self.LinesLoaded = self.TranscriptObj.paragraphs

    def UpdateCurrentContents(self, action):
        """ This method maintains a LIMITED load of data in the editor control, rather than having all
            the data present all the time. In wxPython 2.9.4.0 and 3.0.0.0, the wxRichTextCtrl becomes
            VERY slow during editing for very large documents.  (eg. a 7000 line document can take 4
            seconds per key press near the beginning of the document!) """

        # Set the number of lines that should be included in a transcript segment loaded into the editor
        numberOfLinesInControl = 200
        # If no Transcript Object is defined ...
        if self.TranscriptObj == None:
            # ... we can skip this!
            return

        # If we're entering edit mode, we need to limit the amount of text in the control ...
        if action == 'EnterEditMode':

            # ... and if the transcript has over numberOfLinesInControl lines long AFTER the current window ...
            if self.NumberOfLines - self.PositionToXY(self.HitTestPos((3, self.GetRect()[3] - 10))[1])[1] > numberOfLinesInControl:
                # ... determine the number of lines to load into the control 
                linesToLoad = self.PositionToXY(self.HitTestPos((3, self.GetRect()[3] - 10))[1])[1]
                linesToLoad = linesToLoad - (linesToLoad % numberOfLinesInControl) + numberOfLinesInControl
                # If we should load fewer than ALL the lines ...
                if linesToLoad < self.TranscriptObj.paragraphs:
                    # Create a temporary popup dialog ...
                    loadDlg = Dialogs.PopupDialog(self, _("Loading %d lines") % linesToLoad, _("Loading your transcript.\nPlease wait...."))

                    # Initialize text
                    text = ''
                    # Add the appropriate number of lines (i.e. paragraphs)
                    for x in range(self.TranscriptObj.paragraphPointers[linesToLoad]):
                        text += self.TranscriptObj.lines[x] + '\n'
                    # Add closing XML to made our text a LEGAL XML document
                    text += '  </paragraphlayout>\n'
                    text += '</richtext>'
                        
                    # Load the XML Data held in the transcript's text field
                    self.LoadXMLData(text, clearDoc=False)

                    # Delete the popup dialog.
                    loadDlg.Destroy()

                    # Update the indicator for the number of lines loaded
                    self.LinesLoaded = linesToLoad
                # If we shouls load ALL the lines ...
                else:
                    # ... update the indicator for the number of lines loaded to the total number of paragraphs
                    self.LinesLoaded = self.TranscriptObj.paragraphs
            # Otherwise ...
            else:
                # ... update the indicator for the number of lines loaded to the total number of paragraphs
                self.LinesLoaded = self.TranscriptObj.paragraphs

        # if we're leaving Edit mode, we need to recover the text that had been left out of the control ...
        elif action == 'LeaveEditMode':
            # if there are lines beyond what is currently in the Editor control ...
            if self.LinesLoaded > 0 and self.LinesLoaded != self.TranscriptObj.paragraphs:
                # If the transcript has been changed ...
                if self.IsModified():
                    # ... get the text currently in the control ...
                    currenttext = self.GetFormattedSelection('XML')
                    # ... and break it into lines
                    currentlines = currenttext.split('\n')
                    # Delete the last TWO lines from the loaded text, as they close off the XML too early
                    del(currentlines[-1])
                    del(currentlines[-1])
                    # See if the (formerly) 3rd to last line closes a ParagraphLayout XML tag set
                    if currentlines[-1].strip() == '</paragraphlayout>':
                        # ... and if so, delete that too!
                        del(currentlines[-1])

                    # For all of the original Transcript that falls AFTER what we have loaded in the Text Control ...
                    for x in range(self.TranscriptObj.paragraphPointers[self.LinesLoaded], len(self.TranscriptObj.lines)):
                        # ... add these lines to what we got from the Text Control.
                        currentlines.append(self.TranscriptObj.lines[x])

                    # re-initialize CurrentText
                    currenttext = ""
                    # Concatenate all the current LINES into the current TEXT
                    for x in range(len(currentlines)):
                        currenttext += currentlines[x] + '\n'
                    # Now load the cumulated text into the Text Control
                    self.LoadXMLData(currenttext, clearDoc=False)
                    # Note that the text HAS changed in the Text Control
                    self.MarkDirty()
                # If the transcript has NOT been changed ...
                else:
                    # ... restore the original transcript's text to the Text Control
                    self.LoadXMLData(self.TranscriptObj.text, clearDoc = False)

    def HideTimeCodeData(self):
            """ Hide the Time Code Data, which should NEVER be visible to users """
            # Let's look for time codes and hide the data
            
            # NOTE:  This will ONLY apply to RTF transcripts exported by Transana 2.42 and earlier.  By definition,
            #        these transcripts CANNOT have images, and therefore the offset values returned by FindText will
            #        be adequate.  Transcript that might have time codes AND images will have the time codes already
            #        formatted correctly, so if some time codes get skipped below, it doesn't matter.
            
            # Start at the beginning of the text
            pos = 0
            # Run to the end of the text
            endPos = self.GetTextLength()
            # Find the first time code character
            nextTC = self.FindText(pos, endPos, TIMECODE_CHAR)
            # As long as there are additional time codes to find ...
            while nextTC > -1:
                # Find the END of the time code we just found
                endTC = self.FindText(nextTC, endPos, '>')
                # Select the time code character itself
                self.SetSelection(nextTC + 1, endTC + 1)

                # NOTE:  That doesn't work quite right for transcripts exported to RTF from Transana 2.50.  We
                #        need an additional correction here!  Otherwise, every other time-coded segment is hidden!
                #
                # Get the TEXT of the current selection
                tmp = self.GetStringSelection()
                # If there is a ">" character BEFORE THE END of the text ...
                if len(tmp) > tmp.find('>') + 1:
                    # ... then we need to CORRECT the END TIME CODE position marker.
                    endTC = nextTC + tmp.find('>') + 1

                # Format the Time Code using the Time Code style
                self.SetStyle(richtext.RichTextRange(nextTC, nextTC + 1), self.txtTimeCodeAttr)
                # Format the Time Code data using the Hidden style
                self.SetStyle(richtext.RichTextRange(nextTC + 1, endTC + 1), self.txtHiddenAttr)
                # Find the next time code, if there is one
                nextTC = self.FindText(endTC, endPos, TIMECODE_CHAR)

    def load_timecodes(self):
        """Scan the document for timecodes and add to internal list."""
        # Clear the existing time codes list
        self.timecodes = []
        # Get the text to scan
        txt = self.GetText()
        # Define the string to search for
        findstr = TIMECODE_CHAR + "<"
        # Locate the time code using string.find()
        i = txt.find(findstr, 0)
        # As long as there are more time codes to find ...
        while i >= 0:
            # ... find the END of the time code data, i.e. the next ">" character
            endi = txt.find(">", i)
            # Extract the Time Code Data
            timestr = txt[i+2:endi]
            # Trap exceptions
            try:
                # Conver the time code data to an integer and add it to the TimeCodes list
                self.timecodes.append(int(timestr))
            # If an exception arises (because of inability to convert the time code) ...
            except:
                # ... then just ignore that time code.  It's probably defective.
                pass
            # Look for the next time code.  Result will be -1 if NOT FOUND
            i = txt.find(findstr, i+1)

    def save_transcript(self, continueEditing=True):
        """ Save the transcript to the database.
            continueEditing is used for Partial Transcript Editing only. """
        # Create a popup telling the user about the save (needed for large files)
        self.saveDlg = Dialogs.PopupDialog(None, _("Saving..."), _("Saving your transcript.\nPlease wait...."))
        # Let's try to remember the cursor position
        self.SaveCursor()
        # If Partial Transcript editing is enabled ...
        if TransanaConstants.partialTranscriptEdit:
            # If we have only part of the transcript in the editor, we need to restore the full transcript
            self.UpdateCurrentContents('LeaveEditMode')
        # We can't save with Time Codes showing!  Remember the initial status, and hide them
        # if they are showing.
        initCodesVis = self.codes_vis
        if not initCodesVis:
            self.show_codes()
        # We shouldn't save with Time Code Values showing!  Remember the initial status for later.
        initTimeCodeValueStatus = self.timeCodeDataVisible
        # If Time Code Values are showing ...
        if self.timeCodeDataVisible:
            # ... then hide them for now.
            self.changeTimeCodeValueStatus(False)
        # If we have a defined Transcript Object ...
        if self.TranscriptObj:
            # Note whether the transcript has changed
            self.TranscriptObj.has_changed = self.modified()
            # Get the transcript data in XML format
            self.TranscriptObj.text = self.GetFormattedSelection('XML')
            # Write it to the database
            self.TranscriptObj.db_save()
        # If time codes were showing, show them again.
        if not initCodesVis:
            self.hide_codes()
        # If Time Code Values were showing, show them again.
        if initTimeCodeValueStatus:
            self.changeTimeCodeValueStatus(True)
        # Let's try restoring the Cursor Position when all is said and done.
        self.RestoreCursor()
        # Mark the Edit Control as unmodified.
	self.DiscardEdits()
        # Destroy the Save Popup Dialog
        self.saveDlg.Destroy()
        # If Partial Transcript editing is enabled ...
        if TransanaConstants.partialTranscriptEdit and continueEditing:
            # If we have only part of the transcript in the editor, we need to restore the partial transcript state following save
            self.UpdateCurrentContents('EnterEditMode')

    def export_transcript(self, fname):
        """Export the transcript to an RTF file."""
        # If Partial Transcript editing is enabled ...
        if TransanaConstants.partialTranscriptEdit:
            # If we have only part of the transcript in the editor, we need to restore the full transcript
            self.UpdateCurrentContents('LeaveEditMode')
            self.Refresh()

        # See if there are any time codes in the text.  If not ...
        if self.timecodes == []:
            # ... then we can ignore the whole issue of time-code stripping
            result = wx.ID_YES
        # If there ARE time codes ...
        else:
            # We want to ask the user whether we should include time codes or not.  Create the prompt
            prompt = unicode(_("Do you want to include Transana Time Codes (and their hidden data) in the file?\n(This preserves time codes when you re-import the transcript into Transana.)"), "utf8")
            # Create a dialog box for the question
            dlg = Dialogs.QuestionDialog(self.parent, prompt)
            # Display the dialog box and get the user response
            result = dlg.LocalShowModal()
            # Destroy the dialog box
            dlg.Destroy()

        # If the user does NOT want Time Codes ...
        if result == wx.ID_NO:
            # Remember the Edited Status of the RTC
            isModified = self.IsModified()

            # Get the contents of the RTC buffer
            originalText = self.GetFormattedSelection('XML')
            # Remove the Time Codes
            strippedText = self.StripTimeCodes(originalText)
            # Now put the altered contents of the buffer back into the control!
            # (The RichTextXMLHandler automatically clears the RTC.)
            try:
                # Create an IO String of the stripped text
                stream = cStringIO.StringIO(strippedText)
                # Create an XML Handler
                handler = richtext.RichTextXMLHandler()
                # Load the XML text via the XML Handler.
                # Note that for XML, the RTC BUFFER is passed.
                handler.LoadStream(self.GetBuffer(), stream)
            # exception handling
            except:
                import traceback
                print "XML Handler Load failed"
                print
                print sys.exc_info()[0], sys.exc_info()[1]
                print traceback.print_exc()
                print
                pass

        # If saving an RTF file ...
        if fname[-4:].lower() == '.rtf':
            # ... save the document in Rich Text Format
            self.SaveRTFDocument(fname)
        # If saving an XML file ...
        elif fname[-4:].lower() == '.xml':
            # ... save the document in XML format
            self.SaveXMLDocument(fname)

        # If the user did NOT want Time Codes ...
        if result == wx.ID_NO:
            # ... we need to put the original contents of the buffer back into the control!
            try:
                # Create an IO String of the stripped text
                stream = cStringIO.StringIO(originalText)
                # Create an XML Handler
                handler = richtext.RichTextXMLHandler()
                # Load the XML text via the XML Handler.
                # Note that for XML, the RTC BUFFER is passed.
                handler.LoadStream(self.GetBuffer(), stream)
            # exception handling
            except:
                import traceback
                print "XML Handler Load failed"
                print
                print sys.exc_info()[0], sys.exc_info()[1]
                print traceback.print_exc()
                print
                pass

            # If the RTC was modified BEFORE this export ...
            if isModified:
                # ... then mark it as modified
                self.MarkDirty()
            # If the RTC was NOT modified before this export ...
            else:
                # ... then mark it as clean.  (yeah, one of these is probably not needed.)
                self.DiscardEdits()

        # If Partial Transcript editing is enabled ...
        if TransanaConstants.partialTranscriptEdit:
            # If we have only part of the transcript in the editor, we need to restore the partial transcript state following save
            self.UpdateCurrentContents('EnterEditMode')

    def set_font(self, font_face, font_size, font_fg=0x000000, font_bg=0xffffff):
        """Change the current font or the font for the selected text."""
        self.SetFont(font_face, font_size, font_fg, font_bg)
    
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
        """ Report the current setting for BOLD """
        return self.GetBold()
    
    def set_italic(self, enable=-1):
        """Set italic state for current font or for the selected text."""
        if self.get_read_only():
            return
        if enable == -1:
            enable = not self.GetItalic()
        self.SetItalic(enable)

    def get_italic(self):
        """ Report the current setting for ITALICS """
        return self.GetItalic()

    def set_underline(self, enable=-1):
        """Set underline state for current font or for the selected text."""
        if self.get_read_only():
            return
        if enable == -1:
            enable = not self.GetUnderline()
        self.SetUnderline(enable)

    def get_underline(self):
        """ Report the current setting for UNDERLINE """
        return self.GetUnderline()
        
    def show_codes(self, showPopup=False):
        """Make encoded text in document visible."""
        if showPopup:
            # Create a popup telling the user about the change (needed for large files)
            tmpDlg = Dialogs.PopupDialog(None, _("Showing Time Codes..."), _("Showing time codes.\nPlease wait...."))
        self.changeTimeCodeHiddenStatus(False)
        self.codes_vis = 1
        if showPopup:
            # Destroy the Popup Dialog
            tmpDlg.Destroy()
        
    def hide_codes(self, showPopup=False):
        """Make encoded text in document visible."""
        if showPopup:
            # Create a popup telling the user about the change (needed for large files)
            tmpDlg = Dialogs.PopupDialog(None, _("Hiding Time Codes..."), _("Hiding time codes.\nPlease wait...."))
        self.changeTimeCodeHiddenStatus(True)
        self.codes_vis = 0
        if showPopup:
            # Destroy the Popup Dialog
            tmpDlg.Destroy()

    def show_timecodevalues(self, visible):
        """ Make Time Code value in Human Readable form visible or hidden """
        # Create a popup telling the user about the change (needed for large files)
        if not visible:
            tmpDlg = Dialogs.PopupDialog(None, _("Hiding Time Codes..."), _("Hiding time code values.\nPlease wait...."))
        else:
            tmpDlg = Dialogs.PopupDialog(None, _("Showing Time Codes..."), _("Showing time code values.\nPlease wait...."))
        # Just passing through.
        self.changeTimeCodeValueStatus(visible)
        # Destroy the Popup Dialog
        tmpDlg.Destroy()

    def codes_visible(self):
        """Return 1 if encoded text is visible."""
        return self.codes_vis

    def changeTimeCodeHiddenStatus(self, hiddenVal):
        """ Changes the Time Code marks (but not the time codes themselves) between visible and invisble styles. """
        # Let's also remember if the transcript has already been modified.  This value WILL get changed, but maybe it shouldn't be.
        initModified = self.modified()
        # Let's try to remember the cursor position.  (self.SaveCursor() doesn't work here!)
        (savedPosition, savedSelection) = (self.GetCurrentPos(), self.GetSelection())
        # Move the cursor to the beginning of the document
        self.GotoPos(0)

        # Let's find each time code mark and update it with the new style.  
        for tc in self.timecodes:
            # Find the Timecode.  scroll_to_time() adjusts the time code by 2 ms, so we have to compensate for that here!
            if self.scroll_to_time(tc - 2):
                # Note the Cursor's Current Position
                curpos = self.GetCurrentPos() + 1

                # The time code in position 0 of the document doesn't get hidden correctly!  This adjusts for that!
                if curpos < 1:
                    curpos = 1

                # Get the range for the Time Code character itself.  It starts the character BEFORE the insertion point.
                r = richtext.RichTextRange(curpos - 1, curpos)
                # If we're hiding time codes ...
                if hiddenVal:
                    # ... set its style to Hidden
                    self.SetStyle(r, self.txtHiddenAttr)
                # If we're displaying time codes ...
                else:
                    # ... set its style to Time Code
                    self.SetStyle(r, self.txtTimeCodeAttr)

        # Restore the Cursor Position when all is said and done.  (self.RestoreCursor() doesn't work!)
        # If there's no saved Selection ...
        if savedSelection == (-2, -2):
            # ... make a small selection based on the saved position.  If it's not the last character in the document ...
            if savedPosition < self.GetLastPosition():
                # ... select the next character
                self.SetSelection(savedPosition, savedPosition+1)
            # If it IS the last character and the transcript has at least one character ...
            elif savedPosition > 0:
                # ... select the previous character
                self.SetSelection(savedPosition - 1, savedPosition)
            # Show the current selection
            self.ShowCurrentSelection()
            # Now clear the selection we just made
            self.SetCurrentPos(savedPosition)
        # if there IS a saved selection ...
        else:
            # ... select what used to be selected ...
            self.SetSelection(savedSelection[0], savedSelection[1])
            # ... and show the current selection
            self.ShowCurrentSelection()
        # Start Exception Handling
        try:
            # Update the Transcript Control
            self.Update()
        # If there's a PyAssertionError ...
        except wx._core.PyAssertionError, x:
            # ... we can safely ignore it!
            pass
        
        # If we did not think the document was modified before we showed the time code data ...
        if not initModified:
            # ... then mark the data as unchanged.
            self.DiscardEdits()

    def changeTimeCodeValueStatus(self, visible):
        """ Change visibility of the Time Code Values """
        # Set the Wait cursor
        self.parent.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
        # We can change this even if the Transcript is read-only, but we need to remember the current
        # state so we can return the transcript to read-only if needed.

        # The whole SaveCursor() / RestoreCursor() thing doesn't work well here, because the character numbers just change too
        # much.  Let's maintain transcript position by time code instead.
        (tcBefore, tcAfter) = self.get_selected_time_range()
        
        initReadOnly = self.get_read_only()
        # Let's also remember if the transcript has already been modified.  This value WILL get changed, but maybe it shouldn't be.
        initModified = self.modified()

        # Let's show all the hidden text of the time codes.  This doesn't work without it!
        if not self.codes_vis:
            # Remember that time codes were hidden ...
            wereHidden = True
            # ... and show them
            self.show_all_hidden()
        else:
            # Remember that time codes were showing
            wereHidden = False
        # Put the transcript into Edit mode so we can change it.
        self.set_read_only(False)

        # Let's iterate through every pre-defined regular expression about time codes.  (I think there's only one.  I'm not sure why Nate did it this way.)
        for tcs in self.HIDDEN_REGEXPS:
            # Get a list(?) of all the time code sequences in the text
            tcSequences = tcs.findall(self.GetText())
            # Initialize the Time Code End position to zero.  
            tcEndPos = 0
            # Now iterate through each Time Code in the RegEx list
            for TC in tcSequences:

                # Find the next Time Code in the RTF control, starting at the end point of the previous time code for efficiency's sake.
                tcStartPos = self.GetValue().find(TC, tcEndPos, self.GetLength())  # self.FindText(tcEndPos, self.GetLength(), TC)
                # Remember the end point of the current time code, used to start the next search.
                tcEndPos = self.GetValue().find('>', tcStartPos, self.GetLength())  # self.FindText(tcStartPos, self.GetLength(), '>') + 1
                tcEndPosAdjusted = self.FindText(tcStartPos, self.GetLength(), TC) + len(TC)
                # Move the cursor to the end of the time code's hidden data
                self.GotoPos(tcEndPosAdjusted)

                # Build the text of the time value.  Take parentheses, and add the conversion of the time code data, which is extracted from
                # the Time Code from the RegEx.
                text = '(' + Misc.time_in_ms_to_str(int(TC[2:-1])) + ')'
                # Note the length of the time code text
                lenText = len(text)

#                print "TrancriptEditor_RTC.changeTimeCodeValueStatus():", tcCounter, len(tcSequences)

                # If we're going to SHOW the time code data ...
                if visible:
                    # Insert the text
                    self.WriteText(text)

                    self.SetStyle(richtext.RichTextRange(tcEndPosAdjusted, tcEndPosAdjusted + lenText), self.txtTimeCodeHRFAttr)
                # If we're gong to HIDE the time code data ...
                else:
                    # Let's look at the end of the time code for the opening paragraph character.  This probably signals that the user hasn't
                    # messed with the text, which they could do.  If they mess with it, they're stuck with it!
                    if self.GetCharAt(tcEndPosAdjusted) == ord('('):

                        hrtcStartPos = tcEndPosAdjusted
                        hrtcEndPos = tcEndPosAdjusted + lenText  # self.GetValue().find(')', hrtcStartPos, self.GetLength())
                        # Select the character following the time code data ...
                        self.SetSelection(hrtcStartPos, hrtcEndPos)
                        # ... and get rid of it!
                        self.DeleteSelection()

        # Change the Time Code Data Visible flag to indicate the new state
        self.timeCodeDataVisible = visible

        # We better hide all the hidden text for the time codes again, if they were hidden.
        if wereHidden:
            self.hide_all_hidden()

        # If we were in read-only mode ...
        if initReadOnly:
            # ... return to read-only mode
            self.set_read_only(True)

        # If we did not think the document was modified before we showed the time code data ...
        if not initModified:
            # ... then mark the data as unchanged.
            self.DiscardEdits()
        # Restore the normal cursor
        self.parent.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

        # The cursor position is lost because of the change of the text.  Go to the start of the document.
        self.SetInsertionPoint(0)
        # Call CheckFormatting to update the Format / Font values
        self.CheckFormatting()
        # Scroll to the starting time code position to make the proper segment of the transcript visible
        self.scroll_to_time(tcBefore)
        # Start Exception Handling
        try:
            # Update the Transcript Control
            self.Update()
        # If there's a PyAssertionError ...
        except wx._core.PyAssertionError, x:
            # ... we can safely ignore it!
            pass

    def show_all_hidden(self):
        """Make encoded text in document visible."""
        self.codes_vis = 1

    def hide_all_hidden(self):
        """Make encoded text in document visible."""
        self.codes_vis = 0

    def find_text(self, txt, direction, flags=0):
        """Find text in document."""
        # If we currently have the search text as the selection AND we're looking for the NEXT instance ...
        if (self.GetStringSelection().upper() == txt.upper()) and (direction == 'next'):
            # ... then the current insertion point should be one space beyond the current insertion point
            ip = self.GetInsertionPoint() + 1
        # Otherwise ...
        else:
            # ... not the current insertion point
            ip = self.GetInsertionPoint()
        # Search doesn't work right on transcripts after the first one without this code.
        # If there is no defined Insertion Point ...
        if ip < 0:
            # ... select the start of the document
            ip = 0
            
        # RTC doesn't provide a FIND function.  We'll use Python's string find because it's FAST.
        #
        # But once Images are inserted in the RTC, the Python FIND function
        # produces incorrect results, off by one space per image inserted above the found text.
        # We have to compensate for that.

        # If we're looing for the NEXT instance of the search text ...
        if direction == 'next':
            # ... get the unformatted text from the control starting with the current insertion point,
            # and converted to all lower case characters
            ctrlText = self.GetValue()[ip:].lower()
            # Use Python's string.FIND to find the next instance of the search text
            newPos = ctrlText.find(txt.lower())
            # If a next instance is found ...
            if (newPos > -1):
                # Move the control's selection to the next instance
                self.SetSelection(ip + newPos, ip + newPos + len(txt))

                # Compensate for incorrect results due to IMAGES in the Rich Text Ctrl.
                
                # If the selected text doesn't match the FIND text ...  (this will be repeated once PER IMAGE ABOVE)
                while (newPos < self.GetLastPosition() - len(txt)) and (self.GetStringSelection().upper() != txt.upper()):
                    # ... move the position one space to the right
                    newPos += 1
                    # ... and re-do the selection.  
                    self.SetSelection(ip + newPos, ip + newPos + len(txt))

        # If we're looking for the PREVIOUS instance of the search text ...
        elif direction == 'back':
            # ... get the unformatted text from the control up to the current insertion point,
            # and converted to all lower case characters
            ctrlText = self.GetValue()[:ip - 1].lower()
            # Use Python's string.RFIND to find the PREVIOUS instance of the search text
            newPos = ctrlText.rfind(txt.lower())
            # Unfortunately, because of the RTC location problem caused by images, this rfind might not have worked.
            # It may just have returned the SAME instance of the search text.  (The more images above the current 
            # position and the shorter the search text, the more likely!)  So let's find the NEXT earlier instance
            # of the search text too, just in case.

            # So get the unformatted text from the control up to the first found instance,
            # and converted to all lower case characters
            ctrlText = self.GetValue()[:newPos - 1].lower()
            # Use Python's string.RFIND to find the PREVIOUS instance of the search text
            earlierPos = ctrlText.rfind(txt.lower())
            # If a first earlier instance of the search text was found ...
            if newPos > -1:
                # ... move the control's selection to the first earlier instance
                self.SetSelection(newPos, newPos + len(txt))

                # Compensate for incorrect results due to IMAGES in the Rich Text Ctrl.
            
                # If the selected text doesn't match the FIND text ...  (this will be repeated once PER IMAGE ABOVE)
                while (newPos < self.GetLastPosition()) and (self.GetStringSelection().upper() != txt.upper()):
                    # ... move the position one space to the right
                    newPos += 1
                    # ... and re-do the selection.  
                    self.SetSelection(newPos, newPos + len(txt))
            # If the first RFIND only found the instance where we already are, AND if there IS a second earler
            # instance of the search text ...
            if (newPos + len(txt) == ip) and (earlierPos > -1):
                # ... move the control's selection to the first earlier instance
                self.SetSelection(earlierPos, earlierPos + len(txt))

                # Compensate for incorrect results due to IMAGES in the Rich Text Ctrl.
            
                # If the selected text doesn't match the FIND text ...  (this will be repeated once PER IMAGE ABOVE)
                while (earlierPos < newPos) and (self.GetStringSelection().upper() != txt.upper()):
                    # ... move the position one space to the right
                    earlierPos += 1
                    # ... and re-do the selection.  
                    self.SetSelection(earlierPos, earlierPos + len(txt))
        else:
            print "Unknown search direction:", direction

        # Determine the Range of the current selection
        selRange = self.GetSelectionRange()
        # If there IS a selection ...
        if selRange != (-1, -1):
            # Create a temporary RichTextAttr object to hold the style
            selStyle = richtext.RichTextAttr()
            # Get the Style associated with the selection
            self.GetStyleForRange(selRange, selStyle)
            # if the selection's style is HIDDEN ...
            if self.CompareFormatting(selStyle, self.txtHiddenAttr, False):
                # ... then call the find_text method recursively to find a NON-HIDDEN version of the text
                self.find_text(txt, direction, flags)

        # Make sure the current selection is showing on screen
        self.ShowCurrentSelection()

    def insert_timecode(self, time_ms=-1):
        """Insert a timecode in the current cursor position of the
        Transcript.  The parameter time_ms is optional and will default
        to the current Video Position if not used."""
        if self.get_read_only():
            # Don't do it in read-only mode
            return
        
        if TransanaConstants.demoVersion and (self.GetLength() > 10000):
            prompt = _("The Transana Demonstration limits the size of Transcripts.\nYou have reached the limit and cannot edit this transcript further.")
            tempDlg = Dialogs.InfoDialog(self, prompt)
            tempDlg.ShowModal()
            tempDlg.Destroy()
            return

        # Get the time codes around the current selection        
        (prevTimeCode, nextTimeCode) = self.get_selected_time_range()
        # If a time value is passed in ...
        if time_ms >= 0:
            # ... note that time value
            timepos = time_ms
        # If no time value is passed in ...
        else:
            # ... use the current media position
            timepos = self.parent.ControlObject.GetVideoPosition()

        # Check that the proposed time code falls in the range for the current selection.
        # The second line here allows you to insert a timecode at 0.0 if there isn't already one there.
        if (len(self.timecodes) == 0) or (prevTimeCode < timepos) and ((timepos < nextTimeCode) or (nextTimeCode == -1)) \
           or ((timepos == 0) and (prevTimeCode == 0.0) and (self.timecodes[0] > 0)):
            
            # Insert the Time Code
            self.InsertTimeCode(timepos)
            # If time code data is visible ...
            if self.timeCodeDataVisible:
                # ... work out the time code value ...
                tcText = '(' + Misc.time_in_ms_to_str(int(timepos)) + ')'
                # If we're using RTC ...
                if TransanaConstants.USESRTC:
                    # ... shift one space to the left.
                    # This is because the time code data ends with a space, but time code data should be placed
                    # BEFORE that space.
                    self.SetInsertionPoint(self.GetInsertionPoint() - 1)
                    # Note the current insertion point
                    ip = self.GetInsertionPoint()
                    # ... and insert it into the text
                    self.WriteText(tcText)
                    # Format the Time Code's Human Readable value
                    self.SetStyle(richtext.RichTextRange(ip, ip + len(tcText)), self.txtTimeCodeHRFAttr)
                    # Now shift one space to the right to get past that space that ends the time code data
                    self.SetInsertionPoint(ip + len(tcText) + 1)
                # If we're not using RTC, we should NEVER be here!!
                else:
                    # ... and insert it into the text
                    self.InsertStyledText(tcText, len(tcText))
            # Update the RTC
            self.Refresh()
            # Update the 'timecodes' list, putting it in the right spot
            i = 0
            if len(self.timecodes) > 0:
                # the index variable (i) must be less than the number of elements in
                # self.timecodes to avoid an index error when inserting a selection timecode
                # at the end of the Transcript.
                while (i < len(self.timecodes) and (self.timecodes[i] < timepos)):
                    i = i + 1
            # Add the time code value to the TimeCodes array
            self.timecodes.insert(i, timepos)
        # If the proposed time code is out of sequence ...
        else:
            # ... build an error message.
            msg = _('Time Code Sequence error.\nYou are trying to insert a Time Code at %s\nbetween time codes at %s and %s.')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                msg = unicode(msg, 'utf8')
            # If there is no "later" time code ...
            if nextTimeCode == -1:
                # ... and if we're in an Episode ...
                if type(self.parent.ControlObject.currentObj) == type(Episode.Episode()):
                    # ... use the end of the media file
                    nextTimeCode = self.parent.ControlObject.currentObj.tape_length
                # ... and if we're in a Clip ...
                else:
                    # ... use the end of the clip
                    nextTimeCode = self.parent.ControlObject.currentObj.clip_stop
            # Display the error message
            errordlg = Dialogs.ErrorDialog(self, msg % (Misc.time_in_ms_to_str(timepos), Misc.time_in_ms_to_str(prevTimeCode), Misc.time_in_ms_to_str(nextTimeCode)))
            errordlg.ShowModal()
            errordlg.Destroy()

    def insert_timed_pause(self, start_ms, end_ms):
        """ Insert a timed pause """
        # If we're in read only mode ...
        if self.get_read_only():
            # ... don't do it
            return
        # Get the Time Range for the current text position
        (prevTimeCode, nextTimeCode) = self.get_selected_time_range()
        # Issue 231 -- Selection Insert fails if it is after the last known time code.
        # If the last time code is undefined, use the media length for an Episode or the
        # Clip Stop Point for a Clip.  nextTimeCode = -1 signals this!
        if nextTimeCode == -1:
            # If we're in an Episode ...
            if type(self.parent.ControlObject.currentObj) == type(Episode.Episode()):
                # ... then use the Media File Length
                nextTimeCode = self.parent.ControlObject.currentObj.tape_length
            # If we're in a Clip ...
            else:
                # ... use the Clip End position
                nextTimeCode = self.parent.ControlObject.currentObj.clip_stop
        # If the time code before the cursor is later than the Timed Pause's beginning ...
        if prevTimeCode > start_ms:
            # ... display an error message.
            msg = _('Time Code Sequence error.\nYou are trying to insert a Time Code at %s\nbetween time codes at %s and %s.')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                msg = unicode(msg, 'utf8')
            errordlg = Dialogs.ErrorDialog(self, msg % (Misc.time_in_ms_to_str(start_ms), Misc.time_in_ms_to_str(prevTimeCode), Misc.time_in_ms_to_str(nextTimeCode)))
            errordlg.ShowModal()
            errordlg.Destroy()
        # If the time code after the cursor is earlier than the Timed Pause's end ...
        elif nextTimeCode < end_ms:
            # ... display an error message.
            msg = _('Time Code Sequence error.\nYou are trying to insert a Time Code at %s\nbetween time codes at %s and %s.')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                msg = unicode(msg, 'utf8')
            errordlg = Dialogs.ErrorDialog(self, msg % (Misc.time_in_ms_to_str(end_ms), Misc.time_in_ms_to_str(prevTimeCode), Misc.time_in_ms_to_str(nextTimeCode)))
            errordlg.ShowModal()
            errordlg.Destroy()
        # If there are no errors ...
        else:

            # If you make your Visualization Selection before you go into Edit Mode, problems occur.  The following code solves
            # three problems that I identified.

            # If there is a SELECTION ...
            if self.GetSelection() != (-2, -2):
                # ... place the cursor at the end of the selection
                self.SetInsertionPoint(self.GetSelection()[1])

            # If there is no Insertion Point ...
            if self.GetInsertionPoint() == -2:
                # ... set the insertion point to 0  (This happens at the start of a blank file!)
                self.SetInsertionPoint(0)

            # If we are at the END of the file ...
            if self.GetInsertionPoint() == self.GetLastPosition():
                # ... add a space ...
                self.WriteText(' ')
                # ... and move to before that space.
                self.SetInsertionPoint(self.GetLastPosition() - 1)

            # Determine (and round) the time space
            time_span_secs = (end_ms - start_ms) / 1000.0
            time_span_secs = round(time_span_secs, 1)
            # Insert the starting time code
            self.insert_timecode(start_ms)
            # Format and insert the Time Span in Jeffersonian format
            timespan = "(%.1f)" % time_span_secs
            self.WriteText(timespan)
            # Insert the ending time code
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
        ms = ms + 2

        # If no timecodes exist, or we're in Edit mode, ignore this!
        if (len(self.timecodes) == 0):  # or not self.GetReadOnly():
            return False

        # Temporarily halt screen updates
        self.Freeze()
        
        # Find the timecodes that are on either side of what we want
        # Initialize "Before" to the start of the file
        tcBefore = -1
        # Initializae "After" to the end of the file (-1)
        tcAfter = -1

        # Iterate through the time code list
        for timecode in self.timecodes:

            # If the entry in the time code list is less than the current media position ...
            if timecode < ms:
                # ... update the Before value
                tcBefore = timecode
            # If the entry in the time code list is greater than the current media position ...
            else:
                # ... then we can stop looking for our Before position and we should capture THIS as our After position
                tcAfter = timecode
                # We can stop iterating once we reach this point.
                break

        # If the current position is before the first time code ...
        if tcBefore == -1:
            # ... start at the first position in the RichTextCtrl
            start = 0
        # Otherwise ...
        else:
            # ... let's get the character position of the Before time code
            start = self.FindText(0, self.GetTextLength(), "%s<%d>" % (TIMECODE_CHAR, tcBefore))

        # If the current position is after the last time code ...
        if tcAfter == -1:
            # ... end at the last position in the RichTextCtrl
            end = self.GetTextLength()
        # Otherwise ...
        else:
            # ... let's get the character position of the After time code
            end = self.FindText(0, self.GetTextLength(), "%s<%d>" % (TIMECODE_CHAR, tcAfter))

        # Let's get the current selection position
        pos = self.GetSelection()
        # If there IS A PREVIOUS TIME CODE ...
        if tcBefore > -1:
            # ... offset the selection so the Time Code Data is NOT included
            offset = len("%s<%d>" % (TIMECODE_CHAR, tcBefore))
        # If there's NO PREVIOUS TIME CODE ...
        else:
            # ... we don't need an offset.
            offset = 0
        # Screen updates are okay again
        self.Thaw()

        # Note the current time code
        self.current_timecode = tcBefore

        # Let's compare the current selection to the values we just found.  If they're different ...
        if pos != (start + offset, end):
            # ... then update the selection in the transcript ...
            self.SetSelection(start + offset, end)

            # ... we may need to scroll to see the new selection
            self.ShowCurrentSelection()

            # return True to indicate that the selection was moved
            return True
        # If we already have the correct text selection ...
        else:
            # ... return False to indicate that we have not moved anything
            return False
        
    def cursor_find(self, text):
        """Move the cursor to the next occurrence of given text in the
        transcript (for word tracking)."""
        # We first try searching from the current cursor position
        # for efficiency reasons (most of the time you're jumping just
        # ahead of the current cursor position).
        pos = self.FindText(self.GetCurrentPos(), self.GetLength()-1, text)
        if pos >= 0:
            # If using Unicode, we need to move one position to the right.
            if 'unicode' in wx.PlatformInfo:
                pos += 1
            try:
                self.GotoPos(pos)
            except wx._core.PyAssertionError, x:
                pass
        else:
            # Try searching in reverse
            pos = self.FindText(self.GetCurrentPos(), 0, text)
            if pos >= 0:
                # If using Unicode, we need to move one position to the right.
                if 'unicode' in wx.PlatformInfo:
                    pos += 2
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
        
        curpos = self.GetSelectionStart()
        endpos = self.FindText(curpos + 1, self.GetLength()-1, text)

        # If not found, try again from the beginning!
        if endpos ==  -1:
            endpos = self.FindText(0, self.GetLength(), text)
        # If not found, select the end of the document
        if endpos == -1:
            self.GetLength()

        # When searching for time codes for positioning of the selection, we end up selecting
        # the time code symbol and "<" that starts the time code.  We don't want to do that.
        if 'unicode' in wx.PlatformInfo:
            # Becaues the time code symbols vary with platform, we need to build a list of the characters
            # in the timecode character manually
            tcList = []
            tcList.append(ord('<'))
            for ch in TIMECODE_CHAR:
                tcList.append(ord(ch))
            # The first part of the time code characters doesn't get included under Unicode.  Add it here.
            if 'unicode' in wx.PlatformInfo:
                tcList.append(194)

            # If any of these characters is just to the left of where the end position is, move the end position to the left by 1 position
            while (endpos > 1) and (self.GetCharAt(endpos-1) in tcList):
                endpos -= 1
        else:
            while (endpos > 1) and (self.GetCharAt(endpos-1) in [ord('<'), ord(TIMECODE_CHAR)]):
                endpos -= 1

        # NOTE:  Calling SetSelection only caused the loss of the highlight for Locate Clip in Episode.
        # Calling SetCurrentPos() and SetAnchor maintains the highlight correctly.
        self.SetSelection(curpos, endpos)

        self.ShowPosition(self.GetInsertionPoint())

    def GetTextBetweenTimeCodes(self, startTime, endTime):
        """ Get the text between the time codes indicated """
        # This method is used for Episode Transcript Change Propagation.
        # Let's try to remember the cursor position
        self.SaveCursor()

        # Initialize start time code to zero, in case the first clip has no leading time code
        startTimeCode = 0
        # If the Start Time exactly matches an existing time code ...
        if startTime in self.timecodes:
            # ... then we can just use it.
            startTimeCode = startTime
        # If the Start Time does NOT exactly match an existing time code ...
        else:
            # ... iterate through the existing time codes ...
            for time in self.timecodes:
                # ... and find the time code immediately BEFORE the Start Time
                if time < startTime:
                    startTimeCode = time
                else:
                    break
        # If the End Time exactly matches an existing time code ...
        if endTime in self.timecodes:
            # ... then we can just use it.
            endTimeCode = endTime
        # If the End Time does NOT exactly match an existing time code ...
        else:
            # Default endTimeCode to 0 in order to avoid problems for transcripts that totally lack time codes
            endTimeCode = 0
            # ... iterate through the existing time codes ...
            for time in self.timecodes:
                # ... and find the time code immediately AFTER the End Time
                if time < endTime:
                    endTimeCode = time
                else:
                    endTimeCode = time
                    break
            # Check to see if the end time is AFTER the last time code.
            if endTime > endTimeCode:
                # If so, use the end time
                endTimeCode = endTime
        # Initialize the text to blank
        text = ''

        # Select between the time codes we found above
        self.scroll_to_time(startTime)
        self.select_find(str(endTime))

        # Set the text to the XML version of what we now have selected
        xmlText = self.GetFormattedSelection('XML', selectionOnly=True)  # self.GetXMLBuffer(select_only=True)
        # Let's try restoring the Cursor Position when all is said and done.
        self.RestoreCursor()
        # Return the start and end times that were found along with the text between them.
        return (startTimeCode, endTimeCode, xmlText)

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

    def set_read_only(self, state=True):
        """Enable or disable read-only mode, to prevent the transcript
        from being modified."""
        # Change the Read Only state in the RichTextEditCtrl
        self.SetReadOnly(state)
        # If AutoSave is ON and we are in a Transcript Window ...
        if TransanaGlobal.configData.autoSave and isinstance(self.parent, TranscriptionUI_RTC._TranscriptDialog):
            # If we are switching to READ ONLY ...
            if state:
                # ... we can turn OFF the AutoSave Timer
                self.autoSaveTimer.Stop()
            # If we are switching to EDIT MODE ...
            else:
                # ... we should turn ON the AutoSave Timer.  600,000 is every TEN MINUTES
                self.autoSaveTimer.Start(600000)

    def get_read_only(self):
        """ Report the current Read Only / Edit Mode status """
        return self.GetReadOnly()

    def find_timecode_before_cursor(self, pos):
        """Return the position of the first timecode before the given cursor position."""
        return self.FindText(pos, 0, "%s<" % TIMECODE_CHAR)
        
    def get_selected_time_range(self):
        """Get the time range of the currently selected text.  Return a tuple with the start and end times in milliseconds."""
        # If the transcript is long and lacks time codes (i.e. has just been imported), this can
        # be VERY slow.  I'm adding escape clauses to speed this method up when we're after the
        # last time code.  We assume people will time-code from the beginning of a transcript.

        # If the transcript has NO timecodes ...
        if len(self.timecodes) == 0:
            # ... return that we start at the beginning and finish at the end.
            return (0, -1)
        
        # Get the position of the start of the current selection
        pos = self.GetSelectionStart()
        # remember the initial selected position, just in case.
        initialPos = pos
        # Initialize the variable where we'll collect Time Code data
        numStr = ''

        # While the current character is not a time code (and we're not at the document start) ...
        while (pos > 0) and (self.GetCharAt(pos) != 164):
            # ... move towards the front of the document.  (i.e. find the preceding time code)
            pos -= 1
        # If we get to the start of the document without finding a time code ...
        if pos == 0 and ((initialPos == 0) or (self.GetCharAt(pos) != 164)):
            # ... then our time position is 0, the start of the media file.
            numStr = "0"
        # If we have found a time code ...
        else:
            # Initialize a blank string for collecting the time code value
            numStr = ""
            # Initialize the flag that tracks if we're adding numbers yet
            startAddingNumbers = False
            # Process text until we get to a ">" character (which closes time code data) or the document end.
            while (self.GetCharAt(pos) != 62) and (pos < self.GetTextLength() - 1):
                # Once we find a time code character ...
                if (self.GetCharAt(pos) == 164):
                        # ... we can start adding digits to the time code data
                    startAddingNumbers = True
                # If we're adding data to the time code data AND the character is a DIGIT ...
                if startAddingNumbers and (self.GetCharAt(pos) < 256) and (chr(self.GetCharAt(pos)) in string.digits):
                    # ... add it to the time code value string
                    numStr += chr(self.GetCharAt(pos))
                # increment the string position being examined
                pos += 1
                
        # Convert the time code value string to a long integer
        try:
            start_timecode = long(numStr)
        # If an exception occurs here ...
        except:
            # ... just assume the time code value is 0.
            start_timecode = 0
        # If we are positioned AFTER the LAST time-code ...
        if (start_timecode in self.timecodes) and (self.timecodes.index(start_timecode) == len(self.timecodes) - 1):
            # ... return the start time code and the end of the file
            return (start_timecode, -1)

        # Get the position of the end of the current selection
        pos = self.GetSelectionEnd()
        # Initialize the string for gathering the time code data
        numStr = ''
        # While the current character is not a time code (and we're not at the document end) ...
        while (pos < self.GetTextLength() - 1) and (self.GetCharAt(pos) != 164):
            # ... move towards the front of the document.  (i.e. find the following time code)
            pos += 1

        # If we get to the end of the document without finding a time code ...
        if pos == self.GetTextLength():
            # ... then our time position is -1, the end of the media file.
            numStr = "-1"
        # If we have found a time code ...
        else:
            # Initialize a blank string for collecting the time code value
            numStr = ""
            # We shouldn't start collecting time code data until we find a time code!!  Initialize the flag for that.
            startAddingNumbers = False
            # Process text until we get to a ">" character (which closes time code data) or the document end.
            while (self.GetCharAt(pos) != 62) and (pos < self.GetTextLength() - 1):
                # If we find a Time Code ...
                if (self.GetCharAt(pos) == 164):
                    # ... then we can start adding digits to the time code's time value.
                    startAddingNumbers = True
                # If we should be collecting time code data AND the character is a DIGIT ...
                if startAddingNumbers and (self.GetCharAt(pos) < 256) and (chr(self.GetCharAt(pos)) in string.digits):
                    # ... add it to the time code value string
                    numStr += chr(self.GetCharAt(pos))
                # increment the string position being examined
                pos += 1
        # Convert the time code value string to a long integer
        try:
            end_timecode = long(numStr)
        # If an exception occurs here ...
        except:
            # ... just assume the time code value is -1.
            end_timecode = -1
        # If you have two consecutive time codes with no text between, this routine is producing the
        # WRONG RESULTS!!  You get what should be the end time code for both values.
        # So if start_timecode and end_timecode are the same ...
        if start_timecode == end_timecode:
            # ... if we're not at the FIRST time code in the document ...
            if (start_timecode > self.timecodes[0]):
                # ... then select the next earliest time code as the correct start_timecode
                start_timecode = self.timecodes[self.timecodes.index(start_timecode) - 1]
            # If you ARE at the first time code ...
            elif start_timecode == 0:

                if len(self.timecodes) > 1:
                    end_timecode = self.timecodes[1]
                else:
                    end_timecode = -1
            else:
                # ... then our start should be 0
                start_timecode = 0
        return (start_timecode, end_timecode)

    def ClearDoc(self, skipUnlock = False):
        """ Clear the Transcript Window """
        # If the current Transcript is locked ...
        if (self.TranscriptObj != None) and (self.TranscriptObj.isLocked) and not skipUnlock:
            # ... unlock it.  (Saving has already been taken care of.)
            self.TranscriptObj.unlock_record()
        # Make the RichTextEditCtrl editable!
        self.set_read_only(False)
        # Clear the document from the control
        RichTextEditCtrl.ClearDoc(self)
        # Reset the media time to 0
        self.TimePosition = 0
        # Clear the Transcript Object
        self.TranscriptObj = None
        # Clear the time code list
        self.timecodes = []
        # Clear the current time code pointer
        self.current_timecode = -1
        # Make the control read-only
        # THIS CAUSES A BUG!!
        # Load a transcript, choose File > New, then add a new transcript from RTF with a graphic at the top.
        # The transcript doesn't load correctly!  The top image is shown at the bottom of the transcript!
        # BUT REMOVING IT MEANS THE TRANSCRIPT ISN'T READ ONLY INITIALLY, WHICH IS BAD.
        self.set_read_only(True)
        
        # Signal that time codes are not visible
        self.codes_vis = 0
        # Signal that Time Code Data should be hidden too.
        self.timeCodeDataVisible = False

    def PrevTimeCode(self, tc=None):
        """Return the timecode immediately before the current one."""
        # If no time code is passed ...
        if tc == None:
            # ... use the current time code value
            tc = self.current_timecode
        # If the current time code is -1 ...
        if tc == -1:
            # ... then return 0 to signal to start at the beginning of the file
            return 0

        # Start Exception Handling
        try:
            # Try to get the Index in the time codes list of the item BEFORE the current time code value (hence -1)
            i = self.timecodes.index(tc) - 1
            # If our index is 0 or higher ...
            while i >= 0:
                # If the Time Code List value is less than our original time code ...
                if self.timecodes[i] < tc:
                    # ... then return the value
                    return self.timecodes[i]
                # Otherwise, move a value earlier in the list and try again
                i = i - 1
            # If we exit the while loop without RETURN, then return 0 to signal to start at the beginning of the file
            return 0
        # If an exception is raised ...  (probably an index out of range call.)
        except:
            if DEBUG:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
            # ... return 0 to signal to start at the beginning of the file
            return 0

    def NextTimeCode(self, tc=None):
        """Return the timecode immediately after the current one."""
        # If no time code is passed ...
        if tc == None:
            # ... use the current time code value
            tc = self.current_timecode
        # If the current time code is -1 ...
        if tc == -1:
            # ... and if there are time code values in the list ...
            if len(self.timecodes) > 0:
                # ... then return the FIRST value in the list.
                return self.timecodes[0]

        # Start Exception Handling
        try:
            # Try to get the Index in the time codes list of the item AFTER the current time code value (hence +1)
            i = self.timecodes.index(tc) + 1
            # If our index is less than the highest value ...
            while i < len(self.timecodes):
                # If the Time Code List value is greater than our original time code ...
                if self.timecodes[i] > tc:
                    # ... then return the value
                    return self.timecodes[i]
                # Otherwise, move a value later in the list and try again
                i = i + 1
            # If we get here, we got to the end of the list.  Just return the list's highest value
            return self.timecodes[-1]
        # If an exception is rasied ...  (probably an index out of range call.)
        except:
            if DEBUG:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
            # If there are time codes in the time codes list ...
            if len(self.timecodes) > 0:
                # ... return the highest value in the list
                return self.timecodes[-1]
            else:
                # ... return -1 to signal failure
                return -1

    def OnKeyDown(self, event):
        """ Called when a key is pressed down.  All characters are upper case.  """

        # ... let's try to remember the cursor position.  (We've had problems with the cursor moving during transcription on the Mac.)
        self.SaveCursor()

        # See if the ControlObject wants to handle the key that was pressed.
        if isinstance(self.parent, TranscriptionUI_RTC._TranscriptDialog) and \
           self.parent.ControlObject.ProcessCommonKeyCommands(event):
            # If so, we're done here.
            return

        # We are having trouble with duplicate key events in certain circumstances.  This should eliminate that.
        # Keys to be excluded from calling the RichTextEditCtrl's OnKeyDown() handler include:
        #    Ctrl-Cursor Up
        #    Ctrl-Cursor Down
        #    Ctrl-H
        #    Ctrl-O
        # These are the Special Symbols in Jeffersonian Notation!!
        if not event.ControlDown() or \
           not (event.GetKeyCode() in [wx.WXK_UP, wx.WXK_DOWN, ord("B"), ord("H"), ord("I"), ord("N"), ord("O"), ord("P"), ord("U")]):
            # Call the RichTextEditCtrl's On Key Down handler
            RichTextEditCtrl.OnKeyDown(self, event)

        # It might be necessary to block the "event.Skip()" call.  Assume for the moment that it is not.
        blockSkip = False

        # If the Control key is being held down ...
        if event.ControlDown():
            try:
                # Get the Code for the key that was pressed
                c = event.GetKeyCode()
                # NOTE:  NON-ASCII keys must be processed first, as the chr(c) call raises an exception!
                if c == wx.WXK_UP:
                    # Ctrl-Cursor-Up inserts the Up Arrow / Rising Intonation symbol
                    self.InsertRisingIntonation()
                    return
                elif c == wx.WXK_DOWN:
                    # Ctrl-Cursor-Down inserts the Down Arrow / Falling Intonation symbol
                    self.InsertFallingIntonation()
                    return
                elif c == ord("H"):
                    # Ctrl-H inserts the High Dot / Inbreath symbol
                    self.InsertInBreath()
                    return
                elif c == ord("O"):
                    # Ctrl-O inserts the Open Dot / Whispered Speech symbol
                    self.InsertWhisper()
                    return
                elif c == ord("B"):
                    self.set_bold()
                    self.StyleChanged(self)
                    blockSkip = True
                elif c == ord("U"):
                    self.set_underline()
                    self.StyleChanged(self)
                    # Block the Skip() call to prevent Ctrl-U from calling the Lower Case function that apparently wx.STC provides(?)
                    blockSkip = True
                elif c == ord("I"):
                    self.set_italic()
                    self.StyleChanged(self)
                    blockSkip = True
                elif c == ord("K"):
                    # CTRL-K: Quick Clip Shortcut Key
                    # Ask the Control Object for help creating a Quick Clip
                    self.parent.ControlObject.CreateQuickClip()
                    return
                
            except:
                pass    # Non-ASCII value key pressed

        # If Control is NOT pressed ...
        else:
            # Because of time codes and hidden text, we need a bit of extra code here to make sure the cursor is not left in
            # the middle of hidden text and to prevent accidental deletion of hidden time codes.
            c = event.GetKeyCode()
            curpos = self.GetCurrentPos()
            cursel = self.GetSelection()

            # If the are moving to the LEFT with the cursor ...
            if (c == wx.WXK_LEFT):
                # ... and we come to a TIMECODE Character ...
                if (curpos > 0) and (self.GetCharAt(curpos - 1) < 256) and (self.GetCharAt(curpos - 1) == ord('>')) and (self.IsStyleHiddenAt(curpos)):
                    # ... then we need to find the start of the time code data, signalled by the TIMECODE character ...
                    while (self.GetCharAt(curpos - 1) != ord(TIMECODE_CHAR)):
                        curpos -= 1
                    # If Unicode, we need to move 2 more characters to the left.  This is because of the way the wxSTC
                    # handles the 2-byte timecode character.
                    if 'unicode' in wx.PlatformInfo:
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
                        self.SetCurrentPos(cursel[1])
                        self.SetAnchor(curpos)
            # If the are moving to the RIGHT with the cursor ...
            elif (c == wx.WXK_RIGHT):
                # ... and we come to a TIMECODE Character ...
                # The evaluation to determine we've come to a time code is a little weird under Unicode.
                if 'unicode' in wx.PlatformInfo:
                    if 'wxMac' in wx.PlatformInfo:
                        evaluation = (self.GetCharAt(curpos) == 194) and (self.GetCharAt(curpos + 1) == 164)  #167)
                    else:
                        evaluation = (self.GetCharAt(curpos) == 194) and (self.GetCharAt(curpos + 1) == 164)
                else:
                    evaluation = (self.GetCharAt(curpos) == ord(TIMECODE_CHAR))
                if evaluation:
                    # ... then we need to find the end of the time code data, signalled by the '>' character ...
                    while self.GetCharAt(curpos) != ord('>'):
                        curpos += 1
                    # If Time Codes are not visible, we need one more character here.  It doesn't make sense
                    # to me, as we should be at the end of the time code data, but we DO need this.
                    if not(self.codes_vis):
                        curpos += 1

                    # Check the formatting and update if needed, so the formatting will be correct, not time-code red.
                    # if the current styling is HIDDEN ...
                    if self.GetStyleAt(curpos) == self.STYLE_HIDDEN:
                        # ... and the style spec is for time codes ...
                        if self.style_specs[self.style] == 'timecode':
                            # ... then reset to the default style
                            self.style = 0

                    # If time code data is visible ...
                    if self.timeCodeDataVisible and self.GetCharAt(curpos+1) == ord('('):
                        # ... then as long as we're inside the time code data ...
                        while self.GetCharAt(curpos + 1) in [ord('0'), ord('1'), ord('2'), ord('3'), ord('4'), ord('5'),
                                                             ord('6'), ord('7'), ord('8'), ord('9'), ord(':') ,ord('.'),
                                                             ord('('), ord(')')]:
                            # ... we should keep moving to the right.
                            curpos += 1
                            # Once we hit the close parents, we should stop though.
                            if self.GetCharAt(curpos) == ord(')'):
                                break
                                
                    # If you cursor over a time code while making a selection, the selection was getting lost with
                    # the original code.  Instead, determine if a selection is being made, and if so, make a new
                    # selection appropriately.
                    # If these values differ, we're selecting rather than merely moving.
                    if cursel[0] == cursel[1]:
                        # Position the cursor after the hidden timecode data
                        self.GotoPos(curpos)
                    else:
                        self.SetCurrentPos(cursel[0])
                        self.SetAnchor(curpos)

            # DELETE KEY pressed
            elif (c == wx.WXK_DELETE):
                # Delete is handled by RichTextEditCtrl.OnKeyDown().
                # Prevent doubling of delete by blocking event.skip()
                blockSkip = True

            # BACKSPACE KEY pressed
            elif (c == wx.WXK_BACK):
                # Backspace is handled by RichTextEditCtrl.OnKeyDown().
                # Prevent doubling of backspace by blocking event.skip()
                blockSkip = True

##            elif event.AltDown() and (c == wx.WXK_F1):
##
##                (s, e) = self.get_selected_time_range()
##
##                print "Current Pos:", s, Misc.time_in_ms_to_str(s), e, Misc.time_in_ms_to_str(e), self.GetLastPosition()
##                
##                print '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'
##                print self.timecodes
##                print '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'

            elif c == wx.WXK_F12:
                # F12 is Quick Save
                self.save_transcript()
                blockSkip = True

            # If anything not explicitly handled is entered ...
            else:
                pass

        # if the Skip() call should not be blocked ...
        if not(blockSkip) and not event.ControlDown():
            # ... skip to the next (parent) level event handler.
            event.Skip()

        if isinstance(self.parent, TranscriptionUI_RTC._TranscriptDialog):
            # Always check the styles to see if the Transcript Toolbar needs to be updated.  But we need to defer the call
            # until everything else has been handled so it's always correct.
            wx.CallAfter(self.StyleChanged, self)

    def OnKey(self, event):
        """Called when a character key is pressed.  Works with case-sensitive characters.  """

        # If the style is currently hidden ...
        if self.IsStyleHiddenAt(self.GetInsertionPoint()):
            # ... then change the current text style to the current style
            self.SetDefaultStyle(self.txtAttr)

        # Call the RichTextEditCtrl's On Key handler
        RichTextEditCtrl.OnKey(self, event)

    def OnKeyUp(self, event):
        """ Catch the release of each key """
        # Call the super object's OnKeyUp method
        RichTextEditCtrl.OnKeyUp(self, event)
        # If a Style Changed method has been defined ...
        if self.StyleChanged != None:
            # ... make sure any style change is reflected on screen
            self.StyleChanged(self)

#        if event.GetKeyCode() != wx.WXK_SHIFT:
#            print "TranscriptEditor_RTC.OnKeyUp() (WITH SKIP)", event.GetKeyCode(), wx.WXK_SHIFT

#       DISABLED because this causes this method to be called TWICE on each keystroke.            
#        event.Skip()

    def OnAutoSave(self, event):
        """ Process the AutoSave Timer Event """
        # Determine what control HAS focus
        focusCtrl = self.FindFocus()
        # If the media is playing ...
        if self.parent.ControlObject.IsPlaying():
            # ... remember it was playing ...
            isPlaying = True
            # ... and pause it for a moment
            self.parent.ControlObject.Pause()
        # If the media is NOT playing ...
        else:
            # ... remember it was NOT playing
            isPlaying = False
        
        # Freeze the control for a moment
        self.Freeze()
        # Save the current window label
        lbl = self.parent.GetLabel()
        # Replace the label to indicate we are auto-saving
        self.parent.SetLabel(_("Auto-saving ..."))
        # Save the transcript
        if TransanaConstants.partialTranscriptEdit:
            self.save_transcript(continueEditing=True)
        else:
            self.save_transcript()
        # Restore the window label to its original value
        self.parent.SetLabel(lbl)
        # Okay, we're done
        self.Thaw()

        # If the media was playing ...
        if isPlaying:
            # ... tell it to play again!
            self.parent.ControlObject.Play()

        # Start exception handling
        try:
            # Try to return focus to whatever control had it
            focusCtrl.SetFocus()
        # Handle exceptions
        except:
            # If the original control that had focus is not available, set focus
            # to the transcript!
            self.SetFocus()
            
##        # If Partial Transcript editing is enabled ...
##        if TransanaConstants.partialTranscriptEdit:
##            # If we have only part of the transcript in the editor, we need to restore the partial transcript state following save
##            self.UpdateCurrentContents('EnterEditMode')

    def CheckTimeCodesAtSelectionBoundaries(self):
        """ Check the start and end of a selection and make sure neither is in the middle of a time code """
        # We need to make sure the cursor is not positioned between a time code symbol and the time code data, which unfortunately
        # can happen.  Preventing this is the sole function of this section of this method.

        # NOTE:  This sort of code seems to appear in at least 3 spots.
        #          1.  TranscriptEditor.OnLeftUp
        #          2.  TranscriptEditor.OnStartDrag
        #          3.  RichTextEditCtrl.__ProcessDocAsRTF
        #        I just added the third, as the first two weren't adequate under Unicode
        #   *** Better look at select_find() too!!
        #        DKW, 5/8/2006.  I'm going with the theory that I don't need to evaluate in RichTextEditCtrl any more,
        #                        since I think I just corrected the Unicode problem here.

        # Let's see if any change is made.
        selChanged = False
        # We have a click-drag Selection.  We need to check the start and the end points.
        selStart = self.GetSelectionStart()
        selEnd = self.GetSelectionEnd()
        # First, see if we have a click Position or a click-drag Selection.
        if (selStart != selEnd):
            # Get the selected text
            selString = self.GetStringSelection()
            # Get the STYLE of the first character in the selection
            defStyle = self.GetStyleAt(selStart)
            # if the first character is a time code or has any of the time code-related formats ...
            if (len(selString) > 0) and \
               ((selString[0] == TIMECODE_CHAR) or \
                (self.CompareFormatting(defStyle, self.txtTimeCodeAttr, fullCompare=False) or \
                self.CompareFormatting(defStyle, self.txtHiddenAttr, fullCompare=False) or \
                self.CompareFormatting(defStyle, self.txtTimeCodeHRFAttr, fullCompare=False))):
                # ... indicate that the selection should start after the time code data
                selStart = selStart + selString.find('>') + 1
                selChanged = True
            # if the selected text ends with a time code and time code data, or a time code and incomplete time code data
            if (selString.find('>', selString.rfind(TIMECODE_CHAR)) == len(selString) - 1) or \
               (selString.rfind(TIMECODE_CHAR) > selString.rfind('>')):
                # ... move the selection end to just before the final time code.
                selEnd = self.GetSelectionStart() + selString.rfind(TIMECODE_CHAR)
                selChanged = True

        else:
            # We have a click Position.  We just need to check the current position.
            curPos = self.GetCurrentPos()
            # Let's see if we are between a Time code and its Data
            if 'unicode' in wx.PlatformInfo:
                if 'wxMac' in wx.PlatformInfo:
                    evaluation = (self.GetCharAt(curPos - 1) == 194) and (self.GetCharAt(curPos) == 164)  #167)
                    evaluation2 = (self.GetCharAt(curPos - 2) == 194) and (self.GetCharAt(curPos - 1) == 164)  #167)
                    evaluation = evaluation or evaluation2
                else:
                    evaluation = (self.GetCharAt(curPos - 1) == 194) and (self.GetCharAt(curPos) == 164)
                    evaluation2 = (self.GetCharAt(curPos - 2) == 194) and (self.GetCharAt(curPos - 1) == 164)
                    evaluation = evaluation or evaluation2
            else:
                evaluation = (curPos > 0) and (self.GetCharAt(curPos - 1) == ord(TIMECODE_CHAR)) and (self.GetCharAt(curPos) == ord('<')) 

            if evaluation:
                # Let's find the position of the end of the Time Code
                while self.GetCharAt(curPos - 1) != ord('>'):
                    curPos += 1
                self.GotoPos(curPos)
                self.SetAnchor(curPos)
                self.SetCurrentPos(curPos)

        # If the selection has changed ...
        if selChanged:
            # ... then actually change the selection!
            wx.CallAfter(self.SetSelection, selStart, selEnd)

    def OnStartDrag(self, event, copyToClipboard=False):
        """Called on the initiation of a Drag within the Transcript."""
        # If we don't have a Transcript Object ...
        if not self.TranscriptObj:
            # ... abort the Start Drag 'cause the interface is empty.
            return

        # Let's get the time code boundaries.  This will return a start_time of 0 if there's not initial time code,
        # and an end_time of -1 if there's no ending time code.
        (start_time, end_time) = self.get_selected_time_range()

        # If we're creating a Clip from a Clip, we may need to make some minor adjustments to the clips start and stop times and
        # to the video/audio checkbox data.
        # First, let's see if we're in an Episode Transcript or a Clip Transcript.
        if self.TranscriptObj.clip_num != 0:
            # We're in an Clip Transcript.  If we don't have a starting time code ...
            if start_time == 0:
                # ... we need to use the CLIP's start as the sub-clip's start time, not the start of the video (0:00:00.0)!
                start_time = self.parent.ControlObject.GetVideoStartPoint()
            # Get the initial Video Checkbox data from the Video Window
            videoCheckboxData = self.parent.ControlObject.GetVideoCheckboxDataForClips(start_time)
            # Create a temporary variable for the Clip, just to make things easier
            tmpClip = self.parent.ControlObject.currentObj
            # Get a temporary copy of the source Episode for comparison purposes
            tmpEpisode = Episode.Episode(tmpClip.episode_num)

            # First, let's find out if the EPISODE's main file is included in the Clip.
            if tmpEpisode.media_filename == tmpClip.media_filename:
                # If so, its video would be checked, and the CLIP's audio value tells us if its audio would be checked.
                videoCheckboxData = [(True, tmpClip.audio)]
            # If not ...
            else:
                # ... then neither box would be checked.
                videoCheckboxData = [(False, False)]
            # Now let's iterate through the EPISODE's additional media files
            for addEpVidData in tmpEpisode.additional_media_files:
                # If this Episode Additional Media File is the Clip's Main Media File ...
                if (addEpVidData['filename'] == tmpClip.media_filename):
                    # ... then the Video Checkbox would have been checked, and the object tells us if audio was too.
                    videoCheckboxData.append((True, addEpVidData['audio']))
                # If the Clip's main media file was already established, we need to check the Episode additional files
                # against the Clip Additional media files.
                else:
                    # Assume the video will not be found ...
                    vidFound = False
                    # ... and that the audio should NOT be included.
                    audIncluded = False
                    # Iterate throught the CLIP additional media files ...
                    for addClipVidData in tmpClip.additional_media_files:
                        # If the Clip Media File matches the Episode Media File ...
                        if addClipVidData['filename'] == addEpVidData['filename']:
                            # ... then we've found the Media File!
                            vidFound = True
                            # Note if audio should be included, getting data from the Clip Object
                            audIncluded = addClipVidData['audio']
                            # Stop looking!
                            break
                    # Add info to videoCheckboxData to indicate whether the Episode's media file and audio were checked when
                    # the source clip was created.
                    videoCheckboxData.append((vidFound, audIncluded))

        # If we are clipping from an Episode ...
        else:
            # ... we need to get the Video Checkbox Data from the Video Window.  No further adjustment is needed.
            videoCheckboxData = self.parent.ControlObject.GetVideoCheckboxDataForClips(start_time)
            
        # If we don't have an end Time Code ...
        if end_time == -1:
            # We need the VideoEndPoint.  This is accurate for either Episode Transcripts or Clip Transcripts.
            end_time = self.parent.ControlObject.GetVideoEndPoint()

        # If there is no selection in the Transcript ...
        if not self.HasSelection():
            # ... get the text between the nearest time codes
            (start_time, end_time, xmlText) = self.GetTextBetweenTimeCodes(start_time, end_time)
        # Otherwise ...
        else:
            # ... let's get the selected Transcript text in XML format
            xmlText = self.GetFormattedSelection('XML', selectionOnly=True)  # self.GetXMLBuffer(select_only=1)

        # Create a ClipDragDropData object with all the data we need to create a Clip
        data = DragAndDropObjects.ClipDragDropData(self.TranscriptObj.number, self.TranscriptObj.episode_num, \
                start_time, end_time, xmlText, self.GetStringSelection(), videoCheckboxData)
        # let's convert that object into a portable string using cPickle. (cPickle is faster than Pickle.)
        pdata = cPickle.dumps(data, 1)
        # Create a CustomDataObject with the format of the ClipDragDropData Object
        cdo = wx.CustomDataObject(wx.CustomDataFormat("ClipDragDropData"))
        # Put the pickled data object in the wxCustomDataObject
        cdo.SetData(pdata)

        # If we are supposed to copy the data to the Clip Board ...
        if copyToClipboard:
            # Open the Clipboard
            wx.TheClipboard.Open()
            # ... then copy the data to the clipboard!
            wx.TheClipboard.SetData(cdo)
            # Close the Clipboard
            wx.TheClipboard.Close()
            
        else:
            # Put the data in the DropSource object
            tds = TranscriptDropSource(self.parent)
            tds.SetData(cdo)

            # Initiate the drag operation.
            # NOTE:  Trying to use a value of wx.Drag_CopyOnly to resolve a Mac bug (which it didn't) caused
            #        Windows to stop allowing Clip Creation!
            dragResult = tds.DoDragDrop(wx.Drag_AllowMove)

            # Okay, it turns out that with Drag-and-Drop turned off on the Mac, this code wasn't ever being called anyway.
            # So now here it is, years later, (9/2008) and we're using wxPython 2.8.7.1 and now I'm trying to enable Drag-
            # and-Drop on the Mac because Apple seems to have fixed the bug in QuickTime that made it so we couldn't use
            # it before.
            #
            # On the Mac, if we drag from the transcript, there's still some weirdness.  If we're in read-only mode,
            # the selection gets un-selected.  If we're in edit more, the selection loses all formatting.  So this is
            # still an ugly hack, but I still can't figure out what else to do to get around this problem.

            # If we're on the Mac ...
            if ("wxMac" in wx.PlatformInfo):
                # If the "selection" attribute doesn't exist ... (This came up with multi-transcript quick clips where the
                # selection in the second transcript occurred automatically!)
                if not hasattr(self, 'selection'):
                    # ... then create it!!
                    self.selection = self.GetSelection()
                # ... if we're in Edit mode ...
                if(not self.get_read_only()):
                    # ... temporarily slip the RTF control into read-only mode (not all of Transana) ...
                    self.set_read_only(True)
                    # ... and signal that we should slip back into edit mode "later" ...
                    wx.CallAfter(self.set_read_only, False)
                # ... and restore the original Selection
                self.SetSelection(self.selection[0], self.selection[1])

        # Reset the cursor following the Drag/Drop event
        self.parent.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

    def PositionAfter(self):
        """ We need a mechanism for setting or restoring the cursor position using wx.CallAfter().  This method enables that. """
        self.GotoPos(self.pos)
        self.ShowCurrentSelection()

    def SetSelectionAfter(self):
        """ We need a mechanism for setting or restoring a transcript selection using wx.CallAfter().  This method enables that. """
        self.SetCurrentPos(self.selection[0])
        self.SetAnchor(self.selection[1])

    def OnLeftDown(self, event):
        """ Left Mouse Button Down event """

        ## # Get the Transcript Scale from the Configuration Data
        ## scaleFactor = TransanaGlobal.configData.transcriptScale
        ## # Set the Transcript Scale Factor.  (self.GetScaleX() and self.GetScaleY() do not work !!!)
        ## scale = (scaleFactor, scaleFactor)

        # Note the Mouse Position.  Used in OnLeftUp so see if de-selection is appropriate.
        self.mousePosition = event.GetPosition()

        if event.GetEventObject().HasSelection():
            # Determine the start and end character numbers of the current selection
            textSelection = event.GetEventObject().GetSelection()
            # If we're using wxPython 2.8.x.x ...
            if wx.VERSION[:2] == (2, 8):
                # ... use HitTest()
                mousePos = event.GetEventObject().HitTest(event.GetPosition())[1]
            # If we're using a later wxPython version ...
            else:
                (posx, posy) = event.GetPosition()
                ## posx /= scale[0]
                ## posy /= scale[1]
                # ... use HitTestPos()
                mousePos = event.GetEventObject().HitTestPos((posx, posy))[1]

            # If the Mouse Character is inside the selection ...
            if (textSelection[0] <= mousePos) and (mousePos < textSelection[1]):
                self.canDrag = True
            else:
                self.canDrag = False
                event.Skip()

        else:
            # The Mac doesn't show the cursor properly in wxPython 3.0.0.0.
            self.GetCaret().Show()

            self.canDrag = False
            event.Skip()

    def OnLeftUp(self, event):
        """ Left Mouse Button Up event """
        # Call the parent control's OnLeftUp event.  Without this, located before SaveCursor(), you can't click out of a selection!
        RichTextEditCtrl.OnLeftUp(self, event)
        # Save the original Cursor Position / Selection so it can be restored later.  Otherwise, we occasionally can't
        # make a new selection.
        self.SaveCursor()
        # Note the current Position if PositionAfter needs to be called
        self.pos = self.GetCurrentPos()
        # Set the selection in case SetSelectionAfter gets called later
        self.selection = self.GetSelection()
        # We need to make sure the cursor is not positioned between a time code symbol and the time code data, which unfortunately
        # can happen.
        self.CheckTimeCodesAtSelectionBoundaries()
        # The code that is now indented was causing an error if you tried to edit the Transcript from the
        # Clip Properties form, as the parent didn't have a ControlObject property.
        # Therefore, let's test that we're coming from the Transcript Window before we do this test.
        if type(self.parent).__name__ == '_TranscriptDialog':
            # Get the Start and End times from the time codes on either side of the cursor
            (segmentStartTime, segmentEndTime) = self.get_selected_time_range()
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
            # If we're in Edit mode ...
            else:
                # we don't update the Video Selection, but we still need to update the Selection Text for ALL transcripts
                wx.CallLater(200, self.parent.ControlObject.UpdateSelectionTextLater, self.parent.transcriptWindowNumber)

        # See if the mouse has moved.  If the mouse HAS moved ...
        if (self.mousePosition != event.GetPosition()):
            # ... and restore the original Selection
            self.SetSelection(self.selection[0], self.selection[1])
            
        # If the mouse has NOT moved, the following code allows "click-in-selection" deselection of a transcript highlight.
        else:
            # If a DRAG is possible ...
            if self.canDrag:
                # ... then set the insertion point to the mouse's current position (We're IN the selection)
                # If we're using wxPython 2.8.x.x ...
                if wx.VERSION[:2] == (2, 8):
                    # ... use HitTest()
                    self.SetInsertionPoint(event.GetEventObject().HitTest(event.GetPosition())[1])
                # If we're using a later wxPython version ...
                else:
                    # ... use HitTestPos()
                    self.SetInsertionPoint(event.GetEventObject().HitTestPos(event.GetPosition())[1])
            # If a DRAG is NOT possible ...
            else:
                # ... then set the insertion point to the last good position
                self.SetInsertionPoint(self.pos)
                # ... and re-call the parent's OnLeftUp() method
                RichTextEditCtrl.OnLeftUp(self, event)
        # If the style has changed ...
        if self.StyleChanged != None:
            # ... make sure the style change is processed
            self.StyleChanged(self)

        # Get the text style for the current position
        textAttr = self.GetStyleAt(self.GetInsertionPoint())
        # If we're at the end of the document or at a Time Code ...
        if (self.GetInsertionPoint() == self.GetLastPosition()) or \
           self.CompareFormatting(textAttr, self.txtHiddenAttr, fullCompare=False) or \
           self.CompareFormatting(textAttr, self.txtTimeCodeAttr, fullCompare=False) or \
           self.CompareFormatting(textAttr, self.txtTimeCodeHRFAttr, fullCompare=False):
            # ... then check Formatting to make sure our current style is correct
            self.CheckFormatting()

        ip = self.GetInsertionPoint()
        # If we're not at the first character in the document AND
        # we're not at the first character following a time code ...
        if (ip > 1) and not self.IsStyleHiddenAt(ip - 1):
            # ... update the style to the style for the character PRECEEDING the cursor
            self.txtAttr = self.GetStyleAt(ip - 1)
        # Otherwise ...
        else:
            # update the style to the style for the character FOLLOWING the cursor
            self.txtAttr = self.GetStyleAt(ip)

    def OnMotion(self, event):
        """ Process the EVT_MOTION event for the Rich Text Ctrl """
        # Call the parent EVT_MOTION event
        event.Skip()
        # Also call the RichTextEditCtrl's OnMotion event
        RichTextEditCtrl.OnMotion(self, event)
        # If we are DRAGGING ...
        if (self.canDrag and event.GetEventObject().HasSelection() and event.Dragging()):
            # ... signal the start of a DRAG event
            self.OnStartDrag(event)
            # We were getting odd double-drops with wxPython 2.8.12.0.  Signalling that we can't DRAG any more prevents this!
            self.canDrag = False

    def OnRightDown(self, event):
        """ Handle Right mouse button Down event """
        # We need to DISABLE this event!
        pass

    def OnRightUp(self, event):
        """ Right-clicking should handle Video Play Control rather than providing the
            traditional right-click editing control menu """
        # If nothing is loaded in the main interface ...
        if (self.parent.ControlObject == None) or (self.parent.ControlObject.currentObj == None):
            # ... then exit right here!  There's nothing to do.
            return
        # If we do not already have a cursor position saved, save it
        self.parent.ControlObject.SaveAllTranscriptCursors()

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

            if self.get_read_only():
                self.scroll_to_time(segmentStartTime)
        # Play or Pause the video, depending on its current state
        self.parent.ControlObject.PlayStop()

    def SaveCursor(self):
        """ Capture the current cursor position so that it can be restored later """
        self.cursorPosition = (self.GetCurrentPos(), self.GetSelection())

    def RestoreCursor(self):
        """ Restore the Cursor position following right-click play control operation """
        # If we have a stored Cursor Position ...
        if (self.cursorPosition != 0):
            if (self.cursorPosition[1] == (-2, -2)) or (self.cursorPosition[1][1] == -2) or (self.cursorPosition[1][0] == self.cursorPosition[1][1]):
                self.SetCurrentPos(self.cursorPosition[0])
            else:
                self.SetSelection(self.cursorPosition[1][0], self.cursorPosition[1][1])

            #  Only scroll if we're not in Edit Mode.
            if self.get_read_only():
                # now scroll so that the selection start is shown.
                self.ShowCurrentSelection()
            # Once the cursor position has been reset, we need to clear out the Cursor Position Data
            self.cursorPosition = 0

    def AutoTimeCode(self):
        """ Auto-timecode a transcript with fixed interval time codes """
        # Prepare a message for the user
        msg = _('Please enter the desired time-code interval in seconds.')
        # Create a text entry dialog box
        dlg = wx.TextEntryDialog(self, msg, _('Fixed-Increment Time Codes'), '')
        # Center the dialog on the screen, not the transcript window
        dlg.CentreOnScreen()
        # Get input from the user
        result = dlg.ShowModal()
        # If the user presses OK ...
        if result == wx.ID_OK:
            # Start exception handling
            try:
                # Convert the text input to a real number
                adjustValue = float(dlg.GetValue())
                # If the input is <= 0, exit, signalling failure
                if adjustValue <= 0:
                    return False
                # If we have an Episode loaded ...
                if type(self.parent.ControlObject.currentObj) == type(Episode.Episode()):
                    # ... then we will time code from the beginning (zero) to the end of the media file
                    start = 0
                    end = self.parent.ControlObject.currentObj.tape_length
                # If we have a Clip ...
                else:
                    # ... then we will time code from the clip start to the clip end.
                    # (It's unlikely this will get called, as Clips generally already have transcripts!)
                    start = self.parent.ControlObject.currentObj.clip_start
                    end = self.parent.ControlObject.currentObj.clip_stop
                # For the specified time range, using the specified interval ...
                for tc in range(start, end, int(adjustValue * 1000)):
                    # ... go to the end of the document (which keeps moving!) ...
                    self.GotoPos(self.GetLength())
                    # ... insert the appropriate time code ...
                    self.insert_timecode(tc)
                    # ... and insert a couple of blank lines at the end
                    self.WriteText(' \n')

                # Now that we've placed the time codes, go to the beginning of the document
                self.GotoPos(0)

            # Exception handling
            except:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                # Build the error message
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Error in Time Code increment.\n%s\n%s'), 'utf8')
                else:
                    prompt = _('Error in Time Code increment.\n%s\n%s')
                # Display the error message
                errordlg = Dialogs.ErrorDialog(self, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
                adjustValue = 0
        # Destroy the prompt dialog
        dlg.Destroy()
        # Signal the calling routine whether the user entered a legal value
        return ((result == wx.ID_OK) and (adjustValue > 0))
    
    def AdjustIndexes(self, adjustmentAmount):
        """ Adjust Transcript Time Codes by the specified amount """
        # Let's try to remember the cursor position
        self.SaveCursor()
        
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
            # Find the time code
            self.cursor_find("%s<%d>" % (TIMECODE_CHAR, self.timecodes[loop]))
            # Remember the starting position, adjusted for the width of the Time Code Character
            start = self.GetCurrentPos() - len(TIMECODE_CHAR)
            # We need to determine the end position the hard way.  select_find() was giving the wrong
            # answer because of the CheckTimeCodesAtSelectionBoundaries() call.
            # So start at the beginning of the time code ...
            end = start
            # ... and keep moving until we find the first ">" character, which closes the time code.
            while self.GetCharAt(end - 1) != ord('>'):
                end += 1

            # Now select the smaller selection, which should just be the Time Code Data
            self.SetSelection(start, end)

            # Finally, replace the old time code data with the new time code data.
            # First delete the old data
            self.DeleteSelection()
            # Then insert the new data as hidden text (which is why we can't just plug it in above.)
            self.InsertTimeCode(self.timecodes[loop] + int(adjustmentAmount* 1000))
            # Adjust the local list of Transcript time codes too!
            self.timecodes[loop] = self.timecodes[loop] + int(adjustmentAmount * 1000)
       

        # We better hide all the hidden text for the time codes again
        self.hide_all_hidden()
        # We also need to reset self.codes_vis, which was incorrectly changed by hide_all_hidden()
        self.codes_vis = codes_vis
            
        # Okay, this might not work because of changes we've made to the transcript, but let's
        # try restoring the Cursor Position when all is said and done.
        self.RestoreCursor()
        self.Update()

    def TextTimeCodeConversion(self):
        """ Convert Text (H:MM:SS.hh) Time Codes to Transana's Format """
        # Create a popup telling the user about the conversion (needed for large files)
        convertDlg = Dialogs.PopupDialog(None, _("Converting..."), _("Converting Time Codes.\nPlease wait...."))
        # Go to the beginning of the transcript
        self.GotoPos(0)
        # Define a Regular Expression to find "(H:MM:SS.hh)" format time code information
        regex = "([0-9]+:[0-5][0-9]:[0-5][0-9].[0-9]+)"
        # Get the String (plain text) value of the contents of the editor
        transcriptText = self.GetValue()
        # Execute the Regular Expression so we can iterate through the results
        regexResults = re.finditer(regex, transcriptText)
        # For each result found by the Regular Expression search ...
        for regexResult in regexResults:
            # ... Get the string value for the time, in HH:MM:SS.hh format, excluding the surrounding parentheses
            tcString = transcriptText[regexResult.start() + 1 : regexResult.end() - 1]
            # Convert the string to a Time Code value in milliseconds
            tcVal = Misc.time_in_str_to_ms(tcString)
            # Find the text in the Transcript that matches the current Regular Expression result
            self.find_text(transcriptText[regexResult.start() : regexResult.end()], 'next')
            # If we are looking at a value larger than the last Time Code entered ...
            if (len(self.timecodes) == 0) or (tcVal > self.timecodes[-1]):
                # ... delete the current selection, which is the (H:MM:SS.hh) string 
                self.DeleteSelection()
                # ... insert the new Time Code in Transana Format
                self.InsertTimeCode(tcVal)
                # ... add the new Time Code to the Time Codes list
                self.timecodes.append(tcVal)
        # Go to the beginning of the transcript
        self.GotoPos(0)
        # Destroy the popup
        convertDlg.Destroy()

    def IsStyleHiddenAt(self, pos):
        """ Is the style at the specified position hidden? """
        # Get the style at the specified position
        textAttr = self.GetStyleAt(pos)
        # and compare it to the Hidden Time Code Data style
        return self.CompareFormatting(textAttr, self.txtHiddenAttr, fullCompare=False)

    def CallFormatDialog(self, tabToShow=0):
        """ Trigger the Format Dialog, either updating the font settings for the selected text or
            changing the the font settingss for the current cursor position. """
        # Let's try to remember the cursor position, getting the first character of the first and last lines
        firstChar = self.GetFirstVisiblePosition()
        # If we're using wxPython 2.8.x.x ...
        if wx.VERSION[:2] == (2, 8):
            # ... use HitTest()
            lastChar = self.HitTest(wx.Point(5, self.GetSize()[1] - 10))[1]
        # If we're using a later wxPython version ...
        else:
            # ... use HitTestPos()
            lastChar = self.HitTestPos(wx.Point(5, self.GetSize()[1] - 10))[1]

        # Set the Wait cursor
        self.parent.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))

        # Create a RichTextAttr object to hold the style
        currentStyle = richtext.RichTextAttr()

        # If we don't have a text selection, we can just get the wxFontData and go.
        currentSelection = self.GetSelection()

        # If we have NO selection ...
        if currentSelection[0] == currentSelection[1]:
            # If we're not at the beginning of the document ...
            if self.GetCaretPosition() >= 0:
                # ... let's use the Caret Position
                pos = self.GetCaretPosition()
            # Otherwise ...
            else:
                # ... let's use the Insertion Point
                pos = self.GetInsertionPoint()

            # Get the current style
            currentStyle = self.GetStyleAt(pos)

            # We need to select only VALID formatting for the format dialog.  Therefore, we need to skip
            # time codes, hidden text, and human readable time info.
            while ((self.GetCharAt(pos) == 10) or \
                   (self.GetCharAt(pos) == ord(TIMECODE_CHAR)) or \
                   self.CompareFormatting(currentStyle, self.txtTimeCodeAttr, fullCompare=False) or \
                   self.CompareFormatting(currentStyle, self.txtHiddenAttr, fullCompare=False) or \
                   self.CompareFormatting(currentStyle, self.txtTimeCodeHRFAttr, fullCompare=False)) and \
                  (pos < self.GetLastPosition()) or self.IsStyleHiddenAt(pos + 1):
                pos += 1
                currentStyle = self.GetStyleAt(pos)                
        # If we have a selection ...
        else:
            # ... get the style for the selection
            self.GetStyleForRange(self.GetSelectionRange(), currentStyle)

        # Get the font associated with that style
        currentFont = currentStyle.GetFont()
        # If there's a problem with that font ...
        if not currentFont.IsOk():
            # ... then load the Default Font
            currentFont = self.GetDefaultStyle().GetFont()

        # First, get the initial values for the Font Dialog.  This will match the
        # formatting of the LAST character in the selection.
        fontData = FormatDialog.FormatDef()
        fontData.fontFace = currentFont.GetFaceName()
        fontData.fontSize = currentFont.GetPointSize()
        if currentFont.GetWeight() == wx.FONTWEIGHT_BOLD:
            fontData.fontWeight = FormatDialog.fd_BOLD
        else:
            fontData.fontWeight = FormatDialog.fd_OFF
        if currentFont.GetStyle() == wx.FONTSTYLE_ITALIC:
            fontData.fontStyle = FormatDialog.fd_ITALIC
        else:
            fontData.fontStyle = FormatDialog.fd_OFF
        if currentFont.GetUnderlined():
            fontData.fontUnderline = FormatDialog.fd_UNDERLINE
        else:
            fontData.fontUnderline = FormatDialog.fd_OFF
        fontData.fontColorDef = currentStyle.GetTextColour()
        fontData.fontBackgroundColorDef = currentStyle.GetBackgroundColour()

        fontData.paragraphAlignment = currentStyle.GetAlignment()
        fontData.paragraphLeftIndent = currentStyle.GetLeftIndent()
        fontData.paragraphLeftSubIndent = currentStyle.GetLeftSubIndent()
        fontData.paragraphRightIndent = currentStyle.GetRightIndent()
        fontData.paragraphLineSpacing = currentStyle.GetLineSpacing()
        fontData.paragraphSpaceBefore = currentStyle.GetParagraphSpacingBefore()
        fontData.paragraphSpaceAfter = currentStyle.GetParagraphSpacingAfter()
        fontData.tabs = currentStyle.GetTabs()

        # Now we need to iterate through the selection and look for any characters with different font values.
        for selPos in range(currentSelection[0], currentSelection[1]):
            # Get the Font Attributes of the current Character
            tmpStyle = self.GetStyleAt(selPos)
            tmpFont = tmpStyle.GetFont()

            if tmpFont.IsOk():
                isTC = (self.CompareFormatting(tmpStyle, self.txtTimeCodeAttr, fullCompare=False)) or \
                       (self.CompareFormatting(tmpStyle, self.txtHiddenAttr, fullCompare=False)) or \
                       (self.CompareFormatting(tmpStyle, self.txtTimeCodeHRFAttr, fullCompare=False))
            else:
                isTC = False

            # We don't touch the settings for TimeCodes or Hidden TimeCode Data, so these characters can be ignored.
            # Also check that the tmpFont is valid, or we can't do the comparisons.
            if (not isTC) and tmpFont.IsOk():
                
                # Now look for specs that are different, and flag the TransanaFontDef object if one is found.
                # If the the Symbol Font is used, we ignore this.  (We don't want to change the Font Face of Special Characters.)
                if (fontData.fontFace != None) and (tmpFont.GetFaceName() != 'Symbol') and (tmpFont.GetFaceName() != fontData.fontFace):
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.fontFace)
                    fontData.fontFace = None

                if (fontData.fontSize != None) and (tmpFont.GetPointSize() != fontData.fontSize):
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.fontSize)
                    fontData.fontSize = None

                if (fontData.fontWeight != FormatDialog.fd_AMBIGUOUS) and \
                   ((tmpFont.GetWeight() == wx.FONTWEIGHT_BOLD) and (fontData.fontWeight == FormatDialog.fd_OFF)) or \
                   ((tmpFont.GetWeight() != wx.FONTWEIGHT_BOLD) and (fontData.fontWeight == FormatDialog.fd_BOLD)):
                    fontData.fontWeight = FormatDialog.fd_AMBIGUOUS

                if (fontData.fontStyle != FormatDialog.fd_AMBIGUOUS) and \
                   ((tmpFont.GetStyle() == wx.FONTSTYLE_ITALIC) and (fontData.fontStyle == FormatDialog.fd_OFF)) or \
                   ((tmpFont.GetStyle() != wx.FONTSTYLE_ITALIC) and (fontData.fontStyle == FormatDialog.fd_ITALIC)):
                    fontData.fontStyle = FormatDialog.fd_AMBIGUOUS

                if (fontData.fontUnderline != FormatDialog.fd_AMBIGUOUS) and \
                   ((tmpFont.GetUnderlined()) and (fontData.fontUnderline == FormatDialog.fd_OFF)) or \
                   ((not tmpFont.GetUnderlined()) and (fontData.fontUnderline == FormatDialog.fd_UNDERLINE)):
                    fontData.fontUnderline = FormatDialog.fd_AMBIGUOUS

                if (fontData.fontColorDef != None) and (fontData.fontColorDef != tmpStyle.GetTextColour()):
                    fontData.fontColorName = ''
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.fontColorDef)
                    fontData.fontColorDef = None
        
                if (fontData.fontBackgroundColorDef != None) and (fontData.fontBackgroundColorDef != tmpStyle.GetBackgroundColour()):
                    fontData.fontBackgroundColorName = ''
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.fontBackgroundColorDef)
                    fontData.fontBackgroundColorDef = None

                if (fontData.paragraphAlignment != None) and (fontData.paragraphAlignment != tmpStyle.GetAlignment()):
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.paragraphAlignment)
                    fontData.paragraphAlignment = None

                if (fontData.paragraphLeftIndent != None) and (fontData.paragraphLeftIndent != tmpStyle.GetLeftIndent()):
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.paragraphLeftIndent)
                    fontData.paragraphLeftIndent = None

                if (fontData.paragraphLeftSubIndent != None) and (fontData.paragraphLeftSubIndent != tmpStyle.GetLeftSubIndent()):
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.paragraphLeftSubIndent)
                    fontData.paragraphLeftSubIndent = None

                if (fontData.paragraphRightIndent != None) and (fontData.paragraphRightIndent != tmpStyle.GetRightIndent()):
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.paragraphRightIndent)
                    fontData.paragraphRightIndent = None

                if (fontData.paragraphLineSpacing != None) and (fontData.paragraphLineSpacing != tmpStyle.GetLineSpacing()):
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.paragraphLineSpacing)
                    fontData.paragraphLineSpacing = None

                if (fontData.paragraphSpaceBefore != None) and (fontData.paragraphSpaceBefore != tmpStyle.GetParagraphSpacingBefore()):
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.paragraphSpaceBefore)
                    fontData.paragraphSpaceBefore = None

                if (fontData.tabs != []) and (fontData.tabs != tmpStyle.GetTabs()):
                    # These should be functionally equivalent, but apparently not!  Use the second to avoid problems
                    # when you've imported a file with font problems
                    # del(fontData.tabs)
                    fontData.tabs = []

        # If there's no SELECTION, paragraph formatting may not work right.  This is especially true if we are on
        # a blank line.  This code attempts for fix these problems.

        # If there's no selection ...
        if currentSelection[0] == currentSelection[1]:
            # note the current insertion point ...
            ip = self.GetInsertionPoint()
            # Insert a space into the text ...
            self.WriteText(' ')
            # ... and select that space so that we're formatting SOMETHING.
            # Since the space is selected, it will be over-written by the first character the user types.  Most users
            # won't even notice this, I suspect.
            self.SetSelection(ip, ip+1)
        # If there IS a selection ...
        else:
            # ... set the insertion point indicator to None
            ip = None

        # Set the cursor back to normal
        self.parent.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

        originalFontData = fontData.copy()
        # Make a COPY of the font to change.  If you don't, changes to a selection seem to affect the whole paragraph!!
        fontToChange = fontData.copy()

        # Create the Format Dialog.
        # Note:  We used to use the wx.FontDialog, but this proved inadequate for a number of reasons.
        #        It offered very few font choices on the Mac, and it couldn't handle font ambiguity.
        formatDialog = FormatDialog.FormatDialog(self, fontToChange, tabToShow)

        self.Freeze()

        # Display the FormatDialog and get the user feedback
        if formatDialog.ShowModal() == wx.ID_OK:
            # Begin an Undo Batch for formatting
            self.BeginBatchUndo('FormatDialog')
            # If we don't have a text selection, we can just update the current font settings.
            if currentSelection[0] == currentSelection[1]:
                # Get the wxFontData from the Font Dialog
                newFontDef = formatDialog.GetFormatDef()

#                print "TranscriptEditor_RTC.CallFormatDialog():"
#                self.PrintTextAttr("Before:", currentStyle)
#                print "AFTER:", newFontDef
#                print

##                # Set the current font
##                self.SetTxtStyle(fontColor = newFontDef.fontColorDef,
##                                 fontBgColor = newFontDef.fontBackgroundColorDef,
##                                 fontFace = newFontDef.fontFace,
##                                 fontSize = newFontDef.fontSize,
##                                 fontBold = (newFontDef.fontWeight != 0),
##                                 fontItalic = (newFontDef.fontStyle != 0),
##                                 fontUnderline = newFontDef.fontUnderline,
##                                 parAlign = newFontDef.paragraphAlignment,
##                                 parLeftIndent = (newFontDef.paragraphLeftIndent, newFontDef.paragraphLeftSubIndent),
##                                 parRightIndent = newFontDef.paragraphRightIndent,
##                                 parLineSpacing = newFontDef.paragraphLineSpacing,
##                                 parSpacingBefore = newFontDef.paragraphSpaceBefore,
##                                 parSpacingAfter = newFontDef.paragraphSpaceAfter,
##                                 parTabs = newFontDef.tabs)

                # Problem:  Changing font and paragraph styles at the same time applies the font change to
                #           the whole paragraph when it should only be applied to the insertion point.

                # Solution:  Separate out the formatting calls and only apply those that change.

                # Hypothesis:  Changing all the Paragraph characteristics BEFORE moving onto the font
                #              characteristics may also help.

                if newFontDef.paragraphAlignment != currentStyle.GetAlignment():
                    self.SetTxtStyle(parAlign = newFontDef.paragraphAlignment)

                if (newFontDef.paragraphLeftIndent != currentStyle.GetLeftIndent()) or \
                   (newFontDef.paragraphLeftSubIndent != currentStyle.GetLeftSubIndent()):
                    self.SetTxtStyle(parLeftIndent = (newFontDef.paragraphLeftIndent, newFontDef.paragraphLeftSubIndent))

                if newFontDef.paragraphRightIndent != currentStyle.GetRightIndent():
                    self.SetTxtStyle(parRightIndent = newFontDef.paragraphRightIndent)

                if newFontDef.paragraphLineSpacing != currentStyle.GetLineSpacing():
                    self.SetTxtStyle(parLineSpacing = newFontDef.paragraphLineSpacing)

                if newFontDef.paragraphSpaceBefore != currentStyle.GetParagraphSpacingBefore():
                    self.SetTxtStyle(parSpacingBefore = newFontDef.paragraphSpaceBefore)

                if newFontDef.paragraphSpaceAfter != currentStyle.GetParagraphSpacingAfter():
                    self.SetTxtStyle(parSpacingAfter = newFontDef.paragraphSpaceAfter)

                if newFontDef.tabs != currentStyle.GetTabs():
                    self.SetTxtStyle(parTabs = newFontDef.tabs)

                if newFontDef.fontFace != currentFont.GetFaceName():
                    self.SetTxtStyle(fontFace = newFontDef.fontFace)

                if newFontDef.fontSize != currentFont.GetPointSize():
                    self.SetTxtStyle(fontSize = newFontDef.fontSize)

                if newFontDef.fontWeight != (currentFont.GetWeight() == wx.FONTWEIGHT_BOLD):
                    self.SetTxtStyle(fontBold = (newFontDef.fontWeight != 0))

                if newFontDef.fontStyle != (currentFont.GetStyle() == wx.FONTSTYLE_ITALIC):
                    self.SetTxtStyle(fontItalic = (newFontDef.fontStyle != 0))

                if newFontDef.fontUnderline != currentFont.GetUnderlined():
                    self.SetTxtStyle(fontUnderline = newFontDef.fontUnderline)

                if newFontDef.fontColorDef != currentStyle.GetTextColour():
                    self.SetTxtStyle(fontColor = newFontDef.fontColorDef)

                if newFontDef.fontBackgroundColorDef != currentStyle.GetBackgroundColour():
                    self.SetTxtStyle(fontBgColor = newFontDef.fontBackgroundColorDef)

            else:
                # Set the Wait cursor
                self.parent.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
                # If the selection is 1000 characters or more ...
                if (currentSelection[1] - currentSelection[0] > 1000):
                    # ... create a popup telling the user about the formatting (needed for large selections)
                    formatDlg = Dialogs.PopupDialog(None, _("Formatting..."), _("Formatting your transcript selection.\nPlease wait...."))
                else:
                    formatDlg = None

                # NEW MODEL -- Only update those attributes not flagged as ambiguous.  This is necessary
                # when processing a selection
                # Get the TransanaFontDef data from the Font Dialog.
                newFontData = formatDialog.GetFormatDef()

                # Now we need to iterate through the selection and update the font information.
                # It doesn't work to try to apply formatting to the whole block, as ambiguous attributes
                # lose their values.

                # It appears that we need to separate FONT attributes from PARAGRAPH atttibutes to prevent
                # font attributes from being applied to more than just the current selection when both
                # types of attributes are updated at once in wxPython 2.9.5.0

                # Create a Text Attribute Object for FONT attributes
                tmpFontAttr = richtext.RichTextAttr()
                # Create a Text Attribute Object for the Paragraph attributes
                tmpParagraphAttr = richtext.RichTextAttr()

                # We also need to track whether either of these types of attributes gets changed.
                fontAttrChanged = False
                paragraphAttrChanged = False

                # Initialize a variable to hold the last style, so we can see if style has changed
                lastStyle = None
                # See if Underline was turned ON
                setUnderline = (newFontData.fontUnderline == FormatDialog.fd_UNDERLINE)

                if DEBUG:
                    print "TranscriptEditor_RTC.CallFormatDialog()", currentSelection, setUnderline
                    print newFontData
                    print

                try:
                    if DEBUG:
                        print "TranscriptEditor_RTC.CallFormatDialog():"
                        print "Selection:", currentSelection
                        self.PrintTextAttr("initial FONT attributes", tmpFontAttr)
                        self.PrintTextAttr("initial PARAGRAPH attributes", tmpParagraphAttr)
                        print

                    # For each character position in the current selection ...
                    for selPos in range(currentSelection[0], currentSelection[1]):

                        # Get the Font Attributes of the current Character
                        tmpStyle = self.GetStyleAt(selPos)
                        tmpFont = tmpStyle.GetFont()

                        if tmpFont.IsOk():
                            isTC = (self.CompareFormatting(tmpStyle, self.txtTimeCodeAttr, fullCompare=False)) or \
                                   (self.CompareFormatting(tmpStyle, self.txtHiddenAttr, fullCompare=False)) or \
                                   (self.CompareFormatting(tmpStyle, self.txtTimeCodeHRFAttr, fullCompare=False))
                        else:
                            isTC = False

                        # if this is the LAST character of a selection AND
                        # if we are at the first character of a new block AND
                        # if this character is NOT a time code AND
                        # if the font is valid 
                        # (then we have a special circumstance where we need to do this at the START if the
                        #  conditional block or else the style change never gets applied!)
                        if (selPos == currentSelection[1] - 1) and\
                           (lastStyle is None) and \
                           (not isTC) and \
                           tmpFont.IsOk():
                            # Note the new style
                            lastStyle = tmpStyle
                            # Note the new Font
                            lastFont = tmpFont
                            # Note the starting Position
                            updateStart = selPos

                        # if we don't already have a lastStyle defined ... (we are at the first character of a new block)
                        if (not lastStyle is None):
                            # if we are at a Time Code OR
                            # if the formatting is changing OR
                            # if we're at the end of our selection OR
                            # if we have run up against an invalid font ...
                            if isTC or \
                               not self.CompareFormatting(lastStyle, tmpStyle, fullCompare=True) or \
                               (selPos == currentSelection[1] - 1) or \
                               (not tmpFont.IsOk()):

                            # Now alter those characteristics of the previous block that are not ambiguous in the newFontData.
                            # Where the specification is ambiguous, use the old value from attrs.

                                # We don't want to change the font of special symbols!  Therefore, we don't change
                                # the font name for anything in Symbol font.
                                if (newFontData.fontFace != None) and \
                                   (newFontData.fontFace != '') and \
                                   (lastFont.GetFaceName() != 'Symbol') and \
                                   (newFontData.fontFace != originalFontData.fontFace):
                                    tmpFontAttr.SetFontFaceName(newFontData.fontFace)
                                    fontAttrChanged = True

                                    if DEBUG:
                                        print "Set font to:", newFontData.fontFace
                                    
                                if (newFontData.fontSize != None) and \
                                   (newFontData.fontSize != originalFontData.fontSize):
                                    tmpFontAttr.SetFontSize(newFontData.fontSize)
                                    fontAttrChanged = True

                                    if DEBUG:
                                        print "Set font size to:", newFontData.fontSize
                                    
                                if (newFontData.fontWeight != originalFontData.fontWeight):
                                    if newFontData.fontWeight == FormatDialog.fd_BOLD:
                                        tmpFontAttr.SetFontWeight(wx.FONTWEIGHT_BOLD)
                                        fontAttrChanged = True

                                        if DEBUG:
                                            print "Set Bold to: BOLD"
                                        
                                    elif newFontData.fontWeight == FormatDialog.fd_OFF:
                                        tmpFontAttr.SetFontWeight(wx.FONTWEIGHT_NORMAL)
                                        fontAttrChanged = True

                                        if DEBUG:
                                            print "Set Bold to: normal"
                                    
                                if (newFontData.fontStyle != originalFontData.fontStyle):
                                    if newFontData.fontStyle == FormatDialog.fd_ITALIC:
                                        tmpFontAttr.SetFontStyle(wx.FONTSTYLE_ITALIC)
                                        fontAttrChanged = True

                                        if DEBUG:
                                            print "Set Italics to Italics"
                                        
                                    elif newFontData.fontStyle == FormatDialog.fd_OFF:
                                        tmpFontAttr.SetFontStyle(wx.FONTSTYLE_NORMAL)
                                        fontAttrChanged = True

                                        if DEBUG:
                                            print "Set Italics to normal"

                                if (newFontData.fontUnderline != originalFontData.fontUnderline):
                                    if newFontData.fontUnderline == FormatDialog.fd_UNDERLINE:
                                        tmpFontAttr.SetFontUnderlined(True)
                                        fontAttrChanged = True

                                        if DEBUG:
                                            print "Turn Underline ON"
                                        
                                    elif newFontData.fontUnderline == FormatDialog.fd_OFF:
                                        tmpFontAttr.SetFontUnderlined(False)
                                        fontAttrChanged = True

                                        if DEBUG:
                                            print "Turn Underline off"
                                        
                                if (newFontData.fontColorDef != originalFontData.fontColorDef):
                                    if newFontData.fontColorDef != None:
                                        color = newFontData.fontColorDef
                                        # Now apply the font settings for the current character
                                        tmpFontAttr.SetTextColour(color)
                                        fontAttrChanged = True

                                        if DEBUG:
                                            print "Color changed to", newFontData.fontColorName

                                if (newFontData.fontBackgroundColorDef != originalFontData.fontBackgroundColorDef):
                                    if newFontData.fontBackgroundColorDef != None:
                                        color = newFontData.fontBackgroundColorDef
                                        # Now apply the font settings for the current character
                                        tmpFontAttr.SetBackgroundColour(color)
                                        fontAttrChanged = True

                                        if DEBUG:
                                            print "Background Color changed to", newFontData.fontBackgroundColorDef

                                if (newFontData.paragraphAlignment != originalFontData.paragraphAlignment):
                                    if newFontData.paragraphAlignment != None:
                                        tmpParagraphAttr.SetAlignment(newFontData.paragraphAlignment)
                                        paragraphAttrChanged = True

                                        if DEBUG:
                                            print "Set paragraph alignment to:", newFontData.paragraphAlignment

                                tmpLeftSubIndent = originalFontData.paragraphLeftSubIndent
                                tmpLeftIndent = originalFontData.paragraphLeftIndent
                                if (newFontData.paragraphLeftSubIndent != originalFontData.paragraphLeftSubIndent):
                                    if newFontData.paragraphLeftSubIndent != None:
                                        tmpLeftSubIndent = newFontData.paragraphLeftSubIndent
                                        tmpParagraphAttr.SetLeftIndent(tmpLeftIndent, tmpLeftSubIndent)
                                        paragraphAttrChanged = True

                                        if DEBUG:
                                            print "Set paragraph left sub-indent to:", newFontData.paragraphLeftSubIndent

                                if (newFontData.paragraphLeftIndent != originalFontData.paragraphLeftIndent):
                                    if newFontData.paragraphLeftIndent != None:
                                        tmpLeftIndent = newFontData.paragraphLeftIndent
                                        tmpParagraphAttr.SetLeftIndent(tmpLeftIndent, tmpLeftSubIndent)
                                        paragraphAttrChanged = True

                                        if DEBUG:
                                            print "Set paragraph left indent to:", tmpLeftIndent, tmpLeftSubIndent

                                if (newFontData.paragraphRightIndent != originalFontData.paragraphRightIndent):
                                    if newFontData.paragraphRightIndent != None:
                                        tmpParagraphAttr.SetRightIndent(newFontData.paragraphRightIndent)
                                        paragraphAttrChanged = True

                                        if DEBUG:
                                            print "Set paragraph right indent to:", newFontData.paragraphRightIndent

                                if (newFontData.paragraphLineSpacing != originalFontData.paragraphLineSpacing):
                                    if newFontData.paragraphLineSpacing != None:
                                        tmpParagraphAttr.SetLineSpacing(newFontData.paragraphLineSpacing)
                                        paragraphAttrChanged = True

                                        if DEBUG:
                                            print "Set paragraph line spacing to:", newFontData.paragraphLineSpacing
                                        
                                if (newFontData.paragraphSpaceBefore != originalFontData.paragraphSpaceBefore):
                                    if newFontData.paragraphSpaceBefore != None:
                                        tmpParagraphAttr.SetParagraphSpacingBefore(newFontData.paragraphSpaceBefore)
                                        paragraphAttrChanged = True

                                        if DEBUG:
                                            print "Set paragraph spacing before:", newFontData.paragraphSpaceBefore

                                if (round(newFontData.paragraphSpaceAfter) != originalFontData.paragraphSpaceAfter):
                                    if newFontData.paragraphSpaceAfter != None:
                                        tmpParagraphAttr.SetParagraphSpacingAfter(newFontData.paragraphSpaceAfter)
                                        paragraphAttrChanged = True

                                        if DEBUG:
                                            print "Set Paragraph Space After:", newFontData.paragraphSpaceAfter, originalFontData.paragraphSpaceAfter
                                        
                                if (newFontData.tabs != originalFontData.tabs):
                                    if newFontData.tabs != None:
                                        tmpParagraphAttr.SetTabs(newFontData.tabs)
                                        paragraphAttrChanged = True

                                        if DEBUG:
                                            print "Set tabs to:", newFontData.tabs, originalFontData.tabs
                                        
                                if selPos == currentSelection[1] - 1:
                                    updateEnd = selPos + 1
                                else:
                                    updateEnd = selPos
                                
                                # Apply the style to the previous block
                                if fontAttrChanged:
                                    self.SetStyle((updateStart, updateEnd), tmpFontAttr)

                                    if DEBUG:
                                        self.PrintTextAttr("Setting Font: (%d, %d)" % (updateStart, updateEnd), tmpFontAttr)
                                        
                                if paragraphAttrChanged:
                                    self.SetStyle((updateStart, updateEnd), tmpParagraphAttr)

                                    if DEBUG:
                                        self.PrintTextAttr("Setting Paragraph: (%d, %d)" % (updateStart, updateEnd), tmpParagraphAttr)
                                
                                # Since we've just applied the style, we can reset lastStyle to None
                                lastStyle = None

                        # if we are at the first character of a new block AND
                        # this character is NOT a time code AND
                        # the font is valid ...
                        if (lastStyle is None) and (not isTC) and tmpFont.IsOk():
                            # Note the new style
                            lastStyle = tmpStyle
                            # Note the new Font
                            lastFont = tmpFont
                            # Note the starting Position
                            updateStart = selPos

                except:

                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                # There's a bug in wxWidgets on Windows such that underlining doesn't always show up correctly.
                # The following code changes the window size slightly so that a re-draw is forced, thus
                # causing the underlining to show up.  Update() and Refresh() don't work.

                # if we're turning Underline ON ...
                if setUnderline:
                    # If we're on Windows, resize the parent window to force the redraw ...
                    if ('wxMSW' in wx.PlatformInfo):

                        # ... find the size of the parent window
                        size = self.parent.GetSizeTuple()
                        # Move the Insertion Point to the end of the selection (or this won't work!)
                        self.SetInsertionPoint(currentSelection[1])
                        # Shrink the parent window slightly
                        self.parent.SetSize((size[0], size[1] - 5))
                        # Set the Parent Window back to the original size
                        self.parent.SetSize(size)

                    if DEBUG:
                        print "surroundingUnderline correction DISABLED"

                # Mark the transcript as changed.
                self.MarkDirty()

                if (currentSelection[1] - currentSelection[0] > 1000):
                    # Create a popup telling the user about the load (needed for large files)
                    if formatDlg:
                        formatDlg.Destroy()

                # Set the cursor back to normal
                self.parent.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

            # End the Formatting Undo Batch
            self.EndBatchUndo()

        # Destroy the Font Dialog Box, now that we're done with it.
        formatDialog.Destroy()

        # Enable control updating now that all changes have been processed
        self.Thaw()

        if self.StyleChanged != None:
            self.StyleChanged(self)

        # If we added a space for formatting purposes ...
        if ip != None:
            # ... we need to get rid of the space
            self.DeleteSelection()

        # Let's try restoring the screen position when all is said and done.  For unknown reasons, working with fonts scrolls
        # the control
        if self.GetFirstVisiblePosition() != firstChar:
            # The ShowPosition calls get over-ridden somewhere, so using wx.CallAfter is necessary so they'll be called last.
            # First, show the first character
            wx.CallLater(100, self.ShowPosition, firstChar)
            # Since we generally see scrolling UP, we need to show the LAST Character to restore the screen position.
            wx.CallLater(125, self.ShowPosition, lastChar)

        # We've probably taken the focus from the editor.  Let's return it.
        self.SetFocus()

    def InsertImage(self, imgFile = None):
        """ Insert an image from a file into the Transcript """
        # If no image name is passed, prompt the user to select an image file
        if imgFile == None:
            # Create a dialog box for requesting an image file
            dlg = wx.FileDialog(self, wildcard=TransanaConstants.imageFileTypesString, style=wx.OPEN)
            # Set the File Filter to acceptable graphics types
            dlg.SetFilterIndex(1)
            # Get a file selection from the user
            if dlg.ShowModal() == wx.ID_OK:
                # If the user clicks OK, get the name of the selected file
                imgFile = dlg.GetPath()
            # Destroy the File Dialog
            dlg.Destroy()
        # If an image file name is provided one way or the other ...
        if imgFile != None:
            # Create a wxImage from the file
            image = wx.Image(imgFile)
            # If the image imported okay ...
            if image.IsOk():
                # Get the original dimensions of the image
                imgWidth = image.GetWidth()
                imgHeight = image.GetHeight()
                if TransanaGlobal.configData.maxTranscriptImageWidth == 1:
                    # We need the SMALLER of the current image size and the current Transcript Window size
                    # (Adjust width for scrollbar size!)
                    maxWidth = min(520, float(imgWidth), (self.GetSize()[0] - 20.0) * 0.98)
                else:
                    # We need the SMALLER of the current image size and the current Transcript Window size
                    # (Adjust width for scrollbar size!)
                    maxWidth = min(float(imgWidth), (self.GetSize()[0] - 20.0) * 0.98)
                # It doesn't make sense to limit the image's height.
                # When we have multiple transcripts, window heights can be VERY small!
                maxHeight = imgHeight  # min(float(imgHeight), self.GetSize()[1] * 0.98)
                # Determine the scaling factor for adjusting the image size
                scaleFactor = min(maxWidth / float(imgWidth), maxHeight / float(imgHeight))
                # If the image needs to be re-scaled ...
                if scaleFactor != 1.0:
                    # ... rescale the image to fit in the current Transcript window.  Use slower high quality rescale.
                    image.Rescale(int(imgWidth * scaleFactor), int(imgHeight * scaleFactor), quality=wx.IMAGE_QUALITY_HIGH)
                # Add the image to the transcript
                self.WriteImage(image)

        # Check to see if the transcript has exceeded the transcript object maximum size.
        if len(self.GetFormattedSelection('XML')) > TransanaGlobal.max_allowed_packet - 80000:   # 8300000
            # If so, generate an error message.
            prompt = _("Inserting this image caused the transcript to exceed Transana's maximum transcript size.")
            prompt += '\n' + _("You may need to use fewer images or break this transcript into multiple parts.")
            prompt += '\n' + _("The image will be removed from the trancript.") + "  " + _("You should save your transcript immediately.")

            errDlg = Dialogs.ErrorDialog(self, prompt)
            errDlg.ShowModal()
            errDlg.Destroy()
            
            # Remove the image that caused the overflow.
            self.Undo()


class TranscriptEditorDropTarget(wx.PyDropTarget):
    """ Required to make the TranscriptEditor Control a viable Drop Target """
    
    def __init__(self, editor):
        """ Initialize Drop Target functionality """
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
        if (self.editor.TranscriptObj != None) and self.GetData():
            # Extract actual data passed by DataTreeDropSource
            sourceData = cPickle.loads(self.data.GetData())
            # With the addition of multi-select in the Database Tree, we might get a single item here in the sourceData,
            # or we might get a list of node data elements.  Let's create a List for all elements to go into (to keep
            # everything parallel)
            # Create an empty list to hold the sourceData items
            sourceList = []
            # If we have received a list ...
            if isinstance(sourceData, list):
                # ... add each element from sourceData ...
                for element in sourceData:
                    # ... to our sourceList list
                    sourceList.append(element)
                # ... and catch the FIRST element as if it were a single-item sourceData object
                sourceData = sourceList[0]
            # If we have received a single element from sourceData ...
            else:
                # ... add that element to our sourceList list
                sourceList.append(sourceData)

            # If a Keyword Node is dropped ...
            if sourceData.nodetype == 'KeywordNode':
                # See if we're creating a QuickClip
                if (TransanaConstants.macDragDrop or ('wxMac' in wx.PlatformInfo)) and TransanaGlobal.configData.quickClipMode:
                    # Get the clip start and stop times from the transcript
                    (startTime, endTime) = self.editor.get_selected_time_range()
                    # Initialize the Keyword List
                    kwList = []
                    # Iterate through the sourceLIst list ...
                    for element in sourceList:
                        # ... and extract the keyword data for each element
                        kwList.append((element.parent, element.text))
                    # Determine whether we're creating a Clip from an Episode Transcript
                    if self.editor.TranscriptObj.clip_num == 0:
                        # If so, get the Episode data
                        transcriptNum = self.editor.TranscriptObj.number
                        episodeNum = self.editor.TranscriptObj.episode_num
                        # If thre's no specified end time ...
                        if endTime <= 0:
                            # ... then the end of the Episode is the end of the clip!
                            endTime = self.editor.parent.ControlObject.currentObj.tape_length
                    # If we are not working from an Episode, we're working from a Clip and are sub-clipping.
                    else:
                        # Get the source Clip data.  We can skip loading the Clip Transcript to save load time
                        tempClip = Clip.Clip(self.editor.TranscriptObj.clip_num, skipText=True)
                        # This gets the CORRECT Trancript Record, even with multiple transcripts
                        transcriptNum = self.editor.TranscriptObj.source_transcript
                        episodeNum = tempClip.episode_num
                        # If the sub-clip starts at the start of the parent clip ... 
                        if startTime == 0:
                            # ... we need to grab the parent clip's start time
                            startTime = tempClip.clip_start
                        # if the sub-clip ends at the end of the parent clip ...
                        if endTime <= 0:
                            # ... we need to grab the parent clip's end time
                            endTime = tempClip.clip_stop
                    # If the text selection is blank, we need to send a blank rather than RTF for nothing
                    (startPos, endPos) = self.editor.GetSelection()
                    if startPos == endPos:
                        text = ''
                    else:
                        text = self.editor.GetFormattedSelection('XML', selectionOnly=True)  # GetXMLBuffer(select_only=1)
                    # Get the Clip Data assembled for creating the Quick Clip
                    clipData = DragAndDropObjects.ClipDragDropData(transcriptNum, episodeNum, startTime, endTime, text, self.editor.GetStringSelection(), self.editor.parent.ControlObject.GetVideoCheckboxDataForClips(startTime))
                    # I'm sure this is horrible form, but I don't know how else to do this from here!
                    dbTree = self.editor.parent.ControlObject.DataWindow.DBTab.tree
                    # Create the Quick Clip
                    DragAndDropObjects.CreateQuickClip(clipData, sourceData.parent, sourceData.text, dbTree, extraKeywords=kwList[1:])
                # If we're NOT creating a Quick Clip, we're adding a keyword to the current Episode or Clip
                else:
                    # Determine where the Transcript was loaded from
                    if self.editor.TranscriptObj:
                        # If we've got a Clip loaded ...
                        if self.editor.TranscriptObj.clip_num != 0:
                            # ... assemble Clip data
                            targetType = 'Clip'
                            targetLabel = _('Clip')
                            targetRecNum = self.editor.TranscriptObj.clip_num
                            # We can skip loading the Clip Transcript to save load time
                            clipObj = Clip.Clip(targetRecNum, skipText=True)
                            targetName = clipObj.id
                        # Otherwise, we have an Episode loaded ...
                        else:
                            # ... so assemble Episode data
                            targetType = 'Episode'
                            targetLabel = _('Episode')
                            targetRecNum = self.editor.TranscriptObj.episode_num
                            epObj = Episode.Episode(targetRecNum)
                            targetName = epObj.id

                        # For each keyword in the sourceList list ...
                        for element in sourceList:
                            # ... add the keyword to the current element by simulating a keyword drop on the appropriate target
                            DragAndDropObjects.DropKeyword(self.editor, element, targetType, targetName, targetRecNum, 0, confirmations=True)  # confirmations=(len(sourceList) == 1))
                    else:
                        # No transcript Object loaded, do nothing
                        pass

#            We could create a regular clip here if a Collection were dropped!  (Maybe later.)                    
#            else:
#                print "something other than a Keyword node was dropped"

        return d  # what is returned signals the source what to do
                  # with the original data (move, copy, etc.)  In this
                  # case we just return the suggested value given to us.


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

            # Regular Clips are dropped on Collections, Clips, or Snapshots.  Quick Clips are dropped on Keywords.
            # Text dropped on a Keyword Group can create a Keyword.
            if (self.parent.ControlObject.GetDatabaseTreeTabObjectNodeType() == 'CollectionNode') or \
               (self.parent.ControlObject.GetDatabaseTreeTabObjectNodeType() == 'ClipNode') or \
               (self.parent.ControlObject.GetDatabaseTreeTabObjectNodeType() == 'SnapshotNode') or \
               (self.parent.ControlObject.GetDatabaseTreeTabObjectNodeType() == 'KeywordNode') or\
               (self.parent.ControlObject.GetDatabaseTreeTabObjectNodeType() == 'KeywordGroupNode'):
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
