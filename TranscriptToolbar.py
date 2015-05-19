# Copyright (C) 2003 - 2014 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the TranscriptToolbar class as part of the Editors
component.
"""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx
# import the wxPython RichTextCtrl
import wx.richtext as richtext
# import Transana modules
import Clip
import Dialogs
import Episode
import KeywordListEditForm
import TransanaConstants
import TransanaExceptions
import TransanaGlobal
# Import Transana's Images
import TransanaImages
# import Python's os and sys modules
import os, sys

class TranscriptToolbar(wx.ToolBar):
    """This class consists of a toolbar for use with a TranscriptEditor
    object.  It inherits from the wxToolbar class.  This class is mostly
    self-sustaining, and does not require much in terms of a public
    interface.  Its objects are intended to be created, and then left alone
    until destroyed."""

    def __init__(self, parent, id=-1):
        """Initialize an TranscriptToolbar object."""
        if 'wxGTK' in wx.PlatformInfo:
            size = wx.Size(560, 30)
        else:
            size = wx.Size(470, 30)
        # Create a ToolBar as self
        wx.ToolBar.__init__(self, parent, id, wx.DefaultPosition, size, wx.TB_HORIZONTAL | wx.BORDER_SIMPLE | wx.TB_FLAT | wx.TB_TEXT)
        # remember the parent
        self.parent = parent
        # Set the Toolbar Bitmap size
        self.SetToolBitmapSize((16, 16))
        # Keep a list of the tools placed on the toolbar so they're more easily manipulated
        self.tools = []

        # Create an Undo button
        self.CMD_UNDO_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_UNDO_ID, TransanaImages.Undo16.GetBitmap(),
                        shortHelpString=_('Undo action')))
        wx.EVT_MENU(self, self.CMD_UNDO_ID, self.OnUndo)

        self.AddSeparator()
  
        # Bold, Italic, Underline buttons
        self.CMD_BOLD_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_BOLD_ID, TransanaGlobal.GetImage(TransanaImages.Bold),
                        isToggle=1, shortHelpString=_('Bold text')))
        wx.EVT_MENU(self, self.CMD_BOLD_ID, self.OnBold)

        self.CMD_ITALIC_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_ITALIC_ID, TransanaGlobal.GetImage(TransanaImages.Italic),
                        isToggle=1, shortHelpString=_("Italic text")))
        wx.EVT_MENU(self, self.CMD_ITALIC_ID, self.OnItalic)
       
        self.CMD_UNDERLINE_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_UNDERLINE_ID, TransanaGlobal.GetImage(TransanaImages.Underline),
                        isToggle=1, shortHelpString=_("Underline text")))
        wx.EVT_MENU(self, self.CMD_UNDERLINE_ID, self.OnUnderline)

        self.AddSeparator()

        # Jeffersonian Symbols
        self.CMD_RISING_INT_ID = wx.NewId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_UP, wx.ART_TOOLBAR, (16,16))
        self.tools.append(self.AddTool(self.CMD_RISING_INT_ID, bmp,
                        shortHelpString=_("Rising Intonation")))
        wx.EVT_MENU(self, self.CMD_RISING_INT_ID, self.OnInsertChar)
        
        self.CMD_FALLING_INT_ID = wx.NewId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_DOWN, wx.ART_TOOLBAR, (16,16))
        self.tools.append(self.AddTool(self.CMD_FALLING_INT_ID, bmp,
                        shortHelpString=_("Falling Intonation")))
        wx.EVT_MENU(self, self.CMD_FALLING_INT_ID, self.OnInsertChar) 
       
        self.CMD_AUDIBLE_BREATH_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_AUDIBLE_BREATH_ID, TransanaGlobal.GetImage(TransanaImages.AudibleBreath),
                        shortHelpString=_("Audible Breath")))
        wx.EVT_MENU(self, self.CMD_AUDIBLE_BREATH_ID, self.OnInsertChar)
    
        self.CMD_WHISPERED_SPEECH_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_WHISPERED_SPEECH_ID, TransanaGlobal.GetImage(TransanaImages.WhisperedSpeech),
                        shortHelpString=_("Whispered Speech")))
        wx.EVT_MENU(self, self.CMD_WHISPERED_SPEECH_ID, self.OnInsertChar)
      
        self.AddSeparator()

        # Add show / hide timecodes button
        self.CMD_SHOWHIDE_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_SHOWHIDE_ID, TransanaGlobal.GetImage(TransanaImages.TimeCode16),
                        isToggle=1, shortHelpString=_("Show/Hide Time Code Indexes")))
        wx.EVT_MENU(self, self.CMD_SHOWHIDE_ID, self.OnShowHideCodes)

        # Add show / hide timecodes button
        self.CMD_SHOWHIDETIME_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_SHOWHIDETIME_ID, TransanaGlobal.GetImage(TransanaImages.TimeCodeData16),
                                       TransanaGlobal.GetImage(TransanaImages.TimeCodeData16),
                                       isToggle=1, shortHelpString=_("Show/Hide Time Code Values")))
        wx.EVT_MENU(self, self.CMD_SHOWHIDETIME_ID, self.OnShowHideValues)

        # Add read only / edit mode button
        self.CMD_READONLY_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_READONLY_ID, TransanaGlobal.GetImage(TransanaImages.ReadOnly16),
                                       TransanaGlobal.GetImage(TransanaImages.ReadOnly16),
                                       isToggle=1, shortHelpString=_("Edit/Read-only select")))
        wx.EVT_MENU(self, self.CMD_READONLY_ID, self.OnReadOnlySelect)

        # Add Formatring button
        self.CMD_FORMAT_ID = wx.NewId()
        # ... get the graphic for the Format button ...
        bmp = wx.ArtProvider_GetBitmap(wx.ART_HELP_SETTINGS, wx.ART_TOOLBAR, (16,16))
        # ... and create a Format button on the tool bar.
        self.tools.append(self.AddTool(self.CMD_FORMAT_ID, bmp, shortHelpString=_("Format")))
        wx.EVT_MENU(self, self.CMD_FORMAT_ID, self.OnFormat)

        self.AddSeparator()

        # Add QuickClip button
        self.CMD_QUICKCLIP_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_QUICKCLIP_ID, TransanaGlobal.GetImage(TransanaImages.QuickClip16),
                        shortHelpString=_("Create Quick Clip")))
        wx.EVT_MENU(self, self.CMD_QUICKCLIP_ID, self.OnQuickClip)

        # Add Edit keywords button
        self.CMD_KEYWORD_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_KEYWORD_ID, TransanaGlobal.GetImage(TransanaImages.KeywordRoot16),
                        shortHelpString=_("Edit Keywords")))
        wx.EVT_MENU(self, self.CMD_KEYWORD_ID, self.OnEditKeywords)

        # Add Save Button
        self.CMD_SAVE_ID = wx.NewId()
        self.tools.append(self.AddTool(self.CMD_SAVE_ID, TransanaGlobal.GetImage(TransanaImages.Save16),
                        shortHelpString=_("Save Transcript")))
        wx.EVT_MENU(self, self.CMD_SAVE_ID, self.OnSave)

        self.AddSeparator()

        # Add Propagate Changes Button
        # First, define the ID for this button
        self.CMD_PROPAGATE_ID = wx.NewId()
        # Now create the button and add it to the Tools list
        self.tools.append(self.AddTool(self.CMD_PROPAGATE_ID, TransanaGlobal.GetImage(TransanaImages.Propagate),
                        shortHelpString=_("Propagate Changes")))
        # Link the button to the appropriate event handler
        wx.EVT_MENU(self, self.CMD_PROPAGATE_ID, self.OnPropagate)

        self.AddSeparator()
        
        # Add Multi-Select Button
        # First, define the ID for this button
        self.CMD_MULTISELECT_ID = wx.NewId()
        # Now create the button and add it to the Tools list
        self.tools.append(self.AddTool(self.CMD_MULTISELECT_ID, TransanaGlobal.GetImage(TransanaImages.MultiSelect),
                        shortHelpString=_("Match Selection in Other Transcripts")))
        # Link the button to the appropriate event handler
        wx.EVT_MENU(self, self.CMD_MULTISELECT_ID, self.OnMultiSelect)

        # Add Multiple Transcript Play Button
        # First, define the ID for this button
        self.CMD_PLAY_ID = wx.NewId()
        # Now create the button and add it to the Tools list
        self.tools.append(self.AddTool(self.CMD_PLAY_ID, TransanaImages.Play.GetBitmap(),
                        shortHelpString=_("Play Multiple Transcript Selection")))
        # Link the button to the appropriate event handler
        wx.EVT_MENU(self, self.CMD_PLAY_ID, self.OnMultiPlay)

        self.AddSeparator()

        # SEARCH moved to TranscriptionUI because you can't put a TextCtrl on a Toolbar on the Mac!
        # Set the Initial State of the Editing Buttons to "False"
        for x in (self.CMD_UNDO_ID, self.CMD_BOLD_ID, self.CMD_ITALIC_ID, self.CMD_UNDERLINE_ID, \
                  self.CMD_RISING_INT_ID, self.CMD_FALLING_INT_ID, \
                  self.CMD_AUDIBLE_BREATH_ID, self.CMD_WHISPERED_SPEECH_ID, self.CMD_FORMAT_ID, \
                  self.CMD_PROPAGATE_ID, self.CMD_MULTISELECT_ID, self.CMD_PLAY_ID):
            self.EnableTool(x, False)

##        # On Windows, we need to display a "Transcript Health" indicator related to GDI resources.
##        # I tried to handle this through changing the "Save" button graphic in the Toolbar, but that
##        # caused the Search box to be disabled when the graphic was updated.  Weird.
##        if 'wxMSW' in wx.PlatformInfo:
##            # Create Bitmap objects for the four images used to populate the health indicator
##            self.GDIBmpAll = TransanaGlobal.GetImage(TransanaImages.StopLightAll)
##            self.GDIBmpRed = TransanaGlobal.GetImage(TransanaImages.StopLightRed)
##            self.GDIBmpYellow = TransanaGlobal.GetImage(TransanaImages.StopLightYellow)
##            self.GDIBmpGreen = TransanaGlobal.GetImage(TransanaImages.StopLightGreen)
##            # Create a Static Bitmap to display the Transcript Health indicator.  Start it with the neutral graphic
##            self.GDIBmp = wx.StaticBitmap(self, -1, self.GDIBmpAll)
##            self.AddControl(self.GDIBmp)
##            # Set the proper Tool Tip
##            self.GDIBmp.SetToolTip(wx.ToolTip(_("Transcript Healthy")))

        # Add Quick Search tools
        # Start with the Search Backwards button
        self.CMD_SEARCH_BACK_ID = wx.NewId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_BACK, wx.ART_TOOLBAR, (16,16))
        self.tools.append(self.AddTool(self.CMD_SEARCH_BACK_ID, bmp,
                        shortHelpString=_("Search backwards")))
        wx.EVT_MENU(self, self.CMD_SEARCH_BACK_ID, self.parent.OnSearch)

        # Add the Search Text box
        self.searchText = wx.TextCtrl(self, -1, size=(100, 20), style=wx.TE_PROCESS_ENTER)
        self.AddControl(self.searchText)
        self.Bind(wx.EVT_TEXT_ENTER, self.parent.OnSearch, self.searchText)

        # Add the Search Forwards button
        self.CMD_SEARCH_NEXT_ID = wx.NewId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_FORWARD, wx.ART_TOOLBAR, (16,16))
        self.tools.append(self.AddTool(self.CMD_SEARCH_NEXT_ID, bmp,
                        shortHelpString=_("Search forwards")))
        wx.EVT_MENU(self, self.CMD_SEARCH_NEXT_ID, self.parent.OnSearch)

        self.AddSeparator()

        # Add the Selection Label, which indicates the time position of the current selection
        self.selectionText = wx.StaticText(self, -1, "", size=wx.Size(200, 20))
        self.AddControl(self.selectionText)

    def GetNextId(self):
        """Get a new event ID to use for the toolbar objects."""
        idVal = wx.NewId()
        return idVal
        
    def ClearToolbar(self):
        """Clear buttons to default state."""
        # Reset toggle buttons to OFF.  This does not cause any events
        # to be emitted (only affects GUI state)
        self.ToggleTool(self.CMD_BOLD_ID, False)
        self.ToggleTool(self.CMD_ITALIC_ID, False)
        self.ToggleTool(self.CMD_UNDERLINE_ID, False)
        self.ToggleTool(self.CMD_READONLY_ID, False)
        self.ToggleTool(self.CMD_SHOWHIDE_ID, False)
        self.ToggleTool(self.CMD_SHOWHIDETIME_ID, False)
        self.UpdateEditingButtons()
        # Clear the Search Text
        self.parent.ClearSearch()
        
    def OnUndo(self, evt):
        """ Implement Undo Button """
        self.parent.editor.undo()

    def OnBold(self, evt):
        """ Implement Bold Button """
        bold_state = self.GetToolState(self.CMD_BOLD_ID)
        self.parent.editor.set_bold(bold_state)
        
    def OnItalic(self, evt):
        """ Implement Italics Button """
        italic_state = self.GetToolState(self.CMD_ITALIC_ID)
        self.parent.editor.set_italic(italic_state)
 
    def OnUnderline(self, evt):
        """ Implement Underline Button """
        underline_state = self.GetToolState(self.CMD_UNDERLINE_ID)
        self.parent.editor.set_underline(underline_state)

    def OnInsertChar(self, evt):
        """ Insert Jeffersonian Special Character """
        # Determine which button was pressed
        idVal = evt.GetId()
        # Rising Intonation
        if idVal == self.CMD_RISING_INT_ID:
            # Insert the Up Arrow / Rising Intonation symbol
            self.parent.editor.InsertRisingIntonation()
        # Falling Intonation
        elif idVal == self.CMD_FALLING_INT_ID:
            # Insert the Down Arrow / Falling Intonation symbol
            self.parent.editor.InsertFallingIntonation()
        # Audible Breath
        elif idVal == self.CMD_AUDIBLE_BREATH_ID:
            # Insert the High Dot / Inbreath symbol
            self.parent.editor.InsertInBreath()
        # Whisper
        elif idVal == self.CMD_WHISPERED_SPEECH_ID:
            # Insert the Open Dot / Whispered Speech symbol
            self.parent.editor.InsertWhisper()
        
    def OnShowHideCodes(self, evt):
        """ Implement Show / Hide Button """
        # Get the button's "indent" state
        show_codes = self.GetToolState(self.CMD_SHOWHIDE_ID)
        # Depending on the indent state, show or hide the time codes
        if show_codes:
            self.parent.editor.show_codes(showPopup=True)
        else:
            self.parent.editor.hide_codes(showPopup=True)

    def OnShowHideValues(self, event):
        """ Implement Show / Hide Time Code Values """
        self.parent.editor.show_timecodevalues(self.GetToolState(self.CMD_SHOWHIDETIME_ID))

    def OnReadOnlySelect(self, evt):
        """ Implement Read Only / Edit Mode """
        # If there's no object currently loaded in the interface ...
        if self.parent.editor.TranscriptObj == None:
            # ... then we can't edit it, can we?  Let's get out of here.
            return
        # Get the button's "indent" state
        can_edit = self.GetToolState(self.CMD_READONLY_ID)
        # If leaving edit mode, prompt for save if necessary.
        if not can_edit:
            if not self.parent.ControlObject.SaveTranscript(1, transcriptToSave=self.parent.transcriptWindowNumber):
                # Reset the Toolbar
                self.ClearToolbar()
                # User chose to not save, revert back to database version
                tobj = self.parent.editor.TranscriptObj
                # Set to None so that it doesn't ask us to save twice, since
                # normally when we load a Transcript with one already loaded,
                # it prompts to save for changes.
                self.parent.editor.TranscriptObj = None                
                if tobj:
                    # Determine whether we have a Clip Transcript (not pickled) or an Episode Transcript (pickled)
                    # and load it.
                    if tobj.clip_num > 0:
                        # If we're using the Rich Text Ctrl ...
                        if TransanaConstants.USESRTC:
                            # Load the Clip Transcript
                            self.parent.editor.load_transcript(tobj)
                        # If we're using the Styled Text Ctrl ...
                        else:
                            # The Clip Transcript is never pickled.
                            self.parent.editor.load_transcript(tobj)
                    else:
                        # If we're using the Rich Text Ctrl ...
                        if TransanaConstants.USESRTC:
                            # Load the Episode Transcript
                            self.parent.editor.load_transcript(tobj)
                        else:
                            # The Episode Transcript will always have been pickled in this circumstance.
                            self.parent.editor.load_transcript(tobj, dataType='pickle')
            # Drop the Record Lock
            self.parent.editor.TranscriptObj.unlock_record()
            # Set the Read only state based on the button's indent
            self.parent.editor.set_read_only(not can_edit)
            # update the Toolbar's state
            self.UpdateEditingButtons()
        # if entering Edit Mode ...
        else:
            try:
                # Remember the original Last Save time
                oldLastSaveTime = self.parent.editor.TranscriptObj.lastsavetime
                # Try to get a record lock
                self.parent.editor.TranscriptObj.lock_record()
                # If the Transcript Object was updated during the Record Lock (due to having been
                # edited in the interim by another user) we need to refresh the editor!
                # Check the new LastSaveTime with the original one.
                if oldLastSaveTime != self.parent.editor.TranscriptObj.lastsavetime:
                    msg = _('This Transcript has been updated since you originally loaded it!\nYour copy of the record will be refreshed to reflect the changes.')
                    dlg = Dialogs.InfoDialog(self.parent, msg)
                    dlg.ShowModal()
                    dlg.Destroy()
                    # If Time Code Data is displayed ...
                    if self.GetToolState(self.CMD_SHOWHIDETIME_ID):
                        # ... signal that it was being shown ...
                        timeCodeDataShowing = True
                        # ... and HIDE it by toggling the button and calling the event!  We don't want it to be propagated.
                        self.ToggleTool(self.CMD_SHOWHIDETIME_ID, False)
                        self.OnShowHideValues(evt)
                    # If Time Code Data is NOT displayed ...
                    else:
                        # ... then note that it isn't so we know not to re-display it later.
                        timeCodeDataShowing = False
                    # Determine whether we have a Clip Transcript (not pickled) or an Episode Transcript (pickled)
                    # and load it.
                    if self.parent.editor.TranscriptObj.clip_num > 0:
                        # The Clip Transcript is never pickled.
                        self.parent.editor.load_transcript(self.parent.editor.TranscriptObj)
                    else:
                        if TransanaConstants.USESRTC:
                            # Load the Episode Transcript
                            self.parent.editor.load_transcript(self.parent.editor.TranscriptObj)
                        else:
                            # The Episode Transcript will always have been pickled in this circumstance.
                            self.parent.editor.load_transcript(self.parent.editor.TranscriptObj, dataType='pickle')
                    # reloading the Transcript unfortunately unlocks the record.  I can't figure out
                    # a clever way to avoid this, so let's just re-lock the record.
                    self.parent.editor.TranscriptObj.lock_record()
                    # If Time Code Data was being displayed ...
                    if timeCodeDataShowing:
                        # ... RE-DISPLAY it by toggling the button and calling the event!
                        self.ToggleTool(self.CMD_SHOWHIDETIME_ID, True)
                        self.OnShowHideValues(evt)
                # Set the Read only state based on the button's indent
                self.parent.editor.set_read_only(not can_edit)
                # update the Toolbar's state
                self.UpdateEditingButtons()

                # If we're using the Rich Text Ctrl ...
                if TransanaConstants.USESRTC:
                    # There's a bit of weirdness with the RTC, or more likely with my implementation of it.
                    # New transcripts don't have proper values for Paragraph Alignment or Line Spacing.
                    # I tried to add them when initializing the control, but that caused problems with
                    # RTF import.  The following code tries to fix that problem when the USER first opens
                    # a new transcript.  (That's why it's here instead of elsewhere.)
                    if self.parent.editor.GetLastPosition() == 0:
                        self.parent.editor.SetTxtStyle(parAlign = wx.TEXT_ALIGNMENT_LEFT, parLineSpacing = wx.TEXT_ATTR_LINE_SPACING_NORMAL)
                    # Call CheckFormatting() to set the Default and Basic Styles correctly
                    self.parent.editor.CheckFormatting()

            # Process Record Lock exception
            except TransanaExceptions.RecordLockedError, e:
                self.parent.editor.TranscriptObj.lastsavetime = oldLastSaveTime
                self.ToggleTool(self.CMD_READONLY_ID, not self.GetToolState(self.CMD_READONLY_ID))
                if self.parent.editor.TranscriptObj.id != '':
                    rtype = _('Transcript')
                    idVal = self.parent.editor.TranscriptObj.id
                    TransanaExceptions.ReportRecordLockedException(rtype, idVal, e)
                else:
                    if 'unicode' in wx.PlatformInfo:
                        msg = unicode(_('You cannot proceed because you cannot obtain a lock on the Clip Transcript.\n'), 'utf8') + \
                              unicode(_('The record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = msg % (e.user)
                    else:
                        msg = _('You cannot proceed because you cannot obtain a lock on the Clip Transcript.\n') + \
                              _('The record is currently locked by %s.\nPlease try again later.') 
                    dlg = Dialogs.ErrorDialog(self.parent, msg)
                    dlg.ShowModal()
                    dlg.Destroy()
            # Process Record Not Found exception
            except TransanaExceptions.RecordNotFoundError, e:
                # Reject the attempt to go into Edit mode by toggling the button back to its original state
                self.ToggleTool(self.CMD_READONLY_ID, not self.GetToolState(self.CMD_READONLY_ID))
                if self.parent.editor.TranscriptObj.id != '':
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(_('You cannot proceed because %s "%s" cannot be found.'), 'utf8') + \
                              unicode(_('\nIt may have been deleted by another user.'), 'utf8')
                        msg = msg % (unicode(_('Transcript'), 'utf8'), self.parent.editor.TranscriptObj.id)
                    else:
                        msg = _('You cannot proceed because %s "%s" cannot be found.') + \
                              _('\nIt may have been deleted by another user.') % (_('Transcript'), self.parent.editor.TranscriptObj.id)
                else:
                    msg = _('You cannot proceed because the Clip Transcript cannot be found.\n') + \
                          _('It may have been deleted by another user.')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8')
                dlg = Dialogs.ErrorDialog(self.parent, msg)
                dlg.ShowModal()
                dlg.Destroy()
                # Clear the deleted objects from the Transana Interface.  Otherwise, problems arise.
                self.parent.ControlObject.ClearAllWindows()

    def OnFormat(self, event):
        """ Format Text """
        # Call the ControlObject's Format Dialog method
        self.parent.ControlObject.TranscriptCallFormatDialog()
        
    def UpdateEditingButtons(self):
        """ Update the Toolbar Buttons depending on the Edit State """
        # Enable/Disable editing buttons
        can_edit = not self.parent.editor.get_read_only()
        for x in (self.CMD_UNDO_ID, self.CMD_BOLD_ID, self.CMD_ITALIC_ID, self.CMD_UNDERLINE_ID, \
                  self.CMD_RISING_INT_ID, self.CMD_FALLING_INT_ID, \
                  self.CMD_AUDIBLE_BREATH_ID, self.CMD_WHISPERED_SPEECH_ID, self.CMD_FORMAT_ID, \
                  self.CMD_PROPAGATE_ID):
            self.EnableTool(x, can_edit)
        # Enable/Disable Transcript menu Items
        self.parent.ControlObject.SetTranscriptEditOptions(can_edit)
##        # If on Windows and we are LEAVING Edit Mode ...
##        if ('wxMSW' in wx.PlatformInfo) and not can_edit:
##            # ... change the Transcript Health Indicator to Neutral
##            self.GDIBmp.SetBitmap(self.GDIBmpAll)

    def UpdateMultiTranscriptButtons(self, enable):
        """ Update the Toolbar Buttons depending on the enable parameter """
        for x in (self.CMD_MULTISELECT_ID, self.CMD_PLAY_ID):
            self.EnableTool(x, enable)

    def OnQuickClip(self, event):
        """ Create a Quick Clip """
        # Call upon the Control Object to create the Quick Clip
        self.parent.ControlObject.CreateQuickClip()

    def OnEditKeywords(self, evt):
        """ Implement the Edit Keywords button """
        # Determine if a Transcript is loaded, and if so, what kind
        if self.parent.editor.TranscriptObj != None:
            # Initialize a list where we can keep track of clip transcripts that are locked because they are in Edit mode.
            TranscriptLocked = []
            try:
                # If an episode/clip has multiple transcripts, we could run into lock problems.  Let's try to detect that.
                # (This is probably poor form from an object-oriented standpoint, but I don't know a better way.)
                # For each currently-open Transcript window ...
                for trWin in self.parent.ControlObject.TranscriptWindow:
                    # ... note if the transcript is currently locked.
                    TranscriptLocked.append(trWin.dlg.editor.TranscriptObj.isLocked)
                    # If it is locked ...
                    if trWin.dlg.editor.TranscriptObj.isLocked:
                        # Leave Edit Mode, which will prompt about saving the Transcript.
                        # a) toggle the button
                        trWin.dlg.toolbar.ToggleTool(trWin.dlg.toolbar.CMD_READONLY_ID, not trWin.dlg.toolbar.GetToolState(trWin.dlg.toolbar.CMD_READONLY_ID))
                        # b) call the event that responds to the button state change
                        trWin.dlg.toolbar.OnReadOnlySelect(evt)

                # If the underlying Transcript object has a clip number, we're working with a CLIP.
                if self.parent.editor.TranscriptObj.clip_num > 0:
                    # Finally, we can load the Clip object.
                    obj = Clip.Clip(self.parent.editor.TranscriptObj.clip_num)
                # Otherwise ...
                else:
                    # ... load the Episode
                    obj = Episode.Episode(self.parent.editor.TranscriptObj.episode_num)
            # Process Record Not Found exception
            except TransanaExceptions.RecordNotFoundError, e:
                msg = _('You cannot proceed because the %s cannot be found.')
                # If the Transcript does not have a clip number, 
                if self.parent.editor.TranscriptObj.clip_num == 0:
                    prompt = _('Episode')
                else:
                    prompt = _('Clip')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(msg, 'utf8') % unicode(prompt, 'utf8') + \
                          unicode(_('\nIt may have been deleted by another user.'), 'utf8')
                else:
                    msg = msg % prompt + _('\nIt may have been deleted by another user.')
                dlg = Dialogs.ErrorDialog(self.parent, msg)
                dlg.ShowModal()
                dlg.Destroy()
                # Clear the deleted objects from the Transana Interface.  Otherwise, problems arise.
                self.parent.ControlObject.ClearAllWindows()
                return
            try:
                # Lock the data record
                obj.lock_record()
                # Determine the title for the KeywordListEditForm Dialog Box
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("Keywords for %s"), 'utf8')
                else:
                    prompt = _("Keywords for %s")
                dlgTitle = prompt % obj.id
                # Extract the keyword List from the Data object
                kwlist = []
                for kw in obj.keyword_list:
                    kwlist.append(kw)
                    
                # Create/define the Keyword List Edit Form
                dlg = KeywordListEditForm.KeywordListEditForm(self.parent, -1, dlgTitle, obj, kwlist)
                # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
                contin = True
                # While the "continue" flag is True ...
                while contin:
                    # if the user pressed "OK" ...
                    try:
                        # Show the Keyword List Edit Form and process it if the user selects OK
                        if dlg.ShowModal() == wx.ID_OK:
                            # Clear the local keywords list and repopulate it from the Keyword List Edit Form
                            kwlist = []
                            for kw in dlg.keywords:
                                kwlist.append(kw)

                            # Copy the local keywords list into the appropriate object
                            obj.keyword_list = kwlist

                            # If we are dealing with an Episode ...
                            if isinstance(obj, Episode.Episode):
                                # Check to see if there are keywords to be propagated
                                self.parent.ControlObject.PropagateEpisodeKeywords(obj.number, obj.keyword_list)

                            # Save the Data object
                            obj.db_save()

                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                if isinstance(obj, Episode.Episode):
                                    msg = 'Episode %d' % obj.number
                                    msgObjType = 'Episode'
                                    msgObjClipEpNum = 0
                                elif isinstance(obj, Clip.Clip):
                                    msg = 'Clip %d' % obj.number
                                    msgObjType = 'Clip'
                                    msgObjClipEpNum = obj.episode_num
                                else:
                                    msg = ''
                                    msgObjType = 'None'
                                    msgObjClipEpNum = 0
                                if msg != '':
                                    if TransanaGlobal.chatWindow != None:
                                        # Send the "Update Keyword List" message
                                        TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                            # If any Keyword Examples were removed, remove them from the Database Tree
                            for (keywordGroup, keyword, clipNum) in dlg.keywordExamplesToDelete:
                                self.parent.ControlObject.RemoveDataWindowKeywordExamples(keywordGroup, keyword, clipNum)

                            # Update the Data Window Keywords Tab (this must be done AFTER the Save)
                            self.parent.ControlObject.UpdateDataWindowKeywordsTab()

                            # Update the Keyword Visualizations
                            self.parent.ControlObject.UpdateKeywordVisualization()

                            # Notify other computers to update the Keyword Visualization as well.
                            if not TransanaConstants.singleUserVersion:
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % (msgObjType, obj.number, msgObjClipEpNum))

                            # If we do all this, we don't need to continue any more.
                            contin = False

                        # If the user pressed Cancel ...
                        else:
                            # ... then we don't need to continue any more.
                            contin = False

                    # Handle "SaveError" exception
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # Refresh the Keyword List, if it's a changed Keyword error
                        dlg.refresh_keywords()
                        # Highlight the first non-existent keyword in the Keywords control
                        dlg.highlight_bad_keyword()

                    # Handle other exceptions
                    except:
                        if DEBUG:
                            import traceback
                            traceback.print_exc(file=sys.stdout)
                        # Display the Exception Message, allow "continue" flag to remain true
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_("Exception %s: %s"), 'utf8')
                        else:
                            prompt = _("Exception %s: %s")
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                        errordlg.ShowModal()
                        errordlg.Destroy()

                # Unlock the Data Object
                obj.unlock_record()

            except TransanaExceptions.RecordLockedError, e:
                """Handle the RecordLockedError exception."""
                if isinstance(obj, Episode.Episode):
                    rtype = _('Episode')
                elif isinstance(obj, Clip.Clip):
                    rtype = _('Clip')
                idVal = obj.id
                TransanaExceptions.ReportRecordLockedException(rtype, idVal, e)

            # Process Record Not Found exception
            except TransanaExceptions.RecordNotFoundError, e:
                msg = _('You cannot proceed because the %s cannot be found.')
                prompt = _('Transcript')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(msg, 'utf8') % unicode(prompt, 'utf8') + \
                          unicode(_('\nIt may have been deleted by another user.'), 'utf8')
                dlg = Dialogs.ErrorDialog(self.parent, msg)
                dlg.ShowModal()
                dlg.Destroy()
                # Clear the deleted objects from the Transana Interface.  Otherwise, problems arise.
                self.parent.ControlObject.ClearAllWindows()

    def OnSave(self, evt):
        """ Implement the Save Button """
        self.parent.ControlObject.SaveTranscript()

    def OnPropagate(self, event):
        """ Implement Propagate Changes button """
        # If Time Code Data is displayed ...
        if self.GetToolState(self.CMD_SHOWHIDETIME_ID):
            # ... signal that it was being shown ...
            timeCodeDataShowing = True
            # ... and HIDE it by toggling the button and calling the event!  We don't want it to be propagated.
            self.ToggleTool(self.CMD_SHOWHIDETIME_ID, False)
            self.OnShowHideValues(event)
        # If Time Code Data is NOT displayed ...
        else:
            # ... then note that it isn't so we know not to re-display it later.
            timeCodeDataShowing = False
        
        # Call the Propagate Changes method in the Control Object
        self.parent.ControlObject.PropagateChanges(self.parent.transcriptWindowNumber)

        # If Time Code Data was being displayed ...
        if timeCodeDataShowing:
            # ... RE-DISPLAY it by toggling the button and calling the event!
            self.ToggleTool(self.CMD_SHOWHIDETIME_ID, True)
            self.OnShowHideValues(event)

    def OnMultiSelect(self, event):
        """ Implements the "Match Selection in Multiple Transcripts" button """
        self.parent.ControlObject.MultiSelect(self.parent.transcriptWindowNumber)

    def OnMultiPlay(self, event):
        """ Implements the "Play" button in the multi-transcript environment """
        self.parent.ControlObject.MultiPlay()

    def OnStyleChange(self, editor):
        """This event handler is setup in the higher level Transcript Window,
        which instructs the Transcript editor to call this function when
        the current style changes automatically, but not programatically."""
        self.ToggleTool(self.CMD_BOLD_ID, editor.get_bold())
        self.ToggleTool(self.CMD_ITALIC_ID, editor.get_italic())
        self.ToggleTool(self.CMD_UNDERLINE_ID, editor.get_underline())

    def ChangeLanguages(self):
        """ Update all on-screen prompts to the new language """
        # Update the Speed Button Tool Tips
        self.SetToolShortHelp(self.CMD_UNDO_ID, _('Undo action'))
        self.SetToolShortHelp(self.CMD_BOLD_ID, _('Bold text'))
        self.SetToolShortHelp(self.CMD_ITALIC_ID, _("Italic text"))
        self.SetToolShortHelp(self.CMD_UNDERLINE_ID, _("Underline text"))
        self.SetToolShortHelp(self.CMD_RISING_INT_ID, _("Rising Intonation"))
        self.SetToolShortHelp(self.CMD_FALLING_INT_ID, _("Falling Intonation"))
        self.SetToolShortHelp(self.CMD_AUDIBLE_BREATH_ID, _("Audible Breath"))
        self.SetToolShortHelp(self.CMD_WHISPERED_SPEECH_ID, _("Whispered Speech"))
        self.SetToolShortHelp(self.CMD_SHOWHIDE_ID, _("Show/Hide Time Code Indexes"))
        self.SetToolShortHelp(self.CMD_SHOWHIDETIME_ID, _("Show/Hide Time Code Values"))
        self.SetToolShortHelp(self.CMD_READONLY_ID, _("Edit/Read-only select"))
        self.SetToolShortHelp(self.CMD_FORMAT_ID, _("Format"))
        self.SetToolShortHelp(self.CMD_QUICKCLIP_ID, _("Create Quick Clip"))
        self.SetToolShortHelp(self.CMD_KEYWORD_ID, _("Edit Keywords"))
        self.SetToolShortHelp(self.CMD_SAVE_ID, _("Save Transcript"))
        self.SetToolShortHelp(self.CMD_PROPAGATE_ID, _("Propagate Changes"))
        self.SetToolShortHelp(self.CMD_MULTISELECT_ID, _("Match Selection in Other Transcripts"))
        self.SetToolShortHelp(self.CMD_PLAY_ID, _("Play Multiple Transcript Selection"))
        self.SetToolShortHelp(self.CMD_SEARCH_BACK_ID, _("Search backwards"))
        self.SetToolShortHelp(self.CMD_SEARCH_NEXT_ID, _("Search forwards"))
