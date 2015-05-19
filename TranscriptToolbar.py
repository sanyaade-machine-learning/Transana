# Copyright (C) 2003 - 2007 The Board of Regents of the University of Wisconsin System 
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
# import Transana modules
import Clip
import Dialogs
import Episode
import KeywordListEditForm
import TransanaConstants
import TransanaExceptions
import TransanaGlobal
import TranscriptEditor
 

class TranscriptToolbar(wx.ToolBar):
    """This class consists of a toolbar for use with a TranscriptEditor
    object.  It inherits from the wxToolbar class.  This class is mostly
    self-sustaining, and does not require much in terms of a public
    interface.  Its objects are intended to be created, and then left alone
    until destroyed."""

    def __init__(self, parent, id=-1):
        """Initialize an TranscriptToolbar object."""
        # Create a ToolBar as self
        wx.ToolBar.__init__(self, parent, id, wx.DefaultPosition,
                            wx.Size(325, 30), wx.TB_HORIZONTAL \
                                    | wx.NO_BORDER | wx.TB_FLAT | wx.TB_TEXT)
        # remember the parent
        self.parent = parent
        # Keep a list of the tools placed on the toolbar so they're more easily manipulated
        self.tools = []

        # Create an Undo button
        self.CMD_UNDO_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_UNDO_ID, wx.Bitmap("images/Undo16.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_('Undo action')))
        wx.EVT_MENU(self, self.CMD_UNDO_ID, self.OnUndo)

        self.AddSeparator()
  
        # Bold, Italic, Underline buttons
        self.CMD_BOLD_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_BOLD_ID, wx.Bitmap("images/Bold.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1, shortHelpString=_('Bold text')))
        wx.EVT_MENU(self, self.CMD_BOLD_ID, self.OnBold)

        self.CMD_ITALIC_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_ITALIC_ID, wx.Bitmap("images/Italic.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1, shortHelpString=_("Italic text")))
        wx.EVT_MENU(self, self.CMD_ITALIC_ID, self.OnItalic)
       
        self.CMD_UNDERLINE_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_UNDERLINE_ID, wx.Bitmap("images/Underline.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1, shortHelpString=_("Underline text")))
        wx.EVT_MENU(self, self.CMD_UNDERLINE_ID, self.OnUnderline)

        self.AddSeparator()

        # Jeffersonian Symbols
        self.CMD_RISING_INT_ID = self.GetNextId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_UP, wx.ART_TOOLBAR, (16,16))
        self.tools.append(self.AddTool(self.CMD_RISING_INT_ID, bmp,
                        shortHelpString=_("Rising Intonation")))
        wx.EVT_MENU(self, self.CMD_RISING_INT_ID, self.OnInsertChar)
        
        self.CMD_FALLING_INT_ID = self.GetNextId()
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_DOWN, wx.ART_TOOLBAR, (16,16))
        self.tools.append(self.AddTool(self.CMD_FALLING_INT_ID, bmp,
                        shortHelpString=_("Falling Intonation")))
        wx.EVT_MENU(self, self.CMD_FALLING_INT_ID, self.OnInsertChar) 
       
        self.CMD_AUDIBLE_BREATH_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_AUDIBLE_BREATH_ID, wx.Bitmap("images/AudibleBreath.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_("Audible Breath")))
        wx.EVT_MENU(self, self.CMD_AUDIBLE_BREATH_ID, self.OnInsertChar)
    
        self.CMD_WHISPERED_SPEECH_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_WHISPERED_SPEECH_ID, wx.Bitmap("images/WhisperedSpeech.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_("Whispered Speech")))
        wx.EVT_MENU(self, self.CMD_WHISPERED_SPEECH_ID, self.OnInsertChar)
      
        self.AddSeparator()

        # Add show / hide timecodes button
        self.CMD_SHOWHIDE_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_SHOWHIDE_ID, wx.Bitmap("images/TimeCode16.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1, shortHelpString=_("Show/Hide Time Code Indexes")))
        wx.EVT_MENU(self, self.CMD_SHOWHIDE_ID, self.OnShowHideCodes)

        # Add read only / edit mode button
        self.CMD_READONLY_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_READONLY_ID, wx.Bitmap("images/ReadOnly16.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1, shortHelpString=_("Edit/Read-only select")))
        wx.EVT_MENU(self, self.CMD_READONLY_ID, self.OnReadOnlySelect)
        
        self.AddSeparator()

        # Add Edit keywords button
        self.CMD_KEYWORD_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_KEYWORD_ID, wx.Bitmap("images/KeywordRoot16.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_("Edit Keywords")))
        wx.EVT_MENU(self, self.CMD_KEYWORD_ID, self.OnEditKeywords)

        # Add Save Button
        self.CMD_SAVE_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_SAVE_ID, wx.Bitmap("images/Save16.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_("Save Transcript")))
        wx.EVT_MENU(self, self.CMD_SAVE_ID, self.OnSave)

        self.AddSeparator()

        # SEARCH moved to TranscriptionUI because you can't put a TextCtrl on a Toolbar on the Mac!

        # Set the Initial State of the Editing Buttons to "False"
        for x in (self.CMD_UNDO_ID, self.CMD_BOLD_ID, self.CMD_ITALIC_ID, self.CMD_UNDERLINE_ID, \
                    self.CMD_RISING_INT_ID, self.CMD_FALLING_INT_ID, \
                    self.CMD_AUDIBLE_BREATH_ID, self.CMD_WHISPERED_SPEECH_ID):
            self.EnableTool(x, False)
        
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
            self.parent.editor.show_codes()
        else:
            self.parent.editor.hide_codes()

    def OnReadOnlySelect(self, evt):
        """ Implement Read Only / Edit Mode """
        # Get the button's "indent" state
        can_edit = self.GetToolState(self.CMD_READONLY_ID)
        # If leaving edit mode, prompt for save if necessary.
        if not can_edit:
            if not self.parent.ControlObject.SaveTranscript(1):
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
                        # The Clip Transcript is never pickled.
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
                    # Determine whether we have a Clip Transcript (not pickled) or an Episode Transcript (pickled)
                    # and load it.
                    if self.parent.editor.TranscriptObj.clip_num > 0:
                        # The Clip Transcript is never pickled.
                        self.parent.editor.load_transcript(self.parent.editor.TranscriptObj)
                    else:
                        # The Episode Transcript will always have been pickled in this circumstance.
                        self.parent.editor.load_transcript(self.parent.editor.TranscriptObj, dataType='pickle')
                    # reloading the Transcript unfortunately unlocks the record.  I can't figure out
                    # a clever way to avoid this, so let's just re-lock the record.
                    self.parent.editor.TranscriptObj.lock_record()
                # Set the Read only state based on the button's indent
                self.parent.editor.set_read_only(not can_edit)
                # update the Toolbar's state
                self.UpdateEditingButtons()
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
                        msg = msg % (_('Transcript'), self.parent.editor.TranscriptObj.id)
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

    def UpdateEditingButtons(self):
        """ Update the Toolbar Buttons depending on the Edit State """
        # Enable/Disable editing buttons
        can_edit = not self.parent.editor.get_read_only()
        for x in (self.CMD_UNDO_ID, self.CMD_BOLD_ID, self.CMD_ITALIC_ID, self.CMD_UNDERLINE_ID, \
                    self.CMD_RISING_INT_ID, self.CMD_FALLING_INT_ID, \
                    self.CMD_AUDIBLE_BREATH_ID, self.CMD_WHISPERED_SPEECH_ID):
            self.EnableTool(x, can_edit)
        # Enable/Disable Transcript menu Items
        self.parent.ControlObject.SetTranscriptEditOptions(can_edit)

    def OnEditKeywords(self, evt):
        """ Implement the Edit Keywords button """
        # Determine if a Transcript is loaded, and if so, what kind
        if self.parent.editor.TranscriptObj != None:
            # Default that the transcript was NOT locked, which means we weren't in Edit mode.
            clipTranscriptLocked = False
            # If the Transcript has a clip number, load the Clip
            if self.parent.editor.TranscriptObj.clip_num > 0:
                # If the Clip Transcript is locked, we need to save it first and unlock it.
                if self.parent.editor.TranscriptObj.isLocked:
                    # Note that the transcript was locked, which means we HAD to be in Edit Mode
                    clipTranscriptLocked = True
                    # Leave Edit Mode, which will prompt about saving the Transcript.
                    # a) toggle the button
                    self.ToggleTool(self.CMD_READONLY_ID, not self.GetToolState(self.CMD_READONLY_ID))
                    # b) call the event that responds to the button state change
                    self.OnReadOnlySelect(evt)
                    # Get the "Last Save Time" value
                    lastSaveTime = self.parent.editor.TranscriptObj.lastsavetime
                # Finally, we can load the Clip object
                obj = Clip.Clip(self.parent.editor.TranscriptObj.clip_num)
            # Otherwise ...
            else:
                # ... load the Episode
                obj = Episode.Episode(self.parent.editor.TranscriptObj.episode_num)
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
                # Show the Keyword List Edit Form and process it if the user selects OK
                if dlg.ShowModal() == wx.ID_OK:
                    # Clear the local keywords list and repopulate it from the Keyword List Edit Form
                    kwlist = []
                    for kw in dlg.keywords:
                        kwlist.append(kw)

                    # Copy the local keywords list into the appropriate object
                    obj.keyword_list = kwlist

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

                    
                # Unlock the Data Object
                obj.unlock_record()

                # Load the revised Transcript.  (This prevents a Last Save Time error in MU.)
                self.parent.editor.TranscriptObj.db_load_by_num(self.parent.editor.TranscriptObj.number)
                # If we used to be in Edit Mode (flagged earlier) ...
                if clipTranscriptLocked:
                    # ... toggle the Edit Mode button ...
                    self.ToggleTool(self.CMD_READONLY_ID, not self.GetToolState(self.CMD_READONLY_ID))
                    # ... and call the event associated with toggling the button.  This puts us back in Edit
                    # mode and locks the Clip Transcript.
                    self.OnReadOnlySelect(evt)
                    
            except TransanaExceptions.RecordLockedError, e:
                """Handle the RecordLockedError exception."""
                if isinstance(obj, Episode.Episode):
                    rtype = _('Episode')
                elif isinstance(obj, Clip.Clip):
                    rtype = _('Clip')
                idVal = obj.id
                TransanaExceptions.ReportRecordLockedException(rtype, idVal, e)

    def OnSave(self, evt):
        """ Implement the Save Button """
        self.parent.ControlObject.SaveTranscript()
            	
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
        self.SetToolShortHelp(self.CMD_READONLY_ID, _("Edit/Read-only select"))
        self.SetToolShortHelp(self.CMD_KEYWORD_ID, _("Edit Keywords"))
        self.SetToolShortHelp(self.CMD_SAVE_ID, _("Save Transcript"))
