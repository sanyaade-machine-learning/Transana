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

"""This module implements the TranscriptToolbar class as part of the Editors
component.
"""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

import wx
import Dialogs
import TranscriptEditor
import TransanaConstants
import TransanaExceptions
import TransanaGlobal
import Clip
import Episode
import KeywordListEditForm
 

class TranscriptToolbar(wx.ToolBar):
    """This class consists of a toolbar for use with a TranscriptEditor
    object.  It inherits from the wxToolbar class.  This class is mostly
    self-sustaining, and does not require much in terms of a public
    interface.  Its objects are intended to be created, and then left alone
    until destroyed."""

    def __init__(self, parent, id=-1):
        """Initialize an TranscriptToolbar object."""
        wx.ToolBar.__init__(self, parent, id, wx.DefaultPosition,
                            wx.Size(400, 30), wx.TB_HORIZONTAL \
                                    | wx.NO_BORDER | wx.TB_TEXT)
        self.parent = parent
        self.tools = []
        self.next_id = 0

        self.CMD_UNDO_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_UNDO_ID,
                        wx.Bitmap("images/Undo16.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_('Undo action')))
        wx.EVT_MENU(self, self.CMD_UNDO_ID, self.OnUndo)

        self.AddSeparator()
  
        # Bold, Italic, Underline buttons
        self.CMD_BOLD_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_BOLD_ID,
                        wx.Bitmap("images/Bold.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1,
                        shortHelpString=_('Bold text')))
        wx.EVT_MENU(self, self.CMD_BOLD_ID, self.OnBold)

        self.CMD_ITALIC_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_ITALIC_ID,
                        wx.Bitmap("images/Italic.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1,
                        shortHelpString=_("Italic text")))
        wx.EVT_MENU(self, self.CMD_ITALIC_ID, self.OnItalic)
       
        self.CMD_UNDERLINE_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_UNDERLINE_ID,
                        wx.Bitmap("images/Underline.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1,
                        shortHelpString=_("Underline text")))
        wx.EVT_MENU(self, self.CMD_UNDERLINE_ID, self.OnUnderline)

        self.AddSeparator()
        
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
        self.tools.append(self.AddTool(self.CMD_AUDIBLE_BREATH_ID,
                        wx.Bitmap("images/AudibleBreath.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_("Audible Breath")))
        wx.EVT_MENU(self, self.CMD_AUDIBLE_BREATH_ID, self.OnInsertChar)
    
        self.CMD_WHISPERED_SPEECH_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_WHISPERED_SPEECH_ID,
                        wx.Bitmap("images/WhisperedSpeech.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_("Whispered Speech")))
        wx.EVT_MENU(self, self.CMD_WHISPERED_SPEECH_ID, self.OnInsertChar)
      
        self.AddSeparator()
        self.CMD_SHOWHIDE_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_SHOWHIDE_ID,
                        wx.Bitmap("images/TimeCode16.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1,
                        shortHelpString=_("Show/Hide Time Code Indexes")))
        wx.EVT_MENU(self, self.CMD_SHOWHIDE_ID, self.OnShowHideCodes)
        
        self.CMD_READONLY_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_READONLY_ID,
                        wx.Bitmap("images/ReadOnly16.xpm", wx.BITMAP_TYPE_XPM),
                        isToggle=1,
                        shortHelpString=_("Edit/Read-only select")))
        wx.EVT_MENU(self, self.CMD_READONLY_ID, self.OnReadOnlySelect)
         
        self.AddSeparator()
        self.CMD_KEYWORD_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_KEYWORD_ID,
                        wx.Bitmap("images/KeywordRoot16.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_("Edit Keywords")))
        wx.EVT_MENU(self, self.CMD_KEYWORD_ID, self.OnEditKeywords)
         
        self.CMD_SAVE_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_SAVE_ID,
                        wx.Bitmap("images/Save16.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_("Save Transcript")))
        wx.EVT_MENU(self, self.CMD_SAVE_ID, self.OnSave)

        self.CMD_CLIP_ID = self.GetNextId()
        self.tools.append(self.AddTool(self.CMD_CLIP_ID,
                        wx.Bitmap("images/Clip16.xpm", wx.BITMAP_TYPE_XPM),
                        shortHelpString=_("Select Clip Text")))
        wx.EVT_MENU(self, self.CMD_CLIP_ID, self.OnClipSelect)

        # Set the Initial State of the Editing Buttons to "False"
        for x in (self.CMD_UNDO_ID, self.CMD_BOLD_ID, self.CMD_ITALIC_ID, self.CMD_UNDERLINE_ID, \
                    self.CMD_RISING_INT_ID, self.CMD_FALLING_INT_ID, \
                    self.CMD_AUDIBLE_BREATH_ID, self.CMD_WHISPERED_SPEECH_ID):
            self.EnableTool(x, False)
        
        #bmp = wx.ArtProvider_GetBitmap(wx.ART_NORMAL_FILE, wx.ART_TOOLBAR, (16,16))
        #self.tools.append(self.AddTool(3, bmp,
        #                shortHelpString=""))
        #wx.EVT_MENU(self, 3, self.OnDummy)
        
        #self.tools.append(self.AddTool(3, 
        #                wx.Bitmap("images/Note16.xpm", wx.BITMAP_TYPE_XPM),
        #                shortHelpString="Second button"))
        
    def GetNextId(self):
        """Get a new event ID to use for the toolbar objects."""
        self.next_id = self.next_id + 1
        return self.next_id - 1
        
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
        
    def OnFirstButton(self, evt):
        #wx.MessageBox("First button", "Blah", wx.OK) 
        self.parent.editor.SaveRTFDocument("test_save.rtf")
        rtfdata = self.parent.editor.GetRTFBuffer()
        f = open("testbuf.rtf", "wt")
        f.write(rtfdata)
        f.close()

    def OnDummy(self, evt):
        raise TransanaExceptions.NotImplementedError

    def OnUndo(self, evt):
        self.parent.editor.undo()

    def OnBold(self, evt):
        bold_state = self.GetToolState(self.CMD_BOLD_ID)
        self.parent.editor.set_bold(bold_state)
        
    def OnItalic(self, evt):
        italic_state = self.GetToolState(self.CMD_ITALIC_ID)
        self.parent.editor.set_italic(italic_state)
 
    def OnUnderline(self, evt):
        underline_state = self.GetToolState(self.CMD_UNDERLINE_ID)
        self.parent.editor.set_underline(underline_state)

    def OnInsertChar(self, evt):
        id = evt.GetId()
        c = ""
        if id == self.CMD_RISING_INT_ID:
            # Ctrl-Cursor-Up inserts the Up Arrow / Rising Intonation symbol
            self.parent.editor.InsertRisingIntonation()
        elif id == self.CMD_FALLING_INT_ID:
            # Ctrl-Cursor-Down inserts the Down Arrow / Falling Intonation symbol
            self.parent.editor.InsertFallingIntonation()
        elif id == self.CMD_AUDIBLE_BREATH_ID:
            # Ctrl-H inserts the High Dot / Inbreath symbol
            self.parent.editor.InsertInBreath()
        elif id == self.CMD_WHISPERED_SPEECH_ID:
            # Ctrl-O inserts the Open Dot / Whispered Speech symbol
            self.parent.editor.InsertWhisper()
        else:
            return
        
    def OnShowHideCodes(self, evt):
        show_codes = self.GetToolState(self.CMD_SHOWHIDE_ID)
        if show_codes:
            self.parent.editor.show_codes()
        else:
            self.parent.editor.hide_codes()

    def OnReadOnlySelect(self, evt):
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

            self.parent.editor.TranscriptObj.unlock_record()
            self.parent.editor.set_read_only(not can_edit)
            self.UpdateEditingButtons()
        else:
            try:
                oldLastSaveTime = self.parent.editor.TranscriptObj.lastsavetime
                self.parent.editor.TranscriptObj.lock_record()
                # If the Transcript Object was updated during the Record Lock (do to having been
                # edited in the interim by another user) we need to refresh the editor!
                # Check the new LastSaveTime with the original one.
                if oldLastSaveTime != self.parent.editor.TranscriptObj.lastsavetime:
                    msg = 'This Transcript has been updated since you originally loaded it!\nYour copy of the record will be refreshed to reflect the changes.'
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
                
                self.parent.editor.set_read_only(not can_edit)
                self.UpdateEditingButtons()
            except TransanaExceptions.RecordLockedError, e:
                self.parent.editor.TranscriptObj.lastsavetime = oldLastSaveTime
                self.ToggleTool(self.CMD_READONLY_ID, not self.GetToolState(self.CMD_READONLY_ID))
                if self.parent.editor.TranscriptObj.id != '':
                    rtype = _('Transcript')
                    idVal = self.parent.editor.TranscriptObj.id
                    TransanaExceptions.ReportRecordLockedException(rtype, idVal, e)
                else:
                    msg = _('You cannot proceed because you cannot obtain a lock on the Clip Transcript.\n' + \
                            'The record is currently locked by %s.\nPlease try again later.') 
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8') % (e.user)
                    dlg = Dialogs.ErrorDialog(self.parent, msg)
                    dlg.ShowModal()
                    dlg.Destroy()
            except TransanaExceptions.RecordNotFoundError, e:
                self.ToggleTool(self.CMD_READONLY_ID, not self.GetToolState(self.CMD_READONLY_ID))
                if self.parent.editor.TranscriptObj.id != '':
                    msg = _('You cannot proceed because %s "%s" cannot be found.' + \
                            '\nIt may have been deleted by another user.')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8') % (_('Transcript'), self.parent.editor.TranscriptObj.id)
                else:
                    msg = _('You cannot proceed because the Clip Transcript cannot be found.\n' + \
                            'It may have been deleted by another user.')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8')
                dlg = Dialogs.ErrorDialog(self.parent, msg)
                dlg.ShowModal()
                dlg.Destroy()


    def UpdateEditingButtons(self):
        # Enable/Disable editing buttons
        can_edit = not self.parent.editor.get_read_only()
        for x in (self.CMD_UNDO_ID, self.CMD_BOLD_ID, self.CMD_ITALIC_ID, self.CMD_UNDERLINE_ID, \
                    self.CMD_RISING_INT_ID, self.CMD_FALLING_INT_ID, \
                    self.CMD_AUDIBLE_BREATH_ID, self.CMD_WHISPERED_SPEECH_ID):
            self.EnableTool(x, can_edit)
        # Enable/Disable Transcript menu Items
        self.parent.ControlObject.SetTranscriptEditOptions(can_edit)

    def OnEditKeywords(self, evt):
        # Determine if a Transcript is loaded, and if so, what kind
        if self.parent.editor.TranscriptObj != None:
            # If the Transcript has a clip number, load the Clips
            if self.parent.editor.TranscriptObj.clip_num > 0:
                obj = Clip.Clip(self.parent.editor.TranscriptObj.clip_num)
            # Otherwise load the Episode
            else:
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
                        elif isinstance(obj, Clip.Clip):
                            msg = 'Clip %d' % obj.number
                        else:
                            msg = ''
                        if msg != '':
                            if TransanaGlobal.chatWindow != None:
                                # Send the "Update Keyword List" message
                                TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                    # If any Keyword Examples were removed, remove them from the Database Tree
                    for (keywordGroup, keyword, clipNum) in dlg.keywordExamplesToDelete:
                        self.parent.ControlObject.RemoveDataWindowKeywordExamples(keywordGroup, keyword, clipNum)

                    # Update the Data Window Keywords Tab (this must be done AFTER the Save)
                    self.parent.ControlObject.UpdateDataWindowKeywordsTab()
                    
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

    def OnSave(self, evt):
        self.parent.ControlObject.SaveTranscript()
            	
    def OnClipSelect(self, event):
        self.parent.editor.OnStartDrag(event)

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
        self.SetToolShortHelp(self.CMD_CLIP_ID, _("Select Clip Text"))

        
# Public methods
    def enable(self):
        """Enable the toolbar."""
    def disable(self):
        """Disable toolbar and all buttons."""

# Private methods    

# Public properties
