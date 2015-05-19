# Copyright (C) 2003 - 2010 The Board of Regents of the University of Wisconsin System 
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

""" A Utility Program to deal with propagating Episode Transcript and Clip Change propagation. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "PropagateEpisodeChanges DEBUG is ON!!"

# import wxPython
import wx
# import Python's sys module
import sys
# import Transana's Clip object
import Clip
# import Transana's Database interface
import DBInterface
# import Transana's Miscellaneous functions
import Misc
# import Transana's Constants
import TransanaConstants
# import Transana's Exception definitions
import TransanaExceptions
# import Transana's Global Definitions
import TransanaGlobal
# import Transana's Transcript Object
import Transcript
# import Transana's Transcript Editor
import TranscriptEditor

# Define IDs for the "Update All" and "Skip" buttons on the confirmation form
ID_UPDATEALL = wx.NewId()
ID_SKIP = wx.NewId()

class PropagateEpisodeChanges(wx.Dialog):
    """ This window displays the Propagate Episode Changes report form. """
    def __init__(self, parent):
        # Remember the parent, which should be a ControlObject
        self.parent = parent
        # Define the main Frame for the Propagate Changes report Window
        wx.Dialog.__init__(self, self.parent.MenuWindow, -1, _("Episode Transcript Change Propagation"), size = (710,650), style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|wx.NO_FULL_REPAINT_ON_RESIZE)
        # Set the background to White
        self.SetBackgroundColour(wx.WHITE)
        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create a Sizer for the form
        box = wx.BoxSizer(wx.VERTICAL)

        # Create a label for the Memo section
        txtMemo = wx.StaticText(self, -1, _("Episode Transcript Change Propagation Report"))
        # Put the label in the Memo Sizer, with a little padding
        box.Add(txtMemo, 0, wx.ALL, 6)
        # Add a TextCtrl for the Report text.  This is read only, as it is filled programmatically.
        self.memo = wx.TextCtrl(self, -1, style = wx.TE_MULTILINE | wx.TE_WORDWRAP | wx.TE_READONLY)
        # Put the Memo control in the Memo Sizer
        box.Add(self.memo, 1, wx.EXPAND | wx.ALL, 6)
        
        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Put an expandable point in the button sizer so everything else can be right justified
        btnSizer.Add((1, 1), 1, wx.EXPAND)
        # Add an OK button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Put the OK button on the button sizer
        btnSizer.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        # Add a Help button
        btnHelp = wx.Button(self, -1, _("Help"))
        # Link the Help button to the OnHelp method
        btnHelp.Bind(wx.EVT_BUTTON, self.OnHelp)
        # Add the Help button to the Button sizer
        btnSizer.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        # Add the button sizer to the main sizer
        box.Add(btnSizer, 0, wx.EXPAND)

        # Attach the Form's Main Sizer to the form
        self.SetSizer(box)
        # Set AutoLayout on
        self.SetAutoLayout(True)
        # Lay out the form
        self.Layout()
        # Set the minimum size for the form.
        self.SetSizeHints(minW = 600, minH = 440)
        # Center the form on the screen
        self.CentreOnScreen()

        # Clear the Report of contents
        self.memo.Clear()

        # Determine the time code preceding the cursor position in the Transcript.  (DON'T USE the video positions, which
        # may not have been updated if you are in Edit mode!)
        timeCodeFromTranscript = self.parent.TranscriptWindow[self.parent.activeTranscript].dlg.editor.get_selected_time_range()[0]
        # request clips that include the current transcript cursor position
        clipList = DBInterface.list_of_clips_by_episode(self.parent.TranscriptWindow[self.parent.activeTranscript].dlg.editor.TranscriptObj.episode_num, timeCodeFromTranscript)
        # But unfortunately, this does not produce a COMPLETE Clip List in some cases.
        # If a time code that starts a clip is deleted, the previous list may not contain that clip.
        # Therefore we need to look for more clips.
        # Let's get a time just before the END of the clip.
        timeCodeFromTranscript2 = self.parent.TranscriptWindow[self.parent.activeTranscript].dlg.editor.get_selected_time_range()[1] - 10
        # request clips that include the end time code based on cursor position
        clipList2 = DBInterface.list_of_clips_by_episode(self.parent.TranscriptWindow[self.parent.activeTranscript].dlg.editor.TranscriptObj.episode_num, timeCodeFromTranscript2)
        # Iterate through the list ...
        for clip in clipList2:
            # ... looking for clips that were NOT included in the earlier list ...
            if clip not in clipList:
                # ... and appending them to the list.
                clipList.append(clip)
        # This still misses clips that have had BOTH time codes removed, but I can't figure out how to fix that.  It shouldn't happen very often.
        
        # If no clips are returned ...
        if len(clipList) == 0:
            # ... inform the user ...
            prompt = _("No clips were found that include %s.")
            if 'unicode' in wx.PlatformInfo:
                prompt = unicode(prompt, 'utf8')
            # ... by adding a message to the report
            self.memo.AppendText(prompt % Misc.time_in_ms_to_str(self.parent.TranscriptWindow[self.parent.activeTranscript].dlg.editor.get_selected_time_range()[0]))

        # Let's start a database Transaction
        # First, let's get a database cursor
        dbCursor = DBInterface.get_db().cursor()
        dbCursor.execute('BEGIN')
        
        # Initially, the user has not indicated that s/he wants to update ALL the clips.
        acceptAll = False
        # Initialize Results to ID-SKIP.  If no Clips are found, it's the same as if the user pressed "Skip".
        results = ID_SKIP
        # Iterate through the Clip List ...
        for clip in clipList:
            # Let's load the Clip
            tempClip = Clip.Clip(clip['ClipNum'])
            # If the clip has only one Transcript ...
            if len(tempClip.transcripts) == 1:
                # ... then give the Clip Transcript the Clip's start and stop times.  (They only get assigned for multi-transcript clips!)
                # (NOTE:  Legacy Clips from Transana before version 2.30 require this.)
                tempClip.transcripts[0].clip_start = tempClip.clip_start
                tempClip.transcripts[0].clip_stop = tempClip.clip_stop
            # For every transcript if each clip ...
            for clipTranscript in tempClip.transcripts:
                # Get the text from the Episode Transcript that matches the Clip TRANSCRIPT's start and stop times.  (NOT JUST THE CLIP'S!!)
                # The return values include the Episode Transcript's time code boundaries, in case they've changed since the
                # clip was created, as well as the NEW text.
                (start, end, text) = self.parent.TranscriptWindow[self.parent.activeTranscript].dlg.editor.GetTextBetweenTimeCodes(clipTranscript.clip_start, clipTranscript.clip_stop)

                # Check the start and end times to make sure neither has changed.  THEY MUST MATCH EXACTLY or we won't propagate to that clip.
                # Also check that the Clip's originating Transcript Number matches the current Episode Transcripts's number, that is,
                # that this clip was indeed taken from THIS transcript.  The one exception to this rule is if the clip has been orphaned,
                # which will probably only be known if the data has been through export/import since the clips was orphaned.
                if (start == clipTranscript.clip_start) and \
                   (end == clipTranscript.clip_stop) and \
                   ((clipTranscript.source_transcript == self.parent.TranscriptWindow[self.parent.activeTranscript].dlg.editor.TranscriptObj.number) or \
                    (clipTranscript.source_transcript == 0)):
                    
                    # Start Exception handling.
                    try:
                        # If the user hasn't signal that they want to update all clips ...
                        if not acceptAll:
                            # ... create the Accept Clip Transcript Changes form ...
                            acceptClipChanges = AcceptClipTranscriptChanges(self, tempClip.number, tempClip.transcripts.index(clipTranscript), text, helpString="Propagate Transcript Changes")
                            # ... and display it, capturing the user feedback.
                            results = acceptClipChanges.GetResults()
                            # Close (and Destroy) the form
                            acceptClipChanges.Destroy()
                            # If the user wants to update all Clips ...
                            if results == ID_UPDATEALL:
                                # ... then changing this variable will prevent the need for further user intervention.
                                acceptAll = True
                        # If the user presses "Update" (OK) or has pressed "Update All" ...
                        if acceptAll or (results == wx.ID_OK):
                            # Lock the Clip (will raise an exception of you can't)
                            clipTranscript.lock_record()
                            # substitute the new text for the old text
                            clipTranscript.text = text
                            # Save the Clip
                            clipTranscript.db_save()
                            # unlock the clip
                            clipTranscript.unlock_record()
                            # Finally, indicate success in the Report
                            self.memo.AppendText(_("Transcript updated for clip"))
                        # If the user indicates we should skip ONE clip ...
                        elif results == ID_SKIP:
                            # ... indicate that in the report and don't do anything else.
                            self.memo.AppendText(_("Transcript change skipped for clip"))
                        # If the user indicates we should CANCEL Transcript propagation ...
                        elif results == wx.ID_CANCEL:
                            # ... indicate that in the report.  The rest of Cancel is implemented later, after we've added
                            #     the full Clip information to the report.
                            self.memo.AppendText(_("Transcript propagation cancelled at clip"))
                    # If a RecordLocked Exception is raised ...
                    except TransanaExceptions.RecordLockedError:
                        # ... indicate that in the report.
                        self.memo.AppendText(_("ERROR: Record lock error for clip"))
                    # If a SaveError exception is raised (I don't know how this might happen!)
                    except TransanaExceptions.SaveError, e:
                        # build the prompt for the user ...
                        prompt = _('ERROR: Save error "%s" for clip')
                        if 'unicode' in wx.PlatformInfo:
                            prompt = unicode(prompt, 'utf8')
                        # ... and report the error in the report.
                        self.memo.AppendText(prompt % e.reason)
                        # unlock the clip
                        clipTranscript.unlock_record()
                # If a clip has Time Code boundary or Transcript Source issues ...
                else:
                    # If the time code boundaries are okay ...
                    if (clipTranscript.source_transcript != self.parent.TranscriptWindow[self.parent.activeTranscript].dlg.editor.TranscriptObj.number) or \
                       (clipTranscript.source_transcript == 0):
                        # ... then we have a Transcript Source error to add to the report.
                        self.memo.AppendText(_("SKIP: A different transcript was used to create clip"))
                    # Otherwise ...
                    else:
                        # ... we have a Time Code Boundary problem to report to the user.
                        self.memo.AppendText(_("ERROR: Clip boundaries don't match transcript time codes for clip"))
                    # If either of these errors occurs, it's the same as if the user pressed "Skip".
                    results = ID_SKIP

                # We need to build the prompt for the Clip information, which is added to ALL of the above prompts
                prompt = _('in collection')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                # Add the clip data to the report.
                self.memo.AppendText(' "%s" %s "%s".  (%s - %s)\n\n' % (tempClip.id, prompt, tempClip.collection_id, Misc.time_in_ms_to_str(clipTranscript.clip_start), Misc.time_in_ms_to_str(clipTranscript.clip_stop)))
                # If the user pressed "Cancel" ...
                if results == wx.ID_CANCEL:
                    # ... we should STOP processing TRANSCRIPTS!
                    break
            # If the user pressed "Cancel" ...
            if results == wx.ID_CANCEL:
                # ... we should STOP processing CLIPS too!
                break

        # If the user pressed Cancel ...
        if results == wx.ID_CANCEL:
            # ... undo Clip Transcript changes in the database ...
            dbCursor.execute("ROLLBACK")
            # ... and inform the user that the changes were cancelled.  To avoid confusion, clear the report information already generated.
            self.memo.Clear()
            self.memo.AppendText(_("Episode Transcript Change Propagation was cancelled.  Therefore, no clips were updated."))
        # If the user pressed anything except Cancel ...
        else:
            # ... commit the changes to the database
            dbCursor.execute("COMMIT")
            
        # Update the contents of the memo
        self.memo.Update()
        # Finally, we can show the dialog so the user can see the report!
        self.ShowModal()

    def OnHelp(self, event):
        """ Method to use when the Help Button is pressed """
        # Locate the Menu Window ...
        if TransanaGlobal.menuWindow != None:
            # .. and use that to get to the standard Help Call method.
            TransanaGlobal.menuWindow.ControlObject.Help("Propagate Transcript Changes")
        

class PropagateClipChanges(wx.Dialog):
    """ This window displays the Propagate Clip Changes report form. """
    def __init__(self, parent, originalClip, sourceTranscriptIndex, newTranscriptText, newClipID=None, newKeywordList=None):
        # If changes get propagated to the currently loaded clip, the data can get stale, which can cause problems.
        # So if we have a Clip loaded in the main interface, it shares a name with the clip data being propagated,
        # and it's not the CURRENT clip ...
        if isinstance(parent.ControlObject.currentObj, Clip.Clip) and \
           (originalClip.id == parent.ControlObject.currentObj.id) and \
           (originalClip.number != parent.ControlObject.currentObj.number):
            # ... not the clip's number so the data can be updated at the end of this process
            CurrentClipToUpdate = parent.ControlObject.currentObj.number
        # If we don't meet those criteria ...
        else:
            # ... then we don't need to update the current clip.
            CurrentClipToUpdate = 0
            
        # Define the main Frame for the Propagate Changes report Window
        wx.Dialog.__init__(self, parent, -1, _("Clip Change Propagation"), size = (710,650), style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|wx.NO_FULL_REPAINT_ON_RESIZE)
        # Set the background to White
        self.SetBackgroundColour(wx.WHITE)
        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create a Sizer for the form
        box = wx.BoxSizer(wx.VERTICAL)

        # Create a label for the Memo section
        txtMemo = wx.StaticText(self, -1, _("Clip Change Propagation Report"))
        # Put the label in the Memo Sizer, with a little padding
        box.Add(txtMemo, 0, wx.ALL, 6)
        # Add a TextCtrl for the Report text.  This is read only, as it is filled programmatically.
        self.memo = wx.TextCtrl(self, -1, style = wx.TE_MULTILINE | wx.TE_WORDWRAP | wx.TE_READONLY)
        # Put the Memo control in the Memo Sizer
        box.Add(self.memo, 1, wx.EXPAND | wx.ALL, 6)
        
        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Put an expandable point in the button sizer so everything else can be right justified
        btnSizer.Add((1, 1), 1, wx.EXPAND)
        # Add an OK button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Put the OK button on the button sizer
        btnSizer.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        # Add a Help button
        btnHelp = wx.Button(self, -1, _("Help"))
        # Link the Help button to the OnHelp method
        btnHelp.Bind(wx.EVT_BUTTON, self.OnHelp)
        # Add the Help button to the Button sizer
        btnSizer.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        # Add the button sizer to the main sizer
        box.Add(btnSizer, 0, wx.EXPAND)

        # Attach the Form's Main Sizer to the form
        self.SetSizer(box)
        # Set AutoLayout on
        self.SetAutoLayout(True)
        # Lay out the form
        self.Layout()
        # Set the minimum size for the form.
        self.SetSizeHints(minW = 600, minH = 440)
        # Center the form on the screen
        self.CentreOnScreen()

        # Clear the Report of contents
        self.memo.Clear()
        # If no Clip ID is passed in ...
        if newClipID == None:
            # ... then use the original Clip ID.  It hasn't changed.
            newClipID = originalClip.id
        # If no new keyword list is passed in ...
        if newKeywordList == None:
            # ... then use the original keyword list.  It hasn't changed.
            newKeywordList = originalClip.keyword_list

        # ***************************************************
        #  Could this code be optimized by leaving the newKeywordList == None and skipping Keyword Processing if there is none to do?
        # ***************************************************

        # If we are passed a Transcript Index ...
        if sourceTranscriptIndex > -1:
            # ... we pass that index on to the Clip List ...
            clipListTranscriptIndex = sourceTranscriptIndex
            # Request clips that are copies of the current clip
            clipList = DBInterface.list_of_clip_copies(originalClip.id, originalClip.transcripts[clipListTranscriptIndex].source_transcript, originalClip.transcripts[clipListTranscriptIndex].clip_start, originalClip.transcripts[clipListTranscriptIndex].clip_stop)
        # ... otherwise we have to handle requests from multiple transcripts ...
        else:
            # Initialize the clipList
            clipList = []
            # Iterate through all the transcripts from the original clip
            for clipListTranscriptIndex in range(len(originalClip.transcripts)):
                # Request clips that are copies of the current clip for the indicated transcript
                tempClipList = DBInterface.list_of_clip_copies(originalClip.id, originalClip.transcripts[clipListTranscriptIndex].source_transcript, originalClip.transcripts[clipListTranscriptIndex].clip_start, originalClip.transcripts[clipListTranscriptIndex].clip_stop)
                # Iterate through the clips in this list
                for clipListData in tempClipList:
                    needToAdd = True
                    for clipListData2 in clipList:
                        if clipListData[:3] == clipListData2[:3]:
                            needToAdd = False
                            break
                    if needToAdd:
                        clipList.append(clipListData)

        # If we have NO Transcript Index ...
        if sourceTranscriptIndex == -1:
            # ... then we need to iterate through the clips in the clip list to find instances of Originating Clip that need to be removed
            for (num, colnum, clid, trnum) in clipList:
                # If the original clip is returned in the list of copies ...
                if (originalClip.number == num) and \
                   (originalClip.collection_num == colnum) and \
                   (originalClip.id == clid):
                    # ... then remove it from the list.  It's already been updated!
                    clipList.remove((num, colnum, clid, trnum))
        # If we HAVE a Transcript Index ...
        else:
            
            # If the original clip is returned in the list of copies ...
            if (originalClip.number, originalClip.collection_num, originalClip.id, originalClip.transcripts[sourceTranscriptIndex].number) in clipList:
                # ... then remove it from the list.  It's already been updated!
                clipList.remove((originalClip.number, originalClip.collection_num, originalClip.id, originalClip.transcripts[sourceTranscriptIndex].number))

        # If no clips are returned ...
        if len(clipList) == 0:
            # ... inform the user ...
            prompt = _("No clips were found that match %s.")
            if 'unicode' in wx.PlatformInfo:
                prompt = unicode(prompt, 'utf8')
            # ... by adding a message to the report
            self.memo.AppendText(prompt % originalClip.id)

        # Let's start a database Transaction
        # First, let's get a database cursor
        dbCursor = DBInterface.get_db().cursor()
        dbCursor.execute('BEGIN')
        # We need to track MU Messages to be sent.  This way, if the CANCEL button is pressed, we can skip the messages!
        messageCache = []
        
        # Initially, the user has not indicated that s/he wants to update ALL the clips.
        acceptAll = False
        # Initialize Results to ID-SKIP.  If no Clips are found, it's the same as if the user pressed "Skip".
        results = ID_SKIP
        # If the user cancels, we may need to undo some of the changes that get made to the system!
        undoData = []
        # Iterate through the Clip List ...
        for clip in clipList:
            # Start Exception handling.
            try:
                # ... load the current clip
                clipObj = Clip.Clip(clip[0])
                # Let's remember the original Clip ID.
                oldClipID = clipObj.id
                # If we have a Transcript Index ...
                if sourceTranscriptIndex > -1:
                    # We need the appropriate transcript's INDEX.  Initialize to -1
                    trIndex = -1
                    # Interate through transcripts ...
                    for tr in clipObj.transcripts:
                        # ... if the transcript's number is the one we're looking for (matches the SourceTranscript) ...
                        if tr.number == clip[3]:
                            # ... remember the INDEX for the Transcript, so we know which one to update.
                            trIndex = {clipObj.transcripts.index(tr) : newTranscriptText}
                            break
                # If we don't have a Transcript Index, we need to create a dictionary of transcripts!
                else:
                    # Initialize a dictionary
                    trIndex = {}
                    
                    # The trick here is to make sure that ALL COMBINATIONS of transcripts of same and different sizes get
                    # detected and processed.  If we update a clip with just transcripts 2 and 3, we need to be sure that
                    # clips with transcripts 1 and 3 as well as 2 and 4 get caught and processed.
                    
                    # Loop through all clipObj transcripts
                    for x in range(len(clipObj.transcripts)):
                        # Loop through all newTranscriptText transcripts
                        for y in range(len(newTranscriptText)):
                            # If both transcripts come from the same source transcript, we have a match.
                            if clipObj.transcripts[x].source_transcript == newTranscriptText[y].source_transcript:
                                # We need to record the matches for processing.
                                trIndex[x] = newTranscriptText[y].text
                                # Once we've found the match, no need to keep looking within newTranscriptText.
                                break

                # For each key in the transcripts dictionary ...
                for key in trIndex.keys():
                    # If the user hasn't signal that they want to update all clips AND we've found an eligible transcript (index != -1) ...
                    if (not acceptAll) and (trIndex != -1):
                        # See if we're looking at single transcript situations OR the first transcript in a list ...
                        if (sourceTranscriptIndex > -1) or (key == 0):
                            # ... create the Accept Clip Transcript Changes form WITH KEYWORDS ...
                            acceptClipChanges = AcceptClipTranscriptChanges(self, clip[0], key, trIndex[key], newKeywordList, helpString="Propagate Clip Changes")
                        # If it's not a single or the first, ...
                        else:
                            # ... LEAVE THE KEYWORDS OFF!
                            acceptClipChanges = AcceptClipTranscriptChanges(self, clip[0], key, trIndex[key], helpString="Propagate Clip Changes")
                        # ... and display it, capturing the user feedback.
                        results = acceptClipChanges.GetResults()
                        # Close (and Destroy) the form
                        acceptClipChanges.Destroy()
                        # If the user wants to update all Clips ...
                        if results == ID_UPDATEALL:
                            # ... then changing this variable will prevent the need for further user intervention.
                            acceptAll = True
                    # If the user presses "Update" (OK) or has pressed "Update All" ...
                    if acceptAll or (results == wx.ID_OK):
                        # Get a list of Keyword Examples for the current clip.  We don't want to lose this information in the propagation.
                        keywordExamples = DBInterface.list_all_keyword_examples_for_a_clip(clip[0])
                        # Lock the Clip (will raise an exception of you can't)
                        clipObj.lock_record()
                        # update the Clip ID
                        clipObj.id = newClipID
                        # substitute the new transcript text for the old text
                        clipObj.transcripts[key].text = trIndex[key]
                        if (sourceTranscriptIndex > -1) or (key == 0):
                            # Clear the old keywords from the clip
                            clipObj.clear_keywords()
                            # Iterate through all the clips in the NEW keyword list
                            for clipKeyword in newKeywordList:
                                # If the keyword is also in the Keyword Examples list ...
                                if (clipKeyword.keywordGroup, clipKeyword.keyword, clipObj.number, originalClip.id) in keywordExamples:
                                    # ... add it as an example keyword
                                    clipObj.add_keyword(clipKeyword.keywordGroup, clipKeyword.keyword, example=1)
                                # If it's NOT an example ...
                                else:
                                    # ... add it as a non-example keyword
                                    clipObj.add_keyword(clipKeyword.keywordGroup, clipKeyword.keyword)
                            # We don't want Keyword Example clips to get lost because of propagation.
                            # Therefore, iterate through the Keyword Examples list
                            for kw in keywordExamples:
                                # Check to see if the keyword is already added to the clip above.  If NOT ...
                                if not clipObj.has_keyword(kw[0], kw[1]):
                                    # ... add the keyword as an example keyword
                                    clipObj.add_keyword(kw[0], kw[1], example=1)
                                    # This will be unexpected by the user, so let's add a note to the user about this.
                                    prompt = _('Clip "%s" in collection "%s" retained keyword "%s : %s" because it is a keyword example.') + '\n\n'
                                    if 'unicode' in wx.PlatformInfo:
                                        prompt = unicode(prompt, 'utf8')
                                    # Add the prompt to the memo to communicate this to the user.
                                    self.memo.AppendText(prompt % (oldClipID, clipObj.GetNodeString(False), kw[0], kw[1]))

                        # Save the Clip
                        clipObj.db_save()
                        # unlock the clip
                        clipObj.unlock_record()
                        
                        # Finally, indicate success in the Report
                        prompt = _('Transcript %d of clip "%s" in collection "%s" has been updated.')
                        if 'unicode' in wx.PlatformInfo:
                            prompt = unicode(prompt, 'utf8')
                        # Add the message to the report
                        self.memo.AppendText(prompt % (key + 1, oldClipID, clipObj.GetNodeString(False)) + '\n\n')
                    # If the user indicates we should skip ONE clip ...
                    elif results == ID_SKIP:
                        # ... indicate that in the report and don't do anything else.
                        prompt = _('Clip change has been skipped for transcript %d of clip "%s" in collection "%s".')
                        if 'unicode' in wx.PlatformInfo:
                            prompt = unicode(prompt, 'utf8')
                        self.memo.AppendText(prompt % (key + 1, oldClipID, clipObj.GetNodeString(False)) + '\n\n')
                    # If the user indicates we should CANCEL Clip propagation ...
                    elif results == wx.ID_CANCEL:
                        # ... indicate that in the report.  The rest of Cancel is implemented later, after we've added
                        #     the full Clip information to the report.
                        self.memo.AppendText(_("Clip change propagation has been cancelled."))
                        # ... we should STOP processing Clips!  So stop iterating!
                        break
                # Initialize the Chat Message
                msg = ""
                # If the Clip ID has changed, we need to update instances of the Clip ID
                if oldClipID != clipObj.id:
                    # Iterate through all the Keyowrd Examples for this clip ...
                    for (kwg, kw, clipNumber, clipID) in keywordExamples:
                        # Build the keyword Example Node List for the OLD Clip ID
                        nodeList = (_('Keywords'), kwg, kw, originalClip.id)
                        # Select the Keyword Example Node
                        exampleNode = parent.tree.select_Node(nodeList, 'KeywordExampleNode')
                        # Update the Keyword Example Node to the NEW Clip ID
                        parent.tree.SetItemText(exampleNode, clipObj.id)
                        # Add a Keyword Example record to the Undo list in case the user Cancels.
                        undoData.append(('KWE', clipObj.number, clipObj.episode_num, clipObj.id, originalClip.id, kwg, kw))
                        # If we're in the Multi-User mode, we need to send a message about the change
                        if not TransanaConstants.singleUserVersion:
                            # Begin constructing the message with the old and new names for the node
                            msg = " >|< %s >|< %s" % (originalClip.id, clipObj.id)
                            # Get the full Node Branch by climbing it to two levels above the root
                            while (parent.tree.GetItemParent(parent.tree.GetItemParent(exampleNode)) != parent.tree.GetRootItem()):
                                # Update the selected node indicator
                                exampleNode = parent.tree.GetItemParent(exampleNode)
                                # Prepend the new Node's name on the Message with the appropriate seperator
                                msg = ' >|< ' + parent.tree.GetItemText(exampleNode) + msg
                            # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                            # This must be untranslated to avoid problems in mixed-language environments.
                            # Prepend these on the Messsage
                            msg = "KeywordExampleNode >|< Keywords" + msg

                            if DEBUG:
                                print 'Message to send = "RN %s"' % msg

                            # Cache the Rename Node message for later processing
                            if TransanaGlobal.chatWindow != None:
                                messageCache.append("RN %s" % msg)

                    # now build the Node List for the Clip itself.  First get the Node data from the Clip ...
                    clipNodeData = clipObj.GetNodeData(False)
                    # ... then add the Collections Root and the original Clip ID to make the full Node List.
                    nodeList = (_("Collections"),) + clipNodeData + (originalClip.id,)
                    # Rename the Tree node
                    parent.tree.rename_Node(nodeList, 'ClipNode', clipObj.id)
                    # Add a Clip record to the Undo list in case the user Cancels.
                    undoData.append(('Clip', clipObj.number, clipObj.episode_num, clipObj.id, originalClip.id, clipNodeData))
                    # If we're in the Multi-User mode, we need to send a message about the change
                    if not TransanaConstants.singleUserVersion:
                        # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                        # This must be untranslated to avoid problems in mixed-language environments.
                        # Prepend these on the Messsage
                        msg = "ClipNode >|< Collections >|< "
                        for node in nodeList[1:]:
                            # Prepend the new Node's name on the Message with the appropriate seperator
                            msg += node + ' >|< ' 
                        # Begin constructing the message with the old and new names for the node
                        msg += clipObj.id

                        if DEBUG:
                            print 'Message to send = "RN %s"' % msg.encode('latin1')
                            print

                        # Cache the Rename Node message for later processing
                        if TransanaGlobal.chatWindow != None:
                            messageCache.append("RN %s" % msg)

                # See if the Keyword visualization needs to be updated.
                parent.ControlObject.UpdateKeywordVisualization()
                            
                # Now let's communicate with other Transana instances if we're in Multi-user mode
                if not TransanaConstants.singleUserVersion:
                    # Build the message to update Keyword Visualizations.
                    msg = 'Clip %d %d' % (clipObj.number, clipObj.episode_num)

                    if DEBUG:
                        print 'Message to send = "UKL %s"' % msg

                    if TransanaGlobal.chatWindow != None:
                        # Cache the Update Keyword messages for later processing
                        messageCache.append("UKL %s" % msg)
                        messageCache.append("UKV %s" % msg)

            # If a RecordLocked Exception is raised ...
            except TransanaExceptions.RecordLockedError, e:
                # ... indicate that in the report.
                prompt = _('ERROR: Clip "%s" in collection "%s" is locked and cannot be updated.')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                self.memo.AppendText(prompt % (oldClipID, clipObj.GetNodeString(False)) + '\n\n')

            # If a SaveError exception is raised -- Duplicate Clip ID error, for example
            except TransanaExceptions.SaveError, e:
                # build the prompt for the user ...
                prompt = _('ERROR: Save error "%s" for clip "%s" in collection "%s"')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                # ... and report the error in the report.
                self.memo.AppendText(prompt % (e.reason, oldClipID, clipObj.GetNodeString(False)) + '\n\n')
                # unlock the clip
                clipObj.unlock_record()

            except:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)

            # If the user pressed "Cancel" ...
            if results == wx.ID_CANCEL:
                # ... we should STOP processing Clips!  So stop iterating!
                break

        # If the user pressed Cancel ...
        if results == wx.ID_CANCEL:
            # ... undo Clip Transcript changes in the database ...
            dbCursor.execute("ROLLBACK")

            # If the user presses Cancel, we have to reverse the changes that have already been made to the
            # user interface (local and MU).  This probably isn't a great model (change, then undo if Cancelled),
            # but cancel should be rare.
            # So first, iterate through the Undo list to see what records need to be undone.
            for undoRec in undoData:
                # If the record is a Clip record ...
                if undoRec[0] == 'Clip':
                    # ... assemble the node list based on the values in the Undo record
                    nodeList = (_("Collections"),) + undoRec[5] + (undoRec[3],)
                    # Rename the correct Tree node
                    parent.tree.rename_Node(nodeList, 'ClipNode', undoRec[4])

                    # If we're in the Multi-User mode, we need to send a message about the change
                    if not TransanaConstants.singleUserVersion:
                        # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                        # This must be untranslated to avoid problems in mixed-language environments.
                        # Prepend these on the Messsage
                        msg = "ClipNode >|< Collections >|< "
                        for node in nodeList[1:]:
                            # Prepend the new Node's name on the Message with the appropriate seperator
                            msg += node + ' >|< ' 
                        # Begin constructing the message with the old and new names for the node
                        msg += undoRec[4]

                        if DEBUG:
                            print 'Message to send = "RN %s"' % msg.encode('latin1')
                            print

                        # Cache the Rename Node message for later processing
                        if TransanaGlobal.chatWindow != None:
                            messageCache.append("RN %s" % msg)

                # If the record is a Keyword Example record ...
                else:
                    # .. Build the Keyword Example Node List for the OLD Clip ID
                    nodeList = (_('Keywords'), undoRec[5], undoRec[6], undoRec[3])
                    # Select the Keyword Example Node
                    exampleNode = parent.tree.select_Node(nodeList, 'KeywordExampleNode')
                    # Update the Keyword Example Node to the NEW Clip ID
                    parent.tree.SetItemText(exampleNode, undoRec[4])
                    # If we're in the Multi-User mode, we need to send a message about the change
                    if not TransanaConstants.singleUserVersion:
                        # Begin constructing the message with the old and new names for the node
                        msg = " >|< %s >|< %s" % (undoRec[2], undoRec[3])
                        # Get the full Node Branch by climbing it to two levels above the root
                        while (parent.tree.GetItemParent(parent.tree.GetItemParent(exampleNode)) != parent.tree.GetRootItem()):
                            # Update the selected node indicator
                            exampleNode = parent.tree.GetItemParent(exampleNode)
                            # Prepend the new Node's name on the Message with the appropriate seperator
                            msg = ' >|< ' + parent.tree.GetItemText(exampleNode) + msg
                        # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                        # This must be untranslated to avoid problems in mixed-language environments.
                        # Prepend these on the Messsage
                        msg = "KeywordExampleNode >|< Keywords" + msg

                        if DEBUG:
                            print 'Message to send = "RN %s"' % msg

                        # Cache the Rename Node message for later processing
                        if TransanaGlobal.chatWindow != None:
                            messageCache.append("RN %s" % msg)

                # Now let's communicate with other Transana instances if we're in Multi-user mode
                if not TransanaConstants.singleUserVersion:
                    # Build the message to update Keyword Visualizations.
                    msg = 'Clip %d' % undoRec[1]

                    if DEBUG:
                        print 'Message to send = "UKL %s"' % msg

                    if TransanaGlobal.chatWindow != None:
                        # Cache the Update Keywords messages for later processing
                        messageCache.append("UKL %s" % msg)
                        messageCache.append("UKV %s %s" % (msg, undoRec[2]))
                
            # ... and inform the user that the changes were cancelled.  To avoid confusion, clear the report information already generated.
            self.memo.Clear()
            self.memo.AppendText(_("Clip Change Propagation was cancelled.  Therefore, no clips were updated."))
        # If the user pressed anything except Cancel ...
        else:
            # ... commit the changes to the database
            dbCursor.execute("COMMIT")

            # If we're in MU and have a chat window ...
            if TransanaGlobal.chatWindow != None:
                # Iterate through the cached messages
                for message in messageCache:
                    # Send the messages one at a time
                    TransanaGlobal.chatWindow.SendMessage(message)
            # If the current clip needs to be updated ...
            if CurrentClipToUpdate != 0:
                # ... then update the current clip to the latest data
                parent.ControlObject.LoadClipByNumber(CurrentClipToUpdate)
                    
        # Update the contents of the memo
        self.memo.Update()
        # Finally, we can show the dialog so the user can see the report!
        self.ShowModal()

    def OnHelp(self, event):
        """ Method to use when the Help Button is pressed """
        # Locate the Menu Window ...
        if TransanaGlobal.menuWindow != None:
            # .. and use that to get to the standard Help Call method.
            TransanaGlobal.menuWindow.ControlObject.Help("Propagate Clip Changes")
        

class AcceptClipTranscriptChanges(wx.Dialog):
    """ This dialog shows the user proposed Clip changes due to Transcript Change Propagation and asks for approval. """
    def __init__(self, parent, clipNum, clipTranscriptNum, newText, newKeywordList=None, helpString="Propagate Transcript Changes"):
        """ Initialization Parameters are parent window, the clip Number, the proposed NEW clip transcript text,
            [the new Keyword list], and [the Help ID] """
        # Remember the Help ID for use if the Help button is pressed.
        self.helpString = helpString
        # Load the Clip Object
        tempClip = Clip.Clip(clipNum)
        # Define the main Frame for the Accept Clip Transcript Changes Window
        wx.Dialog.__init__(self, parent, -1, _("Review Transcript Changes"), size = (610,550), style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|wx.NO_FULL_REPAINT_ON_RESIZE)
        # Set the background to White
        self.SetBackgroundColour(wx.WHITE)
        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create a Sizer for the form
        box = wx.BoxSizer(wx.VERTICAL)

        # Create a Horizontal Sizer for the first row
        hBox1 = wx.BoxSizer(wx.HORIZONTAL)
        # Create a Vertical Sizer for the data element Header and data element
        vBox1 = wx.BoxSizer(wx.VERTICAL)
        # Create the Clip ID label
        clipIDHdr = wx.StaticText(self, -1, _("Clip ID"))
        # Add the label to the data element sizer
        vBox1.Add(clipIDHdr, 0)
        # Create the Clip ID data box
        clipID = wx.TextCtrl(self, -1, tempClip.id)
        # Disable the data box
        clipID.Enable(False)
        # Add the data element to the data element sizer
        vBox1.Add(clipID, 0, wx.EXPAND)
        # Create a Vertical Sizer for the data element Header and data element
        vBox2 = wx.BoxSizer(wx.VERTICAL)
        # Create the Collection ID label
        collectionIDHdr = wx.StaticText(self, -1, _("Collection ID"))
        # Add the label to the data element sizer
        vBox2.Add(collectionIDHdr, 0)
        # Create the Collection ID data box
        collectionID = wx.TextCtrl(self, -1, tempClip.GetNodeString(False))
        # Disable the data box
        collectionID.Enable(False)
        # Add the data element to the data element sizer
        vBox2.Add(collectionID, 0, wx.EXPAND)

        # Add the Clip ID data element Sizer to the first row Horizontal Sizer
        hBox1.Add(vBox1, 1, wx.ALL | wx.EXPAND, 6)
        # Add the Collection ID data element Sizer to the first row Horizontal Sizer
        hBox1.Add(vBox2, 2, wx.ALL | wx.EXPAND, 6)
        # Add the first row Horizontal Sizer to the Main Sizer
        box.Add(hBox1, 0, wx.EXPAND)

        # Create a Horizontal Sizer for the second row
        hBox2 = wx.BoxSizer(wx.HORIZONTAL)
        # Create a Vertical Sizer for the data element Header and data element
        vBox3 = wx.BoxSizer(wx.VERTICAL)
        # Create the Clip Start label
        clipStartHdr = wx.StaticText(self, -1, _("Clip Start"))
        # Add the label to the data element sizer
        vBox3.Add(clipStartHdr, 0)
        # Create the Clip Start data box
        clipStart = wx.TextCtrl(self, -1, Misc.time_in_ms_to_str(tempClip.clip_start))
        # Disable the data box
        clipStart.Enable(False)
        # Add the data element to the data element sizer
        vBox3.Add(clipStart, 0, wx.EXPAND)
        # Create a Vertical Sizer for the data element Header and data element
        vBox4 = wx.BoxSizer(wx.VERTICAL)
        # Create the Clip Stop label
        clipStopHdr = wx.StaticText(self, -1, _("Clip Stop"))
        # Add the label to the data element sizer
        vBox4.Add(clipStopHdr, 0)
        # Create the Clip Stop data box
        clipStop = wx.TextCtrl(self, -1, Misc.time_in_ms_to_str(tempClip.clip_stop))
        # Disable the data box
        clipStop.Enable(False)
        # Add the data element to the data element sizer
        vBox4.Add(clipStop, 0, wx.EXPAND)
        # Create a Vertical Sizer for the data element Header and data element
        vBox5 = wx.BoxSizer(wx.VERTICAL)
        # Create the Clip Length label
        clipLengthHdr = wx.StaticText(self, -1, _("Clip Length"))
        # Add the label to the data element sizer
        vBox5.Add(clipLengthHdr, 0)
        # Create the Clip Length data box
        clipLength = wx.TextCtrl(self, -1, Misc.time_in_ms_to_str(tempClip.clip_stop - tempClip.clip_start))
        # Disable the data box
        clipLength.Enable(False)
        # Add the data element to the data element sizer
        vBox5.Add(clipLength, 0, wx.EXPAND)

        # Add the Clip Start data element Sizer to the second row Horizontal Sizer
        hBox2.Add(vBox3, 1, wx.ALL | wx.EXPAND, 6)
        # Add the Clip Stop data element Sizer to the second row Horizontal Sizer
        hBox2.Add(vBox4, 1, wx.ALL | wx.EXPAND, 6)
        # Add the Clip Length data element Sizer to the second row Horizontal Sizer
        hBox2.Add(vBox5, 1, wx.ALL | wx.EXPAND, 6)
        # Add the second row Horizontal Sizer to the Main Sizer
        box.Add(hBox2, 0, wx.EXPAND)

        # Create a horizontal band for the current clip data
        hBox3 = wx.BoxSizer(wx.HORIZONTAL)
        # Create a vertical sizer for the Transcript
        vBox6 = wx.BoxSizer(wx.VERTICAL)

        # Create the Current Transcript label
        oldTranscriptHdr = wx.StaticText(self, -1, _("Current Transcript"))
        # Put the header on the Transcript sizer
        vBox6.Add(oldTranscriptHdr, 0)

        # Load the Old Transcript into an RTF Control so the RTF Encoding won't show
        oldTranscript = TranscriptEditor.TranscriptEditor(self)
        # Make the old transcript writable so it can be populated
        oldTranscript.SetReadOnly(False)
        # Set up the Progess Dialog
        oldTranscript.ProgressDlg = wx.ProgressDialog("Loading Transcript", \
                                               "Reading document stream", \
                                                maximum=100, \
                                                style=wx.PD_AUTO_HIDE)
        # Load the Old Transcript from the temporary clip
        oldTranscript.LoadRTFData(tempClip.transcripts[clipTranscriptNum].text)
        # Clean up the Progress Dialog
        oldTranscript.ProgressDlg.Destroy()
        # This doesn't work!  Hidden text remains visible.
#        oldTranscript.StyleSetVisible(oldTranscript.STYLE_HIDDEN, False)
        # Set the Visibility flag
        oldTranscript.codes_vis = 0
        # Scan transcript for Time Codes
        oldTranscript.load_timecodes()
        # Display the time codes
        oldTranscript.show_codes()
        # Make the old transcript Read Only
        oldTranscript.SetReadOnly(True)
        # Add the old transcript to the transcript sizer
        vBox6.Add(oldTranscript, 1, wx.EXPAND)
        # Add the transcript sizer to the current clip data sizer
        hBox3.Add(vBox6, 3, wx.EXPAND)
        # If there are keywords to be displayed ...
        if newKeywordList != None:
            # Create a vertical sizer for the keywords
            vBox7 = wx.BoxSizer(wx.VERTICAL)
            # Create a label for the keywords
            oldKeywordListHdr = wx.StaticText(self, -1, _("Current Keywords"))
            # Put the label on the keywords sizer
            vBox7.Add(oldKeywordListHdr, 0)
            # Create a list box for the keywords
            oldKeywords = wx.ListBox(self, -1)
            # Add the list box to the keyword sizer
            vBox7.Add(oldKeywords, 1, wx.EXPAND)
            # Add the keyword sizer to the current clip band
            hBox3.Add(vBox7, 1, wx.EXPAND | wx.LEFT, 6)
            # Iterate through the current clip's keyword list ...
            for clipKeyword in tempClip.keyword_list:
                # ... and add the keywords to the keyword list
                oldKeywords.Append(clipKeyword.keywordPair)

        # Add the current clip sizer to the main sizer
        box.Add(hBox3, 1, wx.EXPAND | wx.ALL, 6)

        # Create a horizontal band for the new clip data
        hBox4 = wx.BoxSizer(wx.HORIZONTAL)
        # Create a vertical sizer for the Transcript
        vBox8 = wx.BoxSizer(wx.VERTICAL)
        
        # Create the New Transcript label
        newTranscriptHdr = wx.StaticText(self, -1, _("New Transcript"))
        # Add the label to the main sizer
        vBox8.Add(newTranscriptHdr, 0)

        # Load the Transcript into an RTF Control so the RTF Encoding won't show
        newTranscript = TranscriptEditor.TranscriptEditor(self)
        # Make the new transcript writable so it can be populated
        newTranscript.SetReadOnly(False)
        # Set up the Progess Dialog
        newTranscript.ProgressDlg = wx.ProgressDialog("Loading Transcript", \
                                               "Reading document stream", \
                                                maximum=100, \
                                                style=wx.PD_AUTO_HIDE)
        # Load the Old Transcript from the method parameter
        newTranscript.LoadRTFData(newText)
        # Clean up the Progress Dialog
        newTranscript.ProgressDlg.Destroy()

        # This doesn't work!  Hidden text remains visible.
#        newTranscript.StyleSetVisible(newTranscript.STYLE_HIDDEN, False)
        # Set the Visibility flag
        newTranscript.codes_vis = 0
        # Scan transcript for Time Codes
        newTranscript.load_timecodes()
        # Display the time codes
        newTranscript.show_codes()
        # Make the new transcript Read Only
        newTranscript.SetReadOnly(True)

        # Add the new transcript to the transcript sizer
        vBox8.Add(newTranscript, 1, wx.EXPAND)
        # Add the transcript sizer to the new clip data sizer
        hBox4.Add(vBox8, 3, wx.EXPAND)
        # If there are keywords to be displayed ...
        if newKeywordList != None:
            # Create a vertical sizer for the keywords
            vBox9 = wx.BoxSizer(wx.VERTICAL)
            # Create a label for the keywords
            newKeywordListHdr = wx.StaticText(self, -1, _("New Keywords"))
            # Put the label on the keywords sizer
            vBox9.Add(newKeywordListHdr, 0)
            # Create a list box for the keywords
            newKeywords = wx.ListBox(self, -1)
            # Add the list box to the keyword sizer
            vBox9.Add(newKeywords, 1, wx.EXPAND)
            # Add the keyword sizer to the new clip band
            hBox4.Add(vBox9, 1, wx.EXPAND | wx.LEFT, 6)
            # Iterate through the new clip's keyword list ...
            for clipKeyword in newKeywordList:
                # ... and add the keywords to the keyword list
                newKeywords.Append(clipKeyword.keywordPair)

        # Add the new clip sizer to the main sizer
        box.Add(hBox4, 1, wx.EXPAND | wx.ALL, 6)

        # Create a Horizontal Sizer for the Buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Put an expandable point in the button sizer so everything else can be right justified
        btnSizer.Add((1, 1), 1, wx.EXPAND)
        # Add the "Update" button, using the OK button ID
        btnOK = wx.Button(self, wx.ID_OK, _("Update"))
        # Bind the Update button to the button handler
        btnOK.Bind(wx.EVT_BUTTON, self.OnButton)
        # Add the Update button to the Button Sizer
        btnSizer.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        # Add the "Update All" button, using the UPDATEALL button ID
        btnAcceptAll = wx.Button(self, ID_UPDATEALL, _("Update All"))
        # Bind the Update All button to the button handler
        btnAcceptAll.Bind(wx.EVT_BUTTON, self.OnButton)
        # Add the Update All button to the Button Sizer
        btnSizer.Add(btnAcceptAll, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        # Add the "Skip" button, using the SKIP button ID
        btnSkip = wx.Button(self, ID_SKIP, _("Skip"))
        # Bind the Skip button to the button handler
        btnSkip.Bind(wx.EVT_BUTTON, self.OnButton)
        # Add the Skip button to the Button Sizer
        btnSizer.Add(btnSkip, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        # Add the "Cancel" button, using the CANCEL button ID
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        # Bind the Cancel button to the button handler
        btnCancel.Bind(wx.EVT_BUTTON, self.OnButton)
        # Add the Cancel button to the Button Sizer
        btnSizer.Add(btnCancel, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        box.Add(btnSizer, 0, wx.ALIGN_BOTTOM | wx.EXPAND)
        # Add the "Help" button
        btnHelp = wx.Button(self, -1, _("Help"))
        # Bind the Help button to the Help button handler
        btnHelp.Bind(wx.EVT_BUTTON, self.OnHelp)
        # Add the Help button to the Button Sizer
        btnSizer.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)

        # initialize the Results flag to CANCEL
        self.results = wx.ID_CANCEL

        # Attach the Form's Main Sizer to the form
        self.SetSizer(box)
        # Set AutoLayout on
        self.SetAutoLayout(True)
        # Lay out the form
        self.Layout()
        # Set the minimum size for the form.
        self.SetSizeHints(minW = 500, minH = 340)
        # Center the form on the screen
        self.CentreOnScreen()

    def OnButton(self, event):
        """ Button Pressed handler for most buttons """
        # Set the Results flag to the calling Button's ID
        self.results = event.GetId()
        # Close the form
        self.Close()
        
    def OnHelp(self, event):
        """ Method to use when the Help Button is pressed """
        # Locate the Menu Window ...
        if TransanaGlobal.menuWindow != None:
            # .. and use that to get to the standard Help Call method.
            TransanaGlobal.menuWindow.ControlObject.Help(self.helpString)
        
    def GetResults(self):
        """ The Accept Clip Transcript Changes form's main processing routine """
        # Display the form modally
        self.ShowModal()
        # Return the ID of the button that was pressed
        return self.results
