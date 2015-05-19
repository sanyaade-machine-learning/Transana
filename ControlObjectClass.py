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

"""This module implements the Control Object class for Transana,
which is responsible for managing communication between the
four main windows.  Each object (Menu, Visualization, Video, Transcript,
and Data) should communicate only with the Control Object, not with
each other.
"""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare'

DEBUG = False
if DEBUG:
    print "ControlObjectClass DEBUG is ON!"

# Import wxPython
import wx

# import Transana's Constants
import TransanaConstants
# Import the Menu Constants
import MenuSetup
# Import Transana's Global Values
import TransanaGlobal
# import the Transana Series Object definition
import Series
# import the Transana Episode Object definition
import Episode
# import the Transana Transcript Object definition
import Transcript
# import the Transana Collection Object definition
import Collection
# import the Transana Clip Object definition
import Clip
# import the Transana Miscellaneous Routines
import Misc
# import Transana Database Interface
import DBInterface
# import Transana's Dialogs
import Dialogs
# import Transana's DragAndDrop Objects for Quick Clip creation
import DragAndDropObjects
# import Transana File Management System
import FileManagement
# import Play All Clips
import PlayAllClips
# import the Episode Transcript Change Propagation tool
import PropagateEpisodeChanges
# import Transana's Exceptions
import TransanaExceptions
# Import Transana's Transcript User Interface for creating supplemental Transcript Windows
if TransanaConstants.USESRTC:
    import TranscriptionUI_RTC as TranscriptionUI
else:
    import TranscriptionUI
# import Python's os module
import os
# import Python's sys module
import sys
# import Python's pickle module
import pickle


class ControlObject(object):
    """ The ControlObject operationalizes all inter-window and inter-object communication and control.
        All objects should speak only to the ControlObject, not to each other directly.  The purpose of
        this is to allow greater modularity of code, so that modules can be swapped in and out in with
        changes affecting only this object if the APIs change.  """
    def __init__(self):
        """ Initialize the ControlObject """
        # Define Objects that need controlling (initializing to None)
        self.MenuWindow = None
        self.VideoWindow = None
        # There may be multiple Transcript Windows.  We'll use a List to keep track of them.
        self.TranscriptWindow = []
        # We need to know what transcript is "Active" (most recently selected) at any given point.  -1 signals none.
        self.activeTranscript = -1
        self.VisualizationWindow = None
        self.DataWindow = None
        self.PlayAllClipsWindow = None
        self.NotesBrowserWindow = None
        self.ChatWindow = None

        # Initialize variables
        self.VideoFilename = ''         # Video File Name
        self.VideoStartPoint = 0        # Starting Point for video playback in Milliseconds
        self.VideoEndPoint = 0          # Ending Point for video playback in Milliseconds
        self.WindowPositions = []       # Initial Screen Positions for all Windows, used for Presentation Mode
        self.TranscriptNum = []         # Transcript record # LIST loaded
        self.currentObj = None          # Currently loaded Object (Episode or Clip)
        # Have the Export Directory default to the Video Root, but then remember its changed value for the session
        self.defaultExportDir = TransanaGlobal.configData.videoPath
        self.playInLoop = False         # Should we loop playback?
        self.LoopPresMode = None        # What presentation mode are we ignoring while Looping?
        self.shuttingDown = False       # We need to signal when we want to shut down to prevent problems
                                        # with the Visualization Window's IDLE event trying to call the
                                        # VideoWindow after it's been destroyed.
        
    def Register(self, Menu='', Video='', Transcript='', Data='', Visualization='', PlayAllClips='', NotesBrowser='', Chat=''):
        """ The ControlObject can extert control only over those objects it knows about.  This method
            provides a way to let the ControlObject know about other objects.  This infrastructure allows
            for objects to be swapped in and out.  For example, if you need a different video window
            that supports a format not available on the current one, you can hide the current one, show
            a new one, and register that new one with the ControlObject.  Once this is done, the new
            player will handle all tasks for the program.  """
        # This function expects parameters passed by name and "registers" the components that
        # need to be available to the ControlObject to be controlled.  To remove an
        # object registration, pass in "None"
        if Menu != '':
            self.MenuWindow = Menu                       # Define the Menu Window Object
        if Video != '':
            self.VideoWindow = Video                     # Define the Video Window Object
        if Transcript != '':
            # Add the Transcript Window reference to the list of Transcript Windows
            self.TranscriptWindow.append(Transcript)
            # Add the Transcript Number to the list of Transcript Numbers
            self.TranscriptNum.append(0)
            # Set the new Transcript to be the Active Transcript
            self.activeTranscript = len(self.TranscriptWindow) - 1
        if Data != '':
            self.DataWindow = Data                       # Define the Data Window Object
        if Visualization != '':
            self.VisualizationWindow = Visualization     # Define the Visualization Window Object
        if PlayAllClips != '':
            self.PlayAllClipsWindow = PlayAllClips       # Define the Play All Clips Window Object
        if NotesBrowser != '':
            self.NotesBrowserWindow = NotesBrowser             # Define the Notes Browser Window Object
        if Chat != '':
            self.ChatWindow = Chat                       # Define the Chat Window Object

    def CloseAll(self):
        """ This method closes all application windows and cleans up objects when the user
            quits Transana. """
        # Closing the MenuWindow will automatically close the Transcript, Data, and Visualization
        # Windows in the current setup of Transana, as these windows are all defined as child dialogs
        # of the MenuWindow.
        self.MenuWindow.Close()
        # VideoWindow needs to be closed explicitly.
        self.VideoWindow.close()

    def LoadTranscript(self, series, episode, transcript):
        """ When a Transcript is identified to trigger systemic loading of all related information,
            this method should be called so that all Transana Objects are set appropriately. """
        # Before we do anything else, let's save the current transcript if it's been modified.
        if self.TranscriptWindow[self.activeTranscript].TranscriptModified():
            self.SaveTranscript(1, cleardoc=1)
        # activeTranscript 0 signals we should reset everything in the interface!
        if self.activeTranscript == 0:
            clearAll = True
        else:
            clearAll = False
        # Clear all Windows
        self.ClearAllWindows(resetMultipleTranscripts = clearAll)
        # Because transcript names can be identical for different episodes in different series, all parameters are mandatory.
        # They are:
        #   series      -  the Series associated with the desired Transcript
        #   episode     -  the Episode associated with the desired Transcript
        #   transcript  -  the Transcript to be displayed in the Transcript Window
        seriesObj = Series.Series(series)                                    # Load the Series which owns the Episode which owns the Transcript
        episodeObj = Episode.Episode(series=seriesObj.id, episode=episode)   # Load the Episode in the Series that owns the Transcript
        # Set the current object to the loaded Episode
        self.currentObj = episodeObj
        transcriptObj = Transcript.Transcript(transcript, ep=episodeObj.number)

        # Load the Transcript in the Episode in the Series
        # reset the video start and end points
        self.VideoStartPoint = 0                                     # Set the Video Start Point to the beginning of the video
        self.VideoEndPoint = 0                                       # Set the Video End Point to 0, indicating that the video should not end prematurely
        
        # Remove any tabs in the Data Window beyond the Database Tab
        self.DataWindow.DeleteTabs()

        if self.LoadVideo(self.currentObj):    # Load the video identified in the Episode
            # Delineate the appropriate start and end points for Video Control.  (Required to prevent Waveform Visualization problems)
            self.SetVideoSelection(0, 0)

            # Force the Visualization to load here.  This ensures that the Episode visualization is shown
            # rather than the Clip visualization when Locating a Clip
            self.VisualizationWindow.OnIdle(None)
            
            # Identify the loaded Object
            prompt = _('Transcript "%s" for Series "%s", Episode "%s"')
            if self.activeTranscript > 0:
                prompt = '** ' + prompt + ' **'
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Set the window's prompt
            self.TranscriptWindow[self.activeTranscript].dlg.SetTitle(prompt % (transcriptObj.id, seriesObj.id, episodeObj.id))
            # If we have only one video file ...
            if len(self.currentObj.additional_media_files) == 0:
                # Identify the loaded media file
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Video Media File: "%s"'), 'utf8')
                else:
                    prompt = _('Video Media File: "%s"')
                # Place the file name in the video window's Title bar
                self.VideoWindow.SetTitle(prompt % episodeObj.media_filename)
            # If there are multiple videos ...
            else:
                # Just label the video window generically.  There's not room for file names.
                self.VideoWindow.SetTitle(_("Video"))
            # Open Transcript in Transcript Window
            self.TranscriptWindow[self.activeTranscript].LoadTranscript(transcriptObj) #flies off to transcriptionui.py
            # Add the Transcript Number to the list that tracks the numbers of the open transcripts
            self.TranscriptNum[self.activeTranscript] = transcriptObj.number
            
            # Add the Episode Clips Tab to the DataWindow
            self.DataWindow.AddEpisodeClipsTab(seriesObj=seriesObj, episodeObj=episodeObj)

            # Add the Selected Episode Clips Tab, initially set to the beginning of the video file
            # TODO:  When the Transcript Window updates the selected text, we need to update this tab in the Data Window!
            self.DataWindow.AddSelectedEpisodeClipsTab(seriesObj=seriesObj, episodeObj=episodeObj, TimeCode=0)

            # Add the Keyword Tab to the DataWindow
            self.DataWindow.AddKeywordsTab(seriesObj=seriesObj, episodeObj=episodeObj)
            # Enable the transcript menu item options
            self.MenuWindow.SetTranscriptOptions(True)

            if TransanaConstants.USESRTC:
                # After two seconds, call the EditorPaint method of the Transcript Dialog (in the TranscriptionUI_RTC file)
                # This causes improperly placed line numers to "correct" themselves!
                wx.CallLater(2000, self.TranscriptWindow[self.activeTranscript].dlg.EditorPaint, None)

             # Set focus to the new Transcript's Editor (so that CommonKeys work on the Mac)
            self.TranscriptWindow[self.activeTranscript].dlg.editor.SetFocus()
        # If the video won't load ...
        else:
            # Clear the interface!
            self.ClearAllWindows()
            # We only want to load the File Manager in the Single User version.  It's not the appropriate action
            # for the multi-user version!
            if TransanaConstants.singleUserVersion:
                # Create a File Management Window
                fileManager = FileManagement.FileManagement(self.MenuWindow, -1, _("Transana File Management"))
                # Set up, display, and process the File Management Window
                fileManager.Setup(showModal=True)
                # Destroy the File Manager window
                fileManager.Destroy()

    def LoadClipByNumber(self, clipNum):
        """ When a Clip is identified to trigger systematic loading of all related information,
            this method should be called so that all Transana Objects are set appropriately. """
        # Before we do anything else, let's save the current transcript if it's been modified.
        if self.TranscriptWindow[self.activeTranscript].TranscriptModified():
            self.SaveTranscript(1, cleardoc=1)
        # Set Active Transcript to 0 to signal close of all existing secondary Transcript Windows
        self.activeTranscript = 0
        # Clear all Windows
        self.ClearAllWindows()
        # Load the Clip based on the ClipNumber
        clipObj = Clip.Clip(clipNum)
        # Set the current object to the loaded Episode
        self.currentObj = clipObj
        # Load the Collection that contains the loaded Clip
        collectionObj = Collection.Collection(clipObj.collection_num)
        # set the video start and end points to the start and stop points defined in the clip
        self.VideoStartPoint = clipObj.clip_start                     # Set the Video Start Point to the Clip beginning
        self.VideoEndPoint = clipObj.clip_stop                        # Set the Video End Point to the Clip end
        
        # Load the video identified in the Clip
        if self.LoadVideo(self.currentObj):
            # If we have only one video file ...
            if len(self.currentObj.additional_media_files) == 0:
                # Identify the loaded media file
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Video Media File: "%s"'), 'utf8')
                else:
                    prompt = _('Video Media File: "%s"')
                # Place the file name in the video window's Title bar
                self.VideoWindow.SetTitle(prompt % clipObj.media_filename)
            # If there are multiple videos ...
            else:
                # Just label the video window generically.  There's not room for file names.
                self.VideoWindow.SetTitle(_("Video"))
            # Delineate the appropriate start and end points for Video Control
            self.SetVideoSelection(self.VideoStartPoint, self.VideoEndPoint)
            # Identify the loaded Object
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                str = unicode(_('Transcript for Collection "%s", Clip "%s"'), 'utf8') % (collectionObj.GetNodeString(), clipObj.id)
            else:
                str = _('Transcript for Collection "%s", Clip "%s"') % (collectionObj.GetNodeString(), clipObj.id)
            # The Mac doesn't clean up around frame titles!
            # (The Mac centers titles, while Windows left-justifies them and should not get the leading spaces!)
            if 'wxMac' in wx.PlatformInfo:
                str = "               " + str + "               "
            self.TranscriptWindow[self.activeTranscript].dlg.SetTitle(str)
            # Open the first Clip Transcript in Transcript Window (activeTranscript is ALWAYS 0 here!)
            self.TranscriptWindow[self.activeTranscript].LoadTranscript(clipObj.transcripts[0])
            # Open the remaining clip transcripts in additional transcript windows.
            for tr in clipObj.transcripts[1:]:
                self.OpenAdditionalTranscript(tr.number, isEpisodeTranscript=False)
                self.TranscriptWindow[len(self.TranscriptWindow) - 1].dlg.SetTitle(str)

            # Remove any tabs in the Data Window beyond the Database Tab.  (This was moved down to late in the
            # process due to problems on the Mac documented in the DataWindow object.)
            self.DataWindow.DeleteTabs()

            # Add the Keyword Tab to the DataWindow
            self.DataWindow.AddKeywordsTab(collectionObj=collectionObj, clipObj=clipObj)

            # Get the current selection(s) from the Database Tree
            selItems = self.DataWindow.DBTab.tree.GetSelections()
            # If there are one or more items selected ...
            if len(selItems) >= 1:
                # ... get the item data from the first selection
                selData = self.DataWindow.DBTab.tree.GetPyData(selItems[0])
            # If NO items are selected ...
            else:
                # ... then there's no item data to get
                selData = None
            # If no items are selected or the item selected is NOT a Search Collection or Search Clip ...
            if (selData == None) or not (selData.nodetype in ['SearchCollectionNode', 'SearchClipNode']):
                # Let's make sure this clip is displayed in the Database Tree
                nodeList = (_('Collections'),) + self.currentObj.GetNodeData()
                # Now point the DBTree (the notebook's parent window's DBTab's tree) to the loaded Clip
                self.DataWindow.DBTab.tree.select_Node(nodeList, 'ClipNode')
            
            # Enable the transcript menu item options
            self.MenuWindow.SetTranscriptOptions(True)

            return True
        else:
            # Remove any tabs in the Data Window beyond the Database Tab
            self.DataWindow.DeleteTabs()

            # We only want to load the File Manager in the Single User version.  It's not the appropriate action
            # for the multi-user version!
            if TransanaConstants.singleUserVersion:
                # Create a File Management Window
                fileManager = FileManagement.FileManagement(self.MenuWindow, -1, _("Transana File Management"))
                # Set up, display, and process the File Management Window
                fileManager.Setup(showModal=True)
                # Destroy the File Manager window
                fileManager.Destroy()

            return False

    def OpenAdditionalTranscript(self, transcriptNum, seriesID='', episodeID='', isEpisodeTranscript=True):
        """ Open an additional Transcript without replacing the current one """
        # Create a new Transcript Window
        newTranscriptWindow = TranscriptionUI.TranscriptionUI(TransanaGlobal.menuWindow, includeClose=True)
        # Register this new Transcript Window with the Control Object (self)
        self.Register(Transcript=newTranscriptWindow)
        # Register the Control Object (self) with the new Transcript Window
        newTranscriptWindow.Register(self)
        # Get out Transcript object from the database
        transcriptObj = Transcript.Transcript(transcriptNum)
        # If we have an Episode Transcript, it needs a Window title.  (Clip titles are handled in the calling routine.)
        if isEpisodeTranscript:
            # If we haven't been sent an Episode ID ...
            if episodeID == '':
                # ... get the Episode data based on the Transcript Object ...
                episodeObj = Episode.Episode(transcriptObj.episode_num)
                # ... and note the Episode ID
                episodeID = episodeObj.id
            # If we haven't been sent the Series ID ...
            if seriesID == '':
                # ... get the Series data based on the Episode object ...
                seriesObj = Series.Series(episodeObj.series_num)
                # ... and note the Series ID
                seriesID = seriesObj.id
            # Identify the loaded Object
            prompt = _('Transcript "%s" for Series "%s", Episode "%s"')
            if self.activeTranscript > 0:
                prompt = '** ' + prompt + ' **'
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(prompt, 'utf8')
            # Set the window's prompt
            newTranscriptWindow.dlg.SetTitle(prompt % (transcriptObj.id, seriesID, episodeID))
        # Load the transcript text into the new transcript window
        newTranscriptWindow.LoadTranscript(transcriptObj)
        # Add the new transcript's number to the list that tracks the numbers of the open transcripts.
        self.TranscriptNum[self.activeTranscript] = transcriptObj.number

        # Now we need to arrange the various Transcript windows.
        # if Auto Arrange is enabled ...
        if TransanaGlobal.configData.autoArrange:
            self.AutoArrangeTranscriptWindows()
        # If Auto Arrange is OFF
        else:
            # Determine the position and size of the LAST Transcript Window
            (left, top) = self.TranscriptWindow[self.activeTranscript - 1].dlg.GetPositionTuple()
            (width, height) = self.TranscriptWindow[self.activeTranscript - 1].dlg.GetSizeTuple()
            # Make the new Transcript offset from the last transcript and just a little smaller
            self.TranscriptWindow[self.activeTranscript].dlg.SetDimensions(left + 16, top + 16, width - 16, height - 16)
        # Display the new Transcript window
        newTranscriptWindow.Show()
        newTranscriptWindow.UpdatePosition(self.VideoWindow.GetCurrentVideoPosition())

        # Enable the Multiple Transcript buttons
        for x in range(len(self.TranscriptWindow)):
            self.TranscriptWindow[x].dlg.toolbar.UpdateMultiTranscriptButtons(True)

        # Set focus to the new Transcript's Editor (so that CommonKeys work on the Mac)
        self.TranscriptWindow[self.activeTranscript].dlg.editor.SetFocus()

        if DEBUG:
            print "ControlObjectClass.OpenAdditionalTranscript(%d)  %d" % (transcriptNum, self.activeTranscript)
            for x in range(len(self.TranscriptWindow)):
                print x, self.TranscriptWindow[x].transcriptWindowNumber, self.TranscriptNum[x]
            print
        
    def CloseAdditionalTranscript(self, transcriptNum):
        """ Close a secondary transcript """
        # If we're closeing a transcript other than the active transscript ...
        if self.activeTranscript != transcriptNum:
            # ... remember which transcript WAS active ...
            prevActiveTranscript = self.activeTranscript
            # ... and make the one we're supposed to close active.
            self.activeTranscript = transcriptNum
        # If we're closing the active transcript ...
        else:
            # ... then focus should switch to the top transcript, # 0
            prevActiveTranscript = 0
        # If the prevActiveTranscript is about to be closed, we need to reduce it by one to avoid
        # problems on the Mac.
        if prevActiveTranscript == len(self.TranscriptWindow) - 1:
            prevActiveTranscript = self.activeTranscript - 1
        # Before we do anything else, let's save the current transcript if it's been modified.
        if self.TranscriptWindow[transcriptNum].TranscriptModified():
            self.SaveTranscript(1, cleardoc=1)
        if transcriptNum == 0:
            (left, top) = self.TranscriptWindow[0].dlg.GetPositionTuple()
            self.TranscriptWindow[1].dlg.SetPosition(wx.Point(left, top))
        # ... remove it from the Transcript Window list
        del(self.TranscriptWindow[transcriptNum])
        # ... and remove it from the Transcript Numbers list
        del(self.TranscriptNum[transcriptNum])
        # When all the Transcript Windows are closed, rearrrange the screen
        self.AutoArrangeTranscriptWindows()
        # We need to update the window numbers of the transcript windows.
        for x in range(len(self.TranscriptWindow)):
            # Update the TranscriptUI object
            self.TranscriptWindow[x].transcriptWindowNumber = x
            # Also update the TranscriptUI's Dialog object.  (This is crucial)
            self.TranscriptWindow[x].dlg.transcriptWindowNumber = x
        # Set the frame focus to the Previous active transcript (I'm not convinced this does anything!)
        self.TranscriptWindow[prevActiveTranscript].dlg.SetFocus()
        # Update the Active Transcript number
        self.activeTranscript = prevActiveTranscript
        # If there's only one transcript left ...
        if len(self.TranscriptWindow) == 1:
            # ... Disable the Multiple Transcript buttons
            self.TranscriptWindow[0].dlg.toolbar.UpdateMultiTranscriptButtons(False)

    def SaveAllTranscriptCursors(self):
        """ Save the current cursor position or selection for all open Transcript windows """
        # For each Transcript Window ...
        for trWin in self.TranscriptWindow:
            # ... save the cursorPosition
            trWin.dlg.editor.SaveCursor()

    def RestoreAllTranscriptCursors(self):
        """ Restore the previously saved cursor position or selection for all open Transcript windows """
        # For each Transcript Window ...
        for trWin in self.TranscriptWindow:
            # ... if it HAS a saved cursorPosition ...
            if trWin.dlg.editor.cursorPosition != 0:
                # ... restore the cursor position or selection
                trWin.dlg.editor.RestoreCursor()

    def AutoArrangeTranscriptWindows(self):
        # If we have more than one window ...
        if len(self.TranscriptWindow) > 1:
            # ... define a style that includes the Close Box.  (System_Menu is required for Close to show on Windows in wxPython.)
            style = wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS | wx.SYSTEM_MENU | wx.CLOSE_BOX
        # If there's only one window...
        else:
            # ... then we don't want the close box
            style = wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS
        # Reset the style for the top window
        self.TranscriptWindow[0].dlg.SetWindowStyleFlag(style)
        # Some style changes require a refresh
        self.TranscriptWindow[0].dlg.Refresh()
        # We need to arrange the transcripts if we're leaving Play All Clips mode or if we're in All Windows presentation mode
        if (self.PlayAllClipsWindow == None) or \
           (self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL)):
            # Determine the position and size of the first Transcript
            (left, top) = self.TranscriptWindow[0].dlg.GetPositionTuple()
            (width, height) = self.TranscriptWindow[0].dlg.GetSizeTuple()
            # Get the size of the full screen
            (x, y, w, h) = wx.Display(0).GetClientArea()  # wx.ClientDisplayRect()
            # We don't want the height of the first Transcript window, but the size of the space for all Transcript windows.
            # We assume that it extends from the top of the first Transcript window to the bottom of the whole screen.
            height = h - top
            # We need an adjustment for the Mac.  I don't know why exactly.  It might have to do with the height of the menu bar.
            if 'wxMac' in wx.PlatformInfo:
                height += 20
            # If there's only ONE Media Player AND AutoArrange is turned ON ...
            if (len(self.VideoWindow.mediaPlayers) == 1) and TransanaGlobal.configData.autoArrange:
                # ... the width from the Transcript may very well be incorrect.  Let's grab the width from the Visualization Window
                (width, vh) = self.VisualizationWindow.GetSizeTuple()
            # Initialize a Window Counter
            cnt = 0
            # Iterate through all the Transcript Windows
            for win in self.TranscriptWindow:
                # Increment the counter
                cnt += 1
                # Set the position of each window so they evenly fill up the Transcript space
                win.dlg.SetDimensions(left, top + int((cnt-1) * (height / len(self.TranscriptWindow))), width, int(height / len(self.TranscriptWindow)))

    def ClearAllWindows(self, resetMultipleTranscripts = True):
        """ Clears all windows and resets all objects """
        # Let's stop the media from playing
        self.VideoWindow.Stop()
        # Prompt for save if transcript modifications exist
        self.SaveTranscript(1)
        if resetMultipleTranscripts:
            self.activeTranscript = 0
        # Reset the ControlObject's TranscriptNum
        self.TranscriptNum[self.activeTranscript] = 0
        # Clear Transcript Window
        self.TranscriptWindow[self.activeTranscript].ClearDoc()
        # Identify the loaded Object
        str = _('Transcript')
        self.TranscriptWindow[self.activeTranscript].dlg.SetTitle(str)

        # Clear the Menu Window (Reset menus to initial state)
        self.MenuWindow.ClearMenus()
        # Clear Visualization Window
        self.VisualizationWindow.ClearVisualization()
        # Clear the Video Window
        self.VideoWindow.ClearVideo()
        # Clear the Video Filename as well!
        self.VideoFilename = ''
        # Identify the loaded media file
        str = _('Video')
        self.VideoWindow.SetTitle(str)
        
        # If we are resetting multiple transcripts ...
        if resetMultipleTranscripts:
            # While there are additional Transcript windows open ...
            while len(self.TranscriptWindow) > 1:
                # Save the transcript
                self.SaveTranscript(1, transcriptToSave=len(self.TranscriptWindow) - 1)
                
                # Clear Transcript Window
                self.TranscriptWindow[len(self.TranscriptWindow) - 1].ClearDoc()
                self.TranscriptWindow[len(self.TranscriptWindow) - 1].dlg.Close()
            # When all the Transcritp Windows are closed, rearrrange the screen
            self.AutoArrangeTranscriptWindows()
                    
        # Clear the Data Window
        self.DataWindow.ClearData()
        # Clear the currently loaded object, as there is none
        self.currentObj = None
        
        # Force the screen updates
        # there can be an issue with recursive calls to wxYield, so trap the exception ...
        try:
            wx.Yield()
        # ... and ignore it!
        except:
            pass

    def GetNewDatabase(self):
        """ Close the old database and open a new one. """
        # set the active transcript to 0 so multiple transcript will be cleared
        self.activeTranscript = 0
        # Clear all existing Data
        self.ClearAllWindows()
        # If we're in multi-user ...
        if not TransanaConstants.singleUserVersion:
            # ... stop the Connection Timer so it won't fire while the Database is closed
            TransanaGlobal.connectionTimer.Stop()
        # Close the existing database connection
        DBInterface.close_db()
        # Reset the global encoding to UTF-8 if the Database supports it
        if TransanaGlobal.DBVersion >= u'4.1':
            TransanaGlobal.encoding = 'utf8'
        # Otherwise, if we're in Russian, change the encoding to KOI8r
        elif TransanaGlobal.configData.language == 'ru':
            TransanaGlobal.encoding = 'koi8_r'
        # If we're in Chinese, change the encoding to the appropriate Chinese encoding
        elif TransanaGlobal.configData.language == 'zh':
            TransanaGlobal.encoding = TransanaConstants.chineseEncoding
        # If we're in East Europe Encoding, change the encoding to 'iso8859_2'
        elif TransanaGlobal.configData.language == 'easteurope':
            TransanaGlobal.encoding = 'iso8859_2'
        # If we're in Greek, change the encoding to 'iso8859_7'
        elif TransanaGlobal.configData.language == 'el':
            TransanaGlobal.encoding = 'iso8859_7'
        # If we're in Japanese, change the encoding to cp932
        elif TransanaGlobal.configData.language == 'ja':
            TransanaGlobal.encoding = 'cp932'
        # If we're in Korean, change the encoding to cp949
        elif TransanaGlobal.configData.language == 'ko':
            TransanaGlobal.encoding = 'cp949'
        # Otherwise, fall back to Latin-1
        else:
            TransanaGlobal.encoding = 'latin1'
        # If a new database login fails three times, we need to close the program.
        # Initialize a counter to track that.
        logonCount = 1
        # Flag if Logon succeeds
        loggedOn = False
        # Keep trying for three tries or until successful
        while (logonCount <= 3) and (not loggedOn):
            # Increment logon counter
            logonCount += 1
            # Call up the Username and Password Dialog to get new connection information
            if DBInterface.establish_db_exists():
                # Now update the Data Window
                self.DataWindow.DBTab.tree.refresh_tree()
                # Indicate successful logon
                loggedOn = True
            # If logon fails, inform user and offer to try again twice.
            elif logonCount <= 3:
                # Create a Dialog Box
                dlg = Dialogs.QuestionDialog(self.MenuWindow, _('Transana was unable to connect to the database.\nWould you like to try again?'),
                                         _('Transana Database Connection'))
                # If the user does not want to try again, set the counter to 4, which will cause the program to exit
                if dlg.LocalShowModal() == wx.ID_NO:
                    logonCount = 4
                # Clean up the Dialog Box
                dlg.Destroy()
            # If we're in multi-user and we successfully logged in ...
            if not TransanaConstants.singleUserVersion and loggedOn:
                # ... start the Connection Timer.  This attempts to prevent the "Connection to Database Lost" error by
                # running a very small query every 10 minutes.  See Transana.py.
                TransanaGlobal.connectionTimer.Start(600000)
        # If the Database Connection fails ...
        if not loggedOn:
            # ... Close Transana
            self.MenuWindow.OnFileExit(None)

    def ShowDataTab(self, tabValue):
        """ Changes the visible tab in the notebook in the Data Window """
        if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            # Display the Keywords Tab
            self.DataWindow.nb.SetSelection(tabValue)

    def ProcessCommonKeyCommands(self, event):
        """ Process keyboard commands common to several of Transana's main windows """
        # Assume the key WILL be processed here in this this method
        keyProcessed = True
        # Extract the key code from the event passed in
        c = event.GetKeyCode()
        # Determine if there are modifiers
        hasMods = event.AltDown() or event.ControlDown() or event.CmdDown() or event.ShiftDown()
        # Note whether there is something loaded in the main interface
        loaded = (self.currentObj != None)

        # F1 = Focus on Menu Window
        if (c == wx.WXK_F1) and not hasMods:
            # Set the focus on the Menu Window
            self.MenuWindow.SetFocus()

        # F2 = Focus on Visualization Window
        elif (c == wx.WXK_F2) and not hasMods:
            # Set the focus to the Visualization Window
            self.VisualizationWindow.SetFocus()

        # F3 = Focus on Video Window
        elif (c == wx.WXK_F3) and not hasMods:
            # Set the focus to the Video Window
            self.VideoWindow.SetFocus()

        # F4 = Focus on Transcript Window
        elif (c == wx.WXK_F4) and not hasMods:
            # Determine where the focus currently is (I don't exactly understand why this works.  It's something about
            # this being a "static function")
            tmpFocVal = self.TranscriptWindow[self.activeTranscript].dlg.editor.FindFocus()
            # Now set the focus to the currently active transcript window
            self.TranscriptWindow[self.activeTranscript].dlg.editor.SetFocus()
            # If the focus didn't change ...
            if tmpFocVal == self.TranscriptWindow[self.activeTranscript].dlg.editor.FindFocus():
                # ... then the Transcript Window already HAD focus.  So see if there is more than one Transcript Window ...
                if len(self.TranscriptWindow) > 1:
                    # ... if so, see if the active window is NOT the highest numbered transcript window.
                    if self.activeTranscript < len(self.TranscriptWindow) - 1:
                        # If NOT, increment the Transcript Window by one
                        self.activeTranscript += 1
                    # If we're on the highest-numbered transcript window ...
                    else:
                        # ... then increment back to the start, window zero
                        self.activeTranscript = 0
                    # Now set the focus to the NEW Transcript Window
                    self.TranscriptWindow[self.activeTranscript].dlg.editor.SetFocus()

        # F5 = Focus on Data Window
        elif (c == wx.WXK_F5) and not hasMods:
            # If the Data Window is currently showing the Database tab ...
            if self.DataWindow.nb.GetPageText(self.DataWindow.nb.GetSelection()) == unicode(_("Database"), 'utf8'):
                # ... set the focus to the Database Tree on that tab
                self.DataWindow.DBTab.tree.SetFocus()
            # Otherwise ...
            else:
                # ... just focus on the Data Window Notebook control.
                self.DataWindow.nb.SetFocus()

        # F6 = Toggle Edit / Read Only Mode
        elif (c == wx.WXK_F6) and not hasMods and loaded:
            # Set the focus to the Transcript
            self.TranscriptWindow[self.activeTranscript].dlg.editor.SetFocus()
            # Toggle the Read Only Button on the Transcript Toolbar
            self.TranscriptWindow[self.activeTranscript].dlg.toolbar.ToggleTool(
                self.TranscriptWindow[self.activeTranscript].dlg.toolbar.CMD_READONLY_ID,
                self.TranscriptWindow[self.activeTranscript].dlg.editor.get_read_only())
            # Emulate the Press of the Read Only Button by calling its event directly
            self.TranscriptWindow[self.activeTranscript].dlg.toolbar.OnReadOnlySelect(event)

        # F12 and Ctrl-F12 (for Mac) are Quick Save
        elif (c == wx.WXK_F12) and not (event.AltDown() or event.CmdDown() or event.ShiftDown()) and loaded:
            # if the transcript is in EDIT mode ...
            if not self.TranscriptWindow[self.activeTranscript].dlg.editor.get_read_only():
                # ... save it
                self.TranscriptWindow[self.activeTranscript].dlg.editor.save_transcript()

        # Ctrl-A is Rewind 10 seconds
        elif (c == ord('A')) and event.ControlDown() and loaded:
            # Get the current video position
            vpos = self.GetVideoPosition()
            # Rewind 10 seconds
            self.SetVideoStartPoint(vpos-10000)
            # Explicitly tell Transana to play to the end of the Episode/Clip
            self.SetVideoEndPoint(-1)
            # Play should always be initiated on Ctrl-A, with no auto-rewind
            self.Play(False)

        # Ctrl-D is Stop / Start without Rewind
        elif (c == ord('D')) and event.ControlDown() and loaded:
            if not self.IsPlaying():
                # Explicitly tell Transana to play to the end of the Episode/Clip
                self.SetVideoEndPoint(-1)
            # Play/Pause without rewinding
            self.PlayPause(False)

        # Ctrl-F is Fast Forward 10 seconds
        elif (c == ord('F')) and event.ControlDown() and loaded:
            # Get the current video position
            vpos = self.GetVideoPosition()
            # Fast Forward 10 seconds
            self.SetVideoStartPoint(vpos+10000)
            # Explicitly tell Transana to play to the end of the Episode/Clip
            self.SetVideoEndPoint(-1)
            # Play should always be initiated on Ctrl-F, with no auto-rewind
            self.Play(False)

        # Ctrl-P is Play Previous Time-Coded Segment
        elif (c == ord("P")) and event.ControlDown() and loaded:
            # Get the value for the previous time code
            start_timecode = self.TranscriptWindow[self.activeTranscript].dlg.editor.PrevTimeCode()
            # If there WAS a Previous Segment ....
            if (start_timecode > -1) and self.TranscriptWindow[self.activeTranscript].dlg.editor.get_read_only():
                # Move the Video Start Point to this time position
                self.SetVideoStartPoint(start_timecode)
                # Explicitly tell Transana to play to the end of the Episode/Clip
                self.SetVideoEndPoint(-1)
                # Play should always be initiated on Ctrl-P
                self.Play(0)

        # Ctrl-N is Play Next Time-Coded Segment
        elif (c == ord("N")) and event.ControlDown() and loaded:
            # Get the value for the next time code
            start_timecode = self.TranscriptWindow[self.activeTranscript].dlg.editor.NextTimeCode()
            # If there WAS a Next Segment ...
            if (start_timecode > -1) and self.TranscriptWindow[self.activeTranscript].dlg.editor.get_read_only():
                # Move the Video Start Point to this time position
                self.SetVideoStartPoint(start_timecode)
                # Explicitly tell Transana to play to the end of the Episode/Clip
                self.SetVideoEndPoint(-1)
                # Play should always be initiated on Ctrl-P
                self.Play(0)

        # Ctrl-S is Stop / Start with Rewind
        elif (c == ord('S')) and event.ControlDown() and loaded:
            if not self.IsPlaying():
                # Explicitly tell Transana to play to the end of the Episode/Clip
                self.SetVideoEndPoint(-1)
            # Play/Pause with auto-rewind
            self.PlayPause(True)

        # Ctrl-T inserts a Time Code
        elif (c == ord('T')) and event.ControlDown() and loaded:
            self.TranscriptWindow[self.activeTranscript].dlg.editor.insert_timecode()

        # Ctrl-. is increases playback speed, if possible
        elif (c == ord('.')) and event.ControlDown() and loaded:
            # Ctrl-period increases playback speed by 10% of normal speed
            self.ChangePlaybackSpeed('faster')

        # Ctrl-, is decreases playback speed, if possible
        elif (c == ord(',')) and event.ControlDown() and loaded:
            # Ctrl-comma decreases playback speed by 10% of normal speed
            self.ChangePlaybackSpeed('slower')

        # Ctrl-M is Shapshot (iMage)
        elif (c == ord('M')) and event.ControlDown() and loaded:
            # If possible ...
            if self.VideoWindow.btnSnapshot:
                # Take a Snapshot
                self.VideoWindow.OnSnapshot(event)

        # Otherwise ...
        else:
            # ... the key press had NOT been processed by this method
            keyProcessed = False

        # Let the calling method know if this method processed the key appropriately
        return keyProcessed

    def InsertTimecodeIntoTranscript(self):
        """ Insert a Timecode into the Transcript(s) """
        # For each Transcript window ...
        for trWin in self.TranscriptWindow:
            # ... if the transcript is in Edit mode ...
            if not trWin.dlg.editor.get_read_only():
                # ... get the transcript's selection ...
                selection = trWin.dlg.editor.GetSelection()
                # ... and only if it's a position, not a selection ...
                if selection[0] == selection[1]:
                    # ... then insert the time code.
                    trWin.InsertTimeCode()

    def InsertSelectionTimecodesIntoTranscript(self, startPos, endPos):
        """ Insert a timed pause into the Transcript """
        self.TranscriptWindow[self.activeTranscript].InsertSelectionTimeCode(startPos, endPos)

    def SetTranscriptEditOptions(self, enable):
        """ Change the Transcript's Edit Mode """
        self.MenuWindow.SetTranscriptEditOptions(enable)

    def ActiveTranscriptReadOnly(self):
        return self.TranscriptWindow[self.activeTranscript].dlg.editor.get_read_only()

    def TranscriptUndo(self, event):
        """ Send an Undo command to the Transcript """
        self.TranscriptWindow[self.activeTranscript].TranscriptUndo(event)

    def TranscriptCut(self, event):
        """ Send a Cut command to the Transcript """
        self.TranscriptWindow[self.activeTranscript].TranscriptCut(event)

    def TranscriptCopy(self, event):
        """ Send a Copy command to the Transcript """
        self.TranscriptWindow[self.activeTranscript].TranscriptCopy(event)

    def TranscriptPaste(self, event):
        """ Send a Paste command to the Transcript """
        self.TranscriptWindow[self.activeTranscript].TranscriptPaste(event)

    def TranscriptCallFormatDialog(self, tabToOpen=0):
        """ Tell the TranscriptWindow to open the Format Dialog """
        self.TranscriptWindow[self.activeTranscript].CallFormatDialog(tabToOpen)

    def TranscriptInsertImage(self, fileName = None):
        """ Tell the TranscriptWindow to insert an image """
        self.TranscriptWindow[self.activeTranscript].InsertImage(fileName)

    def Help(self, helpContext):
        """ Handles all calls to the Help System """
        # Getting this to work both from within Python and in the stand-alone executable
        # has been a little tricky.  To get it working right, we need the path to the
        # Transana executables, where Help.exe resides, and the file name, which tells us
        # if we're in Python or not.
        (path, fn) = os.path.split(sys.argv[0])
        
        # If the path is not blank, add the path seperator to the end if needed
        if (path != '') and (path[-1] != os.sep):
            path = path + os.sep

        programName = os.path.join(path, 'Help.py')

        if "__WXMAC__" in wx.PlatformInfo:
            # NOTE:  If we just call Help.Help(), you can't actually do the Tutorial because
            # the Help program's menus override Transana's, and there's no way to get them back.
            # instead of the old call:
            
            # Help.Help(helpContext)
            
            # NOTE:  I've tried a bunch of different things on the Mac without success.  It seems that
            #        the Mac doesn't allow command line parameters, and I have not been able to find
            #        a reasonable method for passing the information to the Help application to tell it
            #        what page to load.  What works is to save the string to the hard drive and 
            #        have the Help file read it that way.  If the user leave Help open, it won't get
            #        updated on subsequent calls, but for now that's okay by me.
            
            helpfile = open(os.getenv("HOME") + '/TransanaHelpContext.txt', 'w')
            pickle.dump(helpContext, helpfile)
            helpfile.flush()
            helpfile.close()

            # On OS X 10.4, when Transana is packed with py2app, the Help call stopped working.
            # It seems we have to remove certain environment variables to get it to work properly!
            # Let's investigate environment variables here!
            envirVars = os.environ
            if 'PYTHONHOME' in envirVars.keys():
                del(os.environ['PYTHONHOME'])
            if 'PYTHONPATH' in envirVars.keys():
                del(os.environ['PYTHONPATH'])
            if 'PYTHONEXECUTABLE' in envirVars.keys():
                del(os.environ['PYTHONEXECUTABLE'])

            os.system('open -a TransanaHelp.app')

        else:
            # NOTE:  If we just call Help.Help(), you can't actually do the Tutorial because 
            # modal dialogs prevent you from focussing back on the Help Window to scroll or
            # advance the Tutorial!  Instead of the old call:
        
            # Help.Help(helpContext)

            # we'll use Python's os.spawn() to create a seperate process for the Help system
            # to run in.  That way, we can go back and forth between Transana and Help as
            # independent programs.
        
            # Make the Help call differently from Python and the stand-alone executable.
            if fn.lower() == 'transana.py':
                # for within Python, we call python, then the Help code and the context
                os.spawnv(os.P_NOWAIT, 'python.bat', [programName, helpContext])
            else:
                # The Standalone requires a "dummy" parameter here (Help), as sys.argv differs between the two versions.
                os.spawnv(os.P_NOWAIT, path + 'Help', ['Help', helpContext])


    # Private Methods
        
    def LoadVideo(self, currentObj):  # (self, Filename, mediaStart, mediaLength):
        """ This method handles loading a video in the video window and loading the
            corresponding Visualization in the Visualization window. """
        # Get the primary file name
        Filename = currentObj.media_filename
        # Get the additional files, if any.
        additionalFiles = currentObj.additional_media_files
        # If we have a Episode ...
        if isinstance(currentObj, Episode.Episode):
            # Initialize the offset value for an Episode to 0, since the 
            offset = 0
            # ... the mediaStart is 0 and the mediaLength is the media file length
            mediaStart = 0
            mediaLength = currentObj.tape_length
            # Signal that we have an Episode
            imgType = 'Episode'
        # If we have a Clip ...
        elif isinstance(currentObj, Clip.Clip):
            # Initialize the offset value for a Clip to the Clip's offset value
            offset = currentObj.offset
            # ... the mediaStart is the clip start and the mediaLength is the clip stop - clip start
            mediaStart = currentObj.clip_start
            mediaLength = currentObj.clip_stop - currentObj.clip_start
            # Signal that we have a Clip
            imgType = 'Clip'
        # See if the primary media file exists
        success = os.path.exists(Filename)
        # If the primary file exists ...
        if success:
            # ... we can iterate through the rest of the media files ...
            for vid in currentObj.additional_media_files:
                # ... to see if they exist as well.
                success = os.path.exists(vid['filename'])
                # As soon as one doesn't exist, we can quit looking.
                if not success:
                    Filename = vid['filename']
                    break
        # If one or more of the Media Files doesn't exist, display an error message.
        if not success:
            # We need a different message for single-user and multi-user Transana if the video file cannot be found.
            if TransanaConstants.singleUserVersion:
                # If it does not exist, display an error message Dialog
                prompt = _('Media File "%s" cannot be found.\nPlease locate this media file and press the "Update Database" button.\nThen reload the Transcript or Clip that failed.')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
            else:
                # If it does not exist, display an error message Dialog
                prompt = _('Media File "%s" cannot be found.\nPlease make sure your video root directory is set correctly, and that the video file exists in the correct location.\nThen reload the Transcript or Clip that failed.')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
            dlg = Dialogs.ErrorDialog(self.MenuWindow, prompt % Filename)
            dlg.ShowModal()
            dlg.Destroy()
        else:
            # If the Visualization Window is visible, open the Visualization in the Visualization Window.
            # Loading Visualization first prevents problems with video being locked by Media Player
            # and thus unavailable for wceraudio DLL/Shared Library for audio extraction (in theory).

            # Load the waveform for the appropriate media files with its current start and length.
            self.VisualizationWindow.load_image(imgType, Filename, additionalFiles, offset, mediaStart, mediaLength)

            # Now that the Visualization is done, load the video in the Video Window
            self.VideoFilename = Filename                # Remember the Video File Name

            # Open the video(s) in the Video Window if the file is found
            self.VideoWindow.open_media_file()
        # Let the calling routine know if we were successful
        return success

    def ClearVisualizationSelection(self):
        """ Clear the current selection from the Visualization Window """
        self.VisualizationWindow.ClearVisualizationSelection()

    def ChangeVisualization(self):
        """ Triggers a complete refresh of the Visualization Window.  Needed for changing Visualization Style. """
        # Capture the Transcript Window's cursor position
        self.SaveAllTranscriptCursors()
        # Update the Visualization Window
        self.VisualizationWindow.Refresh()
        # Restore the Transcript Window's cursor
        self.RestoreAllTranscriptCursors()

    def UpdateKeywordVisualization(self):
        """ If the Keyword Visualization is displayed, update it based on something that could change the keywords
            in the display area. """
        self.VisualizationWindow.UpdateKeywordVisualization()

    def Play(self, setback=False):
        """ This method starts video playback from the current video position. """
        # If we do not already have a cursor position saved, save it
        if self.TranscriptWindow[self.activeTranscript].dlg.editor.cursorPosition == 0:
            self.TranscriptWindow[self.activeTranscript].dlg.editor.SaveCursor()
        # If Setback is requested (Transcription Ctrl-S)
        if setback:
            # Get the current Video position
            videoPos = self.VideoWindow.GetCurrentVideoPosition()
            if type(self.currentObj).__name__ == 'Episode':
                videoStart = 0
            elif type(self.currentObj).__name__ == 'Clip':
                videoStart = self.currentObj.clip_start
            else:
                # Get the current Video marker
                videoStart = self.VideoWindow.GetVideoStartPoint()
            # Assertation: videoPos >= videoStart
            # Find the configured Setback Size (convert to milliseconds)
            setbackSize = TransanaGlobal.configData.transcriptionSetback * 1000
            # If you are further into the video than the Seback Size ...
            if videoPos - videoStart > setbackSize:
                # ... jump back in the video by the setback size
                self.VideoWindow.SetCurrentVideoPosition(videoPos - setbackSize)
            # If the setback would take you to before the beginning of video marker ...
            else:
                # ... jump to the beginning of the video marker
                self.VideoWindow.SetCurrentVideoPosition(videoStart)
        # We need to explicitly set the Clip Endpoint, if it's not known.  (It might be unknown in the ControlObject OR the VideoWindow!)
        # If nothing is loaded, currentObj will be None.  Check to avoid an error.
        if ((self.VideoEndPoint <= 0) or (self.VideoWindow.GetVideoEndPoint() <= 0)) and (self.currentObj != None):
            if type(self.currentObj).__name__ == 'Episode':
                videoEnd = self.VideoWindow.GetMediaLength()
            elif type(self.currentObj).__name__ == 'Clip':
                videoEnd = self.currentObj.clip_stop
            self.SetVideoEndPoint(videoEnd)
        # Play the Video
        self.VideoWindow.Play()

    def Stop(self):
        """ This method stops video playback.  Stop causes the video to be repositioned at the VideoStartPoint. """
        self.VideoWindow.Stop()
        self.RestoreAllTranscriptCursors()

    def Pause(self):
        """ This method pauses video playback.  Pause does not alter the video position, so play will continue from where pause was called. """
        self.VideoWindow.Pause()

    def PlayPause(self, setback=False):
        """ If the video is playing, this pauses it.  If the video is paused, this will make it play. """
        if self.VideoWindow.IsPlaying():
            self.Pause()
        elif self.VideoWindow.IsPaused() or self.VideoWindow.IsStopped():
            self.Play(setback)
        else: # If not playing, paused or stopped, then video not loaded yet
            pass

    def PlayStop(self, setback=False):
        """ If the video is playing, this pauses it.  If the video is paused, this will make it play. """
        if self.VideoWindow.IsPlaying():
            self.Stop()
        elif self.VideoWindow.IsPaused() or self.VideoWindow.IsStopped():
            self.Play(setback)
        else: # If not playing, paused or stopped, then video not loaded yet
            pass

    def PlayLoop(self, startPlay):
        """ Start or stop Looped Playback """
        # Remember if we're supposed to play in a loop
        self.playInLoop = startPlay
        # If we're supposed to play in a loop ...
        if self.playInLoop:
            # Looping doesn't work with Presentation Modes.  You can't stop it without the Visualization Window!
            # Therefore, we need to remember which Presentation Mode we started in.
            if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
                self.LoopPresMode = 'All'
            elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_VIDEO):
                self.LoopPresMode = 'Video'
            elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_TRANS):
                self.LoopPresMode = 'Transcript'
            elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_AUDIO):
                self.LoopPresMode = 'Audio'
            # Now force All Windows Presentation Mode
            self.MenuWindow.menuBar.optionsmenu.Check(MenuSetup.MENU_OPTIONS_PRESENT_ALL, True)
            # ... start playing.
            self.Play()
        # If we're supposed to STOP playing in a loop ...
        else:
            # ... then stop playing!
            self.Stop()
            # Restore the old Presentation Mode
            if self.LoopPresMode == 'Video':
                self.MenuWindow.menuBar.optionsmenu.Check(MenuSetup.MENU_OPTIONS_PRESENT_VIDEO, True)
            elif self.LoopPresMode == 'Transcript':
                self.MenuWindow.menuBar.optionsmenu.Check(MenuSetup.MENU_OPTIONS_PRESENT_TRANS, True)
            elif self.LoopPresMode == 'Audio':
                self.MenuWindow.menuBar.optionsmenu.Check(MenuSetup.MENU_OPTIONS_PRESENT_AUDIO, True)
            # and clear out the mode that was remembered
            self.LoopPresMode = None

    def IsPlaying(self):
        """ Indicates whether the video is playing or not. """
        return self.VideoWindow.IsPlaying()

    def IsPaused(self):
        """ Indicates whether the video is paused or not. """
        return self.VideoWindow.IsPaused()

    def IsLoading(self):
        """ Indicates whether the video is loading into the Player or not. """
        return self.VideoWindow.IsLoading()

    def GetVideoStartPoint(self):
        """ Return the current Video Starting Point """
        return self.VideoStartPoint
    
    def SetVideoStartPoint(self, TimeCode):
        """ Set the Starting Point for video segment definition.  0 is the start of the video.  TimeCode is the nunber of milliseconds from the beginning. """
        # If we are passed a negative time code ...
        if TimeCode < 0:
            # ... set the time code to 0, the start of the video
            TimeCode = 0
        self.VideoWindow.SetVideoStartPoint(TimeCode)
        self.VideoStartPoint = TimeCode

    def GetVideoEndPoint(self):
        """ Return the current Video Ending Point """
        if self.VideoEndPoint > 0:
            return self.VideoEndPoint
        else:
            return self.VideoWindow.GetMediaLength()

    def SetVideoEndPoint(self, TimeCode):
        """ Set the Stopping Point for video segment definition.  0 is the end of the video.  TimeCode is the nunber of milliseconds from the beginning. """
        self.VideoWindow.SetVideoEndPoint(TimeCode)
        self.VideoEndPoint = TimeCode

    def GetVideoSelection(self):
        """ Return the current video starting and ending points """
        return (self.VideoStartPoint, self.VideoEndPoint)

    def SetVideoSelection(self, StartTimeCode, EndTimeCode, UpdateSelectionText=True):
        """ Set the Starting and Stopping Points for video segment definition.  TimeCodes are in milliseconds from the beginning. """
        # For each Transcript Window ...
        for trWin in self.TranscriptWindow:
            # ... if the Window is Read Only (not in Edit mode) ...
            if trWin.dlg.editor.get_read_only():
                # Sometime the cursor is positioned at the end of the selection rather than the beginning, which can cause
                # problems with the highlight.  Let's fix that if needed.
                if trWin.dlg.editor.GetCurrentPos() != trWin.dlg.editor.GetSelection()[0]:
                    (start, end) = trWin.dlg.editor.GetSelection()
                    trWin.dlg.editor.SetCurrentPos(start)
                    trWin.dlg.editor.SetAnchor(end)
                # If Word Tracking is ON ...
                if TransanaGlobal.configData.wordTracking:
                    # ... highlight the full text of the video selection
                    trWin.dlg.editor.scroll_to_time(StartTimeCode)
                    if EndTimeCode > 0:
                        trWin.dlg.editor.select_find(str(EndTimeCode))

        if EndTimeCode <= 0:
            if type(self.currentObj).__name__ == 'Episode':
                EndTimeCode = self.VideoWindow.GetMediaLength()
            elif type(self.currentObj).__name__ == 'Clip':
                EndTimeCode = self.currentObj.clip_stop
        self.SetVideoStartPoint(StartTimeCode)
        self.SetVideoEndPoint(EndTimeCode)
        # The SelectedEpisodeClips window was not updating on the Mac.  Therefore, this was added,
        # even if it might be redundant on Windows.
        if (not self.IsPlaying()) or (self.TranscriptWindow[self.activeTranscript].UpdatePosition(StartTimeCode)):
            if self.DataWindow.SelectedEpisodeClipsTab != None:
                self.DataWindow.SelectedEpisodeClipsTab.Refresh(StartTimeCode)
        # If we should update Selection Text (true at all times other than when clearing the Visualization Window ...)
        if UpdateSelectionText:
            # Update the Selection Text in the current Transcript Window.  But it needs just a tick before the cursor position is set correctly.
            wx.CallLater(50, self.UpdateSelectionTextLater, self.activeTranscript)
        
    def UpdatePlayState(self, playState):
        """ When the Video Player's Play State Changes, we may need to adjust the Screen Layout
            depending on the Presentation Mode settings. """
        # If media playback is stopping and we're supposed to be playing in a loop ...
        if (playState == TransanaConstants.MEDIA_PLAYSTATE_STOP) and self.playInLoop:
            # ... then re-start media playback
            self.Play()
        # If the video is STOPPED, return all windows to normal Transana layout
        elif (playState == TransanaConstants.MEDIA_PLAYSTATE_STOP) and (self.PlayAllClipsWindow == None):
            # When Play is intiated (below), the positions of windows gets saved if they are altered by Presentation Mode.
            # If this has happened, we need to put the screen back to how it was before when Play is stopped.
            if len(self.WindowPositions) != 0:
                # Reset the AutoArrange (which was temporarily disabled for Presentation Mode) variable based on the Menu Setting
                TransanaGlobal.configData.autoArrange = self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_AUTOARRANGE)
                # Reposition the Video Window to its original Position (self.WindowsPositions[2])
                self.VideoWindow.SetDims(self.WindowPositions[2][0], self.WindowPositions[2][1], self.WindowPositions[2][2], self.WindowPositions[2][3])
                # Unpack the Transcript Window Positions
                for winNum in range(len(self.WindowPositions[3])):
                    # if we're NOT using the Rich Text Ctrl (ie. we are using the Styled Text Ctrl) ...
                    if not TransanaConstants.USESRTC:
                        # The Mac has a different base zoom factor than Windows
                        if 'wxMac' in wx.PlatformInfo:
                            zoomFactor = 3
                        else:
                            zoomFactor = 0
                        # Zoom the Transcript window back to normal size
                        self.TranscriptWindow[winNum].dlg.editor.SetZoom(zoomFactor)
                    # Reposition each Transcript Window to its original Position (self.WindowsPositions[3])
                    self.TranscriptWindow[winNum].SetDims(self.WindowPositions[3][winNum][0], self.WindowPositions[3][winNum][1], self.WindowPositions[3][winNum][2], self.WindowPositions[3][winNum][3])
                # Show the Menu Bar
                self.MenuWindow.Show(True)
                # Show the Visualization Window
                self.VisualizationWindow.Show(True)
                # Show the Video Window
                self.VideoWindow.Show(True)
                # Show all Transcript Windows
                for trWindow in self.TranscriptWindow:
                    trWindow.Show(True)
                # Show the Data Window
                self.DataWindow.Show(True)
                # Clear the saved Window Positions, so that if they are moved, the new settings will be saved when the time comes
                self.WindowPositions = []
            # Reset the Transcript Cursors
            self.RestoreAllTranscriptCursors()
                
        # If the video is PLAYED, adjust windows to the desired screen layout,
        # as indicated by the Presentation Mode selection
        elif playState == TransanaConstants.MEDIA_PLAYSTATE_PLAY:
            # If we are starting up from the Video Window, save the Transcript Cursor.
            # Detecting that the Video Window has focus is hard, as there are different video window implementations on
            # different platforms.  Therefore, let's see if it's NOT the Transcript or the Waveform, which are easier to
            # detect.
            if (type(self.MenuWindow.FindFocus()) != type(self.TranscriptWindow[self.activeTranscript].dlg.editor)) and \
               ((self.MenuWindow.FindFocus()) != (self.VisualizationWindow.waveform)):
                self.TranscriptWindow[self.activeTranscript].dlg.editor.SaveCursor()
            # See if Presentation Mode is NOT set to "All Windows" and do all changes common to the other Presentation Modes
            if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL) == False:
                # See if we have already noted the Window Positions.
                if len(self.WindowPositions) == 0:
                    # If not...
                    # Temporarily disable AutoArrange, as it interferes with Presentation Mode
                    TransanaGlobal.configData.autoArrange = False
                    # Get the Window Positions for all Transcript windows
                    transcriptWindowPositions = []
                    for trWin in self.TranscriptWindow:
                        transcriptWindowPositions.append(trWin.GetDimensions())
                    # Save the Window Positions prior to Presentation Mode rearrangement
                    self.WindowPositions = [self.MenuWindow.GetRect(),
                                            self.VisualizationWindow.GetDimensions(),
                                            self.VideoWindow.GetDimensions(),
                                            transcriptWindowPositions,
                                            self.DataWindow.GetDimensions()]
                # Hide the Menu Window
                self.MenuWindow.Show(False)
                # Hide the Visualization Window
                self.VisualizationWindow.Show(False)
                # Hide the Data Window
                self.DataWindow.Show(False)
                # Determine the size of the screen
                (left, top, width, height) = wx.Display(0).GetClientArea()  # wx.ClientDisplayRect()

                # See if Presentation Mode is set to "Video Only"
                if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_VIDEO):
                    # Hide the Transcript Windows
                    for trWindow in self.TranscriptWindow:
                        trWindow.Show(False)
                    # If there is a PlayAllClipsWindow, reset it's size and layout
                    if self.PlayAllClipsWindow != None:
                        # Set the Video Window to take up almost the whole Client Display area
                        self.VideoWindow.SetDims(left + 1, top + 1, width - 2, height - 61)
                        # Set the Window Position in the PlayAllClips Dialog
                        self.PlayAllClipsWindow.xPos = left + 1
                        self.PlayAllClipsWindow.yPos = height - 58
                        # We need a bit more adjustment on the Mac
                        if 'wxMac' in wx.PlatformInfo:
                            self.PlayAllClipsWindow.yPos += 24
                        self.PlayAllClipsWindow.SetRect(wx.Rect(self.PlayAllClipsWindow.xPos, self.PlayAllClipsWindow.yPos, width - 2, 56))
                        # Make the PlayAllClipsWindow the focus
                        self.PlayAllClipsWindow.SetFocus()
                    # If there's NO play all clips window within Video Only presentation mode ...
                    else:
                        # Set the Video Window to take up almost the whole Client Display area
                        self.VideoWindow.SetDims(left + 1, top + 1, width - 2, height - 2)
                        # ... let's create a controller for the video by creating a PlayAllClips window with the current object only
#                        PlayAllClips.PlayAllClips(controlObject=self, singleObject=self.currentObj)
                        # ... When we're done play all the clips, we need to call UpdatePlayState recursively to reset the display!
#                        self.UpdatePlayState(TransanaConstants.MEDIA_PLAYSTATE_STOP)

                # See if Presentation Mode is set to "Video and Transcript"
                elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_TRANS):
                    # Set the screen proportions, currently 70% video, 30% transcript
                    dividePt = 70 / 100.0
                    # We need to make a slight adjustment for the Mac for the menu height
                    if 'wxMac' in wx.PlatformInfo:
                        height += TransanaGlobal.menuHeight
                    # Set the Video Window to take up the top portion of the Client Display Area
                    self.VideoWindow.SetDims(left + 1, top + 2, width - 2, int(dividePt * height) - 3)
                    # If there is a PlayAllClipsWindow, reset it's size and layout
                    if self.PlayAllClipsWindow != None:
                        # Set the Window Position in the PlayAllClips Dialog
                        self.PlayAllClipsWindow.xPos = left + 1
                        self.PlayAllClipsWindow.yPos = int(dividePt * height) + 1
                        if 'wxMac' in wx.PlatformInfo:
                            self.PlayAllClipsWindow.yPos -= 20
                        self.PlayAllClipsWindow.SetRect(wx.Rect(self.PlayAllClipsWindow.xPos, self.PlayAllClipsWindow.yPos, width - 2, 56))
                        # Make the PlayAllClipsWindow the focus
                        self.PlayAllClipsWindow.SetFocus()
                    # Set the Transcript Window to take up the bottom portion of the Client Display Area
                    self.TranscriptWindow[0].SetDims(left + 1, int(dividePt * height) + 1, width - 2, int((1.0 - dividePt) * height) - 2)
                    # Hide the other Transcript Windows
                    for trWindow in self.TranscriptWindow[1:]:
                        trWindow.Show(False)
                    # if we're NOT using the Rich Text Ctrl (ie. we are using the Styled Text Ctrl) ...
                    if not TransanaConstants.USESRTC:
                        # Set the Transcript Zoom Factor
                        if 'wxMac' in wx.PlatformInfo:
                            zoomFactor = 20
                        else:
                            zoomFactor = 14
                        # Zoom in the Transcript window to make the text larger
                        self.TranscriptWindow[0].dlg.editor.SetZoom(zoomFactor)

                # See if Presentation Mode is set to "Audio and Transcript"
                elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_AUDIO):
                    # Hide the Video Window to get Audio with no Video
                    self.VideoWindow.Show(False)
                    # We need to make a slight adjustment for the Mac for the menu height
                    if 'wxMac' in wx.PlatformInfo:
                        height += TransanaGlobal.menuHeight
                    # Calculate the height each Transcript should be
                    winHeight = int((float(height) - top - 2.0) / float(len(self.TranscriptWindow)))
                    # For each Transcript Window:
                    for trWinNum in range(len(self.TranscriptWindow)):
                        # if we're NOT using the Rich Text Ctrl (ie. we are using the Styled Text Ctrl) ...
                        if not TransanaConstants.USESRTC:
                            # Set the Transcript Zoom Factor
                            if 'wxMac' in wx.PlatformInfo:
                                zoomFactor = 20
                            else:
                                zoomFactor = 14
                            # Zoom in the Transcript window to make the text larger
                            self.TranscriptWindow[trWinNum].dlg.editor.SetZoom(zoomFactor)
                        # Set the Transcript Window to take up the entire Client Display Area
                        self.TranscriptWindow[trWinNum].SetDims(left + 1, trWinNum * winHeight + top, width - 2, winHeight)
                        

                    # If there is a PlayAllClipsWindow, reset it's size and layout
                    if self.PlayAllClipsWindow != None:
                        # Set the Window Position in the PlayAllClips Dialog
                        self.PlayAllClipsWindow.xPos = left + 1
                        self.PlayAllClipsWindow.yPos = top
                        if 'wxMac' in wx.PlatformInfo:
                            winHeight = self.TranscriptWindow[0].dlg.GetSizeTuple()[1] - self.TranscriptWindow[0].dlg.GetClientSizeTuple()[1] + 20
                        else:
                            winHeight = self.TranscriptWindow[0].dlg.GetSizeTuple()[1] - self.TranscriptWindow[0].dlg.GetClientSizeTuple()[1] + 30
                        self.PlayAllClipsWindow.SetRect(wx.Rect(self.PlayAllClipsWindow.xPos, self.PlayAllClipsWindow.yPos, width - 2, winHeight))
                        # Make the PlayAllClipsWindow the focus
                        self.PlayAllClipsWindow.SetFocus()

    def ChangePlaybackSpeed(self, direction):
        """ Change the media playback speed on the fly """
        # Pass the request through to the Video Window
        self.VideoWindow.ChangePlaybackSpeed(direction)

    def GetDatabaseDims(self):
        """ Return the dimensions of the Database control. Note that this only returns the Database Tree Tab location.  """
        # Determine the Screen Position of the top left corner of the Tree Control
        (treeLeft, treeTop) = self.DataWindow.DBTab.tree.ClientToScreenXY(1, 1)
        # Determine the width and height of the tree control
        (width, height) = self.DataWindow.DBTab.tree.GetSizeTuple()
        # Return the Database Tree Tab position and size information
        return (treeLeft, treeTop, width, height)

    def GetTranscriptDims(self):
        """ Return the dimensions of the transcript control.  Note that this only includes the transcript itself
        and not the whole Transcript window (including toolbars, etc). """
        return self.TranscriptWindow[self.activeTranscript].GetTranscriptDims()

    def GetCurrentTranscriptObject(self):
        """ Returns a Transcript Object for the Transcript currently loaded in the Transcript Editor """
        return self.TranscriptWindow[self.activeTranscript].GetCurrentTranscriptObject()

    def GetTranscriptSelectionInfo(self):
        """ Returns information about the current selection in the transcript editor """
        # We need to know the time codes that bound the current selection
        (startTime, endTime) = self.TranscriptWindow[self.activeTranscript].dlg.editor.get_selected_time_range()
        # we need to know the text of the current selection
        # If it's blank, we need to send a blank rather than RTF for nothing
        (startPos, endPos) = self.TranscriptWindow[self.activeTranscript].dlg.editor.GetSelection()
        # If there's no current selection ...
        if startPos == endPos:
            # ... get the text between the nearest time codes.
            (st, end, text) = self.TranscriptWindow[self.activeTranscript].dlg.editor.GetTextBetweenTimeCodes(startTime, endTime)
        else:
            if TransanaConstants.USESRTC:
                text = self.TranscriptWindow[self.activeTranscript].dlg.editor.GetFormattedSelection('XML', selectionOnly=True)
            else:
                text = self.TranscriptWindow[self.activeTranscript].dlg.editor.GetRTFBuffer(select_only=1)
        # We also need to know the number of the original Transcript Record
        if self.TranscriptWindow[self.activeTranscript].dlg.editor.TranscriptObj.clip_num == 0:
            # If we have an Episode Transcript, we need the Transcript Number
            originalTranscriptNum = self.TranscriptWindow[self.activeTranscript].dlg.editor.TranscriptObj.number
        else:
            # If we have a Clip Transcript, we need the original Transcript Number, not the Clip Transcript Number.
            # We can get that from the ControlObject's "currentObj", which in this case will be the Clip!
            originalTranscriptNum = self.currentObj.transcripts[self.activeTranscript].source_transcript
        return (originalTranscriptNum, startTime, endTime, text)

    def GetMultipleTranscriptSelectionInfo(self):
        """ Returns information about the current selection(s) in the transcript editor(s) """
        # Initialize a list for the function results
        results = []
        # Iterate through the transcript windows
        for trWindow in self.TranscriptWindow:
            # We need to know the time codes that bound the current selection in the current transcript window
            (startTime, endTime) = trWindow.dlg.editor.get_selected_time_range()
            # If start is 0 ...
            if startTime == 0:
                # ... and we're in a Clip ...
                if isinstance(self.currentObj, Clip.Clip):
                    # ... then the start should be the Clip Start
                    startTime = self.currentObj.clip_start
            # If there is not following time code ...
            if endTime <= 0:
                # ... and we're in a Clip ...
                if isinstance(self.currentObj, Clip.Clip):
                    # ... use the Clip Stop value ...
                    endTime = self.currentObj.clip_stop
                # ... otherwise ...
                else:
                    # ... use the length of the media file
                    endTime = self.GetMediaLength(entire = True)
            # we need to know the text of the current selection in the current transcript window
            # If it's blank, we need to send a blank rather than RTF for nothing
            (startPos, endPos) = trWindow.dlg.editor.GetSelection()
            if startPos == endPos:
                text = ''
            else:
                #text = trWindow.dlg.editor.GetRTFBuffer(select_only=1)
                if TransanaConstants.USESRTC:
                    text = trWindow.dlg.editor.GetFormattedSelection('XML', selectionOnly=True)
                else:
                    text = trWindow.dlg.editor.GetRTFBuffer(select_only=1)
            # We also need to know the number of the original Transcript Record.  If we have an Episode ....
            if trWindow.dlg.editor.TranscriptObj.clip_num == 0:
                # ... we need the Transcript Number, which we can get from the Transcript Window's editor's Transcript Object
                originalTranscriptNum = trWindow.dlg.editor.TranscriptObj.number
            # If we have a Clip ...
            else:
                # ... we need the original Transcript Number, not the Clip Transcript Number.
                # We can get that from the ControlObject's "currentObj", which in this case will be the Clip!
                # We have to pull the source_transcript value from the correct transcript number!
                originalTranscriptNum = self.currentObj.transcripts[self.TranscriptWindow.index(trWindow)].source_transcript
            # Now we can place this transcript's results into the Results list
            results.append((originalTranscriptNum, startTime, endTime, text))
        return results

    def GetDatabaseTreeTabObjectNodeType(self):
        """ Get the Node Type of the currently selected object in the Database Tree in the Data Window """
        return self.DataWindow.DBTab.tree.GetObjectNodeType()

    def SetDatabaseTreeTabCursor(self, cursor):
        """ Change the shape of the cursor for the database tree in the data window """
        self.DataWindow.DBTab.tree.SetCursor(wx.StockCursor(cursor))

    def GetVideoPosition(self):
        """ Returns the current Time Code from the Video Window """
        return self.VideoWindow.GetCurrentVideoPosition()

    def GetVideoCheckboxDataForClips(self, videoPos):
        """ Return the data about the media players checkboxes needed for Clip Creation """
        return self.VideoWindow.GetVideoCheckboxDataForClips(videoPos)

    def VideoCheckboxChange(self):
        """ Detect and adjust to changes in the status of the media player checkboxes """
        # If the check boxes change, the visualization should reflect that.  To accomplish that, we just need to
        # have the Visualization redraw itself when there is idle time.
        self.VisualizationWindow.redrawWhenIdle = True
        
    def UpdateVideoPosition(self, currentPosition):
        """ This method accepts the currentPosition from the video window and propagates that position to other objects """
        # There's a weird glitch with Play All Clips when switching from one multi-transcript clip to the next.
        # Somehow, activeTranscript is getting set to a TranscriptWindow that hasn't yet been created, and that's
        # causing a problem HERE.  The following lines of code fix it.  I haven't been able to track down the cause.
        if self.activeTranscript >= len(self.TranscriptWindow):
            # We need to reset the activeTranscript to 0.  It gets reset later when this line is triggered.
            self.activeTranscript = 0

        # If we do not already have a cursor position saved, and there is a defined cursor position, save it
        if (self.TranscriptWindow[self.activeTranscript].dlg.editor.cursorPosition == 0) and \
           (self.TranscriptWindow[self.activeTranscript].dlg.editor.GetCurrentPos() != 0) and \
           (self.TranscriptWindow[self.activeTranscript].dlg.editor.GetSelection() != (0, 0)):
            self.TranscriptWindow[self.activeTranscript].dlg.editor.SaveCursor()
            
        if self.VideoEndPoint > 0:
            mediaLength = self.VideoEndPoint - self.VideoStartPoint
        else:
            mediaLength = self.VideoWindow.GetMediaLength()
        self.VisualizationWindow.UpdatePosition(currentPosition)

        # Update Transcript position.  If Transcript position changes,
        # then also update the selected Clips tab in the Data window.
        # NOTE:  self.IsPlaying() check added because the SelectedEpisodeClips Tab wasn't updating properly
        if (not self.IsPlaying()) or (self.TranscriptWindow[self.activeTranscript].UpdatePosition(currentPosition)):
            if self.DataWindow.SelectedEpisodeClipsTab != None:
                self.DataWindow.SelectedEpisodeClipsTab.Refresh(currentPosition)

        # Update all Transcript Windows
        for winNum in range(len(self.TranscriptWindow)):
            self.TranscriptWindow[winNum].UpdatePosition(currentPosition)
            self.UpdateSelectionTextLater(winNum)

    def UpdateSelectionTextLater(self, winNum):
        """ Update the Selection Text after the application has had a chance to update the Selection information """
        # When closing windows, we run into trouble.  Check to be sure the window exists to start!
        if winNum in range(len(self.TranscriptWindow)):
            # Get the current selection
            selection = self.TranscriptWindow[winNum].dlg.editor.GetSelection()
            # If we have a point rather than a selection ...
            if selection[0] == selection[1]:
                # We don't need a label
                lbl = ""
            # If we have a selection rather than a point ...
            else:
                # ... we first need to get the time range of the current selection.
                (start, end) = self.TranscriptWindow[winNum].dlg.editor.get_selected_time_range()
                # If start is 0 ...
                if start == 0:
                    # ... and we're in a Clip ...
                    if isinstance(self.currentObj, Clip.Clip):
                        # ... then the start should be the Clip Start
                        start = self.currentObj.clip_start
                # If there is not following time code ...
                if end <= 0:
                    # ... and we're in a Clip ...
                    if isinstance(self.currentObj, Clip.Clip):
                        # ... use the Clip Stop value ...
                        end = self.currentObj.clip_stop
                    # ... otherwise ...
                    else:
                        # ... use the length of the media file
                        end = self.GetMediaLength(entire = True)
                # Then we build the label.
                lbl = _("Selection:  %s - %s")
                if 'unicode' in wx.PlatformInfo:
                    lbl = unicode(lbl, 'utf8')
                lbl = lbl % (Misc.time_in_ms_to_str(start), Misc.time_in_ms_to_str(end))
            # Now display the label on the Transcript Window.
            self.TranscriptWindow[winNum].UpdateSelectionText(lbl)

    def GetMediaLength(self, entire = False):
        """ This method returns the length of the entire video/media segment """
        try:
            if not(entire): # Return segment length
                # if the end point is not defined (possibly due to media length not yet being available) 
                if self.VideoEndPoint <= 0:
                    # Get the length of the longest adjusted media file
                    videoLength = self.VideoWindow.GetMediaLength()
                    # Subtract the video start point, to get segment length
                    mediaLength = videoLength - self.VideoStartPoint

                # If the video end point is defined ...
                else:
                    # ... as long as the length is positive ...
                    if self.VideoEndPoint - self.VideoStartPoint > 0:
                        # ... get the current segment length
                        mediaLength = self.VideoEndPoint - self.VideoStartPoint
                    # If the length is negative ...
                    else:
                        # ... use the total media length as the end point
                        mediaLength = self.VideoWindow.GetMediaLength() - self.VideoStartPoint

                # Sometimes video files don't know their own length because it hasn't been available before.
                # This seems to be a good place to detect and correct that problem before it starts to cause problems,
                # such as in the Keyword Map.

                # First, let's see if an episode is currently loaded that doesn't have a proper length.
                if (isinstance(self.currentObj, Episode.Episode) and \
                   (self.currentObj.media_filename == self.VideoFilename) and \
                   (self.currentObj.tape_length <= 0) and \
                    (self.VideoWindow.mediaPlayers[0].GetMediaLength() > 0)):
                    # Start exception handling, so record lock errors can be ignored
                    try:
                        # Try to lock the record
                        self.currentObj.lock_record()
                        # Get the media length from the first Video Window, not the VideoWindow object, which reports longest adjusted media file length.
                        self.currentObj.tape_length = self.VideoWindow.mediaPlayers[0].GetMediaLength()
                        # for each additional media window ...
                        for x in range(1, len(self.VideoWindow.mediaPlayers)):
                            # ... get the length for each additional media file.  (We need them all.)
                            self.currentObj.additional_media_files[x - 1]['length'] = self.VideoWindow.mediaPlayers[x].GetMediaLength()
                        # Save the object
                        self.currentObj.db_save()
                        # Unlock the record
                        self.currentObj.unlock_record()
                    # If an exception occurs (most likely a Record Lock exception)
                    except:
                        # it can be ignored.
                        pass

                # Return the calculated value
                return mediaLength
            # If the entire length was requested ...
            else:
                # Return length of longest adjusted media file
                return self.VideoWindow.GetMediaLength()
        except:
            # If an exception is raised, most likely we're shutting down and have lost the VideoWindow.  Just return 0.
            return 0
        
    def UpdateVideoWindowPosition(self, left, top, width, height):
        """ This method receives screen position and size information from the Video Window and adjusts all other windows accordingly """
        if TransanaGlobal.configData.autoArrange:
            # This method behaves differently when there are multiple Video Players open
            if (self.currentObj != None) and (len(self.currentObj.additional_media_files) > 0):
                # Get the current dimensionf of the Video Window
                (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                # Update other windows based on this information
                self.UpdateWindowPositions('Video', wleft, wtop + wheight)
            else:

                # NOTE:  Transana panics and crashes if we try to replace this block with the more rational
                # self.UpdateWindowPositions('Video', ... ) code.  Let's not do that.

                if False:                
                    # Get the current dimensionf of the Video Window
                    (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                    # Update other windows based on this information
                    self.UpdateWindowPositions('Video', wleft, YLower=wtop + wheight)

                else:
                    # NOTE:  We only need to trigger Visualization and Data windows' SetDims method to resize everything!

                    # Visualization Window adjusts WIDTH only to match shift in video window
                    (wleft, wtop, wwidth, wheight) = self.VisualizationWindow.GetDimensions()
                    self.VisualizationWindow.SetDims(wleft, wtop, left - wleft - 4, wheight)
                    # Data Window matches Video Window's width and shifts top and height to accommodate shift in video window
                    (wleft, wtop, wwidth, wheight) = self.DataWindow.GetDimensions()
                    self.DataWindow.SetDims(left, top + height + 4, width, wheight - (top + height + 4 - wtop))

            # Play All Clips Window matches the Data Window's WIDTH
            if self.PlayAllClipsWindow != None:
                (parentLeft, parentTop, parentWidth, parentHeight) = self.DataWindow.GetRect()
                (left, top, width, height) = self.PlayAllClipsWindow.GetRect()
                if (parentWidth != width) or (parentLeft != left):
                    self.PlayAllClipsWindow.SetDimensions(parentLeft, top, parentWidth, height)

    def UpdateWindowPositions(self, sender, X, YUpper=-1, YLower=-1):
        """ This method updates all window sizes/positions based on the intersection point passed in.
            X is the horizontal point at which the visualization and transcript windows end and the
            video and data windows begin.
            YUpper is the vertical point where the visualization window ends and the transcript window begins.
            YLower is the vertical point where the video window ends and the data window begins. """

        # NOTE:  This routine was originally written when only one media file could be displayed.  That's why it's
        #        a little more awkward in the multiple-video case.  The data passed in assumes a single video, and
        #        we have to adjust for that.
        
        # We need to adjust the Window Positions to accomodate multiple transcripts!
        # Basically, if we are not in the "first" transcript, we need to substitute the first transcript's
        # "Top position" value for the one sent by the active window.
        if (sender == 'Transcript'):
            YUpper = self.TranscriptWindow[0].dlg.GetPositionTuple()[1] - 4
        # If Auto-Arrange is enabled, resizing one window may alter the positioning of others.
        if TransanaGlobal.configData.autoArrange:
            # If YUpper is NOT passed in ...
            if YUpper == -1:
                # Set it to the BOTTOM of the Visualization Window
                (wleft, wtop, wwidth, wheight) = self.VisualizationWindow.GetDimensions()
                YUpper = wheight + wtop
            # if YLower is NOT passed in ...
            if YLower == -1:
                # Set it to the BOTTOM of the Video Window
                (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                YLower = wheight + wtop
                
            # This method behaves differently when there are multiple Video Players open
            if (self.currentObj != None) and (len(self.currentObj.additional_media_files) > 0):
                # Get the current dimensions of all windows
                visualDims = self.VisualizationWindow.GetDimensions()
                transcriptDims = self.TranscriptWindow[0].GetDimensions()
                videoDims = self.VideoWindow.GetDimensions()
                dataDims = self.DataWindow.GetDimensions()
                # If the Visualization window has been changed ...
                if sender == 'Visualization':
                    # Visual changes Video X, width, height
                    videoDims = (X + 4, videoDims[1], videoDims[2] + (videoDims[0] - X - 4), YUpper - videoDims[1])
                    # Visual changes Transcript Y, height
                    transcriptDims = (transcriptDims[0], YUpper + 4, transcriptDims[2], transcriptDims[3] + (transcriptDims[1] - YUpper - 4))
                    # Visual changes Data Y, height
                    dataDims = (dataDims[0], YUpper + 4, dataDims[2], dataDims[3] + (dataDims[1] - YUpper - 4))
                # If the Video window has been changed ...
                elif sender == 'Video':
                    # Video changes Visual width and height
                    visualDims = (visualDims[0], visualDims[1], X - visualDims[0] - 4, YUpper - visualDims[1])
                    # Video changes Transcript Y and height
                    transcriptDims = (transcriptDims[0], YUpper + 4, transcriptDims[2], transcriptDims[3] + (transcriptDims[1] - YUpper - 4))
                    # Video changes Data Y, height
                    dataDims = (dataDims[0], YUpper + 4, dataDims[2], dataDims[3] + (dataDims[1] - YUpper - 4))
                # If the Transcript window has been changed ...
                elif sender == 'Transcript':
                    # Transcript changes Visual height
                    visualDims = (visualDims[0], visualDims[1], visualDims[2], YUpper - visualDims[1])
                    # Transcript changes Video height
                    videoDims = (videoDims[0], videoDims[1], videoDims[2], YUpper - videoDims[1])
                    # Transcript changes Data X, Y, width, height
                    dataDims = (X + 4, YUpper + 4, dataDims[2] + (dataDims[0] - X - 4), dataDims[3] + (dataDims[1] - YUpper - 4))
                # If the Data window has been changed ...
                elif sender == 'Data':
                    # Data changes Visual height
                    visualDims = (visualDims[0], visualDims[1], visualDims[2], YLower - visualDims[1])
                    # Data changes Transcript Y, width, height
                    transcriptDims = (transcriptDims[0], YLower + 4, X - transcriptDims[0], transcriptDims[3] + (transcriptDims[1] - YLower - 4))
                    # Data changes Video height
                    videoDims = (videoDims[0], videoDims[1], videoDims[2], YLower - videoDims[1])
                # We need to signal that we are resizing everything to reduce redundant OnSize calls
                TransanaGlobal.resizingAll = True
                # Vertical Adjustments
                # Adjust Visualization Window
                if sender != 'Visualization':
                    self.VisualizationWindow.SetDims(visualDims[0], visualDims[1], visualDims[2], visualDims[3])
                # Adjust Transcript Window
                if sender != 'Transcript':
                    self.TranscriptWindow[0].SetDims(transcriptDims[0], transcriptDims[1], transcriptDims[2], transcriptDims[3])
                # If there are multiple transcripts, we need to adjust them all
                if len(self.TranscriptWindow) > 1:
                    self.AutoArrangeTranscriptWindows()
                # Adjust Video Window
                if sender != 'Video':
                    self.VideoWindow.SetDims(videoDims[0], videoDims[1], videoDims[2], videoDims[3])
                # Adjust Data Window
                if sender != 'Data':
                    self.DataWindow.SetDims(dataDims[0], dataDims[1], dataDims[2], dataDims[3])
                # We're done resizing all windows and need to re-enable OnSize calls that would otherwise be redundant
                TransanaGlobal.resizingAll = False
            # If we have only a single video window ...
            else:
                # Adjust Visualization Window
                if sender != 'Visualization':
                    (wleft, wtop, wwidth, wheight) = self.VisualizationWindow.GetDimensions()
                    self.VisualizationWindow.SetDims(wleft, wtop, X - wleft, YUpper - wtop)
                # Adjust Transcript Window
                if sender != 'Transcript':
                    (wleft, wtop, wwidth, wheight) = self.TranscriptWindow[0].GetDimensions()
                    self.TranscriptWindow[0].SetDims(wleft, YUpper + 4, X - wleft, wheight + (wtop - YUpper - 4))
                # If there are multiple transcripts, we need to adjust them all
                if len(self.TranscriptWindow) > 1:
                    self.AutoArrangeTranscriptWindows()
                # Adjust Video Window
                if sender != 'Video':
                    (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                    self.VideoWindow.SetDims(X + 4, wtop, wwidth + (wleft - X - 4), YLower - wtop)
                # Adjust Data Window
                if sender != 'Data':
                    (wleft, wtop, wwidth, wheight) = self.DataWindow.GetDimensions()
                    self.DataWindow.SetDims(X + 4, YLower + 4, wwidth + (wleft - X - 4), wheight + (wtop - YLower - 4))

    def VideoSizeChange(self):
        """ Signal that the Video Size has been changed via the Options > Video menu """
        # If there is a media files loaded ...
        if self.currentObj != None:
            # ... resize the video window.  This will trigger changes in all the other windows as appropriate.
            self.VideoWindow.OnSizeChange()

    def SaveTranscript(self, prompt=0, cleardoc=0, transcriptToSave=-1):
        """Save the Transcript to the database if modified.  If prompt=1,
        prompt the user to confirm the save.  Return 1 if Transcript was
        saved or unchanged, and 0 if user chose to discard changes.  If
        cleardoc=1, then the transcript will be cleared if the user chooses
        to not save."""
        # NOTE:  When the user presses their response to dlg below, it can shift the focus if there are multiple
        #        transcript windows open!  Therefore, remember which transcript we're working on now.
        if transcriptToSave == -1:
            transcriptToSave = self.activeTranscript
        # Was the document modified?
        if self.TranscriptWindow[transcriptToSave].TranscriptModified():
            result = wx.ID_YES
           
            if prompt:
                if self.TranscriptWindow[transcriptToSave].dlg.editor.TranscriptObj.clip_num > 0:
                    pmpt = _("The Clip Transcript has changed.\nDo you want to save it before continuing?")
                else:
                    pmpt = _('Transcript "%s" has changed.\nDo you want to save it before continuing?')
                    if 'unicode' in wx.PlatformInfo:
                        pmpt = unicode(pmpt, 'utf8')
                    pmpt = pmpt % self.TranscriptWindow[transcriptToSave].dlg.editor.TranscriptObj.id
                dlg = Dialogs.QuestionDialog(None, pmpt, _("Question"))
                result = dlg.LocalShowModal()
                dlg.Destroy()
                self.activeTranscript = transcriptToSave
            
            if result == wx.ID_YES:
                try:
                    self.TranscriptWindow[transcriptToSave].SaveTranscript()
                    return 1
                except TransanaExceptions.SaveError, e:
                    dlg = Dialogs.ErrorDialog(None, e.reason)
                    dlg.ShowModal()
                    dlg.Destroy()
                    return 1
            else:
                # If the user does not want to save the edits, discard them to avoid duplicate SAVE prompts
                self.TranscriptWindow[transcriptToSave].dlg.editor.DiscardEdits()
                if cleardoc:
                    self.TranscriptWindow[transcriptToSave].ClearDoc()
                return 0
        return 1

    def SaveTranscriptAs(self):
        """Export the Transcript to an RTF file."""
        dlg = wx.FileDialog(None, defaultDir=self.defaultExportDir,
                            wildcard=_("Rich Text Format (*.rtf)|*.rtf|XML Format (*.xml)|*.xml"), style=wx.SAVE)
        if dlg.ShowModal() == wx.ID_OK:
            # The Default Export Directory should use the last-used value for the session but reset to the
            # video root between sessions.
            self.defaultExportDir = dlg.GetDirectory()
            fname = dlg.GetPath()
            # Mac doesn't automatically append the file extension.  Do it if necessary.
            if (dlg.GetFilterIndex() == 0) and (not fname.upper().endswith(".RTF")):
                fname += '.rtf'
            elif (dlg.GetFilterIndex() == 1) and (not fname.upper().endswith(".XML")):
                fname += '.xml'
            if os.path.exists(fname):
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A file named "%s" already exists.  Do you want to replace it?'), 'utf8')
                else:
                    prompt = _('A file named "%s" already exists.  Do you want to replace it?')
                dlg2 = Dialogs.QuestionDialog(None, prompt % fname,
                                        _('Transana Confirmation'))
                dlg2.CentreOnScreen()
                if dlg2.LocalShowModal() == wx.ID_YES:
                    self.TranscriptWindow[self.activeTranscript].SaveTranscriptAs(fname)
                dlg2.Destroy()
            else:
                self.TranscriptWindow[self.activeTranscript].SaveTranscriptAs(fname)
        dlg.Destroy()

    def PropagateChanges(self, transcriptWindowNumber):
        """ Propagate changes in an Episode transcript down to derived clips """
        # First, let's save the changes in the Transcript.  We don't want to propagate changes, then end up
        # not saving them in the source!
        if self.SaveTranscript(prompt=1):
            # If we are working with an Episode Transcript ...
            if type(self.currentObj).__name__ == 'Episode':
                # Start up the Propagate Episode Transcript Changes tool
                propagateDlg = PropagateEpisodeChanges.PropagateEpisodeChanges(self)
            # If we are working with a Clip Transcript ...
            elif type(self.currentObj).__name__ == 'Clip':
                # If the user has updated the clip's Keywords, self.currentObj will NOT reflect this.
                # Therefore, we need to load a new copy of the clip to get the latest keywords for propagation.
                tempClip = Clip.Clip(self.currentObj.number)
                if TransanaConstants.USESRTC:
                    # Start up the Propagate Clip Changes tool
                    propagateDlg = PropagateEpisodeChanges.PropagateClipChanges(self.MenuWindow,
                                                                                self.currentObj,
                                                                                transcriptWindowNumber,
                                                                                self.TranscriptWindow[transcriptWindowNumber].dlg.editor.GetFormattedSelection('XML'),
                                                                                newKeywordList=tempClip.keyword_list)
                else:
                    # Start up the Propagate Clip Changes tool
                    propagateDlg = PropagateEpisodeChanges.PropagateClipChanges(self.MenuWindow,
                                                                                self.currentObj,
                                                                                transcriptWindowNumber,
                                                                                self.TranscriptWindow[transcriptWindowNumber].dlg.editor.GetRTFBuffer(),
                                                                                newKeywordList=tempClip.keyword_list)

        # If the user chooses NOT to save the Transcript changes ...
        else:
            # ... let them know that nothing was propagated!
            dlg = Dialogs.InfoDialog(None, _("You must save the transcript if you want to propagate the changes."))
            dlg.ShowModal()
            dlg.Destroy()

    def PropagateEpisodeKeywords(self, episodeNum, newKeywordList):
        """ When Episode Keywords are added, this will allow the user to propagate new keywords to all
            Clips created from that Episode if desired. """
        # Get the Episode Object
        tmpEpisode = Episode.Episode(episodeNum)
        # Initialize a list of keywords that have been added to the Episode
        keywordsToAdd = []
        # Iterate through the NEW Keywords list
        for kw in newKeywordList:
            # Initialize that the new keyword has NOT been found
            found = False
            # Iterate through the OLD Keyword list (The "in" operator doesn't work here.)
            for kw2 in tmpEpisode.keyword_list:
                # See if the new Keyword matches the old Keyword
                if (kw.keywordGroup == kw2.keywordGroup) and (kw.keyword == kw2.keyword):
                    # If so, flag it as found ...
                    found = True
                    # ... and stop iterating
                    break
            # If the NEW Keyword was NOT found ...
            if not found:
                # ... add it to the list of keywords to add to Clips
                keywordsToAdd.append(kw)

        # Get the list of clips created from this Episode
        clipList = DBInterface.list_of_clips_by_episode(episodeNum)
        # If there are Clips that have been created from this episode AND Keywords have been added to the Episode ...
        if (len(clipList) > 0) and (len(keywordsToAdd) > 0):
            prompt = unicode(_("Do you want to add the new keywords to all clips created from Episode %s?"), 'utf8')
            # ... build a dialog to prompt the user about adding them to Clips
            tmpDlg = Dialogs.QuestionDialog(self.MenuWindow, prompt % tmpEpisode.id,
                                            _("Episode Keyword Propagation"), noDefault = True)
            # Prompt the user.  If the user says YES ...
            if tmpDlg.LocalShowModal() == wx.ID_YES:
                # Create a Progress Dialog
                prompt = unicode(_("Adding keywords to Episode %s"), 'utf8')
                progress = wx.ProgressDialog(_('Episode Keyword Propagation'), prompt % tmpEpisode.id, style = wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)
                progress.Centre()
                # Initialize the Clip Counter for the progess dialog
                clipCount = 0.0
                # Iterate through the Clip list 
                for clip in clipList:
                    # increment the clip counter for the progress dialog
                    clipCount += 1.0
                    # Load the Clip.
                    tmpClip = Clip.Clip(clip['ClipNum'])
                    # Start Exception Handling
                    try:
                        # Lock the Clip
                        tmpClip.lock_record()
                        # Add the new Keywords
                        for kw in keywordsToAdd:
                            tmpClip.add_keyword(kw.keywordGroup, kw.keyword)
                        # Save the Clip
                        tmpClip.db_save()
                        # Unlock the Clip
                        tmpClip.unlock_record()

                        if not TransanaConstants.singleUserVersion:
                            # We need to update the Keyword Visualization for the current ClipObject
                            if DEBUG:
                                print 'Message to send = "UKV %s %s %s"' % ('Clip', tmpClip.number, tmpClip.episode_num)
                                
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("UKL %s %s" % ('Clip', tmpClip.number))
                                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Clip', tmpClip.number, tmpClip.episode_num))

                        # Increment the Progress Dialog
                        progress.Update(int((clipCount / len(clipList) * 100)))
                    # Handle Exceptions
                    except TransanaExceptions.RecordLockedError, e:
                        prompt = unicode(_('New keywords were not added to Clip "%s"\nin collection "%s"\nbecause the clip record was locked by %s.'), 'utf8')
                        errDlg = Dialogs.ErrorDialog(self.MenuWindow, prompt % (tmpClip.id, tmpClip.GetNodeString(False), e.user))
                        errDlg.ShowModal()
                        errDlg.Destroy()
                # Destroy the Progress Dialog
                progress.Destroy()

                # Need to Update the Keyword Visualization
                self.UpdateKeywordVisualization()

                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    # We need to update the Episode Keyword Visualization
                    if DEBUG:
                        print 'Message to send = "UKV %s %s %s"' % ('Episode', episodeNum, 0)
                        
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKL %s %s" % ('Episode', episodeNum))
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', episodeNum, 0))

            # Destroy the User Prompt dialog
            tmpDlg.Destroy()

    def MultiSelect(self, transcriptWindowNumber):
        """ Make selections in all other transcripts to match the selection in the identified transcript """
        # Determine the start and end times of the selection in the identified transcript window
        (start, end) = self.TranscriptWindow[transcriptWindowNumber].dlg.editor.get_selected_time_range()
        # If start is 0 ...
        if start == 0:
            # ... and we're in a Clip ...
            if isinstance(self.currentObj, Clip.Clip):
                # ... then the start should be the Clip Start
                start = self.currentObj.clip_start
        # If there is not following time code ...
        if end <= 0:
            # ... and we're in a Clip ...
            if isinstance(self.currentObj, Clip.Clip):
                # ... use the Clip Stop value ...
                end = self.currentObj.clip_stop
            # ... otherwise ...
            else:
                # ... use the length of the media file
                end = self.GetMediaLength(entire = True)
        # Iterate through all Transcript Windows
        for trWin in self.TranscriptWindow:
            # If we have a transcript window other than the identified one ...
            if trWin.transcriptWindowNumber != transcriptWindowNumber:
                # ... highlight the full text of the video selection
                trWin.dlg.editor.scroll_to_time(start)
                trWin.dlg.editor.select_find(str(end))
                # Check for time codes at the selection boundaries
                trWin.dlg.editor.CheckTimeCodesAtSelectionBoundaries()
                
            # Once selections are set (later), update the Selection Text
            wx.CallLater(200, self.UpdateSelectionTextLater, trWin.transcriptWindowNumber)
                
    def MultiPlay(self):
        """ Play the current video based on selections in multiple transcripts """
        # Save the cursors for all transcripts (!)
        self.SaveAllTranscriptCursors()
        # Get the Transcript Selection information from all transcript windows.
        transcriptSelectionInfo = self.GetMultipleTranscriptSelectionInfo()
        # Initialize the clip start time to the end of the media file
        earliestStartTime = self.GetMediaLength(True)
        # Initialise the clip end time to the beginning of the media file
        latestEndTime = 0
        # Iterate through the Transcript Selection Info gathered above
        for (transcriptNum, startTime, endTime, text) in transcriptSelectionInfo:
            # If the transcript HAS a selection ...
            if text != "":
                # Check to see if this transcript starts before our current earliest start time, but only if
                # if actually contains text.
                if (startTime < earliestStartTime) and (text != ''):
                    # If so, this is our new earliest start time.
                    earliestStartTime = startTime
                # Check to see if this transcript ends after our current latest end time, but only if
                # if actually contains text.
                if (endTime > latestEndTime) and (text != ''):
                    # If so, this is our new latest end time.
                    latestEndTime = endTime
        # Set the Video Selection to the boundary times
        self.SetVideoSelection(earliestStartTime, latestEndTime)
        # Play the video selection!
        self.Play()

    def UpdateDataWindow(self):
        """ Update the Data Window, as when the "Update Database Window" command is issued """
        # NOTE:  This is called in MU when one user imports a database while another user is connected.
        # Tell the Data Window's Database Tree Tab's Tree to refresh itself
        self.DataWindow.DBTab.tree.refresh_tree()

    def DataWindowHasSearchNodes(self):
        """ Returns the number of Search Nodes in the DataWindow's Database Tree """
        searchNode = self.DataWindow.DBTab.tree.select_Node((_('Search'),), 'SearchRootNode')
        return self.DataWindow.DBTab.tree.ItemHasChildren(searchNode)

    def RemoveDataWindowKeywordExamples(self, keywordGroup, keyword, clipNum):
        """ Remove Keyword Examples from the Data Window """
        # First, remove the Keyword Example from the Database Tree
        # Load the specified Clip record.  Skip the Transcript to speed the process up.
        tempClip = Clip.Clip(clipNum, skipText=True)
        # Prepare the Node List for removing the Keyword Example Node
        nodeList = (_('Keywords'), keywordGroup, keyword, tempClip.id)
        # Call the DB Tree's delete_Node method.  Include the Clip Record Number so the correct Clip entry will be removed.
        self.DataWindow.DBTab.tree.delete_Node(nodeList, 'KeywordExampleNode', tempClip.number)

    def UpdateDataWindowKeywordsTab(self):
        """ Update the Keywords Tab in the Data Window """
        # If the Keywords Tab is the currently displayed tab ...
        if self.DataWindow.nb.GetPageText(self.DataWindow.nb.GetSelection()) == unicode(_('Keywords'), 'utf8'):
            # ... then refresh the Tab
            self.DataWindow.KeywordsTab.Refresh()

    def CreateQuickClip(self):
        """ Trigger the creation of a Quick Clip from outside of the Database Tree """

        # Get the list of selected Nodes in the Database Tree
        dbTreeSelections = self.DataWindow.DBTab.GetSelectedNodeInfo()

        # The selection list must not be empty , and Keywords MUST be selected, or we don't know what keyword to base the Quick Clip on
        if (len(dbTreeSelections) > 0) and (dbTreeSelections[0][3] == 'KeywordNode'):
            # Get the Transcript Selection information from the ControlObject, since we can't communicate with the
            # TranscriptEditor directly.
            (transcriptNum, startTime, endTime, text) = self.GetTranscriptSelectionInfo()
            # Initialize the Episode Number to 0
            episodeNum = 0
            # If our source is an Episode ...
            if isinstance(self.currentObj, Episode.Episode):
                # ... we can just use the ControlObject's currentObj's object number
                episodeNum = self.currentObj.number
                # If we are at the end of a transcript and there are no later time codes, Stop Time will be -1.
                # This is, of course, incorrect, and we must replace it with the Episode Length.
                if endTime <= 0:
                    endTime = self.VideoWindow.GetMediaLength()
            # If our source is a Clip ...
            elif isinstance(self.currentObj, Clip.Clip):
                # ... we need the ControlObject's currentObj's originating episode number
                episodeNum = self.currentObj.episode_num
                # Sometimes with a clip, we get a startTime of 0 from the TranscriptSelectionInfo() method.
                # This is, of course, incorrect, and we must replace it with the Clip Start Time.
                if startTime == 0:
                    startTime = self.currentObj.clip_start
                # Sometimes with a clip, we get an endTime of 0 from the TranscriptSelectionInfo() method.
                # This is, of course, incorrect, and we must replace it with the Clip Stop Time.
                if endTime <= 0:
                    endTime = self.currentObj.clip_stop
            # We now have enough information to populate a ClipDragDropData object to pass to the Clip Creation method.
            clipData = DragAndDropObjects.ClipDragDropData(transcriptNum, episodeNum, startTime, endTime, text, videoCheckboxData=self.GetVideoCheckboxDataForClips(startTime))

            # Let's assemble the keyword list
            kwList = []
            # For each selected node in the DB Tree ...
            for selection in dbTreeSelections:
                # ... get the Node information
                (nodeName, nodeRecNum, nodeParent, nodeType) = selection
                # ... and add the keyword info to the Keyword List
                kwList.append((nodeParent, nodeName))
            # Pass the accumulated data to the CreateQuickClip method, which is in the DragAndDropObjects module
            # because drag and drop is an alternate way to create a Quick Clip.
            DragAndDropObjects.CreateQuickClip(clipData, kwList[0][0], kwList[0][1], self.DataWindow.DBTab.tree, extraKeywords = kwList[1:])

        # If there is something OTHER than a Keyword selected in the Database Tree ...
        elif (len(dbTreeSelections) > 0):
            # ... create an error message
            msg = _("You must select a Keyword in the Data Tree to create a Quick Clip this way.")
            if 'unicode' in wx.PlatformInfo:
                msg = unicode(msg, 'utf8')
            # Display the error message and then clean up.
            dlg = Dialogs.ErrorDialog(None, msg)
            dlg.ShowModal()
            dlg.Destroy()

    def ChangeLanguages(self):
        """ Update all screen components to reflect change in the selected program language """
        self.ClearAllWindows()

##        # Let's look at the issue of database encoding.  We only need to do something if the encoding is NOT UTF-8
##        # or if we're on Windows single-user version.
##        if (TransanaGlobal.encoding != 'utf8') or \
##           (('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion)):
##            # If it's not UTF-*, then if it is Russian, use KOI8r
##            if TransanaGlobal.configData.language == 'ru':
##                newEncoding = 'koi8_r'
##            # If it's Chinese, use the appropriate Chinese encoding
##            elif TransanaGlobal.configData.language == 'zh':
##                newEncoding = TransanaConstants.chineseEncoding
##            # If it's Eastern European Encoding, use 'iso8859_2'
##            elif TransanaGlobal.configData.language == 'easteurope':
##                newEncoding = 'iso8859_2'
##            # If it's Greek, use 'iso8859_7'
##            elif TransanaGlobal.configData.language == 'el':
##                newEncoding = 'iso8859_7'
##            # If it's Japanese, use cp932
##            elif TransanaGlobal.configData.language == 'ja':
##                newEncoding = 'cp932'
##            # If it's Korean, use cp949
##            elif TransanaGlobal.configData.language == 'ko':
##                newEncoding = 'cp949'
##            # Otherwise, fall back to UTF-8
##            else:
##                newEncoding = 'utf8'
##
##            # If we're changing encodings, we need to do a little work here!
##            if newEncoding != TransanaGlobal.encoding:
##                msg = _('Database encoding is changing.  To avoid potential data corruption, \nTransana must close your database before proceeding.')
##                tmpDlg = Dialogs.InfoDialog(None, msg)
##                tmpDlg.ShowModal()
##                tmpDlg.Destroy()
##
##                # We should get a new database.  This call will actually update our encoding if needed!
##                self.GetNewDatabase()
                
        self.MenuWindow.ChangeLanguages()
        self.VisualizationWindow.ChangeLanguages()
        self.DataWindow.ChangeLanguages()
        self.VideoWindow.ChangeLanguages()
        # Updating the Data Window automatically updates the Headers on the Video and Transcript windows!
        for x in range(len(self.TranscriptWindow)):
            self.TranscriptWindow[x].ChangeLanguages()
        # If we're in multi-user mode ...
        if not TransanaConstants.singleUserVersion:
            # We need to update the ChatWindow too
            self.ChatWindow.ChangeLanguages()

    def AutoTimeCodeEnableTest(self):
        """ Test to see if the Fixed-Increment Time Code menu item should be enabled """
        # See if the transcript has some time codes.  If it does, we cannot enable the menu item.
        return len(self.TranscriptWindow[self.activeTranscript].dlg.editor.timecodes) == 0
        
    def AutoTimeCode(self):
        """ Add fixed-interval time codes to a transcript """
        # Ask the Transcript Editor to handle AutoTimeCoding and let us know if it worked.
        # "result" indicates whether the menu item needs to be disabled!
        result = self.TranscriptWindow[self.activeTranscript].dlg.editor.AutoTimeCode()
        # Return the function result obtained
        return result
        
    def AdjustIndexes(self, adjustmentAmount):
        """ Adjust Transcript Time Codes by the specified amount """
        self.TranscriptWindow[self.activeTranscript].AdjustIndexes(adjustmentAmount)

    def __repr__(self):
        """ Return a string representation of information about the ControlObject """
        tempstr = "Control Object contents:\nVideoFilename = %s\nVideoStartPoint = %s\nVideoEndPoint = %s\n"  % (self.VideoFilename, self.VideoStartPoint, self.VideoEndPoint)
        tempstr += 'Current open transcripts: %d (%d)\n' % (len(self.TranscriptWindow), self.activeTranscript)
        return tempstr.encode('utf8')

    def _get_activeTranscript(self):
        """ "Getter" for the activeTranscript property """
        # We need to return the activeTranscript value
        return self._activeTranscript

    def _set_activeTranscript(self, transcriptNum):
        """ "Setter" for the activeTranscript property """
        # Initiate exception handling.  (Shutting down Transana generates exceptions here!)
        try:
            # Iterate through the defined Transcript Windows
            for x in range(len(self.TranscriptWindow)):
                # Get the current window title
                title = self.TranscriptWindow[x].dlg.GetTitle()
                # If the current window is NOT the new active trancript, yet is labeled as
                # the active transcript (i.e. is LOSING focus) ...
                if (x != transcriptNum) and (title[:2] == '**') and (title[-2:] == '**'):
                    # ... remove the asterisks from the title.  (CallAfter resolves timing problems)
                    # But skip this if we are shutting down, as trying to set the title of a deleted
                    # window causes an exception!
                    if not self.shuttingDown:
                        wx.CallAfter(self.TranscriptWindow[x].dlg.SetTitle, title[3:-3])
                        
                # If the current window IS the new active transcript, but is not yet labeled
                # as the active transcript (i.e. is GAINING focus) ...
                if (x == transcriptNum) and (title[:2] != '**') and (title[-2:] != '**') and \
                   (len(self.TranscriptWindow) > 1):
                    # ... create a prompt that puts asterisks on either side of the window title ...
                    prompt = '** %s **'
                    if 'unicode' in wx.PlatformInfo:
                        prompt = unicode(prompt, 'utf8')
                    # ... and set the window title to this new prompt
                    self.TranscriptWindow[x].dlg.SetTitle(prompt % title)

                # Set the Menus to match the active transcript's Edit state
                if (x == transcriptNum):
                    # Enable or disable the transcript menu item options
                    self.MenuWindow.SetTranscriptEditOptions(not self.TranscriptWindow[x].dlg.editor.get_read_only())
        except:

            if DEBUG:
                print "Exception in ControlObjectClass._set_activeTranscript()"
            
            # We can ignore it.  This only happens when shutting down Transana.
            pass
        # Set the underlying data value to the new window number
        self._activeTranscript = transcriptNum

    # define the activeTranscript property for the ControlObject.
    # Doing this as a property allows automatic labeling of the active transcript window.
    activeTranscript = property(_get_activeTranscript, _set_activeTranscript, None)
