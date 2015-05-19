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
# import Transana File Management System
import FileManagement

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
        # Define Objects that need controlling (initializing to None)
        self.MenuWindow = None
        self.VideoWindow = None
        self.TranscriptWindow = None
        self.VisualizationWindow = None
        self.DataWindow = None
        self.PlayAllClipsWindow = None
        self.ChatWindow = None

        # Initialize variables
        self.VideoFilename = ''         # Video File Name
        self.VideoStartPoint = 0        # Starting Point for video playback in Milliseconds
        self.VideoEndPoint = 0          # Ending Point for video playback in Milliseconds
        self.WindowPositions = []       # Initial Screen Positions for all Windows, used for Presentation Mode
        self.TranscriptNum = -1         # Transcript record # loaded
        self.currentObj = None          # Currently loaded Object (Episode or Clip)
        
    def Register(self, Menu='', Video='', Transcript='', Data='', Visualization='', PlayAllClips='', Chat=''):
        """ The ControlObject can extert control only over those objects it knows about.  This method
            provides a way to let the ControlObject know about other objects.  This infrastructure allows
            for objects to be swapped in and out.  For example, if you need a different video window
            that supports a format not available on the current one, you can hide the current one, show
            a new one, and register that new one with the ControlObject.  Once this is done, the new
            player will handle all tasks for the program.  """
        # This function expects parameters passed by name and "registers" the components that
        # need to be available to the ControlObject to be controlled.  To remove an
        # objecct registration, pass in "None"
        if Menu != '':
            self.MenuWindow = Menu                       # Define the Menu Window Object
        if Video != '':
            self.VideoWindow = Video                     # Define the Video Window Object
        if Transcript != '':
            self.TranscriptWindow = Transcript           # Define the Transcript Window Object
        if Data != '':
            self.DataWindow = Data                       # Define the Data Window Object
        if Visualization != '':
            self.VisualizationWindow = Visualization     # Define the Visualization Window Object
        if PlayAllClips != '':
            self.PlayAllClipsWindow = PlayAllClips       # Define the Play All Clips Window Object
        if Chat != '':
            self.ChatWindow = Chat                       # Define the Chat Window Object

    def CloseAll(self):
        """ This method closes all application windows and cleans up objects when the user
            quits Transana. """
        # Closing the MenuWindow will automatically close the Transcript, Data, and Visualization
        # Windows in the current setup of Transana, as these windows are all defined as child dialogs
        # of the MenuWindow.
        self.MenuWindow.Close()
        # VideoWindow is a wxFrame, rather than a wxDialog like the other windows.  Therefore,
        # it needs to be closed explicitly.
        self.VideoWindow.close()

    def LoadTranscript(self, series, episode, transcript):
        """ When a Transcript is identified to trigger systemic loading of all related information,
            this method should be called so that all Transana Objects are set appropriately. """
        # Before we do anything else, let's save the current transcript if it's been modified.
        if self.TranscriptWindow.TranscriptModified():
            self.SaveTranscript(1, cleardoc=1)
        # Clear all Windows
        self.ClearAllWindows()
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

        if self.LoadVideo(episodeObj.media_filename, 0, episodeObj.tape_length):    # Load the video identified in the Episode

            # Force the Visualization to load here.  This ensures that the Episode visualization is shown
            # rather than the Clip visualization when Locating a Clip
            self.VisualizationWindow.OnIdle(None)
            
            # Identify the loaded Object
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Transcript "%s" for Series "%s", Episode "%s"'), 'utf8')
            else:
                prompt = _('Transcript "%s" for Series "%s", Episode "%s"')
            self.TranscriptWindow.dlg.SetTitle(prompt % (transcriptObj.id, seriesObj.id, episodeObj.id))
            # Identify the loaded media file
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Video Media File: "%s"'), 'utf8')
            else:
                prompt = _('Video Media File: "%s"')
            self.VideoWindow.frame.SetTitle(prompt % episodeObj.media_filename)
            # Open Transcript in Transcript Window
            self.TranscriptWindow.LoadTranscript(transcriptObj) #flies off to transcriptionui.py

            self.TranscriptNum = transcriptObj.number
            
            # Add the Episode Clips Tab to the DataWindow
            self.DataWindow.AddEpisodeClipsTab(seriesObj=seriesObj, episodeObj=episodeObj)

            # Add the Selected Episode Clips Tab, initially set to the beginning of the video file
            # TODO:  When the Transcript Window updates the selected text, we need to update this tab in the Data Window!
            self.DataWindow.AddSelectedEpisodeClipsTab(seriesObj=seriesObj, episodeObj=episodeObj, TimeCode=0)

            # Add the Keyword Tab to the DataWindow
            self.DataWindow.AddKeywordsTab(seriesObj=seriesObj, episodeObj=episodeObj)
            # Enable the transcript menu item options
            self.MenuWindow.SetTranscriptOptions(True)

            # When an Episode is first loaded, we don't know how long it is.  
            # Deal with missing episode length.
            if episodeObj.tape_length == 0:
                # The video has been loaded in the Media Player now, so this should work.
                episodeObj.tape_length = self.GetMediaLength()
                # If we now know the Media Length...
                if episodeObj.tape_length != 0:
                    # Let's try to save the Episode Object, since we've added information
                    try:
                        episodeObj.lock_record()
                        episodeObj.db_save()
                        episodeObj.unlock_record()
                    except:
                        pass
        else:
            # Create a File Management Window
            fileManager = FileManagement.FileManagement(self.MenuWindow, -1, _("Transana File Management"))
            # Set up, display, and process the File Management Window
            fileManager.Setup(showModal=True)

    def LoadClipByNumber(self, clipNum):
        """ When a Clip is identified to trigger systematic loading of all related information,
            this method should be called so that all Transana Objects are set appropriately. """
        # Before we do anything else, let's save the current transcript if it's been modified.
        if self.TranscriptWindow.TranscriptModified():
            self.SaveTranscript(1, cleardoc=1)
        # Load the Clip based on the ClipNumber
        clipObj = Clip.Clip(clipNum)
        # Set the current object to the loaded Episode
        self.currentObj = clipObj
        # Load the Collection that contains the loaded Clip
        collectionObj = Collection.Collection(clipObj.collection_num)
        # Load the Clip Transcript
        transcriptObj = Transcript.Transcript(clip=clipObj.number)    # Load the Clip Transcript
        # set the video start and end points to the start and stop points defined in the clip
        self.VideoStartPoint = clipObj.clip_start                     # Set the Video Start Point to the Clip beginning
        self.VideoEndPoint = clipObj.clip_stop                        # Set the Video End Point to the Clip end
        
        # Load the video identified in the Clip
        if self.LoadVideo(clipObj.media_filename, clipObj.clip_start, clipObj.clip_stop - clipObj.clip_start):
            # Identify the loaded Object
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                str = unicode(_('Transcript for Collection "%s", Clip "%s"'), 'utf8') % (collectionObj.id, clipObj.id)
            else:
                str = _('Transcript for Collection "%s", Clip "%s"') % (collectionObj.id, clipObj.id)
            # The Mac doesn't clean up around frame titles!
            # (The Mac centers titles, while Windows left-justifies them and should not get the leading spaces!)
            if 'wxMac' in wx.PlatformInfo:
                str = "               " + str + "               "
            self.TranscriptWindow.dlg.SetTitle(str)
            # Identify the loaded media file
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                str = unicode(_('Video Media File: "%s"'), 'utf8')
            else:
                str = _('Video Media File: "%s"')
            self.VideoWindow.frame.SetTitle(str % clipObj.media_filename)
            # Delineate the appropriate start and end points for Video Control
            self.SetVideoSelection(self.VideoStartPoint, self.VideoEndPoint)
            # Open the Clip Transcript in Transcript Window
            self.TranscriptWindow.LoadTranscript(transcriptObj)

            # Remove any tabs in the Data Window beyond the Database Tab.  (This was moved down to late in the
            # process due to problems on the Mac documented in the DataWindow object.)
            self.DataWindow.DeleteTabs()

            # Add the Keyword Tab to the DataWindow
            self.DataWindow.AddKeywordsTab(collectionObj=collectionObj, clipObj=clipObj)
            # Enable the transcript menu item options
            self.MenuWindow.SetTranscriptOptions(True)

            return True
        else:
            # Remove any tabs in the Data Window beyond the Database Tab
            self.DataWindow.DeleteTabs()

            # Create a File Management Window
            fileManager = FileManagement.FileManagement(self.MenuWindow, -1, _("Transana File Management"))
            # Set up, display, and process the File Management Window
            fileManager.Setup(showModal=True)

            return False


    def ClearAllWindows(self):
        """ Clears all windows and resets all objects """
        # Prompt for save if transcript modifications exist
        self.SaveTranscript(1)

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
        self.VideoWindow.frame.SetTitle(str)
        # Clear Transcript Window
        self.TranscriptWindow.ClearDoc()
        # Also reset the ControlObject's TranscriptNum
        self.TranscriptNum = 0
        # Identify the loaded Object
        str = _('Transcript')
        self.TranscriptWindow.dlg.SetTitle(str)
        # Clear the Data Window
        self.DataWindow.ClearData()
        # Clear the currently loaded object, as there is none
        self.currentObj = None
        # Force the screen updates
        wx.Yield()

    def GetNewDatabase(self):
        """ Close the old database and open a new one. """
        # Clear all existing Data
        self.ClearAllWindows()
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
                dlg = wx.MessageDialog(self.MenuWindow, _('Transana was unable to connect to the database.\nWould you like to try again?'),
                                         _('Transana Database Connection'), wx.YES_NO | wx.ICON_QUESTION)
                # If the user does not want to try again, set the counter to 4, which will cause the program to exit
                if dlg.ShowModal() == wx.ID_NO:
                    logonCount = 4
                # Clean up the Dialog Box
                dlg.Destroy()
                
        # If the Database Connection fails ...
        if not loggedOn:
            # ... Close Transana
            self.MenuWindow.OnFileExit(None)

    def ShowDataTab(self, tabValue):
        if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            # Display the Keywords Tab
            self.DataWindow.nb.SetSelection(tabValue)

    def InsertTimecodeIntoTranscript(self):
        self.TranscriptWindow.InsertTimeCode()

    def InsertSelectionTimecodesIntoTranscript(self, startPos, endPos):
        self.TranscriptWindow.InsertSelectionTimeCode(startPos, endPos)

    def SetTranscriptEditOptions(self, enable):
        self.MenuWindow.SetTranscriptEditOptions(enable)

    def TranscriptUndo(self, event):
        self.TranscriptWindow.TranscriptUndo(event)

    def TranscriptCut(self):
        self.TranscriptWindow.TranscriptCut()

    def TranscriptCopy(self):
        self.TranscriptWindow.TranscriptCopy()

    def TranscriptPaste(self):
        self.TranscriptWindow.TranscriptPaste()

    def TranscriptCallFontDialog(self):
        self.TranscriptWindow.CallFontDialog()

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
        
        # print "ControlObject.help(): '%s' '%s' '%s'" % (path, fn, programName)

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
            
            file = open('/Applications/Transana 2/TransanaHelpContext.txt', 'w')
            pickle.dump(helpContext, file)
            file.flush()
            file.close()            
            
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
                os.spawnv(os.P_NOWAIT, 'python', [programName, helpContext])
            else:
                # The Standalone requires a "dummy" parameter here (Help), as sys.argv differs between the two versions.
                os.spawnv(os.P_NOWAIT, path + 'Help', ['Help', helpContext])


    # Private Methods
        
    def LoadVideo(self, Filename, mediaStart, mediaLength):
        """ This method handles loading a video in the video window and loading the
            corresponding Visualization in the Visualization window. """
        # Assume this will succeed
        success = True
        # Check for the existence of the Media File
        if not os.path.exists(Filename):
            # If it does not exist, display an error message Dialog
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Media File "%s" cannot be found.\nPlease locate this media file and press the "Update Database" button.\nThen reload the Transcript or Clip that failed.'), 'utf8')
            else:
                prompt = _('Media File "%s" cannot be found.\nPlease locate this media file and press the "Update Database" button.\nThen reload the Transcript or Clip that failed.')
            dlg = Dialogs.ErrorDialog(self.MenuWindow, prompt % Filename)
            dlg.ShowModal()
            dlg.Destroy()
            # Indicate that LoadVideo failed.
            success = False
        else:
            # If the Visualization Window is visible, open the Visualization in the Visualization Window.
            # Loading Visualization first prevents problems with video being locked by Media Player
            # and thus unavailable for wceraudio DLL/Shared Library for audio extraction (in theory).
            self.VisualizationWindow.load_image(Filename, mediaStart, mediaLength)

            # Now that the Visualization is done, load the video in the Video Window
            self.VideoFilename = Filename                # Remember the Video File Name

            # Open the video in the Video Window if the file is found
            self.VideoWindow.open_media_file(Filename)
        # Let the calling routine know if we were successful
        return success

    def ClearVisualizationSelection(self):
        self.VisualizationWindow.ClearVisualizationSelection()

    def Play(self, setback=False):
        """ This method starts video playback from the current video position. """
        # If we do not already have a cursor position saved, save it
        if self.TranscriptWindow.dlg.editor.cursorPosition == 0:
            self.TranscriptWindow.dlg.editor.cursorPosition = (self.TranscriptWindow.dlg.editor.GetCurrentPos(), self.TranscriptWindow.dlg.editor.GetSelection())

        # If Setback is requested (Transcription Ctrl-S)
        if setback:
            # Get the current Video position
            videoPos = self.VideoWindow.GetCurrentVideoPosition()

            if type(self.currentObj).__name__ == 'Episode':
                videoStart = 0
            elif type(self.currentObj).__name__ == 'Clip':
                videoStart = self.currentObj.clipStart
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

        # Because of differences in the Media Player technologies on Windows and the Mac, we need
        # to explicitly set the Clip Endpoint if we're on a Mac
        if ("__WXMAC__" in wx.PlatformInfo) and (self.VideoEndPoint == -1):
            if type(self.currentObj).__name__ == 'Episode':
                videoEnd = self.currentObj.tape_length
            elif type(self.currentObj).__name__ == 'Clip':
                videoEnd = self.currentObj.clip_stop
                
            self.SetVideoEndPoint(videoEnd)

        # Play the Video
        self.VideoWindow.Play()

    def Stop(self):
        """ This method stops video playback.  Stop causes the video to be repositioned at the VideoStartPoint. """
        self.VideoWindow.Stop()
        if self.TranscriptWindow.dlg.editor.cursorPosition != 0:
            self.TranscriptWindow.dlg.editor.RestoreCursor()

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
        if self.VideoEndPoint > 0:
            return self.VideoEndPoint
        else:
            return self.VideoWindow.GetMediaLength()

    def SetVideoEndPoint(self, TimeCode):
        """ Set the Stopping Point for video segment definition.  0 is the end of the video.  TimeCode is the nunber of milliseconds from the beginning. """
        self.VideoWindow.SetVideoEndPoint(TimeCode)
        self.VideoEndPoint = TimeCode

    def GetVideoSelection(self):
        return (self.VideoStartPoint, self.VideoEndPoint)

    def SetVideoSelection(self, StartTimeCode, EndTimeCode):
        """ Set the Starting and Stopping Points for video segment definition.  TimeCodes are in milliseconds from the beginning. """
        if self.TranscriptWindow.dlg.editor.get_read_only():

            if DEBUG:
                print "ControlObjectClass.SetVideoSelection(): editor position before =", self.TranscriptWindow.dlg.editor.GetCurrentPos(), self.TranscriptWindow.dlg.editor.GetSelection()

            # Sometime the cursor is positioned at the end of the selection rather than the beginning, which can cause
            # problems with the highlight.  Let's fix that if needed.
            if self.TranscriptWindow.dlg.editor.GetCurrentPos() != self.TranscriptWindow.dlg.editor.GetSelection()[0]:

                if DEBUG:
                    print "ControlObjectClass.SetVideoSelection() Correction!!"

                (start, end) = self.TranscriptWindow.dlg.editor.GetSelection()
                self.TranscriptWindow.dlg.editor.SetCurrentPos(start)
                self.TranscriptWindow.dlg.editor.SetAnchor(end)
                
            # If Word Tracking is ON ...
            if TransanaGlobal.configData.wordTracking:
                # ... highlight the full text of the video selection
                self.TranscriptWindow.dlg.editor.scroll_to_time(StartTimeCode)

            if DEBUG:
                print "ControlObjectClass.SetVideoSelection(): editor position after scroll_to_time() =", self.TranscriptWindow.dlg.editor.GetCurrentPos(), self.TranscriptWindow.dlg.editor.GetSelection()
                
            if EndTimeCode > 0:
                self.TranscriptWindow.dlg.editor.select_find(str(EndTimeCode))

            if DEBUG:
                print "ControlObjectClass.SetVideoSelection(): editor position after select_find() =", self.TranscriptWindow.dlg.editor.GetCurrentPos(), self.TranscriptWindow.dlg.editor.GetSelection()
                
        if EndTimeCode == 0:
            if type(self.currentObj).__name__ == 'Episode':
                EndTimeCode = self.VideoWindow.GetMediaLength()
            elif type(self.currentObj).__name__ == 'Clip':
                EndTimeCode = self.currentObj.clip_stop
            
        self.SetVideoStartPoint(StartTimeCode)
        self.SetVideoEndPoint(EndTimeCode)
        # The SelectedIEpisodeClips window was not updating on the Mac.  Therefore, this was added,
        # even if it might be redundant on Windows.
        if (not self.IsPlaying()) or (self.TranscriptWindow.UpdatePosition(StartTimeCode)):
            if self.DataWindow.SelectedEpisodeClipsTab != None:
                self.DataWindow.SelectedEpisodeClipsTab.Refresh(StartTimeCode)
        
        
    def UpdatePlayState(self, playState):
        """ When the Video Player's Play State Changes, we may need to adjust the Screen Layout
            depending on the Presentation Mode settings. """

        # print "ControlObject.UpdatePlayState():", playState, TransanaConstants.MEDIA_PLAYSTATE_STOP, TransanaConstants.MEDIA_PLAYSTATE_PLAY, TransanaConstants.MEDIA_PLAYSTATE_PAUSE

        # print "ControlObject.UpdatePlayState(): PlayAllClipsWindow =", self.PlayAllClipsWindow
        
        # If the video is STOPPED, return all windows to normal Transana layout
        if (playState == TransanaConstants.MEDIA_PLAYSTATE_STOP) and (self.PlayAllClipsWindow == None):
            # When Play is intiated (below), the positions of windows gets saved if they are altered by Presentation Mode.
            # If this has happened, we need to put the screen back to how it was before when Play is stopped.
            if len(self.WindowPositions) != 0:
                # Reset the AutoArrange (which was temporarily disabled for Presentation Mode) variable based on the Menu Setting
                TransanaGlobal.configData.autoArrange = self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_AUTOARRANGE)
                # Reposition the Video Window to its original Position (self.WindowsPositions[2])
                self.VideoWindow.SetDims(self.WindowPositions[2][0], self.WindowPositions[2][1], self.WindowPositions[2][2], self.WindowPositions[2][3])
                # Reposition the Transcript Window to its original Position (self.WindowsPositions[3])
                self.TranscriptWindow.SetDims(self.WindowPositions[3][0], self.WindowPositions[3][1], self.WindowPositions[3][2], self.WindowPositions[3][3])
                # Show the Menu Bar
                self.MenuWindow.Show(True)
                # Show the Visualization Window
                self.VisualizationWindow.Show(True)
                # Show the Transcript Window
                self.TranscriptWindow.Show(True)
                # Show the Data Window
                self.DataWindow.Show(True)
                # Clear the saved Window Positions, so that if they are moved, the new settings will be saved when the time comes
                self.WindowPositions = []
            # Reset the Transcript Cursor
            self.TranscriptWindow.dlg.editor.RestoreCursor()
                
        # If the video is PLAYED, adjust windows to the desired screen layout,
        # as indicated by the Presentation Mode selection
        elif playState == TransanaConstants.MEDIA_PLAYSTATE_PLAY:
            # If we are starting up from the Video Window, save the Transcript Cursor.
            # Detecting that the Video Window has focus is hard, as there are different video window implementations on
            # different platforms.  Therefore, let's see if it's NOT the Transcript or the Waveform, which are easier to
            # detect.
            if (type(self.MenuWindow.FindFocus()) != type(self.TranscriptWindow.dlg.editor)) and \
               ((self.MenuWindow.FindFocus()) != (self.VisualizationWindow.waveform)):
                self.TranscriptWindow.dlg.editor.cursorPosition = (self.TranscriptWindow.dlg.editor.GetCurrentPos(), self.TranscriptWindow.dlg.editor.GetSelection())

            # See if Presentation Mode is NOT set to "All Windows" and do all changes common to the other Presentation Modes
            if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL) == False:
                # See if we have already noted the Window Positions.
                if len(self.WindowPositions) == 0:
                    # If not...
                    # Temporarily disable AutoArrange, as it interferes with Presentation Mode
                    TransanaGlobal.configData.autoArrange = False
                    # Save the Window Positions prior to Presentation Mode rearrangement
                    self.WindowPositions = [self.MenuWindow.GetRect(),
                                            self.VisualizationWindow.GetDimensions(),
                                            self.VideoWindow.GetDimensions(),
                                            self.TranscriptWindow.GetDimensions(),
                                            self.DataWindow.GetDimensions()]
                # Hide the Menu Window
                self.MenuWindow.Show(False)
                # Hide the Visualization Window
                self.VisualizationWindow.Show(False)
                # Hide the Data Window
                self.DataWindow.Show(False)
                # Determine the size of the screen
                (left, top, width, height) = wx.ClientDisplayRect()

                # See if Presentation Mode is set to "Video Only"
                if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_VIDEO):
                    # Hide the Transcript Window
                    self.TranscriptWindow.Show(False)
                    # Set the Video Window to take up almost the whole Client Display area
                    self.VideoWindow.SetDims(left + 2, top + 2, width - 4, height - 4)
                    # If there is a PlayAllClipsWindow, reset it's size and layout
                    if self.PlayAllClipsWindow != None:
                        # Set the Window Position in the PlayAllClips Dialog
                        self.PlayAllClipsWindow.xPos = left + 2
                        self.PlayAllClipsWindow.yPos = height - 58
                        self.PlayAllClipsWindow.SetRect(wx.Rect(self.PlayAllClipsWindow.xPos, self.PlayAllClipsWindow.yPos, width - 4, 56))
                        # Make the PlayAllClipsWindow the focus
                        self.PlayAllClipsWindow.SetFocus()

                # See if Presentation Mode is set to "Video and Transcript"
                if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_TRANS):
                    # Set the Video Window to take up the top 70% of the Client Display Area
                    self.VideoWindow.SetDims(left + 2, top + 2, width - 4, int(0.7 * height) - 3)
                    # Set the Transcript Window to take up the bottom 30% of the Client Display Area
                    self.TranscriptWindow.SetDims(left + 2, int(0.7 * height) + 1, width - 4, int(0.3 * height) - 4)
                    # If there is a PlayAllClipsWindow, reset it's size and layout
                    if self.PlayAllClipsWindow != None:
                        # Set the Window Position in the PlayAllClips Dialog
                        self.PlayAllClipsWindow.xPos = left + 2
                        self.PlayAllClipsWindow.yPos = int(0.7 * height) - 58
                        self.PlayAllClipsWindow.SetRect(wx.Rect(self.PlayAllClipsWindow.xPos, self.PlayAllClipsWindow.yPos, width - 4, 56))
                        # Make the PlayAllClipsWindow the focus
                        self.PlayAllClipsWindow.SetFocus()
        

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
        return self.TranscriptWindow.GetTranscriptDims()

    def GetCurrentTranscriptObject(self):
        """ Returns a Transcript Object for the Transcript currently loaded in the Transcript Editor """
        return self.TranscriptWindow.GetCurrentTranscriptObject()

    def GetDatabaseTreeTabObjectNodeType(self):
        return self.DataWindow.DBTab.tree.GetObjectNodeType()

    def SetDatabaseTreeTabCursor(self, cursor):
        self.DataWindow.DBTab.tree.SetCursor(wx.StockCursor(cursor))

    def GetVideoPosition(self):
        """ Returns the current Time Code from the Video Window """
        return self.VideoWindow.GetCurrentVideoPosition()
        
    def UpdateVideoPosition(self, currentPosition):
        """ This method accepts the currentPosition from the video window and propogates that position to other objects """
        # If we do not already have a cursor position saved, and there is a defined cursor position, save it
        if (self.TranscriptWindow.dlg.editor.cursorPosition == 0) and \
           (self.TranscriptWindow.dlg.editor.GetCurrentPos() != 0) and \
           (self.TranscriptWindow.dlg.editor.GetSelection() != (0, 0)):
            self.TranscriptWindow.dlg.editor.cursorPosition = (self.TranscriptWindow.dlg.editor.GetCurrentPos(), self.TranscriptWindow.dlg.editor.GetSelection())

        if self.VideoEndPoint > 0:
            mediaLength = self.VideoEndPoint - self.VideoStartPoint
        else:
            mediaLength = self.VideoWindow.GetMediaLength()
        self.VisualizationWindow.UpdatePosition(currentPosition)

        # Update Transcript position.  If Transcript position changes,
        # then also update the selected Clips tab in the Data window.
        # NOTE:  self.IsPlaying() check added because the SelectedEpisodeClips Tab wasn't updating properly
        if (not self.IsPlaying()) or (self.TranscriptWindow.UpdatePosition(currentPosition)):
            if self.DataWindow.SelectedEpisodeClipsTab != None:
                self.DataWindow.SelectedEpisodeClipsTab.Refresh(currentPosition)

        
    def GetMediaLength(self, entire = False):
        """ This method returns the length of the entire video/media segment """
        if not(entire): # Return segment length
            if self.VideoEndPoint == 0:
                mediaLength = self.VideoWindow.GetMediaLength() - self.VideoStartPoint
            else:
                if self.VideoEndPoint - self.VideoStartPoint > 0:
                    mediaLength = self.VideoEndPoint - self.VideoStartPoint
                else:
                    mediaLength = self.VideoWindow.GetMediaLength() - self.VideoStartPoint
            return mediaLength
        else: # Return length of entire video 
            return self.VideoWindow.GetMediaLength()
        
    def UpdateVideoWindowPosition(self, left, top, width, height):
        """ This method receives screen position and size information from the Video Window and adjusts all other windows accordingly """
        if TransanaGlobal.configData.autoArrange:
            # Visualization Window adjusts WIDTH only to match shift in video window
            (wleft, wtop, wwidth, wheight) = self.VisualizationWindow.GetDimensions()
            self.VisualizationWindow.SetDims(wleft, wtop, left - wleft - 4, wheight)

            # NOTE:  We only need to trigger Visualization and Data windows' SetDims method to resize everything!

            # Data Window matches Video Window's width and shifts top and height to accommodate shift in video window
            (wleft, wtop, wwidth, wheight) = self.DataWindow.GetDimensions()
            self.DataWindow.SetDims(left, top + height + 4, width, wheight - (top + height + 4 - wtop))

            # Play All Clips Window matches the Data Window's WIDTH
            if self.PlayAllClipsWindow != None:
                (parentLeft, parentTop, parentWidth, parentHeight) = self.DataWindow.GetRect()
                (left, top, width, height) = self.PlayAllClipsWindow.GetRect()
                if (parentWidth != width):
                    self.PlayAllClipsWindow.SetDimensions(parentLeft, top, parentWidth, height)

    def UpdateWindowPositions(self, sender, X, YUpper=-1, YLower=-1):
        """ This method updates all window sizes/positions based on the intersection point passed in.
            X is the horizontal point at which the visualization and transcript windows end and the
            video and data windows begin.
            YUpper is the vertical point where the visualization window ends and the transcript window begins.
            YLower is the vertical point where the video window ends and the data window begins. """

        # If Auto-Arrange is enabled, resizing one window may alter the positioning of others.
        if TransanaGlobal.configData.autoArrange:

            if YUpper == -1:
                (wleft, wtop, wwidth, wheight) = self.VisualizationWindow.GetDimensions()
                YUpper = wheight + wtop      
            if YLower == -1:
                (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                YLower = wheight + wtop
                
            if sender != 'Visualization':
                # Adjust Visualization Window
                (wleft, wtop, wwidth, wheight) = self.VisualizationWindow.GetDimensions()
                self.VisualizationWindow.SetDims(wleft, wtop, X - wleft, YUpper - wtop)

            if sender != 'Transcript':
                # Adjust Transcript Window
                (wleft, wtop, wwidth, wheight) = self.TranscriptWindow.GetDimensions()
                self.TranscriptWindow.SetDims(wleft, YUpper + 4, X - wleft, wheight + (wtop - YUpper - 4))

            if sender != 'Video':
                # Adjust Video Window
                (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                self.VideoWindow.SetDims(X + 4, wtop, wwidth + (wleft - X - 4), YLower - wtop)

            if sender != 'Data':
                # Adjust Data Window
                (wleft, wtop, wwidth, wheight) = self.DataWindow.GetDimensions()
                self.DataWindow.SetDims(X + 4, YLower + 4, wwidth + (wleft - X - 4), wheight + (wtop - YLower - 4))

    def VideoSizeChange(self):
        """ Signal that the Video Size has been changed via the Options > Video menu """
        # Resize the video window.  This will trigger changes in all the other windows as appropriate.
        self.VideoWindow.frame.OnSizeChange()

    def SaveTranscript(self, prompt=0, cleardoc=0):
        """Save the Transcript to the database if modified.  If prompt=1,
        prompt the user to confirm the save.  Return 1 if Transcript was
        saved or unchanged, and 0 if user chose to discard changes.  If
        cleardoc=1, then the transcript will be cleared if the user chooses
        to not save."""
        # Was the document modified?
        if self.TranscriptWindow.TranscriptModified():
            result = wx.ID_YES
           
            if prompt:
                dlg = wx.MessageDialog(None, \
                    _("Transcript has changed.  Do you want to save it before continuing?"), \
                    _("Question"), wx.YES_NO | wx.ICON_QUESTION)
                result = dlg.ShowModal()
                dlg.Destroy()
            
            if result == wx.ID_YES:
                self.TranscriptWindow.SaveTranscript()
                return 1
            else:
                if cleardoc:
                    self.TranscriptWindow.ClearDoc()
                return 0
        return 1

    def SaveTranscriptAs(self):
        """Export the Transcript to an RTF file."""
        dlg = wx.FileDialog(None, wildcard="*.rtf", style=wx.SAVE)
        if dlg.ShowModal() == wx.ID_OK:
            fname = dlg.GetPath()
            if os.path.exists(fname):
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A file named "%s" already exists.  Do you want to replace it?'), 'utf8')
                else:
                    prompt = _('A file named "%s" already exists.  Do you want to replace it?')
                dlg2 = wx.MessageDialog(None, prompt % fname,
                                        _('Transana Confirmation'), style = wx.YES_NO | wx.ICON_QUESTION | wx.STAY_ON_TOP)
                dlg2.CentreOnScreen()
                if dlg2.ShowModal() == wx.ID_YES:
                    self.TranscriptWindow.SaveTranscriptAs(fname)
                dlg2.Destroy()
            else:
                self.TranscriptWindow.SaveTranscriptAs(fname)
        dlg.Destroy()

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
        # Load the specified Clip record
        tempClip = Clip.Clip(clipNum)
        # Prepare the Node List for removing the Keyword Example Node
        nodeList = (_('Keywords'), keywordGroup, keyword, tempClip.id)
        # Call the DB Tree's delete_Node method.  Include the Clip Record Number so the correct Clip entry will be removed.
        self.DataWindow.DBTab.tree.delete_Node(nodeList, 'KeywordExampleNode', tempClip.number)

    def UpdateDataWindowKeywordsTab(self):
        """ Update the Keywords Tab in the Data Window """
        # If the Keywords Tab is the currently displayed tab ...
        if self.DataWindow.nb.GetPageText(self.DataWindow.nb.GetSelection()) == _('Keywords'):
            # ... then refresh the Tab
            self.DataWindow.KeywordsTab.Refresh()

    def ChangeLanguages(self):
        """ Update all screen components to reflect change in the selected program language """
        self.ClearAllWindows()

        # Let's look at the issue of database encoding.  If it's UTF-8, don't change anything.
        if TransanaGlobal.encoding != 'utf8':
            # If it's not UTF-*, then if it is Russian, use KOI8r
            if TransanaGlobal.configData.language == 'ru':
                newEncoding = 'koi8_r'
            # If it's Chinese, use the appropriate Chinese encoding
            elif TransanaGlobal.configData.language == 'zh':
                newEncoding = TransanaConstants.chineseEncoding
            # If it's Japanese, use cp932
            elif TransanaGlobal.configData.language == 'ja':
                newEncoding = 'cp932'
            # If it's Korean, use cp949
            elif TransanaGlobal.configData.language == 'ko':
                newEncoding = 'cp949'
            # Otherwise, fall back to Latin-1
            else:
                newEncoding = 'latin1'
            # If we're changing encodings, we need to do a little work here!
            if newEncoding != TransanaGlobal.encoding:
                msg = _('Database encoding is changing.  To avoid potential data corruption, \nTransana must close your database before proceeding.')
                tmpDlg = Dialogs.InfoDialog(None, msg)
                tmpDlg.ShowModal()
                tmpDlg.Destroy()

                # We should get a new database.  This call will actually update our encoding if needed!
                self.GetNewDatabase()
                
        self.MenuWindow.ChangeLanguages()
        self.VisualizationWindow.ChangeLanguages()
        self.DataWindow.ChangeLanguages()
        # Updating the Data Window automatically updates the Headers on the Video and Transcript windows!
        self.TranscriptWindow.ChangeLanguages()

    def AdjustIndexes(self, adjustmentAmount):
        """ Adjust Transcript Time Codes by the specified amount """
        self.TranscriptWindow.AdjustIndexes(adjustmentAmount)

    def __repr__(self):
        tempstr = "Control Object contents:\nVideoFilename = %s\nVideoStartPoint = %s\nVideoEndPoint = %s"
        return tempstr % (self.VideoFilename, self.VideoStartPoint, self.VideoEndPoint)

