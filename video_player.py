# Copyright (C) 2006 - 2012 The Board of Regents of the University of Wisconsin System 
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

""" This unit handles interaction with the wxMedia Control used in displaying video
    on the Windows and Mac platforms.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "video_player DEBUG is ON!!!!"

UPDATE_PROGRESS_INTERVAL = 100

# import Python's OS module
import os
# import Python's sys module
import sys
# import Python's Traceback module for error debugging
import traceback
# import wxPython
import wx
# import the wxPython MediaCtrl
import wx.media

# Import Transana's Dialogs
import Dialogs
# import the Transana Global data
import TransanaGlobal
# import the Transana VideoWindow
import VideoWindow

# Declare the main VideoPlayer class, designed to interact with the rest of Transana
class VideoPlayer(wx.Panel):
    """ Media Player Panel control for the Video Window.  This is based on wxMediaCtrl. """
        
    def __init__(self, parent=None, pos=(30, 30), size=wx.Size(480, 420), includeCheckBoxes=False, offset=0, playerNum=-1, formPar=None):
        """ Initialize the Media Player Panel object """
        # In some instances, because of the way sizers work (I think), we have to distinguish between the FORM's parent, which may be
        # a Panel on form, and the CONTROL's parent, which has methods for handling changes in media position etc.
        # If the FORM Parent does NOT differ from the CONTROL Parent ...
        if formPar == None:
            # ... then we can use the CONTROL parent as the Form Parent
            formPar = parent
        # We need to know the Media Player CONTROL's parent
        self.parent = parent
        # Create a Panel to hold the Media Player, using the FORM parent.  The panel wants to process characters.
        wx.Panel.__init__(self, formPar, -1, size=(358, 285), style=wx.WANTS_CHARS)
        # Remember the includeCheckBoxes setting
        self.includeCheckBoxes = includeCheckBoxes
        # Remember the (optional) offset value
        self.offset = offset
        # remember the (optional) playerNum value
        self.playerNum = playerNum
        # Initialize the Play State
        self.playState = None
        # Initialize the Loading State variable
        self.isLoading = False
        # We need a flag that indicates that we want to Play as soon as the Load is complete
        self.playWhenLoaded = False
        # We're having trouble knowing the media length.  Let's track whether we know it yet.
        self.mediaLengthKnown = False
        # The Mac requires the following so that versions look more similar across platforms
        if 'wxMac' in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)
        # Set the background color to Black.  
        self.SetBackgroundColour(wx.WHITE)
        # If Transana is installed in a folder that contains accented characters, we need to adjust the programDir to handle that
        if ('unicode' in wx.PlatformInfo) and isinstance(TransanaGlobal.programDir, str):
            TransanaGlobal.programDir = TransanaGlobal.programDir.decode('cp1250')
        # Build the full file name for the Splash Screen image
        imgFileName = os.path.join(TransanaGlobal.programDir, 'images', 'splash.gif')
        # Get the Splash Srceen image
        self.splashImage = wx.Image(imgFileName)
        # Get the size of the Video Player Panel
        (width, height) = self.GetClientSize()
        # Get a copy of the original image so the original is preserved, doesn't lose resolution.
        self.splash = self.splashImage.Copy()
        # Rescale the image to the size of the panel
        self.splash.Rescale(width, height)
        # Convert the image to a bitmap
        self.graphic = wx.BitmapFromImage(self.splash)
        # Place the Splash graphic on the panel
        dc = wx.BufferedDC(wx.ClientDC(self), self.graphic)

        # Define timer for progress notification to other windows
        timerID = wx.NewId()
        self.ProgressNotification = wx.Timer(self, timerID)
        wx.EVT_TIMER(self, timerID, self.OnProgressNotification)
        
        # Define the Key Down Event Handler
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)

        # set default back end to QuickTime, as then either type of media can be opened.  If we
        # set it to DirectShow for Windows, we can't later load Quicktime files
        if 'wxMSW' in wx.PlatformInfo:
            self.backend = wx.media.MEDIABACKEND_QUICKTIME
        elif 'wxGTK' in wx.PlatformInfo:
            self.backend = wx.media.MEDIABACKEND_GSTREAMER
        else:
            self.backend = wx.media.MEDIABACKEND_QUICKTIME
        # Initialize the Media Player to None
        self.movie = None
        # Create the Media Player
        self.CreateMediaPlayer()
        # Hide the Media Player
        self.movie.Show(False)
        # Bind the OnMediaLoaded event, so we know when video loading is COMPLETE
        self.Bind(wx.media.EVT_MEDIA_LOADED, self.OnMediaLoaded)
        # Bind the OnMediaStop event so we know when the media file has stopped playing.
        self.Bind(wx.media.EVT_MEDIA_STOP, self.OnMediaStop)

        self.Bind(wx.EVT_PAINT, self.OnPaint)
        self.Bind(wx.EVT_RIGHT_UP, self.OnRightUp)
        
        # Bind the Close event handler
        wx.EVT_CLOSE(self, self.OnCloseWindow)
        # Initialize some variables
        self.done = 0
        self.VideoStartPoint = 0
        self.VideoEndPoint = 0
        self.FileName = ""
        # Set initial playback speed based on configuration setting
        self.Rate = TransanaGlobal.configData.videoSpeed / 10.0

    def CreateMediaPlayer(self, flNm=""):
        """ Create the actual Media Player component """
        # If the Progress Notification timer is running, STOP it!!
        self.ProgressNotification.Stop()
        # Freeze the interface to speed up updates
        self.Freeze()
        
        # If there is a media player already defined ...
        if self.movie:
            # ...destroy it ...
            self.movie.Destroy()
            # ... and remove all evidence of it!
            self.movie = None
            # If we are including check boxes ...
            if self.includeCheckBoxes:
                # ... then destroy the existing checkboxes.
                self.includeInClip.Destroy()
                self.playAudio.Destroy()
                # If we are on a Player with a Syncronization indicator ...
                if self.playerNum > 1:
                    # ... destroy the synchronization indicator
                    self.synch.Destroy()
        # Create a new Media Player control
        self.movie = wx.media.MediaCtrl(self, szBackend=self.backend)
        # Create a sizer to handle the media control
        box = wx.BoxSizer(wx.VERTICAL)
        box.Add(self.movie, 1, wx.EXPAND | wx.ALL, 0)
        # If we are including check boxes ...
        if self.includeCheckBoxes:
            # ... create a checkbox sizer
            hSizer = wx.BoxSizer(wx.HORIZONTAL)
            # If we have multiple players ...
            if self.playerNum > 0:
                # add a left spacer
                hSizer.Add((5, 0), 0)
                # Add a color swatch to indicate the waveform color
                colorDefs = (wx.RED, wx.BLUE, wx.GREEN, wx.CYAN)
                # Create an empty bitmap
                bmp = wx.EmptyBitmap(16, 16)
                # Create a Device Context for manipulating the bitmap
                dc = wx.BufferedDC(None, bmp)
                # Begin the drawing process
                dc.BeginDrawing()
                # Paint the bitmap white
                dc.SetBackground(wx.Brush(wx.WHITE))
                # Clear the device context
                dc.Clear()
                # Define the pen to draw with
                pen = wx.Pen(wx.Colour(0, 0, 0), 1, wx.SOLID)
                # Set the Pen for the Device Context
                dc.SetPen(pen)
                # Define the brush to paint with in the defined color
                brush = wx.Brush(colorDefs[self.playerNum - 1])
                # Set the Brush for the Device Context
                dc.SetBrush(brush)
                # Draw a black border around the color graphic, leaving a little white space
                dc.DrawRectangle(1, 1, 14, 14)
                # End the drawing process
                dc.EndDrawing()
                # Select a different object into the Device Context, which allows the bmp to be used.
                dc.SelectObject(wx.EmptyBitmap(5,5))
                # Create a graphic object
                graphic = wx.StaticBitmap(self, -1, bmp)
                # Add it to the sizer
                hSizer.Add(graphic, 0)
            # If we've got mulitple players and we're on one later than the first ...
            if self.playerNum > 1:
                # ... create a synchronization indicator
                self.synch = wx.StaticText(self, -1, '0')
                # ... and add the sychronization indicator to the sizer
                hSizer.Add(self.synch, 0, wx.LEFT, 2)
            # add a spacer
            hSizer.Add((0, 0), 1)
            # Add a checkbox for VIDEO clip inclusion
            self.includeInClip = wx.CheckBox(self, -1, _("Include in Clip"))
            # It defaults to "checked"
            self.includeInClip.SetValue(True)
            # Bind a CheckBox Event Handler
            self.includeInClip.Bind(wx.EVT_CHECKBOX, self.OnIncludeCheck)
            # Add the check box and some space to the sizer
            hSizer.Add(self.includeInClip, 0, wx.TOP, 1)
            hSizer.Add((5, 0))
            # Add a checkbox for AUDIO play/mute/inclusion
            self.playAudio = wx.CheckBox(self, -1, _("Play Audio"))
            # It defaults to "checked"
            self.playAudio.SetValue(True)
            # Bind a CheckBox Event Handler
            self.playAudio.Bind(wx.EVT_CHECKBOX, self.OnAudioCheck)
            # Add the check box to the sizer
            hSizer.Add(self.playAudio, 0, wx.TOP, 1)
            # Add a right spacer
            hSizer.Add((0, 0), 1)
            # If we have multiple players ...
            if self.playerNum > 0:
                # Create a graphic object
                graphic = wx.StaticBitmap(self, -1, bmp)
                # Add it to the sizer
                hSizer.Add(graphic, 0)
                # add a left spacer
                hSizer.Add((5, 0), 0)
            # Add the checkbox sizer to the panel's main sizer
            box.Add(hSizer, 0, wx.EXPAND)
        # Set the main sizer
        self.SetSizer(box)
        # Re-bind the key down event
        self.movie.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        # Re-bind the right-click-up event
        self.movie.Bind(wx.EVT_RIGHT_UP, self.OnRightUp)
        # Thaw the interface when updates are complete
        self.Thaw()
        
        # Adjust the media player size to fit the window, now that it's been laid out.
        self.OnSize(None)

    def SetFilename(self, filename, offset=0):
        """ Load a file in a media player.  If an offset is passed, the video will be positioned to that offset. """
        # If a file name is specified
        if filename != '':
            # Show the media player control
            self.movie.Show(True)
            # If we're including check boxes ...
            if self.includeCheckBoxes:
                # ... show the check boxes!
                self.includeInClip.Show(True)
                self.playAudio.Show(True)
            # If we're on Windows, we may need to change media player back ends!
            if ('wxMSW' in wx.PlatformInfo):
                # Break out the file extension
                (videoFilename, videoExtension) = os.path.splitext(filename)
                # If the extension is one that requires the QuickTime Player ...
                if videoExtension.lower() in ['.mov', '.mp4', '.m4v', '.aac']:
                    # ... indicate we need the QuickTime back end
                    backendNeeded = wx.media.MEDIABACKEND_QUICKTIME
                # If the extension is one that requires the Windows Media Player ...
                elif videoExtension.lower() in ['.wmv', '.wma']:
                    # ... indicate we need the WMP10 back end
                    backendNeeded = wx.media.MEDIABACKEND_WMP10
                # If we have MPEG and are on Windows XP or earlier ...
                elif (videoExtension.lower() in ['.mpg', '.mpeg']) and ((sys.getwindowsversion()[0] < 6)):
                    # ... indicate we need the DirectShow back end
                    backendNeeded = wx.media.MEDIABACKEND_DIRECTSHOW
                # If not QuickTime or Windows Media or MPEG on XP ...
                else:
                    # ... Check the Configuration data to see which wxMediaPlayer back end we need.
                    if TransanaGlobal.configData.mediaPlayer == 1:
                        backendNeeded = wx.media.MEDIABACKEND_DIRECTSHOW
                    else:
                        backendNeeded = wx.media.MEDIABACKEND_WMP10
                # If the current back end is different from the needed back end ...
                if self.backend != backendNeeded:
                    # ... signal the back end that we need ...
                    self.backend = backendNeeded
                    # ... and recreate the Media Player using the new back end type.
                    self.CreateMediaPlayer(filename)

                if DEBUG:
                    if self.backend == wx.media.MEDIABACKEND_QUICKTIME:
                        prompt = 'QuickTime'
                    elif self.backend == wx.media.MEDIABACKEND_WMP10:
                        prompt = 'WMP10'
                    elif self.backend == wx.media.MEDIABACKEND_DIRECTSHOW:
                        prompt = 'DirectShow'
                    tmpDlg = Dialogs.InfoDialog(self, prompt)
                    tmpDlg.ShowModal()
                    tmpDlg.Destroy()
                    
            # We need to have a flag that indicates that the video is in the process of loading
            self.isLoading = True
            # We don't know the media length in some back ends until the load is complete.
            self.mediaLengthKnown = False

            # QuickTime on Windows cannot correctly load files if the PATH contains non-CP1252 characters!
            # A chinese file name in a CP1252 path is OK, but an english file name in a chinese path won't load.
            # This code fixes that problem.

            # First, detect Windows and QuickTime
            if ('wxMSW' in wx.PlatformInfo) and (self.backend == wx.media.MEDIABACKEND_QUICKTIME):
                # Change the Python Encoding to cp1252
                wx.SetDefaultPyEncoding('cp1252')
                # Replace backslashes with forward slashes
                filename = filename.replace('\\', '/')
                # Get the Current Working Directory
                originalCWD = os.getcwd()
                # Divide the path and the file name from each other
                (currentPath, currentFileName) = os.path.split(filename)
                # Change the Current Directory to the file's location
                os.chdir(currentPath)

                if DEBUG:
                    print "video_player.SetFileName(1):"
                    print "Path changed to ", currentPath.encode('utf8'), currentPath.encode(sys.getfilesystemencoding()) == os.getcwd(), type(currentPath), type(os.getcwd())
                
                # Encode just the File Name portion of the path
                tmpfilename = currentFileName.encode(sys.getfilesystemencoding())

                if DEBUG:
                    print "tmpfilename =", tmpfilename, os.path.exists(tmpfilename), os.path.exists(currentFileName)
                
            # If we're not on Windows OR we aren't using QuickTime ...
            else:

                if DEBUG:
                    print "video_player.SetFileName(2):"
                    print "filename.encode('utf8') =", filename.encode('utf8')
                
                # ... then the unencoded file name including the full path works just fine!
                tmpfilename = filename
            
            # Try to load the file in the media player.  If successful ...
            if self.movie.Load(tmpfilename):
                # ... remember the file name
                self.FileName = filename
                # Initialize the start and end points
                if offset <= 0:
                    self.VideoStartPoint = 0
                else:
                    self.VideoStartPoint = offset
                self.VideoEndPoint = 0
            # If the file load FAILS ...
            else:
                # Display an error message
                msg = _('Transana was unable to load media file "%s".')
                if 'unicode' in wx.PlatformInfo:
                    msg = unicode(msg, 'utf8')
                dlg = Dialogs.ErrorDialog(self, msg % filename)
                dlg.ShowModal()
                dlg.Destroy()
                # Signal that the media file did not load
                self.isLoading = False
                self.FileName = ""

            # Again, detect Windows and QuickTime
            if ('wxMSW' in wx.PlatformInfo) and (self.backend == wx.media.MEDIABACKEND_QUICKTIME):
                # Reset the Default Python encoding back to UTF8
                wx.SetDefaultPyEncoding('utf8')
                # Reset the Current Working Directory to what it used to be
                os.chdir(originalCWD)

        # If no filename is specified ...
        else:
            # "unload" what might be in the Media Player
            self.movie.Load('')
            # Hide the Media Player
            self.movie.Show(False)
            # If there are checkboxes (multiple videos) ...
            if self.includeCheckBoxes:
                # ... hide them too.
                self.includeInClip.Show(False)
                self.playAudio.Show(False)
            # If the graphic has been hidden, it may be the wrong size.  This fixes that.
            self.OnSize(None)
        # Update the Panel on screen
        self.Update()

        # If on Windows Vista, Windows 7, or later and we are using DirectShow  ...
        if ('wxMSW' in wx.PlatformInfo) and (sys.getwindowsversion()[0] >= 6) and (self.backend == wx.media.MEDIABACKEND_DIRECTSHOW):
            # ... there's a bug in wxPython that prevents DirectShow from triggering wx.EVT_MEDIA_LOADED, so we'll call it
            #     manually here.  wx.CallAfter is not sufficient, as the video needs time to load and CallAfter doesn't wait
            #     long enough.
            wx.CallLater(2000, self.OnMediaLoaded, None)
    
    def GetFilename(self):
        """ Get the name of the currently-loaded media file """
        return self.FileName

    def SetVideoStartPoint(self, TimeCode):
        """Set the current position in ms."""
        # If there's a defined video ...
        if self.movie != None:
            # TimeCode must be an int on the Mac.
            if not isinstance(TimeCode, int):
                TimeCode = int(TimeCode)
            # Set the video start point locally.  DO NOT adjust the time code for the local media player's offset.
            self.VideoStartPoint = TimeCode
            # Set the video player's current position (which DOES adjust for offset)
            self.SetCurrentVideoPosition(TimeCode)
            if self.parent != None:
                # If this is the ONLY media player or the FIRST media player ...
                if self.playerNum in [-1, 1]:
                    # notify the rest of Transana of the change.
                    self.parent.ControlObject.UpdateVideoPosition(self.VideoStartPoint + self.parent.globalOffset)
                   
    def GetVideoStartPoint(self):
        """ Get the current Video Start Point """
        return self.VideoStartPoint
        
    def SetVideoEndPoint(self, TimeCode):
        """ Set the Video End Point """
        # Adjust the time code for the local media player's offset
        self.VideoEndPoint = TimeCode - self.offset

    def GetVideoEndPoint(self):
        """ Get the current Video End Point """
        return self.VideoEndPoint
 
    def GetTimecode(self):
        """Return the current position in ms."""
        # Report the current video position, adjusted for the local media player offset.  This is the VIRTUAL TimeCode position,
        # rather than the ACTUAL TimeCode Position.
        return self.movie.Tell() + self.offset

    def GetMediaLength(self):
        """ Get the length of the current media file, if it's known """
        if not self.mediaLengthKnown:
            return -1
        else:
            # If there's no defined video ...
            if self.movie == None:
                # ... then there's a problem.
                self.mediaLengthKnown = False
                return -1
            else:
                return self.movie.Length()

    def GetState(self):
        """ Get the current state of the Media Player """
        # If there's a defined video ...
        if self.movie != None:
            return self.movie.GetState()
        else:
            return -1
    
    def SetCurrentVideoPosition(self, TimeCode):
        """ Set the current video position. """
        try:
            # TimeCode must be an int on the Mac.
            if not isinstance(TimeCode, int):
                TimeCode = int(TimeCode)
            # On the Mac, the start point can't be less than 0
            if TimeCode < self.offset:
                TimeCode = 0
            else:
                TimeCode -= self.offset
            # On the Mac, the start point can't be after the end.  A 5 ms adjustment (1/6 of a frame) is too small to be noticable.
            if (self.mediaLengthKnown) and (TimeCode > self.GetMediaLength() - 5):
                self.VideoStartPoint = self.GetMediaLength() - 5
                TimeCode = self.GetMediaLength() - 5
            # Find the appropriate spot in the media file
            self.movie.Seek(TimeCode)
        # Trap the PyDeadObjectError, mostly during Play All Clips on PPC Mac
        except wx._core.PyDeadObjectError, e:
            pass
    
    # NOTE:  GetPlayBackSpeed and SetPlayBackSpeed seem to disagree with each other by a factor of 10!!!!
    def GetPlayBackSpeed(self):
        """ Get the current Playback Speed """
        return self.Rate
    
    # NOTE:  GetPlayBackSpeed and SetPlayBackSpeed seem to disagree with each other by a factor of 10!!!!
    def SetPlayBackSpeed(self, playbackSpeed):
        """ Sets the play back speed."""
        # Reset the Rate value.  The Options Setting screen uses values 1 - 20 to represent 0.1 to 2.0, so divide by 10!
        self.Rate = playbackSpeed / 10.0
        # Set the playback Rate, if we're NOT using the QuickTime Back End.  If we are, that would cause the video to
        # play, so we'll skip it.
        if (self.backend != wx.media.MEDIABACKEND_QUICKTIME) or self.IsPlaying():
            self.movie.SetPlaybackRate(self.Rate)
        # Let's adjust the configuration value as well.
        TransanaGlobal.configData.videoSpeed = int(self.Rate * 10)

    def IsPlaying(self):
        """ Is the media player currently playing? """
        return self.GetState() == wx.media.MEDIASTATE_PLAYING

    def IsPaused(self):
        """ Is the media player currently paused? """
        # When we have multiple media players, not all of which have started playing, some media players may be
        # paused while others may be stopped.  Functionally, we need Paused and Stopped to be treated the same.
        return self.GetState() in [wx.media.MEDIASTATE_PAUSED, wx.media.MEDIASTATE_STOPPED]

    def IsStopped(self):
        """ Is the media player currently stopped? """
        # When we have multiple media players, not all of which have started playing, some media players may be
        # paused while others may be stopped.  Functionally, we need Paused and Stopped to be treated the same.
        return self.GetState() in [wx.media.MEDIASTATE_PAUSED, wx.media.MEDIASTATE_STOPPED]

    def IsLoading(self):
        """ Is the media player currently in the process of loading a media file? """
        return self.isLoading

    def PlaySegment(self, t0, t1):
        """Play the segment of the video from timecode t0 to t1."""
        # Stop playback, if needed.
        self.Stop()
        # Set the start point
        self.SetVideoStartPoint(t0)
        # Set the end point
        self.SetVideoEndPoint(t1)
        # Start media playback
        self.Play()
            
    def Play(self):
        """Play the video."""
        # Play All Clips is having trouble starting because "Play()" is called before loading is complete.
        # This structure allows the Play() to be called once the load is done.
        if self.IsLoading():
            self.playWhenLoaded = True
        else:
            self.playWhenLoaded = False
            # The QuickTime back end behaves differently than Windows Media Player.  Play() always plays at normal speed.
            # SetPlaybackRate() starts playback (!) at the requested speed.  But ... SetPlaybackRate doesn't
            # set the Play State correctly, I don't think, due to a bug in wxMediaCtrl, so we'll call both in quick. 
            # succession to get the behavior we want.  When using the DirectShow back end, Play will work regardless
            # of the selected playback speed, so we only need to alter the call to Play() here if we're in QuickTime
            # at a rate other than 1.0
            if (self.Rate != 1.0):  # (self.backend == wx.media.MEDIABACKEND_QUICKTIME) and 
                self.movie.Play()
                self.movie.SetPlaybackRate(self.Rate)
            else:
                self.movie.Play()
        # Start the Progress Notification timer when play starts
        self.ProgressNotification.Start(UPDATE_PROGRESS_INTERVAL)

    def Pause(self):
        """Pause the video. (pause does not reposition the video)"""
        # pause the media playback
        self.movie.Pause()
        # Stop the progress notification times when playback pauses
        self.ProgressNotification.Stop()
        # Call OnProgressNotification ONCE to signal the stop (needed for Looping)
        wx.CallAfter(self.OnProgressNotification, None)

    def Stop(self):
        """Stop the video. (stop playback and return position to Video Start Point)."""
        # Stop the media playback
        self.movie.Stop()
        # Stop the progress notification times when playback stops
        self.ProgressNotification.Stop()
        # Call OnProgressNotification ONCE to signal the stop (needed for Looping)
        wx.CallAfter(self.OnProgressNotification, None)
        # Reset the media position to the StartPoint
        self.SetCurrentVideoPosition(self.VideoStartPoint)
        # Signal Transana that the position has changed.
        self.PostPos()

    def OnCloseWindow(self, event):
        """ Close event handler """
        # Stop the Progress Notification timer
        self.ProgressNotification.Stop()
        # Destroy the current Panel
        self.Destroy()

    def OnMediaLoaded(self, event):
        """ This event is triggered when media loading is complete.  Only then are certain things known about the media file. """
        # We can now know the media length
        self.mediaLengthKnown = True
        # Once the video is loaded, we can determine its size and should react to that.
        self.parent.OnSizeChange()
        self.Update()
        # Often, the Seek has also not occurred correctly.  This detects and fixes that problem.
        if self.VideoStartPoint != self.GetTimecode():
            self.SetCurrentVideoPosition(self.VideoStartPoint)
        # Indicate that the loading process is complete.  (Moved later due to a PlayAllClips problem.)
        self.isLoading = False
        # Play All Clips is having trouble starting because "Play()" is called before loading is complete.
        # This structure allows the Play() to be called once the load is done.
        if self.playWhenLoaded:
            self.Play()
        # If on Windows ...
        elif "wxMSW" in wx.PlatformInfo:
            # ... prevent automatic playback upon load.
            self.Stop()

    def OnPaint(self, event):
        # If the movie player is NOT shown, we need to show the background image
        if not self.movie.IsShown():
            # Get the size of the Video Player Panel
            (width, height) = self.GetClientSize()
            # Get a COPY of the original splash image
            self.splash = self.splashImage.Copy()
            # Loading two consecutive audio files crashes unless we make sure that the window client height is greater than 0
            if height > 0:
                # Rescale the image to the size of the panel
                self.splash.Rescale(width, height)
                # Convert the image to a bitmap
                self.graphic = wx.BitmapFromImage(self.splash)
                # Show the image on the panel background
                dc = wx.BufferedPaintDC(self, self.graphic)
        event.Skip()

    def OnKeyDown(self, event):
        """ Handle Key Down events """
        # See if the Control Object wants to handle the key that was pressed, which is only will if the parent object is a VideoWindow
        if isinstance(self.parent, VideoWindow.VideoWindow) and self.parent.ControlObject.ProcessCommonKeyCommands(event):
            # If so, we're done here.  (Actually, we're done anyway.)
            return

    def OnRightUp(self, event):
        """ Right Mouse Button Up event handler """
        # let the parent control handle this one!
        event.Skip()
        # Explicitly call the parent's OnRightUp method.
        self.parent.OnRightUp(event)

    def OnSizeChange(self):
        """ OBSOLETE.  Handle size changes that are not event-driven """
        
        print "video_player.OnSizeChange() disabled as obsolete"
        
        # Only change the size of the video window if Auto Arrange is ON!
        if FALSE and TransanaGlobal.configData.autoArrange:
            (sizeX, sizeY) = self.movie.GetBestSize()
            # Now that we have a size, let's position the window
            #  Determine the screen size 
            rect = wx.Display(0).GetClientArea()  # wx.ClientDisplayRect()
            #  Get the current position of the Video Window
            pos = self.GetPosition()
            #  Establish the minimum width of the media player control 
            # (if media is audio-only, for example)
            minWidth = max(sizeX, 300)
            # use Movie Height
            minHeight = sizeY 
            # Adjust Video Size, unless you are showing the Splash Screen
            sizeAdjust = TransanaGlobal.configData.videoSize
            # Take Video Size Menu spec into account (50%, 66%, 100%, 150%, 200%)
            minWidth = int(minWidth * sizeAdjust / 100.0)
            minHeight = int(minHeight * sizeAdjust / 100.0)

            # We need to know the height of the Window Header to adjust the size of the Graphic Area
            headerHeight = self.GetSizeTuple()[1] - self.GetClientSizeTuple()[1]
            #  Set the dimensions of the video window as follows:
            #    left:    right-justify to the screen 
            #           (left side of screen - media player width - 3 pixel margin)
            #    top:     leave top position unchanged
            #    width:   use minimum media player width established above
            #    height:  use height of Movie  + ControlBarSize + headerHeight pixels for the player controls
            # rect[0] compensates if the Start Menu is on the left side of the screen
            if self.backend in [wx.media.MEDIABACKEND_DIRECTSHOW, wx.media.MEDIABACKEND_WMP10]:
                controlBarSize = 65
            elif self.backend == wx.media.MEDIABACKEND_QUICKTIME:
                controlBarSize = 10
            else:
                controlBarSize = 100  # This'll be quite noticable!
            self.SetDimensions(rect[0] + rect[2] - minWidth - 3, 
                               pos[1], 
                               minWidth, 
                               minHeight + controlBarSize + headerHeight)
            if self.parent != None:
                self.parent.UpdateVideoWindowPosition(rect[0] + rect[2] - minWidth - 3, 
                                                      pos[1], 
                                                      minWidth, 
                                                      minHeight + controlBarSize + headerHeight)

    def OnSize(self, event):
        """ Process Size Change event and notify the ControlObject """
        # if event is not None (which it can be if this is called from non-event-driven code rather than
        # from a real even, then we should process underlying OnSize events.
        if event != None:
            event.Skip()
        try:
            # We can't update the graphic's size if it's hidden, as that causes problems visually on the Mac.
            if self.movie and self.movie.IsShown():

                # Force the media player to preserve Aspect Ratio
                (sizex, sizey) = self.movie.GetBestSize()
                # If the video HAS a width ...
                if sizex > 0:
                    # ... determine the original aspect ratio of the media file
                    aspectRatio = float(sizey) / float(sizex)
                    # Reset the height of the media player AND the media itself based on width and original aspect ratio
                    self.SetSize((self.GetSize()[0], self.GetSize()[0] * aspectRatio))
                    self.movie.SetSize((self.GetSize()[0], self.GetSize()[0] * aspectRatio))
                    
                # ... refresh the media player
                self.movie.Refresh()
                self.Refresh()
        # Trap the PyDeadObjectError, mostly during Play All Clips on PPC Mac
        except wx._core.PyDeadObjectError:
            pass

    def OnMediaStop(self, event):
        """ This event is triggered when media play stops """
        # See if OTHER media players continue to play.  If so, leave this player at the video end position.  If NOT ...
        if not self.parent.IsPlaying():
            # ... reset the video position back to the Start Point
            wx.CallAfter(self.SetCurrentVideoPosition, self.VideoStartPoint)
            # Signal Transana that the video position has changed
            wx.CallAfter(self.PostPos)
        elif (self.VideoEndPoint <= 0) or (self.VideoEndPoint >= self.movie.Length() - 66):
            # Set the video to the last frame
            wx.CallAfter(self.SetCurrentVideoPosition, self.movie.Length() - 1)

    def OnProgressNotification(self, event):
        try:
            # Detect Play State Change and notify the VideoWindow.
            # See if the playState has changed.
            if (self.movie != None) and (self.playState != self.movie.GetState()):
                # If it has, communicate that to the VideoWindow.  The Mac doesn't seem to differentiate between
                # Stop and Pause the say Windows does, so pass "Stopped" for either Stop or Pause.
                if self.movie.GetState() != wx.media.MEDIASTATE_PLAYING:
                    if self.parent != None:
                        self.parent.UpdatePlayState(wx.media.MEDIASTATE_STOPPED)
                # Pass "Play" for play.
                else:
                    if self.parent != None:
                        self.parent.UpdatePlayState(wx.media.MEDIASTATE_PLAYING)
                # Update the local playState variable
                self.playState = self.movie.GetState()

            # Take this opportunity to see if there are any waiting events.  Leaving this out has the unfortunate side effect
            # of diabling the media player's control bar.
            # There is sometimes a problem with recursive calls to Yield; trap the exception ...
            wx.YieldIfNeeded()

            # The timer that calls this routine runs whether the video is playing or not.  We only need to think
            # about updating the rest of Transana if the video is playing.
            if self.IsPlaying():
                self.PostPos()
        # Trap the PyDeadObjectError, mostly during Play All Clips on PPC Mac
        except wx._core.PyDeadObjectError, e:
            pass
        # ... and ignore it!
        except:
            pass

    def PostPos(self):
        """ Inform other elements of the current video position """
        try:
            # If we are not shutting down Transana (to avoid a crash) ...
            if (self.parent != None) and (self.parent.ControlObject != None) and (not self.parent.ControlObject.shuttingDown):
                tc = self.GetTimecode()
                if self.playerNum == -1:
                    # ... then update the parent (Video Window)'s Video Position with the current time code
                    self.parent.UpdateVideoPosition(tc)
                else:
                    self.parent.UpdateVideoPosition(tc, self.playerNum)
        # Trap the PyDeadObjectError, mostly during Play All Clips on PPC Mac
        except wx._core.PyDeadObjectError, e:
            pass

    def OnIncludeCheck(self, event):
        """ Event handler for the includeInClip Checkbox """
        # If the includeInClip checkbox has been UNCHECKED ...
        if not self.includeInClip.GetValue():
            # ... remember the value of the PlayAudio checkbox ...
            self.includeAudioSetting = self.playAudio.GetValue()
            # ... and then make sure it is UNCHECKED.  If you're not including the media in the clip, you shouldn't hear it!
            self.SetAudioCheck(False)
        # If the includeInClip checkbox has been CHECKED ...
        else:
            # ... then restore the PlayAudio checkbox to its prior setting
            self.SetAudioCheck(self.includeAudioSetting)
        # Notify the parent of the change (so waveform can be updated)
        self.parent.VideoCheckboxChange()
        
    def OnAudioCheck(self, event):
        """ Event handler for Audio checkbox """
        # If the box has just been checked ...
        if self.playAudio.GetValue():
            # ... restore the audio by setting the volume to the previous level
            self.movie.SetVolume(1.0)
            # If the audio is included, the media file MUST be included in Clips.
            self.includeInClip.SetValue(True)
        # If the box has just been un-checked ...
        else:
            # ... set the audio level to 0.0
            self.movie.SetVolume(0.0)
        # Notify the parent of the change (so waveform can be updated)
        self.parent.VideoCheckboxChange()

    def SetAudioCheck(self, val):
        """ Set the Audio Checkbox value """
        # Set the value of the checkbox
        self.playAudio.SetValue(val)
        # Set the Audio Playback Level by triggering the AudioCheck event
        self.OnAudioCheck(None)

    def GetVideoCheckboxDataForClips(self):
        """ Return the data about this media player checkboxes needed for Clip Creation """
        if self.includeCheckBoxes:
            return (self.includeInClip.GetValue(), self.playAudio.GetValue())
        else:
            return None
