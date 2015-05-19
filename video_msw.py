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

"""This unit handles interaction with the Windows Media Player ActiveX Control 
used in displaying video on the Microsoft Windows platform."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare'

import wx
import string
import traceback

if __name__ == '__main__':
    __builtins__._ = wx.GetTranslation

import Dialogs
from TransanaConstants import *
import TransanaConstants
import TransanaGlobal

# Windows Media Playercontrol embedding class.
try:
    import WindowsMediaPlayer
except:
    import sys
    tb = traceback.format_exc()
    errormsg = "%s %s\n%s" %  (sys.exc_info()[0], sys.exc_info()[1], tb)
    raise ImportError(errormsg)   # "Can't load Windows Media Player")

class VideoFrame(wx.Dialog):
    """Video player dialog. Use the 'public' methods to control the player."""
    def __init__(self, parent, parentVideoWindow, 
                 pos=wx.DefaultPosition, size=wx.DefaultSize):
        wx.Dialog.__init__(self, parent, -1, _("Video"), 
                             pos=pos, size=size,
                             style=wx.CAPTION|wx.RESIZE_BORDER)
        
        # Define the parent abstract frame which is used to communicate with 
        # the rest of Transana
        self.parentVideoWindow = parentVideoWindow
        
        ######## Now create the ActiveX class and initialize everything ########
        #
        # Create an ActiveX class of type Windows Media Player
        # ActiveXClass = MakeActiveXClass(VideoControlModule.MediaPlayer, None, self)
        # Create an object of that class
        self.ax = WindowsMediaPlayer.WindowsMediaPlayer(self)
        
        # Cheat and bring constants from WindowsMediaPlayer into the object.
        # Constants can even be made a object of video_msw but making them 
        # part of self.ax makes for better understanding when reading code
        self.ax.constants = WindowsMediaPlayer.constants()
        
        # Set some properties of the WindowsMediaPlayer object.
        # Make sure AutoStart is False.
        self.ax.autostart = False
        # Hide the "display" portion, which is below the Control Buttons.
        self.ax.showdisplay = False 
        
        ######## Define Event Handlers for critical Media Player Events ########
        #
        # We handle many of these events to make sure that the rest of Transana
        # knows what the video player is doing. This is especially important 
        # when we use the video buttons to control playback.
        
        # Get feedback regarding a change in Media Position. Events get fired
        # when moving tracker etc.
        self.Bind(WindowsMediaPlayer.EVT_PositionChange, self.OnPositionChange, self.ax)
        
        # Get feedback regarding a change in Media Ready State. Events get
        # fired when filenames change etc.
        self.Bind(WindowsMediaPlayer.EVT_ReadyStateChange, self.OnReadyStateChange, self.ax)
        
        # TODO:  Implement this on Mac.
        # It turns out that detecting this state change is the only way to tell 
        # when the Play button has been hit in the video window, so this is 
        # necessary for Presentation Mode changes and also to start/stop the 
        # cursor in the Visualization Window#
        # Get feedback regarding a change in Media Play State. Events get fired
        # when the play, pause, stop buttons are pressed.        
        self.Bind(WindowsMediaPlayer.EVT_PlayStateChange, self.OnPlayStateChange, self.ax)
        
        # Now we are ready to play etc.
        
        # TODO: This might not be necessary since our move to WMP from ActiveMovie
        # However, in Windows, you can't just start video playback.  
        # You have to confirm that the video file is loaded first.
        # This is best handled using a Timer.  This timer is used to play 
        # video when it is ready.
        self.ReadyToPlay = wx.Timer(self, CONTROL_READYTOPLAY)
        wx.EVT_TIMER(self, CONTROL_READYTOPLAY, self.OnReadyToPlay)

        # In Windows, you can't just specify a video filename and a position.
        # Specifying the filename causes the ActiveX Control to load the video, 
        # which takes some time, and you can't specify the starting position
        # until that process is done. You can only tell when the video is done 
        # loading by polling the media control's ReadyState until you get the 
        # result you need (amvComplete).  This needs to be handled by a wxTimer,
        # defined here
        self.PositionAfterLoading = wx.Timer(self, CONTROL_POSITIONAFTERLOADING)
        wx.EVT_TIMER(self, CONTROL_POSITIONAFTERLOADING, self.OnPositionAfterLoading)
        
        # Since we changed code from ActiveMovie to WindowsMediaPlayer, the OnTimer 
        # event of ActiveMovie is not fired, meaning we don't get automatic 
        # notification of playback progress. To handle this create a timer which
        # checks CurrentPosition every 10 secs if the video is playing and 
        # post CurrentPosition to VideoWindow 
        self.ProgressNotification = wx.Timer(self, CONTROL_PROGRESSNOTIFICATION)
        wx.EVT_TIMER(self, CONTROL_PROGRESSNOTIFICATION, self.OnProgressNotification)
        
        # TODO: Rewrite
        # The following lines are a workaround. Creating a WindowsMediaPlayer
        # control in a wxDialog doesn't seem to make the control visible until 
        # a file is loaded. So what we do here is load a non-existing file. This
        # causes an exception, but does make the control visible. We then have 
        # to reset the size of the control to the parameters passed in from 
        # VideoWindow. If you find a cleaner way of doing this, feel free to get
        # rid of these lines.
        self.ax.filename = 'images\\splash.gif'
        
        # Set member Rate to 1.0. Use this later right before play.
        self.Rate = 1.0
        
        # Handle the wxPython resize event, and ensure that size changes
        # propagate to other windows.
        wx.EVT_SIZE(self, self.OnSize)
                
        # Handle the wxPython close window event, and ensure that all timers are
        # stopped and the ActiveX Control is destroyed before closing
        wx.EVT_CLOSE(self, self.OnCloseWindow)

    def SetFilename(self, filename):
        """Set the name of the movie file to use."""
        # self.ax.FileName is now synchronous meaning, that this function will
        # not return until the video is loaded. 
        # TODO: Test this and if true remove the ReadyToPlay and 
        # PositionAfterLoading timers
        try:
            self.ax.filename = filename

            # After loading a different video reset the start and end points. Do 
            # this only if a video was loaded, not when FileName was set to "" to 
            # simply unload a video.
            if self.ax.filename != "":
                self.ax.selectionstart = 0.0
                self.ax.selectionend = self.ax.duration
                self.ax.currentposition = 0.0
                self.ax.Stop()
            else:
                self.ax.filename = 'images\\splash.gif'
                self.ax.currentposition = 0.0
                wx.CallAfter(self.ax.Stop)

        except TypeError:
            # Create a string of legal characters for the file names
            allowedChars = TransanaConstants.legalFilenameCharacters
            # check each character in the file name string
            for char in filename:
                # If the character is illegal ...
                if allowedChars.find(char) == -1:
                    msg = _('There is an unsupported character in the Media File Name.\n\n"%s" includes the "%s" character, \nwhich Transana does not support at this time.  Please rename your folders \nand files so that they do not include characters that are not part of English.')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8')
                    dlg = Dialogs.ErrorDialog(self, msg % (filename, char))
                    dlg.ShowModal()
                    dlg.Destroy()
                    break

    def GetFilename(self):
        """Return the name of the currently loaded movie file."""
        return self.ax.filename

    def SetVideoStartPoint(self, TimeCode):
        """Set the start position in ms."""
        # If new SelectionStart is after old SelectionEnd, then we are bound to 
        # get a new SelectionEnd too. (Either some time after new SelectionStart, 
        # if a selection is made or the end of the video (ax.Duration) if a 
        # marker is simply placed. So it's ok to set end point to Duration. 
        # If we don't set end point to after start point we will get 
        # "SelectionStart less than SelectionEnd" error. 
        # Finally, as usual setting anything for a video is possible only 
        # if the video exists, i.e. ax.FileName != ""
        if self.ax.filename != "":
            if TimeCode/1000.0 >= self.ax.selectionend:
                self.ax.selectionend = self.ax.duration
            self.ax.selectionstart = TimeCode/1000.0
        
            # Start the PositionAfterLoading timer, which will detect when the Video
            # loading is complete and call the Positioning routine necessary.
            # Query every 0.1 seconds until the video is loaded
            self.PositionAfterLoading.Start(100)
        
    def GetVideoStartPoint(self):
        """ Gets the start position in ms."""
        return self.ax.selectionstart * 1000.0

    def SetVideoEndPoint(self, TimeCode):
        """Set the end position in ms."""
        # As usual setting anything for a video is possible only 
        # if the video exists, i.e. ax.FileName != ""
        if self.ax.filename != "":
            if TimeCode == -1:
                if type(self.parentVideoWindow.ControlObject.currentObj).__name__ == 'Episode':
                    self.ax.selectionend = self.ax.duration
                elif type(self.parentVideoWindow.ControlObject.currentObj).__name__ == 'Clip':
                    self.ax.selectionend = self.parentVideoWindow.ControlObject.currentObj.clip_stop/1000.0
            else:
                self.ax.selectionend = TimeCode/1000.0
                
    def GetTimecode(self):
        """Return the current position in ms."""
        return int(self.ax.currentposition * 1000.0)
        
    def SetCurrentVideoPosition(self, TimeCode):
        """ Set the current video position. """
        # If the video is supposed to move to before the start, we need to
        # reset the selectionstart.  Ctrl-S doesn't work right without this!
        if self.ax.selectionstart > (TimeCode/1000.0):
            self.ax.selectionstart = TimeCode/1000.0

        self.ax.currentposition = TimeCode/1000.0

    def GetPlayBackSpeed(self):
        return self.Rate

    def SetPlayBackSpeed(self, playBackSpeed):
        """ Sets the play back speed. Divide by 10 to get correct units 
        for Windows Media Player"""
        self.Rate = playBackSpeed/10.0
        # Immediate rate changes are allowed only if a video is loaded
        if self.ax.filename != "":
            # Change the rate immediately so that rate changes when the video 
            # is already playing occur.
            self.ax.rate = self.Rate

    def IsLoading(self):
        """ True if Currently loading a video. """
        return self.ax.readystate == self.ax.constants.mpReadyStateLoading

    def IsPlaying(self):
        """True if currently playing."""
        return self.ax.playstate == self.ax.constants.mpPlaying

    def IsPaused(self):
        """ True if currently paused. """
        return self.ax.playstate == self.ax.constants.mpPaused
        
    def IsStopped(self):
        """ True is currently stopped. """
        return self.ax.playstate == self.ax.constants.mpStopped

    def PlaySegment(self, t0, t1):
        """Play the segment of the video from timecode t0 to t1."""
        self.SetVideoStartPoint(t0)
        self.SetVideoEndPoint(t1)
        self.Play()

    def Play(self):
        """Play the video."""
        # Start the ReadyToPlay timer, polling which polls the ActiveX control
        # every 10ms and plays when ready
        self.ReadyToPlay.Start(100)

    def Pause(self):
        """Pause the video. (pause does not reposition the video)"""
        # Stop sending progress notification messages to VideoWindow
        self.ProgressNotification.Stop()
        self.ax.Pause()

    def Stop(self):
        """Stop the video. (stop playback and return position to Video Start Point)."""
        # Stop sending progress notification messages to VideoWindow
        self.ProgressNotification.Stop()
        self.ax.Stop()

    def GetMediaLength(self):
        # Media Player gives time in Seconds in the Duration function
        # Transana wants time in Milliseconds.
        return long(self.ax.duration * 1000.00)

    def ReadyState(self):
        """ Returns the ActiveMovie control's 'ReadyState' information """
        return self.ax.readystate

    # "Private" methods.
    def OnPositionAfterLoading(self, event):
        """ Positions the video once the file is done loading """
        # If the video is finished loading, set the starting position and 
        # stop the PositionAfterLoading timer
        # If the video is not done loading, this loop will fire again 
        # according to the wxTimer (PositionAfterLoading)'s time interval
        
        if self.ReadyState() == self.ax.constants.amvComplete:
            # Stop the timer
            self.PositionAfterLoading.Stop()

            # Position the video to the specified Start Point
            self.SetCurrentVideoPosition(self.ax.selectionstart * 1000.0)

    def OnReadyToPlay(self, event):
        """ Starts Video Playback when the Media Player is ready """
        # To play the video, the player must be ready AND 
        # the positioning timer must be done
        if (self.ReadyState() == self.ax.constants.amvComplete) and \
           (not self.PositionAfterLoading.IsRunning()):
            self.ax.Play()
            # Start sending axm.CurrentPosition to VideoWindow every 10ms
            self.ProgressNotification.Start(100)
            # Stop the timer 
            self.ReadyToPlay.Stop()
            
    def OnProgressNotification(self, event):
        """If playing checks the axm.CurrentPosition property ever 10ms and 
        passes information to VideoWindow"""
        if (self.IsPlaying()):
            # If playing, check to see if the current segment has ended.
            if self.ax.selectionend and (self.ax.currentposition >= self.ax.selectionend):
                self.Stop()
                # If we're NOT in PlayAllClips mode, ...
                if self.parentVideoWindow.ControlObject.PlayAllClipsWindow == None:
                    # ... reset the video to the original start point.  (This will break PlayAllClips, though.)    
                    self.SetVideoStartPoint(self.GetVideoStartPoint())
            self.PostPos()
        else:
            self.ProgressNotification.Stop()

    def PostPos(self):
        # Notify the parent window of the change in video position

        # print "video_msw.PostPos()", self.GetTimecode() , self.parentVideoWindow.ControlObject.VideoEndPoint

        # This was added by DKW for release 2.05.  While it is perhaps not necessary, my hope is that it
        # will catch a bug that's been reported in Play All Clips, where the audio keeps playing instead of
        # the video moving on to the next clip, that I haven't been able to reproduce.

        # If the current position is greater than the desired Video End Point, stop playback.
        # (NOTE that 0 and -1 are used in different places to indicate that no end point has been set.)
        if (self.parentVideoWindow.ControlObject.VideoEndPoint > 0) and (self.GetTimecode() > self.parentVideoWindow.ControlObject.VideoEndPoint):
            self.Stop()
        self.parentVideoWindow.UpdateVideoPosition(self.GetTimecode())

    def OnCloseWindow(self, event):
        # Stop all timers. This ensures for a cleaner exit for some reason
        self.ProgressNotification.Stop()
        self.ReadyToPlay.Stop()
        self.PositionAfterLoading.Stop()
        # Manually destroy the ActiveMove Control.
        # Python does not exit properly if this line is missing
        self.ax.Destroy()
        self.Destroy()

    # Handlers for ActiveX events.
    def OnReadyStateChange(self, event):
        # The caller of this throws away exceptions, which is not
        # what we want.
        originalSize = self.GetSize()
        try:
            # We need this because the control will automatically reposition
            # itself without properly redrawing after it finishes loading
            # a video.
            if event.ReadyState == self.ax.constants.amvComplete:
                # Fit the Media Player in the Window once the video is loaded
                self.Fit()
                # Inform the rest of the application about the video position
                self.PostPos()

                # Take "Auto-Arrange" into account (don't resize if not checked)
                if TransanaGlobal.configData.autoArrange:
                    # Adjust size of video window to size of the movie loaded
                    self.OnSizeChange()
                # I'm not sure why, but the video window screen size creeps up 
                # when video is loaded. This "else" clause is designed to 
                # prevent that.
                else:
                    self.SetSize(originalSize)
                    
        except:
            print "Exception in ActiveX event handler."
            traceback.print_exc()

    def OnPlayStateChange(self, event): # old_state
        """ Detect change in the Play State and pass that information up to the 
        VideoWindow Object. """
        # Changes in the Play State of the Media Player are used to trigger 
        # Display Layout Changes based on the Presentation Mode settings

        # TODO:  Duplicate this for the Mac.  This feedback is necessary for 
        # implementing Presentation Mode.
        
        # Different Media Players use different constants.  Here, we need to 
        # translate Windows Media Player's States into Transana PlayStates so 
        # they can be properly interpreted by the ControlObject.
        if event.NewState == self.ax.constants.mpStopped:
             play_state = MEDIA_PLAYSTATE_STOP
             # on stop, reset the video to the original StartPoint
             self.SetVideoStartPoint(self.ax.selectionstart * 1000.0)
        elif event.NewState == self.ax.constants.mpPaused:
             play_state = MEDIA_PLAYSTATE_PAUSE
        elif event.NewState == self.ax.constants.mpPlaying:
             # Change the rate here so that rate changes occur when the 
             # video is stopped, paused etc.
             self.ax.rate = self.Rate
             # Start the Progress Notification timer and send             
             self.ProgressNotification.Start(100)
             play_state = MEDIA_PLAYSTATE_PLAY
        else:
             play_state = MEDIA_PLAYSTATE_NONE
        # Pass the Play State information up to the VideoWindow Object
        self.parentVideoWindow.UpdatePlayState(play_state)

        
    def OnPositionChange(self, event):
        self.PostPos()

    def OnSizeChange(self):
        """ Adjust the size of the video window based on the size of the 
        original video and the Options > Video Size setting. """
        #  Determine the screen size 
        rect = wx.ClientDisplayRect()
        #  Get the current position of the Video Window
        pos = self.GetPosition()
        #  Establish the minimum width of the media player control 
        # (if media is audio-only, for example)
        minWidth = max(self.ax.imagesourcewidth, 300)
        # use Movie Height
        minHeight = self.ax.imagesourceheight
        # Adjust Video Size, unless you are showing the Splash Screen
        if self.ax.filename == 'images\\splash.gif':
            sizeAdjust = 100
        else:
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
        #    height:  use height of Movie  + 45 + headerHeight pixels for the player controls
        # rect[0] compensates if the Start Menu is on the left side of the screen
        self.SetDimensions(rect[0] + rect[2] - minWidth - 3, 
                           pos[1], 
                           minWidth, 
                           minHeight + 45 + headerHeight)
        self.parentVideoWindow.UpdateVideoWindowPosition(rect[0] + rect[2] - minWidth - 3, 
                                                         pos[1], 
                                                         minWidth, 
                                                         minHeight + 45 + headerHeight)

    def OnSize(self, event):
        """ Handler for the OnSize event. Set the dimensions of the ActiveX
        control depending on the size of the dialog, and also send messages to 
        the other windows to resize if Auto-Arrange is enabled. """
        
        # First resize the video to fit the new window dimensions
        # Determine size of the wxDialog
        (width, height) = self.GetSize()
        # We need to know the height of the Window Header to adjust the size of the Graphic Area
        headerHeight = self.GetSizeTuple()[1] - self.GetClientSizeTuple()[1]
        # Set the dimensions of the ActiveX control based on size of wxDialog
        # These magic numbers (8, 30) account for the border pixels and the
        # size of the playback controls
        self.ax.SetDimensions(0,0,width-8,height-headerHeight-4)
        
        # Get the current position of the Video Window after resizing
        pos = self.GetPosition()
        # Pass position onto other windows
        self.parentVideoWindow.UpdateVideoWindowPosition(pos[0], pos[1],
                                                         width, height)
        # There seems to be no need to pass the event onto axw OnSize handler
        #self.ax.axw_OnSize(event)

if __name__ == '__main__':
    MENU_FILE_NEW        =  101
    MENU_FILE_OPEN       =  102
    MENU_FILE_EXIT       =  103

    ID_BUTTON_RESIZE      = 1001
    ID_BUTTON_PLAY        = 1002
    ID_BUTTON_PAUSE       = 1003
    ID_BUTTON_STOP        = 1004
    ID_BUTTON_PLAYSEGMENT = 1005

    class ControlWindow(wx.Frame):
        def __init__(self, VideoWindow, pos=(100, 600), size=wx.Size(500, 100)):
            wx.Frame.__init__(self, None, -1, "Video Controller", 
                                pos=pos, size=size,
                                style = wx.DEFAULT_FRAME_STYLE )
            self.SetBackgroundColour(wx.WHITE)

            self.VideoWindow = VideoWindow
        
            menuBar = wx.MenuBar()
            fileMenu = wx.Menu()
            fileMenu.Append(MENU_FILE_NEW, "&New")
            fileMenu.Append(MENU_FILE_OPEN, "&Open")
            fileMenu.Append(MENU_FILE_EXIT, "E&xit")
            menuBar.Append(fileMenu, "&File")

            self.SetMenuBar(menuBar)

            wx.EVT_MENU(self, MENU_FILE_NEW, self.OnFileNew)
            wx.EVT_MENU(self, MENU_FILE_OPEN, self.OnFileOpen)
            wx.EVT_MENU(self, MENU_FILE_EXIT, self.OnFileExit)

            self.ButtonResize = wx.Button(self, ID_BUTTON_RESIZE, "Resize", wx.Point(5, 5))
            wx.EVT_BUTTON(self, ID_BUTTON_RESIZE, self.OnResize)
            self.ButtonPlay = wx.Button(self, ID_BUTTON_PLAY, "Play", wx.Point(85, 5))
            wx.EVT_BUTTON(self, ID_BUTTON_PLAY, self.OnPlay)
            self.ButtonPause = wx.Button(self, ID_BUTTON_PAUSE, "Pause", wx.Point(165, 5))
            wx.EVT_BUTTON(self, ID_BUTTON_PAUSE, self.OnPause)
            self.ButtonStop = wx.Button(self, ID_BUTTON_STOP, "Stop", wx.Point(245, 5))
            wx.EVT_BUTTON(self, ID_BUTTON_STOP, self.OnStop)
            self.ButtonPlaySeg = wx.Button(self, ID_BUTTON_PLAYSEGMENT, "Play Segment", wx.Point(325, 5))
            wx.EVT_BUTTON(self, ID_BUTTON_PLAYSEGMENT, self.OnPlaySegment)

            wx.EVT_CLOSE(self, self.OnCloseWindow)

            self.CreateStatusBar()
            self.SetStatusText('Program opened.')
            self.Show(True)


        def OnCloseWindow(self, event):
            self.VideoWindow.Close()
            self.ProgressNotication.Stop()
            self.ReadyToPlay.Stop()
            self.PositionAfterLoading.Stop()
            self.ax.Destroy()
            self.Destroy()

        def OnFileNew(self, event):
            self.SetStatusText('New.')
            self.VideoWindow.SetFilename('')

        def OnFileOpen(self, event):
            self.SetStatusText('Opening Video File.')
            wildcard = fileTypesString #From TransanaConstants
            dlg = wx.FileDialog(self, "Choose a video", "v:/Demo", "demo.mpg", wildcard, wx.OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode('Opening %s.', 'utf8')
                else:
                    prompt = 'Opening %s.'
                self.SetStatusText(prompt % dlg.GetPath())
                self.VideoWindow.SetFilename(dlg.GetPath())
                self.VideoWindow.SetVideoStartPoint(0)

        def OnFileExit(self, event):
            self.SetStatusText('Program closing.')
            if self.VideoWindow != None:
                self.VideoWindow.Close()
                self.VideoWindow = None
            self.Close()

        def OnResize(self, event):
            self.VideoWindow.SetSize(wx.Size(500, 400))
            self.SetStatusText('Window resized to (500, 400).')

        def OnPlay(self, event):
            self.VideoWindow.Play()
            self.SetStatusText('Play.')
        
        def OnPause(self, event):
            self.VideoWindow.Pause()
            self.SetStatusText('Pause.')

        def OnStop(self, event):
            self.VideoWindow.Stop()
            self.SetStatusText('Stop.')

        def OnPlaySegment(self, event):
            self.SetStatusText('Play Segment (1:00 to 1:15).')
            if self.VideoWindow.GetFilename() == '':
                self.VideoWindow.SetFilename('v:/Demo/Demo.mpg')
            # set the video start and end points to the start and stop points defined in the clip
            self.VideoWindow.PlaySegment(60000, 75000)

    class MyApp(wx.App):
        def OnInit(self):
            self.frame = VideoFrame(parent = None, parentVideoWindow = None)
            self.frame.Show(True)
            controller = ControlWindow(self.frame)
            self.SetTopWindow(controller)
            return True

        def UpdateVideoPosition(self, timecode):
            pass

        def UpdateVideoWindowPosition(self, left, top, width, height):
            pass


    app = MyApp(0)
    app.MainLoop()
