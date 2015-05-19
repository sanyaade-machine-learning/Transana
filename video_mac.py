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

""" This unit handles interaction with the QuickTime Control used in displaying video
    on the Macintosh OS-X platform.  It uses the wxQtMovie control written by 
    Kevin Ollivier.  Thanks Kevin! """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Mark Kim, Rajas Sambhare'

import wx
try:
    import wx.qtmovie    # import the qxQtMovie custom control
except:
    raise ImportError("Can't load Quicktime.  wxQtMovie component not found.")
import sys
import time
import traceback

if __name__ == '__main__':
    __builtins__._ = wx.GetTranslation

from TransanaConstants import *
import TransanaGlobal

        
class VideoFrame(wx.Frame):
        
    """Video player frame. Use the 'public' methods to control the player.

    This frame posts UpdateVideoPosEvent events to advise other
    wxWindow objects of its position as it plays. Use the Advise and
    Unadvise methods (of Observable) to control this."""
        
    def __init__(self, parentVideoWindow = None, pos=wx.DefaultPosition, size=wx.DefaultSize):
        self.size = size
        wx.Frame.__init__(self, None, -1, _("Video"), pos=pos, size=size,
                         style = wx.RESIZE_BORDER | wx.CAPTION )

        # To look right, the Mac needs the Small Window Variant.
        self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        self.SetBackgroundColour(wx.BLACK)
        self.parentVideoWindow = parentVideoWindow
        
        # Bind event handlers
        wx.EVT_LEFT_DOWN(self, self.MouseLeftDown)
        wx.EVT_SIZE(self, self.OnSize)
        self.done = 0
        
        # Initialize QuickTime
        self.movie = wx.qtmovie.wxQtMovie(self, -1, "")
        self.movie.LoadMovie("images/splash.gif")

        # Define playState and Rate
        self.playState = self.movie.GetState()   # wx.qtmovie.wxMEDIA_STATE_STOPPED
        self.Rate = 1.0

        self.VideoStartPoint = 0
        
        # If this value is nonzero, playback will pause at this timecode.
        self.VideoEndPoint = 0
        
        # Define filename
        self.FileName = ""
        
        # Define timer for progress notification to other windows
        self.ProgressNotification = wx.Timer(self, CONTROL_PROGRESSNOTIFICATION)
        wx.EVT_TIMER(self, CONTROL_PROGRESSNOTIFICATION, self.OnProgressNotification)
        self.ProgressNotification.Start(50)
	

    # Pass mouse click down to QuickTime
    def MouseLeftDown(self, event):
        if self.movie.GetState() == movie.wxMEDIA_STATE_PLAYING:
            self.movie.Pause()
        else:
            self.movie.Play()

    def SetFilename(self, filename):
        if filename == "":
            filename = 'images/splash.gif'

        self.FileName = filename
        self.movie.LoadMovie(filename)

        # Set the number of units per second to 1000
        self.movie.SetTimeUnit(1000)

        # Okay, we want to update the Window Size to mathc the video UNLESS we're in PlayAllClips mode. 
#        if self.parentVideoWindow.ControlObject.PlayAllClipsWindow == None:
#            bounds = self.movie.GetDefaultMovieSize()
#            self.SetSize(wx.Size(bounds.x, bounds.y))

        #initialize the start and end points            
        self.SetVideoStartPoint(0)

        self.SetVideoEndPoint(self.movie.GetDuration())

        # Force an OnSizeChange event to align things correctly
        self.OnSizeChange()


    def GetFilename(self):
        return self.FileName

    def SetVideoStartPoint(self, TimeCode):
        """Set the current position in ms."""
        self.VideoStartPoint = TimeCode
        #move the slider to the videostartpoint        
        self.movie.SetCurrentTime(long(TimeCode))
            
    def GetVideoStartPoint(self):
        return self.VideoStartPoint
        
    def SetVideoEndPoint(self, TimeCode):
        self.VideoEndPoint = TimeCode

    def GetTimecode(self):
        """Return the current position in ms."""
        return self.movie.GetCurrentTime()

    def SetCurrentVideoPosition(self, TimeCode):
        """ Set the current video position. """
        self.movie.SetCurrentTime(TimeCode)
    
    def SetPlayBackSpeed(self, playBackSpeed):
        """ Sets the play back speed. Divide by 10 to get correct units for Quicktime"""
        # Setting Playback Speed on the Mac causes the Quicktime Player to Play.  That's not what we want.
        originalPlayState = self.movie.GetState()
        self.Rate = playBackSpeed/10.0
        self.movie.SetPlaySpeed(self.Rate)
        if originalPlayState == wx.qtmovie.wxMEDIA_STATE_PAUSED:
            self.Pause()
        elif originalPlayState == wx.qtmovie.wxMEDIA_STATE_STOPPED:
            self.Stop()

    def IsPlaying(self):
        """True if currently playing."""
        return self.movie.GetState() == wx.qtmovie.wxMEDIA_STATE_PLAYING

    def IsPaused(self):
        """ True if currently paused. """
        return self.movie.GetState() == wx.qtmovie.wxMEDIA_STATE_PAUSED
        
    def IsStopped(self):
        """ True if currently stopped. """
        return self.movie.GetState() == wx.qtmovie.wxMEDIA_STATE_STOPPED
        
    def IsLoading(self):
        """ True if currently loading. """
        return False
    
    def PlaySegment(self, t0, t1):
        """Play the segment of the video from timecode t0 to t1."""
        self.Stop()
        self.SetVideoStartPoint(t0)
        self.SetVideoEndPoint(t1)
        self.Play()
            
    def Play(self):
        """Play the video."""
        self.parentVideoWindow.UpdatePlayState(wx.qtmovie.wxMEDIA_STATE_PLAYING)
        wx.CallAfter(self.movie.Play)

    def Pause(self):
        """Pause the video. (pause does not reposition the video)"""
        self.movie.Pause()
        self.parentVideoWindow.UpdatePlayState(wx.qtmovie.wxMEDIA_STATE_STOPPED)
        
    def Stop(self):
        """Stop the video. (stop playback and return position to Video Start Point)."""
        self.movie.Stop()
        self.parentVideoWindow.UpdatePlayState(wx.qtmovie.wxMEDIA_STATE_STOPPED)
        
    def GetMediaLength(self):
        return self.movie.GetDuration()
	
    def ReadyState(self):
        """ Returns the Movie control's 'ReadyState' information """
        return self.movie.GetLoadState()

    def OnProgressNotification(self, event):
        # Detect Play State Change and notify the VideoWindow.
        # See if the playState has changed.
        if self.playState != self.movie.GetState():
            # If it has, communicate that to the VideoWindow.  The Mac doesn't seem to differentiate between
            # Stop and Pause the say Windows does, so pass "Stopped" for either Stop or Pause.
            if self.movie.GetState() != wx.qtmovie.wxMEDIA_STATE_PLAYING:
                self.parentVideoWindow.UpdatePlayState(wx.qtmovie.wxMEDIA_STATE_STOPPED)
            # Pass "Play" for play.
            else:
                self.parentVideoWindow.UpdatePlayState(wx.qtmovie.wxMEDIA_STATE_PLAYING)
            # Update the local playState variable
            self.playState = self.movie.GetState()

        # If playing, check to see if the current segment has ended.
        if (self.IsPlaying()):
            if self.VideoEndPoint and (self.GetTimecode() >= self.VideoEndPoint):
                self.Stop()
                # If we're NOT in PlayAllClips mode, ...
                if self.parentVideoWindow.ControlObject.PlayAllClipsWindow == None:
                    # ... reset the video to the original start point.  (This will break PlayAllClips, though.)    
                    self.SetVideoStartPoint(self.VideoStartPoint)
                    
            self.PostPos()

    def PostPos(self):
        self.parentVideoWindow.UpdateVideoPosition(self.GetTimecode())

    def OnPositionChange(self, oldPosition, newPosition):
        self.PostPos()

    def OnSizeChange(self):
        """ Adjust the size of the video window based on the size of the original video and the
            Options > Video Size setting. """
        # If Auto-Arrange is un-checked, we ignot this whole routine!
        if TransanaGlobal.configData.autoArrange:
            #  Determine the screen size 
            rect = wx.ClientDisplayRect()
            #  Get the Menu Height, which is where the top of the window should be.
            top = 24

            bounds = self.movie.GetDefaultMovieSize()
    
            #  Establish the minimum width of the media player control (if media is audio-only, for example)
            minWidth = max(bounds.x, 350)

            # use Movie Height
            minHeight = max(bounds.y, 14)

            if self.FileName == 'images/splash.gif':
                sizeAdjust = 100
            else:
                sizeAdjust = TransanaGlobal.configData.videoSize

            # Take Video Size Menu spec into account (50%, 66%, 100%, 150%, 200%)
            minWidth = int(minWidth * sizeAdjust / 100.0)
            minHeight = int(minHeight * sizeAdjust / 100.0)

            #  Set the dimensions of the video window as follows:
            #    left:    right-justify to the screen (left side of screen - media player width - 3 pixel margin)
            #    top:     leave top position unchanged
            #    width:   use minimum media player width established above
            #    height:  use height of Movie  + 10 pixels for the player controls
            self.SetDimensions(rect[2] - minWidth - 3, top, minWidth, minHeight)
            self.parentVideoWindow.UpdateVideoWindowPosition(rect[2] - minWidth - 3, top, minWidth, minHeight + 38)
        else:
            # The video needs to be adjusted to the size of the window.  This tries to do that by resizing the window slightly.
            pos = self.GetPositionTuple()
            size = self.GetSize()
            self.SetDimensions(pos[0], pos[1], size[0]-1, size[1]-1) 
            self.SetDimensions(pos[0], pos[1], size[0], size[1]) 

    def OnSize(self, event):
        """ Handler for OnSize"""
        # Call the wxQtMovie OnSize handler
        event.Skip()
	
        (width, height) = self.GetSize()
        pos = self.GetPosition()
	
        self.parentVideoWindow.UpdateVideoWindowPosition(pos[0], pos[1], width, height)


if __name__ == '__main__':
	
    import TransanaConstants
	
    MENU_FILE_NEW        =  wx.NewId()
    MENU_FILE_OPEN       =  wx.NewId()
    MENU_FILE_EXIT       =  wx.NewId()

    ID_BUTTON_RESIZE      = wx.NewId()
    ID_BUTTON_PLAY        = wx.NewId()
    ID_BUTTON_PAUSE       = wx.NewId()
    ID_BUTTON_STOP        = wx.NewId()
    ID_BUTTON_PLAYSEGMENT = wx.NewId()

    ID_TIMER              = wx.NewId()

    class ControlWindow(wx.Frame):
        def __init__(self, VideoWindow, pos=(100, 600), size=wx.Size(500, 100)):
            wx.Frame.__init__(self, None, -1, "Video Controller", pos=pos, size=size,
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

            self.ButtonPlay = wx.Button(self, ID_BUTTON_PLAY, "Play", wx.Point(5, 5))
            wx.EVT_BUTTON(self, ID_BUTTON_PLAY, self.OnPlay)
            self.ButtonPause = wx.Button(self, ID_BUTTON_PAUSE, "Pause", wx.Point(85, 5))
            wx.EVT_BUTTON(self, ID_BUTTON_PAUSE, self.OnPause)
            self.ButtonStop = wx.Button(self, ID_BUTTON_STOP, "Stop", wx.Point(165, 5))
            wx.EVT_BUTTON(self, ID_BUTTON_STOP, self.OnStop)
#            self.ButtonPlaySeg = wx.Button(self, ID_BUTTON_PLAYSEGMENT, "Play Segment", wx.Point(245, 5))
#            wx.EVT_BUTTON(self, ID_BUTTON_PLAYSEGMENT, self.OnPlaySegment)
#            self.ButtonResize = wx.Button(self, ID_BUTTON_RESIZE, "Resize", wx.Point(3255, 5))
#            wx.EVT_BUTTON(self, ID_BUTTON_RESIZE, self.OnResize)

            self.CreateStatusBar()
            self.SetStatusText('Program opened.')
            self.Show(True)


        def SetVideoWindow(self, videoWindow):
            self.VideoWindow = videoWindow

        def OnFileNew(self, event):
            self.SetStatusText('New.')
            self.VideoWindow.SetFilename('')

        def OnFileOpen(self, event):
            self.SetStatusText('Opening Video File.')
            
            fss = wx.FileSelector(_("Select a media file"),
                        '', #  os.path.dirname(self.obj.media_filename),
                        '', #  os.path.basename(self.obj.media_filename),
                        "", 
                        _(TransanaConstants.fileTypesString), 
                        wx.OPEN | wx.FILE_MUST_EXIST)
	    # If user didn't cancel ..
	    if fss != "":
		self.VideoWindow.SetFilename(fss)
                
        def OnFileExit(self, event):
            self.SetStatusText('Program closing.')
            self.VideoWindow.Close()
            self.VideoWindow = None
            self.Close()

#        def OnResize(self, event):
#            self.VideoWindow.SetSize(wx.Size(500, 400))
#            self.SetStatusText('Window resized to (500, 400).')
            
        def OnPlay(self, event):
            self.SetStatusText('Play.')
            self.VideoWindow.Play()
            
        def OnPause(self, event):
            self.VideoWindow.Pause()
            
        def OnStop(self, event):
            self.VideoWindow.Stop()
            self.SetStatusText('Stop.')

#        def OnPlaySegment(self, event):
#            self.VideoWindow.PlaySegment(4000, 19000)

        def UpdateVideoPosition(self, pos):
            print 'VideoPosition = ', pos

    class MyApp(wx.App):
        def OnInit(self):
            controller = ControlWindow(None)
            self.frame = VideoFrame(controller)
            self.frame.Show(True)
            controller.SetVideoWindow(self.frame)
            self.SetTopWindow(controller)
            return True
            
        def MacMainLoop(self):
            while(self.frame.done == 0):
                self.Dispatch()


    app = MyApp(0)
    app.MacMainLoop()
