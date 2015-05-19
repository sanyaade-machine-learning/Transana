# Copyright (C) 2006 The Board of Regents of the University of Wisconsin System 
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

UPDATE_PROGRESS_INTERVAL = 150

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

# If running as a stand-alone ...
if __name__ == '__main__':
    # ... we need GetTranslations.  Otherwise, it's handled by Transana's infrastructure
    __builtins__._ = wx.GetTranslation
    # We also need to fake some configuration data usually provided by Transana
    class ConfigData(object):
        def __init__(self):
            self.videoSize = 100
    class TransanaGlobalObject(object):
        def __init__(self):
            self.programDir = ''
            self.configData = ConfigData()
    TransanaGlobal = TransanaGlobalObject()
else:
    # import the Transana Global data
    import TransanaGlobal

CONTROL_PROGRESSNOTIFICATION = wx.NewId()

# Declare the main VideoFrame class, designed to interact with the rest of Transana
class VideoFrame(wx.Dialog):
    """ Video player frame.  This is based on wxMediaCtrl, and should allow increased support for video formats over
        previous media control components. """
        
    def __init__(self, parent = None, parentVideoWindow = None, pos=(30, 30), size=wx.Size(480, 420)):
        """ Initialize the Video Frame object """
        # Create a Dialog to hold the Video Window
        wx.Dialog.__init__(self, parent, -1, _("Video"), pos=pos, size=size, style = wx.RESIZE_BORDER | wx.CAPTION )
        # Remember the parent VideoWindow so we can interact with its methods
        self.parentVideoWindow = parentVideoWindow
        # Initialize the Play State
        self.playState = None
        # Initialize the Loading State variable
        self.isLoading = False
        # We need a flag that indicates that we want to Play as soon as the Load is complete
        self.playWhenLoaded = False
        # The Mac requires the following so that versions look more similar across platforms
        if 'wxMac' in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)
        # Set the background color to Black.  
        self.SetBackgroundColour(wx.BLACK)

        imgFileName = os.path.join(TransanaGlobal.programDir, 'images', 'splash.gif')
        self.splashImage = wx.Image(imgFileName)
        (width, height) = self.GetClientSize()
        self.splashImage.Rescale(width, height)
        self.graphic = wx.StaticBitmap(self, -1, self.splashImage.ConvertToBitmap())

        # Define timer for progress notification to other windows
        self.ProgressNotification = wx.Timer(self, CONTROL_PROGRESSNOTIFICATION)
        wx.EVT_TIMER(self, CONTROL_PROGRESSNOTIFICATION, self.OnProgressNotification)

        # set default back end to QuickTime, as then either type of media can be opened.  If we
        # set it to DirectShow for Windows, we can't later load Quicktime files
        self.backend = wx.media.MEDIABACKEND_QUICKTIME
            
        self.movie = None

        self.CreateMediaPlayer()
        self.movie.Show(False)

        self.Bind(wx.media.EVT_MEDIA_LOADED, self.OnMediaLoaded)
        self.Bind(wx.media.EVT_MEDIA_STOP, self.OnMediaStop)
        self.Bind(wx.EVT_SIZE, self.OnSize)

        box = wx.BoxSizer(wx.HORIZONTAL)
        box.Add(self.movie, 1, wx.EXPAND | wx.ALL, 0)
        
        # Bind event handlers
        wx.EVT_CLOSE(self, self.OnCloseWindow)
        self.done = 0
        
        self.VideoStartPoint = 0
        self.VideoEndPoint = 0
        self.FileName = ""
        self.Rate = 1.0

        self.SetSizer(box)
        self.SetAutoLayout(True)
        self.Layout()

    def CreateMediaPlayer(self, flNm=""):
        self.ProgressNotification.Stop()
        if self.movie:
            self.movie.Destroy()
            self.movie = None
            wx.Yield()
            
        self.movie = wx.media.MediaCtrl(self, id=-1, fileName=flNm, szBackend=self.backend)
        box = wx.BoxSizer(wx.HORIZONTAL)
        box.Add(self.movie, 1, wx.EXPAND | wx.ALL, 0)
        self.SetSizer(box)
        self.SetAutoLayout(True)
        self.Layout()
        # There's no way to detect when the Media Player's control buttons are pressed!
        # Therefore, we ALWAYS need the ProgressNotification loop running.
        self.ProgressNotification.Start(UPDATE_PROGRESS_INTERVAL)

    def SetFilename(self, filename):
        if filename != '':
            self.graphic.Show(False)
            self.movie.Show(True)
            if ('wxMSW' in wx.PlatformInfo):
                (videoFilename, videoExtension) = os.path.splitext(filename)
                if videoExtension in ['.mov']:
                    backendNeeded = wx.media.MEDIABACKEND_QUICKTIME
                else:
                    backendNeeded = wx.media.MEDIABACKEND_DIRECTSHOW
                if self.backend != backendNeeded:
                    self.backend = backendNeeded
                    self.CreateMediaPlayer(filename)
            # We need to have a flag that indicates that the video is in the process of loading
            self.isLoading = True
            if self.movie.Load(filename):
                self.FileName = filename
                self.VideoStartPoint = 0
                self.VideoEndPoint = 0
                self.GetSizer().Layout()
            else:
                print "Unable to load file.  Unsupported format?"
                self.isLoading = False
                self.FileName = ""
            self.Update()
        else:
            self.graphic.Show(True)
            self.movie.Load('')
            self.movie.Show(False)
            # If the graphic has been hidden, it may be the wrong size.  This fixes that.
            self.OnSize(None)
            self.Update()
    
    def GetFilename(self):
        return self.FileName

    def SetVideoStartPoint(self, TimeCode):
        """Set the current position in ms."""
        # TimeCode must be an int on the Mac.
        if not isinstance(TimeCode, int):
            TimeCode = int(TimeCode)
        self.VideoStartPoint = TimeCode
        self.movie.Seek(TimeCode)
        if self.parentVideoWindow != None:
            # notify the rest of Transana of the change.  However, calling parentVideoWindow.UpdateVideoPosition calls this method, causing
            # recursion problems, so we make the call to the parentVideoWindow's ControlObject.
            self.parentVideoWindow.ControlObject.UpdateVideoPosition(self.VideoStartPoint)
                   
    def GetVideoStartPoint(self):
        return self.VideoStartPoint
        
    def SetVideoEndPoint(self, TimeCode):
        self.VideoEndPoint = TimeCode

    def GetVideoEndPoint(self):
        return self.VideoEndPoint
 
    def GetTimecode(self):
        """Return the current position in ms."""
        return self.movie.Tell()

    def GetMediaLength(self):
        return self.movie.Length()

    def GetState(self):
        return self.movie.GetState()
    
    def SetCurrentVideoPosition(self, TimeCode):
        """ Set the current video position. """
        self.movie.Seek(TimeCode)

    def GetPlayBackSpeed(self):
        return self.Rate
    
    def SetPlayBackSpeed(self, playbackSpeed):
        """ Sets the play back speed."""
        # Reset the Rate value.  The Options Setting screen uses values 1 - 20 to represent 0.1 to 2.0, so divide by 10!
        self.Rate = playbackSpeed / 10.0
        # Set the playback Rate, if we're NOT using the QuickTime Back End.  If we are, that would cause the video to
        # play, so we'll skip it.
        if (self.backend != wx.media.MEDIABACKEND_QUICKTIME):
            self.movie.SetPlaybackRate(self.Rate)

    def IsPlaying(self):
        return self.GetState() == wx.media.MEDIASTATE_PLAYING

    def IsPaused(self):
        return self.GetState() == wx.media.MEDIASTATE_PAUSED

    def IsStopped(self):
        return self.GetState() == wx.media.MEDIASTATE_STOPPED

    def IsLoading(self):
        return self.isLoading

    def PlaySegment(self, t0, t1):
        """Play the segment of the video from timecode t0 to t1."""
        self.Stop()
        self.SetVideoStartPoint(t0)
        self.SetVideoEndPoint(t1)
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
            if (self.backend == wx.media.MEDIABACKEND_QUICKTIME) and (self.Rate != 1.0):
                self.movie.Play()
                self.movie.SetPlaybackRate(self.Rate)
            else:
                self.movie.Play()

    def Pause(self):
        """Pause the video. (pause does not reposition the video)"""
        self.movie.Pause()

    def Stop(self):
        """Stop the video. (stop playback and return position to Video Start Point)."""
        self.movie.Stop()
        self.movie.Seek(self.VideoStartPoint)
        self.PostPos()

    def OnCloseWindow(self, event):
        self.ProgressNotification.Stop()
        self.Destroy()

    def OnMediaLoaded(self, event):
        # Indicate that the loading process is complete
        self.isLoading = False
        # Putting this here allows it to show up under QuickTime!
        self.movie.ShowPlayerControls(wx.media.MEDIACTRLPLAYERCONTROLS_DEFAULT | wx.media.MEDIACTRLPLAYERCONTROLS_VOLUME)
        # once the video is loaded, we can determine its size and should react to that.
        self.OnSizeChange()
        self.Update()
        # Play All Clips is having trouble starting because "Play()" is called before loading is complete.
        # This structure allows the Play() to be called once the load is done.
        if self.playWhenLoaded:
            self.Play()
        elif "wxMSW" in wx.PlatformInfo:
            self.Stop()
        # Often, the Seek has also not occurred correctly.  This detects and fixes that problem.
        if self.VideoStartPoint != self.GetTimecode():
            self.movie.Seek(self.VideoStartPoint)

    def OnSizeChange(self):
        (sizeX, sizeY) = self.movie.GetBestSize()
        # Now that we have a size, let's position the window
        #  Determine the screen size 
        rect = wx.ClientDisplayRect()
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
        if self.backend == wx.media.MEDIABACKEND_DIRECTSHOW:
            controlBarSize = 65
        elif self.backend == wx.media.MEDIABACKEND_QUICKTIME:
            controlBarSize = 10
        else:
            controlBarSize = 100  # This'll be quite noticable!
        self.SetDimensions(rect[0] + rect[2] - minWidth - 3, 
                           pos[1], 
                           minWidth, 
                           minHeight + controlBarSize + headerHeight)
        if self.parentVideoWindow != None:
            self.parentVideoWindow.UpdateVideoWindowPosition(rect[0] + rect[2] - minWidth - 3, 
                                                             pos[1], 
                                                             minWidth, 
                                                             minHeight + controlBarSize + headerHeight)

    def OnSize(self, event):
        """ Process Size Change event and notify the ControlObject """
        # if event is not None (which it can be if this is called from non-event-driven code rather than
        # from a real even, then we should process underlying OnSize events.
        if event != None:
            event.Skip()
        # We can't update the graphic's size if it's hidden, as that causes problems visually on the Mac.
        if self.graphic.IsShown():
            # For resizing the graphic, we need the ClientSize (sans toolbar)
            (width, height) = self.GetClientSize()

            if (width > 0) and (height > 0):
                self.graphic.Destroy()
                imgFileName = os.path.join(TransanaGlobal.programDir, 'images', 'splash.gif')
                self.splashImage = wx.Image(imgFileName)
                self.splashImage.Rescale(width, height)
                self.graphic = wx.StaticBitmap(self, -1, self.splashImage.ConvertToBitmap(), pos=(0, 0), size=(width, height))
                self.graphic.Refresh()
        # for updating window position, we need the ful window size
        (width, height) = self.GetSize()
        pos = self.GetPosition()
        if self.parentVideoWindow != None:
            self.parentVideoWindow.UpdateVideoWindowPosition(pos[0], pos[1], width, height)

    def OnMediaStop(self, event):
        wx.CallAfter(self.movie.Seek, self.VideoStartPoint)
        wx.CallAfter(self.PostPos)

    def OnProgressNotification(self, event):
        # Detect Play State Change and notify the VideoWindow.
        # See if the playState has changed.
        if self.playState != self.movie.GetState():
            # If it has, communicate that to the VideoWindow.  The Mac doesn't seem to differentiate between
            # Stop and Pause the say Windows does, so pass "Stopped" for either Stop or Pause.
            if self.movie.GetState() != wx.media.MEDIASTATE_PLAYING:
                if self.parentVideoWindow != None:
                    self.parentVideoWindow.UpdatePlayState(wx.media.MEDIASTATE_STOPPED)
            # Pass "Play" for play.
            else:
                if self.parentVideoWindow != None:
                    self.parentVideoWindow.UpdatePlayState(wx.media.MEDIASTATE_PLAYING)
            # Update the local playState variable
            self.playState = self.movie.GetState()

        # Take this opportunity to see if there are any waiting events.  Leaving this out has the unfortunate side effect
        # of diabling the media player's control bar.
        # There is sometimes a problem with recursive calls to Yield; trap the exception ...
        try:
            wx.YieldIfNeeded()
        # ... and ignore it!
        except:
            pass

        # The timer that calls this routine runs whether the video is playing or not.  We only need to think
        # about updating the rest of Transana if the video is playing.
        if self.IsPlaying():
            # If playing, check to see if the current segment has ended.
            if (self.VideoEndPoint > 0) and (self.GetTimecode() >= self.VideoEndPoint):
                self.Stop()
                # If we're NOT in PlayAllClips mode, ...
                if (self != None) and (self.parentVideoWindow.ControlObject.PlayAllClipsWindow == None):
                    # ... reset the video to the original start point.  (This will break PlayAllClips, though.)    
                    self.SetVideoStartPoint(self.VideoStartPoint)


            self.PostPos()

    def PostPos(self):
        if self.parentVideoWindow != None:
            self.parentVideoWindow.UpdateVideoPosition(self.GetTimecode())

        
if __name__ == '__main__':
    # Create a Stand-alone video controller
    import time
    
    MENU_FILE_NEW        =  wx.NewId()
    MENU_FILE_OPEN       =  wx.NewId()
    MENU_FILE_EXIT       =  wx.NewId()

    TIMER_INTERVAL = 100

    class ControlWindow(wx.Frame):
        def __init__(self, VideoWindow, pos=(100, 600), size=wx.Size(500, 100)):

            wx.Frame.__init__(self, None, -1, "Video Controller", pos=pos, size=size,
                             style = wx.RESIZE_BORDER | wx.CAPTION )
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

            self.ButtonPlay = wx.Button(self, -1, "Play")
            self.Bind(wx.EVT_BUTTON, self.OnPlay, self.ButtonPlay)
            self.ButtonPause = wx.Button(self, -1, "Pause")
            self.Bind(wx.EVT_BUTTON, self.OnPause, self.ButtonPause)
            self.ButtonStop = wx.Button(self, -1, "Stop")
            self.Bind(wx.EVT_BUTTON, self.OnStop, self.ButtonStop)
            self.ButtonPlaySeg = wx.Button(self, -1, "Play Segment")
            self.Bind(wx.EVT_BUTTON, self.OnPlaySegment, self.ButtonPlaySeg)
            self.ButtonResize = wx.Button(self, -1, "Resize")
            self.Bind(wx.EVT_BUTTON, self.OnResize, self.ButtonResize)
            self.Rate = 1.0
            self.ButtonRate = wx.Button(self, -1, "Rate")
            self.Bind(wx.EVT_BUTTON, self.OnRate, self.ButtonRate)
        
            btnBox = wx.BoxSizer(wx.HORIZONTAL)
            btnBox.Add(self.ButtonPlay, 1, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.EXPAND, 10)
            btnBox.Add(self.ButtonPause, 1, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.EXPAND, 10)
            btnBox.Add(self.ButtonStop, 1, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.EXPAND, 10)
            btnBox.Add(self.ButtonPlaySeg, 1, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.EXPAND, 10)
            btnBox.Add(self.ButtonResize, 1, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.EXPAND, 10)
            btnBox.Add(self.ButtonRate, 1, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.EXPAND, 10)

            box = wx.BoxSizer(wx.VERTICAL)
            box.Add(btnBox, 1, wx.ALIGN_LEFT | wx.EXPAND, 5)

            self.txt = wx.StaticText(self, -1, 'Video Position:  %10d' % 0)
            txtbox = wx.BoxSizer(wx.HORIZONTAL)
            txtbox.Add(self.txt, 1, wx.EXPAND, 10)
            box.Add(txtbox, 1, wx.ALIGN_LEFT, 5)
            

            self.CreateStatusBar()
            self.SetStatusText('Program opened.')

            self.SetSizer(box)
            self.Fit()
            self.SetSizeHints(minW = self.GetSize()[0], minH = self.GetSize()[1], maxH = self.GetSize()[1])
            self.SetAutoLayout(True)
            self.Layout()
            
            self.Show(True)

            # Create a timer to handle waiting for the video to load
            self.Timer = wx.Timer(self)
            self.Bind(wx.EVT_TIMER, self.OnTimer)
            self.Timer.Start(TIMER_INTERVAL)


        def OnFileNew(self, event):
            self.SetStatusText('New.')
            self.VideoWindow.SetFilename('')

        def OnFileOpen(self, event):
            self.SetStatusText('Opening Video File.')
            dlg = wx.FileDialog(self, message="Select Video File:", defaultDir="E:\Video", defaultFile="", style=wx.OPEN | wx.CHANGE_DIR)
            if dlg.ShowModal() == wx.ID_OK:
                self.VideoWindow.SetFilename(dlg.GetPath())
                self.Rate = 1.0
            dlg.Destroy()
                
        def OnFileExit(self, event):
            self.Timer.Stop()
            self.SetStatusText('Program closing.')
            self.VideoWindow.Close()
            self.VideoWindow = None
            self.Close()

        def OnResize(self, event):
            self.VideoWindow.SetSize(wx.Size(500, 400))
            self.SetStatusText('Window resized to (500, 400).')
            
        def OnPlay(self, event):
            self.SetStatusText('Play.')
            self.Timer.Start(TIMER_INTERVAL)
            self.VideoWindow.Play()
            
        def OnPause(self, event):
            self.VideoWindow.Pause()
            
        def OnStop(self, event):
            self.VideoWindow.Stop()
            self.SetStatusText('Stop.')

        def OnPlaySegment(self, event):
            self.Timer.Start(TIMER_INTERVAL)
            self.VideoWindow.PlaySegment(10000, 19000)

        def OnRate(self, event):
            self.Rate += 0.2
            if self.Rate > 1.5:
                self.Rate = 0.6
            self.VideoWindow.SetPlayBackSpeed(self.Rate * 10)
            
        def OnTimer(self, event):
            if self.VideoWindow.movie != None:
                pos = self.VideoWindow.GetTimecode()
                length = self.VideoWindow.GetMediaLength()
                self.txt.SetLabel('Video Position:  %10d / %10d     State:  %s     Rate: %2.1f     Time:  %s' % (pos, length, self.VideoWindow.GetState(), self.Rate, time.asctime()[-13:-5]))
                if (self.VideoWindow.VideoEndPoint > 0) and (pos > self.VideoWindow.VideoEndPoint):
                    self.OnStop(event)

        def UpdatePlayState(self, state):
            print "UpdatePlayState:", state

            
    class MyApp(wx.App):
        def OnInit(self):
            self.frame = VideoFrame()
            self.frame.Show(True)
            controller = ControlWindow(self.frame)
            self.SetTopWindow(controller)
            return True
            

    app = MyApp(0)
    app.MainLoop()
