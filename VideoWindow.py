# Copyright (C) 2003 - 2005 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the MediaPlayer class as part of the Media component.
It is primarily responsible for making sure the proper platform-specific code
is used, and it provides a common interface for all platform-specific modules
"""

# Most of the methods at this level do nothing more than pass the message on to the
# equivalent platform-specific control frame.

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas A. Sambhare, Nathaniel Case'

import wx
from TransanaExceptions import *
import TransanaGlobal

if wx.Platform == "__WXMSW__":
    from video_msw import *
elif wx.Platform == "__WXMAC__":
    from video_mac import *
else:
    raise GeneralError, \
        "Sorry, no media player support is available on this platform."

class VideoWindow(object):
    """This class implements the media player responsible for various playback
    functions such as start, stop, pause, seek. It also handles video size and
    fullscreen modes."""

    def __init__(self, parent):
        """Initialize a MediaPlayer/Quicktime object"""
    
        """something is broken on the mac:  you can't send the self object"""
        if "__WXMAC__" in wx.PlatformInfo:
            """see http://www.dzug.org/mailinglisten/zope-org-zope/archive/2003/2003-04/1051704653383"""
            self.frame = VideoFrame(parentVideoWindow = self, pos = self.__pos(), size= self.__size())
            # We need to adjust the screen position on the Mac.  I don't know why.
            pos = self.frame.GetPosition()
            self.frame.SetPosition((pos[0], pos[1]-25))
        
        else:        
            self.frame = VideoFrame(parent = parent, parentVideoWindow = self, pos = self.__pos(), size= self.__size())

        # print "VideoWindow:", self.__pos(), self.__size()
        
        self.ControlObject = None            # The ControlObject handles all inter-object communication, initialized to None
      

    def Show(self, value=True):
        """ Show (or Hide) the Video Player Window. """
        self.frame.Show(value)

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.ControlObject=ControlObject


# Public methods
    def open_media_file(self, video_file):
        """Open a media file for playback and reset media."""
        self.frame.SetFilename(video_file)
        
    def close_media_file(self):
        """Clear the media display, unload media file."""
        
    def Play(self):
        """Start playback."""
        self.frame.Play()
        
    def Pause(self):
        """Pause playback."""
        self.frame.Pause()
        
    def Stop(self):
        """Reset media, stop playback or pause mode, set seek to start."""
        self.frame.Stop()

    def SetVideoStartPoint(self, TimeCode):
        """ Sets the Video Starting Point. """
        self.frame.SetVideoStartPoint(TimeCode)
        
    def GetVideoStartPoint(self):
        """ Gets the Video Starting Point. """
        return self.frame.GetVideoStartPoint()

    def SetVideoEndPoint(self, TimeCode):
        """ Sets the Video End Point """
        self.frame.SetVideoEndPoint(TimeCode)

    def GetCurrentVideoPosition(self):
        """ Gets the Current Video Position """
        return self.frame.GetTimecode()

    def SetCurrentVideoPosition(self, TimeCode):
        """ Sets the current Video Position """
        self.frame.SetCurrentVideoPosition(TimeCode)

    def GetPlayBackSpeed(self):
        """ Gets the playback speed """
        return self.frame.GetPlayBackSpeed()

    def SetPlayBackSpeed(self, playBackSpeed):
        """ Sets the playback speed """
        self.frame.SetPlayBackSpeed(playBackSpeed)

    def UpdatePlayState(self, playState):
        """ Take a PlayState Change from the Video Player and pass it on to the ControlObject.
            This is used in determining Screen Layout for Presentation Mode. """
        self.ControlObject.UpdatePlayState(playState)

    def UpdateVideoPosition(self, currentPosition):
        """ Take Video Position information from the Video Player and pass it on to the ControlObject. """
        self.ControlObject.UpdateVideoPosition(currentPosition)
        # If the video is repositioned while not playing, we need to reset the Video Start Point.
        if not self.IsPlaying():
            self.ControlObject.SetVideoStartPoint(currentPosition)

    def UpdateVideoWindowPosition(self, left, top, width, height):
        """ When the Video Window is updated, it should request that all other windows be repositioned accordingly
            by the ControlObject """
        self.ControlObject.UpdateVideoWindowPosition(left, top, width, height)

    def GetMediaLength(self):
        """ Get Media Length information from the Video Player. """
        return self.frame.GetMediaLength()

    def ClearVideo(self):
        """Clear the display."""
        # Set the video filename to nothing.  That should clear everything!!
        self.frame.SetFilename('')

    def GetDimensions(self):
        """ Returns the dimensions of the Video Window """
        (left, top) = self.frame.GetPositionTuple()
        (width, height) = self.frame.GetSizeTuple()
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
        self.frame.SetDimensions(left, top, width, height)

    def play_pause_gb(self):
        """Play if paused, pause if playing. Implicit go-back by 2 seconds, 
        for easy transcription."""
            
    def fast_forward(self):
        """Start playback in fast forward mode if supported by media."""
    def slow_motion(self):
        """Start playback in slow motion mode if supported by media.""" 
    def seek(self, media_position):
        """Update the seekbar position to given value, or return current seek
        position if called with -1."""
    def original_size_screen(self):
        """Set display size to 100%."""
    def onepointfive_size_screen(self):
        """Set display size to 150%."""
    def double_size_screen(self):
        """Set display size to 200%."""
    def full_size_screen(self):
        """Set fullscreen display."""
    def video_and_transcript_screen(self):
        """Set display to presentation mode size."""
    def audio_only_screen(self):
        """Set display for audio only."""
    def set_volume(self):
        """Set volume for playback."""
    def mute_unmute(self):
        """Toggle the mute setting."""
    def supported_media_types(self):
        """Return a list of supported media types."""
        # FIXME: This is just a hardcoded placeholder for now
        return ["mpg", "mp3", "wav"]

    def close(self):
        """ Close the Video Frame """
        self.frame.Close()

    def ReadyState(self):
        """ Returns the "ReadyState" of the video control """
        return self.frame.ReadyState()

    def IsPlaying(self):
        """ Indicates whether the video is currently playing or not """
        return self.frame.IsPlaying()

    def IsPaused(self):
        """ Indicates whether the video is currently paused or not """
        return self.frame.IsPaused()
    
    def IsStopped(self):
        """ Indicates whether the video is currently stopped or not """
        return self.frame.IsStopped()

    def IsLoading(self):
        """ Indicates whether the video is currently loading into the player """
        return self.frame.IsLoading()

# Private methods

    def __size(self):
        """Determine default size of MediaPlayer Frame."""
        rect = wx.ClientDisplayRect()
        width = rect[2] * .28
        height = (rect[3] - TransanaGlobal.menuHeight) * .35
        return wx.Size(width, height)

    def __pos(self):
        """Determine default position of MediaPlayer Frame."""
        rect = wx.ClientDisplayRect()
        (width, height) = self.__size()
        # rect[0] compensates if the Start menu is on the left side of the screen.
        x = rect[0] + rect[2] - width - 3
        # rect[1] compensates if the Start menu is on the top of the screen
        y = rect[1] + TransanaGlobal.menuHeight + 3
        return wx.Point(x, y)

    def _set_media_filename(self, media_filename):
        self._media_filename = media_filename
    def _get_media_filename(self):
        return self._media_filename
    def _del_media_filename(self):
        del self._media_filename

    def _set_media_loaded(self, media_loaded):
        self._media_loaded = media_loaded
    def _get_media_loaded(self):
        return self._media_loaded
    def _del_media_loaded(self):
        del self._media_loaded

    def _set_volume(self, volume):
        self._volume = volume
    def _get_volume(self):
        return self._volume
    def _del_volume(self):
        del self._volume

    def _set_mute(self, mute):
        self._mute = mute
    def _get_mute(self):
        return self._mute
    def _del_mute(self):
        del self._mute

    def _set_playback_rate(self, playback_rate):
        self._playback_rate = playback_rate
    def _get_playback_rate(self):
        return self._playback_rate
    def _del_playback_rate(self):
        del self._playback_rate

    def _set_media_position(self, media_position):
        self._media_position = media_position
    def _get_media_position(self):
        return self._media_position
    def _del_media_position(self):
        del self._media_position

# Public properties
    media_filename = property(_get_media_filename, _set_media_filename,
    _del_media_filename, \
    """Name of the media file.""")
    media_loaded = property(_get_media_loaded, _set_media_loaded,
    _del_media_loaded, \
    """A boolean indicating whether media has been loaded.""")
    volume = property(_get_volume, _set_volume, _del_volume, \
    """Volume for playback.""")
    mute = property(_get_mute, _set_mute, _del_mute, \
    """A boolean indicating whether playback is muted.""")
    playback_rate = property(_get_playback_rate, _set_playback_rate,
    _del_playback_rate, \
    """Number indicating rate of playback, n for n-fast playback, -n for
    n-reverse playback.""")
    media_position = property(_get_media_position, _set_media_position,
    _del_media_position, \
    """Current media position.""")
