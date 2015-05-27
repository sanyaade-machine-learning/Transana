# Copyright (C) 2003 - 2015 The Board of Regents of the University of Wisconsin System 
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

""" This module implements the MediaPlayer class for Transana.  It is primarily responsible
    for window layout and for communicating with component video_player objects. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "VideoWindow DEBUG is ON."

# Import wxPython
import wx
# import wxPython's media component
import wx.media
# Import Transana's Clip object
import Clip
# import Transana's Dialog Boxes
import Dialogs
# Import Transana's Episode object
import Episode
# Import Transana's Media Conversion dialog
import MediaConvert
# Import Transana's Constants
import TransanaConstants
# Import the Transana Exceptions
from TransanaExceptions import *
# Import Transana's Globals
import TransanaGlobal
# Import Transana's Images
import TransanaImages
# Import the Transana Video Player panel
import video_player
# import Python's os module
import os
# import Python's sys module
import sys

class VideoWindow(wx.Dialog):  # (wx.MDIChildFrame)
    """This class implements the main Transana media window. """

    def __init__(self, parent):
        """Initialize the Media Window object"""
        # Initialize a Dialog Box
        wx.Dialog.__init__(self, parent, -1, _("Media"), pos=self.__pos(), size=self.__size(),
#        wx.MDIChildFrame.__init__(self, parent, -1, _("Video"), pos=self.__pos(), size=self.__size(),
                           style = wx.RESIZE_BORDER | wx.CAPTION )
        # We need to adjust the screen position on the Mac.  I don't know why.
#        if "__WXMAC__" in wx.PlatformInfo:
#            pos = self.GetPosition()
#            self.SetPosition((pos[0], pos[1]-25))
#        self.SetBackgroundColour(wx.WHITE)
        # Bind the Size event
        self.Bind(wx.EVT_SIZE, self.OnSize)
        # Bind the Right Click event
        self.Bind(wx.EVT_RIGHT_UP, self.OnRightUp)
        # Bind the Key Down event
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        # Remember the parent window
        self.parent = parent
        # The ControlObject handles all inter-object communication, initialized to None
        self.ControlObject = None
        # Initialize a list of media player components
        self.mediaPlayers = []
        # Initialize a list that knows which media player components have started playing.
        # (Needed for starting media players that begin later than other media players)
        self.hasPlayed = []
        # Create a variable to hold the offset for all media files, needed if we have a CLIP that does not include
        # the FIRST Episode Video!
        self.globalOffset = 0
        # We need to know which player starts FIRST, to be used as the reference for all others
        self.referencePlayer = 0
        # Create the Media Players
        self.CreateMediaPlayers()

        if DEBUG:
            print "VideoWindow.__init__():  Initial size:", self.GetSize()

        wx.CallAfter(self.InitializeSize)

    def InitializeSize(self):
        """ Set the initial size of the media window based on its components.  We want to set the initial
            size of the Media Window to just show the Control Bar """
        # Determine the size of the Media Window, including frame, BEFORE the window is resized.
        # (Width is correct, but will change.)
        (width, height) = self.GetSize()
        # Resize the window so that the Button bar will be the correct size.  (Needed to compensate for different-sized
        # header bars.)
        self.Fit()
        # Determine the size of the Media Window, including frame
        winSize = self.GetSize()
        # Determine the size of the Media Window, exluding frame
        clientSize = self.GetClientSize()
        # Determine the size of the tallest component of the Control Bar, the Play/Pause button
        btnSize = self.btnPlayPause.GetSize()
        # The new height should include the FRAME size (whole window - client size) plus the height of the Play button
        newHeight = winSize[1] - clientSize[1] + btnSize[1]
        # Set the Window's initial size
        self.SetSize((width, newHeight))

#        print "VideoWindow.InitialSize():", winSize, clientSize, btnSize, newHeight, TransanaGlobal.menuHeight, TransanaGlobal.menuHeight + btnSize[1]

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.ControlObject=ControlObject

    def CreateMediaPlayers(self):
        """ Handles any changes in number of media players """
        # Create a variable to hold the offset for all media files, needed if we have a CLIP that does not include
        # the FIRST Episode Video!
        self.globalOffset = 0
        # Initialize the maximum offset
        self.maxOffset = 0
        # We need to know which player starts FIRST, to be used as the reference for all others
        self.referencePlayer = 0
        # Create a vertical box sizer as the main sizer
        vBox = wx.BoxSizer(wx.VERTICAL)
        # Create a horizontal box sizer
        hBox1 = wx.BoxSizer(wx.HORIZONTAL)
        # If there's no control object (yet) or no defined Episode / Clip object to base Media Player creation on ...
        if (self.ControlObject == None) or (self.ControlObject.currentObj == None):
            # ... See if there are defined media players.  If so ...
            if self.mediaPlayers != []:
                # ... then for each media player ...
                for mp in self.mediaPlayers:
                    # ... just stop the movie from playing.  Calling mp.Stop() involves more.
                    mp.movie.Stop()
                    # Stop the media player's ProgressNotification times (needed on Mac)
                    mp.ProgressNotification.Stop()
                    # ... and destroy it
                    mp.Destroy()
                # Indicate that all media players have been removed from the media players list
                self.mediaPlayers = []
                # Initialize the list that knows whether media players have started yet
                self.hasPlayed = []
                # Destroy the Play/Pause button too!
                self.btnPlayPause.Destroy()
                # Destroy the slider too!
                self.videoSlider.Destroy()
                # And the Snapshot button!
                if self.btnSnapshot:
                    self.btnSnapshot.Destroy()
            # Create a media player.  (With no media, it will get a Transana graphic.)
            mediaPlayer = video_player.VideoPlayer(self)
            # Add the media player to the box sizer.
            hBox1.Add(mediaPlayer, 1, wx.ALL | wx.EXPAND, 0)
            # Add the Media Player to the Media Players list
            self.mediaPlayers.append(mediaPlayer)
            # Flag that the player has not yet started
            self.hasPlayed.append(False)
            # After it's done being created, resize it.
            wx.CallAfter(self.mediaPlayers[0].OnSize, None)
        # If there is a defined Episode / Clip object to base Media Player creation on ...
        else:
            # For any existing media players ...
            for mp in self.mediaPlayers:
                # ... just stop the movie from playing.  Calling mp.Stop() involves more.
                mp.movie.Stop()
                # Stop the media player's ProgressNotification times (needed on Mac)
                mp.ProgressNotification.Stop()
                # ... and destroy the media players
                mp.Destroy()
            # Initialize the Media Players list to empty
            self.mediaPlayers = []
            # Initialize the list that knows whether the media players have started yet.
            self.hasPlayed = []
            # Destroy the Play/Pause button too!
            self.btnPlayPause.Destroy()
            # Destroy the slider too!
            self.videoSlider.Destroy()
            # And the Snapshot button!
            if self.btnSnapshot:
                self.btnSnapshot.Destroy()
            # Now we need to examine the file offsets so we can position the various media players correctly to their relative starting points.
            # Initialize a list of offsets
            offset = [0]
            # initialize a list of Audio check values
            audioCheck = []
            # Initialize the maximum offset
            self.maxOffset = 0
            # For each video in the list of additional files ...
            for vid in self.ControlObject.currentObj.additional_media_files:
                # If the offset value is greater than zero ...
                if vid['offset'] >= 0:
                    # ... and if the maximum offset value is zero or greater ...  (Positive current offset, positive max)
                    if self.maxOffset >= 0:
                        # See if the magnitude of the current value is greater than the max offset.
                        # (Must be done BEFORE appending the new value)
                        if self.maxOffset < vid['offset']:
                            # Set the new value as the new Max
                            self.maxOffset = vid['offset']
                        # Add the offset for the new file to the list, using max minus current value to calculate the OFFSET.
                        offset.append(vid['offset'])
                    # If the max offset is NEGATIVE ...  (Positive current offset, negative max)
                    else:
                        # ... and add ?????????????? to the offset list
                        offset.append(vid['offset'] + abs(self.maxOffset))
                        # If the MAGNITUDE of the max offset is smaller than the current offset value ...
                        if abs(self.maxOffset) < vid['offset']:
                            # ... set the current value as the max.
                            self.maxOffset = vid['offset']
                # If the current offset value is negative ...
                else:
                    # ... and if the maximum offset value is zero or greater ...  (Negative current offset, positive max)
                    if self.maxOffset > 0:
                        for x in range(len(offset)):
                            offset[x] += abs(vid['offset'])
                        offset.append(0)
                        # if the magnitude of the max offset is less than the magnitude of the current offset value ...
                        if self.maxOffset < abs(vid['offset']):
                            # ... set the current value as the max.
                            self.maxOffset = vid['offset']
                    # If the max offset is NEGATIVE ...  (Negative current offset, negative max)
                    else:
                        # if the (negative) max offset is LARGER THAN than the (negative) current offset value ...
                        if self.maxOffset > vid['offset']:
                            # ... add the absolute value of the current offset to the offset list
                            for x in range(len(offset)):
                                offset[x] += abs(self.maxOffset - vid['offset'])
                            offset.append(0)
                            # ... set the current value as the max.
                            self.maxOffset = vid['offset']
                        else:
                            offset.append(abs(self.maxOffset - vid['offset']))
                audioCheck.append(vid['audio'])

            # The Reference player will be the one with the 0 offset
            self.referencePlayer = offset.index(0)
            # If we are loading a CLIP and changing the global offset value ...
            if (isinstance(self.ControlObject.currentObj, Clip.Clip)) and (self.globalOffset != self.ControlObject.currentObj.offset):
                # ... get the global offset value from the current object (clip)
                self.globalOffset = self.ControlObject.currentObj.offset
                # If there are more than on media players (single media players are adjusted elsewhere)
                if len(offset) > 1:
                    # Remember the first offset value
                    offsetAdjust = offset[0]
                    # iterate through the offset values
                    for x in range(len(offset)):
                        # UNLESS we're in the first position and the value is already zero ...
                        if (x > 0) or (offset[x] != 0):
                            # ... adjust all offset values such that the first offset value IS zero by subtracting what the first value WAS.
                            # (That's why we remembered it.  It's the first thing to change!)
                            offset[x] -= offsetAdjust
            # now that we've calculated all the offsets, we can pick the largest offset value from the offsets list.
            # (All values in the offset list will be positive.)
            self.maxOffset = max(offset)
            # Create a first media player frame ...
            mediaPlayer = video_player.VideoPlayer(self, includeCheckBoxes=(len(self.ControlObject.currentObj.additional_media_files) > 0),
                                                   offset=offset[0], playerNum=1)
            # ... add it to the Sizer
            hBox1.Add(mediaPlayer, 1, wx.ALL | wx.EXPAND, 0)
            # Get the MAIN media file name ...
            fileName = self.ControlObject.currentObj.media_filename
            # ... and load it in the first media player
            mediaPlayer.SetFilename(fileName)
            # Turn off the Audio Check, if needed
            if isinstance(self.ControlObject.currentObj, Clip.Clip) and (not self.ControlObject.currentObj.audio) and (mediaPlayer.includeCheckBoxes):
                mediaPlayer.SetAudioCheck(False)
            # Add this first media player to the Media Players list
            self.mediaPlayers.append(mediaPlayer)
            # Flag that the player has not yet started
            self.hasPlayed.append(False)
            # initialize a counter for offset values
            offsetCount = 1
            if TransanaConstants.proVersion:
                # For each additional media file ...
                for vid in self.ControlObject.currentObj.additional_media_files:
                    # .. create an additional media player frame ...
                    mediaPlayer = video_player.VideoPlayer(self, includeCheckBoxes=True, offset=offset[offsetCount], playerNum=offsetCount+1)
                    # ... add it to the sizer ...
                    hBox1.Add(mediaPlayer, 1, wx.ALL | wx.EXPAND, 0)
                    # ... add it to the Media Players List ...
                    self.mediaPlayers.append(mediaPlayer)
                    # Indicate that the player has not yet started
                    self.hasPlayed.append(False)
                    # .. and load the appropriate media file.
                    mediaPlayer.SetFilename(vid['filename'])
                    # Turn off the Audio Check, if needed
                    if not vid['audio']:
                        mediaPlayer.SetAudioCheck(False)
                    # increment the counter for offset values
                    offsetCount += 1

        # Add the media players in the first horizontal sizer to the main vertical sizer
        vBox.Add(hBox1, 1, wx.EXPAND)
        # Create a second horizontal sizer
        hBox2 = wx.BoxSizer(wx.HORIZONTAL)
        # Create the Play / Pause button
        self.btnPlayPause = wx.BitmapButton(self, -1, TransanaGlobal.GetImage(TransanaImages.Play), size=(48, 24))
        # Set the Help String
        self.btnPlayPause.SetToolTipString(_("Play"))
        # Add LayoutDirection to prevent problems with Right-To-Left languages
        self.btnPlayPause.SetLayoutDirection(wx.Layout_LeftToRight)
        # Bind the Play / Pause button to its event handler
        self.btnPlayPause.Bind(wx.EVT_BUTTON, self.OnPlayPause)
        # Allow the PlayPause Button to handle Key Down events too
        self.btnPlayPause.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        # Add the Play / Pause button to the second horizontal sizer
        hBox2.Add(self.btnPlayPause, 0)
        # Create the video position slider
        self.videoSlider = wx.Slider(self, -1, 0, 0, 1000, style=wx.SL_HORIZONTAL)
        # Bind the video position slider to its event handler
        self.videoSlider.Bind(wx.EVT_SCROLL, self.OnScroll)
        # Allow the Video Slider to handle Key Down events too
        self.videoSlider.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        # Adjust position on the Mac
        if 'wxMac' in wx.PlatformInfo:
            indent = 6
        else:
            indent = 0
        # Add the slider to the horizontal sizer
        hBox2.Add(self.videoSlider, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, indent)

        # If there's exactly one media player, we want a Snapshot button
        if len(self.mediaPlayers) == 1:
            # Create the Snapshot button
            self.btnSnapshot = wx.BitmapButton(self, -1, TransanaGlobal.GetImage(TransanaImages.Snapshot), size=(48, 24))
            # Set the Help String
            self.btnSnapshot.SetToolTipString(_("Capture Snapshot"))
            # Add LayoutDirection to prevent problems with Right-To-Left languages
            self.btnSnapshot.SetLayoutDirection(wx.Layout_LeftToRight)
            # Bind the Snapshot button to its event handler
            self.btnSnapshot.Bind(wx.EVT_BUTTON, self.OnSnapshot)
            # Allow the Snapshot Button to handle Key Down events too
            self.btnSnapshot.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
            # Add the Snapshot button to the second horizontal sizer
            hBox2.Add(self.btnSnapshot, 0)
        # If we don't create the Snapshot button ...
        else:
            # ... set the varaible to None so we don't try to destroy it!!
            self.btnSnapshot = None

        # Add the second horizontal sizer to the vertical sizer.
        vBox.Add(hBox2, 0, wx.EXPAND)

        # Set the main Sizer
        self.SetSizer(vBox)
        # Enable Auto Layout
        self.SetAutoLayout(True)
        # Lay out the Video Window
        self.Layout()

# Public methods
    def open_media_file(self):
        """Open one or more media files for playback and reset media."""
        # Create the necessary Media Players
        self.CreateMediaPlayers()
        # Once the window is finished rendering, we need to refresh it to clear ghost media file images on the Mac
        wx.CallAfter(self.Refresh)

    def OnPlayPause(self, event):
        """ Event Handler for the Play / Pause Button """
        # If a media file is loaded ...
        if (self.ControlObject.currentObj != None):
            # If the currently showing item is not a Transcript (is a Document) ...
            if not (self.ControlObject.GetCurrentItemType() == 'Transcript'):
                # ... then bring the Transcript to the front of the Document Stack
                self.ControlObject.BringTranscriptToFront()
            # ... tell the control object to play or pause  (Run this through the Control Object so speed control etc. works.)
            self.ControlObject.PlayPause()

    def OnSnapshot(self, event):
        """ Take a Snapshot of the current video frame and insert into Transcript if possible. """
        # If a media file is loaded ...
        if self.ControlObject.currentObj != None:
            # Initialize an error message to no error
            msg = ''
            # If we have an Episode or a Clip ...            
            if isinstance(self.ControlObject.currentObj, Episode.Episode) or \
               isinstance(self.ControlObject.currentObj, Clip.Clip):
                # ... get the media filename from the object
                mediaFile = self.ControlObject.currentObj.media_filename
                # If the transcript is not editable ...
                if self.ControlObject.ActiveTranscriptReadOnly():
                    # ... create an error message
                    msg = _("The current transcript is not editable.  The requested snapshot can be saved to disk, but cannot be inserted into the transcript.")
                    msg += '\n\n' + _("To insert the snapshot into the transcript, press the Edit Mode button on the Document Toolbar to make the transcript editable.")
            # if not (If we have a Document or Quote) ...
            else:
                # ... get the media filename from the media player
                mediaFile = self.mediaPlayers[0].FileName
                # If the document is not editable ...
                if self.ControlObject.ActiveTranscriptReadOnly():
                    # ... create an error message
                    msg = _("The current document is not editable.  The requested snapshot can be saved to disk, but cannot be inserted into the document.")
                    msg += '\n\n' + _("To insert the snapshot into the document, press the Edit Mode button on the Document Toolbar to make the document editable.")
            if msg != '':
                dlg = Dialogs.InfoDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()
            if mediaFile != '':
                # Create the Media Conversion dialog, including Clip Information so we export only the clip segment
                convertDlg = MediaConvert.MediaConvert(self, mediaFile, self.GetCurrentVideoPosition(), snapshot=True)
                # Show the Media Conversion Dialog
                convertDlg.ShowModal()
                # If the user took a snapshop and the image was successfully created ...
                if convertDlg.snapshotSuccess and os.path.exists(convertDlg.txtDestFileName.GetValue() % 1):
                    # ... ask the Control Object to communicate with the transcript to insert this image.
                    self.ControlObject.TranscriptInsertImage(convertDlg.txtDestFileName.GetValue() % 1)
                # We need to explicitly Close the conversion dialog here to force cleanup of temp files in some circumstances
                convertDlg.Close()
                # Destroy the Media Conversion Dialog
                convertDlg.Destroy()
            else:
                msg = _("Cannot take a Snapshot at the moment....")
                dlg = Dialogs.InfoDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()

    def OnScroll(self, event):
        """ Event Handler for the Video Position Scroll Bar """
        if (self.ControlObject.GetCurrentItemType() == 'Transcript'):
            # Determine the upper and lower bounds for the current video segment.
            # If we are showing an Episode ...
            if isinstance(self.ControlObject.currentObj, Episode.Episode):
                # ... then the upper and lower bounds are 0 (the video start) and the length of the media file
                start = 0
                end = self.ControlObject.GetMediaLength(True)
            # If we are showing a Clip ...
            elif isinstance(self.ControlObject.currentObj, Clip.Clip):
                # ... then we should use the clip start and stop points
                start = self.ControlObject.currentObj.clip_start
                end = self.ControlObject.currentObj.clip_stop
            # If neither an Episode nor a Clip is loaded ...
            else:
                # ... then "disable" the scroll bar by not letting it move off of 0
                self.videoSlider.SetValue(0)
                return
            # Determine the correct video position.
            # (Video range * slider position [divided into 1000 segments] plus the media starting position)
            newPos = (end - start) * event.GetPosition() / 1000 + start
            # Set the new video selection
            self.ControlObject.SetVideoSelection(newPos, -1)
        
    def OnKeyDown(self, event):
        """ Handle Key Down events """
        # See if the ControlObject wants to handle the key that was pressed.
        if self.ControlObject.ProcessCommonKeyCommands(event):
            # If so, we're done here.  (We're done anyway!)
            return

    def Play(self):
        """Start playback."""
        # Get the media player position
        mpPos = self.GetCurrentVideoPosition()
        # For each media players ...
        for mp in self.mediaPlayers:
            # If the media current position is contained within the current media player's start - end range ... 
            # [i.e If the media player's offset is less than the current video position AND
            # the current video position (+ 6 because code in video_player sets mp.Timecode to length - 5 if past the end!!)
            # is less than the offset length of the media player ... ]
            if (mp.offset <= mpPos) and (mpPos + 6 < mp.GetMediaLength() + mp.offset):
                # ... tell the media player to Play!
                mp.Play()
                # Signal that this player HAS been played!
                self.hasPlayed[mp.playerNum - 1] = True
            # If the media player's offset is larger than the current video position ...
            else:
                # ... DON'T play it, but signal that is has NOT been played yet.  (When the current position
                # catches up, the media file will start playing automatically.)
                self.hasPlayed[mp.playerNum - 1] = False
        # Change the button image
        self.btnPlayPause.SetBitmapLabel(TransanaGlobal.GetImage(TransanaImages.Pause))
        # Set the Help String
        self.btnPlayPause.SetToolTipString(_("Pause"))
                
    def Pause(self):
        """Pause playback."""
        # For each media players ...
        for mp in self.mediaPlayers:
            # ... tell it to Pause!
            mp.Pause()
        # Change the button image
        self.btnPlayPause.SetBitmapLabel(TransanaGlobal.GetImage(TransanaImages.Play))
        # Set the Help String
        self.btnPlayPause.SetToolTipString(_("Play"))

    def Stop(self):
        """Reset media, stop playback or pause mode, set seek to start."""
        # For each media players ...
        for mp in self.mediaPlayers:
            # ... tell it to Stop!
            mp.Stop()
        # Change the button image
        self.btnPlayPause.SetBitmapLabel(TransanaGlobal.GetImage(TransanaImages.Play))
        # Set the Help String
        self.btnPlayPause.SetToolTipString(_("Play"))

    def SetVideoStartPoint(self, TimeCode):
        """ Sets the Video Starting Point. """
        # if possible ....
        if self.globalOffset != None:
            # Adjust the TimeCode value by the the global offset for a CLIP (which is 0 for an Episode).
            TimeCode -= self.globalOffset
        # For each media players ...
        for mp in self.mediaPlayers:
            # If the new time code is greater than 0 ...
            if TimeCode >= 0:
                # ... then use the time code value
                tcVal = TimeCode
            # If the new time code is LESS than 0 ...
            else:
                # ...if the video is supposed to be playing ...
                if mp.IsPlaying():
                    # ... pause the playback ...
                    mp.Pause()
                    # ... and signal that we have moved before the video start (so video will start when needed)
                    self.hasPlayed[mp.playerNum - 1] = False
                # ... and finally set the time code value to 0, its legal minimum
                tcVal = 0
            # ... and set the video start point, adjusting for the media player's offset
            mp.SetVideoStartPoint(tcVal)

        # Update the video slider position in the video window.
        # Determine the upper and lower bounds for the current video segment.
        # If we are showing an Episode ...
        if isinstance(self.ControlObject.currentObj, Episode.Episode):
            # ... then the upper and lower bounds are 0 (the video start) and the length of the media file
            start = 0
            end = self.ControlObject.GetMediaLength(True)
        # If we are showing a Clip ...
        elif isinstance(self.ControlObject.currentObj, Clip.Clip):
            # ... then we should use the clip start and stop points
            start = self.ControlObject.currentObj.clip_start
            end = self.ControlObject.currentObj.clip_stop
        # If neither an Episode nor a Clip is loaded ...
        else:
            # ... then "disable" the scroll bar by not letting it move off of 0
            self.videoSlider.SetValue(0)
            return
        # Determine the appropriate slider position.
        # If the video length is NOT known ...
        if ((start == 0) and (end <= 0)) or (end - start == 0):
            # ... position the slider at the start
            newPos = 0
        # if the video length is known ...
        else:
            # ... determine the appropriate slider position
            newPos = long(float(TimeCode + self.globalOffset - start) / abs(float(end - start)) * 1000.0)
        # Set the new video selection
        self.videoSlider.SetValue(newPos)

    def GetVideoStartPoint(self):
        """ Gets the Video Starting Point. """
        return self.mediaPlayers[self.referencePlayer].GetVideoStartPoint()

    def SetVideoEndPoint(self, TimeCode):
        """ Sets the Video End Point """
        # Adjust the TimeCode value by the the global offset for a CLIP (which is 0 for an Episode).
        TimeCode -= self.globalOffset
        # For each media players ...
        for mp in self.mediaPlayers:
            # If the time code is <= 0 ...
            if TimeCode <= 0:
                # ... then set the end point to the end of this media file
                tcVal = -1
            # Otherwise ...
            else:
                # ... the end point should be the smaller of the TimeCode or the media file's length.
                tcVal = min(TimeCode, mp.GetMediaLength() + mp.offset)
            # Set the video end point
            mp.SetVideoEndPoint(tcVal)

    def GetVideoEndPoint(self):
        """ Gets the Video End Point """
        if self.referencePlayer == 0 or TransanaConstants.proVersion:
            return self.mediaPlayers[self.referencePlayer].GetVideoEndPoint()
        else:
            return self.mediaPlayers[0].GetVideoEndPoint()

    def GetCurrentVideoPosition(self):
        """ Gets the Current Video Position """
        # We need to find the correct current video position value.  It might be the same for all media players.
        # If some havent' started yet, it might be the minimum value.  If some but not all have stopped, it might be
        # the maximum value.  It's kinda hard to know.  Let's figure it out.

        # Start by finding the minimum value across media players. Start with a large integer...
        minVal = sys.maxint
        # For each media player ...
        for mp in self.mediaPlayers:
            # ... pick the smaller of the previously found minimum value of the current media player's current position, adjusted for offset
            minVal = min(minVal, mp.GetTimecode() - mp.offset)
        # If ANY media player's current position is before the largest offset, we are at the beginning and looking for a minimum value.
        if minVal < self.maxOffset:
            lookForMin = True
            # If we allow multiple media players ...
            if TransanaConstants.proVersion:
                # ... get the time code value from the correct media player
                tcVal = self.mediaPlayers[self.referencePlayer].GetMediaLength()
            # If we only allow one media player ...
            else:
                # ... get the time code value from the only media player!
                tcVal = self.mediaPlayers[0].GetMediaLength()
        # If NO media player's current position is before the largest offset, we are NOT looking for a minimum, so the max value will work.
        else:
            lookForMin = False
            tcVal = 0
        # Iterate through the media players ...
        for mp in self.mediaPlayers:
            # Get the media player's current position
            mpVal = mp.GetTimecode()
            # If the current position is before the offset ...
            if lookForMin:
                tcVal = min(tcVal, mpVal)
            else:
                # ... remembering the larger of the previous Max value or the current position, adjusted for player offsets
                tcVal = max(tcVal, mpVal)
        # Return the Max value
        return tcVal

    def SetCurrentVideoPosition(self, TimeCode):
        """ Sets the current Video Position """
        # Adjust the TimeCode value by the global offset (for Clips.  Global Offset is 0 for Episodes.)
        TimeCode -= self.globalOffset
        # For each media players ...
        for mp in self.mediaPlayers:
            # ... set the current video position
            mp.SetCurrentVideoPosition(TimeCode)

    def GetPlayBackSpeed(self):
        """ Gets the playback speed """
        return self.mediaPlayers[0].GetPlayBackSpeed()

    def SetPlayBackSpeed(self, playBackSpeed):
        """ Sets the playback speed """
        # For each media players ...
        for mp in self.mediaPlayers:
            # ... set the current video position, adjusting for the media player's offset
            mp.SetPlayBackSpeed(playBackSpeed)

    def UpdatePlayState(self, playState):
        """ Take a PlayState Change from the Video Player and pass it on to the ControlObject.
            This is used in determining Screen Layout for Presentation Mode. """
        if self.ControlObject != None:
            self.ControlObject.UpdatePlayState(playState)
        # If the Play/Pause button is defined ...
        if self.btnPlayPause:
            # ... and the media is PLAYING ...
            if playState == wx.media.MEDIASTATE_PLAYING:
                # ... the Play/Pause button image should be "Pause"
                img = TransanaGlobal.GetImage(TransanaImages.Pause)
                # ... with the matching Help text
                helpStr = _("Pause")
            # If the media is NOT playing ...
            else:
                # ... the Play/Pause button image should be "Play"
                img = TransanaGlobal.GetImage(TransanaImages.Play)
                # ... with the matching Help text
                helpStr = _("Play")
            # Display the appropriate image on the Play/Pause button
            self.btnPlayPause.SetBitmapLabel(img)
            # Set the Help String
            self.btnPlayPause.SetToolTipString(helpStr)
            # PPC seems to need this explicit call to update the screen
            self.Refresh()

    def UpdateVideoPosition(self, currentPosition, playerNum=1):
        """ Take Video Position information from the Video Player and pass it on to the ControlObject. """
        # Find the first player that is playing ...
        firstPlayingPlayer = 1
        # Iterate through the players ...
        for mp in self.mediaPlayers:
            # If it's playing ...
            if mp.IsPlaying():
                # ... stop looking.  We found a playing player!
                break
            # Iterate the player counter if it's not playing
            firstPlayingPlayer += 1
        # If this comes from the first-playing-Player ...
        if playerNum == firstPlayingPlayer:
            # ... then update Transana.  (No need to do it for every change in every media player.)
            self.ControlObject.UpdateVideoPosition(currentPosition + self.globalOffset)
            # If the video is repositioned while not playing, we need to reset the Video Start Point.
            if not self.IsPlaying():
                self.ControlObject.SetVideoStartPoint(currentPosition + self.globalOffset)
            # Check to see if the current segment has ended, remembering to adjust for offset.
            if (self.mediaPlayers[firstPlayingPlayer - 1].VideoEndPoint > 0) and \
               (currentPosition >= self.mediaPlayers[firstPlayingPlayer - 1].VideoEndPoint + self.mediaPlayers[firstPlayingPlayer - 1].offset):
                # if so, STOP!
                self.Stop()
                
            # For each media player ...
            for mp in self.mediaPlayers:
                # If it hasn't already been played and the current position has reached its offset and we haven't moved PAST the end of the media file ...
                if (not self.hasPlayed[mp.playerNum - 1]) and (currentPosition >= mp.offset) and (currentPosition < mp.GetMediaLength() + mp.offset):
                    # ... signal that this player HAS been played ...
                    self.hasPlayed[mp.playerNum - 1] = True
                    # ... and tell it to play!
                    mp.Play()

            # Update the video slider position in the video window.
            # Determine the upper and lower bounds for the current video segment.
            # If we are showing an Episode ...
            if isinstance(self.ControlObject.currentObj, Episode.Episode):
                # ... then the upper and lower bounds are 0 (the video start) and the length of the media file
                start = 0
                end = self.ControlObject.GetMediaLength(True)
            # If we are showing a Clip ...
            elif isinstance(self.ControlObject.currentObj, Clip.Clip):
                # ... then we should use the clip start and stop points
                start = self.ControlObject.currentObj.clip_start
                end = self.ControlObject.currentObj.clip_stop
            # If neither an Episode nor a Clip is loaded ...
            else:
                # ... then "disable" the scroll bar by not letting it move off of 0
                self.videoSlider.SetValue(0)
                return
            # Determine the appropriate slider position
            newPos = int(float(currentPosition + self.globalOffset - start) / float(end - start) * 1000.0)
            # Set the new video selection
            self.videoSlider.SetValue(newPos)

    def UpdateVideoWindowPosition(self, left, top, width, height):
        """ When the Video Window is updated, it should request that all other windows be repositioned accordingly
            by the ControlObject """
        if self.ControlObject != None:
            self.ControlObject.UpdateVideoWindowPosition(left, top, width, height)

    def GetMediaLength(self):
        """ Get Media Length information from the Video Player. """
        # We need to find the largest current video position value.  Therefore, initialize a max value variable to -1
        lenMax = -1
        # Iterate through the media players ...
        for mp in self.mediaPlayers:
            # Get the media length
            mediaLen = mp.GetMediaLength()
            # If we have a media length:
            if mediaLen > -1:
                # ... remembering the larger of the previous Max value or the current length, adjusted for player offsets
                lenMax = max(lenMax, mp.GetMediaLength() + mp.offset)
        # Return the Max value
        return lenMax

    def ClearVideo(self):
        """Clear the display."""
        # If there is a defined "current object" ...
        if self.ControlObject.currentObj != None:
            # Clear the current object
            self.ControlObject.currentObj = None
            # Recreate (ie. clear) the Media Players' interface
            self.CreateMediaPlayers()
        # Reset the Video Window to its Initial (very small) size
        self.InitializeSize()
        # Refresh the graphic in the video window
        self.Refresh()

    def GetDimensions(self):
        """ Returns the dimensions of the Video Window """
        (left, top) = self.GetPositionTuple()
        (width, height) = self.GetSizeTuple()
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
        """ Set window dimensions """
        self.SetDimensions(left, top, width, height)

    def close(self):
        """ Close the Video Frame """
        # For each media player ...
        for mp in self.mediaPlayers:
            # ... Just stop the videos.  Calling mp.Stop() involves more than that.
            mp.movie.Stop()
        # Signal to the Control Object that we're shutting down.  This indicates that the Video Window is no longer available.
        if self.ControlObject != None:
            self.ControlObject.shuttingDown = True
        # Actually close the window
        self.Close()

    def ReadyState(self):
        """ Returns the "ReadyState" of the video control """
        return self.mediaPlayers[0].ReadyState()

    def IsPlaying(self):
        """ Indicates whether the video is currently playing or not """
        # If there are no defined Media Players ...
        if self.mediaPlayers == []:
            # ... they can't very well be playing, can they?
            IsPlaying = False
        # If there ARE media players ...
        else:
            # ... initialize to False
            IsPlaying = False
            # initialize the first playing player number to None
            player1 = None
            # Iterate through the media players.
            for mp in self.mediaPlayers:
                # If this player is playing ...
                if mp.IsPlaying():
                    # If it's the FIRST playing player ...
                    if player1 == None:
                        # ... not its player number (zero-based)
                        player1 = mp.playerNum - 1
                        # ... and note its offset
                        offset1 = mp.offset
                    # If it's NOT the first playing player ...
                    else:
                        # This method gets called frequently during video playback.  Therefore, it's an acceptable place
                        # to efficiently check for media synchronization problems.
                        
                        # Get this media player's offset
                        offset2 = mp.offset
                        # Determine the difference between THIS media player's position and the first playing media player's position,
                        # adjusting each for its offset
                        TCDifference = (self.mediaPlayers[player1].movie.Tell() + offset1) - (mp.movie.Tell() + offset2)

                        mp.synch.SetLabel('%d' % TCDifference)
                        
                        # If the offset is more than one frame (33/1000th of a second) ...
                        if abs(TCDifference) > 33:
                            # ... tell the movie to try to match the first media player's position, adjusting for both offsets
                            # (Use a fresh position, as the video may have moved!)
                            mp.movie.Seek(self.mediaPlayers[player1].movie.Tell() + offset1 - offset2)
                    # If ANY player is playing, we want this to be true
                    IsPlaying = True
        # Return the IsPlaying value found.
        return IsPlaying

    def IsPaused(self):
        """ Indicates whether the video is currently paused or not """
        # Default to Paused
        isPaused = True
        # Check all media players
        for mp in self.mediaPlayers:
            # If ANY media player is NOT paused, then this should be False.
            isPaused = isPaused and mp.IsPaused()
        # Return the result
        return isPaused
    
    def IsStopped(self):
        """ Indicates whether the video is currently stopped or not """
        # Default to Stopped
        isStopped = True
        # Check all media players
        for mp in self.mediaPlayers:
            # If ANY media player is NOT stopped, then this should be False.
            isStopped = isStopped and mp.IsStopped()
        # Return the result
        return isStopped

    def IsLoading(self):
        """ Indicates whether the video is currently loading into the player """
        # Default to NOT loading.
        IsLoading = False
        # Check all media players
        for mp in self.mediaPlayers:
            # If ANY media player is loading, then this should be True.
            IsLoading = IsLoading or mp.IsLoading()
        # Return the result
        return IsLoading

    def OnSize(self, event):
        """ Process Size Change event and notify the ControlObject """
        # if event is not None (which it can be if this is called from non-event-driven code rather than
        # from a real event,) then we should process underlying OnSize events.
        if event != None:
            event.Skip()
        # If we are not resizing ALL the windows ...  (otherwise, avoid redundant or recursive calls.)
        if not TransanaGlobal.resizingAll:
            # for updating window position, we need the full window size and position
            (width, height) = self.GetSize()
            pos = self.GetPosition()
            # We can now inform other windows of the size and position of the Video window
            self.UpdateVideoWindowPosition(pos[0], pos[1], width, height)

        # For each Media Player ...
        for mp in self.mediaPlayers:
            # ... adjust to the current Window size
            mp.OnSize(None)
        # Call update to try to resolve the Mac display problem
        self.Update()

    def OnSizeChange(self):
        """ Size Change called programatically from outside the Video Window """
        # If Auto Arrange is enabled ...
        if TransanaGlobal.configData.autoArrange:

            # If there is no "current object" loaded in the main interface ...
            if self.ControlObject.currentObj == None:
                # ... get the size of the graphic in the first media player window
                (sizeX, sizeY) = self.mediaPlayers[0].graphic.GetSize()
            # If there is a current object in the main interface ...
            else:
                #  Establish the minimum size of the media player control (if media is audio-only, for example)
                (sizeX, sizeY) = (0, 0)
                # Check the media players
                for mp in self.mediaPlayers:
                    # Get the size of the video 
                    (newSizeX, newSizeY) = mp.movie.GetBestSize()
                    # We want the dimensions of the largest video.
                    sizeX += newSizeX   # = max(sizeX, newSizeX)
                    sizeY = max(sizeY, newSizeY)

                # If we have no WIDTH (Audio ONLY) ...
                if sizeX == 0:
                    # ... give the media player 1/4 of the screen
                    sizeX = int(wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()[2] * 0.25)
                # Adjust Video Size
                sizeAdjust = TransanaGlobal.configData.videoSize
                # Take Video Size Menu spec into account (50%, 66%, 100%, 150%, 200%)
                sizeX = int(sizeX * sizeAdjust / 100.0)
                sizeY = int(sizeY * sizeAdjust / 100.0)
                # If we have ONLY audio files, we will have sizeY == 0.  This is now okay if there's only one media player,
                # since we don't have a control bar in the media player, but causes problems with multiple media players.
                # (It makes the visualization window too short.)
                if (sizeY == 0) and (len(self.mediaPlayers) > 1):
                    sizeY = int((wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()[3] - TransanaGlobal.menuHeight) * 0.2)  # wx.ClientDisplayRect()
                # if the PlayPause button is defined ...
                if self.btnPlayPause:
                    # ... adjust the vertical size to include it.
                    sizeY += self.btnPlayPause.GetSize()[1]
                # if not ...
                else:
                    # ... then 24 pixels is a good approximation.
                    sizeY += 24
            #  Determine the screen size
            # On Mac ...
            if 'wxMac' in wx.PlatformInfo:
                # ... use wx.Display.GetClientArea
                screenSize = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
                # We need to adjust by 17 pixels
                xAdjust = 17
            # On Windows ...
            else:
                # ... use MenuWindow.GetClientWindow.GetSize
                #screenSize = (0, 0, self.ControlObject.MenuWindow.GetClientWindow().GetSize()[0], self.ControlObject.MenuWindow.GetClientWindow().GetSize()[1])
                screenSize = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()  # wx.ClientDisplayRect()
                # We need to adjust by 23 pixels
                xAdjust = 23
            # now check width against the screen size.  Allow no more than 2/3 of the screen to be take up by video.
            if sizeX > int(screenSize[2] * 0.66):
                # Adjust Height proportionally (first, so we can calculate the proportion!)
                sizeY = int(sizeY * (screenSize[2] * 0.66) / sizeX)
                # Adjust Width to 2/3 of screen
                sizeX = int(screenSize[2] * 0.66)
            # If there are check boxes in the media players ...
            if (len(self.mediaPlayers) > 0) and self.mediaPlayers[0].includeCheckBoxes:
                # ... increase the height of the video window by the height of the check box.
                sizeY += self.mediaPlayers[0].includeInClip.GetSize()[1]
            #  Get the current position of the Video Window
            pos = self.GetPosition()
            # We need to know the height of the Window Header to adjust the size of the Graphic Area
            headerHeight = self.GetSizeTuple()[1] - self.GetClientSizeTuple()[1]
            # Set the final screen dimensions
            self.SetDimensions(screenSize[0] + screenSize[2] - sizeX - xAdjust, pos[1], sizeX + 16, sizeY + headerHeight)
            # Call update to try to resolve the Mac display problem
            self.Update()

    def OnRightUp(self, event):
        """ Right Mouse Button Up event handler """
        # If there is something loaded in the main interface ...
        if self.ControlObject.currentObj != None:
            # ... then play/pause the video THROUGH THE CONTROL OBJECT so that playback rate is honored.
            self.ControlObject.PlayPause()

    def GetVideoCheckboxDataForClips(self, videoPos):
        """ Return the data about media players checkboxes needed for Clip creation """
        # Create an empty list
        result = []
        # For each Media Player ...
        for mp in self.mediaPlayers:
            # Get the checkbox data
            checkboxData = mp.GetVideoCheckboxDataForClips()
            # If there was data returned ...
            if checkboxData != None:
                # ... check for  video being out of bounds, either before the start (offset) or after the end (len + offset)
                # (or a just-loaded Episode, which will have cvp of 0 if in MOV format), ...
                if (videoPos != 0) and ((videoPos < mp.offset) or (videoPos > mp.GetMediaLength() + mp.offset)):
                    # ... and if out of bounds, override the checkboxes to cancel the out-of-bounds media player.
                    checkboxData = (False, False)
                # ... append the media player's Clip Data to the list
                result.append(checkboxData)
        # Return the results
        return result

    def VideoCheckboxChange(self):
        """ Detect and report changes in the status of the media player checkboxes """
        # When the checkboxes change, let the ControlObject know.
        self.ControlObject.VideoCheckboxChange()

    def ChangePlaybackSpeed(self, direction):
        """ Alter the playback speed on the fly by a small amount """
        # Get the current rate of the first Media Player.  Rates range from 0.1 to 2.0.
        rate = self.mediaPlayers[0].GetPlayBackSpeed()
        # If we're increasing the rate ...
        if (direction == 'faster') and (rate < 2.0):
            # ... increase the rate by a tenth.
            rate += 0.1
        # If we're reducing the rate ...
        elif (direction == 'slower') and (rate > 0.15):
            # ... slow the rage by a tenth
            rate -= 0.1
        # For each media player ...
        for mp in self.mediaPlayers:
            # ... set the new playback speed.  Rates here range from 1 to 20, a factor of 10 larger than GetPlayBackSpeed.  Weird.
            mp.SetPlayBackSpeed(rate * 10)

    def ChangeLanguages(self):
        """ Update all prompts for the Video Window when changing interface languages """
        self.btnPlayPause.SetToolTipString(_("Play"))
        self.btnSnapshot.SetToolTipString(_("Snapshot"))

    def GetNewRect(self):
        """ Get (X, Y, W, H) for initial positioning """
        pos = self.__pos()
        size = self.__size()
        return (pos[0], pos[1], size[0], size[1])

        
# Private methods

    def __size(self):
        """Determine default size of MediaPlayer Frame."""
        # Determine which monitor to use and get its size and position
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            primaryScreen = TransanaGlobal.configData.primaryScreen
        else:
            primaryScreen = 0
        rect = wx.Display(primaryScreen).GetClientArea()
        if not 'wxGTK' in wx.PlatformInfo:
            container = rect[2:4]
        else:
            screenDims = wx.Display(primaryScreen).GetClientArea()
            # screenDims2 = wx.Display(primaryScreen).GetGeometry()
            left = screenDims[0]
            top = screenDims[1]
            width = screenDims[2] - screenDims[0]  # min(screenDims[2], 1280 - self.left)
            height = screenDims[3]
            container = (width, height)
        width = container[0] * .282   # rect[2] * .28
        if 'wxMac' in wx.PlatformInfo:
            height = 40
        else:
            # This doesn't really matter.  It gets re-adjusted elsewhere in InitialSize()
            height = (container[1] - TransanaGlobal.menuHeight) * .068  # .339
        return wx.Size(width, height)

    def __pos(self):
        """Determine default position of MediaPlayer Frame."""
        # Determine which monitor to use and get its size and position
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            primaryScreen = TransanaGlobal.configData.primaryScreen
        else:
            primaryScreen = 0
        rect = wx.Display(primaryScreen).GetClientArea()
        if not 'wxGTK' in wx.PlatformInfo:
            container = rect[2:4]
        else:
            # Linux rect includes both screens, so we need to use an alternate method!
            container = TransanaGlobal.menuWindow.GetSize()
        (width, height) = self.__size()
        # rect[0] compensates if the Start menu is on the left side of the screen.
        if 'wxGTK' in wx.PlatformInfo:
            x = rect[0] + min((rect[2] - 10), (1280 - rect[0])) - width
        else:
            x = rect[0] + container[0] - width - 2  # rect[0] + rect[2] - width - 3
        # rect[1] compensates if the Start menu is on the top of the screen
        if 'wxMac' in wx.PlatformInfo:
            y = rect[1] + 2
        else:
            y = rect[1] + TransanaGlobal.menuHeight + 1
        return wx.Point(x, y)
