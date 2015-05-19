# Copyright (C) 2008 - 2009 The Board of Regents of the University of Wisconsin System 
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

""" This unit handles the synchrnonization of media files.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "Synchronize DEBUG is ENABLED!!"

# Import Python modules
import os
import pickle
import sys
import time

# Import wxPython
import wx

# Import Transana Dialogs
import Dialogs
# Import Transana's Graphics Control Class for making the waveforms
import GraphicsControlClass
# Import Transana's Miscellaneous module for time formatting functions
import Misc
# Import Transana's Globals
import TransanaGlobal
# Import the Video Player module
import video_player
# Import Transana's module for creating Waveform Graphics
import WaveformGraphic
# Import Transana's Waveform (Audio Extraction) Progress Dialog
import WaveformProgress

class Synchronize(wx.Dialog):
    """ This tool assists in synchronizing media files. """
    def __init__(self, parent, filename1, filename2, offset=0):
        # Remember the original offset
        self.offset = offset
        # Initialize a Dialog Box
        wx.Dialog.__init__(self, parent, -1, _("Synchronize Media Files"), size = (800,600),
                           style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        # Freeze the dialog.  This prevents screen updates, speeding up the creation process.
        self.Freeze()
        # Set the initial size as the minimum size.
        self.SetSizeHints(minW = self.GetSize()[0], minH = self.GetSize()[1])
        # Set the background to White
        self.SetBackgroundColour(wx.WHITE)
        
        # The Video Player widget expects a Control Object.  We don't need a full-blown Control Object here.  That would
        # interfere with the rest of Transana.  So we'll create a minimal, local Control Object.
        self.ControlObject = SynchronizeControlObject(self)

        # Initialize the Media Player windows to None
        self.mp1 = None
        self.mp2 = None
        # initialize the Waveform Zoom level to 30K, showing about a 30 second segment of video
        self.waveformSpan = 29000.0
        # Initialize video positions based on starting offset.  Add 15 seconds to both so the waveform cursor can be centered in the waveform.
        if offset > 0:
            self.pos1 = offset + 15000
            self.pos2 = 15000
        else:
            self.pos1 = 15000
            self.pos2 = abs(offset) + 15000

        # Signal that the screen has not yet been built, which will disable some methods that interfere with
        # creating the dialog initially.
        self.windowBuilt = False
        # Initialize the Currently Selected Media Player to None (neither)
        self.currentPlayer = None
        # Initialize the current Waveform to None (neither) also
        self.currentWaveform = None

        # Create a main Box Sizer to contain everything.  The main purpose of this is to provide a whitespace border,
        # as the GridBagSizer doesn't appear to offer that.
        mainBoxSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a main Grid Bag Sizer to contain all the actual screen elements.
        self.mainGBSSizer = wx.GridBagSizer(3, 3)
        # Add the first file name to the GBS first row, spanning 4 columns
        txt = wx.StaticText(self, -1, filename1)
        self.mainGBSSizer.Add(txt, (0, 0), (1, 4))
        
        # Create the first media player
        self.mp1 = video_player.VideoPlayer(self, size=(360, 240), playerNum=1)
        # Load the first file into the first media player.
        self.mp1.SetFilename(filename1, offset=self.pos1)
        # Check for File Not Found, and quit if it's not
        if self.mp1.FileName == "":
            return
        # Add the first media player to the GBS second row, spanning 3 columns
        self.mainGBSSizer.Add(self.mp1, (1, 0), (1, 3))
        # Create the first Waveform control
        self.waveform1 = GraphicsControlClass.GraphicsControl(self, -1, wx.Point(1, 1), wx.Size(414, 219), (414, 219), visualizationMode=True, name='mp1')
        # initialize the start position
        self.start1 = 0
        # Create the waveform for the first file, if needed.  If successful, the name of the WAV file is returned.
        self.waveFile1 = self.LoadWaveform(self.waveform1, filename1)
        # If the waveform fails ...
        if self.waveFile1 == False:
            # ... display an error message ...
            print "Waveform FAIL for", filename1
            # .. and exit the Synchronize routine.
            return
        # Add the waveform to the GBS's second row, fourth column, and make it expandable.
        self.mainGBSSizer.Add(self.waveform1, (1, 3), flag=wx.EXPAND)

        # Add a Play button for the top media player
        self.btnPlay1 = wx.Button(self, -1, _("Play Top"))
        # Bind the button to the OnPlay event handler
        self.btnPlay1.Bind(wx.EVT_BUTTON, self.OnPlay)
        # Add the play button to the GBS, third row, first column
        self.mainGBSSizer.Add(self.btnPlay1, (2, 0))
        # Add a Nudge Left button for the top media player
#        self.btnNudgeLeft1 = wx.Button(self, -1, _("Nudge Left"))
        # Bind the button to the OnNudge event handler
#        self.btnNudgeLeft1.Bind(wx.EVT_BUTTON, self.OnNudge)
        # Add the nudge button to the GBS, third row, second column
#        self.mainGBSSizer.Add(self.btnNudgeLeft1, (2, 1))
        # Add a Nudge Right button for the top media player
#        self.btnNudgeRight1 = wx.Button(self, -1, _("Nudge Right"))
        # Bind the button to the OnNudge event handler
#        self.btnNudgeRight1.Bind(wx.EVT_BUTTON, self.OnNudge)
        # Add a small box sizer
        scrollSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the nudge button to the scroll sizer
#        scrollSizer.Add(self.btnNudgeRight1, 0, wx.ALIGN_LEFT)
        # Add a spacer
        scrollSizer.Add((0,0), 1, wx.EXPAND)
        # Add a label
        lbl = wx.StaticText(self, -1, _('Scroll:'))
        # Add the label to the scroll sizer
        scrollSizer.Add(lbl, 0, wx.ALIGN_RIGHT | wx.TOP, 3)
        # Add the scroll sizer to the GBS, third row, third column
        self.mainGBSSizer.Add(scrollSizer, (2, 2), flag=wx.EXPAND)
        # Add a video scroll slider for the top video
        self.slider1 = wx.Slider(self, -1, 0, 0, 500, style = wx.SL_HORIZONTAL)
        # Add a Tool Tip to label the slider
        self.slider1.SetToolTipString(_("Scroll Top Video"))
        # Bind the slider to the OnSlider Event Handler
        self.slider1.Bind(wx.EVT_SCROLL, self.OnSlider)
        # Add the slider to the GBS, third row, fourth column, and make it expandable
        self.mainGBSSizer.Add(self.slider1, (2, 3), flag=wx.EXPAND)

        # Add a Play button for both media players
        self.btnPlay2 = wx.Button(self, -1, _("Play Both"))
        # Bind the button to the OnPlay event handler
        self.btnPlay2.Bind(wx.EVT_BUTTON, self.OnPlay)
        # Add the play button to the GBS, fourth row, first column
        self.mainGBSSizer.Add(self.btnPlay2, (3, 0))
        # Add a TextCtrl for offset
        self.txtOffset = wx.TextCtrl(self, -1, "%10d" % self.offset, style=wx.TE_RIGHT)
        # Add the slider to the GBS, fourth row, second column
        self.mainGBSSizer.Add(self.txtOffset, (3, 1))
        # Add a small box sizer
        zoomSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer
        zoomSizer.Add((0, 0), 1, wx.EXPAND)
        # Add a label
        lbl = wx.StaticText(self, -1, _('Zoom:'))
        # Add the label to the zoom sizer
        zoomSizer.Add(lbl, 0, wx.ALIGN_RIGHT | wx.TOP, 3)
        # Add the zoom sizer to the GBS, fourth row, third column
        self.mainGBSSizer.Add(zoomSizer, (3, 2), flag=wx.EXPAND)
        # Add the waveform zoom slider
        self.waveformZoom = wx.Slider(self, -1, int(self.waveformSpan / 100), 5, 1200, style = wx.SL_HORIZONTAL)
        # Add a Tool Tip to label the slider
        self.waveformZoom.SetToolTipString(_("Waveform Zoom"))
        # Bind the slider to the OnSlider Event Handler
        self.waveformZoom.Bind(wx.EVT_SCROLL, self.OnSlider)
        # Add the slider to the GBS, fourth row, fourth column, and make it expandable
        self.mainGBSSizer.Add(self.waveformZoom, (3, 3), flag=wx.EXPAND)

        # Add a Play button for the bottom media player
        self.btnPlay3 = wx.Button(self, -1, _("Play Bottom"))
        # Bind the button to the OnPlay event handler
        self.btnPlay3.Bind(wx.EVT_BUTTON, self.OnPlay)
        # Add the play button to the GBS, fifth row, first column
        self.mainGBSSizer.Add(self.btnPlay3, (4, 0))
        # Add a Nudge Left button for the bottom media player
#        self.btnNudgeLeft2 = wx.Button(self, -1, _("Nudge Left"))
        # Bind the button to the OnNudge event handler
#        self.btnNudgeLeft2.Bind(wx.EVT_BUTTON, self.OnNudge)
        # Add the nudge button to the GBS, fifth row, second column
#        self.mainGBSSizer.Add(self.btnNudgeLeft2, (4, 1))
        # Add a Nudge Right button for the bottom media player
#        self.btnNudgeRight2 = wx.Button(self, -1, _("Nudge Right"))
        # Bind the button to the OnNudge event handler
#        self.btnNudgeRight2.Bind(wx.EVT_BUTTON, self.OnNudge)
        # Add a small box sizer
        scrollSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the nudge button to the scroll sizer
#        scrollSizer.Add(self.btnNudgeRight2, 0, wx.ALIGN_LEFT)
        # Add a spacer
        scrollSizer.Add((0,0), 1, wx.EXPAND)
        # Add a label
        lbl = wx.StaticText(self, -1, _('Scroll:'))
        # Add the label to the scroll sizer
        scrollSizer.Add(lbl, 0, wx.ALIGN_RIGHT | wx.TOP, 3)
        # Add the scroll sizer to the GBS, fifth row, third column
        self.mainGBSSizer.Add(scrollSizer, (4, 2), flag=wx.EXPAND)
        # Add a video scroll slider for the bottom video
        self.slider2 = wx.Slider(self, -1, 0, 0, 500, style = wx.SL_HORIZONTAL)
        # Add a Tool Tip to label the slider
        self.slider2.SetToolTipString(_("Scroll Bottom Video"))
        # Bind the slider to the OnSlider Event Handler
        self.slider2.Bind(wx.EVT_SCROLL, self.OnSlider)
        # Add the slider to the GBS, fifth row, fourth column, and make it expandable
        self.mainGBSSizer.Add(self.slider2, (4, 3), flag=wx.EXPAND)

        # Add the second file name to the GBS sixth row, spanning 4 columns
        txt = wx.StaticText(self, -1, filename2)
        self.mainGBSSizer.Add(txt, (5, 0), (1, 4))

        # Create the second media player
        self.mp2 = video_player.VideoPlayer(self, size=(360, 240), playerNum=2)
        # Load the second file into the second media player.
        self.mp2.SetFilename(filename2, offset=self.pos2)
        # Check for File Not Found, and quit if it's not
        if self.mp2.FileName == "":
            return
        # Add the second media player to the GBS seventh row, spanning 3 columns
        self.mainGBSSizer.Add(self.mp2, (6, 0), (1, 3))
        # Create the second Waveform control
        self.waveform2 = GraphicsControlClass.GraphicsControl(self, -1, wx.Point(1, 1), wx.Size(414, 219), (414, 219), visualizationMode=True, name='mp2')
        # initialize the start position
        self.start2 = 0
        # Create the waveform for the second file, if needed.  If successful, the name of the WAV file is returned.
        self.waveFile2 = self.LoadWaveform(self.waveform2, filename2)
        # If the waveform fails ...
        if self.waveFile2 == False:
            # ... display an error message ...
            print "Waveform FAIL for", filename2
            # .. and exit the Synchronize routine.
            return

        # Add the waveform to the GBS's seventh row, fourth column, and make it expandable.
        self.mainGBSSizer.Add(self.waveform2, (6, 3), flag=wx.EXPAND)

        # Make the GBS' second row (Media Player / Waveform 1) "growable"
        self.mainGBSSizer.AddGrowableRow(1)
        # Make the GBS' seventh row (Media Player / Waveform 2) "growable"
        self.mainGBSSizer.AddGrowableRow(6)
        # Make the GBS' fourth row (Waveforms) "growable"
        self.mainGBSSizer.AddGrowableCol(3)

        # Add the GBS to the main Box sizer, giving form a 6 pixel border all the way around
        mainBoxSizer.Add(self.mainGBSSizer, 1, wx.EXPAND | wx.ALL, 6)
        # Create a sizer for the form buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a horizontal spacer to the button sizer
        btnSizer.Add((0, 0), 1)
        # Create an OK button
        self.btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Add the OK button to the button sizer
        btnSizer.Add(self.btnOK, 0, wx.ALIGN_RIGHT | wx.TOP | wx.RIGHT | wx.BOTTOM, 6)
        # Create a Cancel button
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        # Add the Cancel button to the button sizer
        btnSizer.Add(btnCancel, 0, wx.ALIGN_RIGHT | wx.TOP | wx.RIGHT | wx.BOTTOM, 6)
        # Create a Help button
        btnHelp = wx.Button(self, -1, _("Help"))
        # Add the Help button to the button sizer
        btnSizer.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.TOP | wx.RIGHT | wx.BOTTOM, 6)
        # Bind a handler to the Help button
        btnHelp.Bind(wx.EVT_BUTTON, self.Help)
        # Add the button sizer to the main form sizer
        mainBoxSizer.Add(btnSizer, 0, wx.EXPAND)

        # Bind events for the mouse entering and leaving the two Waveform diagrams
        self.waveform1.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseEnter)
        self.waveform1.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        self.waveform2.Bind(wx.EVT_ENTER_WINDOW, self.OnMouseEnter)
        self.waveform2.Bind(wx.EVT_LEAVE_WINDOW, self.OnMouseExit)
        # Bind the form's OnSize event
        self.Bind(wx.EVT_SIZE, self.OnSize)

        # Set the main Box sizer as the form's main sizer
        self.SetSizer(mainBoxSizer)
        # Call Thaw to allow the dialog to finally be drawn.
        self.Thaw()

        # Fit (resize) the window based on the controls
        self.Fit()
        # Turn AutoLayout on
        self.SetAutoLayout(True)
        # Lay out the dialog
        self.Layout()
        # Center the dialog on the screen
        self.CenterOnScreen()

        # Now that this is all done, we can signal that the dialog has been built
        self.windowBuilt = True

    def GetData(self):
        """ Shows the Synchronize Dialog (modally) and returns the offset between the files """
        # If the user pressed OK ...
        if self.ShowModal() == wx.ID_OK:
            # ... return the difference between the two media file positions
            return (self.pos1 - self.pos2, self.mp2.GetMediaLength())
        # Otherwise ...
        else:
            # .. return the original offset that was passed in.
            return (self.offset, self.mp2.GetMediaLength())

    def OnMouseEnter(self, event):
        """ Event handler for mousing over a Waveform graphic """
        # Get a pointer to the event's sending object
        obj = event.GetEventObject()
        # If the event was triggered by Waveform 1 ...
        # if event.GetId() == self.waveform1.GetId():  doesn't work on the Mac.  We're getting an out-dated ID for the waveform!
        if obj.name == 'mp1':   
            # ... then we should manipulate Media Player 1
            self.currentPlayer = self.mp1
            # ... and the current Waveform to Waveform 1
            self.currentWaveform = self.waveform1
            # ... and the current start position
            self.currentStart = self.start1
        # If the event was triggered by Waveform 2 ...
        # if event.GetId() == self.waveform2.GetId():  doesn't work on the Mac.  We're getting an out-dated ID for the waveform!
        elif obj.name == 'mp2':
            # ... then we should manipulate Media Player 2
            self.currentPlayer = self.mp2
            # ... and the current Waveform to Waveform 2
            self.currentWaveform = self.waveform2
            # ... and the current start position
            self.currentStart = self.start2

    def OnMouseExit(self, event):
        """ Event handler for mousing out of a Waveform graphic """
        # If we've left the waveform, neither media player is in play.
        self.currentPlayer = None
        # ... and neither waveform is in play.
        self.currentWaveform = None
        # ... and neither start is in play.
        self.currentStart = None

    def OnPlay(self, event):
        """ Event Handler for the Play / Pause buttons """
        # This proved more complicated than I'd expected!  Hang on!

        # We need to know if one or both of the media players is/are currently playing
        isPlaying1 = self.mp1.IsPlaying()
        isPlaying2 = self.mp2.IsPlaying()

        # We change the status of media player 1 IF:
        #   the TOP button is pressed OR
        #   the MIDDLE button is pressed and either the top player is playing or neither player is playing.
        play1 = (event.GetId() == self.btnPlay1.GetId()) or \
                ((event.GetId() == self.btnPlay2.GetId()) and (isPlaying1 or (not isPlaying1 and not isPlaying2)))
        # We change the status of media player 2 IF:
        #   the TOP button is pressed OR
        #   the MIDDLE button is pressed and either the top player is playing or neither player is playing.
        play2 = (event.GetId() == self.btnPlay3.GetId()) or \
                ((event.GetId() == self.btnPlay2.GetId()) and (isPlaying2 or (not isPlaying1 and not isPlaying2)))

        # Get the media player(s) playing correctly.
        # If the TOP or MIDDLE Play button was pressed ...
        if play1:
            # ... and if the top player is playing ...
            if isPlaying1:
                # ... then pause it.
                self.mp1.Pause()
            # If the top player is NOT playing ...
            else:
                # ... then tell it to play.
                self.mp1.Play()
        # If the MIDDLE or BOTTOM Play button was pressed ...
        if play2:
            # ... and if the bottom player is playing ...
            if isPlaying2:
                # ... then pause it.
                self.mp2.Pause()
            # If the bottom player is NOT playing ...
            else:
                # ... then tell it to play.
                self.mp2.Play()

        # Now, we need to update the form buttons.
        if isPlaying1:  # Player 1 was playing before the button press...
            if play1:   # The top button read "Pause"
                self.btnPlay1.SetLabel(_("Play top"))
#                self.btnNudgeLeft1.Enable(True)
#                self.btnNudgeRight1.Enable(True)
        else:           # Player 1 was NOT playing before the button press ...
            if play1:   # The top button read "Play"
                self.btnPlay1.SetLabel(_("Pause top"))
#                self.btnNudgeLeft1.Enable(False)
#                self.btnNudgeRight1.Enable(False)
        if isPlaying2:  # Player 2 was playing before the button press ...
            if play2:   # The bottom button read "Pause"
                self.btnPlay3.SetLabel(_("Play bottom"))
#                self.btnNudgeLeft2.Enable(True)
#                self.btnNudgeRight2.Enable(True)
        else:           # Player 2 was NOT playing before the button press ...
            if play2:   # The bottom button read "Play"
                self.btnPlay3.SetLabel(_("Pause bottom"))
#                self.btnNudgeLeft2.Enable(False)
#                self.btnNudgeRight2.Enable(False)

        # If both players are NOW paused ...
        if (isPlaying1 and play1 and not isPlaying2) or \
           (isPlaying2 and play2 and not isPlaying1) or \
           (isPlaying1 and play1 and isPlaying2 and play2):
            # ... change the middle button to Play
            self.btnPlay2.SetLabel(_("Play both"))
            # You can only EXIT the form with OK if both players are paused!  (This makes the synch more accurate.)
            self.btnOK.Enable(True)
        # If either player is playing ...
        else:
            # ... change the middle button to Pause ...
            self.btnPlay2.SetLabel(_("Pause both"))
            # ... and disable the OK button!
            self.btnOK.Enable(False)
        # Update the Waveforms
        self.UpdateWaveforms()

    def OnNudge(self, event):
        """ Event handler for the two "Nudge" buttons """
        # If either Player is playing, ignore the Nudge because it's not accurate!
        if self.mp1.IsPlaying() or self.mp2.IsPlaying():
            return
        # Determine which button caused the event to fire
        btnId = event.GetId()
        # Set the Frame Size to 33 milliseconds, which is about 1 / 30th of a second.
        frameSize = 33
        # Get the current positions of the media windows
        pos1 = self.mp1.GetTimecode()
        pos2 = self.mp2.GetTimecode()
        # Depending on which button was pressed, determine which position needs to be adjusted which way
        if btnId == self.btnNudgeLeft1.GetId():
            pos1 -= frameSize
        elif btnId == self.btnNudgeRight1.GetId():
            pos1 += frameSize
        elif btnId == self.btnNudgeLeft2.GetId():
            pos2 -= frameSize
        elif btnId == self.btnNudgeRight2.GetId():
            pos2 += frameSize
        # Still depending on which button was pressed, make the actual position adjustment and update the display.
        if btnId in [self.btnNudgeLeft1.GetId(), self.btnNudgeRight1.GetId()]:
            self.mp1.SetCurrentVideoPosition(pos1)
            self.UpdateVideoPosition(pos1, 1)
        elif btnId in [self.btnNudgeLeft2.GetId(), self.btnNudgeRight2.GetId()]:
            self.mp2.SetCurrentVideoPosition(pos2)
            self.UpdateVideoPosition(pos2, 2)

    def OnSlider(self, event):
        """ Event handler for the Media Position and Zoom Sliders """
        # Call the underlying parent event.  (Mac wasn't letting go of the handles!)
        event.Skip()
        # Default the selected media player to None
        mp = None
        # If the top slider fired the event, it affects Media Player 1 and Waveform 1
        if event.GetId() == self.slider1.GetId():
            mp = self.mp1
            mpN = 1
        # If the bottom slider fired the event, if affects Media Player 2 and Waveform 2
        elif event.GetId() == self.slider2.GetId():
            mp = self.mp2
            mpN = 2
        # If the Zoom slider fired the event, if affects BOTH waveforms, but neither media player.
        elif event.GetId() == self.waveformZoom.GetId():
            # Get the new Zoom value from the slider position
            self.waveformSpan = event.GetPosition() * 100
            # Signal that BOTH waveforms need updating
            mpN = -1

        # If one of the media players should be affected ...
        if mp != None:
            # Determine the length of the selected media file
            mediaLength = mp.GetMediaLength()
            # As long as the media length is known ...
            if mediaLength > -1:
                # ... determine the new position based on the slider position and the media length
                newPos = int((event.GetPosition() / 500.0) * mediaLength)
                # Actually set the new position in the appropriate player
                mp.SetCurrentVideoPosition(newPos)
                if mpN == 1:
                    self.pos1 = newPos
                else:
                    self.pos2 = newPos
                # Update the offset text box
                self.txtOffset.SetValue("%10d" % (self.pos1 - self.pos2))
        # Update the waveform(s) as set above.
        self.UpdateWaveforms(mpN)
            

    # Routines required by the Video Player API

    def UpdatePlayState(self, playState):
        """ Update Play State method """
        # This method is required by video_player for Transana, but not for Synchronize.  Therefore, do nothing!
        pass

    def UpdateVideoPosition(self, videoPos, playerNum):
        """ Update the interface based on a new video position """
        # If the form has been completed ... (to prevent problems during form construction)
        if self.windowBuilt:
            # If info is coming from player 1 ...
            if playerNum == 1:
                # Determine the new value from Player 1
                self.pos1 = videoPos
                # Update the value from Player 2 to keep things synchronized
                self.pos2 = self.mp2.GetTimecode()
                # Update Waveform 1
                self.UpdateWaveforms(1)
            # If info is coming from player 2 ...
            elif playerNum == 2:
                # Determine the new value from Player 2
                self.pos2 = videoPos
                # Update the value from Player 1 to keep things synchronized
                self.pos1 = self.mp1.GetTimecode()
                # Update Waveform 2
                self.UpdateWaveforms(2)
            # Update the offset text box
            self.txtOffset.SetValue("%10d" % (self.pos1 - self.pos2))

    # Routines required by the Waveform API

    def PctPosFromTimeCode(self, timecode):
        """ Required in Transana, not in Synchronize """
        return 1

    def TimeCodeFromPctPos(self, pctPos):
        """ Required in Transana, not in Synchronize """
        return 0

    def OnMouseOver(self, x, y, xpct, ypct):
        """ Required in Transana, not in Synchronize """
        pass

    def OnLeftDown(self, x, y, xpct, ypct):
        """ Required in Transana, not in Synchronize """
        pass

    def OnLeftUp(self, x, y, xpct, ypct):
        """ Handle the release of the left mouse button """
        # Set the video position, based on the percentage position in the waveform, the zoom level, and the current time start position
        self.currentPlayer.SetCurrentVideoPosition(xpct * self.waveformSpan + self.currentStart)
        # Determine which player we're using so we can update the correct position variable
        if self.currentPlayer.GetId() == self.mp1.GetId():
            self.pos1 = xpct * self.waveformSpan + self.currentStart
        else:
            self.pos2 = xpct * self.waveformSpan + self.currentStart
        # Update the offset text box
        self.txtOffset.SetValue("%10d" % (self.pos1 - self.pos2))

    def OnRightUp(self, *args):
        """ Required in Transana by both Visualization and Video Player, but not in Synchronize """
        pass
        
    def IsPlaying(self):
        """ Is there a playing player? """
        # Are media players MUST be defined?
        if (self.mp1 != None) and (self.mp2 != None):
            # If so, is either one playing?
            return self.mp1.IsPlaying() or self.mp2.IsPlaying()
        # if not ...
        else:
            # ... then they can't be playing, can they?
            return False

    # Synchronize Routines

    def LoadWaveform(self, waveformGraphicControl, mediaFile):
        """ Load a Waveform, based on the mediaFile, into the GraphicControl """
        # Clear the graphic control
        waveformGraphicControl.Clear()
        # Separate the path from the file name
        (path1, fn1) = os.path.split(mediaFile)
        # Separate the file name from the extension
        (fnroot1, ext1) = os.path.splitext(fn1)
        # Determine the correct WAV file name
        waveFilename1 = os.path.join(TransanaGlobal.configData.visualizationPath, fnroot1 + '.wav')

        # We just have to assume that audio extraction worked.  Signal success!
        dllvalue = 0
        # If the WAV file does NOT exist, we need to do Audio Extraction.
        if not os.path.exists(waveFilename1):
            # Start Exception Handling
            try:
                # If the Waveforms Directory does not exist, create it.
                if not os.path.exists(TransanaGlobal.configData.visualizationPath):
                    # (os.makedirs is a recursive call to create ALL needed folders!)
                    os.makedirs(TransanaGlobal.configData.visualizationPath)

                # Build the progress box's label
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("Extracting %s\nfrom %s"), 'utf8')
                else:
                    prompt = _("Extracting %s\nfrom %s")
                # Create the Waveform Progress Dialog
                progressDialog = WaveformProgress.WaveformProgress(self, waveFilename1, prompt % (waveFilename1, mediaFile))
                # Tell the Waveform Progress Dialog to handle the audio extraction modally.
                progressDialog.Extract(mediaFile, waveFilename1)
                # Okay, we're done with the Progress Dialog here!
                progressDialog.Destroy()
            # handle exceptions
            except UnicodeDecodeError:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Unicode Decode Error:  %s : %s'), 'utf8')
                else:
                    prompt = _('Unicode Decode Error:  %s : %s')
                errordlg = Dialogs.ErrorDialog(self, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
                dllvalue = 1  # Signal that the WAV file was NOT created!
            except:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Unable to create Waveform Directory.\n%s\n%s'), 'utf8')
                else:
                    prompt = _('Unable to create Waveform Directory.\n%s\n%s')
                errordlg = Dialogs.ErrorDialog(self, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
                dllvalue = 1  # Signal that the WAV file was NOT created!                        

                # Close the Progress Dialog when the DLL call is complete
                progressDialog.Close()
        # If the WAV file was found or created ...
        if dllvalue == 0:
            # ... return the WAV file name
            return waveFilename1
        # If the process failed ...
        else:
            # ... signal failure.
            return False

    def UpdateWaveforms(self, waveformToUpdate=-1):
        """ Update one or both of the Waveform displays """
        # If the form is complete and we're not ONLY supposed to do the second waveform ...
        if self.windowBuilt and waveformToUpdate != 2:
            # Clear the first waveform
            self.waveform1.Clear()
            # Get the current position from media player 1
            start1 = self.mp1.GetTimecode()
            # if the current position cursor should fall in the right half of the graphic ...
            if start1 > self.waveformSpan / 2:
                # ... adjust start so the cursor will fall in the middle of the graphic
                start1 -= self.waveformSpan / 2
            # Remember the current start position
            self.start1 = start1
            # Load the new graphic into the Waveform Control.  (Passing ":memory:" rather than a filename causes WaveformGraphic to return a Bitmap object rather than saving it to a file!)
            self.waveform1.backgroundImage = WaveformGraphic.WaveformGraphicCreate([{'filename' : self.waveFile1, 'offset' : 0, 'length' : self.mp1.GetMediaLength()}], ':memory:', start1, self.waveformSpan, self.waveform1.GetSize())
            # Determine the current cursor position
            pos = ((float(self.mp1.GetTimecode() - start1)) / self.waveformSpan)
            # Draw the Waveform 1 cursor
            self.waveform1.DrawCursor(pos)
            # Determine the Waveform 1 Slider position
            sliderPos = int((float(self.mp1.GetTimecode()) / self.mp1.GetMediaLength()) * 500) + 1
            # Set the Waveform 1 slider position
            self.slider1.SetValue(sliderPos)
        # If the form is complete and we're not ONLY supposed to do the first waveform ...
        if self.windowBuilt and waveformToUpdate != 1:
            # Clear the second waveform
            self.waveform2.Clear()
            # Get the current position from media player 2
            start2 = self.mp2.GetTimecode()
            # If the current position shoudl fall on the right half of the graphic ...
            if start2 > self.waveformSpan / 2:
                # ... adjust the start so the cursor will fall in the middle of the graphic
                start2 -= self.waveformSpan / 2
            # Remember the current start position
            self.start2 = start2
            # Load the new graphic into the Waveform Control.  (Passing ":memory:" rather than a filename causes WaveformGraphic to return a Bitmap object rather than saving it to a file!)
            self.waveform2.backgroundImage = WaveformGraphic.WaveformGraphicCreate([{'filename' : self.waveFile2, 'offset' : 0, 'length' : self.mp2.GetMediaLength()}], ':memory:', start2, self.waveformSpan, self.waveform2.GetSize())
            # Determine the current cursor position
            pos = ((float(self.mp2.GetTimecode() - start2)) / self.waveformSpan)
            # Draw the Waveform 2 cursor
            self.waveform2.DrawCursor(pos)
            # Determine the Waveform 2 Slider position
            sliderPos = int((float(self.mp2.GetTimecode()) / self.mp2.GetMediaLength()) * 500) + 1
            # Set the Waveform 2 slider position
            self.slider2.SetValue(sliderPos)

    def OnSizeChange(self):
        """ Size Change method (not event handler!) for the Synchronize Dialog """
        # If the dialog has been completed ...
        if self.windowBuilt:
            # Get the video's "Best Size"
            (x1, y1) = self.mp1.movie.GetBestSize()
            # Determine the proper scaling factor, based on the size of the media player cell and the video height.
            if y1 > 0:
                resizeFactor = (float(self.mainGBSSizer.GetCellSize(6, 0)[1])) / float(y1)
            # If the video doesn't yet have a size, just use 1 as the scaling factor.
            else:
                resizeFactor = 1.0
            # If the video width would exceed 400 ...
            if int(x1 * resizeFactor) > 400:
                # ... adjust the scaling factor down.
                resizeFactor = 400.0 / x1
            # Set the minimum and maximum size for Media Player 1
            self.mp1.SetSizeHints(int(x1 * resizeFactor), int(y1 * resizeFactor), int(x1 * resizeFactor), int(y1 * resizeFactor))
            # Set the actual size of Media Player 1
            self.mp1.SetSize((int(x1 * resizeFactor), int(y1 * resizeFactor)))
            # Refresh Media Player 1
            self.mp1.Refresh()
            # Set the minimum and maximum size for Media Player 2
            self.mp2.SetSizeHints(int(x1 * resizeFactor), int(y1 * resizeFactor), int(x1 * resizeFactor), int(y1 * resizeFactor))
            # Set the actual size of Media Player 2
            self.mp2.SetSize((int(x1 * resizeFactor), int(y1 * resizeFactor)))
            # Refresh Media Player 2
            self.mp2.Refresh()
            # Get the screen size
            (t, l, w, h) = wx.ClientDisplayRect()
            # Set the minimum and maximum sizes for the dialog
            self.SetSizeHints(800, 600, int(0.9 * w), int(0.9 * h))
            # Trigger the sizer to re-layout the form
            self.Layout()
            # Large dimension files in some video formats cause the synchronize window to be rendered
            # too large to fit in the screen.  Check for this.
            if (self.GetSizeTuple()[0] > int(0.9 * w)) or (self.GetSizeTuple()[1] > int(0.9 * h)):
                # If too large, set the window size to the maximum (90% of the size of the screen.)
                self.SetSize((int(0.9 * w), int(0.9 * h)))
                # Center the window on the screen
                self.CenterOnScreen()
            # When the dust settles, redraw the Waveforms.
            wx.CallAfter(self.ResizeWaveforms)

    def ResizeWaveforms(self):
        """ Method useful in resizing and redrawing the Waveform diagrams """
        # If the dialog has been completely rendered ...
        if self.windowBuilt:
            # update the dialog
            self.Update()
            # Get the size of the Sizer cell for Waveform 1
            cellSize1 = self.mainGBSSizer.GetCellSize(1, 3)
            # Get the position of the Sizer cell for Waveform 1
            cellPos1 = self.waveform1.GetPosition()
            # Resize the Waveform to conform to the cell size and video height
            self.waveform1.SetDim(cellPos1[0], cellPos1[1], cellSize1[0], self.mp1.GetSize()[1])
            # Get the size of the Sizer cell for Waveform 2
            cellSize2 = self.mainGBSSizer.GetCellSize(6, 3)
            # Get the position of the Sizer cell for Waveform 1
            cellPos2 = self.waveform2.GetPosition()
            # Resize the Waveform to conform to the cell size and video height
            self.waveform2.SetDim(cellPos2[0], cellPos2[1], cellSize2[0], self.mp2.GetSize()[1])
            # Finally, update the Waveforms
            self.UpdateWaveforms()
            
    def OnSize(self, event):
        """ Size Change event handler for the Synchronize Dialog """
        # If this was triggered programatically, not via an event ...
        if event == None:
            # ... then call the OnSizeChange method using CallAfter
            wx.CallAfter(self.OnSizeChange)
        # If this IS event-driven ...
        else:
            # ... call event.Skip, which hits the Sizer
            event.Skip()
            # ... and then resize the Waveforms
            wx.CallAfter(self.ResizeWaveforms)

    def Help(self, event):
        """ Handles all calls to the Help System """
        # Define the Help Context
        helpContext = 'Synchronize Media Files'
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


class SynchronizeControlObject(object):
    """ Extremely minimal Control Object substitute for the Synchronize module """
    def __init__(self, parent):
        """ Create the Control Object """
        # Remember it's parent
        self.parent = parent
        # We are not shutting down, and never will be here.
        self.shuttingDown = False
