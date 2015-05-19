# Copyright (C) 2008 - 2014 The Board of Regents of the University of Wisconsin System 
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
# Import Transana's Images
import TransanaImages
# Import the Video Player module
import video_player
# Import Transana's module for creating Waveform Graphics
import WaveformGraphic
# Import Transana's Waveform (Audio Extraction) Progress Dialog
import WaveformProgress

class Synchronize(wx.Dialog):
    """ This tool assists in synchronizing media files. """
    def __init__(self, parent, filename1, filename2, offset=0):

        # Signal that the screen has not yet been built, which will disable some methods that interfere with
        # creating the dialog initially.
        self.windowBuilt = False

        # This form requires a minimum resolution of 1024 x 768.  It just doesn't fit at 800 x 600.
        if wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()[3] < 650:  # wx.ClientDisplayRect()
            msg = _('This form requires a screen resolution of 1024 x 768 or higher.')
            dlg = Dialogs.ErrorDialog(parent, msg)
            dlg.ShowModal()
            return

        # Remember the original offset
        self.offset = offset
        # Initialize a Dialog Box
        wx.Dialog.__init__(self, parent, -1, _("Synchronize Media Files"), size = (880, 640),
                           style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        # Freeze the dialog.  This prevents screen updates, speeding up the creation process.
        self.Freeze()
        displayRect = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
        # Set the initial size as the minimum size.
        minX = min(self.GetSize()[0], int(0.9 * displayRect[2]))  # wx.ClientDisplayRect()
        minY = min(self.GetSize()[1], int(0.9 * displayRect[3]))  # wx.ClientDisplayRect()
        self.SetSizeHints(minX, minY, int(0.9 * displayRect[2]), int(0.9 * displayRect[3]))  # wx.ClientDisplayRect()

        # With wxPython 2.9.4.0, panels get assigned a minimum size below which they try not to shrink
        # when the frame shrinks.  This was making the video files become different sizes when smaller
        # than their original size (or perhaps below 320 x 240).  This fixes that.
        self.SetMinSize((10, 10))

        # Set the background to White
        self.SetBackgroundColour(wx.WHITE)

        # The Video Player widget expects a Control Object.  We don't need a full-blown Control Object here.  That would
        # interfere with the rest of Transana.  So we'll create a minimal, local Control Object.
        self.ControlObject = SynchronizeControlObject(self)

        # In order to be able to correctly trap keypress events, we need to place all controls on a Panel on the dialog.
        # First, create a sizer for the panel
        pnlSizer = wx.BoxSizer(wx.VERTICAL)
        # Now create the panel, which wants character input.  Everything else should go on this panel.
        self.pnl = wx.Panel(self, -1, style=wx.WANTS_CHARS)
        # Place the panel on the panel sizer.  
        pnlSizer.Add(self.pnl, 1, wx.EXPAND | wx.ALL, 0)

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

        # Create a main Box Sizer to contain everything.  
        mainBoxSizer = wx.BoxSizer(wx.VERTICAL)
        
        # Create a new Horizontal sizer
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer
        hSizer.Add((10, 1), 0)
        # Add a vertical sizer for the first media player
        vSizer = wx.BoxSizer(wx.VERTICAL)
        # Add the first file name
        txt = wx.StaticText(self.pnl, -1, filename1)
        # Add this to the media player vertical sizer
        vSizer.Add(txt, 0, wx.BOTTOM, 5)
        # Create the first media player
        self.mp1 = video_player.VideoPlayer(self, size=(360, 240), playerNum=1, formPar=self.pnl)
        # Load the first file into the first media player.
        self.mp1.SetFilename(filename1, offset=self.pos1)
        # Check for File Not Found, and quit if it's not
        if self.mp1.FileName == "":
            return
        # Add the media player to the vertical sizer
        vSizer.Add(self.mp1, 8, wx.ALIGN_CENTER | wx.EXPAND, 0)
        # Add the vertical sizer to the horizontal sizer
        hSizer.Add(vSizer, 12, wx.ALIGN_CENTER | wx.EXPAND | wx.ALL, 5)
        # Add a spacer to the horizontal sizer
        hSizer.Add((10, 1), 0)
        # Add a vertical sizer for the second media player
        vSizer = wx.BoxSizer(wx.VERTICAL)
        # Add the second file name
        txt = wx.StaticText(self.pnl, -1, filename2)
        # Add this to the media player vertical sizer
        vSizer.Add(txt, 0, wx.BOTTOM, 5)
        # Create the second media player
        self.mp2 = video_player.VideoPlayer(self, size=(360, 240), playerNum=2, formPar=self.pnl)
        # Load the second file into the second media player.
        self.mp2.SetFilename(filename2, offset=self.pos2)
        # Check for File Not Found, and quit if it's not
        if self.mp2.FileName == "":
            return
        # Add the second media player to the vertical sizer
        vSizer.Add(self.mp2, 8, wx.ALIGN_CENTER | wx.EXPAND, 0)
        # Add the second vertical sizer to the horizontal sizer
        hSizer.Add(vSizer, 12, wx.ALIGN_CENTER | wx.EXPAND | wx.ALL, 5)
        # Add a spacer to the horizontal sizer
        hSizer.Add((10, 1), 0)
        # Add the horizontal sizer, containing both media players, to the main sizer
        mainBoxSizer.Add(hSizer, 1, wx.ALIGN_CENTER | wx.EXPAND | wx.BOTTOM | wx.RIGHT, 5)
        # Now that we have both media players created, get the sizes of the respective media files
        (size1x, size1y) = self.mp1.movie.GetBestSize()
        (size2x, size2y) = self.mp2.movie.GetBestSize()
        # None of these sizes can be 0 to prevent division by 0 errors
        if size1x == 0:
            size1x = 4
        if size1y == 0:
            size1y = 3
        if size2x == 0:
            size2x = 4
        if size2y == 0:
            size2y = 3
        # Determine the aspect ratio of the first media file
        aspectRatio = float(size1y) / float(size1x)
        # Set the initial media player size based on 360 pixels width and use the aspect ratio to calculate the height
        self.mp1.SetSize((360, (360 * aspectRatio)))
        # Set the initial media size based on 360 pixels width and use the aspect ratio to calculate the height
        self.mp1.movie.SetSize((360, (360 * aspectRatio)))
        # Determine the aspect ratio of the first media file
        aspectRatio = float(size2y) / float(size2x)
        # Set the initial media player size based on 360 pixels width and use the aspect ratio to calculate the height
        self.mp2.SetSize((360, (360 * aspectRatio)))
        # Set the initial media size based on 360 pixels width and use the aspect ratio to calculate the height
        self.mp2.movie.SetSize((360, (360 * aspectRatio)))

        # Create a new Horizontal sizer for the whole media control button row
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer
        hSizer.Add((10, 1), 1)
        # Create a new Vertical sizer for the two Left buttons rows
        vSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a new horizontal sizer for the top Left buttons row
        hSubSizer1 = wx.BoxSizer(wx.HORIZONTAL)
        # Determine the button widths for the position control buttons.  If we're on a Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... the buttons need a width of 52 pixels
            btnWidth = 52
        # Otherwise ...
        else:
            # ... the buttons need a width of 38 pixels
            btnWidth = 38
        # Add the first set of Position Control buttons for the LEFT media control
        self.btnLL60 = wx.Button(self.pnl, -1, "< 60", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnLL60, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnLL10 = wx.Button(self.pnl, -1, "< 10", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnLL10, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnLL1 = wx.Button(self.pnl, -1, "< 1", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnLL1, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # Add the top row to the Vertical sizer
        vSizer.Add(hSubSizer1, 0)
        # Create a new horizontal sizer for the bottom Left buttons row
        hSubSizer2 = wx.BoxSizer(wx.HORIZONTAL)
        # Add the second set of Position Control buttons for the LEFT media control
        self.btnLLFifth = wx.Button(self.pnl, -1, "< 0.2", size=(btnWidth, 20))
        hSubSizer2.Add(self.btnLLFifth, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # This one is twice as big (plus the gap) as the others
        self.btnLLFrame = wx.Button(self.pnl, -1, _("< 1 frame"), size=(2 * btnWidth + 4, 20))
        hSubSizer2.Add(self.btnLLFrame, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # Add the bottom row to the Vertical sizer
        vSizer.Add(hSubSizer2, 0)
        # Add the vertical sizer to the whole row horizontal sizer
        hSizer.Add(vSizer, 0)

        # Add a Play button for the top media player
        self.btnPlay1 = wx.BitmapButton(self.pnl, -1, TransanaGlobal.GetImage(TransanaImages.Play), size=(48, 48))
        self.btnPlay1.SetLayoutDirection(wx.Layout_LeftToRight)
        # Bind the button to the OnPlay event handler
        self.btnPlay1.Bind(wx.EVT_BUTTON, self.OnPlay)
        # Add the play button to the sizer
        hSizer.Add(self.btnPlay1, 0, wx.LEFT | wx.RIGHT, 2)
        
        # Create a new Vertical sizer for the two right buttons rows
        vSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a new horizontal sizer for the top right buttons row
        hSubSizer1 = wx.BoxSizer(wx.HORIZONTAL)
        # Add the rest of the Position Control buttons for the LEFT media control
        self.btnLR1 = wx.Button(self.pnl, -1, "1 >", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnLR1, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnLR10 = wx.Button(self.pnl, -1, "10 >", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnLR10, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnLR60 = wx.Button(self.pnl, -1, "60 >", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnLR60, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # Add the top row to the Vertical sizer
        vSizer.Add(hSubSizer1, 0)
        # Create a new horizontal sizer for the bottom Right buttons row
        hSubSizer2 = wx.BoxSizer(wx.HORIZONTAL)
        # Add the second set of Position Control buttons for the LEFT media control
        # This one is twice as big (plus the gap) as the others
        self.btnLRFrame = wx.Button(self.pnl, -1, _("1 frame >"), size=(2 * btnWidth + 4, 20))
        hSubSizer2.Add(self.btnLRFrame, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnLRFifth = wx.Button(self.pnl, -1, "0.2 >", size=(btnWidth, 20))
        hSubSizer2.Add(self.btnLRFifth, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # Add the bottom row to the Vertical sizer
        vSizer.Add(hSubSizer2, 0)
        # Add the vertical sizer to the whole row horizontal sizer
        hSizer.Add(vSizer, 0)

        # Bind the position change buttons to the position change event handler
        self.btnLL60.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnLL10.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnLL1.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnLLFifth.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnLLFrame.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnLRFrame.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnLRFifth.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnLR1.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnLR10.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnLR60.Bind(wx.EVT_BUTTON, self.OnShiftButton)

        # Add a spacer
        hSizer.Add((10, 1), 1)
        # Add a Play button for both media players
        self.btnPlay2 = wx.Button(self.pnl, -1, _("Play Both"), size=(100, 48))
        # Bind the button to the OnPlay event handler
        self.btnPlay2.Bind(wx.EVT_BUTTON, self.OnPlay)
        # Add the play button to the sizer
        hSizer.Add(self.btnPlay2, 0)
        # Add a spacer
        hSizer.Add((10, 1), 1)

        # Create a new Vertical sizer for the two Left buttons rows
        vSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a new horizontal sizer for the top Left buttons row
        hSubSizer1 = wx.BoxSizer(wx.HORIZONTAL)
        # Add the first set of Position Control buttons for the RIGHT media control
        self.btnRL60 = wx.Button(self.pnl, -1, "< 60", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnRL60, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnRL10 = wx.Button(self.pnl, -1, "< 10", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnRL10, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnRL1 = wx.Button(self.pnl, -1, "< 1", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnRL1, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # Add the top row to the Vertical sizer
        vSizer.Add(hSubSizer1, 0)
        # Create a new horizontal sizer for the bottom Left buttons row
        hSubSizer2 = wx.BoxSizer(wx.HORIZONTAL)
        # Add the second set of Position Control buttons for the LEFT media control
        self.btnRLFifth = wx.Button(self.pnl, -1, "< 0.2", size=(btnWidth, 20))
        hSubSizer2.Add(self.btnRLFifth, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # This one is twice as big (plus the gap) as the others
        self.btnRLFrame = wx.Button(self.pnl, -1, _("< 1 frame"), size=(2 * btnWidth + 4, 20))
        hSubSizer2.Add(self.btnRLFrame, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # Add the bottom row to the Vertical sizer
        vSizer.Add(hSubSizer2, 0)
        # Add the vertical sizer to the whole row horizontal sizer
        hSizer.Add(vSizer, 0)
        
        # Add a Play button for the bottom media player
        self.btnPlay3 = wx.BitmapButton(self.pnl, -1, TransanaGlobal.GetImage(TransanaImages.Play), size=(48, 48))
        self.btnPlay3.SetLayoutDirection(wx.Layout_LeftToRight)
        # Bind the button to the OnPlay event handler
        self.btnPlay3.Bind(wx.EVT_BUTTON, self.OnPlay)
        # Add the play button to the sizer
        hSizer.Add(self.btnPlay3, 0, wx.LEFT | wx.RIGHT, 2)
        
        # Create a new Vertical sizer for the two right buttons rows
        vSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a new horizontal sizer for the top right buttons row
        hSubSizer1 = wx.BoxSizer(wx.HORIZONTAL)
        # Add the rest of the Position Control buttons for the RIGHT media control
        self.btnRR1 = wx.Button(self.pnl, -1, "1 >", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnRR1, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnRR10 = wx.Button(self.pnl, -1, "10 >", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnRR10, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnRR60 = wx.Button(self.pnl, -1, "60 >", size=(btnWidth, 20))
        hSubSizer1.Add(self.btnRR60, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # Add the top row to the Vertical sizer
        vSizer.Add(hSubSizer1, 0)
        # Create a new horizontal sizer for the bottom Right buttons row
        hSubSizer2 = wx.BoxSizer(wx.HORIZONTAL)
        # Add the second set of Position Control buttons for the LEFT media control
        # This one is twice as big (plus the gap) as the others
        self.btnRRFrame = wx.Button(self.pnl, -1, _("1 frame >"), size=(2 * btnWidth + 4, 20))
        hSubSizer2.Add(self.btnRRFrame, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        self.btnRRFifth = wx.Button(self.pnl, -1, "0.2 >", size=(btnWidth, 20))
        hSubSizer2.Add(self.btnRRFifth, 0, wx.LEFT | wx.RIGHT | wx.TOP, 2)
        # Add the bottom row to the Vertical sizer
        vSizer.Add(hSubSizer2, 0)
        # Add the vertical sizer to the whole row horizontal sizer
        hSizer.Add(vSizer, 0)
        
        # Bind the position change buttons to the position change event handler
        self.btnRL60.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnRL10.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnRL1.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnRLFifth.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnRLFrame.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnRRFrame.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnRRFifth.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnRR1.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnRR10.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        self.btnRR60.Bind(wx.EVT_BUTTON, self.OnShiftButton)
        # Add a spacer
        hSizer.Add((10, 1), 1)
        # Add the horizontal sizer to the main sizer
        mainBoxSizer.Add(hSizer, 0, wx.EXPAND | wx.BOTTOM, 5)

        # Create a new horizontal sizer
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer
        hSizer.Add((1, 1), 2)
        # Add a time counter for the left media file
        self.lTime = wx.StaticText(self.pnl, -1, "0:00:00.0")
        # Add the control to the horizontal sizer
        hSizer.Add(self.lTime, 0, wx.LEFT | wx.RIGHT, 3)
        # Add a spacer
        hSizer.Add((1, 1), 1)
        # Add a checkbox for control of the LEFT media player
        self.topLeft = wx.CheckBox(self.pnl, -1, " " + _("Focus / Cursor Keys Left"), style=wx.CHK_2STATE | wx.ALIGN_RIGHT)
        # Set the LEFT checkbox as checked by default
        self.topLeft.SetValue(True)
        # Add the checkbox to the horizontal sizer
        hSizer.Add(self.topLeft, 0)
        # Bind the checkbox to the checkbox event handler
        self.topLeft.Bind(wx.EVT_CHECKBOX, self.OnCheckBox)
        # Add a spacer
        hSizer.Add((1, 1), 5)
        # Add a checkbox for control of the LEFT media player
        self.topRight = wx.CheckBox(self.pnl, -1, " " + _("Focus / Cursor Keys Right"), style=wx.CHK_2STATE)
        # Add the checkbox to the horizontal sizer
        hSizer.Add(self.topRight, 0)
        # Bind the checkbox to the checkbox event handler
        self.topRight.Bind(wx.EVT_CHECKBOX, self.OnCheckBox)
        # Add a spacer
        hSizer.Add((1, 1), 1)
        # Add a time counter for the right media file
        self.rTime = wx.StaticText(self.pnl, -1, "0:00:00.0")
        # Add the control to the horizontal sizer
        hSizer.Add(self.rTime, 0, wx.LEFT | wx.RIGHT, 3)
        # Add a spacer
        hSizer.Add((1, 1), 2)
        # Add the horizontal sizer to the main sizer
        mainBoxSizer.Add(hSizer, 0, wx.EXPAND | wx.BOTTOM, 5)

        # Create a new horizontal sizer
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer
        hSizer.Add((10, 1), 0)
        # Add a label
        lbl = wx.StaticText(self.pnl, -1, _('Zoom:'))
        # Add the label to the zoom sizer
        hSizer.Add(lbl, 0)
        # Add the waveform zoom slider
        self.waveformZoom = wx.Slider(self.pnl, -1, int(self.waveformSpan / 200), 5, 1200, style = wx.SL_HORIZONTAL)
        # Add a Tool Tip to label the slider
        self.waveformZoom.SetToolTipString(_("Waveform Zoom"))
        # Bind the slider to the OnSlider Event Handler
        self.waveformZoom.Bind(wx.EVT_SCROLL, self.OnSlider)
        # Add the slider to the sizer
        hSizer.Add(self.waveformZoom, 13)
        # Add a spacer
        hSizer.Add((10, 1), 0)
        # Add the horizontal sizer to the main sizer
        mainBoxSizer.Add(hSizer, 0, wx.EXPAND | wx.BOTTOM, 5)

        # Create a new horizontal sizer
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer
        hSizer.Add((10, 1), 0)
        # Create the first Waveform control
        self.waveform1 = GraphicsControlClass.GraphicsControl(self.pnl, -1, wx.Point(1, 1), wx.Size(760, 140), (760, 140), visualizationMode=False, name='mp1')
        # Clear the graphic control
        self.waveform1.Clear()
        # initialize the start position
        self.start1 = 0
        # initialize the start position
        self.start2 = 0
        # Create the waveform for the first file, if needed.  If successful, the name of the WAV file is returned.
        self.waveFile1 = self.LoadWaveform(filename1)
        # If the waveform fails ...
        if self.waveFile1 == False:
            # ... display an error message ...
            print "Waveform FAIL for", filename1
            # .. and exit the Synchronize routine.
            return
        # Create the waveform for the second file, if needed.  If successful, the name of the WAV file is returned.
        self.waveFile2 = self.LoadWaveform(filename2)
        # If the waveform fails ...
        if self.waveFile2 == False:
            # ... display an error message ...
            print "Waveform FAIL for", filename2.encode('utf8')
            # .. and exit the Synchronize routine.
            return
        # Add the waveform to the sizer
        hSizer.Add(self.waveform1, 10)
        # Add a spacer
        hSizer.Add((10, 1), 0)
        # Add the horizonal sizer to the main sizer
        mainBoxSizer.Add(hSizer, 0, wx.EXPAND)

        # Create a sizer for the form buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer so that the offset control will align properly
        btnSizer.Add((4, 1), 0)

        # Add a TextCtrl for offset
        self.txtOffset = wx.TextCtrl(self, -1, Misc.time_in_ms_to_str(self.offset, True), style=wx.TE_RIGHT | wx.TE_READONLY)

        # Add the offset text to the button sizer
        btnSizer.Add(self.txtOffset, 0, wx.ALL, 6)
        # Add a horizontal spacer to the button sizer
        btnSizer.Add((0, 0), 1)
        # Create an OK button
        self.btnOK = wx.Button(self.pnl, wx.ID_OK, _("OK"))
        
        # Add the OK button to the button sizer
        btnSizer.Add(self.btnOK, 0, wx.ALIGN_RIGHT | wx.TOP | wx.RIGHT | wx.BOTTOM, 6)
        # Create a Cancel button
        self.btnCancel = wx.Button(self.pnl, wx.ID_CANCEL, _("Cancel"))
        # Add the Cancel button to the button sizer
        btnSizer.Add(self.btnCancel, 0, wx.ALIGN_RIGHT | wx.TOP | wx.RIGHT | wx.BOTTOM, 6)
        # Create a Help button
        self.btnHelp = wx.Button(self.pnl, -1, _("Help"))
        # Add the Help button to the button sizer
        btnSizer.Add(self.btnHelp, 0, wx.ALIGN_RIGHT | wx.TOP | wx.RIGHT | wx.BOTTOM, 6)
        # Bind a handler to the Help button
        self.btnHelp.Bind(wx.EVT_BUTTON, self.Help)
        # Add the button sizer to the main form sizer
        mainBoxSizer.Add(btnSizer, 0, wx.EXPAND)

        # Set the form focus to the Play Both button
        self.btnPlay2.SetFocus()

        # All controls on the form need to call the OnKeyDown and OnKeyUp methods to process cursor key presses
        wx.EVT_KEY_UP(self, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self, self.OnKeyDown)
        wx.EVT_KEY_UP(self.mp1, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.mp1, self.OnKeyDown)
        wx.EVT_KEY_UP(self.mp2, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.mp2, self.OnKeyDown)
        wx.EVT_KEY_UP(self.btnPlay1, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.btnPlay1, self.OnKeyDown)
        wx.EVT_KEY_UP(self.btnPlay2, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.btnPlay2, self.OnKeyDown)
        wx.EVT_KEY_UP(self.btnPlay3, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.btnPlay3, self.OnKeyDown)
        wx.EVT_KEY_UP(self.topRight, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.topRight, self.OnKeyDown)
        wx.EVT_KEY_UP(self.topLeft, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.topLeft, self.OnKeyDown)
        wx.EVT_KEY_UP(self.waveformZoom, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.waveformZoom, self.OnKeyDown)
        wx.EVT_KEY_UP(self.waveform1, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.waveform1, self.OnKeyDown)
        wx.EVT_KEY_UP(self.txtOffset, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.txtOffset, self.OnKeyDown)
        wx.EVT_KEY_UP(self.btnOK, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.btnOK, self.OnKeyDown)
        wx.EVT_KEY_UP(self.btnCancel, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.btnCancel, self.OnKeyDown)
        wx.EVT_KEY_UP(self.btnHelp, self.OnKeyUp)
        wx.EVT_KEY_DOWN(self.btnHelp, self.OnKeyDown)

        # Bind the form's OnSize event
        self.Bind(wx.EVT_SIZE, self.OnSize)

        # Bind the form's OnClose event
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # Call Thaw to allow the dialog to finally be drawn.
        self.Thaw()

        # Set the main box sizer as the Panel's Main Sizer
        self.pnl.SetSizer(mainBoxSizer)
        # Set the panel sizer as the dialog's main sizer
        self.SetSizer(pnlSizer)

        # Turn AutoLayout on
        self.SetAutoLayout(True)
        # Lay out the dialog
        self.Layout()
        # Center the dialog on the screen
        self.CenterOnScreen()

        # Now that this is all done, we can signal that the dialog has been built
        self.windowBuilt = True

        # Update the Waveforms on the Mac, as this doesn't seem to happen automatically
        if 'wxMac' in wx.PlatformInfo:
            wx.CallAfter(self.UpdateWaveforms)

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

    def OnKeyDown(self, event):
        """ Method to handle the initiation of a keypress """
        # Get the Key Code of the key that was pressed
        keyPressed = event.GetKeyCode()
        # We are only intercepting cursor left and cursor right.  Therefore,
        # if the was NOT cursor left or cursor right ...
        if not(keyPressed in [wx.WXK_LEFT, wx.WXK_RIGHT]):
            # ... then process the key press normally.
            event.Skip()

    def OnKeyUp(self, event):
        """ Method to handle the release of a keypress """
        # Get the Key Code of the key that was pressed
        keyPressed = event.GetKeyCode()
        # We are only intercepting cursor left and cursor right.  Therefore,
        # if the was cursor left or cursor right ...
        if keyPressed in [wx.WXK_LEFT, wx.WXK_RIGHT]:
            # If the Alt key is held down ...
            if event.AltDown():
                # ... we want to move 1/2 a second
                amtToMove = 500
            # if the Shift key is held down ...
            elif event.ShiftDown():
                # ... we want to move a whole second
                amtToMove = 1000
            # Otherwise ...
            else:
                # we want to move one frame, or approximately 33/1000ths of a second
                amtToMove = 33
            # If the LEFT checkbox is checked ...
            if self.topLeft.IsChecked():
                # ... if Cursor Left is pressed ...
                if keyPressed == wx.WXK_LEFT:
                    # ... we move Player 1 FORWARD the appropriate amount
                    self.mp1.SetCurrentVideoPosition(self.mp1.GetTimecode() + amtToMove)
                    self.pos1 += amtToMove
                # ... if Cursor Right is pressed ...
                else:
                    # ... we move Player 1 BACKWARDS the appropriate amount
                    self.mp1.SetCurrentVideoPosition(self.mp1.GetTimecode() - amtToMove)
                    self.pos1 -= amtToMove
            # If the RIGHT checkbox is checked ...
            else:
                # ... if Cursor Left is pressed ...
                if keyPressed == wx.WXK_LEFT:
                    # ... we move Player 2 FORWARD the appropriate amount
                    self.mp2.SetCurrentVideoPosition(self.mp2.GetTimecode() + amtToMove)
                    self.pos2 += amtToMove
                # ... if Cursor Right is pressed ...
                else:
                    # ... we move Player 2 BACKWARDS the appropriate amount
                    self.mp2.SetCurrentVideoPosition(self.mp2.GetTimecode() - amtToMove)
                    self.pos2 -= amtToMove
            # Update the Waveforms
            self.UpdateWaveforms()
            # Update the offset text box
            self.txtOffset.SetValue(Misc.time_in_ms_to_str(self.pos1 - self.pos2, True))
        # If the was NOT cursor left or cursor right ...
        else:
            # ... then process the key press normally.
            event.Skip()

    def OnCheckBox(self, event):
        """ Method for handling checkbox events """
        # If the LEFT checkbox triggers the event ...
        if event.GetId() == self.topLeft.GetId():
            # ... then the RIGHT checkbox needs to be updated.
            self.topRight.SetValue(not self.topLeft.GetValue())
        # If the RIGHT checkbox triggers the event ...
        elif event.GetId() == self.topRight.GetId():
            self.topLeft.SetValue(not self.topRight.GetValue())
        self.UpdateWaveforms()

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
                self.btnPlay1.SetBitmapLabel(TransanaGlobal.GetImage(TransanaImages.Play))
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
                self.btnPlay1.SetBitmapLabel(TransanaGlobal.GetImage(TransanaImages.Play))
        else:           # Player 1 was NOT playing before the button press ...
            if play1:   # The top button read "Play"
                self.btnPlay1.SetBitmapLabel(TransanaImages.Pause.GetBitmap())
        if isPlaying2:  # Player 2 was playing before the button press ...
            if play2:   # The bottom button read "Pause"
                self.btnPlay3.SetBitmapLabel(TransanaGlobal.GetImage(TransanaImages.Play))
        else:           # Player 2 was NOT playing before the button press ...
            if play2:   # The bottom button read "Play"
                self.btnPlay3.SetBitmapLabel(TransanaImages.Pause.GetBitmap())

        # If both players are NOW paused ...
        if (isPlaying1 and play1 and not isPlaying2) or \
           (isPlaying2 and play2 and not isPlaying1) or \
           (isPlaying1 and play1 and isPlaying2 and play2):
            # ... change the middle button to Play
            self.btnPlay2.SetLabel(_("Play Both"))
            # You can only EXIT the form if both players are paused!  (This makes the synch more accurate, and prevents Destroy problems.)
            self.btnOK.Enable(True)
            self.btnCancel.Enable(True)
        # If either player is playing ...
        else:
            # ... change the middle button to Pause ...
            self.btnPlay2.SetLabel(_("Pause Both"))
            # ... and disable the OK and Cancel buttons!
            self.btnOK.Enable(False)
            self.btnCancel.Enable(False)
        # Update the Waveforms
        self.UpdateWaveforms()

    def OnShiftButton(self, event):
        """ Methods for implementing the Media Player Position Shift buttons """
        # NOTE:  Media player position change was originally implemented using sliders.  This method is preferred
        #        because it is MUCH easier to move both media files precisely the same amount.
        
        # If Left Player 60 seconds earlier button is pressed ...
        if event.GetId() == self.btnLL60.GetId():
            # Time shift is negative 60 seconds
            amtToMove = -60000
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1
        # If Left Player 10 seconds earlier button is pressed ...
        elif event.GetId() == self.btnLL10.GetId():
            # Time shift is negative 10 seconds
            amtToMove = -10000
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1
        # If Left Player 1 second earlier button is pressed ...
        elif event.GetId() == self.btnLL1.GetId():
            # Time shift is negative 1 second
            amtToMove = -1000
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1
        # If Left Player 0.2 second later button is pressed ...
        elif event.GetId() == self.btnLLFifth.GetId():
            # Time shift is negative one-fifth of a second
            amtToMove = -200
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1
        # If Left Player 1 frame earlier button is pressed ...
        elif event.GetId() == self.btnLLFrame.GetId():
            # Time shift is negative 1 frame
            amtToMove = -33
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1
        # If Left Player 1 frame later button is pressed ...
        elif event.GetId() == self.btnLRFrame.GetId():
            # Time shift is positive 1 frame
            amtToMove = 33
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1
        # If Left Player 0.2 second later button is pressed ...
        elif event.GetId() == self.btnLRFifth.GetId():
            # Time shift is positive one-fifth second
            amtToMove = 200
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1
        # If Left Player 1 second later button is pressed ...
        elif event.GetId() == self.btnLR1.GetId():
            # Time shift is positive 1 second
            amtToMove = 1000
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1
        # If Left Player 10 seconds later button is pressed ...
        elif event.GetId() == self.btnLR10.GetId():
            # Time shift is positive 10 seconds
            amtToMove = 10000
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1
        # If Left Player 60 seconds later button is pressed ...
        elif event.GetId() == self.btnLR60.GetId():
            # Time shift is positive 60 seconds
            amtToMove = 60000
            # Media Player affected is player 1
            mp = self.mp1
            # Current position to be changed is position 1
            pos = self.pos1

        # If Right Player 60 seconds earlier button is pressed ...
        elif event.GetId() == self.btnRL60.GetId():
            # Time shift is negative 60 seconds
            amtToMove = -60000
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2
        # If Right Player 10 seconds earlier button is pressed ...
        elif event.GetId() == self.btnRL10.GetId():
            # Time shift is negative 10 seconds
            amtToMove = -10000
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2
        # If Right Player 1 second earlier button is pressed ...
        elif event.GetId() == self.btnRL1.GetId():
            # Time shift is negative 1 second
            amtToMove = -1000
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2
        # If Right Player 0.2 second later button is pressed ...
        elif event.GetId() == self.btnRLFifth.GetId():
            # Time shift is negative one-fifth of a second
            amtToMove = -200
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2
        # If Right Player 1 frame earlier button is pressed ...
        elif event.GetId() == self.btnRLFrame.GetId():
            # Time shift is negative 1 frame
            amtToMove = -33
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2
        # If Right Player 1 frame later button is pressed ...
        elif event.GetId() == self.btnRRFrame.GetId():
            # Time shift is positive 1 frame
            amtToMove = 33
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2
        # If Right Player 0.2 second later button is pressed ...
        elif event.GetId() == self.btnRRFifth.GetId():
            # Time shift is positive one-fifth second
            amtToMove = 200
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2
        # If Right Player 1 second later button is pressed ...
        elif event.GetId() == self.btnRR1.GetId():
            # Time shift is positive 1 second
            amtToMove = 1000
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2
        # If Right Player 10 seconds later button is pressed ...
        elif event.GetId() == self.btnRR10.GetId():
            # Time shift is positive 10 seconds
            amtToMove = 10000
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2
        # If Right Player 60 seconds later button is pressed ...
        elif event.GetId() == self.btnRR60.GetId():
            # Time shift is positive 60 seconds
            amtToMove = 60000
            # Media Player affected is player 2
            mp = self.mp2
            # Current position to be changed is position 2
            pos = self.pos2

        # Get the current position of the desired player
        curpos = mp.GetTimecode()
        # Get the total length of the media for the desired player
        totlen = mp.GetMediaLength()
        # If the desired move would take us to before the start of the media file ...
        if curpos + amtToMove < 0:
            # ... set the media position to 0
            mp.SetCurrentVideoPosition(0)
            pos = 0
        # If the desired move would take use to after the end of the media file ...
        elif curpos + amtToMove > totlen:
            # ... set the media position to just before the end of the file
            mp.SetCurrentVideoPosition(totlen - 10)
            pos = totlen - 10
        # If the desired move falls within the boundaries of the media files ...
        else:
            # ... then make the desired move
            mp.SetCurrentVideoPosition(curpos + amtToMove)
            pos += amtToMove
        # Update the waveforms to reflect the change
        self.UpdateWaveforms()
        # Update the offset text box
        self.txtOffset.SetValue(Misc.time_in_ms_to_str(self.pos1 - self.pos2, True))

    def OnSlider(self, event):
        """ Event handler for the Zoom Sliders """
        # Call the underlying parent event.  (Mac wasn't letting go of the handles!)
        event.Skip()
        # Default the selected media player to None
        mp = None
        # If the Zoom slider fired the event, if affects BOTH waveforms, but neither media player.
        if event.GetId() == self.waveformZoom.GetId():
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
                self.txtOffset.SetValue(Misc.time_in_ms_to_str(self.pos1 - self.pos2, True))
        
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

            # Update the offset text box, with thousandths of a second precision.
            self.txtOffset.SetValue(Misc.time_in_ms_to_str(self.pos1 - self.pos2, True))

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
        pass

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

    def LoadWaveform(self, mediaFile):
        """ Load a Waveform, based on the mediaFile, into the GraphicControl """
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
                progressDialog = WaveformProgress.WaveformProgress(self, prompt % (waveFilename1, mediaFile))
                # Tell the Waveform Progress Dialog to handle the audio extraction modally.
                progressDialog.Extract(mediaFile, waveFilename1)
                # Get the Error Log that may have been created
                errorLog = progressDialog.GetErrorMessages()
                # Okay, we're done with the Progress Dialog here!
                progressDialog.Destroy()
                # If the user cancelled the audio extraction ...
                if (len(errorLog) == 1) and (errorLog[0] == 'Cancelled'):
                    # ... signal that the WAV file was NOT created!
                    dllvalue = 1  
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
        # If the form is complete ...
        if self.windowBuilt:
            # Clear the waveform
            self.waveform1.Clear()
            # Get the current position from the media players
            start1 = self.mp1.GetTimecode()
            start2 = self.mp2.GetTimecode()
            # Display both media player positions
            self.lTime.SetLabel(Misc.time_in_ms_to_str(start1))
            self.rTime.SetLabel(Misc.time_in_ms_to_str(start2))
            # if the current position cursor should fall in the right half of the graphic ...
            if start1 > self.waveformSpan / 2:
                # ... adjust start so the cursor will fall in the middle of the graphic
                start1 -= self.waveformSpan / 2
            # Remember the current start position
            self.start1 = start1
            # if the current position cursor should fall in the right half of the graphic ...
            if start2 > self.waveformSpan / 2:
                # ... adjust start so the cursor will fall in the middle of the graphic
                start2 -= self.waveformSpan / 2
            # Remember the current start position
            self.start2 = start2
            # If the second file starts LATER ...
            if start2 >= start1:
                # ... then we base the waveform on the position of the SECOND file ...
                start = start2
                # ... the FIRST FILE offset is file 2 position minus file 1 position ...
                diff1 = start2 - start1
                # ... and the SECOND FILE offset is 0.
                diff2 = 0
            # If the first file starts later ...
            else:
                # ... then we base the waveform on the position of the FIRST file ...
                start = start1
                # ... the FIRST FILE offset is 0.
                diff1 = 0
                # ... and the SECOND FILE offset is file 1 position minus file 2 position ...
                diff2 = start1 - start2
            
            # If topLeft is checked ...
            if self.topLeft.IsChecked():
                # Determine the file names so that LEFT (Red) will be on top
                filenames = [{'filename' : self.waveFile1, 'offset' : diff1, 'length' : self.mp1.GetMediaLength()},
                             {'filename' : self.waveFile2, 'offset' : diff2, 'length' : self.mp2.GetMediaLength()}]
                # Determine the color order
                waveformColors = (wx.BLUE, wx.RED)
                # Load the new graphic into the Waveform Control.  (Passing ":memory:" rather than a filename causes WaveformGraphic to return a Bitmap object rather than saving it to a file!)
                self.waveform1.backgroundImage = WaveformGraphic.WaveformGraphicCreate(filenames, ':memory:', start, self.waveformSpan, self.waveform1.GetSize(), colors = waveformColors)
            # if topLeft is NOT checked (i.e. topRight is checked) ...
            else:
                # Determine the file names so that RIGHT (Blue) will be on top
                filenames = [{'filename' : self.waveFile2, 'offset' : diff2, 'length' : self.mp2.GetMediaLength()},
                             {'filename' : self.waveFile1, 'offset' : diff1, 'length' : self.mp1.GetMediaLength()}]
                # Determine the color order
                waveformColors = (wx.RED, wx.BLUE)
                # Load the new graphic into the Waveform Control.  (Passing ":memory:" rather than a filename causes WaveformGraphic to return a Bitmap object rather than saving it to a file!)
                self.waveform1.backgroundImage = WaveformGraphic.WaveformGraphicCreate(filenames, ':memory:', start, self.waveformSpan, self.waveform1.GetSize(), colors = waveformColors)
            # Determine the current cursor position
            pos = ((float(self.mp1.GetTimecode() - start1)) / self.waveformSpan)
            # Draw the Waveform cursor
            self.waveform1.DrawCursor(pos)
            # If it has been more than 1 second since the last visualization update ...
            if time.time() - self.waveform1.lastUpdateTime > 1.0:
                # ... force a redraw now.  (This makes media playback more choppy.)
                self.waveform1.Redraw()

    def OnSizeChange(self):
        """ Size Change method (not event handler!) for the Synchronize Dialog """
        # If the dialog has been completed ...
        if self.windowBuilt:
            # Resize the Waveform
            self.ResizeWaveforms()
            # Lay out the form, so the sizer adjusts everything
            self.Layout()
            # Refresh the display to clean things up.
            self.Refresh()
            # After this is DONE, resize the videos to restore the proper aspect ratios for the media files.
            wx.CallAfter(self.ResetVideoSizes)

    def ResetVideoSizes(self):
        """ Resize the Media Files to maintain their aspect ratios """
        # Get the original media sizes for both media controls
        (size1x, size1y) = self.mp1.movie.GetBestSize()
        (size2x, size2y) = self.mp2.movie.GetBestSize()
        # None of these values should be 0 (as for audio files) or we get a division by zero error
        if size1x == 0:
            size1x = 4
        if size1y == 0:
            size1y = 3
        if size2x == 0:
            size2x = 4
        if size2y == 0:
            size2y = 3
        # Determine the desired width of each media control, which is half of the form's width, adjusted for 40 pixels of framing
        x = (self.GetSize()[0] / 2) - 40
        # Determine the desired height of each media control, which is the larger of the two media control heights, adjusted for
        # aspect ratio
        y = max(((x * float(size2y)) / float(size2x)), ((x * float(size1y)) / float(size1x)))
        # If the media player TOP plus the desired height would over-run the Play buttons, which need to be below the
        # media player area ...
        if self.mp1.GetPosition()[1] + y + 5 > self.btnPlay1.GetPosition()[1]:
            # ... re-calculate the desired height to be the maximum height we have space for
            y = self.btnPlay1.GetPosition()[1] - self.mp1.GetPosition()[1] - 10
            # ... and re-calculate the desired width of media player 1 based on that height to maintain aspect ratio
            x = (float(size1x) * y) / float(size1y)
            # Now resize the media player frame and the media itself to fit
            self.mp1.SetSize((x, y))
            self.mp1.movie.SetSize((x, y))
            # ... and re-calculate the desired width of media player 2 based on that height to maintain aspect ratio
            x = (float(size2x) * y) / float(size2y)
            # Now resize the media player frame and the media itself to fit
            self.mp2.SetSize((x, y))
            self.mp2.movie.SetSize((x, y))

        else:
            # Determine the aspect ratio of the original media file in player 1
            aspectRatio = float(size1y) / float(size1x)
            # Determine the proper height based on aspect ratio and width, then resize teh media player frame
            # and the media itself to fit
            self.mp1.SetSize((x, (x * aspectRatio)))
            self.mp1.movie.SetSize((x, (x * aspectRatio)))
            # Determine the aspect ratio of the original media file in player 2
            aspectRatio = float(size2y) / float(size2x)
            # Determine the proper height based on aspect ratio and width, then resize teh media player frame
            # and the media itself to fit
            self.mp2.SetSize((x, (x * aspectRatio)))
            self.mp2.movie.SetSize((x, (x * aspectRatio)))

    def ResizeWaveforms(self):
        """ Method useful in resizing and redrawing the Waveform diagrams """
        # If the dialog has been completely rendered ...
        if self.windowBuilt:
            # update the dialog
            self.Update()
            # Get the size of the Waveform Control
            waveformSize1 = self.waveform1.GetSize()
            # Get the position of the Waveform Control
            waveformPos1 = self.waveform1.GetPosition()
            # Resize the Waveform to conform to the size of the Control.  This forces the embedded diagram to resize.
            self.waveform1.SetDim(waveformPos1[0], waveformPos1[1], waveformSize1[0]-6, waveformSize1[1]-6)
            # Finally, update the Waveforms
            self.UpdateWaveforms()

    def OnSize(self, event):
        """ Size Change event handler for the Synchronize Dialog """
        # When the dust has settled, call the OnSizeChange method
        wx.CallAfter(self.OnSizeChange)

    def OnClose(self, event):
        """ Handle Window Close event """
        # Stop both media players.  Otherwise, the audio can continue and Transana will freeze if the form is closed
        # unexpectedly
        self.mp1.Stop()
        self.mp2.Stop()
        # Call the Frame's Close event
        event.Skip()
        
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
