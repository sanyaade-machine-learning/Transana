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

"""This module implements the Visualization class as part of the Visualization
component."""

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>, Nathaniel Case, Rajas Sambhare'


import wx
import TransanaGlobal
import TransanaConstants
import GraphicsControlClass
import WaveformGraphic
import WaveformProgress       # Waveform Creation Progress Dialog, used in Wave Extraction Callback function
import Dialogs
import Misc
import os
import sys
import string
import ctypes                 # used to access wceraudio DLL/Shared Library

class VisualizationWindow(wx.Dialog):
    """This class encompasses the others into a single wxPython window,
    and provides the primary interface for the Control objects.  This
    object will be passed a memory-mapped wave file as input for the
    audio data to be displayed."""

    def __init__(self, parent):
        """Initialize a VisualizationWindow object."""
        # Positioning is different on Mac and Windows, requiring subtle adjustments in this window
        if '__WXMAC__' in wx.PlatformInfo:
            topAdjust = 0
        else:
            topAdjust = 2
            
        wx.Dialog.__init__(self, parent, -1, _('Visualization'), pos=self.__pos(), size=self.__size(), style=wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS)
        # Set "Window Variant" to small only for Mac to make fonts match better
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # print "Visualization Window:", self.__pos(), self.__size()

        self.SetBackgroundColour(wx.WHITE)
        self.SetSizeHints(250, 100)
        (width, height) = self.GetSize()

        # The ControlObject handles all inter-object communication, initialized to None
        self.ControlObject = None            

        # startPoint and endPoint are used in the process of drag-selecting part of a waveform.
        self.startPoint = 0
        self.endPoint = 0

        # waveformLowerLimit and waveformUpperLimit contain the lower and upper bounds (in milliseconds) of the portion of the waveform
        # that is currently displayed.  These are crucial in all positioning calculations.
        self.waveformLowerLimit = 0
        self.waveformUpperLimit = 0

        # redrawWhenIdle signals that the Waveform picture needs to be drawn when the CPU has time
        self.redrawWhenIdle = False
        # If the video position needs to be set after a waveform redraw, set this value
        self.resetVideoPosition = 0

        # zoomInfo holds information about Zooms, to allow zoom-out in steps matching zoom-ins
        self.zoomInfo = []                   

        # Add the Toolbar at the bottom of the screen
        # NOTE:  Because of the way XP handles the screen, we'll start at the bottom and work our way up!
        lay = wx.LayoutConstraints()
        lay.height.Absolute(22)
        lay.left.SameAs(self, wx.Left, 0)
        lay.bottom.SameAs(self, wx.Bottom, 0)
        lay.right.SameAs(self, wx.Right, 0)        
        self.toolbar = wx.Panel(self, -1, wx.DefaultPosition, wx.DefaultSize)  # , wx.Point(0, height-50), wx.Size(int(width - 8), 24))
        self.toolbar.SetConstraints(lay)

        # Add GUI elements to the Toolbar

        # Add Zoom In Button
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust)
        lay.left.SameAs(self.toolbar, wx.Left, 2)
        lay.width.AsIs()
        lay.height.AsIs()
        self.zoomIn = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_ZOOMIN, wx.Bitmap('images/ZoomIn.xpm', wx.BITMAP_TYPE_XPM))
        self.zoomIn.SetConstraints(lay)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_ZOOMIN, self.OnZoomIn)

        # Add Zoom Out Button
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust)
        lay.left.RightOf(self.zoomIn, 2)
        lay.width.AsIs()
        lay.height.AsIs()
        self.zoomOut = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_ZOOMOUT, wx.Bitmap('images/ZoomOut.xpm', wx.BITMAP_TYPE_XPM))
        self.zoomOut.SetConstraints(lay)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_ZOOMOUT, self.OnZoomOut)

        # Add Zoom to 100% Button
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust)
        lay.left.RightOf(self.zoomOut, 2)
        lay.width.AsIs()
        lay.height.AsIs()
        self.zoom100 = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_ZOOM100, wx.Bitmap('images/Zoom100.xpm', wx.BITMAP_TYPE_XPM))
        self.zoom100.SetConstraints(lay)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_ZOOM100, self.OnZoom100)

        # Place line separating Zoom buttons from Position section
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust - 1)
        lay.left.RightOf(self.zoom100, 6)
        lay.width.AsIs()
        lay.height.AsIs()
        separator1 = wx.StaticLine(self.toolbar, -1, size=wx.Size(2, 22))
        separator1.SetConstraints(lay)

        # Add "Time" label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust + 3)
        lay.left.RightOf(separator1, 6)
        lay.width.AsIs()
        lay.height.AsIs()
        self.lbl_Time = wx.StaticText(self.toolbar, -1, _("Time:"), size=(38, 16))
        self.lbl_Time.SetConstraints(lay)

        # Add "Time" Time label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust + 3)
        lay.left.RightOf(self.lbl_Time, 8)
        lay.width.AsIs()
        lay.height.AsIs()
        self.lbl_Time_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        self.lbl_Time_Time.SetConstraints(lay)

        # Place line separating Zoom buttons from Position section
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust - 1)
        lay.left.RightOf(self.lbl_Time_Time, 6)
        lay.width.AsIs()
        lay.height.AsIs()
        separator2 = wx.StaticLine(self.toolbar, -1, size=wx.Size(2, 22))
        separator2.SetConstraints(lay)

        # Add "Current" label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust)
        lay.left.RightOf(separator2, 6)
        lay.width.Absolute(80)
        lay.height.Absolute(20)
        self.btn_Current = wx.Button(self.toolbar, TransanaConstants.VISUAL_BUTTON_CURRENT, _("Current:"))
        self.btn_Current.SetConstraints(lay)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_CURRENT, self.OnCurrent)

        # Add "Current" Time label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust + 3)
        lay.left.RightOf(self.btn_Current, 8)
        lay.width.AsIs()
        lay.height.AsIs()
        self.lbl_Current_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        self.lbl_Current_Time.SetConstraints(lay)

        # Add "Selected" label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust)
        lay.left.RightOf(self.lbl_Current_Time, 20)
        lay.width.Absolute(80)
        lay.height.Absolute(20)
        self.btn_Selected = wx.Button(self.toolbar, TransanaConstants.VISUAL_BUTTON_SELECTED, _("Selected:"))
        self.btn_Selected.SetConstraints(lay)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_SELECTED, self.OnSelected)

        # Add "Selected" Time label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust + 3)
        lay.left.RightOf(self.btn_Selected, 8)
        lay.width.AsIs()
        lay.height.AsIs()
        self.lbl_Selected_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        self.lbl_Selected_Time.SetConstraints(lay)

        # Add "Total" Time label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust + 3)
        # Adjust for Mac's Window Resize Handle, which covers the decimal digit without this code
        if '__WXMAC__' in wx.PlatformInfo:
            lay.right.SameAs(self.toolbar, wx.Right, 20)
        else:
            lay.right.SameAs(self.toolbar, wx.Right, 8)
        lay.width.AsIs()
        lay.height.AsIs()
        self.lbl_Total_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        self.lbl_Total_Time.SetConstraints(lay)

        # Add "Total" label
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.toolbar, wx.Top, topAdjust + 3)
        lay.right.LeftOf(self.lbl_Total_Time, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        self.lbl_Total = wx.StaticText(self.toolbar, -1, _("Total:"))
        self.lbl_Total.SetConstraints(lay)

        self.toolbar.Layout()
        self.toolbar.SetAutoLayout(True)

        # We need to know the height of the Window Header to adjust the size of the Graphic Area
        headerHeight = self.GetSizeTuple()[1] - self.GetClientSizeTuple()[1]

        # Add the Timeline Panel, which holds the time line and scale information
        lay = wx.LayoutConstraints()
        lay.bottom.Above(self.toolbar, 0)
        lay.left.SameAs(self.toolbar, wx.Left, 0)
        lay.width.AsIs()
        lay.height.AsIs()
        self.timeline = wx.Panel(self, -1, wx.Point(0, height-44-headerHeight), wx.Size(int(width - 6), 24), style=wx.SUNKEN_BORDER)
        self.timeline.SetConstraints(lay)

        # The waveform is held in a GraphicsControlClass object that handles all the waveform display.
        # At this time, the GraphicsControlClass is not aware of Sizers or Constraints
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self, wx.Top, 0)
        lay.bottom.Above(self.timeline, 0)
        lay.left.SameAs(self.toolbar, wx.Left, 0)
        lay.right.SameAs(self.toolbar, wx.Right, 0)

        self.waveform = GraphicsControlClass.GraphicsControl(self, -1, wx.Point(1, 1), wx.Size(int(width-12), int(height-48-headerHeight)), (width-14, height-52-headerHeight), transanaMode=True)
        self.waveform.SetConstraints(lay)

        # Stick some initial "0" values in the timeline
        self.draw_timeline_zero()

        # Set the focus on the Waveform widget
        self.waveform.SetFocus()
        
        self.Layout()
        self.SetAutoLayout(True)

        # We need to adjust the screen position on the Mac.  I don't know why.
        if "__WXMAC__" in wx.PlatformInfo:
            pos = self.GetPosition()
            self.SetPosition((pos[0]-20, pos[1]-25))
        
        #print "init()"
        #print 'headerHeight = %d, graphic height - %d' % (headerHeight, height-44-headerHeight)
        #print 'Window:', self.GetPosition(), self.GetSize(), self.GetClientSize()
        #print 'GraphicControlClass:', (0, 1), (int(width-8), int(height-44-headerHeight)), (width-12, height-48-headerHeight)
        #print 'timeline:', self.timeline.GetPosition(), self.timeline.GetSize()
        #print 'toolbar:', self.toolbar.GetPosition(), self.toolbar.GetSize()
        #print
        

        # GraphicsControlClass is not Sizer or Constraints aware, so we'll to screen positioning the old fashion way, with a
        # Resize Event
        wx.EVT_SIZE(self, self.OnSize)

        # Idle event (draws when idle to prevent multiple redraws while resizing, which are too slow)
        wx.EVT_IDLE(self, self.OnIdle)

        # Let's also capture key presses so we can control video playback during transcription
        # NOTE that we assign this event to the waveform, not to self.
        wx.EVT_KEY_DOWN(self.waveform, self.OnKeyDown)

    def OnSize(self, event):
        """ Handles widget positioning on Resize Event """
        # Determine Frame Size
        (width, height) = self.GetSize()
        # If Waveform Quickload is NOT selected, redraw the Waveform from the Wave File to the appropriate (new) size
        if not(TransanaGlobal.configData.waveformQuickLoad):
            self.redrawWhenIdle = True
        
        (left, top) = self.GetPositionTuple()
        self.ControlObject.UpdateWindowPositions('Visualization', width + left, YUpper = height + top)

        # Determine NEW Frame Size
        (width, height) = self.GetSize()

        # We need to know the height of the Window Header to adjust the size of the Graphic Area
        headerHeight = self.GetSizeTuple()[1] - self.GetClientSizeTuple()[1]

        # Set size of Waveform Window (Do this first, or the timeline is not scaled properly.)
        self.waveform.SetDim(1, 1, width-12, height-49-headerHeight)
        # Position toolbar Panel
        self.toolbar.SetDimensions(0, height-24-headerHeight, width-8, 22)

        # Draw the appropriate time line
        self.draw_timeline(self.waveformLowerLimit, self.waveformUpperLimit - self.waveformLowerLimit)

        # Position timeline Panel
        self.timeline.SetDimensions(0, height-46-headerHeight, width-6, 24)

        #print "OnSize()"
        #print 'headerHeight = %d, graphic height - %d' % (headerHeight, height-44-headerHeight)
        #print 'Window:', self.GetPosition(), self.GetSize(), self.GetClientSize()
        #print 'GraphicControlClass:', (0, 1), (int(width-8), int(height-44-headerHeight)), (width-12, height-48-headerHeight)
        #print 'timeline:', self.timeline.GetPosition(), self.timeline.GetSize()
        #print 'toolbar:', self.toolbar.GetPosition(), self.toolbar.GetSize()
        #print


    def OnIdle(self, event):
        """ Use Idle Time to handle the drawing in this control """
        # Check to see if the waveform control needs to be redrawn
        if self.redrawWhenIdle:
            # Create the appropriate Waveform Graphic
            if self.waveformFilename != '':
                try:
                    WaveformGraphic.WaveformGraphicCreate(self.waveFilename, self.waveformFilename, self.ControlObject.VideoStartPoint, self.ControlObject.GetMediaLength(), self.waveform.canvassize, True)
                    # Load the new graphic into the Waveform Control
                    self.waveform.LoadFile(self.waveformFilename)
                    # Determine NEW Frame Size
                    (width, height) = self.GetSize()
                    # Draw the TimeLine values
                    self.draw_timeline(self.ControlObject.VideoStartPoint, self.ControlObject.GetMediaLength())
                except:
                    # NO Error Message, as it disrupt the program flow if the user chooses not to waveform
                    # Dialogs.ErrorDialog(self, _('Waveform Graphic creation error for file:\n%s\n%s.') % (self.waveFilename, self.waveformFilename)).ShowModal()
                    self.waveformFilename = ''
                    self.ClearVisualization()
                    
            # Remove old Waveform Selection and Cursor data
            self.waveform.ClearTransanaSelection()

            # Signal that the redraw is complete, and does not need to be done again until this flag is altered.
            self.redrawWhenIdle = False
        
            # resetVideoPosition indicates where the video cursor should be set following a Waveform Repaint
            if self.resetVideoPosition > 0:
                # Set the Video StartPoint, as it may have been wiped out previously
                self.ControlObject.SetVideoStartPoint(self.resetVideoPosition)
                # Set the Visualization Object's startPoint too.
                self.startPoint = self.resetVideoPosition
                # Make sure the Position Cursor is drawn on the Waveform
                self.UpdatePosition(self.resetVideoPosition)
                # Clear the resetVideoPosition variable, as we're done.
                self.resetVideoPosition = 0

    def OnKeyDown(self, event):
        """ Captures Key Events to allow this window to control video playback during transcription. """
        if event.ControlDown():
            try:
                c = event.GetKeyCode()
                
                # Ctrl-T -- Insert Time Code
                if chr(c) == "T":
                    self.OnCurrent(event)
                    
                # Ctrl-S -- Start / Pause with setback
                elif chr(c) == "S":
                    self.ControlObject.PlayPause(1)
                    
                # Ctrl-D -- Start / Pause without setback
                elif chr(c) == "D":
                    self.ControlObject.PlayPause(0)

                # Ctrl-A -- Rewind video by 10 seconds
                elif chr(c) == "A":
                    vpos = self.ControlObject.GetVideoPosition()
                    self.ControlObject.SetVideoStartPoint(vpos-10000)
                    # Play should always be initiated on Ctrl-A
                    self.ControlObject.Play(0)

                # CTRL-F -- Advance video by 10 seconds
                elif chr(c) == "F":
                    vpos = self.ControlObject.GetVideoPosition()
                    self.ControlObject.SetVideoStartPoint(vpos+10000)
                    # Play should always be initiated on Ctrl-F
                    self.ControlObject.Play(0)

            except:
                pass

    def OnZoomIn(self, event):
        """ Zoom in on a portion of the Waveform Diagram """
        if ((self.startPoint != 0) or (self.endPoint != 0)) and (self.startPoint != self.endPoint):
            # Keep track of the new position.  This allows the user to zoom back out in the same steps used to zoom in
            self.zoomInfo.append((int(self.startPoint), int(self.endPoint - self.startPoint)))
            # Limit video playback to the selected part of the media by setting the VideoStartPoint and VideoEndPoint in the
            # Control Object
            self.ControlObject.SetVideoSelection(self.startPoint, self.endPoint)

            # Draw the TimeLine values
            self.draw_timeline(self.ControlObject.VideoStartPoint, self.ControlObject.GetMediaLength())
            # assign a temporary filename for the Waveform Graphic
            # self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.bmp'
            self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.png'
            # Signal that the Waveform Graphic should be updated
            self.redrawWhenIdle = True
            
    def OnZoomOut(self, event):
        """ Zoom out on the Waveform Diagram """
        if len(self.zoomInfo) > 1:
            # Remove the last entry from zoomInfo
            self.zoomInfo = self.zoomInfo[:-1]

            # Signal that the video position needs to be reset after the waveform is drawn
            self.resetVideoPosition = self.ControlObject.GetVideoPosition()

            # Change the start and end points to the last values
            self.ControlObject.SetVideoSelection(self.zoomInfo[-1][0], self.zoomInfo[-1][1] + self.zoomInfo[-1][0])

            # Draw the TimeLine values
            self.draw_timeline(self.ControlObject.VideoStartPoint, self.ControlObject.GetMediaLength())
            
            # assign a temporary filename for the Waveform Graphic
            # self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.bmp'
            self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.png'
            self.redrawWhenIdle = True
            # Clear the Start and End points of the Visualization Selection
            self.startPoint = 0
            self.endPoint = 0

    def OnZoom100(self, event):
        """ Zoom all the way out on the Waveform Diagram """
        # Reset the zoomInfo to the original values
        self.zoomInfo = [self.zoomInfo[0]]

        # Signal that the video position needs to be reset after the waveform is drawn
        self.resetVideoPosition = self.ControlObject.GetVideoPosition()

        # Change the start and end points
        self.ControlObject.SetVideoSelection(self.zoomInfo[0][0], self.zoomInfo[0][1] + self.zoomInfo[0][0])

        # Clear the Start and End points of the Visualization Selection
        self.startPoint = 0
        self.endPoint = 0

        # Draw the TimeLine values
        self.draw_timeline(self.ControlObject.VideoStartPoint, self.ControlObject.GetMediaLength())

        if (self.ControlObject.VideoStartPoint == 0) and (self.ControlObject.VideoEndPoint == 0):
            # Load the original waveform
            self.redrawWhenIdle = True
        else:
            # assign a temporary filename for the Waveform Graphic
            # self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.bmp'
            self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.png'
            self.redrawWhenIdle = True


    def OnCurrent(self, event):
        """ Click the "Current" button to place a time code in the Transcript """
        self.ControlObject.InsertTimecodeIntoTranscript()

    def OnSelected(self, event):
        """ Click the "Selected" button to place two time codes in the Transcript with the time gap in between """
        if self.endPoint - self.startPoint > 0.1:
            self.ControlObject.InsertSelectionTimecodesIntoTranscript(self.startPoint, self.endPoint)

        
# Public methods
    def ClearVisualization(self):
        """Clear the display."""
        # Clear zoom level information
        self.zoomInfo = []
        # Clear wave file and waveform file names
        self.waveFilename = ''
        self.waveformFilename = ''
        # Clear the waveform itself
        self.waveform.Clear()
        # Remove old Waveform Selection and Cursor data
        self.waveform.ClearTransanaSelection()
        # Clear the Time Line
        self.draw_timeline_zero()
        # Clear all time labels
        self.lbl_Time_Time.SetLabel(Misc.time_in_ms_to_str(0))
        self.lbl_Selected_Time.SetLabel(Misc.time_in_ms_to_str(0))
        self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(0))
        self.lbl_Total_Time.SetLabel(Misc.time_in_ms_to_str(0))
        # Signal that things should be redrawn
        self.redrawWhenIdle = True

    def ClearVisualizationSelection(self):
        """ Clear the Selection Box from the Visualization Window """
        self.waveform.ClearTransanaSelection()

    def load_image(self, filename, mediaStart, mediaLength):
        """ Causes the proper visualization to be displayed in the Visualization Window when a Video File is loaded. """
        # To start with, initialize the data structure that holds information about Zooms
        self.zoomInfo = [(self.ControlObject.VideoStartPoint, self.ControlObject.VideoEndPoint)]

        # Remember the original File Name that is passed in
        originalFilename = filename
        # Separate path and filename
        (path, filename) = os.path.split(filename)
        # break filename into root filename and extension
        (filenameroot, extension) = os.path.splitext(filename)
        # Build the correct filename for the Wave File
        self.waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
        # Build the correct filename for the Waveform Graphic
        # self.waveformFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.bmp')
        self.waveformFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.png')

        # print "Filename = %s, WaveFileName = %s" % (originalFilename, self.waveFilename)

        # Remove old Waveform Selection and Cursor data
        self.waveform.ClearTransanaSelection()
        # If we are showing the whole waveform and the Waveform graphic exists and QuickLoad is enabled, load the graphic from the file.
        if not(os.path.exists(self.waveformFilename) and TransanaGlobal.configData.waveformQuickLoad):
            # Create a Wave File if none exists!
            if not(os.path.exists(self.waveFilename)):
                # Politely ask the user to create the waveform
                dlg = wx.MessageDialog(self, _("No wave file exists.  Would you like to create one now?"), _("Transana Wave File Creation"), wx.YES_NO | wx.ICON_QUESTION | wx.CENTRE)
                # Check the user's response to the dialog. 
                if dlg.ShowModal() == wx.ID_YES:
                    try:
                        if not os.path.exists(TransanaGlobal.configData.visualizationPath):
                            os.makedirs(TransanaGlobal.configData.visualizationPath)
                        # Build the progress box's label
                        label = _("Extracting %s\nfrom %s") % (self.waveFilename, originalFilename)
                        # If the user accepts, create and display the Progress Dialog
                        progressDialog = WaveformProgress.WaveformProgress(self, self.waveFilename, label)
                        # Problems arise if this is not Modal!  Problems arise if it is!!  Hmmmm.
                        # The compromise is to add the wx.STAY_ON_TOP style, which will at least keep the progress bar
                        # on top, even if it doesn't fully prevent problematic events from being initiated in other windows.
                        progressDialog.Show()

                        # Set the Extraction Parameters to produce small WAV files
                        bits = 8           # Use 8 for the bit rate (Legal values are 8 and 16)
                        decimation = 16    # Use the highest level of decimation (Legal values are 1, 2, 4, 8, and 16)
                        mono = 1           # 1 = mono, 0 = stereo

                        # Define callback function
                        # NOTE:  This is a BOGUS Callback until DLL parameter issues can be resolved.
                        # Right now, the DLL returns doubles, but ctypes only allows int and string pointer returns.
                        # Thus, this routine gets BOGUS, untranslatable data back from the DLL.
                        def Progress(percent, seconds):
                            return progressDialog.Update(percent, seconds)

                        # Create the data structure needed by ctypes to pass the callback function as a Pointer to the DLL/Shared Library
                        callback = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_int, ctypes.c_int)
                        # Define a pointer to the callback function
                        callbackFunction = callback(Progress)

                        # Call the wcerAudio DLL/Shared Library's ExtractAudio function
                        if (os.name == "nt"):
                            dllvalue = ctypes.cdll.wceraudio.ExtractAudio(originalFilename, self.waveFilename, bits, decimation, mono, callbackFunction)
                        else:
                            wceraudio = ctypes.cdll.LoadLibrary("wceraudio.dylib")
                            dllvalue = wceraudio.ExtractAudio(originalFilename, self.waveFilename, bits, decimation, mono, callbackFunction)

                        if dllvalue != 0:
                            try:
                                os.remove(self.waveFilename)
                            except:
                                errordlg = Dialogs.ErrorDialog(self, _('Unable to create waveform for file "%s"\nError Code: %s') % (originalFilename, dllvalue))
                                errordlg.ShowModal()
                                errordlg.Destroy()

                        # Close the Progress Dialog when the DLL call is complete
                        progressDialog.Close()
                        
                    except:
                        errordlg = Dialogs.ErrorDialog(self, _('Unable to create Waveform Directory.\n%s\n%s') % (sys.exc_info()[0], sys.exc_info()[1]))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        dllvalue = 1  # Signal that the WAV file was NOT created!                        
                else:
                    # User declined to create the WAV file now
                    dllvalue = 1  # Signal that the WAV file was NOT created!
                    # Remove whatever image might be there now.  First, clear the file names.
                    self.waveFilename = ''
                    self.waveformFilename = ''
                    # Then clear the waveform itself
                    self.waveform.Clear()
                # Destroy the Dialog that asked to create the Wave file    
                dlg.Destroy()
                # If the user said no or there was a problem with Wave Extraction ...
                if dllvalue != 0:
                    self.waveformFilename = ''  # Signal that there is NO waveform File

        # If this is a Clip, we need to create a temporary Waveform to show only the portion we need to see
        if (mediaStart != 0) and (self.waveformFilename != '') and (os.path.exists(self.waveformFilename)):

            # assign a temporary filename for the Waveform Graphic
            # self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.bmp'
            self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.png'

        # Signal the need to draw the Waveform
        self.redrawWhenIdle = True

        # Show the media position in the Current Time label
        self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(mediaStart))

        # Draw the TimeLine values
        self.draw_timeline(mediaStart, mediaLength)
        

    def draw_timeline(self, mediaStart, mediaLength):
        def GetScaleIncrements(MediaLength):
            # The general rule is to try to get logical interval sizes with 8 or fewer time increments.
            # You always add a bit (20% at the lower levels) of the time interval to the MediaLength
            # because the final time is placed elsewhere and we don't want overlap.
            # This routine covers from 1 second to 18 hours in length.
        
            # media Length of 9 seconds or less = 1 second intervals  
            if MediaLength < 9001: 
                Num = int(round((MediaLength + 200) / 1000.0))
                Interval = 1000
            # media length of 18 seconds or less = 2 second intervals 
            elif MediaLength < 18001:
                Num = int(round((MediaLength + 400) / 2000.0))
                Interval = 2000
            # media length of 30 seconds or less = 5 second intervals 
            elif MediaLength < 30001:
                Num = int(round((MediaLength + 2000) / 5000.0))
                Interval = 5000
            # media length of 50 seconds or less = 5 second intervals 
            elif MediaLength < 50001:
                Num = int(round((MediaLength + 1000) / 5000.0))
                Interval = 5000
            # media Length of 1:30 or less = 10 second intervals      
            elif MediaLength < 90001: 
                Num = int(round((MediaLength + 2000) / 10000.0))
                Interval = 10000
            # media length of 2:50 or less = 20 second intervals       
            elif MediaLength < 160001:
                Num = int(round((MediaLength + 4000) / 20000.0))
                Interval = 20000
            # media length of 4:30 or less = 30 second intervals    
            elif MediaLength < 270001:
                Num = int(round((MediaLength + 6000) / 30000.0))
                Interval = 30000
            # media length of 6:00 or less = 60 second intervals    
            elif MediaLength < 360001:
                Num = int(round((MediaLength + 12000) / 60000.0))
                Interval = 60000
            # media length of 10:00 or less = 60 second intervals      
            elif MediaLength < 600001:
                Num = int(round((MediaLength + 8000) / 60000.0))
                Interval = 60000
            # media length of 16:00 or less = 2 minute intervals   
            elif MediaLength < 960001:
                Num = int(round((MediaLength + 24000) / 120000.0))
                Interval = 120000
            # media length of 40:00 or less = 5 minute intervals
            elif MediaLength < 2400001:
                Num = int(round((MediaLength + 60000) / 300000.0))
                Interval = 300000
            # media length if 1:10:00 or less get 10 minute intervals           
            elif MediaLength < 4200001:
                Num = int(round((MediaLength + 80000) / 600000.0))
                Interval = 600000
            # media length if 2:10:00 or less get 20 minute intervals           
            elif MediaLength < 7800001:
                Num = int(round((MediaLength + 160000) / 1200000.0))
                Interval = 1200000
            # media length if 3:00:00 or less get 30 minute intervals           
            elif MediaLength < 10800001:
                Num = int(round((MediaLength + 240000) / 1800000.0))
                Interval = 1800000
            # media length if 4:00:00 or less get 30 minute intervals           
            elif MediaLength < 14400001:
                Num = int(round((MediaLength + 60000) / 1800000.0))
                Interval = 1800000
            # media length if 9:00:00 or less get 60 minute intervals           
            elif MediaLength < 32400001:
                Num = int(round((MediaLength + 120000) / 3600000.0))
                Interval = 3600000
            # Longer videos get 2 hour intervals
            else:
                Num = int(round((MediaLength + 240000) / 7200000.0))
                Interval = 7200000
            return Num, Interval

        # Positioning is different on Mac and Windows, requiring subtle adjustments in this window
        if '__WXMAC__' in wx.PlatformInfo:
            topAdjust = -3
        else:
            topAdjust = 0
        # Set the values for the waveform Lower and Upper limits
        self.waveformLowerLimit = mediaStart
        self.waveformUpperLimit = mediaStart + mediaLength

        

        # Determine the number of labels and the time interval between labels that should be displayed
        numIncrements, Interval = GetScaleIncrements(mediaLength)
        # Now we can determine the appropriate starting point for our labels!
        startingPoint = int((round(mediaStart / Interval) + 1) * Interval)

        # Clear all the existing labels
        self.timeline.DestroyChildren()
        timeLabels = []

        for loop in range(startingPoint, (numIncrements * Interval) + startingPoint, Interval):
            if loop > 0:
                # Place line marks
                lay = wx.LayoutConstraints()
                lay.top.SameAs(self.timeline, wx.Top, topAdjust)
                (width, height) = self.waveform.GetSizeTuple()
                width = width - 6  # Adjust for size of widget frame
                lay.left.Absolute(int(round(((float(loop - mediaStart)) / mediaLength) * (width))))  
                lay.width.AsIs()
                lay.height.AsIs()
                wx.StaticLine(self.timeline, 0, size=wx.Size(2, 5)).SetConstraints(lay)
                
                # Place time labels
                lay = wx.LayoutConstraints()
                lay.top.SameAs(self.timeline, wx.Top, 4 + topAdjust)
                lay.centreX.Absolute(int(round(((float(loop - mediaStart)) / mediaLength) * (width))))
                lay.width.AsIs()
                lay.height.AsIs()
                wx.StaticText(self.timeline, 1, Misc.time_in_ms_to_str(loop)).SetConstraints(lay)

        self.timeline.Layout()
        self.timeline.SetAutoLayout(True)
        # Show the total Media Time in the Total Time label
        self.lbl_Total_Time.SetLabel(Misc.time_in_ms_to_str(mediaLength))

    def draw_timeline_zero(self):
        # Positioning is different on Mac and Windows, requiring subtle adjustments in this window
        if '__WXMAC__' in wx.PlatformInfo:
            topAdjust = -3
        else:
            topAdjust = 2
        # Set the values for the waveform Lower and Upper limits
        self.waveformLowerLimit = 0
        self.waveformUpperLimit = 0

        # Clear all the existing labels
        self.timeline.DestroyChildren()
        timeLabels = []

        for loop in range(3):
            # Place line marks
            lay = wx.LayoutConstraints()
            lay.top.SameAs(self.timeline, wx.Top, topAdjust)
            lay.left.PercentOf(self.timeline, wx.Width, 25 * (loop+1))
            lay.width.AsIs()
            lay.height.AsIs()
            wx.StaticLine(self.timeline, 0, size=wx.Size(2, 5)).SetConstraints(lay)
            
            # Place time labels
            lay = wx.LayoutConstraints()
            lay.top.SameAs(self.timeline, wx.Top, 4 + topAdjust)
            lay.centreX.PercentOf(self.timeline, wx.Width, 25 * (loop+1))
            lay.width.AsIs()
            lay.height.AsIs()
            wx.StaticText(self.timeline, 1, Misc.time_in_ms_to_str(0)).SetConstraints(lay)
        
        self.timeline.Layout()
        self.timeline.SetAutoLayout(True)

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.ControlObject=ControlObject

    def UpdatePosition(self, currentPosition):
        """ In response to an external event, this will update the Visualization window's indication of Media Position """
        # Show the media position in the Current Time label
        self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(currentPosition))
        # If UpperLimit and LowerLimit have not been set yet, avoid dividing by zero.  Things will be cleaned up later.
        if self.waveformUpperLimit == self.waveformLowerLimit:
            self.waveformUpperLimit = self.waveformUpperLimit + 1
        pos = ((float(currentPosition - self.waveformLowerLimit)) / (self.waveformUpperLimit - self.waveformLowerLimit))
        self.waveform.DrawCursor(pos)
        # print '(%s-%s)/%s = Pos = %s' % (currentPosition, self.waveformLowerLimit, (self.waveformUpperLimit - self.waveformLowerLimit), pos)

    def OnLeftDown(self, x, y, xpct, ypct):
        # If we don't convert this to an int, our SQL gets screwed up in non-English localizations that use commas instead
        # of decimals.
        self.startPoint = int(round(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit))
        # Show the media position in the Current Time label
        self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(self.startPoint))       

    def OnLeftUp(self, x, y, xpct, ypct):
        if self.ControlObject.IsPlaying():
            self.ControlObject.Stop()
        # Distinguish a left-click (positioning start only) from a left-drag (select range)
        if self.startPoint != int(round(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit)):
            self.endPoint = int(round(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit))
            # If the user drags to the left rather than to the right, we need to swap the values!
            if self.endPoint < self.startPoint:
                temp = self.startPoint
                self.startPoint = self.endPoint
                self.endPoint = temp
            # Show the media selection in the Selected Time label
            self.lbl_Selected_Time.SetLabel(Misc.time_in_ms_to_str(self.endPoint - self.startPoint))
        else:
            self.endPoint = 0
            # Clear the media selection in the Selected Time label
            self.lbl_Selected_Time.SetLabel(Misc.time_in_ms_to_str(0))
        # Set the Current Video Selection to the highlighted part of the Waveform
        self.ControlObject.SetVideoSelection(self.startPoint, self.endPoint)
    
    def TimeCodeFromPctPos(self, xpct):
        return int(round(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit))
        
    def PctPosFromTimeCode(self, timecode):
        if self.waveformUpperLimit - self.waveformLowerLimit > 0:
            return float(timecode - self.waveformLowerLimit)/(self.waveformUpperLimit - self.waveformLowerLimit)
        else:
            return 0.0

    def OnRightUp(self, x, y):
        self.ControlObject.PlayPause()

    def OnMouseOver(self, x, y, xpct, ypct):
        self.lbl_Time_Time.SetLabel(Misc.time_in_ms_to_str(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit))

    def GetDimensions(self):
        (left, top) = self.GetPositionTuple()
        (width, height) = self.GetSizeTuple()
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
        self.SetDimensions(left, top, width, height)

    def ChangeLanguages(self):
        self.SetTitle(_('Visualization'))
        self.lbl_Time.SetLabel(_("Time:"))
        self.btn_Current.SetLabel(_("Current:"))
        self.btn_Selected.SetLabel(_("Selected:"))
        self.lbl_Total.SetLabel(_("Total:"))
        

# Private methods    

    def __size(self):
        rect = wx.ClientDisplayRect()
        width = rect[2] * .715
        height = (rect[3] - TransanaGlobal.menuHeight) * .25
        return wx.Size(int(width), int(height))

    def __pos(self):
        rect = wx.ClientDisplayRect()
        # If the Start Menu is on the left side, 0 is incorrect!  Get starting position from ClientDisplayRect.
        x = rect[0]
        # rect[1] compensated if the Start menu is at the top of the screen
        y = rect[1] + TransanaGlobal.menuHeight + 3
        
        return wx.Point(int(x), int(y))

