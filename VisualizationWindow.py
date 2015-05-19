# Copyright (C) 2003 - 2007 The Board of Regents of the University of Wisconsin System 
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

DEBUG = False
if DEBUG:
    print "VisualizationWindow DEBUG is ON."

HYBRIDOFFSET = 80

import wx

if __name__ == '__main__':
    # This module expects i18n.  Enable it here.
    __builtins__._ = wx.GetTranslation

import Clip
import Episode
import KeywordMapClass
import TransanaGlobal
import TransanaConstants
import TransanaExceptions
import GraphicsControlClass
import WaveformGraphic
import WaveformProgress       # Waveform Creation Progress Dialog, used in Wave Extraction Callback function
import Dialogs
import Misc
import locale                 # import locale so we can get the default system encoding for Unicode Waveforming
import os
import sys
import string
import time
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
        # Initialize waveformFilename to an empty string.  (Without this, Locate Clip that can't find the video fails badly.)
        self.waveformFilename = ''

        # heightIsSet notes whether the height of the Keyword Visualization should be automatically adjusted or not.
        self.heightIsSet = False

        # redrawWhenIdle signals that the Waveform picture needs to be drawn when the CPU has time
        self.redrawWhenIdle = False
        # Let's keep track of time since last redraw too
        self.lastRedrawTime = time.time()
        # If the video position needs to be set after a waveform redraw, set this value
        self.resetVideoPosition = 0

        # zoomInfo holds information about Zooms, to allow zoom-out in steps matching zoom-ins
        self.zoomInfo = [(0, -1)]

        if DEBUG:
            print "zoomInfo 1", self.zoomInfo

        box = wx.BoxSizer(wx.VERTICAL)

        # We need to know the height of the Window Header to adjust the size of the Graphic Area
        headerHeight = self.GetSizeTuple()[1] - self.GetClientSizeTuple()[1]

        # The waveform is held in a GraphicsControlClass object that handles all the waveform display.
        # At this time, the GraphicsControlClass is not aware of Sizers or Constraints
        self.waveform = GraphicsControlClass.GraphicsControl(self, -1, wx.Point(1, 1), wx.Size(int(width-12), int(height-48-headerHeight)), (width-14, height-52-headerHeight), visualizationMode=True)
        box.Add(self.waveform, 1, wx.EXPAND, 0)

        # Add the Timeline Panel, which holds the time line and scale information
        self.timeline = wx.Panel(self, -1, style=wx.SUNKEN_BORDER) # wx.Point(0, height-44-headerHeight), wx.Size(int(width - 6), 24), 
        box.Add(self.timeline, 0, wx.EXPAND | wx.ALL, 1)

        # Add the Toolbar at the bottom of the screen
        # NOTE:  Because of the way XP handles the screen, we'll start at the bottom and work our way up!
        self.toolbar = wx.Panel(self, -1, style=wx.SUNKEN_BORDER)

        # Add GUI elements to the Toolbar
        toolbarSizer = wx.BoxSizer(wx.HORIZONTAL)

        # Get the graphic for the Filter button
        bmp = wx.ArtProvider_GetBitmap(wx.ART_LIST_VIEW, wx.ART_TOOLBAR, (16,16))
        # Add Filter Button
        self.filter = wx.BitmapButton(self.toolbar, -1, bmp)
        toolbarSizer.Add(self.filter, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT , 2)
        self.filter.Enable(False)
        wx.EVT_BUTTON(self, self.filter.GetId(), self.OnFilter)

        # Add Zoom In Button
        self.zoomIn = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_ZOOMIN, wx.Bitmap('images/ZoomIn.xpm', wx.BITMAP_TYPE_XPM))
        toolbarSizer.Add(self.zoomIn, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT , 2)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_ZOOMIN, self.OnZoomIn)

        # Add Zoom Out Button
        self.zoomOut = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_ZOOMOUT, wx.Bitmap('images/ZoomOut.xpm', wx.BITMAP_TYPE_XPM))
        toolbarSizer.Add(self.zoomOut, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.RIGHT , 2)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_ZOOMOUT, self.OnZoomOut)

        # Add Zoom to 100% Button
        self.zoom100 = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_ZOOM100, wx.Bitmap('images/Zoom100.xpm', wx.BITMAP_TYPE_XPM))
        toolbarSizer.Add(self.zoom100, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.RIGHT , 2)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_ZOOM100, self.OnZoom100)

        # Place line separating Zoom buttons from Position section
        separator1 = wx.StaticLine(self.toolbar, -1, size=wx.Size(2, 22))
        toolbarSizer.Add(separator1, 0, wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT, 4)

        # Add "Time" label
        self.lbl_Time = wx.StaticText(self.toolbar, -1, _("Time:"), size=(38, 16))
        toolbarSizer.Add(self.lbl_Time, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 5)

        # Add "Time" Time label
        self.lbl_Time_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        toolbarSizer.Add(self.lbl_Time_Time, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 2)

        # Place line separating Zoom buttons from Position section
        separator2 = wx.StaticLine(self.toolbar, -1, size=wx.Size(2, 22))
        toolbarSizer.Add(separator2, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 4)

        toolbarSizer.Add((2,1))
        # Add "Current" label
        self.btn_Current = wx.Button(self.toolbar, TransanaConstants.VISUAL_BUTTON_CURRENT, _("Current:"))
        toolbarSizer.Add(self.btn_Current, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.ALL, 2)  # wx.LEFT | wx.RIGHT, 4)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_CURRENT, self.OnCurrent)
        toolbarSizer.Add((2,1))

        # Add "Current" Time label
        self.lbl_Current_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        toolbarSizer.Add(self.lbl_Current_Time, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 8)

        # Add "Selected" label
        self.btn_Selected = wx.Button(self.toolbar, TransanaConstants.VISUAL_BUTTON_SELECTED, _("Selected:"))
        toolbarSizer.Add(self.btn_Selected, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT, 8)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_SELECTED, self.OnSelected)

        # Add "Selected" Time label
        self.lbl_Selected_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        toolbarSizer.Add(self.lbl_Selected_Time, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 8)

        # expandable Spacer
        toolbarSizer.Add((10, 1), 1, wx.EXPAND, 0)

        # Add "Total" label
        self.lbl_Total = wx.StaticText(self.toolbar, -1, _("Total:"))
        toolbarSizer.Add(self.lbl_Total, 0, wx.ALIGN_RIGHT | wx.ALIGN_CENTER | wx.RIGHT, 8)

        # Add "Total" Time label
        # Adjust for Mac's Window Resize Handle, which covers the decimal digit without this code
        if '__WXMAC__' in wx.PlatformInfo:
            rightSpacer = 20
        else:
            rightSpacer = 8
        self.lbl_Total_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        toolbarSizer.Add(self.lbl_Total_Time, 0, wx.ALIGN_RIGHT | wx.ALIGN_CENTER | wx.RIGHT, rightSpacer)

        self.toolbar.SetSizer(toolbarSizer)
        self.toolbar.Fit()
        self.toolbar.SetAutoLayout(True)
        self.toolbar.Layout()

        box.Add(self.toolbar, 0, wx.ALIGN_BOTTOM | wx.EXPAND | wx.ALL, 1)

        self.kwMap = None

        # Stick some initial "0" values in the timeline
        self.draw_timeline_zero()

        # Set the focus on the Waveform widget
        self.waveform.SetFocus()

        self.SetSizer(box)
        
        self.SetAutoLayout(True)
        self.Layout()

        # We need to adjust the screen position on the Mac.  I don't know why.
        if "__WXMAC__" in wx.PlatformInfo:
            pos = self.GetPosition()
            self.SetPosition((pos[0]-20, pos[1]-25))
        
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

        if DEBUG:
            print "VisualizationWindow: 1 draw_timeline(%s, %s)" % (self.waveformLowerLimit, self.waveformUpperLimit - self.waveformLowerLimit)

        # Position timeline Panel
        self.timeline.SetDimensions(0, height-46-headerHeight, width-6, 24)

    def Refresh(self):
        self.ClearVisualization()
        self.redrawWhenIdle = True
        self.OnIdle(None)

    def UpdateKeywordVisualization(self):
        # Only do something if we've got the Keyword or Hybrid visualization and there is a data object currently loaded
        if (TransanaGlobal.configData.visualizationStyle in ['Keyword', 'Hybrid']) and \
           (self.ControlObject.currentObj != None) and \
           (self.kwMap != None):
            # Update the Visualization's underlying data
            self.kwMap.UpdateKeywordVisualization()
            # Signal that height can be reset
            self.heightIsSet = False
            # Call the resize Height method
            self.resizeKeywordVisualization()

    def OnIdle(self, event):
        """ Use Idle Time to handle the drawing in this control """
        # Check to see if the waveform control needs to be redrawn.  Under the new Media Player, GetMediaLength takes a while to
        # be set, so we should wait for that too.
        if self.redrawWhenIdle  and (not self.ControlObject.shuttingDown) and (self.ControlObject.GetMediaLength() > 0):
            # Remove old Waveform Selection and Cursor data
            self.waveform.ClearTransanaSelection()

            # If no main object if defined ...
            if self.ControlObject.currentObj == None:
                # ... clear the current waveform ...
                self.waveform.Clear()
                # ... signal that the redraw is done ...
                self.redrawWhenIdle = False
                # ... and exit this method.  We're done here.
                return

            if TransanaGlobal.configData.visualizationStyle in ['Waveform', 'Hybrid']:
                # self.waveform.Clear()
                # Create the appropriate Waveform Graphic
                if (self.waveformFilename != ''):
                    try:
                        
                        # The Mac can't handle Unicode WaveFilenames at this point.  We need to upgrade to Python 2.4 for that.
                        # Let's temporarily take care of that.
                        if ('wxMac' in wx.PlatformInfo) and isinstance(self.waveFilename, unicode):
                            self.waveFilename = self.waveFilename.encode('utf8')

                        # We populate the keyword visualization differently for an episode and a clip.
                        if type(self.ControlObject.currentObj) == Episode.Episode:
                            # The starting point of the Keyword Visualization is the start time in the top of the zoomInfo stack!
                            start = self.zoomInfo[-1][0]
                            # The ending point of the Keyword Visualization is the end time in the top of the zoomInfo stack,
                            # unless that value is 0, in which case it's the media's full length.
                            if self.zoomInfo[-1][1] == 0:
                                length = self.ControlObject.GetMediaLength(True)
                            else:
                                length = self.zoomInfo[-1][1]
                        elif type(self.ControlObject.currentObj) == Clip.Clip:
                            # Set the current global video selection based on the Clip.
                            self.ControlObject.SetVideoSelection(self.ControlObject.currentObj.clip_start, self.ControlObject.currentObj.clip_stop) 
                            # we have to know the right start and end points for the waveform.
                            start = self.ControlObject.VideoStartPoint
                            length = self.ControlObject.GetMediaLength()

                        if DEBUG:
                            print "Creating Waveform:", start, length

                        if WaveformGraphic.WaveformGraphicCreate(self.waveFilename, self.waveformFilename, start, length, self.waveform.canvassize, True):
                            # Load the new graphic into the Waveform Control
                            self.waveform.LoadFile(self.waveformFilename)
                            # Determine NEW Frame Size
                            (width, height) = self.GetSize()
                            # Draw the TimeLine values
                            self.draw_timeline(start, length)

                            if DEBUG:
                                print "VisualizationWindow: 2 draw_timeline(%s, %s)" % (start, length)
                            
                        else:
                            self.waveformFilename = ''
                            self.redrawWhenIdle = False
                            
                    # A bug in Python 2.3.5 causes a RuntimeError with some wave files if Unicode filenames are used.
                    # This should be fixed in Python 2.4.2, but we'll leave this code here to prevent ugly errors if it
                    # does occur.
                    except RuntimeError, e:
                        self.waveformFilename = ''
                        self.ClearVisualization()
                        self.redrawWhenIdle = False

                    except:
                        # NO Error Message, as it disrupt the program flow if the user chooses not to waveform
                        if DEBUG and False:
                            dlg = Dialogs.ErrorDialog(self, 'DEBUG (UNTRANSLATED) Waveform Graphic creation error for file:\n%s\n%s.' % (self.waveFilename, self.waveformFilename))
                            dlg.ShowModal()
                            dlg.Destroy()

                        self.waveformFilename = ''
                        self.ClearVisualization()
                        self.redrawWhenIdle = False

                        if DEBUG:
                            # Beware of recursive Yields; trap the exception ...
                            try:
                                wx.Yield()
                            # ... and ignore it!
                            except:
                                pass
                            print sys.exc_info()[0], sys.exc_info()[1]
                            import traceback
                            traceback.print_exc(file=sys.stdout)

                # If we're building a Hybrid visualization, capture the waveform picture.
                # But if waveforFilename has been cleared, we're waveforming and shouldn't do this yet!
                if (TransanaGlobal.configData.visualizationStyle == 'Hybrid') and (self.waveformFilename != ''):
                    # Get the waveform Bitmap and convert it to an Image
                    hybridWaveform = self.waveform.bmpBuffer.ConvertToImage()
                    # Rescale the image so that it matches the size alloted for the Waveform (HYBRIDOFFSET)
                    hybridWaveform.Rescale(hybridWaveform.GetWidth(), HYBRIDOFFSET)

            if TransanaGlobal.configData.visualizationStyle in ['Keyword', 'Hybrid']:
                # Clear the Visualization
                self.waveform.Clear()
                # Enable the Filter button
                self.filter.Enable(True)
                # If there's an existing Keyword Visualization ...
                if self.kwMap != None:
                    # ... remember the values for the filtered Keyword List
                    filteredKeywordList = self.kwMap.filteredKeywordList[:]
                    # ... remember the values from the unfiltered keyword list
                    unfilteredKeywordList = self.kwMap.unfilteredKeywordList[:]
                    # ... remember the keyword color list too.
                    keywordColorList = self.kwMap.keywordColors
                    # Delete the current keyword visualization object
                    del(self.kwMap)
                    # Set the reference to the keyword visualizatoin object to None so we don't get confused.
                    self.kwMap = None
                # If we're creating a brand new Keyword Visualization ...
                else:
                    # Initialize the filtered keyword list ...
                    filteredKeywordList = []
                    # ... the unfiltered keyword list ...
                    unfilteredKeywordList = []
                    # ... and the keyword color list.
                    keywordColorList = None
                # If we're creating a Hybrid Visualization ...
                if TransanaGlobal.configData.visualizationStyle == 'Hybrid':
                    # ... then the Keyword Visualization portion needs a top margin.
                    topOffset = HYBRIDOFFSET
                # If we're just doing a Keyword Visualization ...
                else:
                    # ... then we don't need a top margin
                    topOffset = 0
                # Create a Keyword Visualization object as an embedded graphic, not a free-standing report.
                self.kwMap = KeywordMapClass.KeywordMap(self, -1, "", embedded=True, topOffset=topOffset)
                # We populate the keyword visualization differently for an episode and a clip.
                if type(self.ControlObject.currentObj) == Episode.Episode:
                    # On the Mac, the video length cannot be determined before the video load is complete.
                    # This leads to problems with the self.zoomInfo values, which are set BEFORE the video
                    # load is complete.  Let's try to detect and repair that problem here, after the video
                    # has been loaded.
                    if self.zoomInfo[0][1] == -1:
                        self.zoomInfo[0] = (self.ControlObject.GetVideoStartPoint(), self.ControlObject.GetMediaLength(True))

                        if DEBUG:
                            print "zoomInfo 2", self.zoomInfo
                            
                    # The starting point of the Keyword Visualization is the start time in the top of the zoomInfo stack!
                    kwMapStartPoint = self.zoomInfo[-1][0]
                    # The ending point of the Keyword Visualization is the end time in the top of the zoomInfo stack,
                    # unless that value is 0, in which case it's the media's full length.
                    if self.zoomInfo[-1][1] == 0:
                        kwMapEndPoint = kwMapStartPoint + self.ControlObject.GetMediaLength(True)
                    else:
                        kwMapEndPoint = kwMapStartPoint + self.zoomInfo[-1][1]

                    # Set up the embedded Keyword Visualization, sending it all the data it needs so it can draw or redraw itself.
                    self.kwMap.SetupEmbedded(self.ControlObject.currentObj.number, self.ControlObject.currentObj.series_id, \
                                             self.ControlObject.currentObj.id, kwMapStartPoint, kwMapEndPoint, \
                                             filteredKeywordList = filteredKeywordList, unfilteredKeywordList = unfilteredKeywordList, \
                                             keywordColors = keywordColorList)

                    # Draw the TimeLine values
                    self.draw_timeline(kwMapStartPoint, kwMapEndPoint - kwMapStartPoint)

                    if DEBUG:
                        print "VisualizationWindow: 3 draw_timeline(%s, %s)" % (kwMapStartPoint, kwMapEndPoint - kwMapStartPoint)

                elif type(self.ControlObject.currentObj) == Clip.Clip:
                    # If we're working with a Clip, we need some information from it's source episode.  Let's get the Episode.
                    try:
                        tmpEpisode = Episode.Episode(self.ControlObject.currentObj.episode_num)
                    except TransanaExceptions.RecordNotFoundError, e:
                        tmpEpisode = Episode.Episode()
                        tmpEpisode.series_id = 'None'
                        tmpEpisode.id = 'None'
                    # Set the current global video selection based on the Clip.
                    self.ControlObject.SetVideoSelection(self.ControlObject.currentObj.clip_start, self.ControlObject.currentObj.clip_stop)

                    # Set up the embedded Keyword Visualization, sending it all the data it needs so it can draw or redraw itself.
                    self.kwMap.SetupEmbedded(self.ControlObject.currentObj.episode_num, tmpEpisode.series_id, \
                                             tmpEpisode.id, self.ControlObject.currentObj.clip_start, self.ControlObject.currentObj.clip_stop, \
                                             filteredKeywordList = filteredKeywordList, unfilteredKeywordList = unfilteredKeywordList, \
                                             keywordColors = keywordColorList, clipNum=self.ControlObject.currentObj.number)

                    # Draw the TimeLine values
                    self.draw_timeline(self.ControlObject.VideoStartPoint, self.ControlObject.GetMediaLength())

                    if DEBUG:
                        print "VisualizationWindow: 4 draw_timeline(%s, %s)" % (self.ControlObject.VideoStartPoint, self.ControlObject.GetMediaLength())

                elif self.ControlObject.currentObj != None:
                    self.waveform.AddText('Keyword Visualization - %s not implemented.' % type(self.ControlObject.currentObj), 5, 5)

                # The Keyword / Hybrid visualization height can be self-adjusting.  Let's call that function.
                self.resizeKeywordVisualization()

                # If we're bulding a Hybrid visualization, so far we've created and stored the waveform, then
                # wiped it out in favor of an offset Keyword visualization.  Here, we combine the two
                # visualizations!  If waveforFilename has been cleared, we're waveforming and shouldn't do this yet!
                if (TransanaGlobal.configData.visualizationStyle == 'Hybrid') and (self.waveformFilename != ''):
                    # Here's a trick.  By setting the waveform's backgroundImage but NOT setting the
                    # backgroundGraphicName, you can add a background image to the GraphicsControlClass
                    # that does not resize to fill the image.  The current offset Keyword visualization
                    # with the waveform overlaid as a background image works pretty well!
                    self.waveform.backgroundImage = hybridWaveform

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

    def resizeKeywordVisualization(self):
        """ The Keyword Visualization (and Hybrid) should auto-resize under some circumstances.  This method implements that. """
        # If Auto Arrange is ON and the height has not already been set ...
        if TransanaGlobal.configData.autoArrange and not self.heightIsSet:
            # ... start by getting the current dimensions of the visualization window.
            (a, b, c, d) = self.GetDimensions()
            # We need to make different adjustments based on whether we're on the Mac or Windows.
            if "wxMac" in wx.PlatformInfo:
                # On Mac, the window height is the smaller of the Keyword Visualization + 100 for the rest of the
                # window or 2/3 of the screen height
                newHeight = min(self.kwMap.GetKeywordCount()[1] + 100, 2 * wx.ClientDisplayRect()[3] / 3)
            else:
                # On Windows, the window height is the smaller of the Keyword Visualization + 142 for the rest of the
                # window or 2/3 of the screen height
                newHeight = min(self.kwMap.GetKeywordCount()[1] + 142, 2 * wx.ClientDisplayRect()[3] / 3)
            # Let's say that 1/4 of the screen is the minimum Waveform height!
            newHeight = max(newHeight, round(wx.ClientDisplayRect()[3] / 4))

            # The Hybrid Visualization was losing the waveform when the Filter Dialog was called.  So when this happens ...
            if TransanaGlobal.configData.visualizationStyle == 'Hybrid':
                # ... signal that the whole visualization has to be re-drawn!
                self.redrawWhenIdle = True

            # now let's adjust the window sizes for the main Transana interface.
            self.ControlObject.UpdateWindowPositions('Visualization', c, YUpper = newHeight)
            # once we do this, we don't need to do it again unless something changes.
            self.heightIsSet = True
        
    def OnKeyDown(self, event):
        """ Captures Key Events to allow this window to control video playback during transcription. """
        # Get the Key Code
        c = event.GetKeyCode()

        try:

            # If Shift is Down ...
            if event.ShiftDown():
                # ... we move 1 second
                timePerPixel = 1000.0
            # if Alt is down ...
            elif event.AltDown():
                # ... we move 1/2 second
                timePerPixel = 500.0
            # If Ctrl is Down ...
            elif event.ControlDown():
                # ... we move 1 frame only
                timePerPixel = 33.3667
            # If no modifier key is down ...
            else:
                # We move 1 pixel rather than 1 frame.
                # Determine the size of the waveform diagram
                (width, height) = self.waveform.GetSizeTuple()
                # Adjust for size of widget frame
                width = width - 6
                # Now calculate the amount of TIME a single pixel represents
                timePerPixel = (self.waveformUpperLimit - self.waveformLowerLimit) / width
                # Video Frame size is 1/29.97 of a second, or 33.3667 milliseconds.  1 frame is the minimum size to be moved here!
                timePerPixel = max(timePerPixel, 33.3667)
            # Determine the current position (in milliseconds) of the video
            currentPos = self.ControlObject.GetVideoPosition()

            # Ctrl-A -- Rewind video by 10 seconds
            if (c == ord("A")) and event.ControlDown():
                vpos = self.ControlObject.GetVideoPosition()
                self.ControlObject.SetVideoStartPoint(vpos-10000)
                self.ControlObject.SetVideoEndPoint(0)
                # Play should always be initiated on Ctrl-A
                self.ControlObject.Play(0)

            # Ctrl-D -- Start / Pause without setback
            elif (c == ord("D")) and event.ControlDown():
                self.ControlObject.SetVideoEndPoint(0)
                self.ControlObject.PlayPause(0)

            # CTRL-F -- Advance video by 10 seconds
            elif (c == ord("F")) and event.ControlDown():
                vpos = self.ControlObject.GetVideoPosition()
                self.ControlObject.SetVideoStartPoint(vpos+10000)
                self.ControlObject.SetVideoEndPoint(0)
                # Play should always be initiated on Ctrl-F
                self.ControlObject.Play(0)

            # Ctrl-S -- Start / Pause with setback
            elif (c == ord("S")) and event.ControlDown():
                self.ControlObject.SetVideoEndPoint(0)
                self.ControlObject.PlayPause(1)
                
            # Ctrl-T -- Insert Time Code
            elif (c == ord("T")) and event.ControlDown():
                self.OnCurrent(event)
                    
            # Cursor Left ...
            elif c in [wx.WXK_LEFT, wx.WXK_NUMPAD_LEFT]:
                # ... moves the video the equivalent of 1 pixel earlier in the video
                if currentPos > self.waveformLowerLimit - timePerPixel:
                    self.ControlObject.SetVideoStartPoint(currentPos - timePerPixel)

            # Cursor Right ...
            elif c in [wx.WXK_RIGHT, wx.WXK_NUMPAD_RIGHT]:
                # ... moves the video the equivalent of 1 pixel later in the video
                if currentPos < self.waveformUpperLimit + timePerPixel:
                    self.ControlObject.SetVideoStartPoint(currentPos + timePerPixel)

        except:

            if DEBUG:
                print sys.exc_info()[0], sys.exc_info()[1]
                import traceback
                traceback.print_exc(file=sys.stdout)
            else:
                pass

    def OnFilter(self, event):
        """ Call the Keyword Visualization Filter Dialog """
        # Call the Keyword Visualization Filter Dialog
        self.kwMap.OnFilter(event)
        # The Hybrid Visualization was losing the waveform when the Filter Dialog was called.  So when this happens ...
        if TransanaGlobal.configData.visualizationStyle == 'Hybrid':
            # ... signal that the whole visualization has to be re-drawn!
            self.redrawWhenIdle = True
        else:
            # If we'd cleared the waveform before calling OnFilter, we'd be done.  But that clears the graphic
            # BEFORE bringing up the filter dialog, which I didn't like.  So even though OnFilter re-draws the
            # visualization (albeit incorrectly because we didn't call waveform.Clear()), we'll now clear the
            # waveform and re-draw the keyword visualization.
            self.waveform.Clear()
            self.kwMap.DrawGraph()

    def OnZoomIn(self, event):
        """ Zoom in on a portion of the Waveform Diagram """
        if (self.startPoint < self.endPoint):
            # Keep track of the new position.  This allows the user to zoom back out in the same steps used to zoom in
            self.zoomInfo.append((int(self.startPoint), int(self.endPoint - self.startPoint)))
            # Limit video playback to the selected part of the media by setting the VideoStartPoint and VideoEndPoint in the
            # Control Object
            self.ControlObject.SetVideoSelection(self.startPoint, self.endPoint)

            # Draw the TimeLine values
            self.draw_timeline(self.ControlObject.VideoStartPoint, self.ControlObject.GetMediaLength())
            
            # assign a temporary filename for the Waveform Graphic
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
            self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.png'
            self.redrawWhenIdle = True
            # Clear the Start and End points of the Visualization Selection
            self.startPoint = self.ControlObject.VideoStartPoint
            self.endPoint = self.ControlObject.VideoStartPoint + self.ControlObject.GetMediaLength()

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
        self.zoomInfo = [self.zoomInfo[0]]

        if DEBUG:
            print "zoomInfo 3", self.zoomInfo
            
        self.startPoint = self.zoomInfo[0][0]
        self.endPoint = self.zoomInfo[0][1]
        self.ControlObject.SetVideoSelection(self.startPoint, self.endPoint)
        # Clear the waveform itself
        self.waveform.Clear()
        # Remove old Waveform Selection and Cursor data
        self.waveform.ClearTransanaSelection()
        self.filter.Enable(False)
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
        self.zoomInfo = [(mediaStart, mediaLength)]

        if DEBUG:
            print "zoomInfo 4", self.zoomInfo, mediaStart, mediaLength, self.ControlObject.GetVideoStartPoint(), self.ControlObject.GetVideoEndPoint()
            
        # Let's clear the Visualization Window as we get started.
        self.ClearVisualization()
        # Remember the original File Name that is passed in
        originalFilename = filename
        fn = filename
        # Separate path and filename
        (path, filename) = os.path.split(filename)
        # break filename into root filename and extension
        (filenameroot, extension) = os.path.splitext(filename)
        # Build the correct filename for the Wave File
        self.waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
        # Build the correct filename for the Waveform Graphic
        self.waveformFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.png')
        # Remove old Waveform Selection and Cursor data
        self.waveform.ClearTransanaSelection()

        if self.kwMap != None:
            del(self.kwMap)
            self.kwMap = None
        # reset heightIsSet
        self.heightIsSet = False

        # If we are showing the whole waveform and the Waveform graphic exists and QuickLoad is enabled, load the graphic from the file.
        if not(os.path.exists(self.waveformFilename) and TransanaGlobal.configData.waveformQuickLoad):
            # Create a Wave File if none exists!
            if not(os.path.exists(self.waveFilename)):
                # Mac cannot show this modal dialog when the modal Play All Clips dialog is shown, so check for that.
                # Check the user's response to the dialog. 
                if ((not 'wxMac' in wx.PlatformInfo) or (not self.ControlObject.PlayAllClipsWindow)):
                    # Politely ask the user to create the waveform
#                    dlg = wx.MessageDialog(self, _("No wave file exists.  Would you like to create one now?"), _("Transana Wave File Creation"), wx.YES_NO | wx.ICON_QUESTION | wx.CENTRE)
#                    if dlg.ShowModal() == wx.ID_YES:
                    dlg = Dialogs.QuestionDialog(self, _("No wave file exists.  Would you like to create one now?"), _("Transana Wave File Creation"))
                    if dlg.LocalShowModal() == wx.ID_YES:
                        try:
                            # Turn off OnIdle Redraw during audio extraction.
                            self.redrawWhenIdle = False
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
                            # Build the file name for the extracted audio
                            self.waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
                            # Create the Waveform Progress Dialog
                            self.progressDialog = WaveformProgress.WaveformProgress(self, self.waveFilename, prompt % (self.waveFilename, originalFilename))
                            # Tell the Waveform Progress Dialog to handle the audio extraction modally.
                            self.progressDialog.Extract(originalFilename, self.waveFilename)
                            # Okay, we're done with the Progress Dialog here!
                            self.progressDialog.Destroy()
                            # We just have to assume that audio extraction worked.  Signal success!
                            dllvalue = 0                                

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

                    else:
                        # User declined to create the WAV file now
                        dllvalue = 1  # Signal that the WAV file was NOT created!
                    # Destroy the Dialog that asked to create the Wave file    
                    dlg.Destroy()
                else:
                    # User was not eligible to create the WAV file now
                    dllvalue = 1  # Signal that the WAV file was NOT created!
                # If the user said no or there was a problem with Wave Extraction ...
                if dllvalue != 0:
                    # Remove whatever image might be there now.  First, clear the file names.
                    self.waveFilename = ''
                    self.waveformFilename = ''
                    # Then clear the waveform itself
                    self.waveform.Clear()

        # If this is a Clip, we need to create a temporary Waveform to show only the portion we need to see
        if (mediaStart != 0) and (self.waveformFilename != '') and (os.path.exists(self.waveformFilename)):

            # assign a temporary filename for the Waveform Graphic
            self.waveformFilename = TransanaGlobal.configData.visualizationPath + 'tempWave.png'

        # If we're in Hybrid mode, clear the visualization to prevent waveform contamination!
        if TransanaGlobal.configData.visualizationStyle == 'Hybrid':
            self.ClearVisualization()

        # It's possible we lost the waveformFilename during audio extraction.  It's okay to re-create it here!
        if self.waveformFilename == '':
            # Build the correct filename for the Waveform Graphic
            self.waveformFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.png')

        # Now that audio extraction is complete, signal that it's time to draw the Waveform Diagram during
        # Idle time.
        self.redrawWhenIdle = True
        # Show the media position in the Current Time label
        self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(mediaStart))
        # Draw the TimeLine values
        self.draw_timeline(mediaStart, mediaLength)

        if DEBUG:
            print "VisualizationWindow: 5 draw_timeline(%s, %s)" % (mediaStart, mediaLength)


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
                (width, height) = self.timeline.GetSizeTuple()
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

        self.timeline.SetAutoLayout(True)
        self.timeline.Layout()
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
        
        self.timeline.SetAutoLayout(True)
        self.timeline.Layout()

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
        # Force a redraw at least every 0.2 seconds while the video is playing.  (Mostly for slow Macs)
        if (self.ControlObject.IsPlaying()) and (time.time() - self.lastRedrawTime > 0.2):
            self.waveform.reInitBuffer = True
            # Actually, the video is getting very jumpy when the keyword or hybrid visualization is too complex.
            # Removing this seems to help that.
            # self.waveform.OnIdle(None)

            # While we're at it, let's see if we need to Yield().  This makes the app MUCH more responsive on slow systems. 
            # There is an issue with recursive calls to wxYield, so trap the exception ...
            try:
                wx.YieldIfNeeded()
            # ... and ignore it!
            except:
                pass
            # Let's keep track of time since last redraw
            self.lastRedrawTime = time.time()
            
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
        # Mac was having problems with the default window position being at -20.
        # Don't accept anything less than 0 for the left parameter.
        if left < 0:
            left = 0
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
        """ Change the dimensions of the Visualization Window """
        # Change the Visualization Frame's dimensions
        self.SetDimensions(left, top, width, height)
        # If we've changed the size of the visualization window, we need to redraw the visualization
        self.redrawWhenIdle = True

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

if __name__ == '__main__':
    app = wx.PySimpleApp()
    frame = VisualizationWindow(None)
    frame.ShowModal()
    frame.Destroy()
    app.MainLoop()
