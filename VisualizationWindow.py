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

"""This module implements the Visualization class as part of the Visualization component."""

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>, Nathaniel Case, Rajas Sambhare'

DEBUG = False
if DEBUG:
    print "VisualizationWindow DEBUG is ON."

HYBRIDOFFSET = 80

# Import wxPython
import wx
# Import wxPython's extra buttons
import wx.lib.buttons

# For stand-alone testing...
if __name__ == '__main__':
    # This module expects i18n.  Enable it here.
    __builtins__._ = wx.GetTranslation

# Import Transana's Clip Object
import Clip
# Import Transana's Dialogs
import Dialogs
# Import Transana's Document Object
import Document
# Import Tranana's Episodde Object
import Episode
# Import Transana's GraphsicControlClass, the canvas for the Waveform / Keyword Visualization
import GraphicsControlClass
# Import Transana's Keyword Map Class, responsible for Keyword Visualizations
import KeywordMapClass
# Import Transana's Miscellaneous Functions
import Misc
# Import Transana's Quote Object
import Quote
# Import Transana's Constants
import TransanaConstants
# Import Transana's Exceptions
import TransanaExceptions
# Import Transana's Globals
import TransanaGlobal
# Import Transana Images
import TransanaImages
# Import Transana's Waveform Creation routines
import WaveformGraphic
# Import Transana's Waveform Creation Progress Dialog, used in Wave Extraction Callback function
import WaveformProgress

# Import Python's ctypes module, used to access wceraudio DLL/Shared Library
import ctypes
# Import Python's locale module, used for i18n and unicode handling
import locale
# Import Python's os module
import os
# Import Python's String module
import string
# Import Python's sys module
import sys
# Import Python's time module
import time

class VisualizationWindow(wx.Dialog):
    """ This class creates Transana's Visualization Window, used to display waveforms (for media files) and Keyword Visualizations
        (for media and text data).  These visualizations are intended to provide the user with useful information abou their
        data in a visual form. """

    def __init__(self, parent):
        """Initialize a VisualizationWindow object."""
        # Create the Visualization Window
        wx.Dialog.__init__(self, parent, -1, _('Visualization'), pos=self.__pos(), size=self.__size(), style=wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS)
        # Set "Window Variant" to small only for Mac to use small icons
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # if we're not on Linux ...
        if not 'wxGTK' in wx.PlatformInfo:
            self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_FRAMEBK))
        else:
            self.SetBackgroundColour(wx.WHITE)

        # Define the Visualization Window's minimum size
        self.SetSizeHints(250, 100)
        # Get the Visualization WIndow's current size
        (width, height) = self.GetSize()

        # The ControlObject handles all inter-object communication, initialized to None
        self.ControlObject = None

        # Initialize the Object to be Visualized
        self.VisualizationObject = None
        # Initialize the Visualization Type
        self.VisualizationType = None

        # Create a Dictionary for the Visualization Settings for each object being visualized.
        # The key reflects the object held in the Document Window's current Notebook page.
        self.VisualizationInfo = {}
        
        # startPoint and endPoint are used in the process of drag-selecting part of a waveform.
        self.startPoint = 0
        self.endPoint = 0

        # waveformLowerLimit and waveformUpperLimit contain the lower and upper bounds (in milliseconds) of the portion of the waveform
        # that is currently displayed.  These are crucial in all positioning calculations.
        self.waveformLowerLimit = 0
        self.waveformUpperLimit = 0
        # We need to signal if the Default Visualization should be loaded or not.  (Used for first-time loads?)
        self.loadDefault = True

        # heightIsSet notes whether the height of the Keyword Visualization should be automatically adjusted or not.
        self.heightIsSet = False

        # redrawWhenIdle signals that the Waveform picture needs to be drawn when the CPU has time
        self.redrawWhenIdle = False
        # Initialize a list structure to hold wave file information
        self.waveFilename = []
        # Let's keep track of time since last redraw too
        self.lastRedrawTime = time.time()
        # If the video position needs to be set after a waveform redraw, set this value
        self.resetVideoPosition = 0

        # zoomInfo holds information about Zooms, to allow zoom-out in steps matching zoom-ins
        self.zoomInfo = [(0, -1)]
        # We need to know the height of the Window Header to adjust the size of the Graphic Area
        headerHeight = self.GetSizeTuple()[1] - self.GetClientSizeTuple()[1]

        # Create a main Box sizer
        box = wx.BoxSizer(wx.VERTICAL)

        # The waveform / Keyword Visualization is held in a GraphicsControlClass object that handles the visualization display.
        # At this time, the GraphicsControlClass is not aware of Sizers or Constraints
        self.waveform = GraphicsControlClass.GraphicsControl(self, -1, wx.Point(1, 1), wx.Size(int(width-12), int(height-48-headerHeight)), (width-14, height-52-headerHeight), visualizationMode=True)
        box.Add(self.waveform, 1, wx.EXPAND, 0)

        # Add the Timeline Panel, which holds the time line and scale information
        self.timeline = wx.Panel(self, -1, style=wx.SUNKEN_BORDER)
        box.Add(self.timeline, 0, wx.EXPAND | wx.ALL, 1)

        # Add the Toolbar at the bottom of the screen
        # NOTE:  Because of the way XP handles the screen, we'll start at the bottom and work our way up!
        self.toolbar = wx.Panel(self, -1, size=(width, 100), style=wx.SUNKEN_BORDER)
        self.toolbar.SetMinSize((-1, 32))

        # Add GUI elements to the Toolbar
        toolbarSizer = wx.BoxSizer(wx.HORIZONTAL)

        # Add Filter Button
        self.filter = wx.BitmapButton(self.toolbar, -1, TransanaImages.ArtProv_LISTVIEW.GetBitmap())
        self.filter.SetToolTipString(_("Filter"))
        toolbarSizer.Add(self.filter, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT , 2)
        self.filter.Enable(False)
        wx.EVT_BUTTON(self, self.filter.GetId(), self.OnFilter)

        # Add Zoom In Button
        self.zoomIn = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_ZOOMIN, TransanaGlobal.GetImage(TransanaImages.ZoomIn))
        self.zoomIn.SetToolTipString(_("Zoom In"))
        toolbarSizer.Add(self.zoomIn, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT , 2)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_ZOOMIN, self.OnZoomIn)

        # Add Zoom Out Button
        self.zoomOut = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_ZOOMOUT, TransanaGlobal.GetImage(TransanaImages.ZoomOut))
        self.zoomOut.SetToolTipString(_("Zoom Out"))
        toolbarSizer.Add(self.zoomOut, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT , 2)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_ZOOMOUT, self.OnZoomOut)

        # Add Zoom to 100% Button
        self.zoom100 = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_ZOOM100, TransanaGlobal.GetImage(TransanaImages.Zoom100))
        self.zoom100.SetToolTipString(_("Zoom to 100%"))
        toolbarSizer.Add(self.zoom100, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT , 2)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_ZOOM100, self.OnZoom100)

        # Put in a tiny horizontal spacer
        toolbarSizer.Add((4, 0))

        # Add Create Clip Button
        self.createClip = wx.BitmapButton(self.toolbar, TransanaConstants.VISUAL_BUTTON_CREATECLIP, TransanaGlobal.GetImage(TransanaImages.Clip16))
        self.createClip.SetToolTipString(_("Create Transcript-less Clip"))
        toolbarSizer.Add(self.createClip, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT , 2)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_CREATECLIP, self.OnCreateClip)

        # There is a problem with the lib.buttons.GenBitmapToggleButton in Arabic.  I don't know if it's limited to Arabic,
        # or if other right-to-left languages are also affected because of image reversal.  For the moment, we'll use
        # LayoutDirection to detect this issue.

        # if we have a left-to-right language, use the wx.lib.buttons.GenBitmapToggleButton
        if TransanaGlobal.configData.LayoutDirection == wx.Layout_LeftToRight:
            # Add a button for looping playback
            # We need to use a Generic Bitmap Toggle Button!
            self.loop = wx.lib.buttons.GenBitmapToggleButton(self.toolbar, -1, None)
            # Define the image for the "un-pressed" state
            self.loop.SetBitmapLabel(TransanaGlobal.GetImage(TransanaImages.loop_up))
            # Define the image for the "pressed" state
            self.loop.SetBitmapSelected(TransanaGlobal.GetImage(TransanaImages.loop_down))
            self.loop.SetToolTipString(_("Loop Playback"))
            # Set the button to "un-pressed"
            self.loop.SetToggle(False)
        # If we have a right-to-left language, use a standard BitmapButton
        else:
            # Add a button for looping playback
            # We need to use the regular Bitmap Button!
            self.loop = wx.BitmapButton(self.toolbar, -1, TransanaGlobal.GetImage(TransanaImages.loop_up))
            self.loop.SetToolTipString(_("Loop Playback"))
            # Since we don't have the GetValue function, let's create a variable called Looping to tell us if we're
            # looping or not.  Initialize to False.
            self.looping = False
            
        # Set the button's initial size
        self.loop.SetInitialSize((20, 20))
        # Add the button to the toolbar
        toolbarSizer.Add(self.loop, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT , 2)
        # Bind the OnLoop event handler to the button
        self.loop.Bind(wx.EVT_BUTTON, self.OnLoop)

        spacerSize = 10

        # Place line separating Zoom buttons from Position section
        separator = wx.StaticLine(self.toolbar, -1, size=wx.Size(2, 22))
        toolbarSizer.Add(separator, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 2)

        # Add "Time" label
        self.lbl_Time = wx.StaticText(self.toolbar, -1, _("Time:"))
        toolbarSizer.Add(self.lbl_Time, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 2)

        # Add "Time" Time label
        self.lbl_Time_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        toolbarSizer.Add(self.lbl_Time_Time, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 4)
        toolbarSizer.Add((spacerSize + 30, 1))

        # Place line separating Time from Current
        separator = wx.StaticLine(self.toolbar, -1, size=wx.Size(2, 22))
        toolbarSizer.Add(separator, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 2)

        # Add "Current" label
        self.btn_Current = wx.Button(self.toolbar, TransanaConstants.VISUAL_BUTTON_CURRENT, _("Current:"))
        self.btn_Current.SetToolTipString(_("Insert Time Code"))
        toolbarSizer.Add(self.btn_Current, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 2)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_CURRENT, self.OnCurrent)

        # Add "Current" Time label
        self.lbl_Current_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        toolbarSizer.Add(self.lbl_Current_Time, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 4)
        toolbarSizer.Add((spacerSize + 30, 1))

        # Place line separating Current from Selected
        separator = wx.StaticLine(self.toolbar, -1, size=wx.Size(2, 22))
        toolbarSizer.Add(separator, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 2)

        # Add "Selected" label
        self.btn_Selected = wx.Button(self.toolbar, TransanaConstants.VISUAL_BUTTON_SELECTED, _("Selected:"))
        self.btn_Selected.SetToolTipString(_("Insert Time Span"))
        toolbarSizer.Add(self.btn_Selected, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 2)
        wx.EVT_BUTTON(self, TransanaConstants.VISUAL_BUTTON_SELECTED, self.OnSelected)

        # Add "Selected" Time label
        self.lbl_Selected_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        toolbarSizer.Add(self.lbl_Selected_Time, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 4)
        toolbarSizer.Add((10 * spacerSize, 1))

        # Place line separating Selected from Total
        separator = wx.StaticLine(self.toolbar, -1, size=wx.Size(2, 22))
        toolbarSizer.Add(separator, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 2)

        # Add "Total" label
        self.lbl_Total = wx.StaticText(self.toolbar, -1, _("Total:"))
        toolbarSizer.Add(self.lbl_Total, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 2)

        # Add "Total" Time label
        self.lbl_Total_Time = wx.StaticText(self.toolbar, -1, "0:00:00.0")
        toolbarSizer.Add(self.lbl_Total_Time, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, 4)
        toolbarSizer.Add((1, 0), 1, wx.EXPAND | wx.RIGHT, 2)

        self.toolbar.SetSizer(toolbarSizer)
        self.toolbar.SetAutoLayout(True)
        self.toolbar.Fit()
        self.toolbar.Layout()

        box.Add(self.toolbar, 0, wx.ALIGN_BOTTOM | wx.EXPAND | wx.ALL, 1)

        self.kwMap = None

        # Stick some initial "0" values in the timeline
        self.draw_timeline_zero()

        # Set the focus on the Waveform widget
        self.waveform.SetFocus()

        self.SetSizer(box)
        
        self.SetAutoLayout(True)
        self.Fit()
        self.Layout()

        # We need to adjust the screen position on the Mac.  I don't know why.
#        if "__WXMAC__" in wx.PlatformInfo:
#            pos = self.GetPosition()
#            self.SetPosition((pos[0]-20, pos[1]-25))

        # GraphicsControlClass is not Sizer or Constraints aware, so we'll to screen positioning the old fashion way, with a
        # Resize Event
        wx.EVT_SIZE(self, self.OnSize)

        # Idle event (draws when idle to prevent multiple redraws while resizing, which are too slow)
        wx.EVT_IDLE(self, self.OnIdle)

        # Let's also capture key presses so we can control video playback during transcription
        # NOTE that we assign this event to the waveform, not to self.
        wx.EVT_KEY_DOWN(self.waveform, self.OnKeyDown)

    def SetVisualizationObject(self, visualizationObject):
        """ Set the Object to be visualized """
        # Reset the Zoom Information, so we don't carry it over from the last visualization
        self.zoomInfo = [(0, -1)]
        if self.kwMap != None:
            # Delete the current keyword visualization object
            del(self.kwMap)
            # Set the reference to the keyword visualizatoin object to None so we don't get confused.
            self.kwMap = None

        # If a Visualization Object is passed in ...
        if visualizationObject != None:
            # ... load the current Visualization Settings for known data objects
            self.LoadVisualizationInfo((type(visualizationObject), visualizationObject.number))
        # Set the object to be Visualized (which CAN be None!)
        self.VisualizationObject = visualizationObject
        # Set the type of Visualization based on the object type.
        # If None is passed in ...
        if self.VisualizationObject == None:
            # ... then we have nothing to visualize
            self.VisualizationType = None
        # If we have a Document or a Quote ...
        elif isinstance(self.VisualizationObject, Document.Document) or \
             isinstance(self.VisualizationObject, Quote.Quote):
            # ... we need a Text-Keyword Visualization
            self.VisualizationType = 'Text-Keyword'
            self.redrawWhenIdle = True

            self.createClip.Show(False)
            self.loop.Show(False)
            self.lbl_Time.SetLabel(_("Pos:"))
            self.btn_Current.Enable(False)
            self.lbl_Current_Time.SetLabel("0")
            self.btn_Selected.Enable(False)
            self.lbl_Selected_Time.SetLabel("0")
            self.lbl_Time_Time.SetLabel("0")

        # If we have an Episode ...
        elif isinstance(self.VisualizationObject, Episode.Episode):
            # ... we should use the configuration setting for Media Visualization Style
            self.VisualizationType = TransanaGlobal.configData.visualizationStyle
            # Load the waveform for the appropriate media files with its current start and length.
            self.load_image('Episode', visualizationObject.media_filename, visualizationObject.additional_media_files,
                            0, 0, self.VisualizationObject.tape_length)
            self.createClip.Show(True)
            self.loop.Show(True)
            self.lbl_Time.SetLabel(_("Time:"))
            self.btn_Current.Enable(True)
            self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(0))
            self.btn_Selected.Enable(True)
            self.lbl_Selected_Time.SetLabel(Misc.time_in_ms_to_str(0))
            self.lbl_Time_Time.SetLabel(Misc.time_in_ms_to_str(0))

        # If we have a Clip ...
        elif isinstance(self.VisualizationObject, Clip.Clip):
            # ... we should use the configuration setting for Media Visualization Style
            self.VisualizationType = TransanaGlobal.configData.visualizationStyle

            # Load the waveform for the appropriate media files with its current start and length.
            self.load_image('Clip', visualizationObject.media_filename, visualizationObject.additional_media_files,
                            visualizationObject.offset, visualizationObject.clip_start,
                            visualizationObject.clip_stop - visualizationObject.clip_start)

            self.createClip.Show(True)
            self.loop.Show(True)
            self.lbl_Time.SetLabel(_("Time:"))
            self.btn_Current.Enable(False)
            self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(0))
            self.btn_Selected.Enable(True)
            self.lbl_Selected_Time.SetLabel(Misc.time_in_ms_to_str(0))
            self.lbl_Time_Time.SetLabel(Misc.time_in_ms_to_str(0))

        # If we have another object type ...
        else:
            # ... then we have nothing to visualize
            self.VisualizationType = None

            print "VisualizationWindow.SetVisualizationObject():  Invalid Object Type"

            # Raise an exception!
            raise TransanaExceptions.ProgrammingError("VisualizationWindow.SetVisualizationObject():  Invalid Object Type")

        # ... clear the current waveform ...
        self.waveform.Clear()

        # ... signal that the redraw is done ...
        self.redrawWhenIdle = True

    def LoadVisualizationInfo(self, objKey):
        """ Load existing Visualization information when changing main objects in Transana's main interface """
        # If there is an entry with the requested key ...
        if objKey in self.VisualizationInfo.keys():
            # ... provide the Visualization Information required
            self.zoomInfo = self.VisualizationInfo[objKey]['zoomInfo']
            self.kwMap = self.VisualizationInfo[objKey]['kwMap']

    def SaveVisualizationInfo(self, objKey):
        """ Save existing Visualization information when changing main objects in Transana's main interface """
        # If there is an entry with the requested key ...
        if objKey in self.VisualizationInfo.keys():
            # ... update the Visualization Information required
            self.VisualizationInfo[objKey]['zoomInfo'] = self.zoomInfo
            self.VisualizationInfo[objKey]['kwMap'] = self.kwMap
        # If not ...
        else:
            # ... create a dictionary with the appropriate key ...
            self.VisualizationInfo[objKey] = {}
            # ... adn add the Visualization Information required
            self.VisualizationInfo[objKey]['zoomInfo'] = self.zoomInfo
            self.VisualizationInfo[objKey]['kwMap'] = self.kwMap

    def DeleteVisualizationInfo(self, objKey):
        """ Delete existing Visualization information when removing objects from Transana's main interface """
        # If there is an entry with the requested key ...
        if objKey in self.VisualizationInfo.keys():
            # ... delete that key's information
            del(self.VisualizationInfo[objKey])

    def Refresh(self):
        """ Called to indicate that the Visualization needs to be re-drawn """
        # The Visualization Type may have changed since it was originally set.
        if isinstance(self.VisualizationObject, Episode.Episode) or \
           isinstance(self.VisualizationObject, Clip.Clip):
            # ... we should use the configuration setting for Media Visualization Style
            self.VisualizationType = TransanaGlobal.configData.visualizationStyle
        # Clear the existing Visualization
        self.ClearVisualization()
        # Signal that the Visualization should be re-drawn during Idle time
        self.redrawWhenIdle = True
        # Call the OnIdle handler's method explicitly, so this is in fact handled in real time rather than waiting for Idle time.
        self.OnIdle(None)

    def UpdateKeywordVisualization(self):
        # Only do something if we've got the Keyword or Hybrid visualization and there is a data object currently loaded
        if (self.VisualizationType in ['Keyword', 'Hybrid', 'Text-Keyword']) and \
           (self.ControlObject.currentObj != None) and \
           (self.kwMap != None):
            # Update the Visualization's underlying data
            self.kwMap.UpdateKeywordVisualization()
            # Signal that height can be reset
            self.heightIsSet = False
            # Call the resize Height method
            self.resizeKeywordVisualization()
            # If we are doing a Hybrid visualization ...
            if self.VisualizationType in ['Hybrid', 'Text-Keyword']:
                # ... we need to redraw the visualization (Hybrid: so we don't lose the Waveform!) (Text: to handle position changes)
                self.redrawWhenIdle = True

    def OnIdle(self, event):
        """ Use Idle Time to handle the drawing in this control """
        # Check to see if the waveform control needs to be redrawn.  Under the new Media Player, GetMediaLength takes a while to
        # be set, so we should wait for that too.

        if self.redrawWhenIdle  and (not self.ControlObject.shuttingDown) and \
           ((self.ControlObject.GetMediaLength() > 0) or (self.VisualizationType == 'Text-Keyword')):

            if not self.VisualizationType == 'Text-Keyword':
                # Remove old Waveform Selection and Cursor data
                self.waveform.ClearTransanaSelection()

            # Let's make sure we're not working without a waveform, unless we have Text
            if (len(self.waveFilename) > 0) and (self.VisualizationType != 'Text-Keyword'):
                # Let's detect and correct media length problems present in new Episodes that prevent proper visualization display.
                # If media file length was not yet defined when the waveFilename object was populated ...
                if self.waveFilename[0]['length'] == 0:
                    # If currentObj is NOT loaded (as for Synchronize.py), or if it has a tapelength of 0 ...
                    if (self.ControlObject.currentObj == None) or (self.ControlObject.currentObj.tape_length == 0):
                        # ... then use the ControlObject's GetMediaLength() method to approximate length
                        self.waveFilename[0]['length'] = self.ControlObject.GetMediaLength()
                    # Otherwise ...
                    else:
                        # ... we need to update all of the waveFilename object's length values from the Control Object's CurrentObject.
                        self.waveFilename[0]['length'] = self.ControlObject.currentObj.tape_length

                    # Iterate through the additional filenames (skipping the first, which we just took care of) ...
                    for x in range(1, len(self.waveFilename)):
                        # If the length is not known ...
                        if self.ControlObject.currentObj.additional_media_files[x - 1]['length'] <= 0:
                            # ... then use the ControlObject's GetMediaLength() method to approximate length
                            self.waveFilename[x]['length'] = self.ControlObject.GetMediaLength()
                        # Otherwise ...
                        else:
                            # ... use the actual length of the media file.
                            self.waveFilename[x]['length'] = self.ControlObject.currentObj.additional_media_files[x - 1]['length']

            # If no main object if defined ...
            if (self.ControlObject.currentObj == None):
                # ... clear the current waveform ...
                self.waveform.Clear()
                # ... signal that the redraw is done ...
                self.redrawWhenIdle = False
                # ... and exit this method.  We're done here.
                return

            # Initialize
            waveformGraphicImage = None
            if self.VisualizationType in ['Waveform', 'Hybrid']:
                # Create the appropriate Waveform Graphic
                try:
                    # The Mac can't handle Unicode WaveFilenames at this point.  We need to upgrade to Python 2.4 for that.
                    # Let's temporarily take care of that.
                    if ('wxMac' in wx.PlatformInfo):
                        for flnm in self.waveFilename:
                            if isinstance(flnm['filename'], unicode):
                                flnm['filename'] = flnm['filename'].encode('utf8')
                    # We populate the keyword visualization differently for an episode and a clip.
                    if type(self.ControlObject.currentObj) == Episode.Episode:
                        # The starting point of the Keyword Visualization is the start time in the top of the zoomInfo stack!
                        start = self.zoomInfo[-1][0]
                        # The ending point of the Keyword Visualization is the end time in the top of the zoomInfo stack,
                        # unless that value is <= 0, in which case it's the media's full length.
                        if self.zoomInfo[-1][1] <= 0:
                            length = self.ControlObject.GetMediaLength(True)
                        else:
                            length = self.zoomInfo[-1][1]

#                        print "VisualizationWindow.OnIdle(2a):", start, length, self.zoomInfo
                        
                    elif type(self.ControlObject.currentObj) == Clip.Clip:

                        if False:
                            # Set the current global video selection based on the Clip.
                            self.ControlObject.SetVideoSelection(self.ControlObject.currentObj.clip_start, self.ControlObject.currentObj.clip_stop) 
                            # we have to know the right start and end points for the waveform.
                            start = self.ControlObject.VideoStartPoint
                            length = self.ControlObject.GetMediaLength()
                        else:
                            start = self.zoomInfo[-1][0]
                            length = max(self.zoomInfo[-1][1], 500)
                            # Set the current global video selection based on the Clip.
                            self.ControlObject.SetVideoSelection(start, self.ControlObject.currentObj.clip_stop)

                    # Get the status of the VideoWindow's checkboxes
                    checkboxData = self.ControlObject.GetVideoCheckboxDataForClips(start)
                    # For each media file (with its corresponding checkboxes) ...
                    for x in range(len(checkboxData)):
                        # ... add a "Show" variable to the waveform filename dictionary in the waveFilename list
                        #     that indicates if that waveform should be shown 
                        self.waveFilename[x]['Show'] = checkboxData[x][1]
                    # Create the waveform graphic
                    waveformGraphicImage = WaveformGraphic.WaveformGraphicCreate(self.waveFilename, ':memory:', start, length, self.waveform.canvassize, style='waveform')
                    # If a waveform graphic was created ...
                    if waveformGraphicImage != None:
                        # ... clear the waveform
                        self.waveform.Clear()
                        # ... and set the image as the waveform background
                        self.waveform.SetBackgroundGraphic(waveformGraphicImage)
                        # Determine NEW Frame Size
                        (width, height) = self.GetSize()
                        # Draw the TimeLine values
                        self.draw_timeline(start, length)
                    # If we can't create the graphic ...
                    else:
                        # ... try to function without one.
                        self.redrawWhenIdle = False

                # A bug in Python 2.3.5 causes a RuntimeError with some wave files if Unicode filenames are used.
                # This should be fixed in Python 2.4.2, but we'll leave this code here to prevent ugly errors if it
                # does occur.
                except RuntimeError, e:
                    self.ClearVisualization()
                    self.redrawWhenIdle = False
                except:
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
                if (self.VisualizationType == 'Hybrid'):
                    # Get the waveform Bitmap and convert it to an Image
                    hybridWaveform = waveformGraphicImage
                    if hybridWaveform != None:
                        # Rescale the image so that it matches the size alloted for the Waveform (HYBRIDOFFSET)
                        hybridWaveform.Rescale(hybridWaveform.GetWidth(), HYBRIDOFFSET)

            if self.VisualizationType in ['Keyword', 'Hybrid']:
                # Clear the Visualization
                self.waveform.Clear()

                # Enable the Filter button
                self.filter.Enable(True)
                # If there's an existing Keyword Visualization ...
                if self.kwMap != None:
                    # ... remember the values for the Clip List
                    filteredClipList = self.kwMap.clipFilterList[:]
                    unfilteredClipList = self.kwMap.clipList[:]
                    # ... remember the values for the Snapshot List
                    filteredSnapshotList = self.kwMap.snapshotFilterList[:]
                    unfilteredSnapshotList = self.kwMap.snapshotList[:]
                    # ... remember the values for the filtered Keyword List
                    filteredKeywordList = self.kwMap.filteredKeywordList[:]
                    # ... remember the values from the unfiltered keyword list
                    unfilteredKeywordList = self.kwMap.unfilteredKeywordList[:]
                    # ... remember the keyword color list too.
                    keywordColorList = self.kwMap.keywordColors
                    # ... and remember the configuration name
                    configName = self.kwMap.configName
                    # Delete the current keyword visualization object
                    del(self.kwMap)
                    # Set the reference to the keyword visualizatoin object to None so we don't get confused.
                    self.kwMap = None
                # If we're creating a brand new Keyword Visualization ...
                else:
                    # Initialize the Clip Lists
                    filteredClipList = []
                    unfilteredClipList = []
                    # Initialize teh Snapshot Lists
                    filteredSnapshotList = []
                    unfilteredSnapshotList = []
                    # Initialize the filtered keyword list ...
                    filteredKeywordList = []
                    # ... the unfiltered keyword list ...
                    unfilteredKeywordList = []
                    # ... and the keyword color list.
                    keywordColorList = None
                    # ... and the configuration name
                    configName = ''
                # If we're creating a Hybrid Visualization ...
                if self.VisualizationType == 'Hybrid':
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
                        # Changed from (self.ControlObject.GetVideoStartPoint(), self.ControlObject.GetMediaLength(True))
                        # because Locate Clip in Episode with Keyword Visualization had the WRONG START POINT!
                        self.zoomInfo[0] = (0, self.ControlObject.GetMediaLength(True))
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
                                             filteredClipList = filteredClipList, unfilteredClipList = unfilteredClipList, \
                                             filteredSnapshotList = filteredSnapshotList, unfilteredSnapshotList = unfilteredSnapshotList, \
                                             filteredKeywordList = filteredKeywordList, unfilteredKeywordList = unfilteredKeywordList, \
                                             keywordColors = keywordColorList, configName=configName, loadDefault=self.loadDefault)

                    # Draw the TimeLine values
                    self.draw_timeline(kwMapStartPoint, kwMapEndPoint - kwMapStartPoint)

                elif type(self.ControlObject.currentObj) == Clip.Clip:
                    # If we're working with a Clip, we need some information from it's source episode.  Let's get the Episode.
                    try:
                        tmpEpisode = Episode.Episode(self.ControlObject.currentObj.episode_num)
                    except TransanaExceptions.RecordNotFoundError, e:
                        tmpEpisode = Episode.Episode()
                        tmpEpisode.series_id = 'None'
                        tmpEpisode.id = 'None'

                    start = self.zoomInfo[-1][0]
                    length = max(self.zoomInfo[-1][1], 500)
                    # Set the current global video selection based on the Clip.
                    self.ControlObject.SetVideoSelection(start, self.ControlObject.currentObj.clip_stop) 

                    # Set up the embedded Keyword Visualization, sending it all the data it needs so it can draw or redraw itself.
                    self.kwMap.SetupEmbedded(self.ControlObject.currentObj.episode_num, tmpEpisode.series_id, \
                                             tmpEpisode.id, start, start + length, \
                                             filteredClipList = filteredClipList, unfilteredClipList = unfilteredClipList, \
                                             filteredSnapshotList = filteredSnapshotList, unfilteredSnapshotList = unfilteredSnapshotList, \
                                             filteredKeywordList = filteredKeywordList, unfilteredKeywordList = unfilteredKeywordList, \
                                             keywordColors = keywordColorList, clipNum=self.ControlObject.currentObj.number,
                                             configName=configName, loadDefault=self.loadDefault)

                    # Draw the TimeLine values
                    self.draw_timeline(start, length)

                elif self.ControlObject.currentObj != None:
                    self.waveform.AddText('Keyword Visualization - %s not implemented.' % type(self.ControlObject.currentObj), 5, 5)

                # By this point, we've already loaded the default and don't need to do it again.
                self.loadDefault = False

                # The Keyword / Hybrid visualization height can be self-adjusting.  Let's call that function.
                self.resizeKeywordVisualization()

                # If we're bulding a Hybrid visualization, so far we've created and stored the waveform, then
                # wiped it out in favor of an offset Keyword visualization.  Here, we combine the two
                # visualizations!  If waveforFilename has been cleared, we're waveforming and shouldn't do this yet!
                if (self.VisualizationType == 'Hybrid'):
                    # Here's a trick.  By setting the waveform's backgroundImage but NOT setting the
                    # backgroundGraphicName, you can add a background image to the GraphicsControlClass
                    # that does not resize to fill the image.  The current offset Keyword visualization
                    # with the waveform overlaid as a background image works pretty well!
                    self.waveform.backgroundImage = hybridWaveform

            if self.VisualizationType == 'Text-Keyword':
                # Clear the Visualization
                self.waveform.Clear()

                # Enable the Filter button
                self.filter.Enable(True)
                # If there's an existing Keyword Visualization ...
                if self.kwMap != None:
                    # ... remember the values for the Quote List
                    filteredQuoteList = self.kwMap.quoteFilterList[:]
                    # We DON'T remember the UnfilteredQuoteList because the Character Positions might have changed!
                    unfilteredQuoteList = []
##                    # ... remember the values for the Snapshot List
##                    filteredSnapshotList = self.kwMap.snapshotFilterList[:]
##                    unfilteredSnapshotList = self.kwMap.snapshotList[:]
                    # ... remember the values for the filtered Keyword List
                    filteredKeywordList = self.kwMap.filteredKeywordList[:]
                    # ... remember the values from the unfiltered keyword list
                    unfilteredKeywordList = self.kwMap.unfilteredKeywordList[:]
                    # ... remember the keyword color list too.
                    keywordColorList = self.kwMap.keywordColors
                    # ... and remember the configuration name
                    configName = self.kwMap.configName
                    # Delete the current keyword visualization object
                    del(self.kwMap)
                    # Set the reference to the keyword visualizatoin object to None so we don't get confused.
                    self.kwMap = None
                # If we're creating a brand new Keyword Visualization ...
                else:
                    # Initialize the Quote Lists
                    filteredQuoteList = []
                    unfilteredQuoteList = []
##                    # Initialize the Snapshot Lists
##                    filteredSnapshotList = []
##                    unfilteredSnapshotList = []
                    # Initialize the filtered keyword list ...
                    filteredKeywordList = []
                    # ... the unfiltered keyword list ...
                    unfilteredKeywordList = []
                    # ... and the keyword color list.
                    keywordColorList = None
                    # ... and the configuration name
                    configName = ''

                # Create a Keyword Visualization object as an embedded graphic, not a free-standing report.
                self.kwMap = KeywordMapClass.KeywordMap(self, -1, "", embedded=True, topOffset=0)

                # We populate the keyword visualization differently for a document and a quote.
                if type(self.ControlObject.currentObj) == Document.Document:
                    # The starting point of the Keyword Visualization is the start position in the top of the zoomInfo stack!
                    kwMapStartChar = self.zoomInfo[-1][0]
                    # The ending point of the Keyword Visualization is the end position in the top of the zoomInfo stack,
                    # unless that value is 0, in which case it's the document's full length.
                    if self.zoomInfo[-1][1] <= 0:
                        kwMapEndChar = self.ControlObject.GetDocumentLength()
                    else:
                        kwMapEndChar = self.zoomInfo[-1][1]

                    # Get the total length of the Document
                    totalLength = self.ControlObject.GetDocumentLength()

                    # Documents are different.  If they are being edited, all of the keyword positioning data from
                    # the database will be incorrect!!  That's because keyword positioning is based on character
                    # position, which changes during edits, instead of time code, which does not change during edits.
                    #
                    # As a result, we need to pass in the LIVE Document Object rather than loading it from the database
                    # the way we do with Episodes and Clips.

                    # Set up the embedded Text Keyword Visualization, sending it all the data it needs so it can draw or redraw itself.
                    self.kwMap.SetupTextEmbedded(self.ControlObject.currentObj,
                                                 kwMapStartChar,
                                                 kwMapEndChar,
                                                 totalLength,
                                                 filteredQuoteList = filteredQuoteList,
                                                 unfilteredQuoteList = unfilteredQuoteList, 
##                                                 filteredSnapshotList = filteredSnapshotList,
##                                                 unfilteredSnapshotList = unfilteredSnapshotList, 
                                                 filteredKeywordList = filteredKeywordList,
                                                 unfilteredKeywordList = unfilteredKeywordList, 
                                                 keywordColors = keywordColorList,
                                                 configName=configName,
                                                 loadDefault=self.loadDefault)

                    # Draw the TimeLine values
                    self.draw_timeline_text(kwMapStartChar, kwMapEndChar)

                elif type(self.ControlObject.currentObj) == Quote.Quote:
                    # Determine the initial start and length for the Quote
                    start = self.zoomInfo[-1][0]
                    # Minimum of 4 characters in the Visualization
                    length = max(self.zoomInfo[-1][1], 4)

                    # Set up the embedded Keyword Visualization, sending it all the data it needs so it can draw or redraw itself.
                    # A Quote gets the Source Document object rather than the Quote object itself, if possible!
                    if self.ControlObject.currentObj.source_document_num > 0:
                        try:
                            sourceObj = Document.Document(self.ControlObject.currentObj.source_document_num)
                        except TransanaExceptions.recordNotFoundError:
                            sourceObj = self.ControlObject.currentObj
                    else:
                        sourceObj = self.ControlObject.currentObj
                    self.kwMap.SetupTextEmbedded(sourceObj,
                                                 self.ControlObject.currentObj.start_char,
                                                 self.ControlObject.currentObj.end_char,  #  - self.ControlObject.currentObj.start_char,
                                                 self.ControlObject.currentObj.end_char,
                                                 filteredQuoteList = filteredQuoteList,
                                                 unfilteredQuoteList = unfilteredQuoteList,
##                                                 filteredSnapshotList = filteredSnapshotList,
##                                                 unfilteredSnapshotList = unfilteredSnapshotList,
                                                 filteredKeywordList = filteredKeywordList,
                                                 unfilteredKeywordList = unfilteredKeywordList,
                                                 keywordColors = keywordColorList,
                                                 quoteNum=self.ControlObject.currentObj.number,
                                                 configName=configName,
                                                 loadDefault=self.loadDefault)

                    # Draw the TimeLine
                    self.draw_timeline_text(self.ControlObject.currentObj.start_char,
                                            self.ControlObject.currentObj.end_char)

                # If we have some other type of object ...
                elif self.ControlObject.currentObj != None:
                    # ... add the error message to the visualization!
                    self.waveform.AddText('Text-Keyword Visualization - %s not implemented.' % type(self.ControlObject.currentObj), 5, 5)

                # By this point, we've already loaded the default and don't need to do it again.
                self.loadDefault = False

                # The Keyword / Hybrid visualization height can be self-adjusting.  Let's call that function.
                self.resizeKeywordVisualization()

            # Signal that the redraw is complete, and does not need to be done again until this flag is altered.
            self.redrawWhenIdle = False

            # If we have a Text-Keyword Visualization ...
            if self.VisualizationType == 'Text-Keyword':
                # If we DO NOT have a SELECTION ...
                if (self.endPoint - self.startPoint <= 0) and not isinstance(self.ControlObject.currentObj, Quote.Quote):
                    if (self.waveformUpperLimit - self.waveformLowerLimit) > 0:
                        # Determine the current/new horizontal position within the visualization
                        pos = ((float(self.startPoint - self.waveformLowerLimit)) / (self.waveformUpperLimit - self.waveformLowerLimit))
                    else:
                        pos = 1
                    # Draw the Waveform Cursor
                    self.waveform.DrawCursor(pos)
                # If we DO have a SELECTION ...
                else:
                    # ... assign the start and end points in waveform ...
                    self.waveform.startTime = self.startPoint
                    self.waveform.endTime = self.endPoint
                    # ... and signal that we want the selection to be shown
                    self.waveform.reSetSelection = True
            # If we have a media-based visualization ...
            else:
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
                # If resetVideoPosition == 0 and GetVideoStartPoint() > 0, we may have just located a Clip in the Episode.
                elif self.ControlObject.GetVideoStartPoint() > 0:
                    # Make sure the Position Cursor is drawn on the Waveform
                    self.UpdatePosition(self.ControlObject.GetVideoPosition())

    def resizeKeywordVisualization(self):
        """ The Keyword Visualization (and Hybrid) should auto-resize under some circumstances.  This method implements that. """

        # Disabled for Transana 3.0.

##        print "VisualizationWindow.resizeVisualization() disabled for Transana 3.0"
        
        return


        # If Auto Arrange is ON and the height has not already been set, and there is only ONE media file ...
        # (We don't resize the Visualization if there are multiple media files.)
        if TransanaGlobal.configData.autoArrange and not self.heightIsSet and (len(self.waveFilename) == 1):
            # ... start by getting the current dimensions of the visualization window.
            (a, b, c, d) = self.GetDimensions()
            (x, y, w, h) = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
            # We need to make different adjustments based on whether we're on the Mac or Windows.
            if "wxMac" in wx.PlatformInfo:
                # On Mac, the window height is the smaller of the Keyword Visualization + 100 for the rest of the
                # window or 2/3 of the screen height
                newHeight = min(self.kwMap.GetKeywordCount()[1] + 100, 2 * h / 3)  # wx.ClientDisplayRect()
            else:
                # On Windows, the window height is the smaller of the Keyword Visualization + 142 for the rest of the
                # window or 2/3 of the screen height
                newHeight = min(self.kwMap.GetKeywordCount()[1] + 142, 2 * h / 3)  # wx.ClientDisplayRect()
            # Let's say that 1/4 of the screen is the minimum Waveform height!
            newHeight = max(newHeight, round(h / 4))  # wx.ClientDisplayRect()

            # The Hybrid Visualization was losing the waveform when the Filter Dialog was called.  So when this happens ...
            if self.VisualizationType == 'Hybrid':
                # ... signal that the whole visualization has to be re-drawn!
                self.redrawWhenIdle = True

            # now let's adjust the window sizes for the main Transana interface.
            self.ControlObject.UpdateWindowPositions('Visualization', c + a, YUpper = newHeight + b)
            # once we do this, we don't need to do it again unless something changes.
            self.heightIsSet = True
        
    def OnKeyDown(self, event):
        """ Captures Key Events to allow this window to control video playback during transcription. """

        # See if the Control Object wants to handle the keys that were pressed
        if self.ControlObject.ProcessCommonKeyCommands(event):
            # If Ctrl-A or Ctrl-F are pressed, we need to move the Visualization Window's starting point!
            if event.ControlDown() and (event.GetKeyCode() in [ord("A"), ord("F")]):
                # Set the start point to the current video position
                self.startPoint = self.ControlObject.GetVideoPosition()
                # Set the end point to the end of the document.
                self.endPoint = 0
            # If so, we're done here
            return

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

            # Ctrl-K -- Create Quick Clip
            if (c == ord("K")) and event.ControlDown():
                # Ask the Control Object to create a Quick Clip
                self.ControlObject.CreateQuickClip()

            # Cursor Left ...
            elif c in [wx.WXK_LEFT, wx.WXK_NUMPAD_LEFT]:
                # ... moves the video the equivalent of 1 pixel earlier in the video
                if currentPos > self.waveformLowerLimit - timePerPixel:
                    self.startPoint = currentPos - timePerPixel
                    self.endPoint = 0
                    self.ControlObject.SetVideoSelection(self.startPoint, self.endPoint)
                    self.waveform.SetFocus()

            # Cursor Right ...
            elif c in [wx.WXK_RIGHT, wx.WXK_NUMPAD_RIGHT]:
                # ... moves the video the equivalent of 1 pixel later in the video
                if currentPos < self.waveformUpperLimit + timePerPixel:
                    self.startPoint = currentPos + timePerPixel
                    self.endPoint = 0
                    self.ControlObject.SetVideoSelection(self.startPoint, self.endPoint)
                    self.waveform.SetFocus()

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
        if self.VisualizationType == 'Hybrid':
            # ... signal that the whole visualization has to be re-drawn!
            #     Setting the reset parameter to false signals this is a Hybrid Visualization, preventing the
            #     waveform part of the Hybrid Visualization from being erased.
            self.kwMap.UpdateKeywordVisualization(reset=False)
        else:
            # If we'd cleared the waveform before calling OnFilter, we'd be done.  But that clears the graphic
            # BEFORE bringing up the filter dialog, which I didn't like.  So even though OnFilter re-draws the
            # visualization (albeit incorrectly because we didn't call waveform.Clear()), we'll now clear the
            # waveform and re-draw the keyword visualization.
            self.waveform.Clear()
            self.kwMap.DrawGraph()
        # Remember the new visualization zoom and filter information
        self.SaveVisualizationInfo((type(self.VisualizationObject), self.VisualizationObject.number))
        
    def OnZoomIn(self, event):
        """ Zoom in on a portion of the Waveform Diagram """
        if (self.startPoint < self.endPoint):
            if self.VisualizationType == 'Text-Keyword':
                # Keep track of the new position.  This allows the user to zoom back out in the same steps used to zoom in
                self.zoomInfo.append((int(self.startPoint), int(self.endPoint)))

            else:
                # Keep track of the new position.  This allows the user to zoom back out in the same steps used to zoom in
                self.zoomInfo.append((int(self.startPoint), int(self.endPoint - self.startPoint)))
                # Limit video playback to the selected part of the media by setting the VideoStartPoint and VideoEndPoint in the
                # Control Object
                self.ControlObject.SetVideoSelection(self.startPoint, self.endPoint)
                # Draw the TimeLine values
                self.draw_timeline(self.ControlObject.VideoStartPoint, self.ControlObject.GetMediaLength())
            # Signal that the Waveform Graphic should be updated
            self.redrawWhenIdle = True
            # Remember the new visualization zoom and filter information
            self.SaveVisualizationInfo((type(self.VisualizationObject), self.VisualizationObject.number))

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

            self.redrawWhenIdle = True
            # Clear the Start and End points of the Visualization Selection
            self.startPoint = self.ControlObject.VideoStartPoint
            self.endPoint = self.ControlObject.VideoStartPoint + self.ControlObject.GetMediaLength()
            # Remember the new visualization zoom and filter information
            self.SaveVisualizationInfo((type(self.VisualizationObject), self.VisualizationObject.number))

    def OnZoom100(self, event):
        """ Zoom all the way out on the Waveform Diagram """
        if len(self.zoomInfo) > 1:
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
            # Redraw the waveform
            self.redrawWhenIdle = True
            # Remember the new visualization zoom and filter information
            self.SaveVisualizationInfo((type(self.VisualizationObject), self.VisualizationObject.number))

    def OnCreateClip(self, event):
        """ Create a Transcript-less Clip """
        # If there's no selection (endPoint could be 0 or -1!) ...
        if self.endPoint <= 0:
            # then we can't make a clip
            return

        # Create a Transcript-less Clip.  (This routine can handle both Standard and Quick Clips, depending on
        # what is selected in the Database Tree.)
        self.ControlObject.CreateTranscriptlessClip()

    def OnScroll(self, direction):
        """ Scroll the visualization in the direction specified """
        # If the window had been zoomed ...
        if len(self.zoomInfo) > 1:
            # Let's rename some data representations to make this a bit simpler
            # The left side of the visualization is the LAST zoomInfo element's FIRST value
            currLeft = self.zoomInfo[-1][0]
            # The width of the visualization is the LAST zoomInfo element's SECOND value
            currWidth = self.zoomInfo[-1][1]
            # The right side of the visualization is the left side plus the width!!
            currRight = currLeft + currWidth
            # The total width of the media file is the FIRST zoomInfo element's SECOND value
            totalWidth = self.zoomInfo[0][1]
            # If a Scroll Left has been requested ...
            if (direction == 'Left'):
                # If the is 75% of the width available to the left ...
                if currLeft > int(currWidth * 0.75):
                    # ... scroll 75% to the left by shifting the selected start and end points and adjusting zoomInfo appropriately
                    self.startPoint -= int(currWidth * 0.75)
                    self.endPoint -= int(currWidth * 0.75)
                    self.zoomInfo[-1] = (self.zoomInfo[-1][0] - int(currWidth * 0.75), self.zoomInfo[-1][1])
                # If there is less than 75% of the width available ...
                else:
                    # ... scroll to the left edge by setting the start to 0, the end to width, and adjusting zoomInfo appropriately
                    self.startPoint = 0
                    self.endPoint = currWidth
                    self.zoomInfo[-1] = (0, self.zoomInfo[-1][1])
                # Signal that we need to re-draw the visualization
                self.redrawWhenIdle = True
                
            # If a Scroll Right has been requested...
            elif (direction == 'Right'):
                # If there's a full 75% of width available to shift right ...
                if currRight < totalWidth - int(currWidth * 0.75):
                    # ... scroll 75% to the right by shifting the selected start and end points and adjusting zoomInfo appropriately
                    self.startPoint = currLeft + int(currWidth * 0.75)
                    self.endPoint += int(currWidth * 0.75)
                    self.zoomInfo[-1] = (self.startPoint, self.zoomInfo[-1][1])
                # If there is less than 75% of the width available ...
                else:
                    # ... scroll to the right edge by setting the start to total width - current width, the end to total width,
                    # and adjusting zoomInfo appropriately
                    self.startPoint = totalWidth - currWidth
                    self.endPoint = totalWidth
                    self.zoomInfo[-1] = (self.startPoint, self.zoomInfo[-1][1])

                # Signal that we need to re-draw the visualization
                self.redrawWhenIdle = True

    def OnLoop(self, event):
        """ Click the "Loop" button to initiate a playback loop """
        # If a data object is currently loaded ...
        if self.ControlObject.currentObj != None:
            # If no selection endpoint is defined ...
            if self.endPoint <= 0:
                # ... then set the endpoint for the earlier of the media end (episode or clip end) or 5 seconds from start
                self.endPoint = min(self.zoomInfo[0][0] + self.zoomInfo[0][1], self.startPoint + 5000)
                # Now check for self.zoomInfo[0] having been (0, 0), which would be a problem!
                if self.endPoint == 0:
                    self.endPoint = self.startPoint + 5000
                # Set the video selection to include this new end point
                self.ControlObject.SetVideoSelection(self.startPoint, self.endPoint)

            # If we're using a LtR language ...
            if TransanaGlobal.configData.LayoutDirection == wx.Layout_LeftToRight:
                # Signal the Control Object to start or stop looped playback
                self.ControlObject.PlayLoop(self.loop.GetValue())
            # If we're using a RtL language ...
            else:
                # If we're currently looping ...
                if self.looping:
                    # Set the image for the "un-pressed" state ...
                    self.loop.SetBitmap(TransanaGlobal.GetImage(TransanaImages.loop_up))
                    # ... and signal that we're STOPPING looping
                    self.looping = False
                # If we're NOT currently looping ...
                else:
                    # Set the image for the "pressed" state ...
                    self.loop.SetBitmap(TransanaGlobal.GetImage(TransanaImages.loop_down))
                    # ... and signal that we're STARTING to loop
                    self.looping = True
                # Signal the Control Object to start or stop looped playback
                self.ControlObject.PlayLoop(self.looping)
        # if no data object is currently loaded ...
        else:
            if TransanaGlobal.configData.LayoutDirection == wx.Layout_LeftToRight:
                # then forcibly reject the button press by un-pressing the button.
                self.loop.SetValue(False)
            else:
                # Define the image for the "un-pressed" state
                self.loop.SetBitmap(TransanaGlobal.GetImage(TransanaImages.loop_up))
                self.looping = False

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
        self.startPoint = self.zoomInfo[0][0]
        self.endPoint = self.zoomInfo[0][1]
        # There's a weird bug where RTF files are not loaded correctly if self.endPoint is passed here instead of -1.
        # Since we're CLEARING the visualization, I can't think of any harm that would be done by passing -1 instead.
        self.ControlObject.SetVideoSelection(self.startPoint, -1, UpdateSelectionText=False)
        # Clear the waveform itself
        self.waveform.Clear()
        # Remove old Waveform Selection and Cursor data
        self.waveform.ClearTransanaSelection()
        self.filter.Enable(False)
        # Clear the Time Line
        self.draw_timeline_zero()
        # Clear all time labels
        # If we have a Transcript currently selected, not a Document
        if self.ControlObject.GetCurrentItemType() == 'Transcript':
            self.lbl_Time_Time.SetLabel(Misc.time_in_ms_to_str(0))
            self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(0))
            self.lbl_Selected_Time.SetLabel(Misc.time_in_ms_to_str(0))
        else:
            self.lbl_Time_Time.SetLabel("0")
            self.lbl_Current_Time.SetLabel("0")
            self.lbl_Selected_Time.SetLabel("0")
        self.lbl_Total_Time.SetLabel(Misc.time_in_ms_to_str(0))
        # When clearing, we need to reset loading the Default
        self.loadDefault = True
        # Signal that things should be redrawn
        self.redrawWhenIdle = True

    def ClearVisualizationSelection(self):
        """ Clear the Selection Box from the Visualization Window """
        self.waveform.ClearTransanaSelection()

    def load_image(self, imgType, filename, additionalFiles, offset, mediaStart, mediaLength):
        """ Causes the proper visualization to be displayed in the Visualization Window when a Video File is loaded. """
        # We need to know the minimum offset value.  We know there will be a zero value
        minVal = 0
        # Add the main file data to the file list
        filenameList = [{'filename' : filename, 'offset' : offset, 'length' : mediaLength}]
        # If we have an episode, we need to determine the VIRTUAL length of multiple media files.
        if imgType == 'Episode':
            # Iterate through the additional media files
            for addFile in additionalFiles:
                # Add the file data to the file list
                filenameList.append({'filename' : addFile['filename'], 'offset' : addFile['offset'], 'length' : addFile['length']})
                # See if it has the minimum (largest negative) offset
                minVal = min(minVal, addFile['offset'])
            # This determines the VIRTUAL LENGTH of the multiple media streams
            mediaLength = self.ControlObject.GetMediaLength(True)
            self.zoomInfo = [(0, -1)]
        # If we have a Clip, we just need to set up the filename list
        elif imgType == 'Clip':
            # Iterate through the additional media files
            for addFile in additionalFiles:
                # Add the file data to the file list
                filenameList.append({'filename' : addFile['filename'], 'offset' : addFile['offset'] + offset, 'length' : addFile['length']})
            # To start with, initialize the data structure that holds information about Zooms
            self.zoomInfo = [(mediaStart, mediaLength)]
        # Let's clear the Visualization Window as we get started.
        self.ClearVisualization()
        # Remove old Waveform Selection and Cursor data
        self.waveform.ClearTransanaSelection()
        if self.kwMap != None:
            del(self.kwMap)
            self.kwMap = None
        # reset heightIsSet
        self.heightIsSet = False

        # If the Visualizatoin Object is defined ...
        if self.VisualizationObject != None:
            # ... load the Visualization zoom and filter information for know data objects
            self.LoadVisualizationInfo((type(self.VisualizationObject), self.VisualizationObject.number))
        # Turn off OnIdle Redraw during audio extraction.
        self.redrawWhenIdle = False
        # Initialize a list structure to hold wave file information
        self.waveFilename = []
        try:
            # Let's assume that audio extraction worked.  Signal success!
            dllvalue = 0                                
            # If the Waveforms Directory does not exist, create it.
            if not os.path.exists(TransanaGlobal.configData.visualizationPath):
                # (os.makedirs is a recursive call to create ALL needed folders!)
                os.makedirs(TransanaGlobal.configData.visualizationPath)
            # Initialize a result for the Waveform prompt
            result = wx.ID_NO
            # Let's do audio extraction of all the files first
            for filenameItem in filenameList:
                # Separate path and filename
                (path, filename) = os.path.split(filenameItem['filename'])
                # break filename into root filename and extension
                (filenameroot, extension) = os.path.splitext(filename)
                # Build the correct filename for the Wave File
                waveFilename = os.path.join(TransanaGlobal.configData.visualizationPath, filenameroot + '.wav')
                # Add information to the Waveform Filename list.  Offsets get adjusted for the largest negative value, so they are all 0 or higher!
                self.waveFilename.append({'filename' : waveFilename, 'offset' : filenameItem['offset'] + abs(minVal), 'length' : filenameItem['length']})
                # Create a Wave File if none exists!
                if not(os.path.exists(waveFilename)):
                    # The user only needs to say Yes once, but will be asked for each file if they say No.  See if they've already said Yes.
                    if result != wx.ID_YES:

                        # This totally sucks.
                        #
                        # If I have a Document loaded and go to load an Episode/Transcript that has not been
                        # thought audio extraction, the "dlg.LocalShowModal()" call below causes
                        # self.ControlObject.currentObj to become None.  I have *NO* idea why, other than a
                        # vague sense that there must be a memory leak somehow.  It's not caused by the act
                        # of audio extraction, as it still happens if you choose NOT to extract.  The problem
                        # also still occurs if you skip creation of the QuestionDialog entirely.  It is also
                        # cross-platform.  This is all most puzzling.
                        #
                        # The solution I have found is to create a copy of the object, and to restore
                        # self.ControlObject.currentObj from that copy following the destruction of the
                        # dialog.  I apologize to the programming Gods for this kludge.

                        tmpObj = self.ControlObject.currentObj
                        # Politely ask the user to create the waveform
                        dlg = Dialogs.QuestionDialog(self, _("No wave file exists.  Would you like to create one now?"), _("Transana Wave File Creation"))
                        # Remember the results.
                        result = dlg.LocalShowModal()
                        # Destroy the Dialog that asked to create the Wave file    
                        dlg.Destroy()
                        # If self.ControlObject.currentObj has been mysteriously wiped out ...
                        if self.ControlObject.currentObj == None:
                            # ... restore it.
                            self.ControlObject.currentObj = tmpObj
                    # If the user says Yes, we do audio extraction.
                    if result == wx.ID_YES:
                        try:

                            # NOTE:  We can't use multi-threaded audio extraction here, as it is non-modal,
                            #        and we need this to be modal.  That is, we want further processing
                            #        blocked here until the audio extraction is DONE.
                            
                            # Build the progress box's label
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_("Extracting %s\nfrom %s"), 'utf8')
                            else:
                                prompt = _("Extracting %s\nfrom %s")
                            # Create the Waveform Progress Dialog
                            self.progressDialog = WaveformProgress.WaveformProgress(self, prompt % (waveFilename, filenameItem['filename']))
                            # Tell the Waveform Progress Dialog to handle the audio extraction modally.
                            self.progressDialog.Extract(filenameItem['filename'], waveFilename)
                            # Get the Error Log that may have been created
                            errorLog = self.progressDialog.GetErrorMessages()
                            # Okay, we're done with the Progress Dialog here!
                            self.progressDialog.Destroy()
                            # If the user cancelled the audio extraction ...
                            if (len(errorLog) == 1) and (errorLog[0] == 'Cancelled'):
                                # ... signal that the WAV file was NOT created!
                                dllvalue = 1  

                            # On Windows only, some Unicode files fail the standard audio extraction process because of the unicode
                            # file names.  Let's try to detect that and if we do, let's re-run audio extraction using the OLD method!

                            # First, see if the waveform file is NOT created.
                            elif not os.path.exists(waveFilename):
                                # If not, re-call audio extraction with the old audio extraction method
                                # Create the Waveform Progress Dialog
                                self.progressDialog = WaveformProgress.WaveformProgress(self, prompt % (waveFilename, filenameItem['filename']))
                                # Tell the Waveform Progress Dialog to handle the audio extraction modally.
                                self.progressDialog.Extract(filenameItem['filename'], waveFilename, mode='AudioExtraction-OLD')
                                # Get the Error Log that may have been created
                                errorLog = self.progressDialog.GetErrorMessages()
                                # Okay, we're done with the Progress Dialog here!
                                self.progressDialog.Destroy()
                                # If the user cancelled the audio extraction ...
                                if (len(errorLog) == 1) and (errorLog[0] == 'Cancelled'):
                                    # ... signal that the WAV file was NOT created!
                                    dllvalue = 1
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

                            dllvalue = 1  # Signal that the WAV file was NOT created!                        

                            # Close the Progress Dialog when the DLL call is complete
                            self.progressDialog.Close()

                    else:
                        # User declined to create the WAV file now
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
            return

        # If the user said no or there was a problem with Wave Extraction ...
        if dllvalue != 0:
            # Remove whatever image might be there now.  First, clear the file names.
            self.waveFilename = []
            # Then clear the waveform itself
            self.waveform.Clear()
        else:
            # Separate path and filename
            (path, filename) = os.path.split(filename)
            # break filename into root filename and extension
            (filenameroot, extension) = os.path.splitext(filename)

        # If we're in Hybrid mode, clear the visualization to prevent waveform contamination!
        if self.VisualizationType == 'Hybrid':
            self.ClearVisualization()

        # Now that audio extraction is complete, signal that it's time to draw the Waveform Diagram during
        # Idle time.
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
        if self.VisualizationType == 'Text-Keyword':
            self.waveformUpperLimit = mediaLength
        else:
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

    def draw_timeline_text(self, startChar, documentLength):
        def GetScaleIncrements(documentLength):
            # The general rule is to try to get logical interval sizes with 8 or fewer time increments.
            # You always add a bit (20% at the lower levels) of the documentLength
            # because the final length is placed elsewhere and we don't want overlap.
            # This routine covers from 0 characters to 35,000,000 characters in length.

            # documentLength of < 15 characters = 2 character intervals
            if documentLength < 15: 
                Num = int((documentLength + 2) / 2)
                Interval = 2
            # documentLength of 35 characters or less = 5 character intervals 
            elif documentLength < 35:
                Num = int((documentLength + 4) / 5)
                Interval = 5
            # documentLength of 70 characters or less = 10 character intervals 
            elif documentLength < 70:
                Num = int((documentLength + 10) / 10)
                Interval = 10
            # documentLength of 105 characters or less = 15 character intervals 
            elif documentLength < 105:
                Num = int((documentLength + 10) / 15)
                Interval = 15
            # documentLength of 140 characters or less = 20 character intervals 
            elif documentLength < 140:
                Num = int((documentLength + 10) / 20)
                Interval = 20
            # documentLength of 280 characters or less = 40 character intervals 
            elif documentLength < 280:
                Num = int((documentLength + 10) / 40)
                Interval = 40
            # documentLength of 350 characters or less = 50 character intervals 
            elif documentLength < 350:
                Num = int((documentLength + 10) / 50)
                Interval = 50
            # documentLength of 700 characters or less = 100 character intervals 
            elif documentLength < 700:
                Num = int((documentLength + 100) / 100)
                Interval = 100
            # documentLength of 1400 characters or less = 200 character intervals 
            elif documentLength < 1400:
                Num = int((documentLength + 100) / 200)
                Interval = 200
            # documentLength of 3500 characters or less = 500 character intervals 
            elif documentLength < 3500:
                Num = int((documentLength + 100) / 500)
                Interval = 500
            # documentLength of 7000 characters or less = 1000 character intervals 
            elif documentLength < 7000:
                Num = int((documentLength + 1000) / 1000)
                Interval = 1000
            # documentLength of 14000 characters or less = 2000 character intervals 
            elif documentLength < 14000:
                Num = int((documentLength + 1000) / 2000)
                Interval = 2000
            # documentLength of 35000 characters or less = 5000 character intervals 
            elif documentLength < 35000:
                Num = int((documentLength + 1000) / 5000)
                Interval = 5000
            # documentLength of 70000 characters or less = 10000 character intervals 
            elif documentLength < 70000:
                Num = int((documentLength + 10000) / 10000)
                Interval = 10000
            # documentLength of 140000 characters or less = 20000 character intervals 
            elif documentLength < 140000:
                Num = int((documentLength + 10000) / 20000)
                Interval = 20000
            # documentLength of 350000 characters or less = 50000 character intervals 
            elif documentLength < 350000:
                Num = int((documentLength + 10000) / 50000)
                Interval = 50000
            # documentLength of 700000 characters or less = 100000 character intervals 
            elif documentLength < 700000:
                Num = int((documentLength + 100000) / 100000)
                Interval = 100000
            # documentLength of 1400000 characters or less = 200000 character intervals 
            elif documentLength < 1400000:
                Num = int((documentLength + 100000) / 200000)
                Interval = 200000
            # documentLength of 3500000 characters or less = 500000 character intervals 
            elif documentLength < 3500000:
                Num = int((documentLength + 100000) / 500000)
                Interval = 500000
            # documentLength of 7000000 characters or less = 1000000 character intervals 
            elif documentLength < 7000000:
                Num = int((documentLength + 1000000) / 1000000)
                Interval = 1000000
            # documentLength of 14000000 characters or less = 2000000 character intervals 
            elif documentLength < 14000000:
                Num = int((documentLength + 1000000) / 2000000)
                Interval = 2000000
            # documentLength of 35000000 characters or less = 5000000 character intervals 
            elif documentLength < 35000000:
                Num = int((documentLength + 1000000) / 5000000)
                Interval = 5000000
            else:
                Num = int((documentLength + 10000000) / 10000000)
                Interval = 10000000
            return Num, Interval

        # Positioning is different on Mac and Windows, requiring subtle adjustments in this window
        if '__WXMAC__' in wx.PlatformInfo:
            topAdjust = -3
        else:
            topAdjust = 0
        # Set the values for the waveform Lower and Upper limits
        self.waveformLowerLimit = startChar
        self.waveformUpperLimit = documentLength

        # Determine the number of labels and the time interval between labels that should be displayed
        numIncrements, Interval = GetScaleIncrements(documentLength - startChar)
        # Now we can determine the appropriate starting point for our labels!
        startingPoint = int((round(startChar / Interval) + 1) * Interval)

        # Clear all the existing labels
        self.timeline.DestroyChildren()
        timeLabels = []

        for loop in range(startingPoint, (numIncrements * Interval + startingPoint), Interval):
            if (loop > 0) and (documentLength > 0):
                # Place line marks
                lay = wx.LayoutConstraints()
                lay.top.SameAs(self.timeline, wx.Top, topAdjust)
                (width, height) = self.timeline.GetSizeTuple()
                width = width - 6  # Adjust for size of widget frame
                lay.left.Absolute(int(round(((float(loop - startChar)) / (documentLength - startChar)) * (width))))  
                lay.width.AsIs()
                lay.height.AsIs()
                wx.StaticLine(self.timeline, 0, size=wx.Size(2, 5)).SetConstraints(lay)
                
                # Place time labels
                lay = wx.LayoutConstraints()
                lay.top.SameAs(self.timeline, wx.Top, 4 + topAdjust)

                lay.centreX.Absolute(int(round(((float(loop - startChar)) / (documentLength - startChar)) * (width))))
                lay.width.AsIs()
                lay.height.AsIs()
                wx.StaticText(self.timeline, 1, "%d" % loop).SetConstraints(lay)

        self.timeline.SetAutoLayout(True)
        self.timeline.Layout()
        # Show the total Media Length in the Total Time label
        self.lbl_Total_Time.SetLabel("%d" % (documentLength - startChar))

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
        # If we have a Transcript currently selected, not a Document
        if self.ControlObject.GetCurrentItemType() == 'Transcript':
            # Show the media position in the Current Time label
            self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(currentPosition))
        else:
            self.lbl_Current_Time.SetLabel("0")
        # If UpperLimit and LowerLimit have not been set yet, avoid dividing by zero.  Things will be cleaned up later.
        if self.waveformUpperLimit == self.waveformLowerLimit:
            self.waveformUpperLimit = self.waveformUpperLimit + 1
        # Determine the current/new horizontal position within the waveform
        pos = ((float(currentPosition - self.waveformLowerLimit)) / (self.waveformUpperLimit - self.waveformLowerLimit))
        # Let's catch the problem of indeterminate media file lengths.  By the time this is called, the length should be known.
        if self.zoomInfo[0][1] == -1:
            self.zoomInfo[0] = (0, self.ControlObject.GetMediaLength(True))
        # Check to see if automatic waveform scrolling is necessary
        # First, check to see if we are zoomed in, if we're at the END of the waveform, if there's room to scroll, and if a redraw
        # hasn't already been requested.
        if (len(self.zoomInfo) > 1) and (pos > 0.99) and (self.waveformUpperLimit < self.zoomInfo[0][0] + self.zoomInfo[0][1]) and not self.redrawWhenIdle:
            # If there is a relatively large shift to take place (over 10% of the current width) ...
            if (pos > 1.1):
                # ... then let's just jump directly to the new position
                self.startPoint = currentPosition - int(self.zoomInfo[-1][1] * 0.25)
                self.endPoint = self.startPoint + self.zoomInfo[-1][1]
                if self.endPoint > self.zoomInfo[0][1]:
                    self.startPoint = self.zoomInfo[0][1] - self.zoomInfo[-1][1]
                    self.endPoint = self.zoomInfo[0][1]
                self.zoomInfo[-1] = (self.startPoint, self.zoomInfo[-1][1])
                # Signal that the waveform needs to be redrawn
                self.redrawWhenIdle = True
            # If we have a more modest shift to make ...
            else:
                # ... and call the Scroll event
                self.OnScroll('Right')
        # Second, check to see if we are zoomed in, if we're at the START of the waveform, if there's room to scroll, and if a redraw
        # hasn't already been requested.
        elif (len(self.zoomInfo) > 1) and (pos < 0) and (self.waveformLowerLimit > self.zoomInfo[0][0]) and not self.redrawWhenIdle:
            # If there is a relatively large shift to take place (over 10% of the current width) ...
            if (pos < -0.1):
                # ... then let's just jump directly to the new position
                self.startPoint = currentPosition - int(self.zoomInfo[-1][1] * 0.75)
                if self.startPoint < self.zoomInfo[0][0]:
                    self.startPoint = self.zoomInfo[0][0]
                self.endPoint = self.startPoint + self.zoomInfo[-1][1]
                self.zoomInfo[-1] = (self.startPoint, self.zoomInfo[-1][1])
                # Signal that the waveform needs to be redrawn
                self.redrawWhenIdle = True
            # If we have a more modest shift to make ...
            else:
                # ... and call the Scroll event
                self.OnScroll('Left')
        # If no scrolling is needed ...
        else:
            # ... just draw the new current position on the waveform
            self.waveform.DrawCursor(pos)

        # Force a redraw at least every half second while the video is playing.  (Mostly for slow Macs and when running multi-transcript video)
        if (self.ControlObject.IsPlaying()) and (time.time() - self.lastRedrawTime > 0.5):
            self.waveform.reInitBuffer = True
            # The video was getting very jumpy when the keyword or hybrid visualization is too complex with updates every
            # 0.2 seconds.  Removing the OnIdle line solves that, but then we don't get ANY forced visualization cursor updates.
            self.waveform.OnIdle(None)

            # While we're at it, let's see if we need to Yield().  This makes the app MUCH more responsive on slow systems. 
            # There is an issue with recursive calls to wxYield, so trap the exception ...
            try:
                wx.YieldIfNeeded()
            # ... and ignore it!
            except:
                pass
            # Let's keep track of time since last redraw
            self.lastRedrawTime = time.time()

    def SetDocumentSelection(self, startPos, endPos):
        """ Respond to the selection of text in a Document """
        # Set the Selection in the Visualization Window
        self.startPoint = startPos
        self.endPoint = endPos
        # Signal that he Visualization needs to be redrawn
        self.redrawWhenIdle = True
        # Update the Current Position
        self.lbl_Current_Time.SetLabel("%s" % endPos)
        # Update the Selection Text
        self.lbl_Selected_Time.SetLabel("%d - %d" % (startPos, endPos))

    def OnLeftDown(self, x, y, xpct, ypct):
        """ Mouse Left Down event for the Visualization Window -- over-rides the Waveform's left down! """
        # If we don't convert this to an int, our SQL gets screwed up in non-English localizations that use commas instead
        # of decimals.
        self.startPoint = int(round(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit))
        # If we have a Transcript currently selected, not a Document
        if self.ControlObject.GetCurrentItemType() == 'Transcript':
            # Show the media position in the Current Time label
            self.lbl_Current_Time.SetLabel(Misc.time_in_ms_to_str(self.startPoint))
        elif self.ControlObject.GetCurrentItemType() == 'Document':
            # Show the media position in the Current Time label
            self.lbl_Current_Time.SetLabel("%s" % self.startPoint)
            # The Document, but not the Transcript, needs to signal a Visualization Redraw here
            self.redrawWhenIdle = True      


    def OnLeftUp(self, x, y, xpct, ypct):
        """ Mouse Left Up event for the Visualization Window -- over-rides the Waveform's left up! """
        # If we have a Transcript currently selected, not a Document
        if self.ControlObject.GetCurrentItemType() == 'Transcript':
            # If the media is currently playing ...
            if self.ControlObject.IsPlaying():
                # ... we need to stop it!
                self.ControlObject.Stop()
        # Distinguish a left-click (positioning start only) from a left-drag (select range)
        # If we have a DRAG ...
        if self.startPoint != int(round(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit)):
            # ... set the end point value based on mouse position
            self.endPoint = int(round(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit))
            # If the user drags to the left rather than to the right, we need to swap the values!
            if self.endPoint < self.startPoint:
                temp = self.startPoint
                self.startPoint = self.endPoint
                self.endPoint = temp
            # If we have a Transcript currently selected, not a Document
            if self.ControlObject.GetCurrentItemType() == 'Transcript':
                # Show the media selection in the Selected Time label
                self.lbl_Selected_Time.SetLabel(Misc.time_in_ms_to_str(self.endPoint - self.startPoint))
            else:
                # Show the text selection in the Selected label
                self.lbl_Selected_Time.SetLabel("%d - %d" % (self.startPoint, self.endPoint))
                self.ControlObject.SetCurrentDocumentSelection(self.startPoint, self.endPoint)
                # Signal a redraw to eliminate the selection box from the Visualization
####                self.redrawWhenIdle = True
        # If we have a CLICK ...
        else:
            # if we have a LtR language ...
            if TransanaGlobal.configData.LayoutDirection == wx.Layout_LeftToRight:
                # ... define the self.looping variable based on the button state
                self.looping = self.loop.GetValue()
            # If we're Looping ...
            if self.looping:
                # ... then set the end point to 5 seconds after start the video length, whichever is smaller
                self.endPoint = self.startPoint + min(self.zoomInfo[0][1], 5000)
            # If we're NOT looping ...
            else:
                # ... then set endpoint to 0 to indicate we don't have a selection
                self.endPoint = 0
            # If we have a Transcript currently selected, not a Document
            if self.ControlObject.GetCurrentItemType() == 'Transcript':
                # Clear the media selection in the Selected Time label
                self.lbl_Selected_Time.SetLabel(Misc.time_in_ms_to_str(0))
            else:
                # Show the LACK of a text selection in the Selected label
                self.lbl_Selected_Time.SetLabel("%d" % 0)
                self.ControlObject.SetCurrentDocumentPosition(self.startPoint, (-2, -2))
                # Determine the current/new horizontal position within the visualization
                if (self.waveformUpperLimit - self.waveformLowerLimit) != 0:
                    pos = ((float(self.startPoint - self.waveformLowerLimit)) / (self.waveformUpperLimit - self.waveformLowerLimit))
                else:
                    pos = 0
                self.waveform.DrawCursor(pos)

        # If we have a Transcript currently selected, not a Document
        if self.ControlObject.GetCurrentItemType() == 'Transcript':
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
        # If there is something loaded in the main interface and the CURRENT item is NOT a DOCUMENT
        if (self.ControlObject.currentObj != None) and \
           (self.ControlObject.GetCurrentItemType() == 'Transcript'):
            # ... Play / Pause the video
            self.ControlObject.PlayPause()

    def OnMouseOver(self, x, y, xpct, ypct):

#        print "VisualizationWindow.OnMouseOver():", x, xpct
        
        if self.ControlObject.GetCurrentItemType() == 'Transcript':
            self.lbl_Time_Time.SetLabel(Misc.time_in_ms_to_str(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit))
        else:
            self.lbl_Time_Time.SetLabel("%d" % round(xpct * (self.waveformUpperLimit - self.waveformLowerLimit) + self.waveformLowerLimit))

    def GetDimensions(self):
        (left, top) = self.GetPositionTuple()
        (width, height) = self.GetSizeTuple()
        
        # Mac was having problems with the default window position being at -20.
        # Don't accept anything less than 0 for the left parameter.
        (adjustX, adjustY, adjustW, adjustH) = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
        if left < adjustX:
            left = adjustX
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
        """ Change the dimensions of the Visualization Window """
        # Change the Visualization Frame's dimensions
        self.SetDimensions(left, top, width, height)
        # If we've changed the size of the visualization window, we need to redraw the visualization
        self.redrawWhenIdle = True

    def ChangeLanguages(self):
        self.SetTitle(_('Visualization'))
        self.filter.SetToolTipString(_("Filter"))
        self.zoomIn.SetToolTipString(_("Zoom In"))
        self.zoomOut.SetToolTipString(_("Zoom Out"))
        self.zoom100.SetToolTipString(_("Zoom to 100%"))
        self.createClip.SetToolTipString(_("Create Transcript-less Clip"))
        self.loop.SetToolTipString(_("Loop Playback"))
        self.btn_Current.SetToolTipString(_("Insert Time Code"))
        self.btn_Selected.SetToolTipString(_("Insert Time Span"))
        self.lbl_Time.SetLabel(_("Time:"))
        self.btn_Current.SetLabel(_("Current:"))
        self.btn_Selected.SetLabel(_("Selected:"))
        self.lbl_Total.SetLabel(_("Total:"))

    def GetNewRect(self):
        """ Get (X, Y, W, H) for initial positioning """
        pos = self.__pos()
        size = self.__size()
        return (pos[0], pos[1], size[0], size[1])
        
    def OnSize(self, event):
        """ Handles widget positioning on Resize Event """
        # Determine Frame Size
        (width, height) = self.GetSize()
        # If we're not resizing ALL Transana windows ...   (to avoid recursive OnSize calls)
        if not TransanaGlobal.resizingAll:
            # ... Get the Visualization Window position ...
            (left, top) = self.GetPositionTuple()
            # ... can call the ControlObject's Update Window Position routine with the appropiate parameters
            self.ControlObject.UpdateWindowPositions('Visualization', width + left, YUpper = height + top)

        # Determine NEW Frame Size
        (width, height) = self.GetSize()

        # We need to know the height of the Window Header to adjust the size of the Graphic Area
        headerHeight = self.GetSizeTuple()[1] - self.GetClientSizeTuple()[1]

        # Set size of Waveform Window (Do this first, or the timeline is not scaled properly.)
        self.waveform.SetDim(0, 1, width, height-50-headerHeight)

        # Position toolbar Panel
        self.toolbar.SetDimensions(0, height-28-headerHeight, width, self.toolbar.GetMinSize()[1])

        # Draw the appropriate time line
        self.draw_timeline(self.waveformLowerLimit, self.waveformUpperLimit - self.waveformLowerLimit)

        # Position timeline Panel
        self.timeline.SetDimensions(0, height-50-headerHeight, width, 24)
        # Tell the Waveform to redraw
        self.redrawWhenIdle = True

# Private methods    

    def __size(self):
        # Determine the correct monitor and get its size and position
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            primaryScreen = TransanaGlobal.configData.primaryScreen
        else:
            primaryScreen = 0
        rect = wx.Display(primaryScreen).GetClientArea()

        if not 'wxGTK' in wx.PlatformInfo:
            container = rect[2:4]
        elif 'wxGTK' in wx.PlatformInfo:
            screenDims = wx.Display(primaryScreen).GetClientArea()
            # screenDims2 = wx.Display(primaryScreen).GetGeometry()
            left = screenDims[0]
            top = screenDims[1]
            width = screenDims[2] - screenDims[0]  # min(screenDims[2], 1280 - self.left)
            height = screenDims[3]
            container = (width, height)

        width = container[0] * .71  # rect[2] * .715
        height = (container[1] - TransanaGlobal.menuHeight) * .24  # (rect[3] - TransanaGlobal.menuHeight) * .25

        if DEBUG:
            print "Visualization width:", container[0], container[0] * 0.715, width
            print "VisualizationWindow.__size():", width, height
            
        return wx.Size(int(width), int(height))

    def __pos(self):
        # Determine the correct monitor and get its size and position
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
        # If the Start Menu is on the left side, 0 is incorrect!  Get starting position from wx.Display.
        x = rect[0] + 1 # 1
        # rect[1] compensated if the Start menu is at the top of the screen
#        y = rect[1] + TransanaGlobal.menuHeight + 3

        if 'wxMac' in wx.PlatformInfo:
            y = rect[1] + 2
        else:
            y = rect[1] + TransanaGlobal.menuHeight + 1

        if DEBUG:
            print "VisualizationWindow.__pos():", x, y
            
        return wx.Point(int(x), int(y))

if __name__ == '__main__':
    app = wx.PySimpleApp()
    frame = VisualizationWindow(None)
    frame.ShowModal()
    frame.Destroy()
    app.MainLoop()
