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

"""This module contains functions for manipulating and coding still images in Transana."""

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

DEBUG=False
if DEBUG:
    print "SnapshotWindow DEBUG is ON!!"

# import wxPython
import wx
# import the FloatCanvas, FloatCanvas Resources, and FloatCanvas GUIMode module
from wx.lib.floatcanvas import FloatCanvas, Resources, GUIMode
# Import the Python os and sys modules
import os, sys
# Import Transana's Collection Object
import Collection
# import Transana's Dialogs
import Dialogs
# import Transana's Database Interface
import DBInterface
# Import Transana's Keyword Object
import KeywordObject
# import Transana's Keyword List Edit Form
import KeywordListEditForm
# import the Transana Snapshot Object
import Snapshot
# import Transana's constants
import TransanaConstants
# Import Transana Exceptions
import TransanaExceptions
# Import Transana Globals (for color objects)
import TransanaGlobal
# Import Transana images
import TransanaImages

# Define Menu constants
MENU_FILE_CLEAR                =  wx.NewId()
MENU_FILE_RESTORE              =  wx.NewId()
MENU_FILE_SHOWALLCODING        =  wx.NewId()
MENU_FILE_HIDEALLCODING        =  wx.NewId()
MENU_FILE_INSERT_IN_TRANSCRIPT =  wx.NewId()
MENU_FILE_SAVE_SELECTION_AS    =  wx.NewId()
MENU_FILE_SAVE_WHOLE_IMAGE_AS  =  wx.NewId()
MENU_FILE_CLOSE_ALL            =  wx.NewId()
MENU_FILE_EXIT                 =  wx.NewId()
MENU_POPUP_HIDE                =  wx.NewId()
MENU_POPUP_SENDTOBACK          =  wx.NewId()
MENU_POPUP_DELETE              =  wx.NewId()

class SnapshotWindow(wx.Frame):
    """ This window displays still images and allows coding of those images. """
    def __init__(self, parent, id, title, snapshot, showWindow=True):
        # Because of problems with the wx.EVT_ACTIVATE handler, we need to track if we're in the process of closing
        # this window.  Initially, we're not.
        self.closing = False
        # Retain the parent object (a MenuWindow?)
        self.parent = parent
        # The Snapshot Window can get the Control Object from it's parent, which it needs.
        if parent != None:
            self.ControlObject = self.parent.ControlObject
        else:
            self.ControlObject = None
        # Retain the snapshot object
        self.obj = snapshot
        # Retain showWindow so we don't remove HIDDEN snapshots from the Window Menu!
        self.showWindow = showWindow
        # Create a holder for the Coding Key popup
        self.codingKeyPopup = None
        # Initialize the Bitmap to None
        self.theBitmap = None

        # Check to see if the image file can be found
        if not os.path.exists(self.obj.image_filename):
            # If not, raise an exception
            errmsg = unicode(_("Image file not found:\n%s"), 'utf8')
            raise TransanaExceptions.ImageLoadError(errmsg % self.obj.image_filename)

        # Load the image
        self.bgImage = wx.Image(self.obj.image_filename)
            
        # Make sure the image is loaded, is not corrupt
        if not self.bgImage.IsOk():
            # If not, raise an exception
            errmsg = unicode(_("Unable to load image file:\n%s\nThere may be a problem with the file, or you may\nhave too many Snapshots open."), 'utf8')
            raise TransanaExceptions.ImageLoadError(errmsg % self.obj.image_filename)

        # If the image that is passed in has a defined window size ...
        if self.obj.image_size[0] > 0:
            # ... use that defined size
            width = self.obj.image_size[0]
            height = self.obj.image_size[1]
        # If the image does NOT have a size ...
        else:
            # ... use the default image window size
            width = self.__size()[0]
            height = self.__size()[1]
        # Initialize the Window Frame
        wx.Frame.__init__(self,parent,-1, title, pos = self.__pos(), size = (width, height), style=wx.DEFAULT_FRAME_STYLE|wx.NO_FULL_REPAINT_ON_RESIZE)
        # Set the background to WHITE
        self.SetBackgroundColour(wx.SystemSettings.GetColour(getattr(wx, 'SYS_COLOUR_MENUBAR')))

        # Let's go ahead and keep the menu for non-Mac platforms
        if self.showWindow and (not '__WXMAC__' in wx.PlatformInfo):
            # Menu Bar
            # Create a MenuBar
            menuBar = wx.MenuBar()
            # Build a Menu Object to go into the Menu Bar
            menuFile = wx.Menu()
            # Add the menu items that should appear in the File Menu
            menuFile.Append(MENU_FILE_CLEAR, _("&Remove All Coding"), _("Remove all coding"))
            menuFile.Append(MENU_FILE_RESTORE, _("Restore &Last Save"), _("Restore Last Save"))
            menuFile.Append(MENU_FILE_SHOWALLCODING, _("&Show All Coding"), _("Show all coding"))
            menuFile.Append(MENU_FILE_HIDEALLCODING, _("&Hide All Coding"), _("Hide all coding"))
            menuFile.Append(MENU_FILE_INSERT_IN_TRANSCRIPT, _("&Insert Image in Transcript"), _("Insert the Coded Image into the Transcript"))
            menuFile.AppendSeparator()
            menuFile.Append(MENU_FILE_SAVE_SELECTION_AS, _("Save &Visible Selection As"), _("Save Visible Selection As"))
            menuFile.Append(MENU_FILE_SAVE_WHOLE_IMAGE_AS, _("Save &Whole Image As"), _("Save Whole Image As"))
            menuFile.AppendSeparator()
            menuFile.Append(MENU_FILE_CLOSE_ALL, _("Close &All Snapshots"), _("Close All Snapshots"))
            menuFile.Append(MENU_FILE_EXIT, _("&Close"), _("Close this window"))
            #Place the Menu Item in the Menu Bar
            menuBar.Append(menuFile, _("&File"))

            # Place a Window menu in the Menu Bar
            # First, create the Window menu
            self.menuWindow = wx.Menu()
            # Add this to the menuBar
            menuBar.Append(self.menuWindow, _("Window"))

            # Place the Menu Bar on the Frame
            self.SetMenuBar(menuBar)
            #Define Events for the Menu Items
            wx.EVT_MENU(self, MENU_FILE_CLEAR, self.FileClear)
            wx.EVT_MENU(self, MENU_FILE_RESTORE, self.FileRestore)
            wx.EVT_MENU(self, MENU_FILE_SHOWALLCODING, self.FileRedraw)
            wx.EVT_MENU(self, MENU_FILE_HIDEALLCODING, self.FileRedraw)
            wx.EVT_MENU(self, MENU_FILE_INSERT_IN_TRANSCRIPT, self.OnInsertIntoTranscript)
            wx.EVT_MENU(self, MENU_FILE_SAVE_SELECTION_AS, self.FileSaveSelectionAs)
            wx.EVT_MENU(self, MENU_FILE_SAVE_WHOLE_IMAGE_AS, self.FileSaveAs)
            wx.EVT_MENU(self, MENU_FILE_CLOSE_ALL, self.CloseAllImages)
            wx.EVT_MENU(self, MENU_FILE_EXIT, self.CloseWindow)

        # Bind the Close Event
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        # Bind the Form Activate event
        self.Bind(wx.EVT_ACTIVATE, self.OnEnterWindow)

        # Define the Frame's Main Sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a Panel for the Frame
        self.panel = wx.Panel(self, -1)
        self.panel.SetBackgroundColour(wx.SystemSettings.GetColour(getattr(wx, 'SYS_COLOUR_MENUBAR')))
        
        # Put the Panel on the Sizer
        mainSizer.Add(self.panel, 9, wx.GROW, 4)
        # Create a Sizer for the Panel
        pnlSizer = wx.BoxSizer(wx.VERTICAL)

        # Get a list of all Snapshots in the same collection
        self.snapshotList = DBInterface.list_of_snapshots_by_collectionnum(self.obj.collection_num, True)
        # Initialize values for Previous and Next Snapshots
        self.prevSnapshot = 0
        self.nextSnapshot = 0
        # Determine the list index for the current Snapshot
        index = self.snapshotList.index((self.obj.number, self.obj.id, self.obj.collection_num, self.obj.sort_order))
        # If the current snapshot isn't the first item in the list ...
        if index > 0:
            # ... then remember the previous snapshot's number
            self.prevSnapshot = self.snapshotList[index - 1][0]
        # If the current snapshot isn't the last item in the list ...
        if index < len(self.snapshotList) - 1:
            # ... then remember the next snapshot's number
            self.nextSnapshot = self.snapshotList[index + 1][0]

        if self.showWindow:
            # Create the Toolbar
            self.toolbar = self.CreateToolBar ()  # wx.ToolBar(self.panel)
            # Set the Bitmap Size for the Toolbar
            self.toolbar.SetToolBitmapSize((16, 16))

            # Create the Edit-Mode Tool
            bmp = TransanaImages.ReadOnly16.GetBitmap()
            self.editTool = self.toolbar.AddCheckTool(wx.ID_ANY, bitmap=bmp, shortHelp = _("Edit/Read-only"))
            self.Bind(wx.EVT_TOOL, self.OnToolbar, self.editTool)

            # Create the Pointer Tool
            bmp = Resources.getPointerBitmap()
            self.pointer = self.toolbar.AddRadioTool(wx.ID_ANY, bitmap=bmp, shortHelp=_("Coding Tool"))
            self.Bind(wx.EVT_TOOL, self.OnToolbar, self.pointer)

            # Create the Move Tool
            bmp = Resources.getHandBitmap()
            self.moveTool = self.toolbar.AddRadioTool(wx.ID_ANY, bitmap=bmp, shortHelp=_("Move"))
            self.Bind(wx.EVT_TOOL, self.OnToolbar, self.moveTool)

            # Create the Zoom In Tool
            bmp = Resources.getMagPlusBitmap()
            self.zoomIn = self.toolbar.AddRadioTool(wx.ID_ANY, bitmap=bmp, shortHelp=_("Zoom In"))
            self.Bind(wx.EVT_TOOL, self.OnToolbar, self.zoomIn)

            # Create the Zoom Out Tool
            bmp = Resources.getMagMinusBitmap()
            self.zoomOut = self.toolbar.AddRadioTool(wx.ID_ANY, bitmap=bmp, shortHelp=_("Zoom Out"))
            self.Bind(wx.EVT_TOOL, self.OnToolbar, self.zoomOut)

            # Add Edit keywords button
            bmp = TransanaImages.KeywordRoot16.GetBitmap()
            self.keywordTool = self.toolbar.AddTool(wx.ID_ANY, bitmap=bmp, isToggle=False, shortHelpString = _("Whole Snapshot Keywords"))
            self.Bind(wx.EVT_TOOL, self.OnEditKeywords, self.keywordTool)

            # Add Coding Key button
            bmp = TransanaImages.Keyword16.GetBitmap()
            self.codingKeyTool = self.toolbar.AddTool(wx.ID_ANY, bitmap=bmp, isToggle=False, shortHelpString = _("Show Coding Key"))
            self.Bind(wx.EVT_TOOL, self.OnCodingKey, self.codingKeyTool)

            # Add a Previous button
            self.prevSnapBtn = self.toolbar.AddTool(wx.ID_ANY, bitmap=TransanaImages.ArtProv_BACK.GetBitmap(), isToggle=False, shortHelpString = _("Previous Snapshot"))
            self.Bind(wx.EVT_TOOL, self.OnChangeSnapshot, self.prevSnapBtn)
            # If there is no previous snapshot ...
            if self.prevSnapshot == 0:
                # ... disable the button
                self.toolbar.EnableTool(self.prevSnapBtn.GetId(), False)
                
            # If there is a next snapshot in the collection, add a Next button
            self.nextSnapBtn = self.toolbar.AddTool(wx.ID_ANY, bitmap=TransanaImages.ArtProv_FORWARD.GetBitmap(), isToggle=False, shortHelpString = _("Next Snapshot"))
            self.Bind(wx.EVT_TOOL, self.OnChangeSnapshot, self.nextSnapBtn)
            # If there is no next snapshot ...
            if self.nextSnapshot == 0:
                # ... disable the button
                self.toolbar.EnableTool(self.nextSnapBtn.GetId(), False)

            # Use multiple Separators to create a space between cursor modes and coding tools
            self.toolbar.AddSeparator()
            # Get the Shapshot's Collection record
            tmpCollection = Collection.Collection(self.obj.collection_num)
            # Create the Keyword Group Selector
            txt = wx.StaticText(self.toolbar, wx.ID_ANY, " " + _("Keyword Group:") + " ")
            self.toolbar.AddControl(txt)
            # Get the Keyword Groups to populate the control
            choices = [''] + DBInterface.list_of_keyword_groups()
            self.keyword_group_cb = wx.Choice(self.toolbar, wx.ID_ANY, choices=choices)
            # If there is a Default Keyword Group ...
            if (tmpCollection.keyword_group != '') and (tmpCollection.keyword_group in choices):
                # ... make that the initial selection
                self.keyword_group_cb.SetStringSelection(tmpCollection.keyword_group)
            # If there's no default keyword group ...
            elif len(choices) > 0:
                # ... select the blank element
                self.keyword_group_cb.Select(0)
            self.toolbar.AddControl(self.keyword_group_cb)
            self.keyword_group_cb.Bind(wx.EVT_CHOICE, self.OnKWGSelect)

            # Create the Keyword Selector
            txt = wx.StaticText(self.toolbar, wx.ID_ANY, "   " + _("Keyword:") + " ")
            self.toolbar.AddControl(txt)
            # If we have a Keyword Group ...
            if self.keyword_group_cb.GetStringSelection() != '':
                # ... get that group's Keywords
                choices = [''] + DBInterface.list_of_keywords_by_group(self.keyword_group_cb.GetStringSelection())
            # If there's no Keyword Group ...
            else:
                # ... there are no Keywords to add yet.
                choices = ['This keyword is used for sizing']
            self.keyword_cb = wx.Choice(self.toolbar, wx.ID_ANY, choices=choices)
            # In any case, select the Blank option to start
            self.keyword_cb.Select(0)
            self.toolbar.AddControl(self.keyword_cb)
            self.keyword_cb.Bind(wx.EVT_CHOICE, self.OnKWSelect)

            self.toolbar.AddSeparator()

            # Create the Hide / Show Keyword button
            self.kwShowHideButton = wx.Button(self.toolbar, label=_("Hide"))
            self.toolbar.AddControl(self.kwShowHideButton)
            self.kwShowHideButton.Bind(wx.EVT_BUTTON, self.OnKWShowHideTool)
            self.kwShowHideButton.Enable(False)

            # Realize the Toolbar
            self.toolbar.Realize()
            
            # Create the second Toolbar
            self.toolbar2 = wx.ToolBar(self.panel)
            # Set the Bitmap Size for the Toolbar
            self.toolbar2.SetToolBitmapSize((16, 16))

            # Create the Zoom to 100% Tool
            self.zoomToFull = wx.Button(self.toolbar2, label=_("100%"))
            self.toolbar2.AddControl(self.zoomToFull)
            self.zoomToFull.Bind(wx.EVT_BUTTON, self.OnToolbar)

            # Create the Zoom To Fit Tool
            self.zoomToFit = wx.Button(self.toolbar2, label=_("Fit"))
            self.toolbar2.AddControl(self.zoomToFit)
            self.zoomToFit.Bind(wx.EVT_BUTTON, self.OnToolbar)

            # Create the Insert Coded Image Into Transcript button
            # Get the initial image for the Play / Pause button
            self.pushToTranscript = wx.BitmapButton(self.toolbar2, -1, TransanaImages.Snapshot.GetBitmap(), size=(48, 24))
            self.pushToTranscript.SetToolTipString(_("Insert Coded Image into Transcript"))
            self.toolbar2.AddControl(self.pushToTranscript)
            self.pushToTranscript.Bind(wx.EVT_BUTTON, self.OnInsertIntoTranscript)

            # Use multiple Separators to create a space between cursor modes and coding tools
            self.toolbar2.AddSeparator()
            # Create the Color Selector
            txt = wx.StaticText(self.toolbar2, wx.ID_ANY, " " + _("Color:") + " ")
            self.toolbar2.AddControl(txt)
            # Get the list of colors for populating the control
            choices = []
            # Make a dictionary for looking up the color definintions that match the color names
            self.colorList = {}
            # Iterate through the global Graphics Colors (which can be user-defined!)
            for x in TransanaGlobal.transana_graphicsColorList:
                # We need to exclude WHITE
                if x[1] != (255, 255, 255):
                    # Get the TRANSLATED color name
                    tmpColorName = _(x[0])
                    # If the color name is a string ...
                    if isinstance(tmpColorName, str):
                        # ... convert it to unicode
                        tmpColorName = unicode(tmpColorName, 'utf8')
                    # Add the translated color name to the choice box
                    choices.append(tmpColorName)
                    # Add the color definition to the dictionary, using the translated name as the key
                    self.colorList[tmpColorName] = x
            self.line_color_cb = wx.Choice(self.toolbar2, wx.ID_ANY, choices=choices)
            # Select the first color in the choice box
            self.line_color_cb.SetStringSelection(_(TransanaGlobal.keywordMapColourSet[0]))
            self.toolbar2.AddControl(self.line_color_cb)
            self.line_color_cb.Bind(wx.EVT_CHOICE, self.OnToolbar)
            # Disable Color Selection
            self.line_color_cb.Enable(False)

            # Create the Shape Selection Tool
            txt = wx.StaticText(self.toolbar2, wx.ID_ANY, "  " + _("Code Tool:") + " ")
            self.toolbar2.AddControl(txt)
            self.codeShape = wx.Choice(self.toolbar2, wx.ID_ANY, choices=[_('Rectangle'), _('Ellipse'), _('Line'), _('Arrow')])
            # Set the Shape to Rectangle by default
            self.codeShape.SetStringSelection(_('Rectangle'))
            self.toolbar2.AddControl(self.codeShape)
            self.codeShape.Bind(wx.EVT_CHOICE, self.OnToolbar)
            # Disable Shape Selection
            self.codeShape.Enable(False)

            # Create the Line Width Tool
            txt = wx.StaticText(self.toolbar2, wx.ID_ANY, "  " + _("Line Width:") + " ")
            self.toolbar2.AddControl(txt)
            self.lineSize = wx.Choice(self.toolbar2, wx.ID_ANY, choices=['1', '2', '3', '4', '5', '6'])
            # Set Line Width to 3 by default
            self.lineSize.SetStringSelection('3')
            self.toolbar2.AddControl(self.lineSize)
            self.lineSize.Bind(wx.EVT_CHOICE, self.OnToolbar)
            # Disable Line Width selection
            self.lineSize.Enable(False)

            # Create the Line Style Tool
            txt = wx.StaticText(self.toolbar2, wx.ID_ANY, "  " + _("Line Style:") + " ")
            self.toolbar2.AddControl(txt)
            # ShortDash is indistinguishable from LongDash, at least on Windows, so I've left it out of the options here.
            self.line_style_cb = wx.Choice(self.toolbar2, wx.ID_ANY, choices=[_('Solid'), _('Dot'), _('Dash'), _('Dot Dash')])
            # Set line style to Solid by default
            self.line_style_cb.SetStringSelection(_('Solid'))
            self.toolbar2.AddControl(self.line_style_cb)
            self.line_style_cb.Bind(wx.EVT_CHOICE, self.OnToolbar)
            # Disable Line Style selection
            self.line_style_cb.Enable(False)

            self.toolbar2.AddSeparator()

            # Add a Help button
            self.help = wx.BitmapButton(self.toolbar2, -1, TransanaImages.ArtProv_HELP.GetBitmap(), size=(24, 24))
            self.help.SetToolTipString(_("Help"))
            self.toolbar2.AddControl(self.help)
            self.help.Bind(wx.EVT_BUTTON, self.OnHelp)

            # Realize the second Toolbar
            self.toolbar2.Realize()

            # If we do NOT have a Keyword Group ...
            if self.keyword_group_cb.GetStringSelection() == '':
                # ... clear out the items list, which was just used to size the control
                self.keyword_cb.SetItems([''])

            # Add the Toolbar to the Panel Sizer
            pnlSizer.Add(self.toolbar2, 0, wx.ALL | wx.ALIGN_LEFT | wx.GROW, 0)

        # If we're on a Mac, we need another Toolbar to handle the menu functions!
        if '__WXMAC__' in wx.PlatformInfo:

            # Create the second Toolbar
            self.toolbar3 = wx.ToolBar(self.panel)
            # Set the Bitmap Size for the Toolbar
            self.toolbar3.SetToolBitmapSize((16, 16))

            # Create the Remove All Keywords button
            self.removeAllCodingButton = wx.Button(self.toolbar3, id=MENU_FILE_CLEAR, label=_("Remove All Coding"))
            self.removeAllCodingButton.Bind(wx.EVT_BUTTON, self.FileClear)
            self.toolbar3.AddControl(self.removeAllCodingButton)

            # Create the Show All Keywords button
            self.showAllCodingButton = wx.Button(self.toolbar3, label=_("Show All Coding"))
            self.showAllCodingButton.Bind(wx.EVT_BUTTON, self.FileRedraw)
            self.toolbar3.AddControl(self.showAllCodingButton)

            # Create the Hide All Keywords button
            self.hideAllCodingButton = wx.Button(self.toolbar3, label=_("Hide All Coding"))
            self.hideAllCodingButton.Bind(wx.EVT_BUTTON, self.FileRedraw)
            self.toolbar3.AddControl(self.hideAllCodingButton)

            # Create the Save Visible Selection button
            self.saveVisibleButton = wx.Button(self.toolbar3, label=_("Save Visible"))
            self.saveVisibleButton.Bind(wx.EVT_BUTTON, self.FileSaveSelectionAs)
            self.toolbar3.AddControl(self.saveVisibleButton)

            # Create the Save Whole Image button
            self.saveWholeImageButton = wx.Button(self.toolbar3, label=_("Save Whole Image"))
            self.saveWholeImageButton.Bind(wx.EVT_BUTTON, self.FileSaveAs)
            self.toolbar3.AddControl(self.saveWholeImageButton)

            # German, Spanish, and French require that the supplementary Mac toolbar get split here.
            if TransanaGlobal.configData.language in ['de', 'es', 'fr']:
                # Realize the second Toolbar
                self.toolbar3.Realize()
                # Add the Toolbar to the Panel Sizer
                pnlSizer.Add(self.toolbar3, 0, wx.ALL | wx.ALIGN_LEFT | wx.GROW, 0)
                # Create the third Toolbar
                self.toolbar4 = wx.ToolBar(self.panel)
                # Set the Bitmap Size for the Toolbar
                self.toolbar4.SetToolBitmapSize((16, 16))
                tmpToolbar = self.toolbar4
            else:
                tmpToolbar = self.toolbar3

            # Create the Restore Last Save button
            self.restoreButton = wx.Button(tmpToolbar, label=_("Restore Last Save"))
            self.restoreButton.Bind(wx.EVT_BUTTON, self.FileRestore)
            tmpToolbar.AddControl(self.restoreButton)

            # Italian, Dutch, and Swedish require that the supplementary Mac toolbar get split here.
            if TransanaGlobal.configData.language in ['it', 'nl', 'sv']:
                # Realize the second Toolbar
                self.toolbar3.Realize()
                # Add the Toolbar to the Panel Sizer
                pnlSizer.Add(self.toolbar3, 0, wx.ALL | wx.ALIGN_LEFT | wx.GROW, 0)
                # Create the third Toolbar
                self.toolbar4 = wx.ToolBar(self.panel)
                # Set the Bitmap Size for the Toolbar
                self.toolbar4.SetToolBitmapSize((16, 16))
                tmpToolbar = self.toolbar4
                
            # Create the Close All button
            self.closeAllButton = wx.Button(tmpToolbar, label=_("Close All Snapshots"))
            self.closeAllButton.Bind(wx.EVT_BUTTON, self.CloseAllImages)
            tmpToolbar.AddControl(self.closeAllButton)

            # Create the Close button

            self.closeButton = wx.BitmapButton(tmpToolbar, -1, TransanaImages.Exit.GetBitmap(), size=(24, 24))
            self.closeButton.SetToolTipString(_("Close"))
            self.closeButton.Bind(wx.EVT_BUTTON, self.CloseWindow)
            tmpToolbar.AddControl(self.closeButton)

            # Realize the final Toolbar
            tmpToolbar.Realize()

            # Add the Toolbar to the Panel Sizer
            pnlSizer.Add(tmpToolbar, 1, wx.EXPAND | wx.ALL | wx.ALIGN_LEFT | wx.GROW, 0)

        # Create a FloatCanvas for the image and coding
        self.canvas = FloatCanvas.FloatCanvas(self.panel)
        # Set Layout Direction to Left-to-Right to prevent image reversal in Arabic
        self.canvas.SetLayoutDirection(wx.Layout_LeftToRight)
        # Add the FloatCanvas to the Panel Sizer
        pnlSizer.Add(self.canvas, 1, wx.EXPAND | wx.GROW, 0)
        # Set the Panel Sizer on the Panel and Fit it.
        self.panel.SetSizerAndFit(pnlSizer)

        # The FloatCanvas ScaledBitmap object seems to run into problems when zoomed in too far.  Large images raise
        # an exception on Zoom In with click and Wheel zooms, as well as with selection-based zoomed.  I contacted Chris
        # Barker, who wrote FloatCanvas, and he suggested I try ScaledBitmap2, which requires a 2-step creation process.
        # It seems to work.
        # Add a Scaled Bitmap, converted from the loaded image, to the SnapshotWindow canvas
        bgBitmapObj = FloatCanvas.ScaledBitmap2(self.bgImage,  # wx.BitmapFromImage(self.bgImage),
                                               (0 - (float(self.bgImage.GetWidth()) / 2.0), (float(self.bgImage.GetHeight()) / 2.0)),
                                               Height = self.bgImage.GetHeight(),
                                               Position = "tl")
        self.canvas.AddObject(bgBitmapObj)

        # Set the minimum scale to 1/20th normal size
        self.canvas.MinScale = 0.05
        # Set the maximum scale to 5 times normal size.  (This avoids unsightly but not serious errors with large images.)
        self.canvas.MaxScale = 5.0  # 3.70
        if self.showWindow:
            # Select the Move Tool initially
            self.toolbar.ToggleTool(self.moveTool.GetId(), True)
        # Set the Canvas Mode to the Move Tool initially
        self.canvas.SetMode(GUIMode.GUIMove())

        # Initialize the Coding Object Number to the number of existing objects
        self.objectNum = len(self.obj.codingObjects)
        # Initialize the variable that tracks the MouseDown position to None
        self.mouseDown = None
        # Initialize the variable that tracks the MouseUp position to None
        self.mouseUp = None
        # Initialize the Code Shape to the Rectanble
        self.drawMode = 'Rectangle'  # one of ['Arrow', 'Rectangle', 'Ellipse', 'Line']
        # Initialize the Line Width to 3
        self.lineWidth = 3
        # Initialize the Line Style to Solid
        self.lineStyle = 'Solid'
        # if the Snapshot Window is visible ...
        if self.showWindow:
            # Get the initial color
            color = self.colorList[self.line_color_cb.GetStringSelection()][1]
        # If the window is not visible ...
        else:
            # ... color and self.drawColor are ignored, so we can just set it to black
            color = wx.BLACK
        # Initialize the Line Color to the initial color
        self.drawColor = wx.Colour(color[0], color[1], color[2])

        # Note that events are NOT bound initially
        self.eventsAreBound = False
        # Bind the FloatCanvas Events
        self.BindEvents()

        # Create the Status Bar (to show Coding information)
        self.CreateStatusBar()

        # Set the Frame's Main Sizer
        self.SetSizer(mainSizer)
        # Make the Frame's Sizing automatic
        self.SetAutoLayout(True)
        # Lay out the frame
        self.Layout()

        # If the snapshot that was passed in has a defined Scale ...
        if self.obj.image_scale > 0.0:
            # ... set the canvas' scale to the snapshot's value ...
            self.canvas.Scale = self.obj.image_scale
            # ... and apply the scale change to the canvas.
            self.canvas.SetToNewScale(DrawFlag=True)
        # If the snapshot that was passed in does NOT have a defined Scale ...
        else:
            # ... size the image to fit the frame
            self.canvas.ZoomToBB()
        # Position the image according to the snapshot object's settings
        self.canvas.ViewPortCenter = [self.obj.image_coords[0], self.obj.image_coords[1]]
        # If we're showing the window ...
        if showWindow:
            # Show the Frame
            self.Show(True)
            # Bring this window to the front, so it doesn't get lost on some computers
            wx.CallLater(500, self.Raise)
        # Draw the initial codingObjects
        self.FileRedraw(None)

        # Call Yield so everything gets drawn properly
        wx.GetApp().Yield(True)

    def AddWindowMenuItem(self, itemName, itemNumber):
        """ Add an item to this Snapshot Window's Window menu """
        # Let's go ahead and keep the menu for non-Mac platforms
        if self.showWindow and (not '__WXMAC__' in wx.PlatformInfo):
            # Get an Item ID
            id = wx.NewId()
            # Add the Menu Item
            newItem = self.menuWindow.Append(id, itemName)
            # Add the Snapshot Number to the Menu Item's Help, which isn't shown so can hold this data
            newItem.SetHelp("%s" % itemNumber)
            # Bind the ID to the Menu Handler
            wx.EVT_MENU(self, id, self.OnWindowMenuItem)

    def UpdateWindowMenuItem(self, oldName, oldNumber, newName, newNumber):
        """ Update an item from this Snapshot Window's Window menu when a snapshot has been changed via Prev / Next buttons """
        # Let's go ahead and keep the menu for non-Mac platforms
        if self.showWindow and (not '__WXMAC__' in wx.PlatformInfo):
            # Iterate through all of the Window Menu Items
            for item in self.menuWindow.GetMenuItems():
                # Find the item with the correct name and number
                if (oldName == item.GetLabel()) and (oldNumber == int(item.GetHelp())):
                    # Update the Menu Label and Menu's Help (which indicates the Snapshot Number)
                    item.SetItemLabel(newName)
                    item.SetHelp("%s" % newNumber)
                    # We don't need to look any more
                    break

    def OnWindowMenuItem(self, event):
        """ Handle the Selection of an item in the Window Menu """
        # Let's go ahead and keep the menu for non-Mac platforms
        if self.showWindow and (not '__WXMAC__' in wx.PlatformInfo):
            # Get the name and number of the menu item selected
            itemName = self.menuWindow.GetLabel(event.GetId())
            itemNumber = int(self.menuWindow.GetHelpString(event.GetId()))
            # Have the Control Object select the appropriate Snapshot Window
            self.ControlObject.SelectSnapshotWindow(itemName, itemNumber)

    def DeleteWindowMenuItem(self, itemName, itemNumber):
        """ Remove an item from this Snapshot Window's Window menu """
        # Let's go ahead and keep the menu for non-Mac platforms
        if self.showWindow and (not '__WXMAC__' in wx.PlatformInfo):
            # Iterate through all of the Window Menu Items
            for item in self.menuWindow.GetMenuItems():
                # Find the item with the correct name and number
                if (itemName == self.menuWindow.GetLabel(item.GetId())) and (itemNumber == int(self.menuWindow.GetHelpString(item.GetId()))):
                    # Delete the menu item
                    self.menuWindow.Delete(item.GetId())
                    # We don't need to look any more
                    break

    # The FloatCanvas requires a mechanism for binding and un-binding events
    def BindEvents(self):
        """ Bind the FloatCanvas Events """
        # If the events are not bound ...
        if not self.eventsAreBound:
            # Bind the FloatCanvas' Motion, LeftDown, LeftUp, RightDown, and RightUp events
            self.canvas.Bind(FloatCanvas.EVT_MOTION, self.OnCanvasMotion)
            self.canvas.Bind(FloatCanvas.EVT_LEFT_DOWN, self.OnCanvasLeftDown)
            self.canvas.Bind(FloatCanvas.EVT_LEFT_UP, self.OnCanvasLeftUp)
            self.canvas.Bind(FloatCanvas.EVT_RIGHT_DOWN, self.OnCanvasRightDown)
            self.canvas.Bind(FloatCanvas.EVT_RIGHT_UP, self.OnCanvasRightUp)
            # Note that the events are now bound!
            self.eventsAreBound = True

    def UnbindEvents(self):
        """ Unbind the FloatCanvas Events """
        # Bind the FloatCanvas' Motion, LeftDown, LeftUp, RightDown, and RightUp events
        self.canvas.Unbind(FloatCanvas.EVT_MOTION)
        self.canvas.Unbind(FloatCanvas.EVT_LEFT_DOWN)
        self.canvas.Unbind(FloatCanvas.EVT_LEFT_UP)
        self.canvas.Unbind(FloatCanvas.EVT_RIGHT_DOWN)
        self.canvas.Unbind(FloatCanvas.EVT_RIGHT_UP)
        # Note that the events are now unbound!
        self.eventsAreBound = False

    def OnToolbar(self, event):
        """ Handle presses for many of the Toolbar's buttons """
        # Get the ID of the control that triggered this event
        eventID = event.GetId()
        # If Edit Tool ...
        if eventID == self.editTool.GetId():
            # ... if we're entering EDIT mode ...
            if self.editTool.IsToggled():
                # Start exception handling
                try:
                    # Remember the current Last Save Time
                    tmpLastSaveTime = self.obj.lastsavetime
                    # ... try to lock the Snapshot
                    self.obj.lock_record()

                    # If the Last Save Time was changed during the act of locking the record ...
                    if tmpLastSaveTime != self.obj.lastsavetime:
                        # ... inform the user that the Snapshot has been updated
                        msg = _('This Snapshot has been updated since you originally loaded it!\nYour copy of the record will be refreshed to reflect the changes.')
                        dlg = Dialogs.InfoDialog(self, msg)
                        dlg.ShowModal()
                        dlg.Destroy()

                        # Get the new Scale
                        self.canvas.Scale = self.obj.image_scale
                        # Get the new Position information
                        self.canvas.ViewPortCenter = [self.obj.image_coords[0], self.obj.image_coords[1]]
                        # resize the image window to the new Size
                        self.SetSize((self.obj.image_size[0], self.obj.image_size[1]))
                        # Re-set the Coding Object Number to the number of existing objects
                        self.objectNum = len(self.obj.codingObjects)
                        # Freeze the image
                        self.canvas.Freeze()
                        # Clear the Image (without deleting the coding)
                        self.FileClear(None)
                        # Redraw the coding
                        self.FileRedraw(None)
                        # Thaw the image
                        self.canvas.Thaw()
                    # If a Keyword has been selected ...                        
                    if (self.keyword_group_cb.GetStringSelection() != '') and (self.keyword_cb.GetStringSelection() != ''):
                        # ... enable the Coding Configuration tools
                        self.line_color_cb.Enable(True)
                        self.codeShape.Enable(True)
                        self.lineSize.Enable(True)
                        self.line_style_cb.Enable(True)
                        self.kwShowHideButton.Enable(True)
                        # If the current Keyword Group : Keyword selection in the interface is NOT already defined
                        # in the Snapshot's Keyword Styles ...
                        if not (self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection()) in \
                               self.obj.keywordStyles.keys():
                            # ... then add the current Coding Configuration information to the Snapshot's Keyword Styles
                            self.obj.keywordStyles[(self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection())] = \
                                { 'drawMode'       :  self.drawMode,
                                  'lineColorName'  :  self.colorList[self.line_color_cb.GetStringSelection()][0], # self.line_color_cb.GetStringSelection(),
                                  'lineColorDef'   :  "#%02x%02x%02x" % self.colorList[self.line_color_cb.GetStringSelection()][1], # "#%02x%02x%02x" % TransanaGlobal.transana_colorLookup[self.line_color_cb.GetStringSelection()],
                                  'lineWidth'      :  self.lineSize.GetStringSelection(),
                                  'lineStyle'      :  self.lineStyle  }

                # Handle "RecordLockedError" exception
                except TransanaExceptions.RecordLockedError, e:
                    # Display the Exception information to the user
                    TransanaExceptions.ReportRecordLockedException(_("Snapshot"), self.obj.id, e)
                    # Reject the attempt to go into Edit mode
                    self.toolbar.ToggleTool(self.editTool.GetId(), False)
            # If we are LEAVING EDIT mode ...
            else:
                # ... save the image changes
                self.LeaveEditMode()
                # Disable the Coding Configuration Tools
                self.line_color_cb.Enable(False)
                self.codeShape.Enable(False)
                self.lineSize.Enable(False)
                self.line_style_cb.Enable(False)
                # If the Pointer Tool is selected ...
                if self.pointer.IsToggled():
                    # ... switch to the Move Tool ...
                    self.toolbar.ToggleTool(self.moveTool.GetId(), True)
                    # ... and switch the GUI Mode to match
                    self.canvas.SetMode(GUIMode.GUIMove())
        # If Pointer Tool ...
        elif eventID == self.pointer.GetId():
            # ... set the canvas mode to GUITransana, which DRAWS
            self.canvas.SetMode(GUITransana())
        # If Move Tool ...
        elif eventID == self.moveTool.GetId():
            # ... set the canvas mode to FloatCanvas' GUIMove, which MOVES the canvas
            self.canvas.SetMode(GUIMode.GUIMove())
        # If Zoom In Tool ...
        elif eventID == self.zoomIn.GetId():
            # ... set the canvas mode to Transana's modified version of FloatCanvas' GUIZoomIn, which Zooms
            self.canvas.SetMode(GUIMode.GUIZoomIn())
        # If Zoom Out Tool ...
        elif eventID == self.zoomOut.GetId():
            # ... set the canvas mode to FloatCanvas' GUIZoomOut, which Zooms (out)
            self.canvas.SetMode(GUIMode.GUIZoomOut())
        # If Zoom To Fit Tool ...
        elif eventID == self.zoomToFit.GetId():
            # ... zoom to the canvas' Bounding Box (this accompishes the zoom) ...
            self.canvas.ZoomToBB()
            # ... and shift the focus to the canvas rather than the toolbar
            self.canvas.SetFocus()
        # If Zoom To Full Tool ...
        elif eventID == self.zoomToFull.GetId():
            # ... set the canvas' scale to 1 ...
            self.canvas.Scale= 1
            # ... and apply the scale change to the canvas.
            self.canvas.SetToNewScale(DrawFlag=True)
            # Set the focus to the canvas rather than the toolbar
            self.canvas.SetFocus()
        # If the Shape is selected ...
        elif eventID == self.codeShape.GetId():
            # ... set the Draw Mode to the selected shape
            if self.codeShape.GetStringSelection().encode('utf8') == _('Rectangle'):
                self.drawMode = 'Rectangle'
            elif self.codeShape.GetStringSelection().encode('utf8') == _('Ellipse'):
                self.drawMode = 'Ellipse'
            elif self.codeShape.GetStringSelection().encode('utf8') == _('Line'):
                self.drawMode = 'Line'
            elif self.codeShape.GetStringSelection().encode('utf8') == _('Arrow'):
                self.drawMode = 'Arrow'
        # If the Color is selected ...
        elif eventID == self.line_color_cb.GetId():
            # ... get the color RGB definition for the selected color ...
            color = self.colorList[self.line_color_cb.GetStringSelection()][1] # TransanaGlobal.transana_colorLookup[self.line_color_cb.GetStringSelection()]
            # ... and set the Draw Color to the selected color
            self.drawColor = wx.Colour(color[0], color[1], color[2])
        # If the Line Width is selected ...
        elif eventID == self.lineSize.GetId():
            # ... set the Draw Line Width
            self.lineWidth = int(self.lineSize.GetStringSelection())
        # If the Line Style is selected ...
        elif eventID == self.line_style_cb.GetId():
            # ... set the style to None
            style = None
            # Translate the user's selection to the style the FloatCanvas needs
            if self.line_style_cb.GetStringSelection().encode('utf8') == _('Solid'):
                style = 'Solid'
            elif self.line_style_cb.GetStringSelection().encode('utf8') == _('Dot'):
                style = 'Dot'
            elif self.line_style_cb.GetStringSelection().encode('utf8') == _('Dash'):
                style = 'LongDash'
            # ShortDash is indistinguishable from LongDash, at least on Windows, so it's not supported
#            elif self.line_style_cb.GetStringSelection().encode('utf8') == 'Short Dash':
#                style = 'ShortDash'
            elif self.line_style_cb.GetStringSelection().encode('utf8') == _('Dot Dash'):
                style = 'DotDash'
            # If a valid Style has been selected ...
            if style:
                # ... set the Draw Line Style
                self.lineStyle = style
        # If one of the Coding Configuration controls was used ...
        if eventID in [self.codeShape.GetId(), self.line_color_cb.GetId(), self.lineSize.GetId(), self.line_style_cb.GetId()]:
            # ... update (or add) the current configuration to the Shapshot's Keyword Styles
            self.obj.keywordStyles[(self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection())] = \
                { 'drawMode'       :  self.drawMode,
                  'lineColorName'  :  self.colorList[self.line_color_cb.GetStringSelection()][0], # self.line_color_cb.GetStringSelection(),
                  'lineColorDef'   :  "#%02x%02x%02x" % self.colorList[self.line_color_cb.GetStringSelection()][1], # TransanaGlobal.transana_colorLookup[self.line_color_cb.GetStringSelection()],
                  'lineWidth'      :  self.lineSize.GetStringSelection(),
                  'lineStyle'      :  self.lineStyle  }

            # If there are Draw Objects ...  (If there aren't, calling these lines messes up image position until the image is clicked!)
            if len(self.obj.codingObjects) > 0:
                # Freeze the image
                self.canvas.Freeze()
                # Clear the Image (without deleting coding)
                self.FileClear(None)
                # Redraw the coding
                self.FileRedraw(None)
                # Thaw the Image
                self.canvas.Thaw()

    def OnInsertIntoTranscript(self, event):
        """ Insert the current sized & coded Snapshot into the current editable Transcript! """
        # If the current transcript is in Read Only mode ...
        if self.ControlObject.ActiveTranscriptReadOnly():
            # ... inform the user
            msg = _("The current document is not editable.  The requested snapshot cannot be inserted into the document.")
            msg += '\n\n' + _("To insert the snapshot into the document, press the Edit Mode button on the Document Toolbar to make the document editable.")
            dlg = Dialogs.InfoDialog(self, msg)
            dlg.ShowModal()
            dlg.Destroy()
        else:
            # Create a Temporary Image File Name
            filename = os.path.join(TransanaGlobal.configData.visualizationPath, 'Temp.jpg')
            # Save the current selection to the temporary file
            self.FileSaveSelectionAs(event, filename = filename)
            # Load the TEMP file into Transcript"
            self.ControlObject.TranscriptInsertImage(filename, self.obj.number)

    def OnEditKeywords(self, event):
        """ Edit whole snapshot keywords """
        # Start Exception Handling
        try:
            # If the Snapshot isn't already locked ...
            if not self.obj.isLocked:
                # Remember the current Last Save Time
                tmpLastSaveTime = self.obj.lastsavetime
                # ... lock the Snapshot ...
                self.obj.lock_record()
                # If the Last Save Time was changed during the act of locking the record ...
                if tmpLastSaveTime != self.obj.lastsavetime:
                    # ... inform the user that the Snapshot has been updated
                    msg = _('This Snapshot has been updated since you originally loaded it!\nYour copy of the record will be refreshed to reflect the changes.')
                    dlg = Dialogs.InfoDialog(self, msg)
                    dlg.ShowModal()
                    dlg.Destroy()

                    # Get the new Scale
                    self.canvas.Scale = self.obj.image_scale
                    # Get the new Position information
                    self.canvas.ViewPortCenter = [self.obj.image_coords[0], self.obj.image_coords[1]]
                    # resize the image window to the new Size
                    self.SetSize((self.obj.image_size[0], self.obj.image_size[1]))
                    # Re-set the Coding Object Number to the number of existing objects
                    self.objectNum = len(self.obj.codingObjects)
                    # Freeze the image
                    self.canvas.Freeze()
                    # Clear the Image (without deleting the coding)
                    self.FileClear(None)
                    # Redraw the coding
                    self.FileRedraw(None)
                    # Thaw the image
                    self.canvas.Thaw()

                # ... and remember we locked it here
                lockedRecord = True

                # We need to refresh the Keyword List.
                # See, if someone has deleted a Keyword (or Keyword Group) while this Snapshot was
                # open, it could be out of date without LastSaveTime being updated!
                self.obj.refresh_keywords()

            # If  the Snapshot IS already locked ...
            else:
                # ... remember that we did NOT lock it here
                lockedRecord = False

            # Determine the title for the KeywordListEditForm Dialog Box
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Keywords for %s"), 'utf8')
            else:
                prompt = _("Keywords for %s")
            dlgTitle = prompt % self.obj.id
            # Extract the keyword List from the Data object
            kwlist = []
            for kw in self.obj.keyword_list:
                 kwlist.append(kw)

            # Create/define the Keyword List Edit Form
            dlg = KeywordListEditForm.KeywordListEditForm(self, -1, dlgTitle, self.obj, kwlist)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Show the Keyword List Edit Form and process it if the user selects OK
                    if dlg.ShowModal() == wx.ID_OK:
                        # Clear the local keywords list and repopulate it from the Keyword List Edit Form
                        kwlist = []
                        for kw in dlg.keywords:
                            kwlist.append(kw)

                        # Copy the local keywords list into the appropriate object
                        self.obj.keyword_list = kwlist

                        # If we locked the Snapshot here ...
                        if lockedRecord:
                            # Save the Data object
                            self.obj.db_save()
                            # Re-load the Snapshot so we don't get an error message about having an OLD copy of the Snapshot Data
                            self.obj = Snapshot.Snapshot(self.obj.number)

                            # Update the Keyword Visualization, if needed
                            self.ControlObject.UpdateKeywordVisualization()
                            # Even if this computer doesn't need to update the keyword visualization others, might need to.
                            if not TransanaConstants.singleUserVersion and (self.obj.episode_num != 0):
                                # We need to update the Episode Keyword Visualization
                                if DEBUG:
                                    print 'Message to send = "UKV %s %s %s"' % ('Episode', self.obj.episode_num, 0)
                                    
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', self.obj.episode_num, 0))


                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # Refresh the Keyword List, if it's a changed Keyword error
                    dlg.refresh_keywords()
                    # Highlight the first non-existent keyword in the Keywords control
                    dlg.highlight_bad_keyword()

                # Handle other exceptions
                except:
                    if DEBUG:
                        import traceback
                        traceback.print_exc(file=sys.stdout)
                    # Display the Exception Message, allow "continue" flag to remain true
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_("Exception %s: %s"), 'utf8')
                    else:
                        prompt = _("Exception %s: %s")
                    errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                    errordlg.ShowModal()
                    errordlg.Destroy()
            # If we locked the Snapshot here ...
            if lockedRecord:
                # ... release the record lock
                self.obj.unlock_record()

        # Handle record lock exceptions
        except TransanaExceptions.RecordLockedError, e:
            """Handle the RecordLockedError exception."""
            TransanaExceptions.ReportRecordLockedException(_('Snapshot'), self.obj.id, e)

    def OnCodingKey(self, event):
        """ Display the Snapshot Coding Key """
        # We need to determine what styles are visible.  INitialize a variable to hold the visible styles
        visibleKeywordStyles = {}
        # Iterate through the coding objects ...
        for x in self.obj.codingObjects.keys():
            # If the current Coding Object is visible AND is not yet represented in the Visible Styles dictionary ...
            if (self.obj.codingObjects[x]['visible']) and \
               not ((self.obj.codingObjects[x]['keywordGroup'], self.obj.codingObjects[x]['keyword']) in visibleKeywordStyles.keys()):
                # ... then add its style to the Visible Styles dictionary
                visibleKeywordStyles[(self.obj.codingObjects[x]['keywordGroup'], self.obj.codingObjects[x]['keyword'])] = \
                    self.obj.keywordStyles[(self.obj.codingObjects[x]['keywordGroup'], self.obj.codingObjects[x]['keyword'])]
        # If there is a current Coding Key Popup display ...
        if self.codingKeyPopup != None:
            # ... start exception handling ...
            try:
                # ... close the Coding Key Poup display
                self.codingKeyPopup.Close()
            # Ignore exceptions
            except:
                pass

        # Determine which monitor to use and get its size and position
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            primaryScreen = TransanaGlobal.configData.primaryScreen
        else:
            primaryScreen = 0
        rect = wx.Display(primaryScreen).GetClientArea()  # wx.ClientDisplayRect()

        # Get the position and size of the current Snapshot Window
        pos = self.GetPositionTuple()
        size = self.GetSizeTuple()
        # Create a Coding Key Popup using the visible styles just to the right of the Snapshot Window
        self.codingKeyPopup = CodingKeyWindow(self, visibleKeywordStyles, wx.Point(pos[0] + size[0], pos[1]))
        # Display the Coding Key Popup
        self.codingKeyPopup.Show()
        # Get the position and size of the popup window (which isn't correct until we Show() it!)
        winPos = self.codingKeyPopup.GetRect()
        # If the left + width > monitor size ...
        if winPos[0] + winPos[2] > rect[0] + rect[2]:
            # Adjust the position so the image stays on screen!
            pos = (rect[0] + rect[2]  - winPos[2], winPos[1])
            self.codingKeyPopup.SetPosition(pos)

    def OnChangeSnapshot(self, event):
        """ Handle the Previous and Next Snapshot buttons """

        # Remember the old Snapshot / Snapshot Window / Snapshot Window Menu Item name
        oldSnapshotName = self.obj.id
        oldSnapshotNumber = self.obj.number

        # if we're in Edit Mode ...
        if self.showWindow and self.editTool.IsToggled():
            # ... leave Edit Mode
            self.LeaveEditMode()

            # If LeaveEditMode did NOT leave Edit Mode ...
            if self.editTool.IsToggled():
                # ... there is an error condition and we need to VETO the close!
                event.Veto()
                # Reset the closing flag.  We're not closing after all.
                self.closing = False
                # We need to exit this method now for the Veto to occur properly.
                return

        # If the Coding Key is visible ...
        if self.codingKeyPopup != None:
            # Start Exception handling
            try:
                # ... close the Coding Key
                self.codingKeyPopup.Close()
            # Ignore exceptions
            except:
                pass

        # If the Previous Button was pressed ...
        if event.GetId() == self.prevSnapBtn.GetId():
            # ... load the Previous Snapshot
            newSnapshot = Snapshot.Snapshot(self.prevSnapshot)
        # If the Next Button was pressed ...
        elif event.GetId() == self.nextSnapBtn.GetId():
            # ... load the Next Snapshot
            newSnapshot = Snapshot.Snapshot(self.nextSnapshot)

        # Check to see if the image file can be found
        if not os.path.exists(newSnapshot.image_filename):
            # If not, raise an exception
            errmsg = unicode(_("Image file not found:\n%s"), 'utf8')
            raise TransanaExceptions.ImageLoadError(errmsg % newSnapshot.image_filename)

        # Load the image
        self.bgImage = wx.Image(newSnapshot.image_filename)
            
        # Make sure the image is loaded, is not corrupt
        if not self.bgImage.IsOk():
            # If not, raise an exception
            errmsg = unicode(_("Unable to load image file:\n%s\nThere may be a problem with the file, or you may\nhave too many Snapshots open."), 'utf8')
            raise TransanaExceptions.ImageLoadError(errmsg % newSnapshot.image_filename)

        # Change Window Menu Items in MenuWindow and all Snapshot Windows.  This can be done via the ControlObject.
        # This has to be done NOW in case the Snapshot being loaded is loaded and edited in another window!!
        self.ControlObject.UpdateWindowMenu(oldSnapshotName, oldSnapshotNumber, newSnapshot.id, newSnapshot.number)

        # Now we have to re-load the Snapshot, in case it was updated in another window that just got closed.
        self.obj = Snapshot.Snapshot(newSnapshot.number)

        # Clear the Canvas to remove the previous Snapshot
        self.canvas.ClearAll()

        # Change the Window Name
        self.SetTitle(_("Snapshot") + u' - ' + self.obj.GetNodeString(True))

        # If the image that is passed in has a defined window size ...
        if self.obj.image_size[0] > 0:
            # ... use that defined size
            width = self.obj.image_size[0]
            height = self.obj.image_size[1]
        # If the image does NOT have a size ...
        else:
            # ... use the default image window size
            width = self.__size()[0]
            height = self.__size()[1]

        # Resize the window 
        self.SetSize((width, height))

        # Get a list of all Snapshots in the same collection
        self.snapshotList = DBInterface.list_of_snapshots_by_collectionnum(self.obj.collection_num, True)
        # Initialize values for Previous and Next Snapshots
        self.prevSnapshot = 0
        self.nextSnapshot = 0
        # Determine the list index for the current Snapshot
        index = self.snapshotList.index((self.obj.number, self.obj.id, self.obj.collection_num, self.obj.sort_order))
        # If the current snapshot isn't the first item in the list ...
        if index > 0:
            # ... then remember the previous snapshot's number
            self.prevSnapshot = self.snapshotList[index - 1][0]
            self.toolbar.EnableTool(self.prevSnapBtn.GetId(), True)
        else:
            self.toolbar.EnableTool(self.prevSnapBtn.GetId(), False)
        # If the current snapshot isn't the last item in the list ...
        if index < len(self.snapshotList) - 1:
            # ... then remember the next snapshot's number
            self.nextSnapshot = self.snapshotList[index + 1][0]
            self.toolbar.EnableTool(self.nextSnapBtn.GetId(), True)
        else:
            self.toolbar.EnableTool(self.nextSnapBtn.GetId(), False)

        # The FloatCanvas ScaledBitmap object seems to run into problems when zoomed in too far.  Large images raise
        # an exception on Zoom In with click and Wheel zooms, as well as with selection-based zoomed.  I contacted Chris
        # Barker, who wrote FloatCanvas, and he suggested I try ScaledBitmap2, which requires a 2-step creation process.
        # It seems to work.
        # Add a Scaled Bitmap, converted from the loaded image, to the SnapshotWindow canvas
        bgBitmapObj = FloatCanvas.ScaledBitmap2(self.bgImage,  # wx.BitmapFromImage(self.bgImage),
                                               (0 - (float(self.bgImage.GetWidth()) / 2.0), (float(self.bgImage.GetHeight()) / 2.0)),
                                               Height = self.bgImage.GetHeight(),
                                               Position = "tl")
        self.canvas.AddObject(bgBitmapObj)

        # Set the minimum scale to 1/20th normal size
        self.canvas.MinScale = 0.05
        # Set the maximum scale to 3.7 times normal size.  (This avoids unsightly but not serious errors with large images.)
        self.canvas.MaxScale = 5.0  # 3.70
        if self.showWindow:
            # Select the Move Tool initially
            self.toolbar.ToggleTool(self.moveTool.GetId(), True)
        # Set the Canvas Mode to the Move Tool initially
        self.canvas.SetMode(GUIMode.GUIMove())

        # If the snapshot that was passed in has a defined Scale ...
        if self.obj.image_scale > 0.0:
            # ... set the canvas' scale to the snapshot's value ...
            self.canvas.Scale = self.obj.image_scale
            # ... and apply the scale change to the canvas.
            self.canvas.SetToNewScale(DrawFlag=True)
        # If the snapshot that was passed in does NOT have a defined Scale ...
        else:
            # ... size the image to fit the frame
            self.canvas.ZoomToBB()
        # Position the image according to the snapshot object's settings
        self.canvas.ViewPortCenter = [self.obj.image_coords[0], self.obj.image_coords[1]]
        # Draw the initial codingObjects
        self.FileRedraw(None)

        # Call Yield so everything gets drawn properly
        wx.GetApp().Yield(True)
            
    def OnKWGSelect(self, event):
        """ Handle the selection of a Keyword Group """
        # Initialize a list for Keyword Choices
        choices = ['']
        # If a Keyword Group has been selected ...
        if self.keyword_group_cb.GetStringSelection() != '':
            # ... get that Keyword Group's Keywords from the Database
            choices += DBInterface.list_of_keywords_by_group(self.keyword_group_cb.GetStringSelection())
        # Set the Keyword Choice Box with the appropriate Keyword Choices 
        self.keyword_cb.SetItems(choices)
        # Select the first item in the list, the blank one.
        self.keyword_cb.Select(0)
        # Because we always blank the Keyword, we always need to disable the Coding Configuration Tools
        self.line_color_cb.Enable(False)
        self.codeShape.Enable(False)
        self.lineSize.Enable(False)
        self.line_style_cb.Enable(False)
        self.kwShowHideButton.Enable(False)

    def OnKWSelect(self, event):
        """ Handle the selection of a Keyword """
        # If NO Keyword is selected ...
        if self.keyword_cb.GetStringSelection() == '':
            # ... disable the Coding Configuration Tools
            self.line_color_cb.Enable(False)
            self.codeShape.Enable(False)
            self.lineSize.Enable(False)
            self.line_style_cb.Enable(False)
            self.kwShowHideButton.Enable(False)
        # If a Keyword is selected ...
        else:
            # ... and if the Edit Tool is toggled ...
            if self.editTool.IsToggled():
                # ... enable the Coding Configuration Tools
                self.line_color_cb.Enable(True)
                self.codeShape.Enable(True)
                self.lineSize.Enable(True)
                self.line_style_cb.Enable(True)

            # Let's set the coding specs to match the specified Coding Characteristics for the chosen keyword
            if not (self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection()) in self.obj.keywordStyles.keys():

                # Get the Keyword Definition
                tmpKeyword = KeywordObject.Keyword(self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection())
                                
                # Use the defined Coding Defaults, if they exist, but keep using the current values if not.
                # NOTE that the tmpKeyword may have some values but not others.

                # If no default Keyword Line Color is defined ...
                if tmpKeyword.lineColorName == '':
                    # ... use the current value from the form
                    tmpLineColorName = self.colorList[self.line_color_cb.GetStringSelection()][0] # self.line_color_cb.GetStringSelection()
                    tmpLineColorDef = "#%02x%02x%02x" % self.colorList[self.line_color_cb.GetStringSelection()][1] # TransanaGlobal.transana_colorLookup[self.line_color_cb.GetStringSelection()]
                # If the default Keyword Line Color is defined ...
                else:
                    tmpLineColorName = tmpKeyword.lineColorName
                    tmpLineColorDef = tmpKeyword.lineColorDef

                # If no default Keyword Draw Mode is defined ...
                if tmpKeyword.drawMode == '':
                    # ... use the current value from the form
                    tmpDrawMode = self.drawMode
                # If the default Keyword Draw Mode is defined ...
                else:
                    # ... use the Keyword Default
                    tmpDrawMode = tmpKeyword.drawMode
                    # ... and change the value for the form to use
                    self.drawMode = tmpDrawMode

                # If no default Keyword Line Width is defined ...
                if tmpKeyword.lineWidth == 0:
                    # ... use the current value from the form
                    tmpLineWidth = self.lineWidth
                # If the default Keyword Line Width is defined ...
                else:
                    # ... use the Keyword Default
                    tmpLineWidth = tmpKeyword.lineWidth
                    # ... and change the value for the form to use
                    self.lineWidth = tmpLineWidth

                # If no default Keyword Line Style is defined ...
                if tmpKeyword.lineStyle == '':
                    # ... use the current value from the form
                    tmpLineStyle = self.lineStyle
                # If the default Keyword Line Style is defined ...
                else:
                    # ... use the Keyword Default
                    tmpLineStyle = tmpKeyword.lineStyle
                    # ... and change the value for the form to use
                    self.lineStyle = tmpLineStyle

                # ... save the current selections as this Keyword's Style
                self.obj.keywordStyles[(self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection())] = \
                    { 'drawMode'       :  tmpDrawMode,
                      'lineColorName'  :  tmpLineColorName,
                      'lineColorDef'   :  tmpLineColorDef,
                      'lineWidth'      :  '%s' % tmpLineWidth,
                      'lineStyle'      :  tmpLineStyle  }

            # Get the Keyword Styles for the selected Keyword Group : Keyword from the Snapshot
            keyVals = self.obj.keywordStyles[(self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection())]
            # Set the Draw Mode
            self.drawMode = keyVals['drawMode']
            # Set the Code Shape Tool Selection to match the Draw Mode
            if self.drawMode == 'Arrow':
                self.codeShape.SetStringSelection(_('Arrow'))
            elif self.drawMode == 'Line':
                self.codeShape.SetStringSelection(_('Line'))
            elif self.drawMode == 'Rectangle':
                self.codeShape.SetStringSelection(_('Rectangle'))
            elif self.drawMode == 'Ellipse':
                self.codeShape.SetStringSelection(_('Ellipse'))

            # If the Line Color's Name is in the defined graphics colors ...
            if keyVals['lineColorName'] in TransanaGlobal.transana_colorLookup.keys():
                # ... get the Color Definition by looking it up in the Global List.
                #     (Thus, it will be correct even if the user changes the global color definitions!)
                color = TransanaGlobal.transana_colorLookup[keyVals['lineColorName']]
            # If the Color Name isn't in the defined graphics colors ...
            else:
                # ... use the Color Definition saved with the Snapshot's Keyword Styles
                color = (int(keyVals['lineColorDef'][1:3], 16), int(keyVals['lineColorDef'][3:5], 16), int(keyVals['lineColorDef'][5:7], 16))
            # Set the Line Draw Color to the defined color definition
            self.drawColor = wx.Colour(color[0], color[1], color[2])
            # Set the Color Tool Selection to the Color
            self.line_color_cb.SetStringSelection(_(keyVals['lineColorName']))

            # Set the Line Width
            self.lineWidth = int(keyVals['lineWidth'])
            # Set the Line Size Tool Selection
            self.lineSize.SetStringSelection(keyVals['lineWidth'])

            # Set the Line Style
            self.lineStyle = keyVals['lineStyle']

            # Set the Line Style Tool Selection 
            if self.lineStyle == 'Solid':
                self.line_style_cb.SetStringSelection(_('Solid'))
            elif self.lineStyle == 'Dot':
                self.line_style_cb.SetStringSelection(_('Dot'))
            elif self.lineStyle == 'LongDash':
                self.line_style_cb.SetStringSelection(_('Dash'))
            elif self.lineStyle == 'Dot':
                self.line_style_cb.SetStringSelection(_('Dot'))
            elif self.lineStyle == 'DotDash':
                self.line_style_cb.SetStringSelection(_('Dot Dash'))

            # Enable the Show/Hide Button
            self.kwShowHideButton.Enable(True)
            # We have to figure out what the label should be for the Show/Hide button.  Assume HIDE to start.
            defaultLabel = _('Hide')
            # Get the keys for the coding objects
            keys = self.obj.codingObjects.keys()
            # Iterate through the coding objects
            for key in keys:
                # Get this key's coding object
                obj = self.obj.codingObjects[key]
                # If the coding object's keyword matches the form's keyword selection ...
                if (obj['keywordGroup'], obj['keyword']) == (self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection()):
                    # if ANY instances of the selected keyword are HIDDEN ...
                    if not obj['visible']:
                        # ... then the button should read SHOW
                        defaultLabel = _('Show')
                        # ... and we don't need to look any more
                        break
            # Set the Show/Hide button to the appropriate label
            self.kwShowHideButton.SetLabel(defaultLabel)

    def OnKWShowHideTool(self, event):
        """ Show or Hide the coding for a given keyword """
        # Get the keys for the coding objects
        keys = self.obj.codingObjects.keys()
        # Flag whether anything has changed on screen
        changed = False
        # Iterate through the coding objects
        for key in keys:
            # Get this key's coding object
            obj = self.obj.codingObjects[key]
            # If the coding object's keyword matches the form's keyword selection ...
            if (obj['keywordGroup'], obj['keyword']) == (self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection()):
                # ... if we're SHOWING coding ...
                if self.kwShowHideButton.GetLabel().encode('utf8') == _('Show'):
                    # ... then make the object's coding visible
                    obj['visible'] = True
                # ... if we're HIDING coding ...
                elif self.kwShowHideButton.GetLabel().encode('utf8') == _('Hide'):
                    # ... then make the object's coding hidden
                    obj['visible'] = False
                # Flag that the screen has (probably) changed
                changed = True

        # If the button says Show ...
        if self.kwShowHideButton.GetLabel().encode('utf8') == _('Show'):
            # ... change it's label to Hide
            self.kwShowHideButton.SetLabel(_('Hide'))
        # If the button says Hide ...
        else:
            # ... change it's label to Show
            self.kwShowHideButton.SetLabel(_('Show'))

        # If the image has changed ...    
        if changed:
            # ... update the screen
            self.canvas.Freeze()
            self.FileClear(None)
            self.FileRedraw(None)
            self.canvas.Thaw()

    def OnHelp(self, event):
        # ... call Help!
        self.ControlObject.Help("Snapshot Window")
        
    def OnCanvasMotion(self, event):
        """ EVT_MOTION for the Canvas """
        # Call the FloatCanvas' EVT_MOTION event so we get things like the zoom rubber band box
        event.Skip()

    def OnCanvasLeftDown(self, event):
        """ EVT_LEFT_DOWN for the Canvas """
        # If the Pointer Tool is selected ...
        if self.pointer.IsToggled():
            # ... then note the coordinates where the mouse button was pressed
            self.mouseDown = event.Coords
        # Call the FloatCanvas' EVT_LEFT_DOWN event
        event.Skip()

    def OnCanvasLeftUp(self, event):
        """ EVT_LEFT_UP for the Canvas """
        # Note the coordinates where the mouse button was released
        self.mouseUp = event.Coords
        # If we also know where the Mouse button was pressed, AND we're in Edit Mode AND Color Selection is enabled (so we have a Keyword) ...
        if (self.mouseDown != None) and (self.editTool.IsToggled()) and (self.line_color_cb.IsEnabled()):
            # Determine how much the mouse has moved
            mouseChange = self.mouseUp - self.mouseDown
            # If the mouse has moved at least 5 pixels in either direction ...  (Used to be 10, which was too big!)
            if (abs(mouseChange[0]) > 5) or (abs(mouseChange[1]) > 5):
                # Initialize the Draw Object to None
                drawObj = None
                # If we're drawing an Arrow ...
                if self.drawMode == 'Arrow':
                    # Add the Arrow to the canvas
                    drawObj = self.canvas.AddArrowLine((self.mouseDown[0], self.mouseDown[1], self.mouseUp[0], self.mouseUp[1]),
                                                       LineWidth = self.lineWidth, LineColor = self.drawColor,
                                                       LineStyle = self.lineStyle, ArrowHeadSize = 15, ArrowHeadAngle = 60)

                    # Remember the specifications for this code object
                    self.obj.codingObjects[self.objectNum] = {'x1'             :  self.mouseDown[0],
                                                              'y1'             :  self.mouseDown[1],
                                                              'x2'             :  self.mouseUp[0],
                                                              'y2'             :  self.mouseUp[1],
                                                              'keywordGroup'   :  self.keyword_group_cb.GetStringSelection(),
                                                              'keyword'        :  self.keyword_cb.GetStringSelection(),
                                                              'visible'        :  True}

                # If we're drawing a Line ...
                elif self.drawMode == 'Line':
                    # Add the Line to the canvas
                    drawObj = self.canvas.AddArrowLine((self.mouseDown[0], self.mouseDown[1], self.mouseUp[0], self.mouseUp[1]),
                                                       LineWidth = self.lineWidth, LineColor = self.drawColor,
                                                       LineStyle = self.lineStyle, ArrowHeadSize = 0)

                    # Remember the specifications for this code object
                    self.obj.codingObjects[self.objectNum] = {'x1'             :  self.mouseDown[0],
                                                        'y1'             :  self.mouseDown[1],
                                                        'x2'             :  self.mouseUp[0],
                                                        'y2'             :  self.mouseUp[1],
                                                        'keywordGroup'  :  self.keyword_group_cb.GetStringSelection(),
                                                        'keyword'       :  self.keyword_cb.GetStringSelection(),
                                                        'visible'       :  True}

                # If we're drawing a Rectangle ...
                elif self.drawMode == 'Rectangle':
                    # Calculate the position and size of the Rectangle
                    x1 = self.mouseDown[0]
                    y1 = self.mouseDown[1]
                    x2 = self.mouseUp[0]
                    y2 = self.mouseUp[1]
                    # Add the Rectangle to the canvas
                    drawObj = self.canvas.AddRectangle((x1, y1), (x2 - x1, y2 - y1),
                                                       LineWidth = self.lineWidth, LineColor = self.drawColor,
                                                       LineStyle = self.lineStyle)

                    # Remember the specifications for this code object
                    self.obj.codingObjects[self.objectNum] = {'x1'             :  x1,
                                                        'y1'             :  y1,
                                                        'x2'             :  x2,
                                                        'y2'             :  y2,
                                                        'keywordGroup'  :  self.keyword_group_cb.GetStringSelection(),
                                                        'keyword'       :  self.keyword_cb.GetStringSelection(),
                                                        'visible'       :  True}
                    
                # If we're drawing an Ellipse ...
                elif self.drawMode == 'Ellipse':
                    # Calculate the position and size of the Ellipse
                    x1 = self.mouseDown[0]
                    y1 = self.mouseDown[1]
                    x2 = self.mouseUp[0]
                    y2 = self.mouseUp[1]
                    # Add the Ellipse to the canvas
                    drawObj = self.canvas.AddEllipse((x1, y1), (x2 - x1, y2 - y1),
                                                     LineWidth = self.lineWidth, LineColor = self.drawColor,
                                                     LineStyle = self.lineStyle)

                    # Remember the specifications for this code object
                    self.obj.codingObjects[self.objectNum] = {'x1'             :  x1,
                                                        'y1'             :  y1,
                                                        'x2'             :  x2,
                                                        'y2'             :  y2,
                                                        'keywordGroup'  :  self.keyword_group_cb.GetStringSelection(),
                                                        'keyword'       :  self.keyword_cb.GetStringSelection(),
                                                        'visible'       :  True}

                # If we have a valid draw object ...
                if drawObj:
                    # Number the draw object
                    drawObj.ObjectNum = self.objectNum
                    # Name the draw object after the Keyword Group : Keyword
                    drawObj.Name = "%s : %s" % (self.keyword_group_cb.GetStringSelection(), self.keyword_cb.GetStringSelection())
                    # Specify the "Hit" size as slightly larger than the arrow line itself to make object selection easier
                    drawObj.HitLineWidth = self.lineWidth + 4
                    # Define the draw object's Left Click, Right Click, Enter, and Leave events
                    # These implement mouse-over labelling and the right-click menu
                    drawObj.Bind(FloatCanvas.EVT_FC_RIGHT_DOWN, self.OnItemRightDown)
                    drawObj.Bind(FloatCanvas.EVT_FC_ENTER_OBJECT, self.OnEnterObject)
                    drawObj.Bind(FloatCanvas.EVT_FC_LEAVE_OBJECT, self.OnLeaveObject)
                    # Increment the Draw Object counter
                    self.objectNum += 1
                    # Redraw the FloatCanvas
                    self.canvas.Draw()

        # If we CAN'T draw coding shapes ...
        if (not self.editTool.IsToggled()) or (not self.line_color_cb.IsEnabled()):
            # ... calling this removes the Bounding Box from the screen!!
            self.canvas.Draw()

        # Re-initialize the Mouse Down and Mouse Up coordinates
        self.mouseDown = None
        self.mouseUp = None

    def OnCanvasRightDown(self, event):
        """ EVT_RIGHT_DOWN for the Canvas """
        # Call the FloatCanvas' EVT_RIGHT_DOWN event
        event.Skip()

    def OnCanvasRightUp(self, event):
        """ EVT_RIGHT_UP for the Canvas """
        # Call the FloatCanvas' EVT_RIGHT_UP event
        event.Skip()

    def OnItemRightDown(self, Object):
        """ Handle EVT_RIGHT_DOWN for a Float Canvas Drawing Object """
        # Save the information for the object that was right-clicked so it's accessible later
        self.popupObject = Object
        # Create a Menu
        menu = wx.Menu()
        # Populate the Menu
        menu.Append(MENU_POPUP_HIDE, _("Hide"))
        wx.EVT_MENU(self, MENU_POPUP_HIDE, self.OnPopupMenu)
        menu.Append(MENU_POPUP_SENDTOBACK, _("Send To Back"))
        wx.EVT_MENU(self, MENU_POPUP_SENDTOBACK, self.OnPopupMenu)
        menu.Append(MENU_POPUP_DELETE, _("Delete"))
        wx.EVT_MENU(self, MENU_POPUP_DELETE, self.OnPopupMenu)
        # Have the Menu Pop Up for the User
        self.PopupMenu(menu, self.ScreenToClient(wx.GetMousePosition()))

    def OnEnterObject(self, Object):
        """ EVT_FC_ENTER_OBJCET for objects drawn on the Canvas """
        # Create a string to identify the object (Name is the Keyword!)
        str = Object.Name
        # Show the string in the Status Bar
        self.SetStatusText(str)

    def OnLeaveObject(self, Object):
        """ EVT_FC_LEAVE_OBJECT for objects drawn on the Canvas """
        # Clear the Status Bar on leaving a draw object
        self.SetStatusText("")

    def OnPopupMenu(self, event):
        """ Handle EVT_MENU events from the right-click popup menu """
        # If we have a HIDE command ...
        if event.GetId() == MENU_POPUP_HIDE:
            # ... set the appropriate Snapshot Coding Object's visible property to False
            self.obj.codingObjects[self.popupObject.ObjectNum]['visible'] = False
            # If the current object's Keyword Group : Keyword is currently selected in the interface ...
            if (self.obj.codingObjects[self.popupObject.ObjectNum]['keywordGroup'] == self.keyword_group_cb.GetStringSelection()) and \
               (self.obj.codingObjects[self.popupObject.ObjectNum]['keyword'] == self.keyword_cb.GetStringSelection()):
                # ... we need to change the Show/Hide Button to "Show".
                self.kwShowHideButton.SetLabel(_('Show'))
        # If we have a SEND TO BACK command ...
        elif event.GetId() == MENU_POPUP_SENDTOBACK:
            # ... get an integer one smaller than the smallest Coding Object's key value ...
            minVal = min(self.obj.codingObjects.keys()) - 1
            # ... and change the Coding Object's Key Value to this smallest value (using the
            #     dictionary's pop() method to move the data to a new key!
            self.obj.codingObjects[minVal] = self.obj.codingObjects.pop(self.popupObject.ObjectNum)
        # If we have a DELETE command ...
        elif event.GetId() == MENU_POPUP_DELETE:
            # ... delete the appropriate Coding Object from the Snapshot
            del(self.obj.codingObjects[self.popupObject.ObjectNum])

        # ... update the screen
        self.canvas.Freeze()
        self.FileClear(None)
        self.FileRedraw(None)
        self.canvas.Thaw()

    def FileClear(self, event):
        """ Clear the Coding from the Image
            If event is None, coding objects will be removed from the drawing but not deleted from the
              Snapshot object.  If event.GetId() is MENU_FILE_CLEAR, the coding objects will be deleted
              from the Snapshot. """
        # If this method is called from the File > Clear MENU ...
        if not(event is None) and (event.GetId() == MENU_FILE_CLEAR):
            # Initialize a dictionary of objects to draw on the image
            self.obj.codingObjects = {}
        # If not ...
        else:
            # remember the Shapshot's Scale and Coordinates
            oldScale = self.canvas.Scale
            oldCoords = self.canvas.ViewPortCenter

        # Unbind the FloatCanvas events
        self.UnbindEvents()
        # Initialize the FloatCanvas
        self.canvas.InitAll()
        # Re-bind the FloatCanvas events
        self.BindEvents()
        
        # Set the minimum scale to 1/10th normal size
        self.canvas.MinScale = 0.1
        # Set the maximum scale to 3.7 times normal size.  (This avoids unsightly but not serious errors with large images.)
        self.canvas.MaxScale = 3.70

        # Add a Scaled Bitmap, converted from the loaded image, to the SnapshotWindow canvas
        bgBitmapObj = FloatCanvas.ScaledBitmap2(wx.BitmapFromImage(self.bgImage),
                                               (0 - (float(self.bgImage.GetWidth()) / 2.0), (float(self.bgImage.GetHeight()) / 2.0)),
                                               Height = self.bgImage.GetHeight(),
                                               Position = "tl")
        self.canvas.AddObject(bgBitmapObj)

        # If this is called from the MENU ...
        if not(event is None) and (event.GetId() == MENU_FILE_CLEAR):
            # ... zoom to the Bounding Box, to fit the image on the screen
            self.canvas.ZoomToBB()
        # If NOT called from the MENU ...
        else:
            # ... restore the original Scale ...
            self.canvas.Scale = oldScale
            # ... and apply the scale change to the canvas ...
            self.canvas.SetToNewScale()
            # ... and restore the original Positioning
            self.canvas.ViewPortCenter = oldCoords
        # Draw the new Canvas
        self.canvas.Draw()

    def FileRestore(self, event):
        """ Restore the Last Saved Version of the Snapshot. """
        # If called from the Restore Last Save menu item or button ...
        if (('wxMSW'in wx.PlatformInfo) and (event.GetId() == MENU_FILE_RESTORE)) or \
           (('wxMac'in wx.PlatformInfo) and (event.GetId() == self.restoreButton.GetId())):
            # ... start by clearing the Snapshot to remove new coding added since the last save
            self.FileClear(event)
        
        # If there is a current Coding Key Popup display ...
        if self.codingKeyPopup != None:
            # ... start exception handling ...
            try:
                # ... close the Coding Key Poup display
                self.codingKeyPopup.Close()
            # Ignore exceptions
            except:
                pass
        # Initialize a dictionary of objects to draw on the image
        self.obj.codingObjects = {}
        # Reload the Snapshot Object
        self.obj.db_load(self.obj.number)
        # If a size is defined ...
        if self.obj.image_size != (0, 0):
            # Resize the Window
            self.SetSize(self.obj.image_size)
        # Resize the Snapshot
        # If the snapshot that was passed in has a defined Scale ...
        if self.obj.image_scale > 0.0:
            # ... set the canvas' scale to the snapshot's value ...
            self.canvas.Scale = self.obj.image_scale
            # ... and apply the scale change to the canvas.
            self.canvas.SetToNewScale(DrawFlag=True)
        # If the snapshot that was passed in does NOT have a defined Scale ...
        else:
            # ... size the image to fit the frame
            self.canvas.ZoomToBB()
        # Reposition the Snapshot
        self.canvas.ViewPortCenter = [self.obj.image_coords[0], self.obj.image_coords[1]]
        # Redraw the Snapshot
        self.FileRedraw(event)

    def FileRedraw(self, event):
        """ Redraws the coding on an image.  This is used when the coding has changed to allow the screen
            to reflect those changes. """
        # If called from the Hide All Coding menu item ...
        if not (event is None) and \
           ((event.GetId() == MENU_FILE_HIDEALLCODING) or \
            (('wxMac' in wx.PlatformInfo) and (event.GetId() == self.hideAllCodingButton.GetId()))):
            # ... clear the screen (without deleting the Coding Objects)
            self.FileClear(event)
        # Get the list of keys for the Coding Objects
        keys = self.obj.codingObjects.keys()
        # Sort the keys, since drawing order matters
        keys.sort()
        # For each Coding Object ...
        for key in keys:
            # ... grab a reference to the current Coding object
            obj = self.obj.codingObjects[key]
            # If this was called from the Menu ...
            if not (event is None):
                # ... and it was the Show All Coding menu item ...
                if ((event.GetId() == MENU_FILE_SHOWALLCODING) or \
                    (('wxMac' in wx.PlatformInfo) and (event.GetId() == self.showAllCodingButton.GetId()))):
                    # ... then set ALL objects to Visible
                    obj['visible'] = True
                # ... and it was the Hide All Coding menu item ...
                elif ((event.GetId() == MENU_FILE_HIDEALLCODING) or \
                      (('wxMac' in wx.PlatformInfo) and (event.GetId() == self.hideAllCodingButton.GetId()))):
                    # ... then set ALL object to NOT VISIBLE
                    obj['visible'] = False

            # If the current object is supposed to be VISIBLE ...
            if obj['visible']:
                # Initialize the Draw Object to None
                drawObj = None
                # Get the Keyword Styles for the selected Keyword Group : Keyword from the Snapshot
                keywordStyles = self.obj.keywordStyles[(obj['keywordGroup'], obj['keyword'])]
                # If the Line Color's Name is in the defined graphics colors ...
                if keywordStyles['lineColorName'] in TransanaGlobal.transana_colorLookup.keys():
                    # ... get the Color Definition by looking it up in the Global List.
                    #     (Thus, it will be correct even if the user changes the global color definitions!)
                    color = TransanaGlobal.transana_colorLookup[keywordStyles['lineColorName']]
                # If the Color Name isn't in the defined graphics colors ...
                else:
                    # ... use the Color Definition saved with the Snapshot's Keyword Styles
                    color = (int(keywordStyles['lineColorDef'][1:3], 16), int(keywordStyles['lineColorDef'][3:5], 16), int(keywordStyles['lineColorDef'][5:7], 16))
                # Set the Coding Color to the defined color definition
                codingColor = wx.Colour(color[0], color[1], color[2])
                # If we're drawing an Arrow ...
                if keywordStyles['drawMode'] == 'Arrow':
                    # Add the Arrow to the canvas
                    drawObj = self.canvas.AddArrowLine((obj['x1'], obj['y1'], obj['x2'], obj['y2']),
                                                       LineWidth = int(keywordStyles['lineWidth']),
                                                       LineColor = codingColor,
                                                       LineStyle = keywordStyles['lineStyle'],
                                                       ArrowHeadSize = 15,
                                                       ArrowHeadAngle = 60)

                # If we're drawing a Line ...
                elif keywordStyles['drawMode'] == 'Line':
                    # Add the Line to the canvas
                    drawObj = self.canvas.AddArrowLine((obj['x1'], obj['y1'], obj['x2'], obj['y2']),
                                                       LineWidth = int(keywordStyles['lineWidth']),
                                                       LineColor = codingColor,
                                                       LineStyle = keywordStyles['lineStyle'],
                                                       ArrowHeadSize = 0)

                # If we're drawing a Rectangle ...
                elif keywordStyles['drawMode'] == 'Rectangle':
                    # Add the Rectangle to the canvas
                    drawObj = self.canvas.AddRectangle((obj['x1'], obj['y1']), (obj['x2'] - obj['x1'], obj['y2'] - obj['y1']),
                                                       LineWidth = int(keywordStyles['lineWidth']),
                                                       LineColor = codingColor,
                                                       LineStyle = keywordStyles['lineStyle'])
                
                # If we're drawing an Ellipse ...
                elif keywordStyles['drawMode'] == 'Ellipse':
                    # Add the Ellipse to the canvas
                    drawObj = self.canvas.AddEllipse((obj['x1'], obj['y1']), (obj['x2'] - obj['x1'], obj['y2'] - obj['y1']),
                                                     LineWidth = int(keywordStyles['lineWidth']),
                                                     LineColor = codingColor,
                                                     LineStyle = keywordStyles['lineStyle'])

                # If we have a valid draw object ...
                if drawObj:
                    # Number the draw object
                    drawObj.ObjectNum = key
                    # Name the draw object
                    drawObj.Name = "%s : %s" % (obj['keywordGroup'], obj['keyword'])
                    # Specify the "Hit" size as slightly larger than the arrow line itself to make object selection easier
                    drawObj.HitLineWidth = self.lineWidth + 2
                    # Define the draw object's Left Click, Enter, and Leave events
                    drawObj.Bind(FloatCanvas.EVT_FC_RIGHT_DOWN, self.OnItemRightDown)
                    drawObj.Bind(FloatCanvas.EVT_FC_ENTER_OBJECT, self.OnEnterObject)
                    drawObj.Bind(FloatCanvas.EVT_FC_LEAVE_OBJECT, self.OnLeaveObject)

        # Hide All Codes was causing a zoomed-in image to shift.  This shifts it back!
        self.canvas.SendSizeEvent()
        # Redraw the FloatCanvas
        self.canvas.Draw()

    def CopyBitmap(self):
        """ Provide a copy of the cropped, coded Snapshot image as a wxBitmap.
               NOTE:  The Bitmap object MUST be destroyed by the routine that calls this method! """
        # Reset the Bounding Box to avoid NaN problems
        self.canvas._ResetBoundingBox()
        # Get the image's Bounding Box
        box = self.canvas.BoundingBox
        # Create an empty Bitmap the size of the image
        tmpBMP = wx.EmptyBitmap(self.canvas.GetSize()[0], self.canvas.GetSize()[1])
        # Get the MemoryDC for the empty bitmap
        tempDC = wx.MemoryDC(tmpBMP)
        # Set the bitmap's background colour to WHITE
        tempDC.SetBackground(wx.Brush("white"))
        tempDC.Clear()
        # Create a ClientDC
        tempDC2 = wx.ClientDC(self)
        # Create another MemoryDC (although I'm not sure why!)
        tempDC3 = wx.MemoryDC()
        # Get the image from the hidden Snapshot Window's Canvas and put it in the Device Contexts we just created
        self.canvas._DrawObjects(tempDC, self.canvas._DrawList, tempDC2, box, tempDC3)
        # Pass the bitmap to the calling routine, which needs to Destroy() it
        return tmpBMP

    def FileSaveSelectionAs(self, event, filename = None):
        """ Save the StillImageWindow's Visible Selection with markup """
        # If NO file name is passed in ...
        if filename == None:
            # Create a dialog to prompt the user for a file name and path
            dlg = wx.FileDialog(self, message="Save file as ...", defaultDir=TransanaGlobal.configData.videoPath, 
                                defaultFile="", wildcard="*.jpg", style=wx.SAVE )
            # Display the dialog.
            result = dlg.ShowModal()
            # If the user hits OK ...
            if result == wx.ID_OK:
                # ... get the path
                path = dlg.GetPath()
            # Destroy the File Dialog
            dlg.Destroy()
            # Check for the file extension
            if not(path[-4:].lower() == ".jpg"):
                # ... and add it if needed
                path = path+".jpg"
            # Check to see if the file already exists ...
            if os.path.exists(path):
                # ... and if so, build an error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A file named "%s" already exists.  Do you want to replace it?'), 'utf8')
                else:
                    prompt = _('A file named "%s" already exists.  Do you want to replace it?')
                # Build an error dialog
                dlg2 = Dialogs.QuestionDialog(None, prompt % path)
                # Get the user response
                if dlg2.LocalShowModal() == wx.ID_YES:
                    result = wx.ID_OK
                else:
                    result = wx.ID_CANCEL
                dlg2.Destroy()
        # If a file name is passed in ...
        else:
            # ... fake the User hitting OK
            result = wx.ID_OK
            # ... and use the file name passed in
            path = filename
            # Check for the file extension
            if not(path[-4:].lower() == ".jpg"):
                # ... and add it if needed
                path = path+".jpg"

        # If the user pressed OK
        if result == wx.ID_OK:
            
            # Use the FloatCanvas' SaveAsImage() method to export the image
            self.canvas.SaveAsImage(path, ImageType=wx.BITMAP_TYPE_JPEG)

            # Initialize the Coding Key Bitmap to None
            codingKeyBMP = None
            # Start Exception Handling
            try:
                # Get the Coding Key Bitmap from the Coding Key Popup, if it's open.
                # (It will throw an exception if the Coding Key Popup is not open, ending this process.)
                codingKeyBMP = self.codingKeyPopup.theBitmap
                # Load the Snapshot Bitmap we just saved
                snapshotBMP = wx.Bitmap(path, wx.BITMAP_TYPE_JPEG)
                # Convert the Coding Key Bitmap to an Image
                codingKeyImage = codingKeyBMP.ConvertToImage()
                # Determine the factor for scaling the Coding Key relative to the image.
                # (Snapshot width divided by twice the Coding Key width makes the Coding Key 1/2 the width of the Snapshot)
                scalingFactor = float(snapshotBMP.GetWidth()) / (2.0 * float(codingKeyImage.GetWidth()))
                # For legibility, the Scaling Factor should be no smaller than 0.5, no larger than 5.
                scalingFactor = max(scalingFactor, 0.5)
                scalingFactor = min(scalingFactor, 5.0)
                # Scale the Coding Key Image, converting it to a Bitmap in the process
                codingKeyBMP = wx.BitmapFromImage(codingKeyImage.Scale(scalingFactor * codingKeyImage.GetWidth(), scalingFactor * codingKeyImage.GetHeight(), wx.IMAGE_QUALITY_HIGH))
                # Create an empty bitmap big enough to hold BOTH images
                tmpBMP = wx.EmptyBitmap(max(snapshotBMP.GetSize()[0], codingKeyBMP.GetSize()[0]), snapshotBMP.GetSize()[1] + codingKeyBMP.GetSize()[1] + 4)
                # Get the Device Context for the empty bitmap
                dc = wx.BufferedDC(None, tmpBMP)

                # Begin the Drawing Process
                dc.BeginDrawing()
                # Give the DC a white background
                dc.SetBackground(wx.Brush(wx.Colour(255, 255, 255)))
                # Clear the DC
                dc.Clear()
                # Get a Device Context for the Snapshot
                dc1 = wx.BufferedDC(None, snapshotBMP)
                # Put the Snapshot's DC onto the Empty Image's DC in the top left corner
                dc.Blit(0, 0, snapshotBMP.GetWidth(), snapshotBMP.GetHeight(), dc1, 0, 0)
                # Get a Device Context for the Coding Key
                dc2 = wx.BufferedDC(None, codingKeyBMP)
                # Draw a box around the coding key
                dc2.DrawLine(0, 0, codingKeyBMP.GetWidth()-1, 0)
                dc2.DrawLine(codingKeyBMP.GetWidth()-1, 0, codingKeyBMP.GetWidth()-1, codingKeyBMP.GetHeight()-23)
                dc2.DrawLine(codingKeyBMP.GetWidth()-1, codingKeyBMP.GetHeight()-23, 0, codingKeyBMP.GetHeight()-23)
                dc2.DrawLine(0, codingKeyBMP.GetHeight()-23, 0, 0)
                # Put the Coding Key on the Empty Image's DC, centered horizontally, at the bottom vertically
                dc.Blit(int((tmpBMP.GetWidth() / 2.0) - (codingKeyBMP.GetWidth() / 2.0)), snapshotBMP.GetHeight() + 3, codingKeyBMP.GetWidth(), codingKeyBMP.GetHeight(), dc2, 0, 0)
                # End the Drawing Process
                dc.EndDrawing()

                # Save the formerly empty Bitmap, whose DC we just filled in
                tmpBMP.SaveFile(path, wx.BITMAP_TYPE_JPEG)

            # Handle Exceptions.
            # An exception here indicates the Coding Key Popup was not open, so we don't need to add the Coding Key to the exported image
            except:
                if DEBUG:
                    print "Failed to get Coding Key Bitmap"
                    import sys
                    print sys.exc_info()[0], sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    print

    def FileSaveAs(self, event):
        """ Save the StillImageWindow's entire image with markup """
        # Create a dialog to prompt the user for a file name and path
        dlg = wx.FileDialog(self, message="Save file as ...", defaultDir=TransanaGlobal.configData.videoPath, 
                            defaultFile="", wildcard="*.jpg", style=wx.SAVE )
        # Display the file dialog.
        result = dlg.ShowModal()
        # Destroy the File Dialog
        dlg.Destroy()
        # If the user hits OK ...
        if result == wx.ID_OK:
            # ... get the file path
            path = dlg.GetPath()
            # Check for the file extension
            if not(path[-4:].lower() == ".jpg"):
                # ... and add it if needed
                path = path+".jpg"

            # Check to see if the file already exists ...
            if os.path.exists(path):
                # ... and if so, build an error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A file named "%s" already exists.  Do you want to replace it?'), 'utf8')
                else:
                    prompt = _('A file named "%s" already exists.  Do you want to replace it?')
                # Build an error dialog
                dlg2 = Dialogs.QuestionDialog(None, prompt % path)
                # Get the user response
                result2 = dlg2.LocalShowModal()
                dlg2.Destroy()
                if result2 != wx.ID_YES:
                    # Exit the Save
                    return
            
            # Freeze the image
            self.canvas.Freeze()

            # Remember the original Scale and Center of the current image
            oldScale = self.canvas.Scale
            oldCenter = self.canvas.ViewPortCenter

            # Reset the image Scale to 1, so we get the full size image
            self.canvas.Scale= 1
            # Set the Bounding Box so we get the whole image
            self.canvas._ResetBoundingBox()
            box = self.canvas.BoundingBox

            # Get the image's size
            size = self.canvas.GetSize()
            # Center the full image
            self.canvas.ViewPortCenter = box[0][0] + (size[0] / 2), box[1][1] - (size[1] / 2)
            # Set the Scale, redrawing the image
            self.canvas.SetToNewScale(DrawFlag=True)
            
            # Create an empty bitmap
            snapshotBMP = wx.EmptyBitmap(box[1][0] - box[0][0], box[1][1] - box[0][1])
            # Copy the FloatCanvas image to the empty bitmap
            tempDC = wx.MemoryDC(snapshotBMP)
            tempDC2 = wx.ClientDC(self)
            tempDC3 = wx.MemoryDC()

            # Begin drawing on the DC
            tempDC.BeginDrawing()
            # Get a Device Context for the original image.  The ScaledBitmap2 object won't give the full image!
            dc1 = wx.BufferedDC(None, self.bgImage.ConvertToBitmap())
            # Put the original image's DC onto the Empty Image's DC in the top left corner
            tempDC.Blit(0, 0, self.bgImage.GetWidth(), self.bgImage.GetHeight(), dc1, 0, 0)
            # End drawing on the DC
            tempDC.EndDrawing()

            # Now put the FloatCanvas' images on the DC.  The partial image from the ScaledBitmap2 lines up with the DC blit we just did!
            self.canvas._DrawObjects(tempDC, self.canvas._DrawList, tempDC2, box, tempDC3)

            # Initialize the Coding Key Bitmap to None
            codingKeyBMP = None
            # Start Exception Handling
            try:
                # Get the Coding Key Bitmap from the Coding Key Popup, if it's open.
                # (It will throw an exception if the Coding Key Popup is not open, ending this process.)
                codingKeyBMP = self.codingKeyPopup.theBitmap
                # Convert the Coding Key Bitmap to an Image
                codingKeyImage = codingKeyBMP.ConvertToImage()
                # Determine the factor for scaling the Coding Key relative to the image.
                # (Snapshot width divided by twice the Coding Key width makes the Coding Key 1/2 the width of the Snapshot)
                scalingFactor = float(snapshotBMP.GetWidth()) / (2.0 * float(codingKeyImage.GetWidth()))
                # For legibility, the Scaling Factor should be no smaller than 0.5, no larger than 5.
                scalingFactor = max(scalingFactor, 0.5)
                scalingFactor = min(scalingFactor, 5.0)
                # Scale the Coding Key Image, converting it to a Bitmap in the process
                codingKeyBMP = wx.BitmapFromImage(codingKeyImage.Scale(scalingFactor * codingKeyImage.GetWidth(), scalingFactor * codingKeyImage.GetHeight(), wx.IMAGE_QUALITY_HIGH))
            # Handle Exceptions.
            # An exception here indicates the Coding Key Popup was not open, so we don't need to add the Coding Key to the exported image
            except:
                if DEBUG:
                    print "Failed to get Coding Key Bitmap"
                    import sys
                    print sys.exc_info()[0], sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    print

            # If we didn't get a Coding Bitmap ...
            if codingKeyBMP == None:
                # ... just use the Snapshot Bitmap
                tmpBMP = snapshotBMP
            # If we got a Coding Bitmap, combine it with the Snapshot Bitmap
            else:
                # Create an empty bitmap big enough to hold BOTH images
                tmpBMP = wx.EmptyBitmap(max(snapshotBMP.GetSize()[0], codingKeyBMP.GetSize()[0]), snapshotBMP.GetSize()[1] + codingKeyBMP.GetSize()[1] + 4)
                # Get the Device Context for the empty bitmap
                dc = wx.BufferedDC(None, tmpBMP)

                # Begin the Drawing Process
                dc.BeginDrawing()
                # Give the Bitmap a white background
                dc.SetBackground(wx.Brush(wx.Colour(255, 255, 255)))
                # Clear the DC
                dc.Clear()
                # Get a Device Context for the Snapshot
                dc1 = wx.BufferedDC(None, snapshotBMP)
                # Put the Snapshot's DC onto the Empty Image's DC in the top left corner
                dc.Blit(0, 0, snapshotBMP.GetWidth(), snapshotBMP.GetHeight(), dc1, 0, 0)
                # Get a Device Context for the Coding Key
                dc2 = wx.BufferedDC(None, codingKeyBMP)
                # Draw a box around the coding key
                dc2.DrawLine(0, 0, codingKeyBMP.GetWidth()-1, 0)
                dc2.DrawLine(codingKeyBMP.GetWidth()-1, 0, codingKeyBMP.GetWidth()-1, codingKeyBMP.GetHeight()-23)
                dc2.DrawLine(codingKeyBMP.GetWidth()-1, codingKeyBMP.GetHeight()-23, 0, codingKeyBMP.GetHeight()-23)
                dc2.DrawLine(0, codingKeyBMP.GetHeight()-23, 0, 0)
                # Put the Coding Key on the Empty Image's DC, centered horizontally, at the bottom vertically
                dc.Blit(int((tmpBMP.GetWidth() / 2.0) - (codingKeyBMP.GetWidth() / 2.0)), snapshotBMP.GetHeight() + 3, codingKeyBMP.GetWidth(), codingKeyBMP.GetHeight(), dc2, 0, 0)
                # End the Drawing Process
                dc.EndDrawing()

            # Save the formerly empty Bitmap, whose DC we just filled in
            tmpBMP.SaveFile(path, wx.BITMAP_TYPE_JPEG)

            # Reset the image's Scale and Center to their original values
            self.canvas.Scale = oldScale
            self.canvas.SetToNewScale(DrawFlag=True)
            self.canvas.ViewPortCenter = oldCenter
            # Redraw the image so changes which shouldn't show up don't.
            self.FileRedraw(None)

            # Thaw the image
            self.canvas.Thaw()

    def CloseAllImages(self, event):
        """ Close All Images menu handler """
        # Call the Control Object's Close All Images function
        self.ControlObject.CloseAllImages()

    def LeaveEditMode(self):
        """ Leave Edit Mode (i.e. Save the Snapshot Changes) """
        try:
            # Note the Scale
            self.obj.image_scale = self.canvas.Scale
            # Note the Image Position
            self.obj.image_coords = (self.canvas.ViewPortCenter[0], self.canvas.ViewPortCenter[1])
            # Note the Image Size
            self.obj.image_size = (self.GetSize()[0], self.GetSize()[1])
            # Save the Snapshot
            self.obj.db_save()
            # Unlock the Snapshot Object
            self.obj.unlock_record()
            # Reload the record to update the Last Save Time value
            self.obj.db_load(self.obj.number)
            # Update the Keyword Visualization, if needed
            self.ControlObject.UpdateKeywordVisualization()
            # Multi-user Messaging
            if not TransanaConstants.singleUserVersion:
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if (self.obj.episode_num != 0):
                    # We need to update the Episode Keyword Visualization
                    if DEBUG:
                        print 'Message to send = "UKV %s %s %s"' % ('Episode', self.obj.episode_num, 0)
                        
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', self.obj.episode_num, 0))
                # If this Snapshot is open on other computers, it should be updated.
                if TransanaGlobal.chatWindow != None:
                    TransanaGlobal.chatWindow.SendMessage("US %s" % (self.obj.number))
            # Toggle the Toolbar Button to signal that we have left Edit mode
            self.toolbar.ToggleTool(self.editTool.GetId(), False)
        except TransanaExceptions.SaveError:
            # Display the Error Message, allow "continue" flag to remain true
            errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
            errordlg.ShowModal()
            errordlg.Destroy()

            # Reject the attempt to leave Edit mode
            self.toolbar.ToggleTool(self.editTool.GetId(), True)

    def OnClose(self, event):
        """ The method for the wx.EVT_CLOSE event.  Handles closing the Snapshot Window cleanly. """
        # We need to signal that we are closing this window, so the wx.EVT_ACTIVATE handler doesn't panic.
        self.closing = True
        # if we're in Edit Mode ...
        if self.showWindow and self.editTool.IsToggled():
            # ... leave Edit Mode
            self.LeaveEditMode()

            # If LeaveEditMode did NOT leave Edit Mode ...
            if self.editTool.IsToggled():
                # ... there is an error condition and we need to VETO the close!
                event.Veto()
                # Reset the closing flag.  We're not closing after all.
                self.closing = False
                # We need to exit this method now for the Veto to occur properly.
                return

        # If the Coding Key is visible ...
        if self.codingKeyPopup != None:
            # Start Exception handling
            try:
                # ... close the Coding Key
                self.codingKeyPopup.Close()
            # Ignore exceptions
            except:
                pass
        # If this window was not a HIDDEN Snapshot ...
        if self.showWindow:
            # Remove ControlObject.SnapshotWindows Reference!
            self.ControlObject.RemoveSnapshotWindow(self.obj.id, self.obj.number)

        self.bgImage.Destroy()
        self.bgImage = None

        # Allow the inherited Close method to be called
        event.Skip()

    def OnEnterWindow(self, event):
        """ Handle the wx.EVT_ACTIVATE event, when the window is selected """
        try:
            # If we're not closing this window, we need to do these things.  Otherwise, skip it.
            if not self.closing:
                # Remember the currently selected Keyword Group
                kwg = self.keyword_group_cb.GetStringSelection()
                # Remember the currently selected Keyword
                kw = self.keyword_cb.GetStringSelection()
                # Get a list of all Keyword Groups
                allKwgs = DBInterface.list_of_keyword_groups()
                # Replace the items in the Keyword Group Choice Box, in case new Keyword Group has been added
                self.keyword_group_cb.SetItems([''] + allKwgs)
                # If a Keyword Group was selected ...
                if (kwg != '') and (kwg in allKwgs):
                    # ... re-select that Keyword Group
                    self.keyword_group_cb.SetStringSelection(kwg)
                    # Get a list of all Keywords for the selected Keyword Group
                    allKws = DBInterface.list_of_keywords_by_group(kwg)
                    # Replace the items in the Keywords Choice box, in case new Keywords have been added
                    self.keyword_cb.SetItems([''] + allKws)
                    # If a Keyword is selected ...
                    if (kw != '') and (kw in allKws):
                        # ... re-select that Keyword
                        self.keyword_cb.SetStringSelection(kw)
                # If we have NO Keyword Group ...
                else:
                    # ... clear the Keyword List too!
                    self.keyword_cb.SetItems([''])

        except:

            import sys
            print "SnapshotWindow.OnEnterWindow():", sys.exc_info()[0]
            print sys.exc_info()[1]
            import traceback
            traceback.print_exc(file=sys.stdout)
            print

            pass

    def OnLeaveWindow(self, event):

        print "SnapshotWindow.OnLeaveWindow()"

    def CloseWindow(self, event):
        """ Handle the wx.EVT_MENU event for the Snapshot Window's File > Close Menu. """
        # Close the Window.  (This triggers self.OnClose().)
        self.Close()

    def __size(self):
        """Determine the default size for the Snapshot frame."""
        # Get the size of the first monitor
        rect = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()

        # If we're on Linux, we may not be getting the right screen size value
        if 'wxGTK' in wx.PlatformInfo:
            # rect2 = wx.Display(TransanaGlobal.configData.primaryScreen).GetGeometry()
            width = (rect[2] - rect[0] - 4) * .6

            # Snapshot Compontent should be 80% of the HEIGHT, adjusted for the menu height
            height = (min(rect[3], rect[3]) - rect[1] - 6) * .80

        # If we're on Windows or OS X ...
        else:
            # Snapshot Component should be 60% of the WIDTH
            width = rect[2] * .6

            # Snapshot Compontent should be 80% of the HEIGHT, adjusted for the menu height
            height = (rect[3] - TransanaGlobal.menuHeight) * .80

        # Return the SIZE values
        return wx.Size(width, height)

    def __pos(self):
        """Determine default position of Snapshot Frame."""
        # Get the size of the first monitor
        rect = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
        # rect[0] compensates if Start menu is on Left
        x = rect[0] + 20
        # rect[1] compensates if Start menu is on Top
        if 'wxGTK' in wx.PlatformInfo:
            y = min(rect[1] + TransanaGlobal.menuHeight + 20 + (20 * len(self.ControlObject.SnapshotWindows)), rect[3] - 80)
        else:
            y = min(rect[1] + TransanaGlobal.menuHeight + 20 + (20 * len(self.ControlObject.SnapshotWindows)), rect[3] - 80)
        # Return the POSITION values
        return (x, y)    



# import the numpy module
import numpy

class GUITransana(GUIMode.GUIBase):
    """ This is a custom GUI Mode for Transana, one which shows a Rubber Band box when drawing.
        It is a combination of GUIMode.GUIMouse (for HitTest functionality) and GUIMode.GUIZoomIn
        (for Rubber Band functionality). """
 
    def __init__(self, canvas=None):
        """ Initialize the custom GUIMode """
        # Initialize, inherited from GUIMode.GUIBase
        GUIMode.GUIBase.__init__(self, canvas)
        # Initialize the Starting and Previous Rubber Band Boxes to None
        self.StartRBBox = None
        self.PrevRBBox = None

    def OnLeftDown(self, event):
        """ Detect Left Mouse Down to initate Rubber Band box and Hit Test functions """
        # Call the Canvas' Left Down event
        self.Canvas._RaiseMouseEvent(event, FloatCanvas.EVT_FC_LEFT_DOWN)
        # Rubber Band Box
        self.StartRBBox = numpy.array( event.GetPosition() )
        self.PrevRBBox = None
        self.Canvas.CaptureMouse()
        # Make Selection Hit-Test-able
        EventType = FloatCanvas.EVT_FC_LEFT_DOWN
        if not self.Canvas.HitTest(event, EventType):
            self.Canvas._RaiseMouseEvent(event, EventType)

    def OnLeftUp(self, event):
        """ Detect Left Mouse Up to finish Rubber Band Box and Hit Test functions """
        # Rubber Band Box
        if event.LeftUp() and not self.StartRBBox is None:
            self.PrevRBBox = None
            self.StartRBBox = None
        # Make Selection Hit-Test-able
        EventType = FloatCanvas.EVT_FC_LEFT_UP
        if not self.Canvas.HitTest(event, EventType):
            self.Canvas._RaiseMouseEvent(event, EventType)
        # Call the Canvas' Left Up event
        self.Canvas._RaiseMouseEvent(event, FloatCanvas.EVT_FC_LEFT_UP)

    def OnRightDown(self, event):
        """ Detect Right Mouse Down to initate Rubber Band box and Hit Test functions """
        # Call the Canvas' Right Down event
        self.Canvas._RaiseMouseEvent(event, FloatCanvas.EVT_FC_RIGHT_DOWN)
        # Make Selection Hit-Test-able
        EventType = FloatCanvas.EVT_FC_RIGHT_DOWN
        if not self.Canvas.HitTest(event, EventType):
            self.Canvas._RaiseMouseEvent(event, EventType)

    def OnRightUp(self, event):
        """ Detect Right Mouse Up to finish Rubber Band Box and Hit Test functions """
        # Make Selection Hit-Test-able
        EventType = FloatCanvas.EVT_FC_RIGHT_UP
        if not self.Canvas.HitTest(event, EventType):
            self.Canvas._RaiseMouseEvent(event, EventType)
        # Call the Canvas' Right Up event
        self.Canvas._RaiseMouseEvent(event, FloatCanvas.EVT_FC_RIGHT_UP)

    def OnMove(self, event):
        """ Detect Mouse Motion for Rubber Band Box and Hit Test functions """
        # Hit Test
        # self.Canvas.MouseOverTest(event)

        # Always raise the Move event.
        self.Canvas._RaiseMouseEvent(event,FloatCanvas.EVT_FC_MOTION)
        # Rubber Band Box
        if event.Dragging() and event.LeftIsDown() and not (self.StartRBBox is None):
            xy0 = self.StartRBBox
            xy1 = numpy.array( event.GetPosition() )
            wh  = abs(xy1 - xy0)
            xy_c = (xy0 + xy1) / 2
            dc = wx.ClientDC(self.Canvas)
            dc.BeginDrawing()
            dc.SetPen(wx.Pen('WHITE', 2, wx.SHORT_DASH))
            dc.SetBrush(wx.TRANSPARENT_BRUSH)
            dc.SetLogicalFunction(wx.XOR)
            if self.PrevRBBox:
                dc.DrawRectanglePointSize(*self.PrevRBBox)
            self.PrevRBBox = ( xy_c - wh/2, wh )
            dc.DrawRectanglePointSize( *self.PrevRBBox )
            dc.EndDrawing()
            
    def UpdateScreen(self):
        """
        Update gets called if the screen has been repainted in the middle of a zoom in
        so the Rubber Band Box can get updated
        """
        if self.PrevRBBox is not None:
            dc = wx.ClientDC(self.Canvas)
            dc.SetPen(wx.Pen('WHITE', 2, wx.SHORT_DASH))
            dc.SetBrush(wx.TRANSPARENT_BRUSH)
            dc.SetLogicalFunction(wx.XOR)
            dc.DrawRectanglePointSize(*self.PrevRBBox)


def CodingKeyGraphic(keywordStyle):
    """ Given a keywordStyle, return an Image of that style, used for the Coding Key in several places. """
    # Create an illustration of the coding key!
    tmpBMP = wx.EmptyBitmap(72, 17)
    # Create a Device Context to draw on
    dc = wx.BufferedDC(None, tmpBMP)
    # Begin the Drawing Process
    dc.BeginDrawing()
    # If the lineColor is NOT White ...
    if keywordStyle['lineColorDef'] != u'#ffffff':
        # Give the Bitmap a white background
        dc.SetBackground(wx.Brush(wx.Colour(255, 255, 255)))
    else:
        # Give the Bitmap a black background
        dc.SetBackground(wx.Brush(wx.Colour(0, 0, 0)))
        # Create a Brush.  This fills the shape with black as well
        brush = wx.Brush(wx.Colour(0, 0, 0), style=wx.SOLID)
        # Set the Brush for the Device Context
        dc.SetBrush(brush)
    # Clear the DC
    dc.Clear()
    # Set the Drawing Line Style
    if keywordStyle['lineStyle'] == 'Solid':
        penStyle = wx.SOLID
    elif keywordStyle['lineStyle'] == 'LongDash':
        penStyle = wx.LONG_DASH
    elif keywordStyle['lineStyle'] == 'Dot':
        penStyle = wx.DOT
    elif keywordStyle['lineStyle'] == 'DotDash':
        penStyle = wx.DOT_DASH
    # Build the Color
    colorRed = int(keywordStyle['lineColorDef'][-6:-4], 16)
    colorGreen = int(keywordStyle['lineColorDef'][-4:-2], 16)
    colorBlue = int(keywordStyle['lineColorDef'][-2:], 16)
    # Define the Pen for drawing
    pen = wx.Pen(wx.Colour(colorRed, colorGreen, colorBlue), int(keywordStyle['lineWidth']), penStyle)
    # Set the Pen for the Device Context
    dc.SetPen(pen)
    # Draw the Coding Style Glyph Bitmap
    if keywordStyle['drawMode'] == 'Line':
        # Draw the sample Line with the Pen we just created
        dc.DrawLine(6, 8, 63, 8)
    elif keywordStyle['drawMode'] == 'Arrow':
        # Draw the sample Arrow with the Pen we just created
        dc.DrawLine(6, 8, 63, 8)
        dc.DrawLine(56, 1, 63, 8)
        dc.DrawLine(56, 17, 63, 8)
    elif keywordStyle['drawMode'] == 'Rectangle':
        # Draw the sample Rectangle with the Pen we just created
        dc.DrawRectangle(6, 4, 62, 11)
    elif keywordStyle['drawMode'] == 'Ellipse':
        # Draw the sample Ellipse with the Pen we just created
        dc.DrawEllipse(6, 4, 60, 11)
    # FilledRectangle is created by the Keyword Summary Report to display Keywords that have a
    # defined Color but no defined Coding Shape, etc.
    elif keywordStyle['drawMode'] == 'FilledRectangle':
        # Create a Brush.  This allows us to have a SOLID rather than an OUTLINE shape
        brush = wx.Brush(wx.Colour(colorRed, colorGreen, colorBlue), style=wx.SOLID)
        # Set the Brush for the Device Context
        dc.SetBrush(brush)
        # Draw the sample Rectangle with the Pen we just created
        dc.DrawRectangle(6, 4, 28, 11)
    # End the drawing
    dc.EndDrawing()
    # Select a different object into the Device Context, which allows the Bitmap to be used
    dc.SelectObject(wx.EmptyBitmap(5, 5))
    # Convert the Bitmap to an Image
    tmpImage = tmpBMP.ConvertToImage()
    # Return the image to the calling routine
    return tmpImage

class CodingKeyWindow(wx.Frame):
    """ Create the Coding Key Window and Graphic for the Snapshot Window """
    def __init__(self, parent, keywordStyles, pos):
        """ Initiliaze the Coding Key Window """
        # Remember the Keyword Styles
        self.keywordStyles = keywordStyles
        # Determine the default Coding Key Image size
        self.imageSize = (250, 22 * (len(keywordStyles) + 1))
        # Create the Coding Key Frame
        wx.Frame.__init__(self, parent, -1,
                          _("Snapshot Coding Key"),
                          pos,
                          wx.Size(self.imageSize[0], self.imageSize[1] + 50),
                          style=wx.CAPTION | wx.CLOSE_BOX | wx.RESIZE_BORDER | wx.FRAME_FLOAT_ON_PARENT)
        # Configure for a Custom Background
        self.SetBackgroundStyle(wx.BG_STYLE_CUSTOM)
        # Define the Window Events needed to draw the background on the Window
        self.Bind(wx.EVT_SIZE, self.OnSize)
        self.Bind(wx.EVT_PAINT, self.OnPaint)
        self.Bind(wx.EVT_ERASE_BACKGROUND, self.OnEraseBackground)

    def CreateBitmap(self):
        """ Create the Coding Key graphic """
        # Create and empty Bitmap the size of the Window
        self.theBitmap = wx.EmptyBitmap(self.imageSize[0], self.imageSize[1])
        # Create a Device Context to draw on
        dc = wx.BufferedDC(None, self.theBitmap)
        # Begin the Drawing Process
        dc.BeginDrawing()
        # Give the Bitmap a white background
        dc.SetBackground(wx.Brush(wx.Colour(255, 255, 255)))
        # Clear the DC
        dc.Clear()
        # Set the Font for the Title (14 point, bold, underlined)
        theFont = wx.Font(pointSize=14, family=wx.FONTFAMILY_SWISS, style=wx.FONTSTYLE_NORMAL, weight=wx.FONTWEIGHT_BOLD, underline=True)
        dc.SetFont(theFont)
        # Determine the size of the Title text
        titleSize = dc.GetTextExtent(_("Snapshot Coding Key"))
        # Determine the Window Size
        windowSize = self.GetSizeTuple()
        # Draw the title, centered in the Window
        dc.DrawText(_("Snapshot Coding Key"), windowSize[0]/2 - titleSize[0]/2, 3)
        # Set the Font for the Coding Key labels (12 point plain text)
        theFont = wx.Font(pointSize=12, family=wx.FONTFAMILY_SWISS, style=wx.FONTSTYLE_NORMAL, weight=wx.FONTWEIGHT_NORMAL, underline=False)
        dc.SetFont(theFont)

        # Create a Counter for positioning.  Start at 1 to account for the title
        counter = 1
        # Get the keys for the Keyword Styles
        keys = self.keywordStyles.keys()
        # Sort the keys
        keys.sort()
        # For each key in the sorted keys ...
        for key in keys:
            # ... get the Style graphic as a Bitmap
            bmp = CodingKeyGraphic(self.keywordStyles[key]).ConvertToBitmap()
            # ... draw the Style graphic to the Coding Key image
            dc.DrawBitmap(bmp, 5, counter * 22 + 9)
            # ... add the text for the keyword
            dc.DrawText("%s : %s" % key, bmp.GetSize()[0] + 15, counter * 22 + 9)
            # Determine the size of the prompt text
            promptSize = dc.GetTextExtent("%s : %s" % key)
            # If the prompt is wider than the current window ...
            if (promptSize[0] + bmp.GetSize()[0] + 40) > self.GetSizeTuple()[0]:
                # ... widen the window to accomodate
                self.SetSize((promptSize[0] + bmp.GetSize()[0] + 40, self.GetSizeTuple()[1]))
            # Delete the image
            del(bmp)
            # ... increment the counter
            counter += 1
                                   
        # End the drawing
        dc.EndDrawing()
        
    def OnSize(self, event):
        """ Resize Event for the Coding Key Window """
        # Determine the Window Size
        self.imageSize = self.GetSizeTuple()
        # Re-draw the background image
        self.CreateBitmap()
        # Refresh the control
        self.Refresh()

    def OnEraseBackground(self, event):
        """ This method intentionally left blank! """
        pass

    def OnPaint(self, event):
        """ Paint Event for the Coding Key Window """
        # Get a Paint DC
        dc = wx.BufferedPaintDC(self)
        # Draw the Background Image on the Paint DC
        self.Draw(dc)

    def Draw(self, theDC):
        """ Draw Background for the Coding Key Window """
        # Place the Background Image on the DC passed in!
        if hasattr(self, 'theBitmap') and (self.theBitmap != None):
            theDC.DrawBitmap(self.theBitmap, 0, 0)
        

if __name__ == '__main__':

    class MyApp(wx.App):
        def OnInit(self):
            
            # Select ONE fileName for testing purposes
            # fileName='C:\\Users\\DavidWoods\\Videos\\Images\\NateOnRattlesnake.jpg'
            fileName='C:\\Users\\DavidWoods\\Videos\\Images\\IMG_0912B Taurmina Theater with boat.jpg'
            # fileName='C:\\Users\\DavidWoods\\Videos\\Images\\MomDad2003.jpg'
            # fileName = "C:\\Users\\DavidWoods\\Documents\\Transana Documentation\\Presentations\\20120919 - Ian Baird - Geography\\Photo_small.jpg"
            frame = SnapshotWindow(None, -1, "Still Image Window", fileName)
            self.SetTopWindow(frame)
            return True
          
    app = MyApp(0)
    app.MainLoop()
