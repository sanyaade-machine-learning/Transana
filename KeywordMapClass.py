#Copyright (C) 2002-2007  The Board of Regents of the University of Wisconsin System

#This program is free software; you can redistribute it and/or
#modify it under the terms of the GNU General Public License
#as published by the Free Software Foundation; either version 2
#of the License, or (at your option) any later version.

#This program is distributed in the hope that it will be useful,
#but WITHOUT ANY WARRANTY; without even the implied warranty of
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#GNU General Public License for more details.

#You should have received a copy of the GNU General Public License
#along with this program; if not, write to the Free Software
#Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

"""This module implements a map of keywords that have been applied to the selected Episode"""

__author__ = "David K. Woods <dwoods@wcer.wisc.edu>"

DEBUG = False
if DEBUG:
    print "KeywordMapClass DEBUG is ON!!"

# import Python's os and sys modules
import os, sys
# load wxPython for GUI
import wx
# import MySQLdb for database access
import MySQLdb
# load the GraphicsControl
import GraphicsControlClass
# Load the Printout Class
from KeywordMapPrintoutClass import MyPrintout
# Import Transana's Database Interface
import DBInterface
# Import Transana's Dialogs
import Dialogs
# Import Transana's Filter Dialog
import FilterDialog
# Import Transana's Constants
import TransanaConstants
# import Transana's Globals
import TransanaGlobal
# import Transana Miscellaneous functions
import Misc

# Declare Control IDs
# Menu Item and Toolbar Item for File > Filter
M_FILE_FILTER        =  wx.NewId()
T_FILE_FILTER        =  wx.NewId()
# Menu Item and Toolbar Item for File > Save As
M_FILE_SAVEAS        =  wx.NewId()
T_FILE_SAVEAS        =  wx.NewId()
# Menu Item and Toolbar Item for File > Printer Setup
M_FILE_PRINTSETUP    =  wx.NewId()
T_FILE_PRINTSETUP    =  wx.NewId()
# Menu Item and Toolbar Item for File > Print Preview
M_FILE_PRINTPREVIEW  =  wx.NewId()
T_FILE_PRINTPREVIEW  =  wx.NewId()
# Menu Item and Toolbar Item for File > Print 
M_FILE_PRINT         =  wx.NewId()
T_FILE_PRINT         =  wx.NewId()
# Menu Item and Toolbar Item for File > Exit
M_FILE_EXIT          =  wx.NewId()
T_FILE_EXIT          =  wx.NewId()
# Menu Item and Toolbar Item for Help > Help
M_HELP_HELP          =  wx.NewId()
T_HELP_HELP          =  wx.NewId()
# Series List Combo Box
ID_SERIESLIST        = wx.NewId()
# Episode List Combo Box
ID_EPISODELIST       = wx.NewId()

class KeywordMap(wx.Frame):
    """ This is the main class for the Keyword Map application.
        It can be instantiated as a free-standing report with a frame, or can be
        called as an embedded graphic display for the Visualization window. """
    def __init__(self, parent, ID=-1, title="", embedded=False, topOffset=0):
        # It's always important to remember your ancestors.
        self.parent = parent
        # We do some things differently if we're a free-standing Keyword Map report 
        # or if we're an embedded Keyword Visualization.  
        self.embedded = embedded
        # Remember the topOffset parameter value.  This is used to specify a larger top margin for the keyword visualization,
        # needed for the Hybrid Visualization.
        self.topOffset = topOffset
        #  Create a connection to the database
        DBConn = DBInterface.get_db()
        #  Create a cursor and execute the appropriate query
        self.DBCursor = DBConn.cursor()
        # If we're NOT embedded, we need to create a full frame etc.
        if not self.embedded:
            # Determine the screen size for setting the initial dialog size
            rect = wx.ClientDisplayRect()
            width = rect[2] * .80
            height = rect[3] * .80
            # Create the basic Frame structure with a white background
            self.frame = wx.Frame.__init__(self, parent, ID, title, pos=(10, 10), size=wx.Size(width, height), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL | wx.NO_FULL_REPAINT_ON_RESIZE)
            self.SetBackgroundColour(wx.WHITE)
            # Set the icon
            transanaIcon = wx.Icon(os.path.join(TransanaGlobal.programDir, "images", "Transana.ico"), wx.BITMAP_TYPE_ICO)
            self.SetIcon(transanaIcon)
            # The rest of the form creation is deferred to the Setup method.  To set the form up,
            # we need data from the database, which we cannot get until we have gotten the username
            # and password information.  However, to get that information, we need a Main Form to
            # exist.  Be careful not to confuse Setup() with SetupEmbedded().
        # If we ARE embedded, we just use the parent object's graphics canvas for drawing the Keyword Map.
        else:
            # We need to unbind existing mouse motion event connections to avoid overlapping ToolTips and other problems.
            self.parent.waveform.Unbind(wx.EVT_MOTION)
            # point the local graphic object to the appropriate item in the parent, the Visualization Window's waveform object.
            self.graphic = self.parent.waveform
            # We need to assign the Keyword Map's Mouse Motion event to the existing graphic object.
            self.graphic.Bind(wx.EVT_MOTION, self.OnMouseMotion)
            # The rest of the information is deferred to the SetupEmbedded method.  Be careful not to confuse
            # SetupEmbedded() with Setup().
        # Initialize EpisodeNum
        self.episodeNum = None
        # Initialize Media File to nothing
        self.MediaFile = ''
        # Initialize Media Length to 0
        self.MediaLength = 0
        # Initialize Keyword Lists to empty
        self.unfilteredKeywordList = []
        self.filteredKeywordList = []
        self.clipList = []
        self.clipFilterList = []
        # To be able to show only parts of an Episode Time Line, we need variables for the time boundaries.
        self.startTime = 0
        self.endTime = 0
        self.keywordClipList = {}
        self.configName = ''
        # Initialize variables required to avoid crashes when the visualization has been cleared
        self.graphicindent = 0
        self.Bounds = [1, 1, 1, 1]
        # Create a dictionary of the colors for each keyword.
        self.keywordColors = {'lastColor' : -1}
        if not self.embedded:
            # Get the Configuration values for the Keyword Map Options
            self.barHeight = TransanaGlobal.configData.keywordMapBarHeight
            self.whitespaceHeight = TransanaGlobal.configData.keywordMapWhitespace
            self.hGridLines = TransanaGlobal.configData.keywordMapHorizontalGridLines
            self.vGridLines = TransanaGlobal.configData.keywordMapVerticalGridLines
            # We default to Color Output.  When this was configurable, if a new Map was
            # created in B & W, the colors never worked right afterwards.
            self.colorOutput = True
        else:
            # Get the Configuration values for the Keyword Visualization Options
            self.barHeight = TransanaGlobal.configData.keywordVisualizationBarHeight
            self.whitespaceHeight = TransanaGlobal.configData.keywordVisualizationWhitespace
            self.hGridLines = TransanaGlobal.configData.keywordVisualizationHorizontalGridLines
            self.vGridLines = TransanaGlobal.configData.keywordVisualizationVerticalGridLines
            self.colorOutput = True

    def Setup(self, episodeNum, seriesName, episodeName):
        """ Complete initialization for the free-standing Keyword Map, not the embedded version. """
        # Remember the appropriate Episode information
        self.episodeNum = episodeNum
        self.seriesName = seriesName
        self.episodeName = episodeName
        # indicate that we're not working from a Clip.  (The Keyword Map is never Clip-based.)
        self.clipNum = None
        # You can't have a separate menu on the Mac, so we'll use a Toolbar
        self.toolBar = self.CreateToolBar(wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_TEXT)
        # Get the graphic for the Filter button
        bmp = wx.ArtProvider_GetBitmap(wx.ART_LIST_VIEW, wx.ART_TOOLBAR, (16,16))
        self.toolBar.AddTool(T_FILE_FILTER, bmp, shortHelpString=_("Filter"))
        self.toolBar.AddTool(T_FILE_SAVEAS, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "SaveJPG16.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Save As'))
        self.toolBar.AddTool(T_FILE_PRINTSETUP, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "PrintSetup.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Set up Page'))
        self.toolBar.AddTool(T_FILE_PRINTPREVIEW, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "PrintPreview.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Print Preview'))

        # Disable Print Preview on the Mac
        if 'wxMac' in wx.PlatformInfo:
            self.toolBar.EnableTool(T_FILE_PRINTPREVIEW, False)
            
        self.toolBar.AddTool(T_FILE_PRINT, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Print.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Print'))
        # Get the graphic for Help
        bmp = wx.ArtProvider_GetBitmap(wx.ART_HELP, wx.ART_TOOLBAR, (16,16))
        # create a bitmap button for the Move Down button
        self.toolBar.AddTool(T_HELP_HELP, bmp, shortHelpString=_("Help"))
        self.toolBar.AddTool(T_FILE_EXIT, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Exit.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Exit'))        
        self.toolBar.Realize()
        # Let's go ahead and keep the menu for non-Mac platforms
        if not '__WXMAC__' in wx.PlatformInfo:
            # Add a Menu Bar
            menuBar = wx.MenuBar()                                                 # Create the Menu Bar
            self.menuFile = wx.Menu()                                                   # Create the File Menu
            self.menuFile.Append(M_FILE_FILTER, _("&Filter"), _("Filter report contents"))  # Add "Filter" to File Menu
            self.menuFile.Append(M_FILE_SAVEAS, _("Save &As"), _("Save image in JPEG format"))  # Add "Save As" to File Menu
            self.menuFile.Enable(M_FILE_SAVEAS, False)
            self.menuFile.Append(M_FILE_PRINTSETUP, _("Page Setup"), _("Set up Page")) # Add "Page Setup" to the File Menu
            self.menuFile.Append(M_FILE_PRINTPREVIEW, _("Print Preview"), _("Preview your printed output")) # Add "Print Preview" to the File Menu
            self.menuFile.Enable(M_FILE_PRINTPREVIEW, False)
            self.menuFile.Append(M_FILE_PRINT, _("&Print"), _("Send your output to the Printer")) # Add "Print" to the File Menu
            self.menuFile.Enable(M_FILE_PRINT, False)
            self.menuFile.Append(M_FILE_EXIT, _("E&xit"), _("Exit the Keyword Map program")) # Add "Exit" to the File Menu
            menuBar.Append(self.menuFile, _('&File'))                                     # Add the File Menu to the Menu Bar
            self.menuHelp = wx.Menu()
            self.menuHelp.Append(M_HELP_HELP, _("&Help"), _("Help"))
            menuBar.Append(self.menuHelp, _("&Help"))
            self.SetMenuBar(menuBar)                                              # Connect the Menu Bar to the Frame
        # Link menu items and toolbar buttons to the appropriate methods
        wx.EVT_MENU(self, M_FILE_FILTER, self.OnFilter)                               # Attach File > Filter to a method
        wx.EVT_MENU(self, T_FILE_FILTER, self.OnFilter)                               # Attach Toolbar Filter to a method
        wx.EVT_MENU(self, M_FILE_SAVEAS, self.OnSaveAs)                               # Attach File > Save As to a method
        wx.EVT_MENU(self, T_FILE_SAVEAS, self.OnSaveAs)                               # Attach Toolbar Save As to a method
        wx.EVT_MENU(self, M_FILE_PRINTSETUP, self.OnPrintSetup)                       # Attach File > Print Setup to a method
        wx.EVT_MENU(self, T_FILE_PRINTSETUP, self.OnPrintSetup)                       # Attach Toolbar Print Setup to a method
        wx.EVT_MENU(self, M_FILE_PRINTPREVIEW, self.OnPrintPreview)                   # Attach File > Print Preview to a method
        wx.EVT_MENU(self, T_FILE_PRINTPREVIEW, self.OnPrintPreview)                   # Attach Toolbar Print Preview to a method
        wx.EVT_MENU(self, M_FILE_PRINT, self.OnPrint)                                 # Attach File > Print to a method
        wx.EVT_MENU(self, T_FILE_PRINT, self.OnPrint)                                 # Attach Toolbar Print to a method
        wx.EVT_MENU(self, M_FILE_EXIT, self.CloseWindow)                              # Attach CloseWindow to File > Exit
        wx.EVT_MENU(self, T_FILE_EXIT, self.CloseWindow)                              # Attach CloseWindow to Toolbar Exit
        wx.EVT_MENU(self, M_HELP_HELP, self.OnHelp)
        wx.EVT_MENU(self, T_HELP_HELP, self.OnHelp)

        # Determine the window boundaries
        (w, h) = self.GetClientSizeTuple()
        if (self.seriesName != '') and (self.episodeName != ''):
            self.Bounds = (5, 5, w - 10, h - 25)
        else:
            self.Bounds = (5, 40, w - 10, h - 30)

        # Create the Graphic Area using the GraphicControlClass
        # NOTE:  EVT_LEFT_DOWN, EVT_LEFT_UP, and EVT_RIGHT_UP are caught in GraphicsControlClass and are passed to this routine's
        #        OnLeftDown and OnLeftUp (for both left and right) methods because of the "passMouseEvents" paramter
        self.graphic = GraphicsControlClass.GraphicsControl(self, -1, wx.Point(self.Bounds[0], self.Bounds[1]),
                                                           (self.Bounds[2] - self.Bounds[0], self.Bounds[3] - self.Bounds[1]),
                                                           (self.Bounds[2] - self.Bounds[0], self.Bounds[3] - self.Bounds[1]),
                                                            passMouseEvents=True)

        # Add a Status Bar
        self.CreateStatusBar()

        # Attach the Resize Event
        wx.EVT_SIZE(self, self.OnSize)

        # We'll detect mouse movement in the GraphicsControlClass from out here, as
        # the KeywordMap object is the object that knows what the data is on the graphic.
        self.graphic.Bind(wx.EVT_MOTION, self.OnMouseMotion)

        # Prepare objects for use in Printing
        self.printData = wx.PrintData()
        self.printData.SetPaperId(wx.PAPER_LETTER)

        # Center on the screen
        self.CenterOnScreen()
        # Show the Frame
        self.Show(True)

        if (self.seriesName != '') and (self.episodeName != ''):
            # Clear the drawing
            self.filteredKeywordList = []
            self.unfilteredKeywordList = []
            # Populate the drawing
            self.ProcessEpisode()
            self.DrawGraph()

    def SetupEmbedded(self, episodeNum, seriesName, episodeName, startTime, endTime, filteredKeywordList=[],
                      unfilteredKeywordList = [], keywordColors = None, clipNum=None):
        """ Complete setup for the embedded version of the Keyword Map. """
        # Remember the appropriate Episode information
        self.episodeNum = episodeNum
        self.seriesName = seriesName
        self.episodeName = episodeName
        self.clipNum = clipNum
        # Set the start and end time boundaries (especially important for Clips!)
        self.startTime = startTime
        self.endTime = endTime
        # Toggle the Embedded labels.  (Used for testing mouse-overs only)
        self.showEmbeddedLabels = False
        
        # Determine the graphic's boundaries
        w = self.graphic.getWidth()
        h = self.graphic.getHeight()
        if (self.seriesName != '') and (self.episodeName != ''):
            self.Bounds = (0, 0, w, h - 25)
        else:
            self.Bounds = (0, 0, w, h - 25)
        # If we have a defined Episode (which we always should) ...
        if (self.seriesName != '') and (self.episodeName != ''):
            # set the initial keyword lists
            self.filteredKeywordList = filteredKeywordList[:]
            self.unfilteredKeywordList = unfilteredKeywordList[:]
            # If we got keywordColors, use them!!
            if keywordColors != None:
                self.keywordColors = keywordColors
            # Populate the drawing
            self.ProcessEpisode()
            # Actually draw the graph
            self.DrawGraph()

    # Define the Method that implements Filter
    def OnFilter(self, event):
        """ Implement the Filter Dialog call for Keyword Maps and Keyword Visualizations """
        # Set up parameters for creating the Filter Dialog.  Keyword Map/Keyword Visualization Filter requires episodeNum for the Config Save.
        if not self.embedded:
            # For the keyword map, the form created here is the parent
            parent = self
            title = _("Keyword Map Filter Dialog")
            # Keyword Map wants the Clip Filter
            clipFilter = True
            # Keyword map does not support Keyword Color customization.  (Colors represent Clips!)
            keywordColors = False
            # We want the Options tab
            options = True
            # reportType=1 indicates it is for a Keyword Map.  
            reportType = 1
            # Create a Filter Dialog, passing all the necessary parameters.
            dlgFilter = FilterDialog.FilterDialog(parent, -1, title, reportType=reportType, configName=self.configName,
                                                  reportScope=self.episodeNum, clipFilter=clipFilter, keywordFilter=True, keywordSort=True,
                                                  keywordColor=keywordColors, options=options, startTime=self.startTime, endTime=self.endTime,
                                                  barHeight=self.barHeight, whitespace=self.whitespaceHeight, hGridLines=self.hGridLines,
                                                  vGridLines=self.vGridLines, colorOutput=self.colorOutput)
        else:
            # For the keyword visualization, the parent that was passed in on initialization is the parent
            parent = self.parent
            title = _("Keyword Visualization Filter Dialog")
            # Keyword visualization does NOT want the Clip Filter
            clipFilter = False
            # Keyword visualization wants Keyword Color customization
            keywordColors = True
            # We want the Options tab
            options = True
            # reportType=2 indicates it is for a Keyword Visualization.  
            reportType = 2
            # Create a Filter Dialog, passing all the necessary parameters.
            dlgFilter = FilterDialog.FilterDialog(parent, -1, title, reportType=reportType, configName=self.configName,
                                                  reportScope=self.episodeNum, clipFilter=clipFilter, keywordFilter=True, keywordSort=True,
                                                  keywordColor=keywordColors, options=options, startTime=self.startTime, endTime=self.endTime,
                                                  barHeight=self.barHeight, whitespace=self.whitespaceHeight, hGridLines=self.hGridLines,
                                                  vGridLines=self.vGridLines)
        # If we requested the Clip Filter ...
        if clipFilter:
            # We want the Clips sorted in Clip ID order in the FilterDialog.  We handle that out here, as the Filter Dialog
            # has to deal with manual clip ordering in some instances, though not here, so it can't deal with this.
            self.clipFilterList.sort()
            # Inform the Filter Dialog of the Clips
            dlgFilter.SetClips(self.clipFilterList)
        # Keyword Colors must be specified before Keywords!  So if we want Keyword Colors, ...
        if keywordColors:
            # Inform the Filter Dialog of the colors used for each Keyword
            dlgFilter.SetKeywordColors(self.keywordColors)
        # Inform the Filter Dialog of the Keywords
        dlgFilter.SetKeywords(self.unfilteredKeywordList)
        # Create a dummy error message to get our while loop started.
        errorMsg = 'Start Loop'
        # Keep trying as long as there is an error message
        while errorMsg != '':
            # Clear the last (or dummy) error message.
            errorMsg = ''
            # Show the Filter Dialog and see if the user clicks OK
            if dlgFilter.ShowModal() == wx.ID_OK:
                # If we requested Clip Filtering ...
                if clipFilter:
                    # ... then get the filtered clip data
                    self.clipFilterList = dlgFilter.GetClips()
                # Get the complete list of keywords from the Filter Dialog.  We'll deduce the filter info in a moment.
                # (This preserves the "check" info for later reuse.)
                self.unfilteredKeywordList = dlgFilter.GetKeywords()
                # If we requested Keyword Color data ...
                if keywordColors:
                    # ... then get the keyword color data from the Filter Dialog
                    self.keywordColors = dlgFilter.GetKeywordColors()
                # Reset the Filtered Keyword List
                self.filteredKeywordList = []
                # Iterate through the entire Keword List ...
                for (kwg, kw, checked) in self.unfilteredKeywordList:
                    # ... and determine which keywords were checked.
                    if checked:
                        # Only the checked ones go into the filtered keyword list.
                        self.filteredKeywordList.append((kwg, kw))

                # If we had an Options Tab, extract that data.
                if options:
                    # If we're in the Keyword Map ...
                    if not self.embedded:
                        # Let's get the Time Range data.
                        # Start Time must be 0 or greater.  Otherwise, don't change it!
                        if Misc.time_in_str_to_ms(dlgFilter.GetStartTime()) >= 0:
                            self.startTime = Misc.time_in_str_to_ms(dlgFilter.GetStartTime())
                        else:
                            errorMsg += _("Illegal value for Start Time.\n")
                        
                        # If the Start Time is greater than the media length, reset it to 0.
                        if self.startTime >= self.MediaLength:
                            dlgFilter.startTime.SetValue(Misc.time_in_ms_to_str(0))
                            errorMsg += _("Illegal value for Start Time.\n")

                        # End Time must be at least 0.  Otherwise, don't change it!
                        if (Misc.time_in_str_to_ms(dlgFilter.GetEndTime()) >= 0):
                            self.endTime = Misc.time_in_str_to_ms(dlgFilter.GetEndTime())
                        else:
                            errorMsg += _("Illegal value for End Time.\n")

                        # If the end time is 0 or greater than the media length, set it to the media length.
                        if (self.endTime == 0) or (self.endTime > self.MediaLength):
                            self.endTime = self.MediaLength

                        # Start time cannot equal end time (but this check must come after setting endtime == 0 to MediaLength)
                        if self.startTime == self.endTime:
                            errorMsg += _("Start Time and End Time must be different.")
                            # We need to alter the time values to prevent "division by zero" errors while the Filter Dialog is not modal.
                            self.startTime = 0
                            self.endTime = self.MediaLength

                        # If the Start Time is greater than the End Time, swap them.
                        if (self.endTime < self.startTime):
                            temp = self.startTime
                            self.startTime = self.endTime
                            self.endTime = temp

                        # Get the colorOutput value from the dialog IF we're in the Keyword Map
                        self.colorOutput = dlgFilter.GetColorOutput()

                    # Get the Bar Height and Whitespace Height for both versions of the Keyword Map
                    self.barHeight = dlgFilter.GetBarHeight()
                    self.whitespaceHeight = dlgFilter.GetWhitespace()
                    # we need to store the Bar Height and Whitespace values in the Configuration.
                    if not self.embedded:
                        TransanaGlobal.configData.keywordMapBarHeight = self.barHeight
                        TransanaGlobal.configData.keywordMapWhitespace = self.whitespaceHeight
                    else:
                        TransanaGlobal.configData.keywordVisualizationBarHeight = self.barHeight
                        TransanaGlobal.configData.keywordVisualizationWhitespace = self.whitespaceHeight

                    # Get the Grid Line data from the form
                    self.hGridLines = dlgFilter.GetHGridLines()
                    self.vGridLines = dlgFilter.GetVGridLines()
                    # Store the Grid Line data in the Configuration
                    if not self.embedded:
                        TransanaGlobal.configData.keywordMapHorizontalGridLines = self.hGridLines
                        TransanaGlobal.configData.keywordMapVerticalGridLines = self.vGridLines
                    else:
                        TransanaGlobal.configData.keywordVisualizationHorizontalGridLines = self.hGridLines
                        TransanaGlobal.configData.keywordVisualizationVerticalGridLines = self.vGridLines

            if errorMsg != '':
                errorDlg = Dialogs.ErrorDialog(self, errorMsg)
                errorDlg.ShowModal()
                errorDlg.Destroy()

        # Remember the configuration name for later reuse
        self.configName = dlgFilter.configName
        # Destroy the Filter Dialog.  We're done with it.
        dlgFilter.Destroy()
        # Now we can draw the graph.
        self.DrawGraph()

    # Define the Method that implements Save As
    def OnSaveAs(self, event):
        self.graphic.SaveAs()

    # Define the Method that implements Printer Setup
    def OnPrintSetup(self, event):
        # Destroying the printerDialog also wipes out the printData object in wxPython 2.5.1.5.  Don't do this.
        # printerDialog.Destroy()

        # Let's use PAGE Setup here ('cause you can do Printer Setup from Page Setup.)  It's a better system
        # that allows Landscape on Mac.
        pageSetupDialogData = wx.PageSetupDialogData(self.printData)
        pageSetupDialogData.CalculatePaperSizeFromId()
        pageDialog = wx.PageSetupDialog(self, pageSetupDialogData)
        pageDialog.ShowModal()
        self.printData = wx.PrintData(pageDialog.GetPageSetupData().GetPrintData())
        pageDialog.Destroy()

    # Define the Method that implements Print Preview
    def OnPrintPreview(self, event):
        printout = MyPrintout(_('Transana Keyword Map'), self.graphic)
        printout2 = MyPrintout(_('Transana Keyword Map'), self.graphic)
        self.preview = wx.PrintPreview(printout, printout2, self.printData)
        if not self.preview.Ok():
            self.SetStatusText(_("Print Preview Problem"))
            return
        theWidth = max(wx.ClientDisplayRect()[2] - 180, 760)
        theHeight = max(wx.ClientDisplayRect()[3] - 200, 560)
        frame2 = wx.PreviewFrame(self.preview, self, _("Print Preview"), size=(theWidth, theHeight))
        frame2.Centre()
        frame2.Initialize()
        frame2.Show(True)

    # Define the Method that implements Print
    def OnPrint(self, event):
        pdd = wx.PrintDialogData()
        pdd.SetPrintData(self.printData)
        printer = wx.Printer(pdd)
        printout = MyPrintout(_('Transana Keyword Map'), self.graphic)
        if not printer.Print(self, printout):
            dlg = Dialogs.ErrorDialog(None, _("There was a problem printing this report."))
            dlg.ShowModal()
            dlg.Destroy()
        # NO!  REMOVED to prevent crash on 2nd print attempt following Filter Config.
        # else:
        #     self.printData = printer.GetPrintDialogData().GetPrintData()
        printout.Destroy()

    # Define the Method that closes the Window on File > Exit
    def CloseWindow(self, event):
        self.Close()

    def OnHelp(self, event):
        """ Implement the Filter Dialog Box's Help function """
        # Define the Help Context
        HelpContext = "Keyword Map"
        # If a Help Window is defined ...
        if TransanaGlobal.menuWindow != None:
            # ... call Help!
            TransanaGlobal.menuWindow.ControlObject.Help(HelpContext)

    def OnSize(self, event):
        (w, h) = self.GetClientSizeTuple()
        if not self.embedded:
            if self.Bounds[1] == 5:
                self.Bounds = (5, 5, w - 10, h - 25)
            else:
                self.Bounds = (5, 40, w - 10, h - 30)
        else:
            self.Bounds = (0, 0, w, h - 25)
        if self.episodeName != '':
            self.DrawGraph()
            
    def CalcX(self, XPos):
        """ Determine the proper horizontal coordinate for the given time """
        # Specify a margin width
        if not self.embedded:
            marginwidth = (0.06 * (self.Bounds[2] - self.Bounds[0]))
        else:
            marginwidth = 0
        # The Horizonal Adjustment is the global graphic indent
        hadjust = self.graphicindent
        # The Scaling Factor is the active portion of the drawing area width divided by the total media length
        # The idea is to leave the left margin, self.graphicindent for Keyword Labels, and the right margin
        if self.MediaLength > 0:
            scale = (float(self.Bounds[2]) - self.Bounds[0] - hadjust - 2 * marginwidth) / (self.endTime - self.startTime)
        else:
            scale = 0.0
        # The horizontal coordinate is the left margin plus the Horizontal Adjustment for Keyword Labels plus
        # position times the scaling factor
        res = marginwidth + hadjust + ((XPos - self.startTime) * scale) 
        return int(res)

    def FindTime(self, x):
        """ Given a horizontal pixel position, determine the corresponding time value from
            the video time line """
        # determine the margin width
        if not self.embedded:
            marginwidth = (0.06 * (self.Bounds[2] - self.Bounds[0]))
        else:
            marginwidth = 0
        # The Horizonal Adjustment is the global graphic indent
        hadjust = self.graphicindent
        # The Scaling Factor is the active portion of the drawing area width divided by the total media length
        # The idea is to leave the left margin, self.graphicindent for Keyword Labels, and the right margin
        if self.MediaLength > 0:
            scale = (float(self.Bounds[2]) - self.Bounds[0] - hadjust - 2 * marginwidth) / (self.endTime - self.startTime)
        else:
            scale = 1.0
        # The time is calculated by taking the total width, subtracting the margin values and horizontal indent,
        # and then dividing the result by the scale factor calculated above
        time = int((x - marginwidth - hadjust) / scale) + self.startTime
        return time

    def CalcY(self, YPos):
        """ Determine the vertical position for a given keyword index """
        # For the Keyword Map Report
        if not self.embedded:
            # Spacing is the larger of (12 pixels for label text or the bar height) plus 2 for whitespace
            spacing = max(12, self.barHeight) + self.whitespaceHeight
            # Top margin is 30 for titles plus 28 for the timeline
            topMargin = 30 + (2 * spacing) + self.topOffset
        # For the Keyword Visualization
        else:
            # Top margin to the line CENTER is 5 pixels plus 1/2 line width to account for line thickness.
            # (The line is drawn centered around the line thickness, so this gives us a constant margin of 5
            # pixels to the TOP of the line!)
            topMargin = int(self.barHeight / 2) + 5 + self.topOffset
            # Spacing is the line height plus the whitespace!
            spacing = self.barHeight + self.whitespaceHeight
        return int(spacing * YPos + topMargin)

    def FindKeyword(self, y):
        """ Given a vertical pixel position, determine the corresponding Keyword data """
        # If the graphic is scrolled, the raw Y value does not point to the correct Keyword.
        # Determine the unscrolled equivalent Y position.
        (modX, modY) = self.graphic.CalcUnscrolledPosition(0, y)
        # Each keyword is a fixed number of pixels high, and we need to adjust for the top margin and
        # the amount of graphic scroll too.
        if not self.embedded:
            # Spacing is the larger of (12 pixels for label text or the bar height) plus 2 for whitespace
            spacing = max(12, self.barHeight) + self.whitespaceHeight
            # Top margin is 30 for titles, 28 for timeline, less 1/2 the inter-bar whitespace (bar is 5, whitespace is 9)
            topMargin = 30 + (2 * spacing) - (self.whitespaceHeight / 2) + self.topOffset
        else:
            # Top margin is 5 minus half the whitespace
            topMargin = 5 - (self.whitespaceHeight / 2) + self.topOffset
            # Spacing is bar height plus whitespace
            spacing = self.barHeight + self.whitespaceHeight
        # Calculate the integer index from the pixel position
        kwIndex = int((modY - topMargin) / spacing)
        # We can only return keywords from the part of the graph where there ARE keywords.
        if (kwIndex >= 0) and (kwIndex < len(self.filteredKeywordList)):
            # Determine which keyword the position corresponds to
            (kwg, kw) = self.filteredKeywordList[kwIndex]
            # ... and return it
            return (kwg, kw)
        # If the cursor is out of range ...
        else:
            # ... return None to signal that no keyword is at that position
            return None

    def GetScaleIncrements(self, MediaLength):
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

    def ProcessEpisode(self):
        self.clipFilterList = []
        
        # We need a data struture to hold the data about what clips correspond to what keywords
        self.MediaFile = ''
        self.MediaLength = 0
        # Get Series Number, Episode Number, Media File Name, and Length
        SQLText = """SELECT e.EpisodeNum, e.SeriesNum, e.MediaFile, e.EpLength
                       FROM Episodes2 e, Series2 s
                       WHERE s.SeriesID = %s AND
                             s.SeriesNum = e.SeriesNum AND
                             e.EpisodeID = %s"""
        if 'unicode' in wx.PlatformInfo:
            querySeriesName = self.seriesName.encode(TransanaGlobal.encoding)
            queryEpisodeName = self.episodeName.encode(TransanaGlobal.encoding)
        else:
            querySeriesName = self.seriesName
            queryEpisodeName = self.episodeName
        self.DBCursor.execute(SQLText, (querySeriesName, queryEpisodeName))
        if self.DBCursor.rowcount == 1:
            (EpisodeNum, SeriesNum, MediaFile, EpisodeLength) = self.DBCursor.fetchone()
            # Capture Media File Name and Length for use in the Graph
            self.MediaFile = os.path.split(MediaFile)[1]
            self.MediaLength = EpisodeLength
            # If the end time is 0 or greater than the media length, set it to the media length.
            if (self.endTime == 0) or (self.endTime > self.MediaLength):
                self.endTime = self.MediaLength
        # If we don't have a single record from the database, we probably have an orphaned Clip.
        else:
            # In that case, we can just use the Episode Number passed in by the calling routine ...
            EpisodeNum = self.episodeNum
            # ... and we can set the MediaLength to the end time passed in.
            self.MediaLength = self.endTime

        if self.filteredKeywordList == []:
            # If we deleted the last keyword in a filtered list, the Filter Dialog ended up with
            # duplicate entries.  This should prevent it!!
            self.unfilteredKeywordList = []
            # Get the list of Keywords to be displayed
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                           FROM Clips2 cl, ClipKeywords2 ck
                           WHERE cl.EpisodeNum = %s AND
                                 cl.ClipNum = ck.ClipNum
                           GROUP BY ck.keywordgroup, ck.keyword
                           ORDER BY KeywordGroup, Keyword, ClipStart"""
            self.DBCursor.execute(SQLText, EpisodeNum)
            for (kwg, kw) in self.DBCursor.fetchall():
                kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                self.filteredKeywordList.append((kwg, kw))
                self.unfilteredKeywordList.append((kwg, kw, True))

        # Create the Keyword Placement lines to be displayed.  We need them to be in ClipStart, ClipNum order so colors will be
        # distributed properly across bands.
        SQLText = """SELECT ck.KeywordGroup, ck.Keyword, cl.ClipStart, cl.ClipStop, cl.ClipNum, cl.ClipID, cl.CollectNum
                       FROM Clips2 cl, ClipKeywords2 ck
                       WHERE cl.EpisodeNum = %s AND
                             cl.ClipNum = ck.ClipNum
                       ORDER BY ClipStart, ClipNum, KeywordGroup, Keyword"""
        self.DBCursor.execute(SQLText, EpisodeNum)
        for (kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum) in self.DBCursor.fetchall():
            kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
            kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
            clipID = DBInterface.ProcessDBDataForUTF8Encoding(clipID)
            # If we're dealing with an Episode, self.clipNum will be None and we want all clips.
            # If we're dealing with a Clip, we only want to deal with THIS clip!
            if (self.clipNum == None) or (clipNum == self.clipNum):
                self.clipList.append((kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum))
                if not ((clipID, collectNum, True) in self.clipFilterList):
                    self.clipFilterList.append((clipID, collectNum, True))

    def UpdateKeywordVisualization(self):
        """ Update the Keyword Visualization following something that could have changed it. """
        # If the Keyword Map hasn't been Setup yet, skip this.
        if self.episodeNum == None:
            return
        # Clear the Clip List
        self.clipList = []
        # Clear the Filtered Clip List
        self.clipFilterList = []
        # Clear the graphic itself
        self.graphic.Clear()

        # Before we start, make a COPY of the keyword list so we can check for keywords that are no longer
        # included on the Map and need to be deleted from the KeywordLists
        delList = self.unfilteredKeywordList[:]
        # Now let's create the SQL to get all relevant Clip and Clip Keyword records
        SQLText = """SELECT ck.KeywordGroup, ck.Keyword, cl.ClipStart, cl.ClipStop, cl.ClipNum, cl.ClipID, cl.CollectNum
                       FROM Clips2 cl, ClipKeywords2 ck
                       WHERE cl.EpisodeNum = %s AND
                             cl.ClipNum = ck.ClipNum
                       ORDER BY ClipStart, ClipNum, KeywordGroup, Keyword"""
        # Execute the query
        self.DBCursor.execute(SQLText, self.episodeNum)
        # Iterate through the results ...
        for (kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum) in self.DBCursor.fetchall():
            kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
            kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
            clipID = DBInterface.ProcessDBDataForUTF8Encoding(clipID)
            # If we're dealing with an Episode, self.clipNum will be None and we want all clips.
            # If we're dealing with a Clip, we only want to deal with THIS clip!
            if (self.clipNum == None) or (clipNum == self.clipNum):
                # If a Clip is not found in the clipList ...
                if not ((kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum) in self.clipList):
                    # ... add it to the clipList ...
                    self.clipList.append((kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum))
                    # ... and if it's not in the clipFilter List (which it probably isn't!) ...
                    if not ((clipID, collectNum, True) in self.clipFilterList):
                        # ... add it to the clipFilterList.
                        self.clipFilterList.append((clipID, collectNum, True))

            # If the keyword is not in either of the Keyword Lists, ...
            if not (((kwg, kw) in self.filteredKeywordList) or ((kwg, kw, False) in self.unfilteredKeywordList)):
                # ... add it to both keyword lists.
                self.filteredKeywordList.append((kwg, kw))
                self.unfilteredKeywordList.append((kwg, kw, True))

            # If the keyword is in query results, it should be removed from the list of keywords to be deleted.
            # Check that list for either True or False versions of the keyword!
            if (kwg, kw, True) in delList:
                del(delList[delList.index((kwg, kw, True))])
            if (kwg, kw, False) in delList:
                del(delList[delList.index((kwg, kw, False))])
        # Iterate through ANY keywords left in the list of keywords to be deleted ...
        for element in delList:
            # ... and delete them from the unfiltered Keyword List
            del(self.unfilteredKeywordList[self.unfilteredKeywordList.index(element)])
            # If the keyword is also in the filtered keyword list ...
            if (element[0], element[1]) in self.filteredKeywordList:
                # ... it needs to be deleted from there too!
                del(self.filteredKeywordList[self.filteredKeywordList.index((element[0], element[1]))])

        # Now that the underlying data structures have been corrected, we're ready to redraw the Keyword Visualization
        self.DrawGraph()
        

    def DrawGraph(self):
        """ Actually Draw the Keyword Map """
        self.keywordClipList = {}

        if not self.embedded:
            # Now that we have all necessary information, let's create and populate the graphic
            # Start by destroying the existing control and creating a new one with the correct Canvas Size
            self.graphic.Destroy()
            newheight = max(self.CalcY(len(self.filteredKeywordList) + 1), self.Bounds[3] - self.Bounds[1])
            self.graphic = GraphicsControlClass.GraphicsControl(self, -1, wx.Point(self.Bounds[0], self.Bounds[1]),
                                                                (self.Bounds[2] - self.Bounds[0], self.Bounds[3] - self.Bounds[1]),
                                                                (self.Bounds[2] - self.Bounds[0], newheight + 3),
                                                                passMouseEvents=True)

        self.graphic.SetFontColour("BLACK")
        if 'wxMac' in wx.PlatformInfo:
            self.graphic.SetFontSize(13)
        else:
            self.graphic.SetFontSize(10)
        if not self.embedded:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Series: %s'), 'utf8')
            else:
                prompt = _('Series: %s')
            self.graphic.AddText(prompt % self.seriesName, 2, 2)
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Episode: %s"), 'utf8')
            else:
                prompt = _("Episode: %s")
            self.graphic.AddTextCentered(prompt % self.episodeName, (self.Bounds[2] - self.Bounds[0]) / 2, 2)
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('File: %s'), 'utf8')
            else:
                prompt = _('File: %s')
            self.graphic.AddTextRight(prompt % DBInterface.ProcessDBDataForUTF8Encoding(self.MediaFile), self.Bounds[2] - self.Bounds[1], 2)
            if self.configName != '':
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Filter Configuration: %s'), 'utf8')
                else:
                    prompt = _('Filter Configuration: %s')
                self.graphic.AddText(prompt % self.configName, 2, 16)
                
            Count = 0
            # We want Grid Lines in light gray
            self.graphic.SetColour('LIGHT GREY')
            # Draw the top Grid Line, if appropriate
            if self.hGridLines:
                self.graphic.AddLines([(10, self.CalcY(-1) + 6 + int(self.whitespaceHeight / 2), self.CalcX(self.endTime), self.CalcY(-1) + 6 + int(self.whitespaceHeight / 2))])
            for KWG, KW in self.filteredKeywordList:
                self.graphic.AddText("%s : %s" % (KWG, KW), 10, self.CalcY(Count) - 7)
                # Add Horizontal Grid Lines, if appropriate
                if self.hGridLines and (Count % 2 == 1):
                    self.graphic.AddLines([(10, self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2), self.CalcX(self.endTime), self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2))])
                Count = Count + 1
            # Reset the graphic color following drawing the Grid Lines
            self.graphic.SetColour("BLACK")
                
            self.graphicindent = self.graphic.GetMaxWidth(start=3)
        else:
            # We need 2 pixels to account for the rounded edge of the thick line in the Keyword visualization
            self.graphicindent = 2
            # Draw the top Grid Line, if appropriate
            if self.hGridLines:
                Count = 0
                # We want Grid Lines in light gray
                self.graphic.SetColour('LIGHT GREY')
                for KWG, KW in self.filteredKeywordList:
                    # Add Horizontal Grid Lines, if appropriate
                    if self.hGridLines and (Count % 2 == 1):
                        self.graphic.AddLines([(0, self.CalcY(Count) + 3 + int(self.whitespaceHeight / 2), self.CalcX(self.endTime), self.CalcY(Count) + 3 + int(self.whitespaceHeight / 2))])
                    Count = Count + 1
                # Reset the graphic color following drawing the Grid Lines
                self.graphic.SetColour("BLACK")

        if not self.embedded:
            # If the Media Length is known, display the Time Line
            if self.MediaLength > 0:
                self.graphic.SetThickness(3)
                self.graphic.AddLines([(self.CalcX(self.startTime), self.CalcY(-2), self.CalcX(self.endTime), self.CalcY(-2))])
                # Add Time markers
                self.graphic.SetThickness(1)
                if 'wxMac' in wx.PlatformInfo:
                    self.graphic.SetFontSize(11)
                else:
                    self.graphic.SetFontSize(8)

                X = self.startTime
                self.graphic.AddLines([(self.CalcX(X), self.CalcY(-2) + 1, self.CalcX(X), self.CalcY(-2) + 6)])
                XLabel = Misc.TimeMsToStr(X)
                self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(-2) + 5)
                X = self.endTime
                self.graphic.AddLines([(self.CalcX(X), self.CalcY(-2) + 1, self.CalcX(X), self.CalcY(-2) + 6)])
                XLabel = Misc.TimeMsToStr(X)
                self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(-2) + 5)

                # Add the first and last Vertical Grid Lines, if appropriate
                if self.vGridLines:
                    # We want Grid Lines in light gray
                    self.graphic.SetColour('LIGHT GREY')
                    self.graphic.AddLines([(self.CalcX(self.startTime), self.CalcY(0) - 6 - int(self.whitespaceHeight / 2), self.CalcX(self.startTime), self.CalcY(len(self.filteredKeywordList)) - 6 - int(self.whitespaceHeight / 2))])
                    self.graphic.AddLines([(self.CalcX(self.endTime), self.CalcY(0) - 6 - int(self.whitespaceHeight / 2), self.CalcX(self.endTime), self.CalcY(len(self.filteredKeywordList)) - 6 - int(self.whitespaceHeight / 2))])
                    # Reset the graphic color following drawing the Grid Lines
                    self.graphic.SetColour("BLACK")
                (numMarks, interval) = self.GetScaleIncrements(self.endTime - self.startTime)
                for loop in range(1, numMarks):
                    X = int(round(float(loop) * interval) + self.startTime)
                    self.graphic.AddLines([(self.CalcX(X), self.CalcY(-2) + 1, self.CalcX(X), self.CalcY(-2) + 6)])
                    XLabel = Misc.TimeMsToStr(X)
                    self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(-2) + 5)
                    # Add Vertical Grid Lines, if appropriate
                    if self.vGridLines:
                        # We want Grid Lines in light gray
                        self.graphic.SetColour('LIGHT GREY')
                        self.graphic.AddLines([(self.CalcX(X), self.CalcY(0) - 6 - int(self.whitespaceHeight / 2), self.CalcX(X), self.CalcY(len(self.filteredKeywordList)) - 6 - int(self.whitespaceHeight / 2))])
                        # Reset the graphic color following drawing the Grid Lines
                        self.graphic.SetColour("BLACK")
        else:
            # Add Vertical Grid Lines, if appropriate
            if self.vGridLines:
                # We want Grid Lines in light gray
                self.graphic.SetColour('LIGHT GREY')
                (numMarks, interval) = self.GetScaleIncrements(self.endTime - self.startTime)
                for loop in range(1, numMarks):
                    X = int(round(float(loop) * interval) + self.startTime)
                    self.graphic.AddLines([(self.CalcX(X) - 1, self.CalcY(0) - int(self.whitespaceHeight / 2), self.CalcX(X) - 1, self.CalcY(len(self.filteredKeywordList)) - int(self.whitespaceHeight / 2))])
                # Reset the graphic color following drawing the Grid Lines
                self.graphic.SetColour("BLACK")

        colourindex = self.keywordColors['lastColor']
        lastclip = 0
        # some clip boundary lines for overlapping clips can get over-written, depeding on the nature of the overlaps.
        # Let's create a separate list of these lines, which we'll add to the END of the process so they can't get overwritten.
        overlapLines = []

        # implement colors or gray scale as appropriate
        if self.colorOutput:
            colorSet = TransanaConstants.keywordMapColourSet
            colorLookup = TransanaConstants.transana_colorLookup
        else:
            colorSet = TransanaConstants.keywordMapGraySet
            colorLookup = TransanaConstants.transana_grayLookup
        if self.embedded:
            if 'wxMac' in wx.PlatformInfo:
                self.graphic.SetFontSize(10)
            else:
                self.graphic.SetFontSize(7)
            # Iterate through the keyword list in order ...
            for (KWG, KW) in self.filteredKeywordList:
                # ... and assign colors to Keywords
                if self.keywordColors.has_key((KWG, KW)) and self.colorOutput:
                    colourindex = self.keywordColors[(KWG, KW)]
                else:
                    colourindex = self.keywordColors['lastColor'] + 1
                    if colourindex > len(colorSet) - 1:
                        colourindex = 0
                    self.keywordColors['lastColor'] = colourindex
                    self.keywordColors[(KWG, KW)] = colourindex

                if self.showEmbeddedLabels:
                    self.graphic.AddText("%s : %s" % (KWG, KW), 2, self.CalcY(self.filteredKeywordList.index((KWG, KW))) - 7)

        for (KWG, KW, Start, Stop, ClipNum, ClipName, CollectNum) in self.clipList:
            if ((ClipName, CollectNum, True) in self.clipFilterList) and ((KWG, KW) in self.filteredKeywordList):

                if Start < self.startTime:
                    Start = self.startTime
                if Start > self.endTime:
                    Start = self.endTime
                if Stop > self.endTime:
                    Stop = self.endTime
                if Stop < self.startTime:
                    Stop = self.startTime

                if Start != Stop:
                    self.graphic.SetThickness(self.barHeight)
                    tempLine = []
                    tempLine.append((self.CalcX(Start), self.CalcY(self.filteredKeywordList.index((KWG, KW))), self.CalcX(Stop), self.CalcY(self.filteredKeywordList.index((KWG, KW)))))
                    if not self.embedded:
                        if (ClipNum != lastclip) and (lastclip != 0):
                             if colourindex < len(colorSet) - 1:
                                 colourindex = colourindex + 1
                             else:
                                 colourindex = 0
                    else:
                        colourindex = self.keywordColors[(KWG, KW)]
                            
                    self.graphic.SetColour(colorLookup[colorSet[colourindex]])

                    self.graphic.AddLines(tempLine)
                    
                lastclip = ClipNum

                if DEBUG and KWG == 'Transana Users' and KW == 'DavidW':
                    print "Looking at %s (%d)" % (ClipName, CollectNum)
                            
                # Now add the Clip to the keywordClipList.  This holds all Keyword/Clip data in memory so it can be searched quickly
                # This dictionary object uses the keyword pair as the key and holds a list of Clip data for all clips with that keyword.
                # If the list for a given keyword already exists ...
                if self.keywordClipList.has_key((KWG, KW)):
                    # Let's look for overlap
                    overlapStart = Stop
                    overlapEnd = Start
                    # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                    overlapClips = self.keywordClipList[(KWG, KW)]
                    # Iterate through the Clip List ...
                    for (overlapStartTime, overlapEndTime, overlapClipNum, overlapClipName) in overlapClips:

                        if DEBUG and KWG == 'Transana Users' and KW == 'DavidW':
                            print "Start = %7d, overStart = %7s, Stop = %7s, overEnd = %7s" % (Start, overlapStartTime, Stop, overlapEndTime)

                        # Look for Start between overlapStartTime and overlapEndTime
                        if (Start >= overlapStartTime) and (Start < overlapEndTime):
                            overlapStart = Start

                        # Look for overlapStartTime between Start and Stop
                        if (overlapStartTime >= Start) and (overlapStartTime < Stop):
                            overlapStart = overlapStartTime

                        # Look for Stop between overlapStartTime and overlapEndTime
                        if (Stop > overlapStartTime) and (Stop <= overlapEndTime):
                            overlapEnd = Stop

                        # Look for overlapEndTime between Start and Stop
                        if (overlapEndTime > Start) and (overlapEndTime <= Stop):
                            overlapEnd = overlapEndTime
                            
                    # If we've found an overlap, it will be indicated by Start being less than End!
                    if overlapStart < overlapEnd:
                        # Draw a multi-colored line to indicate overlap
                        overlapThickness = int(self.barHeight/ 3) + 1
                        self.graphic.SetThickness(overlapThickness)
                        if self.colorOutput:
                            self.graphic.SetColour("GREEN")
                        else:
                            self.graphic.SetColour("WHITE")
                        tempLine = [(self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW))), self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW))))]
                        self.graphic.AddLines(tempLine)
                        if self.colorOutput:
                            self.graphic.SetColour("RED")
                        else:
                            self.graphic.SetColour("BLACK")
                        tempLine = [(self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW)))-overlapThickness+1, self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW)))-overlapThickness+1)]
                        self.graphic.AddLines(tempLine)
                        if self.colorOutput:
                            self.graphic.SetColour("BLUE")
                        else:
                            self.graphic.SetColour("GRAY")
                        tempLine = [(self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW)))+overlapThickness, self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW)))+overlapThickness)]
                        self.graphic.AddLines(tempLine)
                        # Let's remember the clip start and stop boundaries, to be drawn at the end so they won't get over-written
                        overlapLines.append(((self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW)))-(self.barHeight / 2), self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW)))+(self.barHeight / 2)),))
                        overlapLines.append(((self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW)))-(self.barHeight / 2), self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW)))+(self.barHeight / 2)),))

                    # ... add the new Clip to the Clip List
                    self.keywordClipList[(KWG, KW)].append((Start, Stop, ClipNum, ClipName))
                # If there is no entry for the given keyword ...
                else:
                    # ... create a List object with the first clip's data for this Keyword Pair key
                    self.keywordClipList[(KWG, KW)] = [(Start, Stop, ClipNum, ClipName)]

        # If we are doing a Keyword Visualization, but there are no Clips in the picture, it can be confusing.
        # Let's place a message on the visualization saying it's intentionally left blank.
        if self.embedded and (lastclip == 0):
            self.graphic.AddText(_("No keywords meet the visualization display criteria."), 5, self.CalcY(0))
                    
        # let's add the overlap boundary lines now
        self.graphic.SetThickness(1)
        self.graphic.SetColour("BLACK")
        for tempLine in overlapLines:
            self.graphic.AddLines(tempLine)

        if not self.embedded:
            
            # Enable tracking of mouse movement over the graphic
            self.graphic.Bind(wx.EVT_MOTION, self.OnMouseMotion)

            if not '__WXMAC__' in wx.PlatformInfo:
                self.menuFile.Enable(M_FILE_SAVEAS, True)
                self.menuFile.Enable(M_FILE_PRINTPREVIEW, True)
                self.menuFile.Enable(M_FILE_PRINT, True)
        else:
            # The DrawGraph routine destroys and recreates self.graphic.  We need to re-point the waveform to it.
            self.parent.waveform = self.graphic

    def GetKeywordCount(self):
        """ Returns the number of keywords in the filtered Keyword List and the size of the image that results """
        return (len(self.filteredKeywordList), len(self.filteredKeywordList) * (self.barHeight + self.whitespaceHeight) + self.topOffset + 4)

    def OnMouseMotion(self, event):
        """ Process the movement of the mouse over the Keyword Map. """
        # If we're in the embedded version, call the graphic's OnMouseMotion event also!
        if self.embedded:
            self.graphic.TransanaOnMotion(event)
            
        # Get the mouse's current position
        x = event.GetX()
        y = event.GetY()
        # Based on the mouse position, determine the time in the video timeline
        time = self.FindTime(x)
        # Based on the mouse position, determine what keyword is being pointed to
        kw = self.FindKeyword(y)
        if not self.embedded:
            # First, let's make sure we're actually on the data portion of the graph
            if (time > 0) and (time < self.MediaLength) and (kw != None):
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(_("Keyword:  %s : %s,  Time: %s"), 'utf8')
                else:
                    prompt = _("Keyword:  %s : %s,  Time: %s")
                # Set the Status Text to indicate the current Keyword and Time values
                self.SetStatusText(prompt % (kw[0], kw[1], Misc.time_in_ms_to_str(time)))
                if (self.keywordClipList.has_key(kw)):
                    # initialize the string that will hold the names of clips being pointed to
                    clipNames = ''
                    # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                    clips = self.keywordClipList[kw]
                    # Iterate through the Clip List ...
                    for (startTime, endTime, clipNum, clipName) in clips:
                        # If the current Time value falls between the Clip's StartTime and EndTime ...
                        if (startTime < time) and (endTime > time):
                            # ... calculate the length of the Clip ...
                            clipLen = endTime - startTime
                            # ... and add the Clip Name and Length to the list of Clips with this Keyword at this Time
                            # First, see if the list is empty.
                            if clipNames == '':
                                # If so, just add the keyword name and time
                                clipNames = "%s (%s)" % (clipName, Misc.time_in_ms_to_str(clipLen))
                            else:
                                # ... add the keyword to the end of the list
                                clipNames += ', ' + "%s (%s)" % (clipName, Misc.time_in_ms_to_str(clipLen))
                    # If any clips are found for the current mouse position ...
                    if (clipNames != ''):
                        # ... add the Clip Names to the ToolTip so they will show up on screen as a hint
                        self.graphic.SetToolTipString(clipNames)
            # If we're not on the data portion of the graph ...
            else:
                # ... set the status text to a blank
                self.SetStatusText('')
        else:
            # We need to call the Visualization Window's "MouseOver" method for updating the cursor's time value
            self.parent.OnMouseOver(x, y, float(x) / self.Bounds[2], float(y) / self.Bounds[3])
            # First, let's make sure we're actually on the data portion of the graph
            if (time > 0) and (time < self.MediaLength) and (kw != None):
                if (self.keywordClipList.has_key(kw)):
                    # initialize the string that will hold the names of clips being pointed to.
                    # We don't actually need to know the names, but this signals that we're at least OVER a Clip.
                    clipNames = ''
                    # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                    clips = self.keywordClipList[kw]
                    # Iterate through the Clip List ...
                    for (startTime, endTime, clipNum, clipName) in clips:
                        # If the current Time value falls between the Clip's StartTime and EndTime ...
                        if (startTime < time) and (endTime > time):
                            # ... calculate the length of the Clip ...
                            clipLen = endTime - startTime
                            # ... and add the Clip LENGTH to the list of Clips with this Keyword at this Time
                            clipNames += " (%s)" % Misc.time_in_ms_to_str(clipLen)
                    # If any clips are found for the current mouse position ...
                    if (clipNames != ''):
                        # ... add the KEYWORD names to the ToolTip so they will show up on screen as a hint
                        self.graphic.SetToolTipString('%s : %s  -  %s' % (kw[0], kw[1], clipNames))

    def OnLeftDown(self, event):
        """ Left Mouse Button Down event """
        # Pass the event to the parent
        event.Skip()
        
    def OnLeftUp(self, event):
        """ Left Mouse Button Up event.  Triggers the load of a Clip. """
        # Pass the event to the parent
        event.Skip()
        # Get the mouse's current position
        x = event.GetX()
        y = event.GetY()
        # Based on the mouse position, determine the time in the video timeline
        time = self.FindTime(x)
        # Based on the mouse position, determine what keyword is being pointed to
        kw = self.FindKeyword(y)
        # Create an empty Dictionary Object for tracking Clip data
        clipNames = {}
        # First, let's make sure we're actually on the data portion of the graph
        if (time > 0) and (time < self.MediaLength) and (kw != None) and (self.keywordClipList.has_key(kw)):
            if 'unicode' in wx.PlatformInfo:
                prompt = unicode(_("Keyword:  %s : %s,  Time: %s"), 'utf8')
            else:
                prompt = _("Keyword:  %s : %s,  Time: %s")
            # Set the Status Text to indicate the current Keyword and Time values
            self.SetStatusText(prompt % (kw[0], kw[1], Misc.time_in_ms_to_str(time)))
            # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
            clips = self.keywordClipList[kw]
            # Iterate through the Clip List ...
            for (startTime, endTime, clipNum, clipName) in clips:
                # If the current Time value falls between the Clip's StartTime and EndTime ...
                if (startTime <= time) and (endTime >= time):
                    # Check to see if this is a duplicate Clip
                    if clipNames.has_key(clipName):
                        # If so, we need to count the number of duplicates.
                        # NOTE:  This is not perfect.  If the Clip Name is a shorter version of another Clip Name, the count
                        #        will be too high.
                        tmpList = clipNames.keys()
                        # Initialize the counter to 1 so our end number will be 1 higher than the number counted
                        cnt = 1
                        # iterate through the list
                        for cl in tmpList:
                            # If we have a match ...
                            if cl.find(clipName) > -1:
                                # ... increment the counter
                                cnt += 1
                        # Add the clipname and counter to the Clip Names dictionary
                        clipNames["%s (%d)" % (clipName, cnt)] = clipNum
                    else:
                        # Add the Clip Name as a Dictionary key pointing to the Clip Number
                        clipNames[clipName] = clipNum
        # If only 1 Clip is found ...
        if len(clipNames) == 1:
            # ... load that clip by looking up the clip's number
            self.parent.KeywordMapLoadClip(clipNames[clipNames.keys()[0]])
            # If left-click, close the Keyword Map.  If not, don't!
            if event.LeftUp():
                # Close the Keyword Map
                self.CloseWindow(event)
        # If more than one Clips are found ..
        elif len(clipNames) > 1:
            # Use a wx.SingleChoiceDialog to allow the user to make the choice between multiple clips here.
            dlg = wx.SingleChoiceDialog(self, _("Which Clip would you like to load?"), _("Select a Clip"),
                                        clipNames.keys(), wx.CHOICEDLG_STYLE)
            # If the user selects a Clip and click OK ...
            if dlg.ShowModal() == wx.ID_OK:
                # ... load the selected clip
                self.parent.KeywordMapLoadClip(clipNames[dlg.GetStringSelection()])
                # Destroy the SingleChoiceDialog
                dlg.Destroy()
                # If left-click, close the Keyword Map.  If not, don't!
                if event.LeftUp():
                    # Close the Keyword Map
                    self.CloseWindow(event)
            # If the user selects Cancel ...
            else:
                # ... destroy the SingleChoiceDialog, but that's all
                dlg.Destroy()
