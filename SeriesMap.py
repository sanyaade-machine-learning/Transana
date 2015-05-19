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

"""This module implements a map of keywords that have been applied to multiple Episodes in a Series"""

__author__ = "David K. Woods <dwoods@wcer.wisc.edu>"

DEBUG = False
if DEBUG:
    print "SeriesMap DEBUG is ON!!"

# import Python's os and sys modules
import os, sys
import string
# load wxPython for GUI
import wx
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

class SeriesMap(wx.Frame):
    """ This is the main class for the Series Map application. """
    def __init__(self, parent, title, seriesNum, seriesName, reportType):
        # reportType 1 is the Sequence Mode, showing relative position of keywords in the Episodes
        # reportType 2 is the Bar Graph mode, showing a bar graph of total time for each keyword
        # reportType 3 is the Percentage mode, showing percentage of total Episode length for each keyword

        # Set the Cursor to the Hourglass while the report is assembled
        TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))

        # It's always important to remember your ancestors.
        self.parent = parent
        # Remember the title
        self.title = title
        # Remember the Report Type
        self.reportType = reportType
        #  Create a connection to the database
        DBConn = DBInterface.get_db()
        #  Create a cursor and execute the appropriate query
        self.DBCursor = DBConn.cursor()
        # Determine the screen size for setting the initial dialog size
        rect = wx.ClientDisplayRect()
        width = rect[2] * .80
        height = rect[3] * .80
        # Create the basic Frame structure with a white background
        self.frame = wx.Frame.__init__(self, parent, -1, title, pos=(10, 10), size=wx.Size(width, height), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL | wx.NO_FULL_REPAINT_ON_RESIZE)
        self.SetBackgroundColour(wx.WHITE)
        # Set the icon
        transanaIcon = wx.Icon(os.path.join(TransanaGlobal.programDir, "images", "Transana.ico"), wx.BITMAP_TYPE_ICO)
        self.SetIcon(transanaIcon)

        # Initialize Media Length to 0
        self.MediaLength = 0
        # Initialize all the data Lists to empty
        self.episodeList = []
        self.filteredEpisodeList = []
        self.clipList = []
        self.clipFilterList = []
        self.unfilteredKeywordList = []
        self.filteredKeywordList = []

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
        # Get the Configuration values for the Series Map Options
        self.barHeight = TransanaGlobal.configData.seriesMapBarHeight
        self.whitespaceHeight = TransanaGlobal.configData.seriesMapWhitespace
        self.hGridLines = TransanaGlobal.configData.seriesMapHorizontalGridLines
        self.vGridLines = TransanaGlobal.configData.seriesMapVerticalGridLines
        self.singleLineDisplay = TransanaGlobal.configData.singleLineDisplay
        self.showLegend = TransanaGlobal.configData.showLegend
        # We default to Color Output.  When this was configurable, if a new Map was
        # created in B & W, the colors never worked right afterwards.
        self.colorOutput = True
        # Get the number of lines per page for multi-page reports
        self.linesPerPage = 66
        # If we have a Series Keyword Sequence Map in multi-line mode ...
        if (self.reportType == 1) and (not self.singleLineDisplay):
            # ... initialize the Episode Name Keyword Lookup Table here.
            self.epNameKWGKWLookup = {}
        # Initialize the Episode Counter, used for vertical placement.
        self.episodeCount= 0
        # We need to be able to look up Episode Lengths for the Bar Graph.  Let's remember them.
        self.episodeLengths = {}

        # Remember the appropriate Episode information
        self.seriesNum = seriesNum
        self.seriesName = seriesName
        # indicate that we're not working from a Clip.  (The Series Maps are never Clip-based.)
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
            self.menuFile.Append(M_FILE_PRINTSETUP, _("Page Setup"), _("Set up Page")) # Add "Printer Setup" to the File Menu
            self.menuFile.Append(M_FILE_PRINTPREVIEW, _("Print Preview"), _("Preview your printed output")) # Add "Print Preview" to the File Menu
            self.menuFile.Enable(M_FILE_PRINTPREVIEW, False)
            self.menuFile.Append(M_FILE_PRINT, _("&Print"), _("Send your output to the Printer")) # Add "Print" to the File Menu
            self.menuFile.Enable(M_FILE_PRINT, False)
            self.menuFile.Append(M_FILE_EXIT, _("E&xit"), _("Exit the Series Map program")) # Add "Exit" to the File Menu
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
        self.Bounds = (5, 5, w - 10, h - 25)

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

        # Populate the drawing
        self.ProcessSeries()
        self.DrawGraph()

        # Restore Cursor to Arrow
        TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

    # Define the Method that implements Filter
    def OnFilter(self, event):
        """ Implement the Filter Dialog call for Series Maps """
        # Set the Cursor to the Hourglass while the report is assembled
        TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
        # Set up parameters for creating the Filter Dialog.  Series Map Filter requires Series Number (as episodeNum) for the Config Save.
        title = string.join([self.title, _("Filter Dialog")], ' ')
        # Series Map wants the Clip Filter
        clipFilter = True
        # Series Map wants Keyword Color customization.
        keywordColors = True
        # We want the Options tab
        options = True
        # reportType=5 indicates it is for a Series Sequence Map.
        # reportType=6 indicates it is for a Series Bar Graph.
        # reportType=7 indicates it is for a Series Percentage Map
        reportType = self.reportType + 4

        # The Series Keyword Sequence Map has all the usual parameters plus Time Range data and the Single Line Display option
        if self.reportType in [1]:
            # Create a Filter Dialog, passing all the necessary parameters.
            dlgFilter = FilterDialog.FilterDialog(self, -1, title, reportType=reportType, configName=self.configName,
                                                  reportScope=self.seriesNum, episodeFilter=True, episodeSort=True,
                                                  clipFilter=clipFilter, keywordFilter=True, keywordSort=True,
                                                  keywordColor=keywordColors, options=options, startTime=self.startTime, endTime=self.endTime,
                                                  barHeight=self.barHeight, whitespace=self.whitespaceHeight, hGridLines=self.hGridLines,
                                                  vGridLines=self.vGridLines, singleLineDisplay=self.singleLineDisplay, showLegend=self.showLegend,
                                                  colorOutput=self.colorOutput)
        elif self.reportType in [2, 3]:
            # Create a Filter Dialog, passing all the necessary parameters.
            dlgFilter = FilterDialog.FilterDialog(self, -1, title, reportType=reportType, configName=self.configName,
                                                  reportScope=self.seriesNum, episodeFilter=True, episodeSort=True,
                                                  clipFilter=clipFilter, keywordFilter=True, keywordSort=True,
                                                  keywordColor=keywordColors, options=options, 
                                                  barHeight=self.barHeight, whitespace=self.whitespaceHeight, hGridLines=self.hGridLines,
                                                  vGridLines=self.vGridLines, showLegend=self.showLegend, colorOutput=self.colorOutput)
        # Sort the Episode List
        self.episodeList.sort()
        # Inform the Filter Dialog of the Episodes
        dlgFilter.SetEpisodes(self.episodeList)
        # If we requested the Clip Filter ...
        if clipFilter:
            # We want the Clips sorted in Clip ID order in the FilterDialog.  We handle that out here, as the Filter Dialog
            # has to deal with manual clip ordering in some instances, though not here, so it can't deal with this.
            self.clipFilterList.sort()
            # Inform the Filter Dialog of the Clips
            dlgFilter.SetClips(self.clipFilterList)
        # Keyword Colors must be specified before Keywords!  So if we want Keyword Colors, ...
        if keywordColors:
            # If we're in grayscale mode, the colors are probably mangled, so let's fix them before
            # we send them to the Filter dialog.
            if not self.colorOutput:
                # A shallow copy of the dictionary object should get the job done.
                self.keywordColors = self.rememberedKeywordColors.copy()
            # Inform the Filter Dialog of the colors used for each Keyword
            dlgFilter.SetKeywordColors(self.keywordColors)
        # Inform the Filter Dialog of the Keywords
        dlgFilter.SetKeywords(self.unfilteredKeywordList)
        # Set the Cursor to the Arrow now that the filter dialog is assembled
        TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
        # Create a dummy error message to get our while loop started.
        errorMsg = 'Start Loop'
        # Keep trying as long as there is an error message
        while errorMsg != '':
            # Clear the last (or dummy) error message.
            errorMsg = ''
            # Show the Filter Dialog and see if the user clicks OK
            if dlgFilter.ShowModal() == wx.ID_OK:
                # Get the Episode Data from the Filter Dialog
                self.episodeList = dlgFilter.GetEpisodes()
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
                    # Only the Series Keyword Sequence Map needs the Time Range options.
                    if self.reportType in [1]:
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

                    # Get the Bar Height and Whitespace Height for all versions of the Series Map
                    self.barHeight = dlgFilter.GetBarHeight()
                    self.whitespaceHeight = dlgFilter.GetWhitespace()
                    # we need to store the Bar Height and Whitespace values in the Configuration.
                    TransanaGlobal.configData.seriesMapBarHeight = self.barHeight
                    TransanaGlobal.configData.seriesMapWhitespace = self.whitespaceHeight

                    # Get the Grid Line data from the form
                    self.hGridLines = dlgFilter.GetHGridLines()
                    self.vGridLines = dlgFilter.GetVGridLines()
                    # Store the Grid Line data in the Configuration
                    TransanaGlobal.configData.seriesMapHorizontalGridLines = self.hGridLines
                    TransanaGlobal.configData.seriesMapVerticalGridLines = self.vGridLines

                    # Only the Series Keyword Sequence Graph needs the Single Line Display Option data.
                    if self.reportType in [1]:
                        # Get the singleLineDisplay value from the dialog
                        self.singleLineDisplay = dlgFilter.GetSingleLineDisplay()
                        # Remember the value.
                        TransanaGlobal.configData.singleLineDisplay = self.singleLineDisplay
                        
                    # Get the showLegend value from the dialog
                    self.showLegend = dlgFilter.GetShowLegend()
                    # Remember the value.  (This doesn't get saved.)
                    TransanaGlobal.configData.showLegend = self.showLegend

                    # Detect if the colorOutput value is actually changing.
                    if (self.colorOutput != dlgFilter.GetColorOutput()):
                        # If we're going from color to grayscale ...
                        if self.colorOutput:
                            # ... remember what the colors were before they get all screwed up by displaying
                            # the graphic without them.
                            self.rememberedKeywordColors = {}
                            self.rememberedKeywordColors = self.keywordColors.copy()
                            
                    # Get the colorOutput value from the dialog 
                    self.colorOutput = dlgFilter.GetColorOutput()

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
        lineHeight = self.CalcY(1) - self.CalcY(0)
        printout = MyPrintout(self.title, self.graphic, multiPage=True, lineStart=self.CalcY(0) - int(lineHeight / 2.0), lineHeight=lineHeight)
        printout2 = MyPrintout(self.title, self.graphic, multiPage=True, lineStart=self.CalcY(0) - int(lineHeight / 2.0), lineHeight=lineHeight)
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
        lineHeight = self.CalcY(1) - self.CalcY(0)
        printout = MyPrintout(self.title, self.graphic, multiPage=True, lineStart=self.CalcY(0) - int(lineHeight / 2.0), lineHeight=lineHeight)
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
        HelpContext = "Series Keyword Graphs"
        # If a Help Window is defined ...
        if TransanaGlobal.menuWindow != None:
            # ... call Help!
            TransanaGlobal.menuWindow.ControlObject.Help(HelpContext)

    def OnSize(self, event):
        """ Handle Resize Events by resizing the Graphic Control and redrawing the graphic """
        (w, h) = self.GetClientSizeTuple()
        if self.Bounds[1] == 5:
            self.Bounds = (5, 5, w - 10, h - 25)
        else:
            self.Bounds = (5, 40, w - 10, h - 30)
        self.DrawGraph()
            
    def CalcX(self, XPos):
        """ Determine the proper horizontal coordinate for the given time """
        # We need to start by defining the legal range for the type of graph we're working with.
        # The Sequence Map is tied to the start and end time variables.
        if self.reportType == 1:
            startVal = self.startTime
            endVal = self.endTime
        # The Bar Graph stretches from 0 to the time line Maximum variable
        elif self.reportType == 2:
            startVal = 0.0
            if self.timelineMax == 0:
                endVal = 1
            else:
                endVal = self.timelineMax
        # The Percentage Graph ranges from 0 to 100!
        elif self.reportType == 3:
            startVal = 0.0
            endVal = 100.0
        # Specify a margin width
        marginwidth = (0.06 * (self.Bounds[2] - self.Bounds[0]))
        # The Horizonal Adjustment is the global graphic indent
        hadjust = self.graphicindent
        # The Scaling Factor is the active portion of the drawing area width divided by the total media length
        # The idea is to leave the left margin, self.graphicindent for Keyword Labels, and the right margin
        if self.MediaLength > 0:
            scale = (float(self.Bounds[2]) - self.Bounds[0] - hadjust - 2 * marginwidth) / (endVal - startVal)
        else:
            scale = 0.0
        # The horizontal coordinate is the left margin plus the Horizontal Adjustment for Keyword Labels plus
        # position times the scaling factor
        res = marginwidth + hadjust + ((XPos - startVal) * scale)
        return int(res)

    def FindTime(self, x):
        """ Given a horizontal pixel position, determine the corresponding time value from
            the video time line """
        # determine the margin width
        marginwidth = (0.06 * (self.Bounds[2] - self.Bounds[0]))
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
        # Spacing is the larger of (12 pixels for label text or the bar height) plus 2 for whitespace
        spacing = max(12, self.barHeight) + self.whitespaceHeight
        # Top margin is 30 for titles plus 28 for the timeline
        topMargin = 30 + (2 * spacing)
        return int(spacing * YPos + topMargin)

    def FindKeyword(self, y):
        """ Given a vertical pixel position, determine the corresponding Keyword data """

        # NOTE:  This method is only valid if self.reportType == 1, the Sequence Map.
        #        Other variations of the Series maps may use different key values for the dictionary.
        if self.reportType != 1:
            return None
        
        # If the graphic is scrolled, the raw Y value does not point to the correct Keyword.
        # Determine the unscrolled equivalent Y position.
        (modX, modY) = self.graphic.CalcUnscrolledPosition(0, y)
        # Now we need to get the keys for the Lookup Dictionary
        keyVals = self.epNameKWGKWLookup.keys()
        # We need the keys to be in order, so we can quit when we've found what we're looking for.
        keyVals.sort()

        # The single-line display and the multi-line display handle the lookup differently, of course.
        # Let's start with the single-line display.
        if self.singleLineDisplay:

            # Initialize the return value to None in case nothing is found.  The single-line version expects an Episode Name.
            returnVal = None
            # We also need a temporary value initialized to None.  Our data structure returns complex data, from which we
            # extract the desired value.
            tempVal = None

            # Iterate through the sorted keys.  The keys are actually y values for the graph!
            for yVal in keyVals:
                # If we find a key value that is smaller than the unscrolled Graphic y position ...
                if yVal <= modY:
                    # ... then we've found a candidate for what we're looking for.  But we keep iterating,
                    # because we want the LARGEST yVal that's smaller than the graphic y value.
                    tempVal = self.epNameKWGKWLookup[yVal]
                # Once our y values are too large ...
                else:
                    # ... we should stop iterating through the (sorted) keys.
                    break

            # If we found a valid data structure ...
            if tempVal != None:
                # ... we can extract the Episode name by looking at the first value of the first value of the first key.
                returnVal = tempVal[tempVal.keys()[0]][0][0]

        # Here, we handle the multi-line display of the Sequence Map.
        else:
            # Initialize the return value to a tuple of three Nones in case nothing is found.
            # The multi-line version expects an Episode Name, Keyword Group, Keyword tuple.
            returnVal = (None, None, None)
            
            # Iterate through the sorted keys.  The keys are actually y values for the graph!
            for yVal in keyVals:
                # If we find a key value that is smaller than the unscrolled Graphic y position ...
                if yVal <= modY:
                    # ... then we've found a candidate for what we're looking for.  But we keep iterating,
                    # because we want the LARGEST yVal that's smaller than the graphic y value.
                    returnVal = self.epNameKWGKWLookup[yVal]
                # Once our y values are too large ...
                else:
                    # ... we should stop iterating through the (sorted) keys.
                    break
        # Return the value we found, or None
        return returnVal

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

    def ProcessSeries(self):

        # Initialize Media Length to 0
        self.MediaLength = 0
        # Initialize all the data Lists to empty
        self.episodeList = []
        self.filteredEpisodeList = []
        self.clipList = []
        self.clipFilterList = []
        self.unfilteredKeywordList = []
        self.filteredKeywordList = []

        if self.reportType == 2:
            epLengths = {}

        # Get Series Number, Episode Number, Media File Name, and Length
        SQLText = """SELECT e.EpisodeNum, e.EpisodeID, e.SeriesNum, e.MediaFile, e.EpLength, s.SeriesID
                       FROM Episodes2 e, Series2 s
                       WHERE s.SeriesNum = e.SeriesNum AND
                             s.SeriesNum = %s
                       ORDER BY EpisodeID """

        self.DBCursor.execute(SQLText, self.seriesNum)

        for (EpisodeNum, EpisodeID, SeriesNum, MediaFile, EpisodeLength, SeriesID) in self.DBCursor.fetchall():
            EpisodeID = DBInterface.ProcessDBDataForUTF8Encoding(EpisodeID)
            SeriesID = DBInterface.ProcessDBDataForUTF8Encoding(SeriesID)
            MediaFile = DBInterface.ProcessDBDataForUTF8Encoding(MediaFile)

            self.episodeList.append((EpisodeID, SeriesID, True))

            if (EpisodeLength > self.MediaLength):
                self.MediaLength = EpisodeLength
                self.endTime = self.MediaLength

            # Remember the Episode's length
            self.episodeLengths[(EpisodeID, SeriesID)] = EpisodeLength

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
                if not (kwg, kw) in self.filteredKeywordList:
                    self.filteredKeywordList.append((kwg, kw))
                if not (kwg, kw, True) in self.unfilteredKeywordList:
                    self.unfilteredKeywordList.append((kwg, kw, True))

            # Sort the Keyword List
            self.unfilteredKeywordList.sort()

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
                    self.clipList.append((kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum, EpisodeID, SeriesID))

                    if not ((clipID, collectNum, True) in self.clipFilterList):
                        self.clipFilterList.append((clipID, collectNum, True))

        # Sort the Keyword List
        self.filteredKeywordList.sort()

    def UpdateKeywordVisualization(self):
        """ Update the Keyword Visualization following something that could have changed it. """
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
        SQLText = """SELECT ck.KeywordGroup, ck.Keyword, cl.ClipStart, cl.ClipStop, cl.ClipNum, cl.ClipID, cl.CollectNum, ep.EpisodeName
                       FROM Clips2 cl, ClipKeywords2 ck, Episodes2 ep
                       WHERE cl.EpisodeNum = %s AND
                             cl.ClipNum = ck.ClipNum AND
                             ep.EpisodeNum = cl.EpisodeNum
                       ORDER BY ClipStart, ClipNum, KeywordGroup, Keyword"""
        # Execute the query
        self.DBCursor.execute(SQLText, self.episodeNum)
        # Iterate through the results ...
        for (kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum, episodeName) in self.DBCursor.fetchall():
            kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
            kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
            clipID = DBInterface.ProcessDBDataForUTF8Encoding(clipID)
            episodeName = DBInterface.ProcessDBDataForUTF8Encoding(episodeName)
            # If we're dealing with an Episode, self.clipNum will be None and we want all clips.
            # If we're dealing with a Clip, we only want to deal with THIS clip!
            if (self.clipNum == None) or (clipNum == self.clipNum):
                # If a Clip is not found in the clipList ...
                if not ((kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum, episodeName, seriesName) in self.clipList):
                    # ... add it to the clipList ...
                    self.clipList.append((kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum, episodeName, seriesName))
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
        """ Actually Draw the Series Map """
        self.keywordClipList = {}

        # Series Keyword Sequence Map, if multi-line display is desired
        if (self.reportType == 1) and (not self.singleLineDisplay):
            epCount = 0
            for (episodeName, seriesName, checked) in self.episodeList:
                if checked:
                    epCount += 1
            # Determine the graphic size needed for the number of episodes times the number of keywords plus two lines
            # for each episode for the episode title and the blank line!
            newheight = max(self.CalcY(epCount * (len(self.filteredKeywordList) + 2)), self.Bounds[3] - self.Bounds[1])

        # Series Keyword Sequence Map's single-line display,
        # Series Keyword Bar Graph, and Series Keyword Percentage Graph all need the data arranged the same way
        else:
            # Initialize a dictionary that will hold information about the bars we're drawing.
            barData = {}
            # We need to know how many Episodes we have on the graph.  Initialize a counter
            self.episodeCount = 0
            for (episodeName, seriesName, checkVal) in self.episodeList:
                if checkVal:
                    # Count all the Episodes that have been "checked" in the Filter.
                    self.episodeCount += 1

            # Now we iterate through the CLIPS.
            for (KWG, KW, Start, Stop, ClipNum, ClipName, CollectNum, episodeName, seriesName) in self.clipList:
                # We make sure they are selected in the Filter, checking the Episode, Clips and Keyword selections
                if ((episodeName, seriesName, True) in self.episodeList) and \
                   ((ClipName, CollectNum, True) in self.clipFilterList) and \
                   ((KWG, KW) in self.filteredKeywordList):
                    # Now we track the start and end times compared to the current display limits
                    if Start < self.startTime:
                        Start = self.startTime
                    if Start > self.endTime:
                        Start = self.endTime
                    if Stop > self.endTime:
                        Stop = self.endTime
                    if Stop < self.startTime:
                        Stop = self.startTime

                    # Set up the key we use to mark overlaps
                    overlapKey = (episodeName, KWG, KW)

                    # If Start and Stop are the same, the Clip is off the graph and should be ignored.
                    if Start != Stop:
                        # If the clip is ON the graph, let's check for overlap with other clips with the same keyword at the same spot
                        if not barData.has_key(overlapKey):
                            barData[overlapKey] = 0
                        # Add the bar length to the bar Data dictionary.
                        barData[overlapKey] += Stop - Start

                    # For the Series Keyword Bar Graph and the Series Keyword Percentage Graph ...
                    if self.reportType in [2, 3]:
                        # Now add the Clip to the keywordClipList.  This holds all Keyword/Clip data in memory so it can be searched quickly
                        # This dictionary object uses the episode name and keyword pair as the key and holds a list of Clip data for all clips with that keyword.
                        # If the list for a given keyword already exists ...
                        if self.keywordClipList.has_key(overlapKey):
                            # Let's look for overlap.  Overlap artificially inflates the size of the bars, and must be eliminated.
                            overlapStart = Stop
                            overlapEnd = Start
                            # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                            overlapClips = self.keywordClipList[overlapKey]
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
                                # We need to SUBTRACT the overlap time from the barData structure.
                                barData[overlapKey] -= overlapEnd - overlapStart

                                if DEBUG:
                                    print "Bar Graph overlap found:", overlapKey, overlapEnd - overlapStart

                            # ... add the new Clip to the Clip List
                            self.keywordClipList[overlapKey].append((Start, Stop, ClipNum, ClipName))
                        # If there is no entry for the given keyword ...
                        else:
                            # ... create a List object with the first clip's data for this Keyword Pair key
                            self.keywordClipList[overlapKey] = [(Start, Stop, ClipNum, ClipName)]

            # once we're done with checking overlaps here, let's clear out this variable,
            # as it may get re-used later for other purposes!
            self.keywordClipList = {}
            if self.showLegend:
                newheight = max(self.CalcY(self.episodeCount + len(self.filteredKeywordList) + 2), self.Bounds[3] - self.Bounds[1])
            else:
                newheight = max(self.CalcY(self.episodeCount), self.Bounds[3] - self.Bounds[1])

        # Now that we have all necessary information, let's create and populate the graphic
        # Start by destroying the existing control and creating a new one with the correct Canvas Size
        self.graphic.Destroy()
        self.graphic = GraphicsControlClass.GraphicsControl(self, -1, wx.Point(self.Bounds[0], self.Bounds[1]),
                                                            (self.Bounds[2] - self.Bounds[0], self.Bounds[3] - self.Bounds[1]),
                                                            (self.Bounds[2] - self.Bounds[0], newheight + 3),
                                                            passMouseEvents=True)
        # Put the header information on the graphic.
        self.graphic.SetFontColour("BLACK")
        if 'wxMac' in wx.PlatformInfo:
            self.graphic.SetFontSize(17)
        else:
            self.graphic.SetFontSize(14)
        self.graphic.AddTextCentered("%s" % self.title, (self.Bounds[2] - self.Bounds[0]) / 2, 1)
        if 'wxMac' in wx.PlatformInfo:
            self.graphic.SetFontSize(13)
        else:
            self.graphic.SetFontSize(10)
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Series: %s'), 'utf8')
        else:
            prompt = _('Series: %s')
        self.graphic.AddText(prompt % self.seriesName, 2, 2)
        if self.configName != '':
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Filter Configuration: %s'), 'utf8')
            else:
                prompt = _('Filter Configuration: %s')
            self.graphic.AddText(prompt % self.configName, 2, 16)

        # Initialize a Line Counter, used for vertical positioning
        Count = 0
        # We'll also need a lookup table for vertical values.
        yValLookup = {}
        # Initialize the Episode Name / Keyword Lookup table.  The multi-line Series Keyword Sequence Map gets a blank first line.
        if (self.reportType == 1) and (not self.singleLineDisplay):
            self.epNameKWGKWLookup = {0 : ('', '', '')}
        else:
            self.epNameKWGKWLookup = {}

        # Now iterate through the Episode list, adding the Episode Names and (if appropriate) the Keywords as an axis label
        for (episodeName, seriesName, episodeShown) in self.episodeList:
            if episodeShown:
                # Add the Episode Name to the vertical axis
                self.graphic.AddText("%s" % episodeName, 4, self.CalcY(Count) - 7)
                # if Keyword Series Sequence Map in multi-line mode ...
                if (self.reportType == 1) and (not self.singleLineDisplay):
                    # ... add a blank lookup line for the blank line, as this line gets no data for that report.
                    self.epNameKWGKWLookup[self.CalcY(Count-1) - int((self.barHeight + self.whitespaceHeight)/2)] = ('', '', '')
                # We want Grid Lines in light gray
                self.graphic.SetColour('LIGHT GREY')
                # if Keyword Series Sequence Map in multi-line mode, we draw Grid Lines and add Keywords to the Vertical Axis.
                if (self.reportType == 1) and (not self.singleLineDisplay):
                    # Draw the top Grid Line, if appropriate
                    if self.hGridLines:
                        self.graphic.AddLines([(10, self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2), self.CalcX(self.endTime), self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2))])
                        gridLineCount = Count
                    Count += 1
                    # Iterate through the Keyword List from the Filter Dialog ...
                    for KWG, KW in self.filteredKeywordList:
                        # ... and add the Keywords to the Vertical Axis.
                        self.graphic.AddText("%s : %s" % (KWG, KW), 10, self.CalcY(Count) - 7)
                        # Add this data to the Y Position Lookup dictionary.
                        yValLookup[(episodeName, KWG, KW)] = Count
                        # Add a Lookup Line for this episodeName, Keyword Group, Keyword combination
                        self.epNameKWGKWLookup[self.CalcY(Count) - int((self.barHeight + self.whitespaceHeight)/2)] = (episodeName, KWG, KW)
                        # Add Horizontal Grid Lines, if appropriate
                        if self.hGridLines and ((Count - gridLineCount) % 2 == 0):
                            self.graphic.AddLines([(10, self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2), self.CalcX(self.endTime), self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2))])
                        # Increment the counter for each Keyword
                        Count = Count + 1
                # If it's NOT the multi-line Sequence Map, the Gridline rules are different, but still need to be handled.
                else:
                    # Add Horizontal Grid Lines, if appropriate
                    if self.hGridLines and (Count % 2 == 1):
                        self.graphic.AddLines([(4, self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2), self.CalcX(self.timelineMax), self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2))])
                    # Add this data to the Y Position Lookup dictionary.
                    yValLookup[episodeName] = Count
                # Increment the counter for each Episode.  (This produces a blank line in the Sequence Map, which is OK.)
                Count += 1

        # If multi-line Sequence Report, we're building the Episode Name / Keyword Lookup table here, otherwise it's later.                    
        if (self.reportType == 1) and (not self.singleLineDisplay):
            # Finish with a blank Lookup Line so the bottom of the chart doesn't give false positive information
            self.epNameKWGKWLookup[self.CalcY(Count-1) - int((self.barHeight + self.whitespaceHeight)/2)] = ('', '', '')
                
        # Reset the graphic color following drawing the Grid Lines
        self.graphic.SetColour("BLACK")
        # After we have the axis values specified but before we draw anything else, we determine the amount the
        # subsequent graphics must be indented to adjust for the size of the text labels.
        self.graphicindent = self.graphic.GetMaxWidth(start=3)

        # Draw the Graph Time Line
        # For the Sequence Map, the timeline is startTime to endTime.
        if (self.reportType == 1):
            # If the Media Length is known, display the Time Line
            if self.MediaLength > 0:
                self.DrawTimeLine(self.startTime, self.endTime)
            # For the Sequence Map, we need to know the maximum Episode time, which is already stored under self.endTime.
            self.timelineMax = self.endTime
        # For the Series Keyword Bar Graph and the Series Keyword Percentage Graph, we need to know the maximum coded
        # time and the episode length for each Episode.
        # For the Bar Graph, we use the longer of Episode Length or Total Episode Coded Time.
        # For the Percentage Graph, we need to know total amount of coded video for each Episode
        elif self.reportType in [2, 3]:
            # Initialize the time line maximum variable
            self.timelineMax = 0
            # Create a dictionary to store the episode times.
            episodeTimeTotals = {}
            # Start by iterating through the Episode List ...
            for (episodeName, seriesName, checked) in self.episodeList:
                if checked:
                    # Initialize the Episode's length to 0
                    episodeTimeTotals[episodeName] = 0
                    # Iterate through the Keyword List
                    for (kwg, kw) in self.filteredKeywordList:
                        # Check to see if we have data for this keyword in this Episode.
                        if barData.has_key((episodeName, kwg, kw)):
                            # If so, add the time to the Episode's total time.
                            episodeTimeTotals[episodeName] += barData[(episodeName, kwg, kw)]
                            # If this Episode is the longest we've dealt with so far ...
                            if episodeTimeTotals[episodeName] > self.timelineMax:
                                # ... note the new time line maximum.
                                self.timelineMax = episodeTimeTotals[episodeName]
                    # If we are building the Bar Graph, ...
                    if self.reportType == 2:
                        # ... we need to adjust the timelineMax value for the length of the whole Episode, if it's larger.
                        self.timelineMax = max(self.timelineMax, self.episodeLengths[(episodeName, seriesName)])
                    
            # The Series Keyword Bar Graph extends from 0 to the timeLineMax value we just determined.
            if self.reportType == 2:
                self.DrawTimeLine(0, self.timelineMax)
            # The Series Keyword Percentage Graph extends from 0% to 100%
            elif (self.reportType == 3):
                self.DrawTimeLine(0, 100)
            # Add the top Horizontal Grid Line, if appropriate
            if self.hGridLines:
                # We want Grid Lines in light gray
                self.graphic.SetColour('LIGHT GREY')
                self.graphic.AddLines([(4, self.CalcY(-1) + 6 + int(self.whitespaceHeight / 2), self.CalcX(self.timelineMax), self.CalcY(-1) + 6 + int(self.whitespaceHeight / 2))])

        # Select the color palate for colors or gray scale as appropriate
        if self.colorOutput:
            colorSet = TransanaConstants.keywordMapColourSet
            colorLookup = TransanaConstants.transana_colorLookup
        else:
            colorSet = TransanaConstants.keywordMapGraySet
            colorLookup = TransanaConstants.transana_grayLookup

        # Set the colourIndex tracker to the last color used.
        colourindex = self.keywordColors['lastColor']

        # Iterate through the keyword list in order ...
        for (KWG, KW) in self.filteredKeywordList:
            # Check to see if the color for this keyword has already been assigned, and if we're using Colors.
            if self.keywordColors.has_key((KWG, KW)) and self.colorOutput:
                # If so, set the index to match the current keyword's assigned color
                colourindex = self.keywordColors[(KWG, KW)]
            # If no color has been assigned, or if we're in GrayScale mode ...
            else:
                # ... assign the next available color to the index
                colourindex = self.keywordColors['lastColor'] + 1
                # If we've exceeded the number of available colors (or shades of gray) ...
                if colourindex > len(colorSet) - 1:
                    # ... that start over again at the beginning of the list.
                    colourindex = 0
                # Update the dictionary element that tracks the last new color assigned.
                self.keywordColors['lastColor'] = colourindex
                # Remember that this color is now associated with this keyword for future reference
                self.keywordColors[(KWG, KW)] = colourindex

        # If we're producing a Series Keyword Sequence Map ..
        if (self.reportType == 1):
            # some clip boundary lines for overlapping clips can get over-written, depeding on the nature of the overlaps.
            # Let's create a separate list of these lines, which we'll add to the END of the process so they can't get overwritten.
            overlapLines = []
            # Iterate through all the Clip/Keyword records in the Clip List ...
            for (KWG, KW, Start, Stop, ClipNum, ClipName, CollectNum, episodeName, seriesName) in self.clipList:
                # Check the clip against the Episode List, the Clip Filter List, and the Keyword Filter list to see if
                # it should be included in the report.
                if ((episodeName, seriesName, True) in self.episodeList) and \
                   ((ClipName, CollectNum, True) in self.clipFilterList) and \
                   ((KWG, KW) in self.filteredKeywordList):

                    # We compare the Clip's Start Time with the Map's boundaries.  We only want the portion of the clip
                    # that falls within the Map's upper and lower boundaries.
                    if Start < self.startTime:
                        Start = self.startTime
                    if Start > self.endTime:
                        Start = self.endTime
                    if Stop > self.endTime:
                        Stop = self.endTime
                    if Stop < self.startTime:
                        Stop = self.startTime

                    # If Start and Stop match, the clip is off the Map and can be ignored.  Otherwise ...
                    if Start != Stop:
                        # ... we start drawing the clip's bar by setting the bar thickness.
                        self.graphic.SetThickness(self.barHeight)
                        # Initialize a variable for building the line's data record
                        tempLine = []
                        # Determine the vertical placement of the line, which requires a different lookup key for the
                        # single-line report than the multi-line report.
                        if self.singleLineDisplay:
                            yPos = self.CalcY(yValLookup[episodeName])
                        else:
                            yPos = self.CalcY(yValLookup[(episodeName, KWG, KW)])
                        # Add the line data
                        tempLine.append((self.CalcX(Start), yPos, self.CalcX(Stop), yPos))
                        # Determine the appropriate color for the keyword
                        colourindex = self.keywordColors[(KWG, KW)]
                        # Tell the graph to use the selected color, using the appropriate lookup table
                        self.graphic.SetColour(colorLookup[colorSet[colourindex]])
                        # Add the line data to the graph
                        self.graphic.AddLines(tempLine)

                        # We need to track the bar positions so that the MouseOver can display data correctly.  We need to do it
                        # later for the multi-line report, but here for the single-line report.
                        if self.singleLineDisplay:
                            # The first stage of the lookup is the Y-coordinate.  If there's not already an
                            # EpisodeNameKeywordGroupKeywordLookup record for this Y-Coordinate ...
                            if not self.epNameKWGKWLookup.has_key(self.CalcY(yValLookup[episodeName]) - int((self.barHeight + self.whitespaceHeight)/2)):
                                # ... create an empty dictionary object for the first part of the Lookup Line
                                self.epNameKWGKWLookup[self.CalcY(yValLookup[episodeName]) - int((self.barHeight + self.whitespaceHeight)/2)] = {}
                            # The second stage of the lookup is the X range in a tuple.  If the X range isn't already in the dictionary,
                            # then add an empty List object for the X range.
                            if not self.epNameKWGKWLookup[self.CalcY(yValLookup[episodeName]) - int((self.barHeight + self.whitespaceHeight)/2)].has_key((self.CalcX(Start), self.CalcX(Stop))):
                                self.epNameKWGKWLookup[self.CalcY(yValLookup[episodeName]) - int((self.barHeight + self.whitespaceHeight)/2)][(self.CalcX(Start), self.CalcX(Stop))] = []
                            # Add a Lookup Line for this Y-coordinate and X range containing the Episode Name, the keyword data,
                            # and the Clip Length.
                            self.epNameKWGKWLookup[self.CalcY(yValLookup[episodeName]) - int((self.barHeight + self.whitespaceHeight)/2)][(self.CalcX(Start), self.CalcX(Stop))].append((episodeName, KWG, KW, Stop - Start))

                    if DEBUG and KWG == 'Transana Users' and KW == 'DavidW':
                        print "Looking at %s (%d)" % (ClipName, CollectNum)

                    # We need to indicate where there is overlap in this map.

                    # We use a different key to mark overlaps depending on whether we're in singleLineDisplay mode or not.
                    if self.singleLineDisplay:
                        overlapKey = (episodeName)
                    else:
                        overlapKey = (episodeName, KWG, KW)
                                
                    # Now add the Clip to the keywordClipList.  This holds all Keyword/Clip data in memory so it can be searched quickly
                    # This dictionary object uses the keyword pair as the key and holds a list of Clip data for all clips with that keyword.
                    # If the list for a given keyword already exists ...
                    if self.keywordClipList.has_key(overlapKey):
                        # Let's look for overlap
                        overlapStart = Stop
                        overlapEnd = Start
                        # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                        overlapClips = self.keywordClipList[overlapKey]
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
                            tempLine = [(self.CalcX(overlapStart), yPos, self.CalcX(overlapEnd), yPos)]
                            self.graphic.AddLines(tempLine)
                            if self.colorOutput:
                                self.graphic.SetColour("RED")
                            else:
                                self.graphic.SetColour("BLACK")
                            tempLine = [(self.CalcX(overlapStart), yPos - overlapThickness+1, self.CalcX(overlapEnd), yPos - overlapThickness+1)]
                            self.graphic.AddLines(tempLine)
                            if self.colorOutput:
                                self.graphic.SetColour("BLUE")
                            else:
                                self.graphic.SetColour("GRAY")
                            tempLine = [(self.CalcX(overlapStart), yPos + overlapThickness, self.CalcX(overlapEnd), yPos + overlapThickness)]
                            self.graphic.AddLines(tempLine)
                            # Let's remember the clip start and stop boundaries, to be drawn at the end so they won't get over-written
                            overlapLines.append(((self.CalcX(overlapStart), yPos - (self.barHeight / 2), self.CalcX(overlapStart), yPos + (self.barHeight / 2)),))
                            overlapLines.append(((self.CalcX(overlapEnd), yPos - (self.barHeight / 2), self.CalcX(overlapEnd), yPos + (self.barHeight / 2)),))

                        # ... add the new Clip to the Clip List
                        self.keywordClipList[overlapKey].append((Start, Stop, ClipNum, ClipName))
                    # If there is no entry for the given keyword ...
                    else:
                        # ... create a List object with the first clip's data for this Keyword Pair key.
                        self.keywordClipList[overlapKey] = [(Start, Stop, ClipNum, ClipName)]

            # For the single-line display only ...
            if self.singleLineDisplay:
                # ... finish with a blank Lookup Line so the bottom of the chart doesn't give false positive information
                self.epNameKWGKWLookup[self.CalcY(self.episodeCount) - int((self.barHeight + self.whitespaceHeight)/2)] = {}
                self.epNameKWGKWLookup[self.CalcY(self.episodeCount) - int((self.barHeight + self.whitespaceHeight)/2)][(0, self.timelineMax)] = [('', '', '', 0)]

            # let's add the overlap boundary lines now
            self.graphic.SetThickness(1)
            self.graphic.SetColour("BLACK")
            for tempLine in overlapLines:
                self.graphic.AddLines(tempLine)

            if not '__WXMAC__' in wx.PlatformInfo:
                self.menuFile.Enable(M_FILE_SAVEAS, True)
                self.menuFile.Enable(M_FILE_PRINTPREVIEW, True)
                self.menuFile.Enable(M_FILE_PRINT, True)

        # For the Series Keyword Bar Graph  and the Series keyword Percentage Graph, which are VERY similar and therefore use the same
        # infrastructure ...
        elif self.reportType in [2, 3]:
            # ... we first iterate through all the Episodes in the Episode List ...
            for (episodeName, seriesName, checked) in self.episodeList:
                # .. and check to see if the Episode should be included.
                if checked:
                    # These graphs are cumulative bar charts.  We need to track the starting place for the next bar.
                    barStart = 0
                    # Create the first part of the Lookup Line, an empty dictionary for the Y coordinate
                    self.epNameKWGKWLookup[self.CalcY(yValLookup[episodeName]) - int((self.barHeight + self.whitespaceHeight)/2)] = {}
                    # Now we iterate through the Filtered Keyword List.  (This gives us the correct presentation ORDER for the bars!)
                    for (kwg, kw) in self.filteredKeywordList:
                        # Now we check to see if there's DATA for this Episode / Keyword combination.
                        if barData.has_key((episodeName, kwg, kw)):
                            # Start by setting the bar thickness
                            self.graphic.SetThickness(self.barHeight)
                            # Initialize a temporary list for accumulating Bar data (not really necessary with this structure, but no harm done.)
                            tempLine = []
                            # If we're drawing the Series Keyword Bar Graph ...
                            if self.reportType == 2:
                                # ... the bar starts at the unadjusted BarStart position ...
                                xStart = self.CalcX(barStart)
                                # ... and ends at the start plus the width of the bar!
                                xEnd = self.CalcX(barStart + barData[(episodeName, kwg, kw)])
                                # The mouseover for this report is the unadjusted length of the bar
                                lookupVal = barData[(episodeName, kwg, kw)]
                            # If we're drawing the Series Keyword Percentage Graph ...
                            elif self.reportType == 3:
                                # This should just be a matter of adjusting barData for episodeTimeTotals[episodeName], which is the total
                                # coded time for each Episode.
                                # ... the bar starts at the adjusted BarStart position ...
                                xStart = self.CalcX(barStart * 100.0 / episodeTimeTotals[episodeName])
                                # ... and ends at the adjusted (start plus the width of the bar)!
                                xEnd = self.CalcX((barStart + barData[(episodeName, kwg, kw)]) * 100.0 / episodeTimeTotals[episodeName])
                                # The mouseover for this report is the adjusted length of the bar, which is the percentage value for the bar!
                                lookupVal = barData[(episodeName, kwg, kw)] * 100.0 / episodeTimeTotals[episodeName]
                            # Build the line to be displayed based on these calculated values
                            tempLine.append((xStart, self.CalcY(yValLookup[episodeName]), xEnd, self.CalcY(yValLookup[episodeName])))
                            # Determine the index for this Keyword's Color
                            colourindex = self.keywordColors[(kwg, kw)]
                            # Tell the graph to use the selected color, using the appropriate lookup table
                            self.graphic.SetColour(colorLookup[colorSet[colourindex]])
                            # Actually add the line to the graph's data structure
                            self.graphic.AddLines(tempLine)
                            # Add a Lookup Line for this Y-coordinate and X range containing the Episode Name, the keyword data,
                            # and the Clip 's Lookup Value determined above.  Note that this is a bit simpler than for the Sequence Map
                            # because we don't have to worry about overlaps.  Thus, the lookup value can just be a tuple instead of having
                            # to be a list of tuples to accomodate overlapping clip/keyword values.
                            self.epNameKWGKWLookup[self.CalcY(yValLookup[episodeName]) - int((self.barHeight + self.whitespaceHeight)/2)][(xStart, xEnd)] = (episodeName, kwg, kw, lookupVal)
                            # The next bar should start where this bar ends.  No need to adjust for the Percentage Graph -- that's handled
                            # when actually placing the bars.
                            barStart += barData[(episodeName, kwg, kw)]

            # Finish with a blank Lookup Line so the bottom of the chart doesn't give false positive information
            self.epNameKWGKWLookup[self.CalcY(self.episodeCount) - int((self.barHeight + self.whitespaceHeight)/2)] = {}
            self.epNameKWGKWLookup[self.CalcY(self.episodeCount) - int((self.barHeight + self.whitespaceHeight)/2)][(0, self.timelineMax)] = ('', '', '', 0)

        # Enable tracking of mouse movement over the graphic
        self.graphic.Bind(wx.EVT_MOTION, self.OnMouseMotion)

        # Add Legend.  The multi-line Series Keyword Sequence Map doesn't get a legend, nor does any report where the showLegend option
        # is turned off.
        if (((self.reportType == 1) and self.singleLineDisplay) or (self.reportType in [2, 3])) and self.showLegend:
            # Skip two lines from the bottom of the report.
            Count +=2
            # Let's place the legend at 1/3 of the way across the report horizontally.
            startX = int((self.Bounds[2] - self.Bounds[0]) / 3.0)
            # Let's place the legend below the report content.
            startY = self.CalcY(Count)
            # To draw a box around the legend, we'll need to track it's end coordinates too.
            endX = startX
            endY = startY

            # For GetTextExtent to work right, we have to make sure the font is set in the graphic context.
            # First, define a font for the current font settings
            font = wx.Font(self.graphic.fontsize, self.graphic.fontfamily, self.graphic.fontstyle, self.graphic.fontweight)
            # Set the font for the graphics context
            self.graphic.SetFont(font)
            # Add a label for the legend
            self.graphic.AddText(_("Legend:"), startX, self.CalcY(Count - 1) - 7)
            endX = startX + 14 + self.graphic.GetTextExtent(_("Legend:"))[0]
            # We'll use a 14 x 12 block to show color.  Set the line thickness
            self.graphic.SetThickness(12)
            # Iterate through teh filtered keyword list (which gives the sorted keyword list) ...
            for (kwg, kw) in self.filteredKeywordList:
                # Determine the color index for this keyword
                colourindex = self.keywordColors[(kwg, kw)]
                # Set the color of the line, using the color lookup for the appropriate color set
                self.graphic.SetColour(colorLookup[colorSet[colourindex]])
                # Add the color box to the graphic
                self.graphic.AddLines([(startX, self.CalcY(Count), startX + 14, self.CalcY(Count) + 14)])
                # Add the text associating the keyword with the colored line we just created
                self.graphic.AddText("%s : %s" % (kwg, kw), startX + 18, self.CalcY(Count) - 7)
                # If the new text extends past the current right-hand boundary ...
                if endX < startX + 14 + self.graphic.GetTextExtent("%s : %s" % (kwg, kw))[0]:
                    # ... note the new right-hand boundary for the box that outlines the legend
                    endX = startX + 14 + self.graphic.GetTextExtent("%s : %s" % (kwg, kw))[0]
                # Note the new bottom boundary for the box that outlines the legend
                endY = self.CalcY(Count) + 14
                # Increment the line counter
                Count += 1
            # Set the line color to black and the line thickness to 1 for the legend bounding box
            self.graphic.SetColour("BLACK")
            self.graphic.SetThickness(1)
            # Draw the legend bounding box, based on the dimensions we've been tracking.
            self.graphic.AddLines([(startX - 6, startY - 24, endX + 6, startY - 24), (endX + 6, startY - 24, endX + 6, endY - 4),
                                   (endX + 6, endY - 4, startX - 6, endY - 4), (startX - 6, endY - 4, startX - 6, startY - 24)])            

    def DrawTimeLine(self, startVal, endVal):
        """ Draw the time line on the Series Map graphic """
        # Set the line thickness to 3
        self.graphic.SetThickness(3)
        # Add a horizontal line from X = start to end ay Y = -2, which will be above the data area of the graph
        self.graphic.AddLines([(self.CalcX(startVal), self.CalcY(-2), self.CalcX(endVal), self.CalcY(-2))])
        # Add Time markers
        self.graphic.SetThickness(1)
        if 'wxMac' in wx.PlatformInfo:
            self.graphic.SetFontSize(11)
        else:
            self.graphic.SetFontSize(8)
        # Add the starting point
        X = startVal
        # Add the line indicator
        self.graphic.AddLines([(self.CalcX(X), self.CalcY(-2) + 1, self.CalcX(X), self.CalcY(-2) + 6)])
        # The Percentage Graph needs a Percent label.  Otherwise, convert to a time representation.
        if self.reportType == 3:
            XLabel = "%d%%" % X
        else:
            XLabel = Misc.TimeMsToStr(X)
        # Add the time label.
        self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(-2) + 5)

        # Add the ending point
        X = endVal
        # Add the line indicator
        self.graphic.AddLines([(self.CalcX(X), self.CalcY(-2) + 1, self.CalcX(X), self.CalcY(-2) + 6)])
        # The Percentage Graph needs a Percent label.  Otherwise, convert to a time representation.
        if self.reportType == 3:
            XLabel = "%d%%" % X
        else:
            XLabel = Misc.TimeMsToStr(X)
        # Add the time label.
        self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(-2) + 5)

        # Add the first and last Vertical Grid Lines, if appropriate
        if self.vGridLines:
            # Determine how far down on the graph the vertical axis lines should go.
            if self.reportType == 1:
                vGridBottom = self.graphic.canvassize[1] - (int(1.75 * max(12, self.barHeight)) + self.whitespaceHeight)
            else:
                vGridBottom = self.CalcY(self.episodeCount - 1) + 7 + int(self.whitespaceHeight / 2)
            # We want Grid Lines in light gray
            self.graphic.SetColour('LIGHT GREY')
            # Add the line for the Start Value
            self.graphic.AddLines([(self.CalcX(self.startTime), self.CalcY(0) - 6 - int(self.whitespaceHeight / 2), self.CalcX(self.startTime), vGridBottom)])
            # Add the line for the End Value
            self.graphic.AddLines([(self.CalcX(endVal), self.CalcY(0) - 6 - int(self.whitespaceHeight / 2), self.CalcX(endVal), vGridBottom)])
            # Reset the graphic color following drawing the Grid Lines
            self.graphic.SetColour("BLACK")

        # Determine the frequency of scale marks for the time line.
        # If we're showing the Percentage Graph ...
        if self.reportType == 3:
            # We'll use marks at every 20%
            numMarks = 5
            interval = 20.0
        # Otherwise ...
        else:
            # We'll use the same logic as the Visualization's Time Axis
            (numMarks, interval) = self.GetScaleIncrements(endVal - startVal)

        # using the incrementation values we just determined ...
        for loop in range(1, numMarks):
            # ... add the intermediate time marks
            X = int(round(float(loop) * interval) + startVal)
            # Add the line indicator
            self.graphic.AddLines([(self.CalcX(X), self.CalcY(-2) + 1, self.CalcX(X), self.CalcY(-2) + 6)])
            # The Percentage Graph needs a Percent label.  Otherwise, convert to a time representation.
            if self.reportType == 3:
                XLabel = "%d%%" % X
            else:
                XLabel = Misc.TimeMsToStr(X)
            # Add the time label.
            self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(-2) + 5)

            # Add Vertical Grid Lines, if appropriate
            if self.vGridLines:
                # We want Grid Lines in light gray
                self.graphic.SetColour('LIGHT GREY')
                # Add the Vertical Grid Line
                self.graphic.AddLines([(self.CalcX(X), self.CalcY(0) - 6 - int(self.whitespaceHeight / 2), self.CalcX(X), vGridBottom)])
                # Reset the graphic color following drawing the Grid Lines
                self.graphic.SetColour("BLACK")

    def GetKeywordCount(self):
        """ Returns the number of keywords in the filtered Keyword List and the size of the image that results """
        return (len(self.filteredKeywordList), len(self.filteredKeywordList) * (self.barHeight + self.whitespaceHeight) + 4)

    def OnMouseMotion(self, event):
        """ Process the movement of the mouse over the Series Map. """
        # Get the mouse's current position
        x = event.GetX()
        y = event.GetY()
        # For the Series Keyword Sequence Map ...
        if (self.reportType == 1):
            # Based on the mouse position, determine the time in the video timeline
            time = self.FindTime(x)
            # Based on the mouse position, determine what keyword is being pointed to
            # We use a different key to mark overlaps depending on whether we're in singleLineDisplay mode or not.
            overlapKey = self.FindKeyword(y)
            # First, let's make sure we're actually on the data portion of the graph
            if (time > 0) and (time < self.MediaLength) and (overlapKey != None) and (overlapKey != '') and (overlapKey != ('', '', '')):
                if self.singleLineDisplay:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_("Episode:  %s,  Time: %s"), 'utf8')
                    else:
                        prompt = _("Episode:  %s,  Time: %s")
                    # Set the Status Text to indicate the current Episode value
                    self.SetStatusText(prompt % (overlapKey, Misc.time_in_ms_to_str(time)))
                else:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_("Episode:  %s,  Keyword:  %s : %s,  Time: %s"), 'utf8')
                    else:
                        prompt = _("Episode:  %s,  Keyword:  %s : %s,  Time: %s")
                    # Set the Status Text to indicate the current Keyword and Time values
                    self.SetStatusText(prompt % (overlapKey[0], overlapKey[1], overlapKey[2], Misc.time_in_ms_to_str(time)))
                if (self.keywordClipList.has_key(overlapKey)):
                    # initialize the string that will hold the names of clips being pointed to
                    clipNames = ''
                    # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                    clips = self.keywordClipList[overlapKey]
                    # For the single-line display ...
                    if self.singleLineDisplay:
                        # Initialize a string for the popup to show
                        clipNames = ''
                        currentRow = None
                        # Get a list of the Lookup dictionary keys.  These keys are top Y-coordinate values
                        keyvals = self.epNameKWGKWLookup.keys()
                        # Sort the keys
                        keyvals.sort()
                        # Iterate through the keys
                        for yVal in keyvals:
                            # We need the largest key value that doesn't exceed the Mouse's Y coordinate
                            if yVal < y:
                                currentRow = self.epNameKWGKWLookup[yVal]
                            # Once the key val exceeds the Mouse position, we can stop looking.
                            else:
                                break

                        # Initialize the Episode Name, Keyword Group, and Keyword variables.
                        epName = KWG = KW = ''
                        # If we have a data record to look at ...
                        if currentRow != None:
                            # Iterate through all the second-level lookup keys, the X ranges ...
                            for key in currentRow.keys():
                                # If the horizontal mouse coordinate falls in the X range of a record ...
                                if (x >= key[0]) and (x < key[1]):
                                    # ... iterate through the records ...
                                    for clipKWRec in currentRow[key]:
                                        # ... extract the Lookup data for the record ...
                                        (epName, KWG, KW, length) = clipKWRec
                                        # ... if it's not the first record in the list, add a comma separator ...
                                        if clipNames != '':
                                            clipNames += ', '
                                        # ... and add the lookup data to the mouseover text string variable
                                        clipNames += "%s : %s (%s)" % (KWG, KW, Misc.time_in_ms_to_str(length))
                    # If we have the Series Keyword Sequence Map multi-line display ...
                    else:
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
            else:
                # ... set the status text to a blank
                self.SetStatusText('')
        # The Series Keyword Bar Graph and the Series Keyword Percentage Graph both work the same way
        elif self.reportType in [2, 3]:
            # Initialize the current Row to None, in case we don't find data under the cursor
            currentRow = None
            # Get a list of the Lookup dictionary keys.  These keys are top Y-coordinate values
            keyvals = self.epNameKWGKWLookup.keys()
            # Sort the keys
            keyvals.sort()
            # Iterate through the keys
            for yVal in keyvals:
                # We need the largest key value that doesn't exceed the Mouse's Y coordinate
                if yVal < y:
                    currentRow = self.epNameKWGKWLookup[yVal]
                # Once the key val exceeds the Mouse position, we can stop looking.
                else:
                    break

            # Initialize the Episode Name, Keyword Group, and Keyword variables.
            epName = KWG = KW = ''
            # If we have a data record to look at ...
            if currentRow != None:
                 # Iterate through all the second-level lookup keys, the X ranges ...
                for key in currentRow.keys():
                    # If the horizontal mouse coordinate falls in the X range of a record ...
                    if (x >= key[0]) and (x < key[1]):
                         # ... extract the Lookup data for the record.  There aren't overlapping records to deal with here.
                        (epName, KWG, KW, length) = currentRow[key]
            # If a data record was found ...
            if KWG != '':
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("Episode:  %s,  Keyword:  %s : %s"), 'utf8')
                else:
                    prompt = _("Episode:  %s,  Keyword:  %s : %s")
                # ... set the Status bar text:
                self.SetStatusText(prompt % (epName, KWG, KW))
                # If we have a Series Keyword Bar Graph ...
                if self.reportType == 2:
                    # ... report Keyword info and Clip Length.
                    self.graphic.SetToolTipString("%s : %s  (%s)" % (KWG, KW, Misc.time_in_ms_to_str(length)))
                # If we have a Series Keyword Percentage Graph ...
                elif self.reportType == 3:
                    # ... report Keyword and Percentage information
                    self.graphic.SetToolTipString("%s : %s  (%3.1f%%)" % (KWG, KW, length))
            # If we've got no data ...
            else:
                # ... reflect that in the Status Text.
                self.SetStatusText('')
                
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
            # If we have a Series Keyword Sequence Map ...
            # (The Bar Graph and Percentage Graph do not have defined Click behaviors!)
            if self.reportType == 1:
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(_("Episode:  %s,  Keyword:  %s : %s,  Time: %s"), 'utf8')
                else:
                    prompt = _("Episode:  %s,  Keyword:  %s : %s,  Time: %s")
                # Set the Status Text to indicate the current Keyword and Time values
                self.SetStatusText(prompt % (kw[0], kw[1], kw[2], Misc.time_in_ms_to_str(time)))
                # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                clips = self.keywordClipList[kw]
                # Iterate through the Clip List ...
                for (startTime, endTime, clipNum, clipName) in clips:
                    # If the current Time value falls between the Clip's StartTime and EndTime ...
                    if (startTime <= time) and (endTime >= time):
                        # Check to see if this is a duplicate Clip
                        if clipNames.has_key(clipName) and (clipNames[clipName] != clipNum):
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
            # If left-click, close the Series Map.  If not, don't!
            if event.LeftUp():
                # Close the Series Map
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
                # If left-click, close the Series Map.  If not, don't!
                if event.LeftUp():
                    # Close the Series Map
                    self.CloseWindow(event)
            # If the user selects Cancel ...
            else:
                # ... destroy the SingleChoiceDialog, but that's all
                dlg.Destroy()
