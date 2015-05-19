#Copyright (C) 2002-2006  The Board of Regents of the University of Wisconsin System

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

import os, sys
# load wxPython for GUI and MySQLdb for Database Access
import wx
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

def TimeMsToStr(TimeVal):
    # Converts Time in Milliseconds to a formatted string
    seconds = int(TimeVal) / 1000
    hours = seconds / (60 * 60)
    seconds = seconds % (60 * 60)
    minutes = seconds / 60
    seconds = round(seconds % 60)
    TempStr = ''
    if hours > 0:
        TempStr = '%s:%02d:%02d' % (hours, minutes, seconds)
    else:
        TempStr = '%s:%02d' % (minutes, seconds)
    return TempStr
        

class KeywordMap(wx.Frame):
    """ This is the Main Window for the Keyword Map application """
    def __init__(self, parent, ID, title):
        # Determine the screen size for setting the initial dialog size
        self.parent = parent
        rect = wx.ClientDisplayRect()
        width = rect[2] * .80
        height = rect[3] * .80
        # Create the basic Frame structure with a white background
        self.frame = wx.Frame.__init__(self, parent, ID, title, pos=(10, 10), size=wx.Size(width, height), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL | wx.NO_FULL_REPAINT_ON_RESIZE)
        self.SetBackgroundColour(wx.WHITE)
        # Set the icon
        transanaIcon = wx.Icon(os.path.join(TransanaGlobal.programDir, "images", "transana.ico"), wx.BITMAP_TYPE_ICO)
        self.SetIcon(transanaIcon)
        # The rest of the form creation is deferred to the Setup method.  To set the form up,
        # we need data from the database, which we cannot get until we have gotten the username
        # and password information.  However, to get that information, we need a Main Form to
        # exist.

    def Setup(self, episodeNum, seriesName, episodeName):
        self.episodeNum = episodeNum
        self.seriesName = seriesName
        self.episodeName = episodeName
        # You can't have a separate menu on the Mac, so we'll use a Toolbar
        self.toolBar = self.CreateToolBar(wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_TEXT)
        # Get the graphic for the Filter button
        bmp = wx.ArtProvider_GetBitmap(wx.ART_LIST_VIEW, wx.ART_TOOLBAR, (16,16))
        self.toolBar.AddTool(T_FILE_FILTER, bmp, shortHelpString=_("Filter"))
        self.toolBar.AddTool(T_FILE_SAVEAS, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "SaveJPG16.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Save As'))
        self.toolBar.AddTool(T_FILE_PRINTSETUP, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "PrintSetup.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Set up Printer'))
        self.toolBar.AddTool(T_FILE_PRINTPREVIEW, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "PrintPreview.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Print Preview'))
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
            self.menuFile.Append(M_FILE_PRINTSETUP, _("Printer Setup"), _("Set up Printer")) # Add "Printer Setup" to the File Menu
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

        # Get all Series Records
        #  Create a connection to the database
        DBConn = DBInterface.get_db()
        #  Create a cursor and execute the appropriate query
        self.DBCursor = DBConn.cursor()

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
            # Populate the drawing
            self.ProcessEpisode()
            self.DrawGraph()

    # Define the Method that implements Filter
    def OnFilter(self, event):
        # Create a Filter Dialog.  reportType=1 indicates it is for a Keyword Map.  Keyword Map Filter requires episodeNum
        # for the Config Save.
        dlgFilter = FilterDialog.FilterDialog(self, -1, _("Keyword Map Filter Dialog"), reportType=1, configName=self.configName,
                                              episodeNum=self.episodeNum, clipFilter=True, keywordFilter=True,
                                              keywordSort=True, timeRange=True, startTime=self.startTime, endTime=self.endTime)
        # We want the Clips sorted in Clip ID order in the FilterDialog.  We handle that out here, as the Filter Dialog
        # has to deal with manual clip ordering in some instances, though not here, so it can't deal with this.
        self.clipFilterList.sort()
        dlgFilter.SetClips(self.clipFilterList)
        dlgFilter.SetKeywords(self.unfilteredKeywordList)
        errorMsg = 'Start Loop'
        while errorMsg != '':
            errorMsg = ''
            if dlgFilter.ShowModal() == wx.ID_OK:
                self.clipFilterList = dlgFilter.GetClips()
                self.unfilteredKeywordList = dlgFilter.GetKeywords()
                self.filteredKeywordList = []
                for (kwg, kw, checked) in self.unfilteredKeywordList:
                    if checked:
                        self.filteredKeywordList.append((kwg, kw))

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

            if errorMsg != '':
                errorDlg = Dialogs.ErrorDialog(self, errorMsg)
                errorDlg.ShowModal()
                errorDlg.Destroy()
            
        self.configName = dlgFilter.configName
        dlgFilter.Destroy()
        self.DrawGraph()

    # Define the Method that implements Save As
    def OnSaveAs(self, event):
        self.graphic.SaveAs()

    # Define the Method that implements Printer Setup
    def OnPrintSetup(self, event):
        printerDialog = wx.PrintDialog(self)
        printerDialog.GetPrintDialogData().SetPrintData(self.printData)
        printerDialog.GetPrintDialogData().SetSetupDialog(True)
        printerDialog.ShowModal()
        self.printData = printerDialog.GetPrintDialogData().GetPrintData()
        # Destroying the printerDialog also wipes out the printData object in wxPython 2.5.1.5.  Don't do this.
        # printerDialog.Destroy()

    # Define the Method that implements Print Preview
    def OnPrintPreview(self, event):
        printout = MyPrintout(self.graphic)
        printout2 = MyPrintout(self.graphic)
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
        printout = MyPrintout(self.graphic)
        if not printer.Print(self, printout):
            dlg = Dialogs.ErrorDialog(None, _("There was a problem printing this report."))
            dlg.ShowModal()
            dlg.Destroy()
        else:
            self.printData = printer.GetPrintDialogData().GetPrintData()
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
        if self.Bounds[1] == 5:
            self.Bounds = (5, 5, w - 10, h - 25)
        else:
            self.Bounds = (5, 40, w - 10, h - 30)
        if self.episodeName != '':
            self.DrawGraph()  # self.ProcessEpisode()

    def CalcX(self, XPos):
        """ Determine the proper horizontal coordinate for the given time """
        # Specify a margin width
        marginwidth = (0.06 * (self.Bounds[2] - self.Bounds[0]))
        # The Horizonal Adjustment is the global graphic indent
        hadjust = self.graphicindent
        # The Scaling Factor is the active portion of the drawing area width divided by the total media length
        # The idea is to leave the left margin, self.graphicindent for Keyword Labels, and the right margin
        if self.MediaLength > 0:
            scale = (self.Bounds[2] - self.Bounds[0] - hadjust - 2 * marginwidth) / (self.endTime - self.startTime)
        else:
            scale = 0.0
        # The horizontal coorinate is 
        # the left margin plus
        # the Horizontal Adjustment for Keyword Labels plus
        # position times the scaling factor
        res = marginwidth + hadjust + ((XPos - self.startTime) * scale) 
        return res

    def CalcY(self, YPos):
        return 14 * YPos + 30

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
        self.unfilteredKeywordList = []
        self.filteredKeywordList = []
        self.clipFilterList = []
        self.startTime = 0
        self.endTime = 0
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
            self.endTime = self.MediaLength

        # Get the list of Keywords to be displayed
        SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                       FROM Clips2 cl, ClipKeywords2 ck
                       WHERE cl.EpisodeNum = %s AND
                             cl.ClipNum = ck.ClipNum
                       GROUP BY ck.keywordgroup, ck.keyword
                       ORDER BY KeywordGroup, Keyword, ClipStart"""
        self.DBCursor.execute(SQLText, EpisodeNum)

        for row in self.DBCursor.fetchall():
            self.filteredKeywordList.append((row[0], row[1]))
            self.unfilteredKeywordList.append((row[0], row[1], True))

        # Create the Keyword Placement lines to be displayed.  We need them to be in ClipStart, ClipNum order so colors will be
        # distributed properly across bands.
        SQLText = """SELECT ck.KeywordGroup, ck.Keyword, cl.ClipStart, cl.ClipStop, cl.ClipNum, cl.ClipID, cl.CollectNum
                       FROM Clips2 cl, ClipKeywords2 ck
                       WHERE cl.EpisodeNum = %s AND
                             cl.ClipNum = ck.ClipNum
                       ORDER BY ClipStart, ClipNum, KeywordGroup, Keyword"""
        self.DBCursor.execute(SQLText, EpisodeNum)
        for (kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum) in self.DBCursor.fetchall():
            self.clipList.append((kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum))
            if not ((clipID, collectNum, True) in self.clipFilterList):
                self.clipFilterList.append((clipID, collectNum, True))

    def DrawGraph(self):
        """ Actually Draw the Keyword Map """
        self.keywordClipList = {}
        # Now that we have all necessary information, let's create and populate the graphic
        # Start by destroying the existing control and creating a new one with the correct Canvas Size
        self.graphic.Destroy()
        newheight = max(self.CalcY(len(self.filteredKeywordList) + 3), self.Bounds[3] - self.Bounds[1])
        self.graphic = GraphicsControlClass.GraphicsControl(self, -1, wx.Point(self.Bounds[0], self.Bounds[1]),
                                                            (self.Bounds[2] - self.Bounds[0], self.Bounds[3] - self.Bounds[1]),
                                                            (self.Bounds[2] - self.Bounds[0], newheight + 3),
                                                            passMouseEvents=True)

        self.graphic.SetFontColour("BLACK")
        self.graphic.SetFontSize(10)
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
            
        Count = 2
        for KWG, KW in self.filteredKeywordList:
            KWG = DBInterface.ProcessDBDataForUTF8Encoding(KWG)
            KW = DBInterface.ProcessDBDataForUTF8Encoding(KW)
            self.graphic.AddText("%s : %s" % (KWG, KW), 10, self.CalcY(Count) - 7)
            Count = Count + 1
            
        self.graphicindent = self.graphic.GetMaxWidth(start=3)

        # If the Media Length is known, display the Time Line
        if self.MediaLength > 0:
            self.graphic.SetThickness(3)
            self.graphic.AddLines([(self.CalcX(self.startTime), self.CalcY(0), self.CalcX(self.endTime), self.CalcY(0))])
            # Add Time markers
            self.graphic.SetThickness(1)
            self.graphic.SetFontSize(8)

            X = self.startTime
            self.graphic.AddLines([(self.CalcX(X), self.CalcY(0) + 1, self.CalcX(X), self.CalcY(0) + 6)])
            XLabel = TimeMsToStr(X)
            self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(0) + 8)
            X = self.endTime
            self.graphic.AddLines([(self.CalcX(X), self.CalcY(0) + 1, self.CalcX(X), self.CalcY(0) + 6)])
            XLabel = TimeMsToStr(X)
            self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(0) + 8)

            (numMarks, interval) = self.GetScaleIncrements(self.endTime - self.startTime)
            for loop in range(1, numMarks):
                X = int(round(float(loop) * interval) + self.startTime)
                self.graphic.AddLines([(self.CalcX(X), self.CalcY(0) + 1, self.CalcX(X), self.CalcY(0) + 6)])
                XLabel = TimeMsToStr(X)
                self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(0) + 8)


        # Get the legal colors for bars in the Keyword Map.  These colors are taken from the TransanaConstants.transana_colorList
        # but are put in a different order for the Keyword Map.
        # ('Green' in line 3 wasn't distinct enough from 'Green Blue' and 'Chartreuse', so I removed it and moved 'Olive' up from line 6.)
        colour = ['Dark Blue', 'Green Blue', 'Gray', 'Light Fuscia',
                  'Blue', 'Dark Green', 'Lavendar', 'Rose',
                  'Light Blue', 'Olive', 'Purple', 'Red',
                  'Cyan', 'Chartreuse', 'Dark Purple', 'Salmon',
                  'Light Aqua', 'Light Green', 'Maroon', 'Orange',
                  'Blue Green', 'Magenta', 'Yellow']
        
        colourindex = 0
        lastclip = 0
        # some clip boundary lines for overlapping clips can get over-written, depeding on the nature of the overlaps.
        # Let's create a separate list of these lines, which we'll add to the END of the process so they can't get overwritten.
        overlapLines = []

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
                    self.graphic.SetThickness(5)
                    tempLine = []
                    tempLine.append((self.CalcX(Start)+2, self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2), self.CalcX(Stop)-2, self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)))
                    if (ClipNum != lastclip) and (lastclip != 0):
                         if colourindex < len(colour) - 1:
                             colourindex = colourindex + 1
                         else:
                             colourindex = 0
                    self.graphic.SetColour(TransanaConstants.transana_colorLookup[colour[colourindex]])

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
                        self.graphic.SetThickness(2)
                        self.graphic.SetColour("RED")
                        tempLine = [(self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)-1, self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)-1)]
                        self.graphic.AddLines(tempLine)
                        self.graphic.SetColour("GREEN")
                        tempLine = [(self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)+1, self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)+1)]
                        self.graphic.AddLines(tempLine)
                        self.graphic.SetColour("BLUE")
                        tempLine = [(self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)+2, self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)+2)]
                        self.graphic.AddLines(tempLine)
                        # Let's remember the clip start and stop boundaries, to be drawn at the end so they won't get over-written
                        overlapLines.append(((self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)-2, self.CalcX(overlapStart), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)+2),))
                        overlapLines.append(((self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)-2, self.CalcX(overlapEnd), self.CalcY(self.filteredKeywordList.index((KWG, KW)) + 2)+2),))

                    # ... add the new Clip to the Clip List
                    self.keywordClipList[(KWG, KW)].append((Start, Stop, ClipNum, ClipName))
                # If there is no entry for the given keyword ...
                else:
                    # ... create a List object with the first clip's data for this Keyword Pair key
                    self.keywordClipList[(KWG, KW)] = [(Start, Stop, ClipNum, ClipName)]

        # let's add the overlap boundary lines now
        self.graphic.SetThickness(1)
        self.graphic.SetColour("BLACK")
        for tempLine in overlapLines:
            self.graphic.AddLines(tempLine)
            
        if not '__WXMAC__' in wx.PlatformInfo:
            self.menuFile.Enable(M_FILE_SAVEAS, True)
            self.menuFile.Enable(M_FILE_PRINTPREVIEW, True)
            self.menuFile.Enable(M_FILE_PRINT, True)

        # Enable tracking of mouse movement over the graphic
        self.graphic.Bind(wx.EVT_MOTION, self.OnMouseMotion)

        

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
            scale = (self.Bounds[2] - self.Bounds[0] - hadjust - 2 * marginwidth) / (self.endTime - self.startTime)
        else:
            scale = 0.0
        # The time is calculated by taking the total width, subtracting the margin values and horizontal indent,
        # and then dividing the result by the scale factor calculated above
        time = int((x - marginwidth - hadjust) / scale) + self.startTime
        return time

    def FindKeyword(self, y):
        """ Given a vertical pixel position, determine the corresponding Keyword data """
        # If the graphic is scrolled, the raw Y value does not point to the correct Keyword.
        # Determine the unscrolled equivalent Y position.
        (modX, modY) = self.graphic.CalcUnscrolledPosition(0, y)
        # Each keyword is 14 pixels high, but we need to adjust for the top margin and
        # the amount of graphic scroll too.
        kwIndex = int((modY - 52) / 14)
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

    def OnMouseMotion(self, event):
        """ Process the movement of the mouse over the Keyword Map. """
        # Get the mouse's current position
        x = event.GetX()
        y = event.GetY()
        # Based on the mouse position, determine the time in the video timeline
        time = self.FindTime(x)
        # Based on the mouse position, determine what keyword is being pointed to
        kw = self.FindKeyword(y)
        # First, let's make sure we're actually on the data portion of the graph
        if (time > 0) and (time < self.MediaLength) and (kw != None):
            # Set the Status Text to indicate the current Keyword and Time values
            self.SetStatusText(_("Keyword:  %s : %s,  Time: %s") % (kw[0], kw[1], Misc.time_in_ms_to_str(time)))
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
            # Set the Status Text to indicate the current Keyword and Time values
            self.SetStatusText(_("Keyword:  %s : %s,  Time: %s") % (kw[0], kw[1], Misc.time_in_ms_to_str(time)))
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
