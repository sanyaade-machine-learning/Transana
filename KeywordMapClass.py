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

import os, sys
# load wxPython for GUI and MySQLdb for Database Access
import wx
import MySQLdb
# load the GraphicsControl
from GraphicsControlClass import GraphicsControl
# Load the Printout Class
from KeywordMapPrintoutClass import MyPrintout
# Import Transana's Database Interface
import DBInterface
# Import Transana's Dialogs
import Dialogs
# Import Transana's Constants
import TransanaConstants
# import Transana's Globals
import TransanaGlobal

# Declare Control IDs
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
# Series List Combo Box
ID_SERIESLIST        = wx.NewId()
# Episode List Combo Box
ID_EPISODELIST       = wx.NewId()

def TimeMsToStr(TimeVal):
    # Converts Time in Milliseconds to a formatted string
    seconds = TimeVal / 1000
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
        rect = wx.ClientDisplayRect()
        width = rect[2] * .80
        height = rect[3] * .80
        # Create the basic Frame structure with a white background
        self.frame = wx.Frame.__init__(self, parent, ID, title, pos=(10, 10), size=wx.Size(width, height), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL | wx.NO_FULL_REPAINT_ON_RESIZE)
        self.SetBackgroundColour(wx.WHITE)
        # Set the icon
        transanaIcon = wx.Icon("images/transana.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(transanaIcon)
        # The rest of the form creation is deferred to the Setup method.  To set the form up,
        # we need data from the database, which we cannot get until we have gotten the username
        # and password information.  However, to get that information, we need a Main Form to
        # exist.

    def Setup(self, seriesName='', episodeName=''):
        self.seriesName = seriesName
        self.episodeName = episodeName
        # You can't have a separate menu on the Mac, so we'll use a Toolbar
        self.toolBar = self.CreateToolBar(wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_TEXT)
        self.toolBar.AddTool(T_FILE_SAVEAS, wx.Bitmap("images/SaveJPG16.xpm", wx.BITMAP_TYPE_XPM), shortHelpString=_('Save As'))
        self.toolBar.AddTool(T_FILE_PRINTSETUP, wx.Bitmap("images/PrintSetup.xpm", wx.BITMAP_TYPE_XPM), shortHelpString=_('Set up Printer'))
        self.toolBar.AddTool(T_FILE_PRINTPREVIEW, wx.Bitmap("images/PrintPreview.xpm", wx.BITMAP_TYPE_XPM), shortHelpString=_('Print Preview'))
        self.toolBar.AddTool(T_FILE_PRINT, wx.Bitmap("images/Print.xpm", wx.BITMAP_TYPE_XPM), shortHelpString=_('Print'))
        self.toolBar.AddTool(T_FILE_EXIT, wx.Bitmap("images/Exit.xpm", wx.BITMAP_TYPE_XPM), shortHelpString=_('Exit'))
        self.toolBar.Realize()
        # Let's go ahead and keep the menu for non-Mac platforms
        if not '__WXMAC__' in wx.PlatformInfo:
            # Add a Menu Bar
            menuBar = wx.MenuBar()                                                 # Create the Menu Bar
            self.menuFile = wx.Menu()                                                   # Create the File Menu
            self.menuFile.Append(M_FILE_SAVEAS, _("Save &As"), _("Save image in JPEG format"))  # Add "Save As" to File Menu
            self.menuFile.Enable(M_FILE_SAVEAS, False)
            self.menuFile.Append(M_FILE_PRINTSETUP, _("Printer Setup"), _("Set up Printer")) # Add "Printer Setup" to the File Menu
            self.menuFile.Append(M_FILE_PRINTPREVIEW, _("Print Preview"), _("Preview your printed output")) # Add "Print Preview" to the File Menu
            self.menuFile.Enable(M_FILE_PRINTPREVIEW, False)
            self.menuFile.Append(M_FILE_PRINT, _("&Print"), _("Send your output to the Printer")) # Add "Print" to the File Menu
            self.menuFile.Enable(M_FILE_PRINT, False)
            self.menuFile.Append(M_FILE_EXIT, _("E&xit"), _("Exit the Keyword Map program")) # Add "Exit" to the File Menu
            menuBar.Append(self.menuFile, _('&File'))                                     # Add the File Menu to the Menu Bar
            self.SetMenuBar(menuBar)                                              # Connect the Menu Bar to the Frame

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

        # Get all Series Records
        #  Create a connection to the database
        DBConn = DBInterface.get_db()
        #  Create a cursor and execute the appropriate query
        self.DBCursor = DBConn.cursor()

        if self.seriesName == '':
            self.DBCursor.execute('SELECT SeriesID FROM Series2 ORDER BY SeriesID')
            #   Get the data from the database
            SeriesIDs = self.DBCursor.fetchall()

            # Create a Choice Box of Series IDs
            #  Create a label for the choice box
            self.lblSeries = wx.StaticText(self, -1, _("Series:"), wx.Point(10,10))
            #  put all the Series IDs into a List Object
            choiceList = [' ']
            for (Series,) in SeriesIDs:
                choiceList.append(Series)
            self.SeriesList = wx.Choice(self, ID_SERIESLIST, wx.Point(60,7), size=(300, 26), choices=choiceList)
            self.SeriesList.SetSelection(0)
            wx.EVT_CHOICE(self, ID_SERIESLIST, self.SeriesSelect)

        if self.episodeName == '':
            # Create a Choice Box for Episodes, even though it's not needed yet
            self.lblEpisode = wx.StaticText(self, -1, _("Episode:"), wx.Point(400,10))
            self.lblEpisode.Show(False)
            # Initialize the list with 10 elements to get a decent-sized dropdown when it is populated later (!)
            choiceList = ['', '', '', '', '', '', '', '', '', '']
            self.EpisodeList = wx.Choice(self, ID_EPISODELIST, wx.Point(450, 7), size = (300, 26), choices = choiceList)
            self.EpisodeList.Show(False)
            wx.EVT_CHOICE(self, ID_EPISODELIST, self.EpisodeSelect)

        # Initialize Media File to nothing
        self.MediaFile = ''
        # Initialize Media Length to 0
        self.MediaLength = 0
        # Initialize Keyword List to empty
        self.KWList = ()

        (w, h) = self.GetClientSizeTuple()
        if (self.seriesName != '') and (self.episodeName != ''):
            self.Bounds = (5, 5, w - 10, h - 25)
        else:
            self.Bounds = (5, 40, w - 10, h - 30)

        self.graphic = GraphicsControl(self, -1, wx.Point(self.Bounds[0], self.Bounds[1]),
                                       (self.Bounds[2] - self.Bounds[0], self.Bounds[3] - self.Bounds[1]),
                                       (self.Bounds[2] - self.Bounds[0], self.Bounds[3] - self.Bounds[1]))

        # Add a Status Bar
        self.CreateStatusBar()

        # Attach the Resize Event
        wx.EVT_SIZE(self, self.OnSize)

        # Prepare objects for use in Printing
        self.printData = wx.PrintData()
        self.printData.SetPaperId(wx.PAPER_LETTER)

        # Center on the screen
        self.CenterOnScreen()
        # Show the Frame
        self.Show(True)

        if (self.seriesName != '') and (self.episodeName != ''):
            # Clear the drawing
            self.KWList = ()
            # Populate the drawing
            self.ProcessEpisode()

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

    def OnSize(self, event):
        (w, h) = self.GetClientSizeTuple()
        if self.Bounds[1] == 5:
            self.Bounds = (5, 5, w - 10, h - 25)
        else:
            self.Bounds = (5, 40, w - 10, h - 30)
        if self.episodeName != '':
            self.ProcessEpisode()

    def SeriesSelect(self, event):
        if event.GetString() != ' ':
            #  Execute the appropriate query
            SQLText = """SELECT EpisodeID
                           FROM Episodes2, Series2
                           WHERE Series2.SeriesID = %s AND
                                 Series2.SeriesNum = Episodes2.SeriesNum
                           ORDER BY EpisodeID"""
            self.DBCursor.execute(SQLText, event.GetString())
            EpisodeIDs = self.DBCursor.fetchall()
            self.EpisodeList.Clear()
            self.EpisodeList.Append(' ')
            for (Episode,) in EpisodeIDs:
                self.EpisodeList.Append(Episode)
            self.EpisodeList.SetSelection(0)
            self.lblEpisode.Show(True)
            self.EpisodeList.Show(True)

            self.seriesName = event.GetString()
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Series %s selected.'), 'utf8')
            else:
                prompt = _('Series %s selected.')
            self.SetStatusText(prompt % self.seriesName)
            self.KWList = ()
            self.MediaFile = ''
            self.MediaLength = 0
        else:
            self.SetStatusText(_('No Series selected.'))
            self.lblEpisode.Show(False)
            self.EpisodeList.Show(False)

        # If you are changing the Series, wipe out the graphic and reset the menus
        self.graphic.Clear()
        self.menuFile.Enable(M_FILE_SAVEAS, False)
        self.menuFile.Enable(M_FILE_PRINTPREVIEW, False)
        self.menuFile.Enable(M_FILE_PRINT, False)

    
    def EpisodeSelect(self, event):
        if event.GetString() != ' ':
            self.episodeName = event.GetString()
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Episode %s selected.'), 'utf8')
            else:
                prompt = _('Episode %s selected.')
            self.SetStatusText(prompt % self.episodeName)
            # Clear the drawing
            self.KWList = ()
            # Populate the drawing
            self.ProcessEpisode()
        else:
            self.SetStatusText(_('No Episode selected.'))
            self.graphic.Clear()
            self.menuFile.Enable(M_FILE_SAVEAS, False)
            self.menuFile.Enable(M_FILE_PRINTPREVIEW, False)
            self.menuFile.Enable(M_FILE_PRINT, False)

    def CalcX(self, XPos):
        """ Determine the proper horizontal coordinate for the given time """

        # Specify a margin width
        marginwidth = (0.06 * (self.Bounds[2] - self.Bounds[0]))
        # The Horizonal Adjustment is the global graphic indent
        hadjust = self.graphicindent
        # The Scaling Factor is the active portion of the drawing area width divided by the total media length
        # The idea is to leave the left margin, self.graphicindent for Keyword Labels, and the right margin
        if self.MediaLength > 0:
            scale = (self.Bounds[2] - self.Bounds[0] - hadjust - 2 * marginwidth) / self.MediaLength  
        else:
            scale = 0.0
        # The horizontal coorinate is 
        # the left margin plus
        # the Horizontal Adjustment for Keyword Labels plus
        # position times the scaling factor
        res = marginwidth + hadjust + (XPos * scale) 

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
        self.KWList = ()
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

        # Get the list of Keywords to be displayed
        SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                       FROM Clips2 cl, ClipKeywords2 ck
                       WHERE cl.EpisodeNum = %s AND
                             cl.ClipNum = ck.ClipNum
                       GROUP BY ck.keywordgroup, ck.keyword
                       ORDER BY KeywordGroup, Keyword, ClipStart"""
        self.DBCursor.execute(SQLText, EpisodeNum)
        self.KWList = list(self.DBCursor.fetchall())

        # Create the Keyword Placement lines to be displayed
        SQLText = """SELECT ck.KeywordGroup, ck.Keyword, cl.ClipStart, cl.ClipStop, cl.ClipNum 
                       FROM Clips2 cl, ClipKeywords2 ck
                       WHERE cl.EpisodeNum = %s AND
                             cl.ClipNum = ck.ClipNum
                       ORDER BY ClipNum, KeywordGroup, Keyword"""
        self.DBCursor.execute(SQLText, EpisodeNum)
        ClipList = self.DBCursor.fetchall()
            
        # Now that we have all necessary information, let's create and populate the graphic
        # Start by destroying the existing control and creating a new one with the correct Canvas Size
        self.graphic.Destroy()
        newheight = max(self.CalcY(len(self.KWList) + 3), self.Bounds[3] - self.Bounds[1])
        self.graphic = GraphicsControl(self, -1, wx.Point(self.Bounds[0], self.Bounds[1]), (self.Bounds[2] - self.Bounds[0], self.Bounds[3] - self.Bounds[1]), (self.Bounds[2] - self.Bounds[0], newheight + 3))

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

        Count = 2
        for KWG, KW in self.KWList:
            KWG = DBInterface.ProcessDBDataForUTF8Encoding(KWG)
            KW = DBInterface.ProcessDBDataForUTF8Encoding(KW)
            self.graphic.AddText("%s : %s" % (KWG, KW), 10, self.CalcY(Count) - 7)
            Count = Count + 1
            
        self.graphicindent = self.graphic.GetMaxWidth(start=3)

        # If the Media Length is known, display the Time Line
        if self.MediaLength > 0:
            self.graphic.SetThickness(3)
            self.graphic.AddLines([(self.CalcX(0), self.CalcY(0), self.CalcX(self.MediaLength), self.CalcY(0))])
            # Add Time markers
            self.graphic.SetThickness(1)
            self.graphic.SetFontSize(8)
            # Add start and end markers
            for loop in range(2):
                X = int(round(float(loop) / 1.0 * self.MediaLength))
                self.graphic.AddLines([(self.CalcX(X), self.CalcY(0) + 1, self.CalcX(X), self.CalcY(0) + 6)])
                XLabel = TimeMsToStr(X)
                self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(0) + 8)

            (numMarks, interval) = self.GetScaleIncrements(self.MediaLength)
            for loop in range(1, numMarks):
                X = int(round(float(loop) * interval))
                self.graphic.AddLines([(self.CalcX(X), self.CalcY(0) + 1, self.CalcX(X), self.CalcY(0) + 6)])
                XLabel = TimeMsToStr(X)
                self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(0) + 8)

        colour = ["BLUE", "RED", "FOREST GREEN", "CORAL", "SLATE BLUE", "BROWN", "CYAN", 
                  "DARK TURQUOISE", "PURPLE", "NAVY", "GREEN", "VIOLET RED"]
        colourindex = 0
        lastclip = 0
        self.graphic.SetThickness(5)
        for (KWG, KW, Start, Stop, ClipNum) in ClipList:
            tempLine = []
            tempLine.append((self.CalcX(Start), self.CalcY(self.KWList.index((KWG, KW)) + 2), self.CalcX(Stop), self.CalcY(self.KWList.index((KWG, KW)) + 2)))
            if (ClipNum != lastclip) and (lastclip != 0):
                 if colourindex < len(colour) - 1:
                     colourindex = colourindex + 1
                 else:
                     colourindex = 0
            self.graphic.SetColour(colour[colourindex])
            self.graphic.AddLines(tempLine)
            lastclip = ClipNum
        if not '__WXMAC__' in wx.PlatformInfo:
            self.menuFile.Enable(M_FILE_SAVEAS, True)
            self.menuFile.Enable(M_FILE_PRINTPREVIEW, True)
            self.menuFile.Enable(M_FILE_PRINT, True)

        
if __name__ == '__main__':

    class MyApp(wx.App):
        def OnInit(self):
            global frame
            # We need to create the main form for the Username Dialog to appear.
            frame = KeywordMap(None, -1, _("Transana Keyword Map"))
            # Let's use a try ... except block to catch database connection errors
            try:
                frame.Setup()
                self.SetTopWindow(frame)
            except MySQLdb.MySQLError, value:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('The following MySQL database error has occurred:\n%s'), 'utf8')
                else:
                    prompt = _('The following MySQL database error has occurred:\n%s')
                dlg = Dialogs.ErrorDialog(frame, prompt % value)
                dlg.ShowModal()
                dlg.Destroy()
                frame.Close()
                return True
            except:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('The following error has occurred:\n%s\n%s'), 'utf8')
                else:
                    prompt = _('The following error has occurred:\n%s\n%s')
                dlg = Dialogs.ErrorDialog(frame, prompt % (sys.exc_type, sys.exc_value))
                dlg.ShowModal()
                dlg.Destroy()
                frame.Close()
                return True
            else:
                return True

    # run the application
    app = MyApp(0)
    app.MainLoop()
