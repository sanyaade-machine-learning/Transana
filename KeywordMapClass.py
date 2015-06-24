#Copyright (C) 2002-2015  The Board of Regents of the University of Wisconsin System

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
# import Python's platform module
import platform
# load wxPython for GUI
import wx
# load the GraphicsControl
import GraphicsControlClass
# Load the Printout Class
from KeywordMapPrintoutClass import MyPrintout
# Load the Collection object
import Collection
# Import Transana's Database Interface
import DBInterface
# Import Transana's Dialogs
import Dialogs
# Import Transana's Document object
import Document
# Import Transana's Episode object
import Episode
# Import Transana's Filter Dialog
import FilterDialog
# import Transana's Keyword Object
import KeywordObject
# import Transana Miscellaneous functions
import Misc
# import Transana's Quote object
import Quote
# import Transana's Constants
import TransanaConstants
# Import Transana's Exceptions
import TransanaExceptions
# import Transana's Globals
import TransanaGlobal
# Import Transana's Images
import TransanaImages

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
    def __init__(self, parent, ID=-1, title="", embedded=False, topOffset=0, controlObject=None):
        # It's always important to remember your ancestors.
        self.parent = parent
        # Remember the title
        self.title = title
        # Initialize the Report Number
        self.reportNumber = 0
        # We do some things differently if we're a free-standing Keyword Map report 
        # or if we're an embedded Keyword Visualization.  
        self.embedded = embedded
        # Remember the topOffset parameter value.  This is used to specify a larger top margin for the keyword visualization,
        # needed for the Hybrid Visualization.
        self.topOffset = topOffset
        # Let's remember the Control Object, if one is passed in
        self.ControlObject = controlObject
        # If a Control Object has been passed in ...
        if self.ControlObject != None:
            # ... register this report with the Control Object (which adds it to the Windows Menu)
            self.ControlObject.AddReportWindow(self)
        #  Create a connection to the database
        DBConn = DBInterface.get_db()
        #  Create a cursor and execute the appropriate query
        self.DBCursor = DBConn.cursor()
        # If we're NOT embedded, we need to create a full frame etc.
        if not self.embedded:
            # Determine the screen size for setting the initial dialog size
            rect = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()  # wx.ClientDisplayRect()
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
        # Initialize DocumentNum and EpisodeNum
        self.documentNum = None
        self.episodeNum = None
        # Initialize CollectionNum and the collection object
        self.collectionNum = None
        self.collection = None
        # Initialize the Text (Document) Object
        self.textObj = None
        # Initialize Media File to nothing
        self.MediaFile = ''
        # Initialize Media Length and Character Length to 0
        self.MediaLength = 0
        self.CharacterLength = 0
        # Initialize Keyword Lists to empty
        self.unfilteredKeywordList = []
        self.filteredKeywordList = []
        # Intialize the Clip List to empty
        self.clipList = []
        self.clipFilterList = []
        # Initialize the Snapshot List to empty
        self.snapshotList = []
        self.snapshotFilterList = []
        # Initialize the Quote List to empty
        self.quoteList = []
        self.quoteFilterList = []
        # To be able to show only parts of an Episode Time Line, we need variables for the time boundaries.
        self.startTime = 0
        self.endTime = 0
        # Initialize the StartChar and EndChar values to -1 to indicate we're dealing with Media rather than Text
        self.startChar = -1
        self.endChar = -1
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
            # Get the colorAsKeyword value
            self.colorAsKeywords = TransanaGlobal.configData.colorAsKeywords
        else:
            # Get the Configuration values for the Keyword Visualization Options
            self.barHeight = TransanaGlobal.configData.keywordVisualizationBarHeight
            self.whitespaceHeight = TransanaGlobal.configData.keywordVisualizationWhitespace
            self.hGridLines = TransanaGlobal.configData.keywordVisualizationHorizontalGridLines
            self.vGridLines = TransanaGlobal.configData.keywordVisualizationVerticalGridLines
            self.colorOutput = True

    def Setup(self, documentNum=None, episodeNum=None, collNum=None, seriesName='', documentName='', episodeName=''):
        """ Complete initialization for the free-standing Keyword Map or Collection Keyword Map, not the embedded version. """
        # Remember the appropriate Document information
        self.documentNum = documentNum
        self.documentName = documentName
        # Remember the appropriate Episode information
        self.episodeNum = episodeNum
        self.episodeName = episodeName
        # Remember the Series / Library information
        self.seriesName = seriesName
        # Remember the appropriate Collection information
        self.collectionNum = collNum
        # indicate that we're not working from a Quote.  (The Keyword Map is never Quote-based.)
        self.quoteNum = None
        # indicate that we're not working from a Clip.  (The Keyword Map is never Clip-based.)
        self.clipNum = None

        # Initialize Time and Character Positions
        self.startTime = -1
        self.endTime = -1
        self.startChar = -1
        self.endChar = -1
        
        # You can't have a separate menu on the Mac, so we'll use a Toolbar
        self.toolBar = self.CreateToolBar(wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_TEXT)
        # Get the graphic for the Filter button
        self.toolBar.AddTool(T_FILE_FILTER, TransanaImages.ArtProv_LISTVIEW.GetBitmap(), shortHelpString=_("Filter"))
        self.toolBar.AddTool(T_FILE_SAVEAS, TransanaImages.SaveJPG16.GetBitmap(), shortHelpString=_('Save As'))
        self.toolBar.AddTool(T_FILE_PRINTSETUP, TransanaImages.PrintSetup.GetBitmap(), shortHelpString=_('Set up Page'))

        self.toolBar.AddTool(T_FILE_PRINTPREVIEW, TransanaImages.PrintPreview.GetBitmap(), shortHelpString=_('Print Preview'))
        # Disable Print Preview on the PPC Mac and for Right-To-Left languages
        if (platform.processor() == 'powerpc') or (TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft):
            self.toolBar.EnableTool(T_FILE_PRINTPREVIEW, False)
            
        self.toolBar.AddTool(T_FILE_PRINT, TransanaImages.Print.GetBitmap(), shortHelpString=_('Print'))

        # create a bitmap button for the Move Down button
        self.toolBar.AddTool(T_HELP_HELP, TransanaImages.ArtProv_HELP.GetBitmap(), shortHelpString=_("Help"))
        self.toolBar.AddTool(T_FILE_EXIT, TransanaImages.Exit.GetBitmap(), shortHelpString=_('Exit'))        
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
            if self.collectionNum == None:
                self.menuFile.Append(M_FILE_EXIT, _("E&xit"), _("Exit the Keyword Map program")) # Add "Exit" to the File Menu
            else:
                self.menuFile.Append(M_FILE_EXIT, _("E&xit"), _("Exit the Collection Keyword Map program")) # Add "Exit" to the File Menu
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
        # Bind the form's EVT_CLOSE method
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # Determine the window boundaries
        (w, h) = self.GetClientSizeTuple()
        # If not embedded ...
        if ((self.seriesName != '') and (self.episodeName != '')) or (self.collectionNum != None):
            self.Bounds = (5, 5, w - 10, h - 25)
        # If embedded ...
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
        TransanaGlobal.CenterOnPrimary(self)

        # If we have a Series name and Document Name, we are doing a Document Keyword Map
        if (self.seriesName != '') and (self.documentName != ''):
            # Load the Text Object, in this case a Document
            self.textObj = Document.Document(self.documentNum)
            # Initialize the Quote Filter List to be empty
            self.quoteFilterList = []
            # Initialize the Snapshot Filter List to be empty
            self.snapshotFilterList = []
            # Clear the drawing
            self.filteredKeywordList = []
            self.unfilteredKeywordList = []
            # Populate the drawing
            self.ProcessDocument()
            # We need to draw the graph before we set the Default filter
            self.DrawGraph()
            # Trigger the load of the Default filter, if one exists.  An event of None signals we're loading the
            # Default config, and the OnFilter method will handle drawing the graph!
            self.OnFilter(None)

        # If we have a Series name and Episode Name, we are doing an Episode Keyword Map
        elif (self.seriesName != '') and (self.episodeName != ''):
            # Initialize the Clip Filter List to be empty
            self.clipFilterList = []
            # Initialize the Snapshot Filter List to be empty
            self.snapshotFilterList = []
            # Clear the drawing
            self.filteredKeywordList = []
            self.unfilteredKeywordList = []
            # Populate the drawing
            self.ProcessEpisode()
            # We need to draw the graph before we set the Default filter
            self.DrawGraph()
            # Trigger the load of the Default filter, if one exists.  An event of None signals we're loading the
            # Default config, and the OnFilter method will handle drawing the graph!
            self.OnFilter(None)

        # If we have a Collection Number, we're doing a Collection Keyword Map
        elif self.collectionNum != None:
            # Create a collection object
            self.collection = Collection.Collection(self.collectionNum)
            # Clear the drawing
            self.filteredKeywordList = []
            self.unfilteredKeywordList = []
            # Populate the drawing
            self.ProcessCollection()
            # We need to draw the graph before we set the Default filter
            self.DrawGraph()
            # Trigger the load of the Default filter, if one exists.  An event of None signals we're loading the
            # Default config, and the OnFilter method will handle drawing the graph!
            self.OnFilter(None)

        # Show the Frame
        self.Show(True)

    def SetupEmbedded(self, episodeNum, seriesName, episodeName, startTime, endTime,
                      filteredClipList=[], unfilteredClipList = [],
                      filteredSnapshotList=[], unfilteredSnapshotList = [],
                      filteredKeywordList=[], unfilteredKeywordList = [],
                      keywordColors = None, clipNum=None, configName='', loadDefault=False):
        """ Complete setup for the embedded version of the Keyword Map. """
        # Remember the appropriate Episode information
        self.episodeNum = episodeNum
        self.seriesName = seriesName
        self.episodeName = episodeName
        self.clipNum = clipNum
        self.configName = configName
        # Set the start and end time boundaries (especially important for Clips!)
        self.startTime = startTime
        self.endTime = endTime
        # Set the StartChar and EndChar values to -1 to indicate we're dealing with Media rather than Text
        self.startChar = -1
        self.endChar = -1
        self.CharacterLength = 0
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
            # Set the initial Clip Lists
            self.clipFilterList = filteredClipList[:]
            self.clipList = unfilteredClipList[:]
            # Set the initial Snapshot Lists
            self.snapshotFilterList = filteredSnapshotList[:]
            self.snapshotList = unfilteredSnapshotList[:]
            # set the initial keyword lists
            self.filteredKeywordList = filteredKeywordList[:]
            self.unfilteredKeywordList = unfilteredKeywordList[:]
            # If we got keywordColors, use them!!
            if keywordColors != None:
                self.keywordColors = keywordColors
            # Populate the drawing
            self.ProcessEpisode()
            # We need to draw the graph before we set the Default filter
            self.DrawGraph()
            # If we need to load the Default Configuration ...
            if loadDefault:
                # We actually need to wipe out the original graphic prior to loading the Default filter!
                self.graphic.Clear()
                # Trigger the load of the Default filter, if one exists.  An event of None signals we're loading the
                # Default config, and the OnFilter method will handle drawing the graph!
                self.OnFilter(None)

    def SetupTextEmbedded(self,
                          textObj,
                          startChar,
                          endChar,
                          totalLength,
                          filteredQuoteList=[],
                          unfilteredQuoteList = [],
##                          filteredSnapshotList=[],
##                          unfilteredSnapshotList = [],
                          filteredKeywordList=[],
                          unfilteredKeywordList = [],
                          keywordColors = None,
                          quoteNum=None,
                          configName='',
                          loadDefault=False):
        """ Complete setup for the embedded version of the Text Keyword Map. """
        # Remember the appropriate object
        self.textObj = textObj
        self.quoteNum = quoteNum
        self.configName = configName
        # Set the StartTime and EndTime values to -1 to indicate we're dealing with Text rather than media
        # Actually, no.  That causes problems with showing a selection in the Text Keyword Visualization!
        # self.startTime = -1
        self.endTime = -1
        # Set the start and end position boundaries (especially important for Quotes!)
        self.startChar = startChar
        self.endChar = endChar
        # Clear the Media Length
        self.MediaLength = 0
        # Remember the total character length
        self.CharacterLength = totalLength
        # Toggle the Embedded labels.  (Used for testing mouse-overs only)
        self.showEmbeddedLabels = False
        
        # Determine the graphic's boundaries
        w = self.graphic.getWidth()
        h = self.graphic.getHeight()
        self.Bounds = (0, 0, w, h - 25)

        # If we have a defined textObj (which we always should) ...
        if isinstance(self.textObj, Document.Document) or isinstance(self.textObj, Quote.Quote):
            # Set the initial Quote Lists
            self.quoteFilterList = filteredQuoteList[:]
            self.quoteList = unfilteredQuoteList[:]
##            # Set the initial Snapshot Lists
##            self.snapshotFilterList = filteredSnapshotList[:]
##            self.snapshotList = unfilteredSnapshotList[:]
            # set the initial keyword lists
            self.filteredKeywordList = filteredKeywordList[:]
            self.unfilteredKeywordList = unfilteredKeywordList[:]
            # If we got keywordColors, use them!!
            if keywordColors != None:
                self.keywordColors = keywordColors
            # Populate the drawing
            self.ProcessDocument()
            # We need to draw the graph before we set the Default filter
            self.DrawGraph()
            # If we need to load the Default Configuration ...
            if loadDefault:
                # We actually need to wipe out the original graphic prior to loading the Default filter!
                self.graphic.Clear()
                # Trigger the load of the Default filter, if one exists.  An event of None signals we're loading the
                # Default config, and the OnFilter method will handle drawing the graph!
                self.OnFilter(None)

    # Define the Method that implements Filter
    def OnFilter(self, event):
        """ Implement the Filter Dialog call for Keyword Maps and Keyword Visualizations """
        # See if we're loading the Default profile.  This is signalled by an event of None!
        if event == None:
            loadDefault = True
        else:
            loadDefault = False
        # Set up parameters for creating the Filter Dialog.  Keyword Map/Keyword Visualization Filter requires episodeNum for the Config Save.
        if not self.embedded:
            # For the keyword map, the form created here is the parent
            parent = self
            # If we have an Episode Keyword Map ...
            if self.episodeNum != None:
                # Set and encode the dialog title
                title = unicode(_("Episode Keyword Map Filter Dialog"), 'utf8')
                # reportType=1 indicates it is for a Keyword Map.  
                reportType = 1
                reportScope = self.episodeNum
                startVal = max(self.startTime, 0)
                endVal = max(self.endTime, 0)
            # If we have a Collection Keyword Map
            elif self.collectionNum != None:
                # Set and encode the dialog title
                title = unicode(_("Collection Keyword Map Filter Dialog"), 'utf8')
                # reportType=16 indicates it is for a Collection Keyword Map.  
                reportType = 16
                reportScope = self.collectionNum
                startVal = max(self.startTime, 0)
                endVal = max(self.endTime, 0)
            # If we have a Document Keyword Map
            elif self.textObj != None:
                # Set and encode the dialog title
                title = unicode(_("Document Keyword Map Filter Dialog"), 'utf8')
                # reportType=17 indicates it is for a Document Keyword Map.  
                reportType = 17
                reportScope = self.textObj.number
                startVal = max(self.startChar, 0)
                endVal = max(self.endChar, 0)
            # See if there are Quotes in the Filter List
            quoteFilter = (len(self.quoteFilterList) > 0)
            # See if there are Clips in the Filter List
            clipFilter = (len(self.clipFilterList) > 0)
            # See if there are Snapshots in the Snapshot Filter List
            snapshotFilter = (len(self.snapshotFilterList) > 0)
            # See if there are Keywords in the Filter List
            keywordFilter = (len(self.unfilteredKeywordList) > 0)
            # Keyword Map and Collection Keyword Map now support Keyword Color customization, at least sometimes
            keywordColors = (len(self.unfilteredKeywordList) > 0)
            # We want the Options tab
            options = True
            # Create a Filter Dialog, passing all the necessary parameters.
            dlgFilter = FilterDialog.FilterDialog(parent,
                                                  -1,
                                                  title,
                                                  reportType=reportType,
                                                  loadDefault=loadDefault,
                                                  configName=self.configName,
                                                  reportScope=reportScope,
                                                  quoteFilter=quoteFilter,
                                                  clipFilter=clipFilter,
                                                  snapshotFilter=snapshotFilter,
                                                  keywordFilter=keywordFilter,
                                                  keywordSort=True,
                                                  keywordColor=keywordColors,
                                                  options=options,
                                                  startTime=startVal,
                                                  endTime=endVal,
                                                  barHeight=self.barHeight,
                                                  whitespace=self.whitespaceHeight,
                                                  hGridLines=self.hGridLines,
                                                  vGridLines=self.vGridLines,
                                                  colorOutput=self.colorOutput,
                                                  colorAsKeywords=self.colorAsKeywords)
        else:
            # For the keyword visualization, the parent that was passed in on initialization is the parent
            parent = self.parent
            title = unicode(_("Keyword Visualization Filter Dialog"), 'utf8')
            # See if there are Quotes in the Filter List
            quoteFilter = (len(self.quoteFilterList) > 0)
            # See if there are Clips in the Filter List
            clipFilter = (len(self.clipFilterList) > 0)
            # See if there are Snapshots in the Snapshot Filter List
            snapshotFilter = (len(self.snapshotFilterList) > 0)
            # See if there are Keywords in the Filter List
            keywordFilter = (len(self.unfilteredKeywordList) > 0)
            # Keyword visualization wants Keyword Color customization
            keywordColors = (len(self.unfilteredKeywordList) > 0)
            # We want the Options tab
            options = True
            # If we have an Episode Keyword Visualization ...
            if self.episodeNum != None:
                # reportType=2 indicates it is for a Keyword Visualization.  
                reportType = 2
                reportScope = self.episodeNum
            # If we have a Document Keyword Map
            elif self.textObj != None:
                # reportType=18 indicates it is for a Document Keyword Visualization.  
                reportType = 18
                reportScope = self.textObj.number
            # Create a Filter Dialog, passing all the necessary parameters.
            dlgFilter = FilterDialog.FilterDialog(parent,
                                                  -1,
                                                  title,
                                                  reportType=reportType,
                                                  loadDefault=loadDefault,
                                                  configName=self.configName,
                                                  reportScope=reportScope,
                                                  quoteFilter=quoteFilter,
                                                  clipFilter=clipFilter,
                                                  snapshotFilter=snapshotFilter,
                                                  keywordFilter=keywordFilter,
                                                  keywordSort=True,
                                                  keywordColor=keywordColors,
                                                  options=options,
                                                  startTime=self.startTime,
                                                  endTime=self.endTime,
                                                  barHeight=self.barHeight,
                                                  whitespace=self.whitespaceHeight,
                                                  hGridLines=self.hGridLines,
                                                  vGridLines=self.vGridLines)
        # If we requested the Quote Filter ...
        if quoteFilter:
            # We want the Quotes sorted in Quote ID order in the FilterDialog.  We handle that out here, as the Filter Dialog
            # has to deal with manual Quote ordering in some instances, though not here, so it can't deal with this.
            self.quoteFilterList.sort()
            # Inform the Filter Dialog of the Quotes
            dlgFilter.SetQuotes(self.quoteFilterList)
        # If we requested the Clip Filter ...
        if clipFilter:
            # We want the Clips sorted in Clip ID order in the FilterDialog.  We handle that out here, as the Filter Dialog
            # has to deal with manual clip ordering in some instances, though not here, so it can't deal with this.
            self.clipFilterList.sort()
            # Inform the Filter Dialog of the Clips
            dlgFilter.SetClips(self.clipFilterList)
        # if there are Snapshots ...
        if snapshotFilter:
            # ... populate the Filter Dialog with Snapshots
            dlgFilter.SetSnapshots(self.snapshotFilterList)
        # if there are Keywords ...
        if keywordFilter:
            # Keyword Colors must be specified before Keywords!  So if we want Keyword Colors, ...
            if keywordColors:
                # Inform the Filter Dialog of the colors used for each Keyword
                dlgFilter.SetKeywordColors(self.keywordColors)
            # Populate the Filter Dialog with Keywords
            dlgFilter.SetKeywords(self.unfilteredKeywordList)
        # Create a dummy error message to get our while loop started.
        errorMsg = 'Start Loop'
        # Keep trying as long as there is an error message
        while errorMsg != '':
            # Clear the last (or dummy) error message.
            errorMsg = ''
            # If we're loading the Default configuration ...
            if loadDefault:
                # ... get the list of existing configuration names.
                profileList = dlgFilter.GetConfigNames()
                # If (translated) "Default" is in the list ...
                # (NOTE that the default config name is stored in English, but gets translated by GetConfigNames!)
                if unicode(_('Default'), 'utf8') in profileList:
                    # ... then signal that we need to load the config.
                    dlgFilter.OnFileOpen(None)
                    # Fake that we asked the user for a filter name and got an OK
                    result = wx.ID_OK
                # If we're loading a Default profile, but there's none in the list, we can skip
                # the rest of the Filter method by pretending we got a Cancel from the user.
                else:
                    result = wx.ID_CANCEL
            # If we're not loading a Default profile ...
            else:
                # ... we need to show the Filter Dialog here.
                result = dlgFilter.ShowModal()
                
            # If the user clicks OK (or we have a Default config)
            if result == wx.ID_OK:
                # If we requested Quote Filtering ...
                if quoteFilter:
                    # ... then get the filtered quote data
                    self.quoteFilterList = dlgFilter.GetQuotes()
                # If we requested Clip Filtering ...
                if clipFilter:
                    # ... then get the filtered clip data
                    self.clipFilterList = dlgFilter.GetClips()
                # If we requested Snapshot Filtering ...
                if snapshotFilter:
                    # ... then get the filtered snapshot data
                    self.snapshotFilterList = dlgFilter.GetSnapshots()
                # If we requested Keyword filtering ...
                if keywordFilter:
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
                        if self.textObj == None:
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
                            if (self.endTime <= 0) or (self.endTime > self.MediaLength):
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
                        else:
                            try:
                                # Let's get the character Range data.
                                # Start Pos must be 0 or greater.  Otherwise, don't change it!
                                if int(dlgFilter.GetStartTime()) >= 0:
                                    self.startChar = int(dlgFilter.GetStartTime())
                                else:
                                    errorMsg += _("Illegal value for Start Position.\n")
                                
                                # If the Start Position is greater than the media length, reset it to 0.
                                if self.startTime >= self.CharacterLength:
                                    dlgFilter.startTime.SetValue('0')
                                    errorMsg += _("Illegal value for Start Position.\n")
                            except:
                                errorMsg += _("Illegal value for Start Position.\n")

                            try:
                                # End Position must be at least 0.  Otherwise, don't change it!
                                if (int(dlgFilter.GetEndTime()) >= 0):
                                    self.endChar = int(dlgFilter.GetEndTime())
                                else:
                                    errorMsg += _("Illegal value for End Position.\n")
                            except:
                                errorMsg += _("Illegal value for End Position.\n")

                            # If the end position is 0 or greater than the character length, set it to the character length.
                            if (self.endChar <= 0) or (self.endChar > self.CharacterLength):
                                self.endChar = self.CharacterLength

                            # Start position cannot equal end position (but this check must come after setting endChar == 0 to CharacterLength)
                            if self.startChar == self.endChar:
                                errorMsg += _("Start Position and End Position must be different.")
                                # We need to alter the time values to prevent "division by zero" errors while the Filter Dialog is not modal.
                                self.startChar = 0
                                self.endChar = self.CharacterLength

                            # If the Start Position is greater than the End Position, swap them.
                            if (self.endChar < self.startChar):
                                temp = self.startChar
                                self.startChar = self.endChar
                                self.endChar = temp

                        # Get the colorOutput value from the dialog IF we're in the Keyword Map
                        self.colorOutput = dlgFilter.GetColorOutput()

                        # Get the colorAsKeywords value from the dialog IF we're in the Keyword Map
                        self.colorAsKeywords = dlgFilter.GetColorAsKeywords()

                    # Get the Bar Height and Whitespace Height for both versions of the Keyword Map
                    self.barHeight = dlgFilter.GetBarHeight()
                    self.whitespaceHeight = dlgFilter.GetWhitespace()
                    # we need to store the Bar Height, Whitespace, and colorAsKeywords values in the Configuration.
                    if not self.embedded:
                        TransanaGlobal.configData.keywordMapBarHeight = self.barHeight
                        TransanaGlobal.configData.keywordMapWhitespace = self.whitespaceHeight
                        TransanaGlobal.configData.colorAsKeywords = self.colorAsKeywords
                    else:
                        TransanaGlobal.configData.keywordVisualizationBarHeight = self.barHeight
                        TransanaGlobal.configData.keywordVisualizationWhitespace = self.whitespaceHeight

                    # Keyword Map Report, Keyword Visualization, the Series Keyword Sequence Map, and the Collection Keyword Map
                    # have Bar height and Whitespace parameters as well as horizontal and vertical grid lines
                    if reportType in [1, 2, 5, 6, 7, 16, 17, 18]:
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
        theWidth = max(wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()[2] - 180, 760)  # wx.ClientDisplayRect()
        theHeight = max(wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()[3] - 200, 560)  # wx.ClientDisplayRect()
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

    def OnClose(self, event):
        """ Handle the Close Event """
        # If the report has a defined Control Object ...
        if self.ControlObject != None:
            # ... remove this report from the Menu Window's Window Menu
            self.ControlObject.RemoveReportWindow(self.title, self.reportNumber)
        # Inherit the parent Close event so things will, you know, close.
        event.Skip()

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
        # If we have data defined in the graph ...
        if (self.episodeName != '') or (self.textObj != None) or (self.collection != None):
            # ... redraw the graph
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

        # If we have a TIME-based Map ...
        if (self.startChar == -1) and (self.endChar == -1):
            lowerVal = self.startTime
            upperVal = self.endTime
            totalLength = self.MediaLength
        # If we have a CHARACTER-based Map ...
        else:
            lowerVal = self.startChar
            upperVal = self.endChar
            totalLength = self.CharacterLength

        # The Scaling Factor is the active portion of the drawing area width divided by the total media length
        # The idea is to leave the left margin, self.graphicindent for Keyword Labels, and the right margin
        if (totalLength > 0) and (upperVal > lowerVal):
            scale = (float(self.Bounds[2]) - self.Bounds[0] - hadjust - 2 * marginwidth) / (upperVal - lowerVal)
        else:
            scale = 0.0
        # The horizontal coordinate is the left margin plus the Horizontal Adjustment for Keyword Labels plus
        # position times the scaling factor
        res = marginwidth + hadjust + ((XPos - lowerVal) * scale)

        # If we are in a Right-To-Left Language ...
        if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
            # ... adjust for a right-to-left graph
            return int(self.Bounds[2] - self.Bounds[0] - res)
        # If we are in a Left-To-Right language ...
        else:
            # ... just return the calculated value
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
        # If we have a TIME-based Map ...
        if (self.startChar == -1) and (self.endChar == -1):
            lowerVal = self.startTime
            upperVal = self.endTime
            totalLength = self.MediaLength
        # If we have a CHARACTER-based Map ...
        else:
            lowerVal = self.startChar
            upperVal = self. endChar
            totalLength = self.CharacterLength
        # The Scaling Factor is the active portion of the drawing area width divided by the total media length
        # The idea is to leave the left margin, self.graphicindent for Keyword Labels, and the right margin
        if (upperVal - lowerVal) > 0:
            scale = (float(self.Bounds[2]) - self.Bounds[0] - hadjust - 2 * marginwidth) / (upperVal - lowerVal)
        else:
            scale = 1.0
        # The time is calculated by taking the total width, subtracting the margin values and horizontal indent,
        # and then dividing the result by the scale factor calculated above
        time = int((x - marginwidth - hadjust) / scale) + lowerVal
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

    def ProcessDocument(self):
        """ Process a Document """
        if isinstance(self.textObj, Document.Document):
            self.CharacterLength = self.textObj.document_length
            if (self.startChar == -1) and (self.endChar == -1):
                self.startChar = 0
                self.endChar = self.CharacterLength
        elif isinstance(self.textObj, Quote.Quote):
            self.CharacterLength = max(self.endChar - self.startChar, 1)
            if (self.startChar == -1) and (self.endChar == -1):
                self.startChar = 0
                self.endChar = self.CharacterLength
                
        # If this is our first time through ...
        if (self.filteredKeywordList == []) and (self.unfilteredKeywordList == []):
            # If we deleted the last keyword in a filtered list, the Filter Dialog ended up with
            # duplicate entries.  This should prevent it!!
            self.unfilteredKeywordList = []
            if isinstance(self.textObj, Document.Document):
                # Get the list of QUOTE Keywords to be displayed
                SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                               FROM Quotes2 q, ClipKeywords2 ck
                               WHERE q.SourceDocumentNum = %s AND
                                     q.QuoteNum = ck.QuoteNum
                               GROUP BY ck.keywordgroup, ck.keyword
                               ORDER BY KeywordGroup, Keyword"""
            elif isinstance(self.textObj, Quote.Quote):
                # Get the list of QUOTE Keywords to be displayed
                SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                               FROM Quotes2 q, ClipKeywords2 ck
                               WHERE q.QuoteNum = %s AND
                                     q.QuoteNum = ck.QuoteNum
                               GROUP BY ck.keywordgroup, ck.keyword
                               ORDER BY KeywordGroup, Keyword"""
            # Adjust the query for sqlite if needed
            SQLText = DBInterface.FixQuery(SQLText)
            self.DBCursor.execute(SQLText, (self.textObj.number, ))
            for (kwg, kw) in self.DBCursor.fetchall():
                kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                if not (kwg, kw) in self.filteredKeywordList:
                    self.filteredKeywordList.append((kwg, kw))
                if not (kwg, kw, True) in self.unfilteredKeywordList:
                    self.unfilteredKeywordList.append((kwg, kw, True))

##            if TransanaConstants.proVersion:
##                # Get the list of WHOLE SNAPSHOT Keywords to be displayed
##                SQLText = """SELECT ck.KeywordGroup, ck.Keyword
##                               FROM Snapshots2 sn, ClipKeywords2 ck
##                               WHERE sn.EpisodeNum = %s AND
##                                     sn.SnapshotNum = ck.SnapshotNum
##                               GROUP BY ck.keywordgroup, ck.keyword
##                               ORDER BY KeywordGroup, Keyword, SnapshotTimeCode"""
##                # Adjust the query for sqlite if needed
##                SQLText = DBInterface.FixQuery(SQLText)
##                self.DBCursor.execute(SQLText, (self.episodeNum, ))
##                for (kwg, kw) in self.DBCursor.fetchall():
##                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
##                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
##                    if not (kwg, kw) in self.filteredKeywordList:
##                        self.filteredKeywordList.append((kwg, kw))
##                    if not (kwg, kw, True) in self.unfilteredKeywordList:
##                        self.unfilteredKeywordList.append((kwg, kw, True))
##
##                # Get the list of SNAPSHOT CODING Keywords to be displayed
##                SQLText = """SELECT ck.KeywordGroup, ck.Keyword
##                               FROM Snapshots2 sn, SnapshotKeywords2 ck
##                               WHERE sn.EpisodeNum = %s AND
##                                     sn.SnapshotNum = ck.SnapshotNum
##                               GROUP BY ck.keywordgroup, ck.keyword
##                               ORDER BY KeywordGroup, Keyword, SnapshotTimeCode"""
##                # Adjust the query for sqlite if needed
##                SQLText = DBInterface.FixQuery(SQLText)
##                self.DBCursor.execute(SQLText, (self.episodeNum, ))
##                for (kwg, kw) in self.DBCursor.fetchall():
##                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
##                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
##                    if not (kwg, kw) in self.filteredKeywordList:
##                        self.filteredKeywordList.append((kwg, kw))
##                    if not (kwg, kw, True) in self.unfilteredKeywordList:
##                        self.unfilteredKeywordList.append((kwg, kw, True))

        # If we haven't loaded a configuration (which contains its own sort order) ...
        if self.configName == '':
            # Sort the Keyword List
            self.unfilteredKeywordList.sort()
            self.filteredKeywordList.sort()
        
        if isinstance(self.textObj, Document.Document):
            # Create the Quote Keyword Placement lines to be displayed.  We need them to be in StartChar, QuoteNum order so colors will be
            # distributed properly across bands.
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, qp.StartChar, qp.EndChar, q.QuoteNum, q.QuoteID, q.CollectNum
                           FROM Quotes2 q, QuotePositions2 qp, ClipKeywords2 ck
                           WHERE q.SourceDocumentNum = %s AND
                                 q.QuoteNum = qp.QuoteNum AND
                                 q.QuoteNum = ck.QuoteNum
                           ORDER BY StartChar, q.QuoteNum, KeywordGroup, Keyword"""
        elif isinstance(self.textObj, Quote.Quote):
            # Create the Quote Keyword Placement lines to be displayed.  We need them to be in StartChar, QuoteNum order so colors will be
            # distributed properly across bands.
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, qp.StartChar, qp.EndChar, q.QuoteNum, q.QuoteID, q.CollectNum
                           FROM Quotes2 q, QuotePositions2 qp, ClipKeywords2 ck
                           WHERE q.QuoteNum = %s AND
                                 q.QuoteNum = qp.QuoteNum AND
                                 q.QuoteNum = ck.QuoteNum
                           ORDER BY StartChar, q.QuoteNum, KeywordGroup, Keyword"""
        # Adjust the query for sqlite if needed
        SQLText = DBInterface.FixQuery(SQLText)
        self.DBCursor.execute(SQLText, (self.textObj.number, ))
        for (kwg, kw, startChar, endChar, quoteNum, quoteID, collectNum) in self.DBCursor.fetchall():
            kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
            kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
            quoteID = DBInterface.ProcessDBDataForUTF8Encoding(quoteID)
            # Handle orphaned Quotes
            if isinstance(self.textObj, Quote.Quote):
                if startChar == -1:
                    startChar = 0
                if endChar == -1:
                    endChar = 1
            # If we're dealing with a Document, self.QuoteNum will be None and we want all Quotes.
            # If we're dealing with a Quote, we only want to deal with THIS Quote!
            if (self.quoteNum == None) or (quoteNum == self.quoteNum):
                if (kwg, kw, startChar, endChar, quoteNum, quoteID, collectNum) not in self.quoteList:
                    if isinstance(self.textObj, Document.Document):
                        self.quoteList.append((kwg, kw, self.textObj.quote_dict[quoteNum][0], self.textObj.quote_dict[quoteNum][1], quoteNum, quoteID, collectNum))
                    elif isinstance(self.textObj, Quote.Quote):
                        self.quoteList.append((kwg, kw, startChar, endChar, quoteNum, quoteID, collectNum))
                if (not ((quoteID, collectNum, True) in self.quoteFilterList)) and \
                   (not ((quoteID, collectNum, False) in self.quoteFilterList)):
                    self.quoteFilterList.append((quoteID, collectNum, True))

##        if TransanaConstants.proVersion:
##            # Create the WHOLE SNAPSHOT Keyword Placement lines to be displayed.  We need them to be in SnapshotTimeCode, SnapshotNum order so colors will be
##            # distributed properly across bands.
##            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, sn.SnapshotTimeCode, sn.SnapshotDuration, sn.SnapshotNum, sn.SnapshotID, sn.CollectNum
##                           FROM Snapshots2 sn, ClipKeywords2 ck
##                           WHERE sn.EpisodeNum = %s AND
##                                 sn.SnapshotNum = ck.SnapshotNum
##                           ORDER BY SnapshotTimeCode, sn.SnapshotNum, KeywordGroup, Keyword"""
##            # Adjust the query for sqlite if needed
##            SQLText = DBInterface.FixQuery(SQLText)
##            self.DBCursor.execute(SQLText, (self.episodeNum, ))
##            for (kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID, collectNum) in self.DBCursor.fetchall():
##                kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
##                kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
##                SnapshotID = DBInterface.ProcessDBDataForUTF8Encoding(SnapshotID)
##                # If we're dealing with an Episode, self.clipNum will be None and we want all clips.
##                # If we're dealing with a Clip, we only want to deal with THIS clip!
##                if (self.clipNum == None):
##                    if (kwg, kw, SnapshotTimeCode, SnapshotTimeCode + SnapshotDuration, SnapshotNum, SnapshotID, collectNum) not in self.snapshotList:
##                        self.snapshotList.append((kwg, kw, SnapshotTimeCode, SnapshotTimeCode + SnapshotDuration, SnapshotNum, SnapshotID, collectNum))
##                    if (not ((SnapshotID, collectNum, True) in self.snapshotFilterList)) and \
##                       (not ((SnapshotID, collectNum, False) in self.snapshotFilterList)):
##                        self.snapshotFilterList.append((SnapshotID, collectNum, True))
##
##            # Create the SNAPSHOT CODING Keyword Placement lines to be displayed.  We need them to be in SnapshotTimeCode, SnapshotNum order so colors will be
##            # distributed properly across bands.
##            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, sn.SnapshotTimeCode, sn.SnapshotDuration, sn.SnapshotNum, sn.SnapshotID, sn.CollectNum
##                           FROM Snapshots2 sn, SnapshotKeywords2 ck
##                           WHERE sn.EpisodeNum = %s AND
##                                 sn.SnapshotNum = ck.SnapshotNum
##                           ORDER BY SnapshotTimeCode, sn.SnapshotNum, KeywordGroup, Keyword"""
##            # Adjust the query for sqlite if needed
##            SQLText = DBInterface.FixQuery(SQLText)
##            self.DBCursor.execute(SQLText, (self.episodeNum, ))
##            for (kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID, collectNum) in self.DBCursor.fetchall():
##                kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
##                kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
##                SnapshotID = DBInterface.ProcessDBDataForUTF8Encoding(SnapshotID)
##                # If we're dealing with an Episode, self.clipNum will be None and we want all clips.
##                # If we're dealing with a Clip, we only want to deal with THIS clip!
##                if (self.clipNum == None):
##                    if (kwg, kw, SnapshotTimeCode, SnapshotTimeCode + SnapshotDuration, SnapshotNum, SnapshotID, collectNum) not in self.snapshotList:
##                        self.snapshotList.append((kwg, kw, SnapshotTimeCode, SnapshotTimeCode + SnapshotDuration, SnapshotNum, SnapshotID, collectNum))
##                    if (not ((SnapshotID, collectNum, True) in self.snapshotFilterList)) and \
##                       (not ((SnapshotID, collectNum, False) in self.snapshotFilterList)):
##                        self.snapshotFilterList.append((SnapshotID, collectNum, True))

    def ProcessEpisode(self):
        """ Process a Keyword Map for an Episode """
        # We need a data struture to hold the data about what clips correspond to what keywords
        self.MediaFile = ''
        self.MediaLength = 0

        # Start Exception Handling
        try:
            # Load the specified Episode
            tmpEpObj = Episode.Episode(num=self.episodeNum)
            # Note the Media File Name (without path) and the Media File Length
            self.MediaFile = os.path.split(tmpEpObj.media_filename)[1]
            self.MediaLength = tmpEpObj.episode_length()
            # If the end time is 0 or greater than the (non-zero) media length, set it to the media length.
            if (self.endTime <= 0) or ((self.endTime > self.MediaLength) and (self.MediaLength > 0)):
                self.endTime = self.MediaLength
        # If we don't have a single record from the database, we probably have an orphaned Clip.
        except TransanaExceptions.RecordNotFoundError:
            # ... and we can set the MediaLength to the end time passed in.
            self.MediaLength = self.endTime
        except:
            import traceback
            traceback.print_exc(file=sys.stdout)

        # If this is our first time through ...
        if (self.filteredKeywordList == []) and (self.unfilteredKeywordList == []):
            # If we deleted the last keyword in a filtered list, the Filter Dialog ended up with
            # duplicate entries.  This should prevent it!!
            self.unfilteredKeywordList = []
            # Get the list of CLIP Keywords to be displayed
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                           FROM Clips2 cl, ClipKeywords2 ck
                           WHERE cl.EpisodeNum = %s AND
                                 cl.ClipNum = ck.ClipNum
                           GROUP BY ck.keywordgroup, ck.keyword
                           ORDER BY KeywordGroup, Keyword, ClipStart"""
            # Adjust the query for sqlite if needed
            SQLText = DBInterface.FixQuery(SQLText)
            self.DBCursor.execute(SQLText, (self.episodeNum, ))
            for (kwg, kw) in self.DBCursor.fetchall():
                kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                if not (kwg, kw) in self.filteredKeywordList:
                    self.filteredKeywordList.append((kwg, kw))
                if not (kwg, kw, True) in self.unfilteredKeywordList:
                    self.unfilteredKeywordList.append((kwg, kw, True))

            if TransanaConstants.proVersion:
                # Get the list of WHOLE SNAPSHOT Keywords to be displayed
                SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                               FROM Snapshots2 sn, ClipKeywords2 ck
                               WHERE sn.EpisodeNum = %s AND
                                     sn.SnapshotNum = ck.SnapshotNum
                               GROUP BY ck.keywordgroup, ck.keyword
                               ORDER BY KeywordGroup, Keyword, SnapshotTimeCode"""
                # Adjust the query for sqlite if needed
                SQLText = DBInterface.FixQuery(SQLText)
                self.DBCursor.execute(SQLText, (self.episodeNum, ))
                for (kwg, kw) in self.DBCursor.fetchall():
                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                    if not (kwg, kw) in self.filteredKeywordList:
                        self.filteredKeywordList.append((kwg, kw))
                    if not (kwg, kw, True) in self.unfilteredKeywordList:
                        self.unfilteredKeywordList.append((kwg, kw, True))

                # Get the list of SNAPSHOT CODING Keywords to be displayed
                SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                               FROM Snapshots2 sn, SnapshotKeywords2 ck
                               WHERE sn.EpisodeNum = %s AND
                                     sn.SnapshotNum = ck.SnapshotNum
                               GROUP BY ck.keywordgroup, ck.keyword
                               ORDER BY KeywordGroup, Keyword, SnapshotTimeCode"""
                # Adjust the query for sqlite if needed
                SQLText = DBInterface.FixQuery(SQLText)
                self.DBCursor.execute(SQLText, (self.episodeNum, ))
                for (kwg, kw) in self.DBCursor.fetchall():
                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                    if not (kwg, kw) in self.filteredKeywordList:
                        self.filteredKeywordList.append((kwg, kw))
                    if not (kwg, kw, True) in self.unfilteredKeywordList:
                        self.unfilteredKeywordList.append((kwg, kw, True))

        # If we haven't loaded a configuration (which contains its own sort order) ...
        if self.configName == '':
            # Sort the Keyword List
            self.unfilteredKeywordList.sort()
            self.filteredKeywordList.sort()
        
        # Create the Clip Keyword Placement lines to be displayed.  We need them to be in ClipStart, ClipNum order so colors will be
        # distributed properly across bands.
        SQLText = """SELECT ck.KeywordGroup, ck.Keyword, cl.ClipStart, cl.ClipStop, cl.ClipNum, cl.ClipID, cl.CollectNum
                       FROM Clips2 cl, ClipKeywords2 ck
                       WHERE cl.EpisodeNum = %s AND
                             cl.ClipNum = ck.ClipNum
                       ORDER BY ClipStart, cl.ClipNum, KeywordGroup, Keyword"""
        # Adjust the query for sqlite if needed
        SQLText = DBInterface.FixQuery(SQLText)
        self.DBCursor.execute(SQLText, (self.episodeNum, ))
        for (kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum) in self.DBCursor.fetchall():
            kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
            kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
            clipID = DBInterface.ProcessDBDataForUTF8Encoding(clipID)
            # If we're dealing with an Episode, self.clipNum will be None and we want all clips.
            # If we're dealing with a Clip, we only want to deal with THIS clip!
            if (self.clipNum == None) or (clipNum == self.clipNum):
                if (kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum) not in self.clipList:
                    self.clipList.append((kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum))
                if (not ((clipID, collectNum, True) in self.clipFilterList)) and \
                   (not ((clipID, collectNum, False) in self.clipFilterList)):
                    self.clipFilterList.append((clipID, collectNum, True))

        if TransanaConstants.proVersion:
            # Create the WHOLE SNAPSHOT Keyword Placement lines to be displayed.  We need them to be in SnapshotTimeCode, SnapshotNum order so colors will be
            # distributed properly across bands.
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, sn.SnapshotTimeCode, sn.SnapshotDuration, sn.SnapshotNum, sn.SnapshotID, sn.CollectNum
                           FROM Snapshots2 sn, ClipKeywords2 ck
                           WHERE sn.EpisodeNum = %s AND
                                 sn.SnapshotNum = ck.SnapshotNum
                           ORDER BY SnapshotTimeCode, sn.SnapshotNum, KeywordGroup, Keyword"""
            # Adjust the query for sqlite if needed
            SQLText = DBInterface.FixQuery(SQLText)
            self.DBCursor.execute(SQLText, (self.episodeNum, ))
            for (kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID, collectNum) in self.DBCursor.fetchall():
                kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                SnapshotID = DBInterface.ProcessDBDataForUTF8Encoding(SnapshotID)
                # If we're dealing with an Episode, self.clipNum will be None and we want all clips.
                # If we're dealing with a Clip, we only want to deal with THIS clip!
                if (self.clipNum == None):
                    if (kwg, kw, SnapshotTimeCode, SnapshotTimeCode + SnapshotDuration, SnapshotNum, SnapshotID, collectNum) not in self.snapshotList:
                        self.snapshotList.append((kwg, kw, SnapshotTimeCode, SnapshotTimeCode + SnapshotDuration, SnapshotNum, SnapshotID, collectNum))
                    if (not ((SnapshotID, collectNum, True) in self.snapshotFilterList)) and \
                       (not ((SnapshotID, collectNum, False) in self.snapshotFilterList)):
                        self.snapshotFilterList.append((SnapshotID, collectNum, True))

            # Create the SNAPSHOT CODING Keyword Placement lines to be displayed.  We need them to be in SnapshotTimeCode, SnapshotNum order so colors will be
            # distributed properly across bands.
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, sn.SnapshotTimeCode, sn.SnapshotDuration, sn.SnapshotNum, sn.SnapshotID, sn.CollectNum
                           FROM Snapshots2 sn, SnapshotKeywords2 ck
                           WHERE sn.EpisodeNum = %s AND
                                 sn.SnapshotNum = ck.SnapshotNum
                           ORDER BY SnapshotTimeCode, sn.SnapshotNum, KeywordGroup, Keyword"""
            # Adjust the query for sqlite if needed
            SQLText = DBInterface.FixQuery(SQLText)
            self.DBCursor.execute(SQLText, (self.episodeNum, ))
            for (kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID, collectNum) in self.DBCursor.fetchall():
                kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                SnapshotID = DBInterface.ProcessDBDataForUTF8Encoding(SnapshotID)
                # If we're dealing with an Episode, self.clipNum will be None and we want all clips.
                # If we're dealing with a Clip, we only want to deal with THIS clip!
                if (self.clipNum == None):
                    if (kwg, kw, SnapshotTimeCode, SnapshotTimeCode + SnapshotDuration, SnapshotNum, SnapshotID, collectNum) not in self.snapshotList:
                        self.snapshotList.append((kwg, kw, SnapshotTimeCode, SnapshotTimeCode + SnapshotDuration, SnapshotNum, SnapshotID, collectNum))
                    if (not ((SnapshotID, collectNum, True) in self.snapshotFilterList)) and \
                       (not ((SnapshotID, collectNum, False) in self.snapshotFilterList)):
                        self.snapshotFilterList.append((SnapshotID, collectNum, True))

    def ProcessCollection(self):
        """ Process a Collection for the Collection Keyword Map variation of the Keyword Map """
        # Initialize the Clip Filter List
        self.clipFilterList = []
        # Initialize the Snapshot Filter List
        self.snapshotFilterList = []
        # We don't have a single Media File here.  Leave it blank
        self.MediaFile = ''
        # Initialize the Media Length, which we will accumulate from the clips
        self.MediaLength = 0

        # We need a data struture to hold the data about what clips and snapshots correspond to what keywords.
        # But we only need to process it once.
        if self.filteredKeywordList == []:
            # If we deleted the last keyword in a filtered list, the Filter Dialog ended up with
            # duplicate entries.  This should prevent it!!
            self.unfilteredKeywordList = []
            # Get the list of CLIP Keywords to be displayed.  This query should do it.
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                           FROM Clips2 cl, ClipKeywords2 ck
                           WHERE cl.CollectNum = %s AND
                                 cl.ClipNum = ck.ClipNum
                           GROUP BY ck.keywordgroup, ck.keyword
                           ORDER BY KeywordGroup, Keyword, ClipStart"""
            # Adjust the query for sqlite if needed
            SQLText = DBInterface.FixQuery(SQLText)
            # Execute the query
            self.DBCursor.execute(SQLText, (self.collectionNum, ))
            # For each record in the query results ...
            for (kwg, kw) in self.DBCursor.fetchall():
                # ... encode the KWG and KW
                kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                # ... and add them to the filtered and unfiltered keyword lists
                self.filteredKeywordList.append((kwg, kw))
                self.unfilteredKeywordList.append((kwg, kw, True))

            if TransanaConstants.proVersion:
                # Get the list of WHOLE SNAPSHOT Keywords to be displayed
                SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                               FROM Snapshots2 sn, ClipKeywords2 ck
                               WHERE sn.CollectNum = %s AND
                                     sn.SnapshotNum = ck.SnapshotNum
                               GROUP BY ck.keywordgroup, ck.keyword
                               ORDER BY KeywordGroup, Keyword, SnapshotTimeCode"""
                # Adjust the query for sqlite if needed
                SQLText = DBInterface.FixQuery(SQLText)
                # Execute the query
                self.DBCursor.execute(SQLText, (self.collectionNum, ))
                # For each record in the query results ...
                for (kwg, kw) in self.DBCursor.fetchall():
                    # ... encode the KWG and KW
                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                    # ... and IF they're not already there, add them to the filtered and unfiltered keyword lists.
                    # Unlike with Clips, a Snapshot can have multiple instances of the same keyword!!
                    if not (kwg, kw) in self.filteredKeywordList:
                        self.filteredKeywordList.append((kwg, kw))
                    if not (kwg, kw, True) in self.unfilteredKeywordList:
                        self.unfilteredKeywordList.append((kwg, kw, True))

                # Get the list of SNAPSHOT CODING Keywords to be displayed
                SQLText = """SELECT ck.KeywordGroup, ck.Keyword
                               FROM Snapshots2 sn, SnapshotKeywords2 ck
                               WHERE sn.CollectNum = %s AND
                                     sn.SnapshotNum = ck.SnapshotNum
                               GROUP BY ck.keywordgroup, ck.keyword
                               ORDER BY KeywordGroup, Keyword, SnapshotTimeCode"""
                # Adjust the query for sqlite if needed
                SQLText = DBInterface.FixQuery(SQLText)
                # Execute the query
                self.DBCursor.execute(SQLText, (self.collectionNum, ))
                # For each record in the query results ...
                for (kwg, kw) in self.DBCursor.fetchall():
                    # ... encode the KWG and KW
                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                    # ... and IF they're not already there, add them to the filtered and unfiltered keyword lists.
                    # Unlike with Clips, a Snapshot can have multiple instances of the same keyword!!
                    if not (kwg, kw) in self.filteredKeywordList:
                        self.filteredKeywordList.append((kwg, kw))
                    if not (kwg, kw, True) in self.unfilteredKeywordList:
                        self.unfilteredKeywordList.append((kwg, kw, True))

        # Sort the Keyword Lists
        self.unfilteredKeywordList.sort()
        self.filteredKeywordList.sort()

        # We need to track what clip we're looking at.  Initialize a variable for that.
        currClip = 0
        # We need to track what snapshot we're looking at too.
        currSnapshot = 0
        # Initialize the Collection Contents dictionary, which allows us to sort Clips and Snapshots in the right order
        collectionOrder = {}

        # Get the CLIP information for the Keyword Placement lines to be displayed.  
        SQLText = """SELECT ck.KeywordGroup, ck.Keyword, cl.ClipStart, cl.ClipStop, cl.ClipNum, cl.ClipID, cl.CollectNum, cl.SortOrder
                       FROM Clips2 cl, ClipKeywords2 ck
                       WHERE cl.CollectNum = %s AND
                             cl.ClipNum = ck.ClipNum
                       ORDER BY SortOrder, KeywordGroup, Keyword"""
        # Adjust the query for sqlite if needed
        SQLText = DBInterface.FixQuery(SQLText)
        # Execute the query
        self.DBCursor.execute(SQLText, (self.collectionNum, ))
        # Iterate through the query results
        for (kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum, sortOrder) in self.DBCursor.fetchall():
            # If there's no entry for this item in the Sort Order ...
            if not collectionOrder.has_key(sortOrder):
                # ... create a new list of Clip Keyword entries for this Sort Order item
                collectionOrder[sortOrder] = [('Clip', kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum)]
            # If there already is an entry for this item in the Sort Order ...
            else:
                # ... append the new Keyword to the list.  (We must allow multiple keywords per clip/snapshot!)
                collectionOrder[sortOrder].append(('Clip', kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum))

        if TransanaConstants.proVersion:
            # Create the WHOLE SNAPSHOT Keyword Placement lines to be displayed.
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, sn.SnapshotTimeCode, sn.SnapshotDuration, sn.SnapshotNum,
                                sn.SnapshotID, sn.CollectNum, sn.SortOrder
                           FROM Snapshots2 sn, ClipKeywords2 ck
                           WHERE sn.CollectNum = %s AND
                                 sn.SnapshotNum = ck.SnapshotNum
                           ORDER BY SnapshotTimeCode, sn.SnapshotNum, KeywordGroup, Keyword"""
            # Adjust the query for sqlite if needed
            SQLText = DBInterface.FixQuery(SQLText)
            # Execute the query
            self.DBCursor.execute(SQLText, (self.collectionNum, ))
            # Iterate through the query results
            for (kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID, collectNum, sortOrder) in self.DBCursor.fetchall():
                # If the Snapshot does not have a defined duration ...
                if SnapshotDuration <= 0:
                    # ... let's give it a temporary length of 10 seconds so it will show up!
                    SnapshotDuration = 10000
                # If there's no entry for this item in the Sort Order ...
                if not collectionOrder.has_key(sortOrder):
                    # ... create a new list of Snapshot Keyword entries for this Sort Order item
                    collectionOrder[sortOrder] = [('Snapshot', kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID,
                                                   collectNum)]
                # If there already is an entry for this item in the Sort Order ...
                else:
                    # ... append the new Keyword to the list.  (We must allow multiple keywords per clip/snapshot!)
                    collectionOrder[sortOrder].append(('Snapshot', kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID,
                                                       collectNum))

            # Create the SNAPSHOT CODING Keyword Placement lines to be displayed.  
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, sn.SnapshotTimeCode, sn.SnapshotDuration, sn.SnapshotNum,
                                sn.SnapshotID, sn.CollectNum, sn.SortOrder
                           FROM Snapshots2 sn, SnapshotKeywords2 ck
                           WHERE sn.CollectNum = %s AND
                                 sn.SnapshotNum = ck.SnapshotNum
                           ORDER BY SnapshotTimeCode, sn.SnapshotNum, KeywordGroup, Keyword"""
            # Adjust the query for sqlite if needed
            SQLText = DBInterface.FixQuery(SQLText)
            # Execute the query
            self.DBCursor.execute(SQLText, (self.collectionNum, ))
            # Iterate through the query results
            for (kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID, collectNum, sortOrder) in self.DBCursor.fetchall():
                # If the Snapshot does not have a defined duration ...
                if SnapshotDuration <= 0.0:
                    # ... let's give it a temporary length of 10 seconds so it will show up!
                    SnapshotDuration = 10000
                # If there's no entry for this item in the Sort Order ...
                if not collectionOrder.has_key(sortOrder):
                    # ... create a new list of Snapshot Keyword entries for this Sort Order item
                    collectionOrder[sortOrder] = [('Snapshot', kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID,
                                                   collectNum)]
                # If there already is an entry for this item in the Sort Order ...
                else:
                    # ... append the new Keyword to the list.  (We must allow multiple keywords per clip/snapshot!)
                    collectionOrder[sortOrder].append(('Snapshot', kwg, kw, SnapshotTimeCode, SnapshotDuration, SnapshotNum, SnapshotID,
                                                       collectNum))
        # Get the Sort Order Keys (which are the Sort Order values)
        keys = collectionOrder.keys()
        # Sort the Sort Order Keys, so they're in Sort Order order!!
        keys.sort()
        # Iterate through the Collection Order dictionary KEYS, that is, the Collection items' Sort Order.
        # (This allows us to combine Clips and Snapshots in the correct Sort Order!)
        for sortOrder in keys:
            # Get the LIST of data records for the next sort order and iterate through that list
            for recData in collectionOrder[sortOrder]:
                # If the next record is a CLIP ...
                if recData[0] == 'Clip':
                    # ... extract the Clip data
                    (recType, kwg, kw, clipStart, clipStop, clipNum, clipID, collectNum) = recData
                    # If we have not yet added THIS clip to total time ...
                    if clipNum != currClip:
                        # ... add this clip's length, plus 100 ms, to the total length of the Collection Keyword Map's time line ...
                        self.MediaLength += (clipStop - clipStart + 100)
                        # ... and update the current clip number so this clip won't be counted again if it has multiple keywords
                        currClip = clipNum
                    # Decode the KWG, KW, and Clip ID
                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                    clipID = DBInterface.ProcessDBDataForUTF8Encoding(clipID)
                    # If we're dealing with a Collection Keyword Report, self.clipNum will be None
                    if (self.clipNum == None):
                        # Add the current Clip/keyword combo to the Clip List, placing it to the right of the last clip.
                        self.clipList.append((kwg, kw, self.MediaLength - (clipStop - clipStart) - 50, self.MediaLength - 50, clipNum, clipID, collectNum))
                        # If the clip ID isn't already in the Clip Filter List ...
                        if not ((clipID, collectNum, True) in self.clipFilterList):
                            # ... add the clip to the Clip Filter List
                            self.clipFilterList.append((clipID, collectNum, True))
                # If the next record is a SNAPSHOT ...
                elif recData[0] == 'Snapshot':
                    # ... extract the Snapshot data
                    (recType, kwg, kw, snapshotTimeCode, snapshotDuration, snapshotNum, snapshotID, collectNum) = recData
                    # If we have not yet added THIS Snapshot to total time ...
                    if snapshotNum != currSnapshot:
                        # ... add this snapshot's length, plus 100 ms, to the total length of the Collection Keyword Map's time line ...
                        self.MediaLength += (snapshotDuration + 100)
                        # ... and update the current snapshot number so this snapshot won't be counted again if it has multiple keywords
                        currSnapshot = snapshotNum
                    # Decode the KWG, KW, and Snapshot ID
                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                    snapshotID = DBInterface.ProcessDBDataForUTF8Encoding(snapshotID)
                    # If we're dealing with a Collection Keyword Report, self.clipNum will be None and we want all data.
                    if (self.clipNum == None):
                        # Snapshots, unlike Clips, can have the same keyword multiple times.  Let's check to see if this Keyword is already
                        # in the Snapshot List.
                        if (kwg, kw, self.MediaLength - snapshotDuration - 50, self.MediaLength - 50, snapshotNum, snapshotID, collectNum) not in self.snapshotList:
                            # Add the current Snapshot/keyword combo to the Snapshot List, placing it to the right of the last snapshot.
                            self.snapshotList.append((kwg, kw, self.MediaLength - snapshotDuration - 50, self.MediaLength - 50, snapshotNum, snapshotID, collectNum))
                        # If the snapshot ID isn't already in the Snapshot Filter List ...
                        if (not ((snapshotID, collectNum, True) in self.snapshotFilterList)) and \
                           (not ((snapshotID, collectNum, False) in self.snapshotFilterList)):
                            # ... add it to the Snapshot Filter List
                            self.snapshotFilterList.append((snapshotID, collectNum, True))

        # When we're done adding clips, we know the total width of the graphic.  Set self.endTime to the accumulated
        # Media Length so the graphic will render correctly.
        self.endTime = self.MediaLength

    def UpdateKeywordVisualization(self, reset=True):
        """ Update the Keyword Visualization following something that could have changed it.
            The reset variable (when false) allows the Hybrid Visualization's Filter box to work! """
        # If the Keyword Map hasn't been Setup yet, skip this.
        if (self.episodeNum == None) and (self.textObj == None):
            return

        # Before we start, make COPIES of the Quote, Clip, and Snapshot Filter Lists so we can check for Quotes, Clips,
        # and Snapshots that should not be displayed due to FILTER settings
        hideQuoteList = self.quoteFilterList[:]
        hideClipList = self.clipFilterList[:]
        hideSnapshotList = self.snapshotFilterList[:]
        # Before we start, make a COPY of the keyword list so we can check for keywords that are no longer
        # included on the Map and need to be deleted from the KeywordLists
        delKeywordList = self.unfilteredKeywordList[:]

        # if reset is true (always except Hybrid Visualization!) ...
        if reset:
            # Clear the Clip List
            self.clipList = []
            # Clear the Filtered Clip List
            self.clipFilterList = []
            # Clear the Snapshot List
            self.snapshotList = []
            # Clear the Filtered Snapshot List
            self.snapshotFilterList = []
            # Clear the Quote List
            self.quoteList = []
            # Clear the Filtered Quote List
            self.quoteFilterList = []
            
        # Clear the graphic itself (Pass on Hybrid Visualization's reset variable!)
        self.graphic.Clear(reset=reset)

        # If we're dealing with a Media File ...
        if self.MediaLength > 0:

            # Now let's create the SQL to get all relevant Clip and Clip Keyword records
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, cl.ClipStart, cl.ClipStop, cl.ClipNum, cl.ClipID, cl.CollectNum
                           FROM Clips2 cl, ClipKeywords2 ck
                           WHERE cl.EpisodeNum = %s AND
                                 cl.ClipNum = ck.ClipNum
                           ORDER BY ClipStart, cl.ClipNum, KeywordGroup, Keyword"""
            # Adjust the query for sqlite if needed
            SQLText = DBInterface.FixQuery(SQLText)
            # Execute the query
            self.DBCursor.execute(SQLText, (self.episodeNum, ))
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
                    if not (kwg, kw) in self.filteredKeywordList:
                        self.filteredKeywordList.append((kwg, kw))
                    if not (kwg, kw) in self.unfilteredKeywordList:
                        self.unfilteredKeywordList.append((kwg, kw, True))

                # If the keyword is in query results, it should be removed from the list of keywords to be deleted.
                # Check that list for either True or False versions of the keyword!
                if (kwg, kw, True) in delKeywordList:
                    del(delKeywordList[delKeywordList.index((kwg, kw, True))])
                if (kwg, kw, False) in delKeywordList:
                    del(delKeywordList[delKeywordList.index((kwg, kw, False))])

            # Now let's do a pass to see if there are any Clips that should be HIDDEN!
            # For each entry in the previous Filter list ...
            for (tmpClipID, tmpCollectionNum, tmpShowStatus) in hideClipList:
                # ... if that entry was HIDDEN and there is a current entry that is SHOWN ...
                if (tmpShowStatus == False) and ((tmpClipID, tmpCollectionNum, True) in self.clipFilterList):
                    # ... note the position of the entry ...
                    index = self.clipFilterList.index((tmpClipID, tmpCollectionNum, True))
                    # ... and update it to HIDE the Clip
                    self.clipFilterList[index] = (tmpClipID, tmpCollectionNum, False)

            # If we're dealing with an Episode, self.clipNum will be None and we want all clips.
            # If we're dealing with a Clip, we don't deal with Snapshots!
            if (self.clipNum == None):
                # Now let's create the SQL to get all relevant WHOLE SNAPSHOT and Clip Keyword records
                SQLText = """SELECT ck.KeywordGroup, ck.Keyword, sn.SnapshotTimeCode, sn.SnapshotDuration, sn.SnapshotNum, sn.SnapshotID, sn.CollectNum
                               FROM Snapshots2 sn, ClipKeywords2 ck
                               WHERE sn.EpisodeNum = %s AND
                                     sn.SnapshotNum = ck.SnapshotNum
                               ORDER BY SnapshotTimecode, sn.SnapshotNum, KeywordGroup, Keyword"""
                # Adjust the query for sqlite if needed
                SQLText = DBInterface.FixQuery(SQLText)
                # Execute the query
                self.DBCursor.execute(SQLText, (self.episodeNum, ))
                # Iterate through the results ...
                for (kwg, kw, snapshotStart, snapshotDuration, snapshotNum, snapshotID, collectNum) in self.DBCursor.fetchall():
                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                    snapshotID = DBInterface.ProcessDBDataForUTF8Encoding(snapshotID)
                    # If a Snapshot is not found in the snapshotList ...
                    if not ((kwg, kw, snapshotStart, snapshotStart + snapshotDuration, snapshotNum, snapshotID, collectNum) in self.snapshotList):
                        # ... add it to the snapshotList ...
                        self.snapshotList.append((kwg, kw, snapshotStart, snapshotStart + snapshotDuration, snapshotNum, snapshotID, collectNum))
                        # ... and if it's not in the snapshotFilter List (which it probably isn't!) ...
                        if not ((snapshotID, collectNum, True) in self.snapshotFilterList):
                            # ... add it to the snapshotFilterList.
                            self.snapshotFilterList.append((snapshotID, collectNum, True))

                    # If the keyword is not in either of the Keyword Lists, ...
                    if not (((kwg, kw) in self.filteredKeywordList) or ((kwg, kw, False) in self.unfilteredKeywordList)):
                        # ... add it to both keyword lists.
                        if not (kwg, kw) in self.filteredKeywordList:
                            self.filteredKeywordList.append((kwg, kw))
                        if not (kwg, kw) in self.unfilteredKeywordList:
                            self.unfilteredKeywordList.append((kwg, kw, True))

                    # If the keyword is in query results, it should be removed from the list of keywords to be deleted.
                    # Check that list for either True or False versions of the keyword!
                    if (kwg, kw, True) in delKeywordList:
                        del(delKeywordList[delKeywordList.index((kwg, kw, True))])
                    if (kwg, kw, False) in delKeywordList:
                        del(delKeywordList[delKeywordList.index((kwg, kw, False))])

                # Now let's create the SQL to get all relevant SNAPSHOT CODING Keyword records
                SQLText = """SELECT ck.KeywordGroup, ck.Keyword, sn.SnapshotTimeCode, sn.SnapshotDuration, sn.SnapshotNum, sn.SnapshotID, sn.CollectNum
                               FROM Snapshots2 sn, SnapshotKeywords2 ck
                               WHERE sn.EpisodeNum = %s AND
                                     sn.SnapshotNum = ck.SnapshotNum
                               ORDER BY SnapshotTimecode, sn.SnapshotNum, KeywordGroup, Keyword"""
                # Adjust the query for sqlite if needed
                SQLText = DBInterface.FixQuery(SQLText)
                # Execute the query
                self.DBCursor.execute(SQLText, (self.episodeNum, ))
                # Iterate through the results ...
                for (kwg, kw, snapshotStart, snapshotDuration, snapshotNum, snapshotID, collectNum) in self.DBCursor.fetchall():
                    kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                    kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                    snapshotID = DBInterface.ProcessDBDataForUTF8Encoding(snapshotID)
                    # If a Snapshot is not found in the snapshotList ...
                    if not ((kwg, kw, snapshotStart, snapshotStart + snapshotDuration, snapshotNum, snapshotID, collectNum) in self.snapshotList):
                        # ... add it to the snapshotList ...
                        self.snapshotList.append((kwg, kw, snapshotStart, snapshotStart + snapshotDuration, snapshotNum, snapshotID, collectNum))
                        # ... and if it's not in the snapshotFilter List (which it probably isn't!) ...
                        if not ((snapshotID, collectNum, True) in self.snapshotFilterList):
                            # ... add it to the snapshotFilterList.
                            self.snapshotFilterList.append((snapshotID, collectNum, True))

                    # If the keyword is not in either of the Keyword Lists, ...
                    if not (((kwg, kw) in self.filteredKeywordList) or ((kwg, kw, False) in self.unfilteredKeywordList)):
                        # ... add it to both keyword lists.
                        if not (kwg, kw) in self.filteredKeywordList:
                            self.filteredKeywordList.append((kwg, kw))
                        if not (kwg, kw) in self.unfilteredKeywordList:
                            self.unfilteredKeywordList.append((kwg, kw, True))

                    # If the keyword is in query results, it should be removed from the list of keywords to be deleted.
                    # Check that list for either True or False versions of the keyword!
                    if (kwg, kw, True) in delKeywordList:
                        del(delKeywordList[delKeywordList.index((kwg, kw, True))])
                    if (kwg, kw, False) in delKeywordList:
                        del(delKeywordList[delKeywordList.index((kwg, kw, False))])

                # Now let's do a pass to see if there are any Snapshots that should be HIDDEN!
                # For each entry in the previous Filter list ...
                for (tmpSnapshotID, tmpCollectionNum, tmpShowStatus) in hideSnapshotList:
                    # ... if that entry was HIDDEN and there is a current entry that is SHOWN ...
                    if (tmpShowStatus == False) and ((tmpSnapshotID, tmpCollectionNum, True) in self.snapshotFilterList):
                        # ... note the position of the entry ...
                        index = self.snapshotFilterList.index((tmpSnapshotID, tmpCollectionNum, True))
                        # ... and update it to HIDE the Snapshot
                        self.snapshotFilterList[index] = (tmpSnapshotID, tmpCollectionNum, False)

        # If we're dealing with a Text Document ...
        elif self.CharacterLength > 0:
            # Now let's create the SQL to get all relevant Quote and Quote Keyword records
            SQLText = """SELECT ck.KeywordGroup, ck.Keyword, qp.StartChar, qp.EndChar, q.QuoteNum, q.QuoteID, q.CollectNum
                           FROM Quotes2 q, QuotePositions2 qp, ClipKeywords2 ck
                           WHERE q.SourceDocumentNum = %s AND
                                 q.QuoteNum = qp.QuoteNum AND
                                 q.QuoteNum = ck.QuoteNum
                           ORDER BY StartChar, q.QuoteNum, KeywordGroup, Keyword"""

            # Adjust the query for sqlite if needed
            SQLText = DBInterface.FixQuery(SQLText)
            # Execute the query
            self.DBCursor.execute(SQLText, (self.textObj.number, ))
            # Iterate through the results ...
            for (kwg, kw, startChar, endChar, quoteNum, quoteID, collectNum) in self.DBCursor.fetchall():
                kwg = DBInterface.ProcessDBDataForUTF8Encoding(kwg)
                kw = DBInterface.ProcessDBDataForUTF8Encoding(kw)
                quoteID = DBInterface.ProcessDBDataForUTF8Encoding(quoteID)
                # If we're dealing with a Document, self.quoteNum will be None and we want all quotes.
                # If we're dealing with a Quote, we only want to deal with THIS quote!
                if (self.quoteNum == None) or (quoteNum == self.quoteNum):
                    # If a Quote is not found in the quoteList ...
                    if not ((kwg, kw, startChar, endChar, quoteNum, quoteID, collectNum) in self.quoteList):
                        # ... add it to the quoteList ...
                        self.quoteList.append((kwg, kw, self.textObj.quote_dict[quoteNum][0], self.textObj.quote_dict[quoteNum][1], quoteNum, quoteID, collectNum))
                        # ... and if it's not in the quoteFilter List (which it probably isn't!) ...
                        if not ((quoteID, collectNum, True) in self.quoteFilterList):
                            # ... add it to the quoteFilterList.
                            self.quoteFilterList.append((quoteID, collectNum, True))

                # If the keyword is not in either of the Keyword Lists, ...
                if not (((kwg, kw) in self.filteredKeywordList) or ((kwg, kw, False) in self.unfilteredKeywordList)):
                    # ... add it to both keyword lists.
                    if not (kwg, kw) in self.filteredKeywordList:
                        self.filteredKeywordList.append((kwg, kw))
                    if not (kwg, kw) in self.unfilteredKeywordList:
                        self.unfilteredKeywordList.append((kwg, kw, True))

                # If the keyword is in query results, it should be removed from the list of keywords to be deleted.
                # Check that list for either True or False versions of the keyword!
                if (kwg, kw, True) in delKeywordList:
                    del(delKeywordList[delKeywordList.index((kwg, kw, True))])
                if (kwg, kw, False) in delKeywordList:
                    del(delKeywordList[delKeywordList.index((kwg, kw, False))])

            # Now let's do a pass to see if there are any Quotes that should be HIDDEN!
            # For each entry in the previous Filter list ...
            for (tmpQuoteID, tmpCollectionNum, tmpShowStatus) in hideQuoteList:
                # ... if that entry was HIDDEN and there is a current entry that is SHOWN ...
                if (tmpShowStatus == False) and ((tmpQuoteID, tmpCollectionNum, True) in self.quoteFilterList):
                    # ... note the position of the entry ...
                    index = self.quoteFilterList.index((tmpQuoteID, tmpCollectionNum, True))
                    # ... and update it to HIDE the Quote
                    self.quoteFilterList[index] = (tmpQuoteID, tmpCollectionNum, False)

        # Iterate through ANY keywords left in the list of keywords to be deleted ...
        for element in delKeywordList:
            # ... and delete them from the unfiltered Keyword List
            del(self.unfilteredKeywordList[self.unfilteredKeywordList.index(element)])
            # If the keyword is also in the filtered keyword list ...
            if (element[0], element[1]) in self.filteredKeywordList:
                # ... it needs to be deleted from there too!
                del(self.filteredKeywordList[self.filteredKeywordList.index((element[0], element[1]))])

        # Now that the underlying data structures have been corrected, we're ready to redraw the Keyword Visualization
        self.DrawGraph()

        # If we are in an embedded keyword visualization, we may need to apply the Default Configuration ...
        if self.embedded:
            # Trigger the load of the Default filter, if one exists.  An event of None signals we're loading the
            # Default config, and the OnFilter method will handle drawing the graph!
            self.OnFilter(None)

    def DrawGraph(self):
        """ Actually Draw the Keyword Map """
        self.keywordClipList = {}
        # We need to remember Snapshot Color for when self.keywordAsColor is False
        # Otherwise, whole snapshot coding may get a different color than detail snapshot coding.
        snapshotColor = {}

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

        if self.CharacterLength > 0:
            totalLength = self.CharacterLength
            startVal = max(self.startChar, 0)
            endVal = self.endChar
        else:
            totalLength = self.MediaLength
            startVal = max(self.startTime, 0)
            endVal = self.endTime

        if not self.embedded:
            # If we're doing a Document Keyword Map ...
            if self.documentNum != None:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Library: %s'), 'utf8')
                else:
                    prompt = _('Library: %s')
                # If we are in a Right-To-Left Language ...
                if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                    self.graphic.AddTextRight(prompt % self.seriesName, self.Bounds[2] - self.Bounds[0] - 2, 2)
                else:
                    self.graphic.AddText(prompt % self.seriesName, 2, 2)
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("Document: %s"), 'utf8')
                else:
                    prompt = _("Document: %s")
                self.graphic.AddTextCentered(prompt % self.documentName, (self.Bounds[2] - self.Bounds[0]) / 2, 2)
            # If we're doing an Episode Keyword Map ...
            elif self.episodeNum != None:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Library: %s'), 'utf8')
                else:
                    prompt = _('Library: %s')
                # If we are in a Right-To-Left Language ...
                if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                    self.graphic.AddTextRight(prompt % self.seriesName, self.Bounds[2] - self.Bounds[0] - 2, 2)
                else:
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
                try:
                    # If we are in a Right-To-Left Language ...
                    if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                        self.graphic.AddText(prompt % DBInterface.ProcessDBDataForUTF8Encoding(self.MediaFile), 10, 2)
                    else:
                        self.graphic.AddTextRight(prompt % DBInterface.ProcessDBDataForUTF8Encoding(self.MediaFile), self.Bounds[2] - self.Bounds[0] - 20, 2)
                except ValueError:
                    # If we are in a Right-To-Left Language ...
                    if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                        self.graphic.AddText(prompt % self.MediaFile, 10, 2)
                    else:
                        self.graphic.AddTextRight(prompt % self.MediaFile, self.Bounds[2] - self.Bounds[0], 2)
            # If we're doing a Collection Keyword Map, not a Keyword Map ...
            elif self.collectionNum != None:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Collection: %s'), 'utf8')
                else:
                    prompt = _('Collection: %s')
                self.graphic.AddTextCentered(prompt % self.collection.GetNodeString(), (self.Bounds[2] - self.Bounds[0]) / 2, 2)
            if self.configName != '':
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Filter Configuration: %s'), 'utf8')
                else:
                    prompt = _('Filter Configuration: %s')
                # If we are in a Right-To-Left Language ...
                if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                    self.graphic.AddTextRight(prompt % self.configName, self.Bounds[2] - self.Bounds[0] - 2, 16)
                else:
                    self.graphic.AddText(prompt % self.configName, 2, 16)
                
            Count = 0
            # We want Grid Lines in light gray
            self.graphic.SetColour('LIGHT GREY')
            # Draw the top Grid Line, if appropriate
            if self.hGridLines:
                # If we are in a Right-To-Left Language ...
                if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                    self.graphic.AddLines([(self.Bounds[2] - self.Bounds[0] - 10, self.CalcY(-1) + 6 + int(self.whitespaceHeight / 2),
                                            self.CalcX(endVal), self.CalcY(-1) + 6 + int(self.whitespaceHeight / 2))])
                else:
                    self.graphic.AddLines([(10, self.CalcY(-1) + 6 + int(self.whitespaceHeight / 2), self.CalcX(endVal), self.CalcY(-1) + 6 + int(self.whitespaceHeight / 2))])

            for KWG, KW in self.filteredKeywordList:
                # If we are in a Right-To-Left Language ...
                if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                    tmpX = self.Bounds[2] - self.Bounds[0] - 10
                    self.graphic.AddTextRight("%s : %s" % (KWG, KW), tmpX, self.CalcY(Count) - 7)
                else:
                    tmpX = 10
                    self.graphic.AddText("%s : %s" % (KWG, KW), tmpX, self.CalcY(Count) - 7)

                # Add Horizontal Grid Lines, if appropriate
                if self.hGridLines and (Count % 2 == 1):
                    # If we are in a Right-To-Left Language ...
                    if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                        self.graphic.AddLines([(self.Bounds[2] - self.Bounds[0] - 10, self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2),
                                                self.CalcX(endVal), self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2))])
                    else:
                        self.graphic.AddLines([(10, self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2), self.CalcX(endVal), self.CalcY(Count) + 6 + int(self.whitespaceHeight / 2))])

                Count = Count + 1
            # Reset the graphic color following drawing the Grid Lines
            self.graphic.SetColour("BLACK")

            # We need to skip the Title lines in determining the Graphic Indent.
            # If we have a Collection Number ...
            if self.collectionNum > 0:
                # ... the Collection Keyword Map has a single Title element
                start = 1
            # if we have a Document Number ...
            elif self.documentNum > 0:
                # ... the Docuemnt Keyword Map has two Title Elements
                start = 2
            # If we DO NOT have a Collection Number ...
            else:
                # ... the Episode Keyword Map has three Title elements
                start = 3
            # If we have a Configuration loaded ...
            if self.configName != '':
                # ... then we have one additional Title element
                start += 1
            # Determine the max width of the Keywords (by skipping the correct number of Title elements!)
            self.graphicindent = self.graphic.GetMaxWidth(start=start)

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
                        # If we are in a Right-To-Left Language ...
                        if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                            self.graphic.AddLines([(self.Bounds[2] - self.Bounds[0], self.CalcY(Count) + 3 + int(self.whitespaceHeight / 2),
                                                    self.CalcX(endVal), self.CalcY(Count) + 3 + int(self.whitespaceHeight / 2))])
                        else:
                            self.graphic.AddLines([(0, self.CalcY(Count) + 3 + int(self.whitespaceHeight / 2), self.CalcX(endVal), self.CalcY(Count) + 3 + int(self.whitespaceHeight / 2))])

                    Count = Count + 1
                # Reset the graphic color following drawing the Grid Lines
                self.graphic.SetColour("BLACK")

        if not self.embedded:

            # If the Media Length is known, display the Time Line
            if totalLength > 0:
                self.graphic.SetThickness(3)
                self.graphic.AddLines([(self.CalcX(startVal), self.CalcY(-2), self.CalcX(endVal), self.CalcY(-2))])
                # Add Time markers
                self.graphic.SetThickness(1)
                if 'wxMac' in wx.PlatformInfo:
                    self.graphic.SetFontSize(11)
                else:
                    self.graphic.SetFontSize(8)

                X = startVal
                self.graphic.AddLines([(self.CalcX(X), self.CalcY(-2) + 1, self.CalcX(X), self.CalcY(-2) + 6)])
                if self.documentNum != None:
                    XLabel = "%s" % X
                else:
                    XLabel = Misc.TimeMsToStr(X)
                self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(-2) + 5)
                X = endVal
                self.graphic.AddLines([(self.CalcX(X), self.CalcY(-2) + 1, self.CalcX(X), self.CalcY(-2) + 6)])
                if self.documentNum != None:
                    XLabel = "%s" % X
                else:
                    XLabel = Misc.TimeMsToStr(X)
                self.graphic.AddTextCentered(XLabel, self.CalcX(X), self.CalcY(-2) + 5)

                # Add the first and last Vertical Grid Lines, if appropriate
                if self.vGridLines:
                    # We want Grid Lines in light gray
                    self.graphic.SetColour('LIGHT GREY')
                    self.graphic.AddLines([(self.CalcX(startVal), self.CalcY(0) - 6 - int(self.whitespaceHeight / 2), self.CalcX(startVal), self.CalcY(len(self.filteredKeywordList)) - 6 - int(self.whitespaceHeight / 2))])
                    self.graphic.AddLines([(self.CalcX(endVal), self.CalcY(0) - 6 - int(self.whitespaceHeight / 2), self.CalcX(endVal), self.CalcY(len(self.filteredKeywordList)) - 6 - int(self.whitespaceHeight / 2))])
                    # Reset the graphic color following drawing the Grid Lines
                    self.graphic.SetColour("BLACK")
                (numMarks, interval) = self.GetScaleIncrements(endVal - startVal)
                for loop in range(1, numMarks):
                    X = int(round(float(loop) * interval) + startVal)
                    self.graphic.AddLines([(self.CalcX(X), self.CalcY(-2) + 1, self.CalcX(X), self.CalcY(-2) + 6)])
                    if self.documentNum != None:
                        XLabel = "%s" % X
                    else:
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
                (numMarks, interval) = self.GetScaleIncrements(endVal - startVal)
                for loop in range(1, numMarks):
                    X = int(round(float(loop) * interval) + startVal)
                    self.graphic.AddLines([(self.CalcX(X) - 1, self.CalcY(0) - int(self.whitespaceHeight / 2), self.CalcX(X) - 1, self.CalcY(len(self.filteredKeywordList)) - int(self.whitespaceHeight / 2))])
                # Reset the graphic color following drawing the Grid Lines
                self.graphic.SetColour("BLACK")

        colourindex = self.keywordColors['lastColor']
        lastclip = 0
        lastsnapshot = 0
        lastQuote = 0
        # some clip boundary lines for overlapping clips can get over-written, depeding on the nature of the overlaps.
        # Let's create a separate list of these lines, which we'll add to the END of the process so they can't get overwritten.
        overlapLines = []

        # implement colors or gray scale as appropriate
        if self.colorOutput:
            colorSet = TransanaGlobal.keywordMapColourSet
            colorLookup = TransanaGlobal.transana_colorLookup
        else:
            colorSet = TransanaGlobal.keywordMapGraySet
            colorLookup = TransanaGlobal.transana_grayLookup
        # Set the font size, differing by platform
        if 'wxMac' in wx.PlatformInfo:
            self.graphic.SetFontSize(10)
        else:
            self.graphic.SetFontSize(7)

        # Iterate through the keyword list in order ...
        for (KWG, KW) in self.filteredKeywordList:
            # ... and assign colors to Keywords
            # If we want COLOR output ...
            if self.colorOutput:
                # If the color is already defined ...
                if self.keywordColors.has_key((KWG, KW)):
                    # ... get the index for the color
                    colourindex = self.keywordColors[(KWG, KW)]
                # If the color has NOT been defined ...
                else:
                    # Load the keyword
                    tmpKeyword = KeywordObject.Keyword(KWG, KW)
                    # If the Default Keyword Color is in the set of defined colors ...
                    if tmpKeyword.lineColorName in colorSet:
                        # ... define the color for this keyword
                        self.keywordColors[(KWG, KW)] = colorSet.index(tmpKeyword.lineColorName)
                    # If the Default Keyword Color is NOT in the defined colors ...
                    elif tmpKeyword.lineColorName != '':
                        # ... add the color name to the colorSet List
                        colorSet.append(tmpKeyword.lineColorName)
                        # ... add the color's definition to the colorLookup dictionary
                        colorLookup[tmpKeyword.lineColorName] = (int(tmpKeyword.lineColorDef[1:3], 16), int(tmpKeyword.lineColorDef[3:5], 16), int(tmpKeyword.lineColorDef[5:7], 16))
                        # ... determine the new color's index
                        colourindex = colorSet.index(tmpKeyword.lineColorName)
                        # ... define the new color for this keyword
                        self.keywordColors[(KWG, KW)] = colourindex
                    # If there is no Default Keyword Color defined
                    else:
                        # ... get the index for the next color in the color list
                        colourindex = self.keywordColors['lastColor'] + 1
                        # If we're at the end of the list ...
                        if colourindex > len(colorSet) - 1:
                            # ... reset the list to the beginning
                            colourindex = 0
                        # ... remember the color index used
                        self.keywordColors['lastColor'] = colourindex
                        # ... define the new color for this keyword
                        self.keywordColors[(KWG, KW)] = colourindex
            # If we want Grayscale output ...
            else:
                # ... get the index for the next color in the color list
                colourindex = self.keywordColors['lastColor'] + 1
                # If we're at the end of the list ...
                if colourindex > len(colorSet) - 1:
                    # ... reset the list to the beginning
                    colourindex = 0
                # ... remember the color index used
                self.keywordColors['lastColor'] = colourindex
                # ... define the new color for this keyword
                self.keywordColors[(KWG, KW)] = colourindex

            # If we're in the Keyword Visualization and showEmbeddedLabels is enabled ...
            # NOTE:  This is ONLY to be used for testing the mouse-overs, not in production!
            if self.embedded and self.showEmbeddedLabels:
                self.graphic.AddText("%s : %s" % (KWG, KW), 2, self.CalcY(self.filteredKeywordList.index((KWG, KW))) - 7)

        # Set a counter for missing colors
        nextColour = 0
        # For each record in the Clip List ...
        for (KWG, KW, Start, Stop, ClipNum, ClipName, CollectNum) in self.clipList:
            # If the record should be displayed based on the Clip and Keyword sections of the Filter Dialog ...
            if ((ClipName, CollectNum, True) in self.clipFilterList) and ((KWG, KW) in self.filteredKeywordList):
                # See if the Clip's start is before the portion of the map being displayed
                if Start < self.startTime:
                    Start = self.startTime
                if Start > self.endTime:
                    Start = self.endTime
                # See if the Clip's end is after the portion of the map being displayed
                if Stop > self.endTime:
                    Stop = self.endTime
                if Stop < self.startTime:
                    Stop = self.startTime
                # If there's some Clip to be displayed ...
                if Start != Stop:
                    # Determine the line thickness
                    self.graphic.SetThickness(self.barHeight)
                    # Initialize a list for Temporary Lines
                    tempLine = []
                    # Add the Coding Line
                    tempLine.append((self.CalcX(Start), self.CalcY(self.filteredKeywordList.index((KWG, KW))),
                                     self.CalcX(Stop), self.CalcY(self.filteredKeywordList.index((KWG, KW)))))
                    # If we're in the Keyword Map and are NOT using Colors as Keywords (i.e., colors are Clips) ....
                    if (not self.embedded) and (not self.colorAsKeywords):
                        # Update the color index here, at the clip transition
                        if (ClipNum != lastclip) and (lastclip != 0):
                            if colourindex < len(colorSet) - 1:
                                colourindex = colourindex + 1
                            else:
                                colourindex = 0
                    # Otherwise ...
                    else:
                        # ... use the keyword's defined color
                        colourindex = self.keywordColors[(KWG, KW)]

                    # Make sure the colourindex is valid
                    if colourindex > len(colorSet) - 1:
                        # If not, we need to reset the colourindex
                        colourindex = nextColour
                        # replace the keywordColors enty, so the keyword is colored consistently!
                        self.keywordColors[(KWG, KW)] = colourindex
                        # Increment the next color counter
                        nextColour += 1
                        # If nextColour gets too large, reset to 0
                        if nextColour >= len(colorSet):
                            nextColour = 0

                    # Set the Color of the line to be drawn
                    self.graphic.SetColour(colorLookup[colorSet[colourindex]])
                    # Add this line to the graphic
                    self.graphic.AddLines(tempLine)

                # Note what Clip is being processed at the moment
                lastclip = ClipNum

                # Now add the Clip to the keywordClipList.  This holds all Keyword/Clip data in memory so it can be searched quickly
                # This dictionary object uses the keyword pair as the key and holds a list of Clip data for all clips with that keyword.
                # If the list for a given keyword already exists ...
                if self.keywordClipList.has_key((KWG, KW)):
                    # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                    overlapClips = self.keywordClipList[(KWG, KW)]
                    # Iterate through the Clip List ...
                    for (objType, overlapStartTime, overlapEndTime, overlapClipNum, overlapClipName) in overlapClips:
                        # Let's look for overlap
                        overlapStart = Stop
                        overlapEnd = Start

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
                    self.keywordClipList[(KWG, KW)].append(('Clip', Start, Stop, ClipNum, ClipName))

                # If there is no entry for the given keyword ...
                else:
                    # ... create a List object with the first clip's data for this Keyword Pair key
                    self.keywordClipList[(KWG, KW)] = [('Clip', Start, Stop, ClipNum, ClipName)]

        # For each record in the Snapshot List ...
        for (KWG, KW, Start, Stop, SnapshotNum, SnapshotName, CollectNum) in self.snapshotList:
            # If the record should be displayed based on the Snapshot and Keyword sections of the Filter Dialog ...
            if ((SnapshotName, CollectNum, True) in self.snapshotFilterList) and ((KWG, KW) in self.filteredKeywordList):
                # See if the Snapshot's start is before the portion of the map being displayed
                if Start < self.startTime:
                    Start = self.startTime
                if Start > self.endTime:
                    Start = self.endTime
                # See if the Snapshot's end is after the portion of the map being displayed
                if Stop > self.endTime:
                    Stop = self.endTime
                if Stop < self.startTime:
                    Stop = self.startTime
                # If there's some Snapshot to be displayed ...
                if Start != Stop:
                    # Determine the line thickness
                    self.graphic.SetThickness(self.barHeight)
                    # Initialize a list for Temporary Lines
                    tempLine = []
                    # Add the Coding Line
                    tempLine.append((self.CalcX(Start), self.CalcY(self.filteredKeywordList.index((KWG, KW))),
                                     self.CalcX(Stop), self.CalcY(self.filteredKeywordList.index((KWG, KW)))))
                    # If we're in the Keyword Map and are NOT using Colors as Keywords (i.e., colors are Clips) ....
                    if (not self.embedded) and (not self.colorAsKeywords):
                        # Update the color index here, at the clip transition
                        if snapshotColor.has_key(SnapshotNum):
                            colourindex = snapshotColor[SnapshotNum]
                        else:
                            # ... get the index for the next color in the color list
                            colourindex = self.keywordColors['lastColor'] + 1
                            # If we're at the end of the list ...
                            if colourindex > len(colorSet) - 1:
                                # ... reset the list to the beginning
                                colourindex = 0
                            # ... remember the color index used
                            self.keywordColors['lastColor'] = colourindex
                            if colourindex < len(colorSet) - 1:
                                colourindex = colourindex + 1
                            else:
                                colourindex = 0
                            snapshotColor[SnapshotNum] = colourindex
                    # Otherwise ...
                    else:
                        # ... use the keyword's defined color
                        colourindex = self.keywordColors[(KWG, KW)]

                    # Set the Color of the line to be drawn
                    self.graphic.SetColour(colorLookup[colorSet[colourindex]])
                    # Add this line to the graphic
                    self.graphic.AddLines(tempLine)

                # Note what Snapshot is being processed at the moment
                lastsnapshot = SnapshotNum

                # Now add the Snapshot to the keywordClipList.  This holds all Keyword/Clip data in memory so it can be searched quickly
                # This dictionary object uses the keyword pair as the key and holds a list of Clip data for all clips with that keyword.
                # If the list for a given keyword already exists ...
                if self.keywordClipList.has_key((KWG, KW)):
                    # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                    overlapClips = self.keywordClipList[(KWG, KW)]
                    # Iterate through the Clip List ...
                    for (objType, overlapStartTime, overlapEndTime, overlapClipNum, overlapClipName) in overlapClips:
                        # Let's look for overlap
                        overlapStart = Stop
                        overlapEnd = Start

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
                    self.keywordClipList[(KWG, KW)].append(('Snapshot', Start, Stop, SnapshotNum, SnapshotName))

                # If there is no entry for the given keyword ...
                else:
                    # ... create a List object with the first clip's data for this Keyword Pair key
                    self.keywordClipList[(KWG, KW)] = [('Snapshot', Start, Stop, SnapshotNum, SnapshotName)]

        # For each record in the Quote List ...
        for (KWG, KW, Start, Stop, QuoteNum, QuoteName, CollectNum) in self.quoteList:
            # If the record should be displayed based on the Quote and Keyword sections of the Filter Dialog ...
            if ((QuoteName, CollectNum, True) in self.quoteFilterList) and ((KWG, KW) in self.filteredKeywordList):
                # See if the Quote's start is before the portion of the map being displayed
                if Start < self.startChar:
                    Start = self.startChar
                if Start > self.endChar:
                    Start = self.endChar
                # See if the Quote's end is after the portion of the map being displayed
                if Stop > self.endChar:
                    Stop = self.endChar
                if Stop < self.startChar:
                    Stop = self.startChar
                # If there's some Quote to be displayed ...
                if Start != Stop:
                    # Determine the line thickness
                    self.graphic.SetThickness(self.barHeight)
                    # Initialize a list for Temporary Lines
                    tempLine = []

                    # Add the Coding Line
                    tempLine.append((self.CalcX(Start), self.CalcY(self.filteredKeywordList.index((KWG, KW))),
                                     self.CalcX(Stop), self.CalcY(self.filteredKeywordList.index((KWG, KW)))))
                    # If we're in the Keyword Map and are NOT using Colors as Keywords (i.e., colors are Quotes) ....
                    if (not self.embedded) and (not self.colorAsKeywords):
                        # Update the color index here, at the quote transition
                        if (QuoteNum != lastQuote) and (lastQuote != 0):
                            if colourindex < len(colorSet) - 1:
                                colourindex = colourindex + 1
                            else:
                                colourindex = 0
                    # Otherwise ...
                    else:
                        # ... use the keyword's defined color
                        colourindex = self.keywordColors[(KWG, KW)]
                    # Set the Color of the line to be drawn
                    self.graphic.SetColour(colorLookup[colorSet[colourindex]])
                    # Add this line to the graphic
                    self.graphic.AddLines(tempLine)

                # Note what Quote is being processed at the moment
                lastQuote = QuoteNum

                # Now add the Quote to the keywordClipList.  This holds all Keyword/Clip data in memory so it can be searched quickly
                # This dictionary object uses the keyword pair as the key and holds a list of Clip data for all clips with that keyword.
                # If the list for a given keyword already exists ...
                if self.keywordClipList.has_key((KWG, KW)):
                    # Get the list of Quotes that contain the current Keyword from the keyword / Clip List dictionary
                    overlapClips = self.keywordClipList[(KWG, KW)]
                    # Iterate through the Overlap Clip List ...
                    for (objType, overlapStartTime, overlapEndTime, overlapClipNum, overlapClipName) in overlapClips:
                        # Let's look for overlap
                        overlapStart = Stop
                        overlapEnd = Start

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

                    # ... add the new Quote to the Clip List
                    self.keywordClipList[(KWG, KW)].append(('Quote', Start, Stop, QuoteNum, QuoteName))

                # If there is no entry for the given keyword ...
                else:
                    # ... create a List object with the first quote's data for this Keyword Pair key
                    self.keywordClipList[(KWG, KW)] = [('Quote', Start, Stop, QuoteNum, QuoteName)]

        # If we are doing a Keyword Visualization, but there are no Clips in the picture, it can be confusing.
        # Let's place a message on the visualization saying it's intentionally left blank.
        if self.embedded and (lastclip == 0) and (lastsnapshot == 0) and (lastQuote == 0):
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
                # We can't enable Print Preview for Right-To-Left languages
                if not (TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft):
                    self.menuFile.Enable(M_FILE_PRINTPREVIEW, True)
                self.menuFile.Enable(M_FILE_PRINTSETUP, True)
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
            if (time > 0) and (kw != None) and \
               (((self.MediaLength > 0) and (time < self.MediaLength)) or \
                ((self.CharacterLength > 0) and (time < self.CharacterLength))):
                # If we have media-based data ...
                if (self.MediaLength > 0) and (time < self.MediaLength):
                    # ... set a prompt for keyword and time
                    prompt = unicode(_("Keyword:  %s : %s,  Time: %s"), 'utf8')
                    # Set the Status Text to indicate the current Keyword and Time values
                    self.SetStatusText(prompt % (kw[0], kw[1], Misc.time_in_ms_to_str(time)))
                # if we have text-based data ...
                elif (self.CharacterLength > 0) and (time < self.CharacterLength):
                    # ... set a prompt for keyword and position
                    prompt = unicode(_("Keyword:  %s : %s,  Position: %s"), 'utf8')
                    # Set the Status Text to indicate the current Keyword and Position values
                    self.SetStatusText(prompt % (kw[0], kw[1], time))
                # If there's a defined keyword in the Keyword Clip List ...
                if (self.keywordClipList.has_key(kw)):
                    # ... and we have media-based data ...
                    if self.MediaLength > 0:
                        # initialize the string that will hold the names of clips being pointed to
                        clipNames = ''
                        # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                        clips = self.keywordClipList[kw]
                        # Iterate through the Clip List ...
                        for (objType, startTime, endTime, clipNum, clipName) in clips:
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
                    # .. and we have text-based data ...
                    elif self.CharacterLength > 0:
                        # initialize the string that will hold the names of quotes being pointed to.
                        quoteNames = ''
                        # Get the list of Quotes that contain the current Keyword from the keyword / Clip List dictionary
                        quotes = self.keywordClipList[kw]
                        # Iterate through the Quote List ...
                        for (objType, startChar, endChar, quoteNum, quoteName) in quotes:
                            # If the current Character value falls between the Quote's StartChar and EndChar ...
                            if (startChar < time) and (endChar > time):
                                # ... calculate the length of the Quote ...
                                quoteLen = endChar - startChar
                                # First, see if the list is empty.
                                if quoteNames == '':
                                    # If so, just add the keyword name and time
                                    quoteNames = "%s (%s)" % (quoteName, quoteLen)
                                else:
                                    # ... and add the Quote Name and Quote LENGTH to the list of Quotes with this Keyword at this Position
                                    quoteNames += ', ' + "%s (%s)" % (quoteName, quoteLen)
                        # If any quotes are found for the current mouse position ...
                        if (quoteNames != ''):
                            # ... add the KEYWORD names to the ToolTip so they will show up on screen as a hint
                            self.graphic.SetToolTipString('%s : %s  -  %s' % (kw[0], kw[1], quoteNames))
            # If we're not on the data portion of the graph ...
            else:
                # ... set the status text to a blank
                self.SetStatusText('')
        else:
            # We need to call the Visualization Window's "MouseOver" method for updating the cursor's time value
            self.parent.OnMouseOver(x, y, float(x) / self.Bounds[2], float(y) / self.Bounds[3])
            # First, let's make sure we're actually on the data portion of the graph
            if (time >= 0) and (kw != None) and \
               (((self.MediaLength > 0) and (time < self.MediaLength)) or \
                ((self.CharacterLength > 0) and (time < self.CharacterLength)) or \
                ((True))):
                if (self.keywordClipList.has_key(kw)):
                    # If we have a Media File ...
                    if self.MediaLength > 0:
                        # initialize the string that will hold the names of clips being pointed to.
                        # We don't actually need to know the names, but this signals that we're at least OVER a Clip.
                        clipNames = ''
                        # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
                        clips = self.keywordClipList[kw]
                        # Iterate through the Clip List ...
                        for (objType, startTime, endTime, clipNum, clipName) in clips:
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
                    # If we have a text document ...
                    elif self.CharacterLength > 0:
                        # initialize the string that will hold the names of quotes being pointed to.
                        # We don't actually need to know the names, but this signals that we're at least OVER a Quote.
                        quoteNames = ''
                        # Get the list of Quotes that contain the current Keyword from the keyword / Clip List dictionary
                        quotes = self.keywordClipList[kw]
                        # Iterate through the Quote List ...
                        for (objType, startChar, endChar, quoteNum, quoteName) in quotes:
                            # If the current Character value falls between the Quote's StartChar and EndChar ...
                            if (startChar <= time) and (endChar > time):
                                # ... calculate the length of the Quote ...
                                quoteLen = endChar - startChar
                                # ... and add the Quote LENGTH to the list of Quotes with this Keyword at this Position
                                quoteNames += " (%s)" % quoteLen
                            # Handle Orphan Quotes
                            elif (startChar == 0) and (endChar == 1):
                                quoteNames += " (%s)" % 0
                        # If any quotes are found for the current mouse position ...
                        if (quoteNames != ''):
                            # ... add the KEYWORD names to the ToolTip so they will show up on screen as a hint
                            self.graphic.SetToolTipString('%s : %s  -  %s' % (kw[0], kw[1], quoteNames))

    def OnLeftDown(self, event):
        """ Left Mouse Button Down event """
        # Pass the event to the parent
        event.Skip()
        
    def OnLeftUp(self, event):
        """ Left Mouse Button Up event.  Triggers the load of a Quote or Clip. """
        # Note if the Control key is pressed
        ctrlPressed = wx.GetKeyState(wx.WXK_CONTROL)
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
        if self.textObj == None:
            maxVal = self.MediaLength
        else:
            maxVal = self.CharacterLength
        # First, let's make sure we're actually on the data portion of the graph
        if (time > 0) and (time < maxVal) and (kw != None) and (self.keywordClipList.has_key(kw)):
            if 'unicode' in wx.PlatformInfo:
                prompt = unicode(_("Keyword:  %s : %s,  Time: %s"), 'utf8')
            else:
                prompt = _("Keyword:  %s : %s,  Time: %s")
            # Set the Status Text to indicate the current Keyword and Time values
            self.SetStatusText(prompt % (kw[0], kw[1], Misc.time_in_ms_to_str(time)))
            # Get the list of Clips that contain the current Keyword from the keyword / Clip List dictionary
            clips = self.keywordClipList[kw]
            # Iterate through the Clip List ...
            for (objType, startTime, endTime, clipNum, clipName) in clips:
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
                        clipNames["%s (%d)" % (clipName, cnt)] = (objType, clipNum)
                    else:
                        # Add the Clip Name as a Dictionary key pointing to the Clip Number
                        clipNames[clipName] = (objType, clipNum)

        # If only 1 Clip is found ...
        if len(clipNames) == 1:
            # Get the data for the clicked object
            (objType, objNum) = clipNames[clipNames.keys()[0]]
            # ... load that object
            self.parent.KeywordMapLoadItem(objType, objNum, ctrlPressed)
                
            # If left-click, close the Keyword Map.  If not, don't!
            if event.LeftUp():
                # Close the Keyword Map
                self.CloseWindow(event)
        # If more than one Clips are found ..
        elif len(clipNames) > 1:
            # Use a wx.SingleChoiceDialog to allow the user to make the choice between multiple clips here.
            dlg = wx.SingleChoiceDialog(self, _("Which Item would you like to load?"), _("Select an Item"),
                                        clipNames.keys(), wx.CHOICEDLG_STYLE)
            # If the user selects a Clip and click OK ...
            if dlg.ShowModal() == wx.ID_OK:
                # Get the data for the clicked object
                (objType, objNum) = clipNames[dlg.GetStringSelection()]
                # ... load that Quote
                self.parent.KeywordMapLoadItem(objType, objNum, ctrlPressed)
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
