# Copyright (C) 2003 - 2014 The Board of Regents of the University of Wisconsin System 
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

"""This file implements the TransanaMenuBar class, which defines the contents
   of the main Transana Menus. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx
# Import Transana's Constants
import TransanaConstants
# Import Transana's Global Variables
import TransanaGlobal
# import Python os and sys modules
import os, sys

# Define Menu ID Constants

# File Menu
MENU_FILE_NEW                   =  wx.NewId()
MENU_FILE_NEWDATABASE           =  wx.NewId()
MENU_FILE_FILEMANAGEMENT        =  wx.NewId()
MENU_FILE_SAVE                  =  wx.NewId()
MENU_FILE_SAVEAS                =  wx.NewId()
MENU_FILE_PRINTTRANSCRIPT       =  wx.NewId()
MENU_FILE_PRINTERSETUP          =  wx.NewId()
MENU_FILE_CLOSE_SNAPSHOTS       =  wx.NewId()
MENU_FILE_CLOSE_REPORTS         =  wx.NewId()
MENU_FILE_EXIT                  =  wx.ID_EXIT   # Constant used to improve Mac standardization

# Trasncript Menu
MENU_TRANSCRIPT_EDIT            =  wx.NewId()
MENU_TRANSCRIPT_EDIT_UNDO       =  wx.NewId()
# 2.51 - the Standard IDs don't work on the Mac.  Menu elements don't get enabled or disabled properly!
if 'wxMac' in wx.PlatformInfo:
    MENU_TRANSCRIPT_EDIT_CUT        =  wx.NewId()  # wx.ID_CUT
    MENU_TRANSCRIPT_EDIT_COPY       =  wx.NewId()  # wx.ID_COPY
    MENU_TRANSCRIPT_EDIT_PASTE      =  wx.NewId()  # wx.ID_PASTE
else:
    MENU_TRANSCRIPT_EDIT_CUT        =  wx.ID_CUT
    MENU_TRANSCRIPT_EDIT_COPY       =  wx.ID_COPY
    MENU_TRANSCRIPT_EDIT_PASTE      =  wx.ID_PASTE
MENU_TRANSCRIPT_FONT            =  wx.NewId()
MENU_TRANSCRIPT_PARAGRAPH       =  wx.NewId()
MENU_TRANSCRIPT_TABS            =  wx.NewId()
MENU_TRANSCRIPT_INSERT_IMAGE    =  wx.NewId()
MENU_TRANSCRIPT_PRINT           =  wx.NewId()
MENU_TRANSCRIPT_PRINTERSETUP    =  wx.NewId()
MENU_TRANSCRIPT_CHARACTERMAP    =  wx.NewId()
MENU_TRANSCRIPT_AUTOTIMECODE    =  wx.NewId()
MENU_TRANSCRIPT_ADJUSTINDEXES   =  wx.NewId()
MENU_TRANSCRIPT_TEXT_TIMECODE   =  wx.NewId()

# Tools Menu
MENU_TOOLS_NOTESBROWSER         =  wx.NewId()
MENU_TOOLS_FILEMANAGEMENT       =  wx.NewId()
MENU_TOOLS_MEDIACONVERSION      =  wx.NewId()
MENU_TOOLS_IMPORT_DATABASE      =  wx.NewId()
MENU_TOOLS_EXPORT_DATABASE      =  wx.NewId()
MENU_TOOLS_COLORCONFIG          =  wx.NewId()
MENU_TOOLS_BATCHWAVEFORM        =  wx.NewId()
MENU_TOOLS_CHAT                 =  wx.NewId()
MENU_TOOLS_RECORDLOCK           =  wx.NewId()

# Options Menu
MENU_OPTIONS_SETTINGS           =  wx.ID_PREFERENCES  # Constant used to improve Mac standardization
MENU_OPTIONS_LANGUAGE           =  wx.NewId()
MENU_OPTIONS_LANGUAGE_EN        =  wx.NewId()
MENU_OPTIONS_LANGUAGE_AR        =  wx.NewId()  # Arabic
MENU_OPTIONS_LANGUAGE_DA        =  wx.NewId()  # Danish
MENU_OPTIONS_LANGUAGE_DE        =  wx.NewId()  # German
MENU_OPTIONS_LANGUAGE_EASTEUROPE = wx.NewId()  # Central and Eastern European Encoding (iso-8859-2)
MENU_OPTIONS_LANGUAGE_EL        =  wx.NewId()  # Greek
MENU_OPTIONS_LANGUAGE_ES        =  wx.NewId()  # Spanish
MENU_OPTIONS_LANGUAGE_FI        =  wx.NewId()  # Finnish
MENU_OPTIONS_LANGUAGE_FR        =  wx.NewId()  # French
MENU_OPTIONS_LANGUAGE_HE        =  wx.NewId()  # Hebrew
MENU_OPTIONS_LANGUAGE_IT        =  wx.NewId()  # Italian
MENU_OPTIONS_LANGUAGE_JA        =  wx.NewId()  # Japanese
MENU_OPTIONS_LANGUAGE_KO        =  wx.NewId()  # Korean
MENU_OPTIONS_LANGUAGE_NL        =  wx.NewId()  # Dutch
MENU_OPTIONS_LANGUAGE_NB        =  wx.NewId()  # Norwegian Bokmal
MENU_OPTIONS_LANGUAGE_NN        =  wx.NewId()  # Norwegian Ny-norsk
MENU_OPTIONS_LANGUAGE_PL        =  wx.NewId()  # Polish
MENU_OPTIONS_LANGUAGE_PT        =  wx.NewId()  # Portuguese
MENU_OPTIONS_LANGUAGE_RU        =  wx.NewId()  # Russian
MENU_OPTIONS_LANGUAGE_SV        =  wx.NewId()  # Swedish
MENU_OPTIONS_LANGUAGE_ZH        =  wx.NewId()  # Chinese

# NOTE:  Adding languages?  Don't forget to update the EVT_MENU_RANGE settings.
#        If you scan through MenuSetup.py and MenuWindow.py for language code and add the language for MySQL
#        in DBInterface.InitializeSingleUserDatabase(), set export encoding in ClipDataExport, set GetNewData()
#        and ChangeLanguages() in ControlObjectClass, you should be all set.

if TransanaConstants.macDragDrop or (not 'wxMac' in wx.PlatformInfo):
    MENU_OPTIONS_QUICK_CLIPS    =  wx.NewId()
MENU_OPTIONS_QUICKCLIPWARNING   =  wx.NewId()
MENU_OPTIONS_WORDTRACK          =  wx.NewId()
MENU_OPTIONS_AUTOARRANGE        =  wx.NewId()
# Options Visualization Style menu
MENU_OPTIONS_VISUALIZATION          =  wx.NewId()
MENU_OPTIONS_VISUALIZATION_WAVEFORM =  wx.NewId()
MENU_OPTIONS_VISUALIZATION_KEYWORD  =  wx.NewId()
MENU_OPTIONS_VISUALIZATION_HYBRID   =  wx.NewId()
# Options menu continued
# Options Video Size menu
MENU_OPTIONS_VIDEOSIZE          =  wx.NewId()
MENU_OPTIONS_VIDEOSIZE_50       =  wx.NewId()
MENU_OPTIONS_VIDEOSIZE_66       =  wx.NewId()
MENU_OPTIONS_VIDEOSIZE_100      =  wx.NewId()
MENU_OPTIONS_VIDEOSIZE_150      =  wx.NewId()
MENU_OPTIONS_VIDEOSIZE_200      =  wx.NewId()
# Options for Multiple Monitors
MENU_OPTIONS_MONITOR_BASE       =  wx.NewId()
# Options Presentation Mode menu
MENU_OPTIONS_PRESENT            =  wx.NewId()
MENU_OPTIONS_PRESENT_ALL        =  wx.NewId()
MENU_OPTIONS_PRESENT_VIDEO      =  wx.NewId()
MENU_OPTIONS_PRESENT_TRANS      =  wx.NewId()
MENU_OPTIONS_PRESENT_AUDIO      =  wx.NewId()

# Window Menu
MENU_WINDOW_ALL_MAIN            =  wx.NewId()
MENU_WINDOW_ALL_SNAPSHOT        =  wx.NewId()
MENU_WINDOW_CLOSE_REPORTS       =  wx.NewId()
MENU_WINDOW_DATA                =  wx.NewId()
MENU_WINDOW_VIDEO               =  wx.NewId()
MENU_WINDOW_TRANSCRIPT          =  wx.NewId()
MENU_WINDOW_VISUALIZATION       =  wx.NewId()
MENU_WINDOW_NOTESBROWSER        =  wx.NewId()
MENU_WINDOW_CHAT                =  wx.NewId()

# Help Menu
MENU_HELP_MANUAL                =  wx.ID_HELP   # Constant used to improve Mac standardization
MENU_HELP_TUTORIAL              =  wx.NewId()
MENU_HELP_NOTATION              =  wx.NewId()
MENU_HELP_WEBSITE               =  wx.NewId()
MENU_HELP_FUND                  =  wx.NewId()
MENU_HELP_ABOUT                 =  wx.ID_ABOUT   # Constant used to improve Mac standardization



class MenuSetup(wx.MenuBar):
    """  MenuSetup defines the Menu Structure for Transana's MenuWindow """
    def __init__(self):
        # Create a Menu Bar
        wx.MenuBar.__init__(self)

        # Build the File menu
        self.filemenu = wx.Menu()
        self.filemenu.Append(MENU_FILE_NEW, _("&New"))
        self.filemenu.Append(MENU_FILE_NEWDATABASE, _("&Change Database"))
        self.filemenu.Append(MENU_FILE_FILEMANAGEMENT, _("File &Management"))
        self.filemenu.AppendSeparator()
        self.filemenu.Append(MENU_FILE_SAVE, _("&Save Transcript"))
        self.filemenu.Append(MENU_FILE_SAVEAS, _("Save Transcript &As"))
        self.filemenu.AppendSeparator()
        self.filemenu.Append(MENU_FILE_PRINTTRANSCRIPT, _("&Print Transcript"))
        self.filemenu.Append(MENU_FILE_PRINTERSETUP, _("P&rinter Setup"))
        self.filemenu.AppendSeparator()
        self.filemenu.Append(MENU_FILE_CLOSE_SNAPSHOTS, _("Close All Snapshots"))
        self.filemenu.Append(MENU_FILE_CLOSE_REPORTS, _("Close All Reports"))
        self.filemenu.Append(MENU_FILE_EXIT, _("E&xit"))
        self.Append(self.filemenu, _("&File"))

        # Build the Transcript menu
        self.transcriptmenu = wx.Menu()
        self.transcriptmenu.Append(MENU_TRANSCRIPT_EDIT_UNDO, _("&Undo\tCtrl-Z"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_EDIT_CUT, _("Cu&t\tCtrl-X"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_EDIT_COPY, _("&Copy\tCtrl-C"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_EDIT_PASTE, _("&Paste\tCtrl-V"))
        self.transcriptmenu.AppendSeparator()
        self.transcriptmenu.Append(MENU_TRANSCRIPT_FONT, _("Format &Font"))
        if TransanaConstants.USESRTC:
            self.transcriptmenu.Append(MENU_TRANSCRIPT_PARAGRAPH, _("Format Paragrap&h"))
            self.transcriptmenu.Append(MENU_TRANSCRIPT_TABS, _("Format Ta&bs"))
        self.transcriptmenu.AppendSeparator()
        if TransanaConstants.USESRTC:
            self.transcriptmenu.Append(MENU_TRANSCRIPT_INSERT_IMAGE, _("&Insert Image"))
            self.transcriptmenu.AppendSeparator()
        self.transcriptmenu.Append(MENU_TRANSCRIPT_PRINT, _("&Print"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_PRINTERSETUP, _("P&rinter Setup"))
        self.transcriptmenu.AppendSeparator()
        # The Character Map is in tough shape. 
        # 1.  It doesn't appear to return Font information along with the Character information, rendering it pretty useless
        # 2.  On some platforms, such as XP, it allows the selection of Unicode characters.  But Transana can't cope
        #     with Unicode at this point.
        # Let's just disable it completely for now.
        # self.transcriptmenu.Append(MENU_TRANSCRIPT_CHARACTERMAP, _("&Character Map"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_AUTOTIMECODE, _("F&ixed-Increment Time Codes"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_TEXT_TIMECODE, _("Text Time Code Conversion"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_ADJUSTINDEXES, _("&Adjust Indexes"))
        self.Append(self.transcriptmenu, _("&Transcript"))

        # Built the Tools menu
        self.toolsmenu = wx.Menu()
        self.toolsmenu.Append(MENU_TOOLS_NOTESBROWSER, _("&Notes Browser"))
        self.toolsmenu.Append(MENU_TOOLS_FILEMANAGEMENT, _("&File Management"))
        self.toolsmenu.Append(MENU_TOOLS_MEDIACONVERSION, _("&Media Conversion"))
        self.toolsmenu.Append(MENU_TOOLS_IMPORT_DATABASE, _("&Import Database"))
        self.toolsmenu.Append(MENU_TOOLS_EXPORT_DATABASE, _("&Export Database"))
        self.toolsmenu.Append(MENU_TOOLS_COLORCONFIG, _("&Graphics Color Configuration"))
        self.toolsmenu.Append(MENU_TOOLS_BATCHWAVEFORM, _("&Batch Waveform Generator"))
        if not TransanaConstants.singleUserVersion:
            self.toolsmenu.Append(MENU_TOOLS_CHAT, _("&Chat Window"))
            self.toolsmenu.Append(MENU_TOOLS_RECORDLOCK, _("&Record Lock Utility"))
        self.Append(self.toolsmenu, _("Too&ls"))
        
        # Build the Options menu
        self.optionsmenu = wx.Menu()
        self.optionsmenu.Append(MENU_OPTIONS_SETTINGS, _("Program &Settings"))
        self.optionslanguagemenu = wx.Menu()

        # English should always be installed
        self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_EN, _("&English"), kind=wx.ITEM_RADIO)
        # Arabic
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'ar', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_AR, _("&Arabic"), kind=wx.ITEM_RADIO)
        # Danish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'da', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_DA, _("&Danish"), kind=wx.ITEM_RADIO)
        # German
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'de', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_DE, _("&German"), kind=wx.ITEM_RADIO)
        # Spanish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'es', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_ES, _("&Spanish"), kind=wx.ITEM_RADIO)
        # Finnish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'fi', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_FI, _("Fi&nnish"), kind=wx.ITEM_RADIO)
        # French
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'fr', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_FR, _("&French"), kind=wx.ITEM_RADIO)
        # Hebrew
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'he', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_HE, _("&Hebrew"), kind=wx.ITEM_RADIO)
        # Italian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'it', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_IT, _("&Italian"), kind=wx.ITEM_RADIO)
        # Dutch
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nl', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_NL, _("D&utch"), kind=wx.ITEM_RADIO)
        # Norwegian Bokmal
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nb', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_NB, _("Norwegian Bokmal"), kind=wx.ITEM_RADIO)
        # Norwegian Ny-norsk
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nn', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_NN, _("Norwegian Ny-norsk"), kind=wx.ITEM_RADIO)
        # Polish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'pl', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_PL, _("&Polish"), kind=wx.ITEM_RADIO)
        # Portuguese
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'pt', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_PT, _("P&ortuguese"), kind=wx.ITEM_RADIO)
        # Russian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'ru', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_RU, _("&Russian"), kind=wx.ITEM_RADIO)
        # Swedish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'sv', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_SV, _("S&wedish"), kind=wx.ITEM_RADIO)
        # Chinese
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'zh', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_ZH, _("&Chinese - Simplified"), kind=wx.ITEM_RADIO)
        self.optionsmenu.AppendMenu(MENU_OPTIONS_LANGUAGE, _("&Language"), self.optionslanguagemenu)
        self.optionsmenu.AppendSeparator()
        if TransanaConstants.macDragDrop or (not 'wxMac' in wx.PlatformInfo):
            self.optionsmenu.Append(MENU_OPTIONS_QUICK_CLIPS, _("&Quick Clip Mode"), kind=wx.ITEM_CHECK)
        self.optionsmenu.Append(MENU_OPTIONS_QUICKCLIPWARNING, _("Show Quick Clip Warning"), kind=wx.ITEM_CHECK)
        self.optionsmenu.Append(MENU_OPTIONS_WORDTRACK, _("Auto Word-&tracking"), kind=wx.ITEM_CHECK)
        self.optionsmenu.Append(MENU_OPTIONS_AUTOARRANGE, _("&Auto-Arrange"), kind=wx.ITEM_CHECK)
        
        self.optionsmenu.AppendSeparator()
        
        # Add a menu for the Visualization Options
        self.optionsvisualizationmenu = wx.Menu()
        self.optionsvisualizationmenu.Append(MENU_OPTIONS_VISUALIZATION_WAVEFORM, _("&Waveform"), kind=wx.ITEM_RADIO)
        self.optionsvisualizationmenu.Append(MENU_OPTIONS_VISUALIZATION_KEYWORD, _("&Keyword"), kind=wx.ITEM_RADIO)
        self.optionsvisualizationmenu.Append(MENU_OPTIONS_VISUALIZATION_HYBRID, _("&Hybrid"), kind=wx.ITEM_RADIO)
        self.optionsmenu.AppendMenu(MENU_OPTIONS_VISUALIZATION, _("Vi&sualization Style"), self.optionsvisualizationmenu)
        
        self.optionsvideomenu = wx.Menu()
        self.optionsvideomenu.Append(MENU_OPTIONS_VIDEOSIZE_50, "&50%", kind=wx.ITEM_RADIO)
        self.optionsvideomenu.Append(MENU_OPTIONS_VIDEOSIZE_66, "&66%", kind=wx.ITEM_RADIO)
        self.optionsvideomenu.Append(MENU_OPTIONS_VIDEOSIZE_100, "&100%", kind=wx.ITEM_RADIO)
        self.optionsvideomenu.Append(MENU_OPTIONS_VIDEOSIZE_150, "15&0%", kind=wx.ITEM_RADIO)
        self.optionsvideomenu.Append(MENU_OPTIONS_VIDEOSIZE_200, "&200%", kind=wx.ITEM_RADIO)
        self.optionsmenu.AppendMenu(MENU_OPTIONS_VIDEOSIZE, _("&Video Size"), self.optionsvideomenu)
        self.optionspresentmenu = wx.Menu()
        self.optionspresentmenu.Append(MENU_OPTIONS_PRESENT_ALL, _("&All Windows"), kind=wx.ITEM_RADIO)
        self.optionspresentmenu.Append(MENU_OPTIONS_PRESENT_VIDEO, _("&Video Only"), kind=wx.ITEM_RADIO)
        self.optionspresentmenu.Append(MENU_OPTIONS_PRESENT_TRANS, _("Video and &Transcript Only"), kind=wx.ITEM_RADIO)
        self.optionspresentmenu.Append(MENU_OPTIONS_PRESENT_AUDIO, _("A&udio and Transcript Only"), kind=wx.ITEM_RADIO)
        self.optionsmenu.AppendMenu(MENU_OPTIONS_PRESENT, _("&Presentation Mode"), self.optionspresentmenu)
        self.Append(self.optionsmenu, _("&Options"))

        # Place the "check" by the correct language based on Transana's configuration settings
        self.SetLanguageMenuCheck(TransanaGlobal.configData.language)
            
        # Set Options Menu items to their initial values based on Configuration Data
        if TransanaConstants.macDragDrop or (not 'wxMac' in wx.PlatformInfo):
            self.optionsmenu.Check(MENU_OPTIONS_QUICK_CLIPS, TransanaGlobal.configData.quickClipMode)
        self.optionsmenu.Check(MENU_OPTIONS_QUICKCLIPWARNING, TransanaGlobal.configData.quickClipWarning)
        self.optionsmenu.Check(MENU_OPTIONS_WORDTRACK, TransanaGlobal.configData.wordTracking)
        self.optionsmenu.Check(MENU_OPTIONS_AUTOARRANGE, TransanaGlobal.configData.autoArrange)

        # Set the Visualization Style to reflect the Configuration Data value
        if TransanaGlobal.configData.visualizationStyle == 'Keyword':
            self.optionsvisualizationmenu.Check(MENU_OPTIONS_VISUALIZATION_KEYWORD, True)
        elif TransanaGlobal.configData.visualizationStyle == 'Hybrid':
            self.optionsvisualizationmenu.Check(MENU_OPTIONS_VISUALIZATION_HYBRID, True)
        else:
            self.optionsvisualizationmenu.Check(MENU_OPTIONS_VISUALIZATION_WAVEFORM, True)
        
        # Set the VideoSize Menu to reflect the Configuration Data value
        if TransanaGlobal.configData.videoSize == 50:
            self.optionsmenu.Check(MENU_OPTIONS_VIDEOSIZE_50, True)
        elif TransanaGlobal.configData.videoSize == 66:
            self.optionsmenu.Check(MENU_OPTIONS_VIDEOSIZE_66, True)
        elif TransanaGlobal.configData.videoSize == 100:
            self.optionsmenu.Check(MENU_OPTIONS_VIDEOSIZE_100, True)
        elif TransanaGlobal.configData.videoSize == 150:
            self.optionsmenu.Check(MENU_OPTIONS_VIDEOSIZE_150, True)
        elif TransanaGlobal.configData.videoSize == 200:
            self.optionsmenu.Check(MENU_OPTIONS_VIDEOSIZE_200, True)
            
        # Disable the automatic "Window" menu on the Mac.  It's not accessible for i18n, etc.
        if 'wxMac' in wx.PlatformInfo:
            wx.MenuBar.SetAutoWindowMenu(False)

        # Build the Window Menu
        self.windowMenu = wx.Menu()
        self.windowMenu.Append(MENU_WINDOW_ALL_MAIN, _("Bring Main Windows to Front"))
        self.windowMenu.Append(MENU_WINDOW_ALL_SNAPSHOT, _("Close All Snapshots"))
        self.windowMenu.Append(MENU_WINDOW_CLOSE_REPORTS, _("Close All Reports"))
        self.windowMenu.AppendSeparator()
        self.windowMenu.Append(MENU_WINDOW_DATA, _("Data"))
        self.windowMenu.Append(MENU_WINDOW_VIDEO, _("Video"))
        self.windowMenu.Append(MENU_WINDOW_TRANSCRIPT, _("Transcript"))
        self.windowMenu.Append(MENU_WINDOW_VISUALIZATION, _("Visualization"))
        self.windowMenu.Append(MENU_WINDOW_NOTESBROWSER, _("Notes Browser"))
        if not TransanaConstants.singleUserVersion:
            self.windowMenu.Append(MENU_WINDOW_CHAT, _("Chat"))
        self.windowMenu.AppendSeparator()
        self.Append(self.windowMenu, _("Window"))

        # Build the Help Menu
        self.helpmenu = wx.Menu()
        self.helpmenu.Append(MENU_HELP_MANUAL, _("&Manual"))
        self.helpmenu.Append(MENU_HELP_TUTORIAL, _("&Tutorial"))
        self.helpmenu.Append(MENU_HELP_NOTATION, _("Transcript &Notation"))
        self.helpmenu.Append(MENU_HELP_WEBSITE, _("&www.transana.org"))
        self.helpmenu.Append(MENU_HELP_FUND, _("&Fund Transana"))
        self.helpmenu.Append(MENU_HELP_ABOUT, _("&About"))
        self.Append(self.helpmenu, _("&Help"))

        wx.App_SetMacHelpMenuTitleName(_("&Help"))

    def SetLanguageMenuCheck(self, language):
        """ Set the check mark in the Language menu by the selected language """
        # Set Language Menu to appropriate value
        if language == 'en':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_EN, True)
        elif language == 'ar':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_AR, True)
        elif language == 'da':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_DA, True)
        elif language == 'de':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_DE, True)
        elif language == 'es':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_ES, True)
        elif language == 'fi':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_FI, True)
        elif language == 'fr':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_FR, True)
        elif language == 'it':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_IT, True)
        elif language == 'nl':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_NL, True)
        elif language == 'nb':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_NB, True)
        elif language == 'nn':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_NN, True)
        elif language == 'pl':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_PL, True)
        elif language == 'pt':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_PT, True)
        elif language == 'ru':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_RU, True)
        elif language == 'sv':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_SV, True)
        elif (language == 'zh'):
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_ZH, True)
