# Copyright (C) 2004 The Board of Regents of the University of Wisconsin System 
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
MENU_FILE_NEW                   =  wx.NewId()  # 101
MENU_FILE_NEWDATABASE           =  wx.NewId()  # 102
MENU_FILE_FILEMANAGEMENT        =  wx.NewId()  # 103
MENU_FILE_SAVE                  =  wx.NewId()  # 104
MENU_FILE_SAVEAS                =  wx.NewId()  # 105
MENU_FILE_PRINTTRANSCRIPT       =  wx.NewId()  # 106
MENU_FILE_PRINTERSETUP          =  wx.NewId()  # 107
MENU_FILE_EXIT                  =  wx.ID_EXIT   # Constant used to improve Mac standardization

# Trasncript Menu
MENU_TRANSCRIPT_EDIT            =  wx.NewId()  # 201
MENU_TRANSCRIPT_EDIT_UNDO       =  wx.NewId()  # 2011
MENU_TRANSCRIPT_EDIT_CUT        =  wx.NewId()  # 2012
MENU_TRANSCRIPT_EDIT_COPY       =  wx.NewId()  # 2013
MENU_TRANSCRIPT_EDIT_PASTE      =  wx.NewId()  # 2014
MENU_TRANSCRIPT_FONT            =  wx.NewId()  # 202
MENU_TRANSCRIPT_PRINT           =  wx.NewId()  # 203
MENU_TRANSCRIPT_PRINTERSETUP    =  wx.NewId()  # 204
MENU_TRANSCRIPT_CHARACTERMAP    =  wx.NewId()  # 205
MENU_TRANSCRIPT_ADJUSTINDEXES   =  wx.NewId()  # 206

# Tools Menu
MENU_TOOLS_FILEMANAGEMENT       =  wx.NewId()  # 301
MENU_TOOLS_IMPORT_DATABASE      =  wx.NewId()  # 302
MENU_TOOLS_EXPORT_DATABASE      =  wx.NewId()  # 303
MENU_TOOLS_CHAT                 =  wx.NewId()  # 304
MENU_TOOLS_BATCHWAVEFORM        =  wx.NewId()  # 305

# Options Menu
MENU_OPTIONS_SETTINGS           =  wx.NewId()  # 401
MENU_OPTIONS_LANGUAGE           =  wx.NewId()  # 402
MENU_OPTIONS_LANGUAGE_EN        =  wx.NewId()  # 4020
MENU_OPTIONS_LANGUAGE_DA        =  wx.NewId()  # 4021  # Danish
MENU_OPTIONS_LANGUAGE_DE        =  wx.NewId()  # 4022  # German
MENU_OPTIONS_LANGUAGE_EL        =  wx.NewId()  # 4023  # Greek
MENU_OPTIONS_LANGUAGE_ES        =  wx.NewId()  # 4024  # Spanish
MENU_OPTIONS_LANGUAGE_FI        =  wx.NewId()  # 4025  # Finnish
MENU_OPTIONS_LANGUAGE_FR        =  wx.NewId()  # 4026  # French
MENU_OPTIONS_LANGUAGE_IT        =  wx.NewId()  # 4027  # Italian
MENU_OPTIONS_LANGUAGE_NL        =  wx.NewId()  # 4028  # Dutch
MENU_OPTIONS_LANGUAGE_PL        =  wx.NewId()  # 4029  # Polish
MENU_OPTIONS_LANGUAGE_RU        =  wx.NewId()  # 4030  # Russian
MENU_OPTIONS_LANGUAGE_SV        =  wx.NewId()  # 4031  # Swedish
# NOTE:  Adding languages?  Don't forget to update the EVT_MENU_RANGE settings.
#        If you scan through MenuSetup.py and MenuWindow.py for language code and add the language for MySQL
#        in DBInterface.InitializeSingleUserDatabase(), you should be all set.
MENU_OPTIONS_WORDTRACK          =  wx.NewId()  # 404
MENU_OPTIONS_AUTOARRANGE        =  wx.NewId()  # 405
MENU_OPTIONS_WAVEFORMQUICKLOAD  =  wx.NewId()  # 406
MENU_OPTIONS_VIDEOSIZE          =  wx.NewId()  # 407
MENU_OPTIONS_VIDEOSIZE_50       =  wx.NewId()  # 4071
MENU_OPTIONS_VIDEOSIZE_66       =  wx.NewId()  # 4072
MENU_OPTIONS_VIDEOSIZE_100      =  wx.NewId()  # 4073
MENU_OPTIONS_VIDEOSIZE_150      =  wx.NewId()  # 4074
MENU_OPTIONS_VIDEOSIZE_200      =  wx.NewId()  # 4075
MENU_OPTIONS_PRESENT            =  wx.NewId()  # 408
MENU_OPTIONS_PRESENT_ALL        =  wx.NewId()  # 4081
MENU_OPTIONS_PRESENT_VIDEO      =  wx.NewId()  # 4082
MENU_OPTIONS_PRESENT_TRANS      =  wx.NewId()  # 4083

# Help Menu
MENU_HELP_MANUAL                =  wx.NewId()  # 501
MENU_HELP_TUTORIAL              =  wx.NewId()  # 502
MENU_HELP_NOTATION              =  wx.NewId()  # 503
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
        self.filemenu.Append(MENU_FILE_EXIT, _("E&xit"))
        self.Append(self.filemenu, _("&File"))

        # Build the Transcript menu
        self.transcriptmenu = wx.Menu()
#        self.transcripteditmenu = wx.Menu()
#        self.transcripteditmenu.Append(MENU_TRANSCRIPT_EDIT_UNDO, _("&Undo\tCtrl-Z"))
#        self.transcripteditmenu.AppendSeparator()
#        self.transcripteditmenu.Append(MENU_TRANSCRIPT_EDIT_CUT, _("Cu&t\tCtrl-X"))
#        self.transcripteditmenu.Append(MENU_TRANSCRIPT_EDIT_COPY, _("&Copy\tCtrl-C"))
#        self.transcripteditmenu.Append(MENU_TRANSCRIPT_EDIT_PASTE, _("&Paste\tCtrl-V"))
#        self.transcriptmenu.AppendMenu(MENU_TRANSCRIPT_EDIT, _("&Edit"), self.transcripteditmenu)
        self.transcriptmenu.Append(MENU_TRANSCRIPT_EDIT_UNDO, _("&Undo\tCtrl-Z"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_EDIT_CUT, _("Cu&t\tCtrl-X"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_EDIT_COPY, _("&Copy\tCtrl-C"))
        self.transcriptmenu.Append(MENU_TRANSCRIPT_EDIT_PASTE, _("&Paste\tCtrl-V"))
        self.transcriptmenu.AppendSeparator()
        self.transcriptmenu.Append(MENU_TRANSCRIPT_FONT, _("&Font"))
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
        self.transcriptmenu.Append(MENU_TRANSCRIPT_ADJUSTINDEXES, _("&Adjust Indexes"))
        self.Append(self.transcriptmenu, _("&Transcript"))

        # Built the Tools menu
        self.toolsmenu = wx.Menu()
        self.toolsmenu.Append(MENU_TOOLS_FILEMANAGEMENT, _("&File Management"))
        self.toolsmenu.Append(MENU_TOOLS_IMPORT_DATABASE, _("&Import Database"))
        self.toolsmenu.Append(MENU_TOOLS_EXPORT_DATABASE, _("&Export Database"))
        if not TransanaConstants.singleUserVersion:
            self.toolsmenu.Append(MENU_TOOLS_CHAT, _("&Chat Window"))
        self.toolsmenu.Append(MENU_TOOLS_BATCHWAVEFORM, _("&Batch Waveform Generator"))
        self.Append(self.toolsmenu, _("Too&ls"))
        
        # Build the Options menu
        self.optionsmenu = wx.Menu()
        self.optionsmenu.Append(MENU_OPTIONS_SETTINGS, _("Program &Settings"))
        self.optionslanguagemenu = wx.Menu()

        # English should always be installed
        self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_EN, _("&English"), kind=wx.ITEM_RADIO)
        # Danish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'da', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_DA, _("&Danish"), kind=wx.ITEM_RADIO)
        # German
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'de', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_DE, _("&German"), kind=wx.ITEM_RADIO)
        # Greek
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'el', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_EL, _("Gree&k"), kind=wx.ITEM_RADIO)
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
        # Italian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'it', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_IT, _("&Italian"), kind=wx.ITEM_RADIO)
        # Dutch
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nl', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_NL, _("D&utch"), kind=wx.ITEM_RADIO)
        # Polish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'pl', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_PL, _("&Polish"), kind=wx.ITEM_RADIO)
        # Russian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'ru', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_RU, _("&Russian"), kind=wx.ITEM_RADIO)
        # Swedish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'sv', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.optionslanguagemenu.Append(MENU_OPTIONS_LANGUAGE_SV, _("S&wedish"), kind=wx.ITEM_RADIO)
        self.optionsmenu.AppendMenu(MENU_OPTIONS_LANGUAGE, _("&Language"), self.optionslanguagemenu)
        self.optionsmenu.AppendSeparator()
        self.optionsmenu.Append(MENU_OPTIONS_WORDTRACK, _("Auto Word-&tracking"), kind=wx.ITEM_CHECK)
        self.optionsmenu.Append(MENU_OPTIONS_AUTOARRANGE, _("&Auto-Arrange"), kind=wx.ITEM_CHECK)
        self.optionsmenu.Append(MENU_OPTIONS_WAVEFORMQUICKLOAD, _("&Waveform Quick-load"), kind=wx.ITEM_CHECK)
        self.optionsmenu.AppendSeparator()
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
        self.optionsmenu.AppendMenu(MENU_OPTIONS_PRESENT, _("&Presentation Mode"), self.optionspresentmenu)
        self.Append(self.optionsmenu, _("&Options"))

        # Set Language Menu to initial value
        if TransanaGlobal.configData.language == 'da':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_DA, True)
        elif TransanaGlobal.configData.language == 'de':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_DE, True)
        elif TransanaGlobal.configData.language == 'el':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_EL, True)
        elif TransanaGlobal.configData.language == 'es':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_ES, True)
        elif TransanaGlobal.configData.language == 'fi':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_FI, True)
        elif TransanaGlobal.configData.language == 'fr':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_FR, True)
        elif TransanaGlobal.configData.language == 'it':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_IT, True)
        elif TransanaGlobal.configData.language == 'nl':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_NL, True)
        elif TransanaGlobal.configData.language == 'pl':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_PL, True)
        elif TransanaGlobal.configData.language == 'ru':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_RU, True)
        elif TransanaGlobal.configData.language == 'sv':
            self.optionslanguagemenu.Check(MENU_OPTIONS_LANGUAGE_SV, True)
            
        # Set Options Menu items to their initial values based on Configuration Data
        self.optionsmenu.Check(MENU_OPTIONS_WORDTRACK, TransanaGlobal.configData.wordTracking)
        self.optionsmenu.Check(MENU_OPTIONS_AUTOARRANGE, TransanaGlobal.configData.autoArrange)
        self.optionsmenu.Check(MENU_OPTIONS_WAVEFORMQUICKLOAD, TransanaGlobal.configData.waveformQuickLoad)
        
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

        # Build the Help Menu
        self.helpmenu = wx.Menu()
        self.helpmenu.Append(MENU_HELP_MANUAL, _("&Manual"))
        self.helpmenu.Append(MENU_HELP_TUTORIAL, _("&Tutorial"))
        self.helpmenu.Append(MENU_HELP_NOTATION, _("Transcript &Notation"))
        self.helpmenu.Append(MENU_HELP_ABOUT, _("&About"))
        self.Append(self.helpmenu, _("&Help"))

        return self
