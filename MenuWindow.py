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

"""This file handles the Transana Menu Window and all associated logic.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case, Rajas Sambhare'

# Import Python os module
import os
# Import Python sys module
import sys
# Import Python's gettext module
import gettext
# import python's webbrowser module
import webbrowser
# import wxPython
import wx
# Import Transana About Box
import About
# import Transana Database Interface
import DBInterface
# import the Transana Dialogs
import Dialogs
# import Transana File Management System
import FileManagement
# Import Transana Menu Setup
import MenuSetup
# import Transana's Notes Browser
import NotesBrowser
# import Database Import
import XMLImport
# import Database Import
import XMLExport
# import Batch Waveform Generator
import BatchWaveformGenerator
# Import Transana Options Settings 
import OptionsSettings
# Import Transana's Constants
import TransanaConstants
# Import Transana Globals
import TransanaGlobal
# Import the Transcript Printing Module
import TranscriptPrintoutClass
# ONLY if we're using the Multi-user version ...
if not TransanaConstants.singleUserVersion:
    # ... import Transana's ChatWindow
    import ChatWindow
# import Transana Record Lock Utility
import RecordLock

# Language-specific labels for the different languages.  
ENGLISH_LABEL = 'English'
DANISH_LABEL = 'Dansk'
GERMAN_LABEL = 'Deutsch'
GREEK_LABEL = 'English prompts, Greek data'
if 'unicode' in wx.PlatformInfo:
    SPANISH_LABEL = u'Espa\u00f1ol'
else:
    SPANISH_LABEL = 'Espanol'
FINNISH_LABEL = 'Finnish'
if 'unicode' in wx.PlatformInfo:
    FRENCH_LABEL = u'Fran\u00e7ais'
else:
    FRENCH_LABEL = 'Francais'
ITALIAN_LABEL = 'Italiano'
DUTCH_LABEL = 'Nederlands'
if 'unicode' in wx.PlatformInfo:
    NORWEGIAN_BOKMAL_LABEL = u'Norv\u00e9gien Bokm\u00e5l'
    NORWEGIAN_NYNORSK_LABEL = u'Norv\u00e9gien Ny-norsk'
else:
    NORWEGIAN_BOKMAL_LABEL = 'Norvegien Bokmal'
    NORWEGIAN_NYNORSK_LABEL = 'Norvegien Ny-norsk'
POLISH_LABEL = 'Polish'
if 'unicode' in wx.PlatformInfo:
    RUSSIAN_LABEL = u'\u0420\u0443\u0441\u0441\u043a\u0438\u0439'
else:
    RUSSIAN_LABEL = 'Russian'
SWEDISH_LABEL = 'Svenska'
CHINESE_LABEL = 'English prompts, Chinese data'
EASTEUROPE_LABEL = _("English prompts, Eastern European data (ISO-8859-2 encoding)")
JAPANESE_LABEL = 'English prompts, Japanese data'
KOREAN_LABEL = 'English prompts, Korean data'

class MenuWindow(wx.Frame):
    """This class contains the frame object for the Transana Menu Bar window."""

    def __init__(self, parent, id, title):
        """Initialize a MenuBarFrame object."""
        # Initialize the Control Object to None so its absence can be detected
        self.ControlObject = None
        # Initialize the height to be used for the Menu Window.
        self.height = TransanaGlobal.menuHeight

        # We need to handle the window differently on Windows vs. Mac.
        # First, Windows ...
        if '__WXMSW__' in wx.Platform:
            screenDims = wx.ClientDisplayRect()
            self.left = screenDims[0]
            self.top = screenDims[1]
            self.width = screenDims[2] - 2
            self.height = screenDims[3] - 2
            winstyle = wx.MINIMIZE_BOX | wx.CLOSE_BOX | wx.RESIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION      # | wx.MAXIMIZE
            
        # on Mac OS-X ...
        elif '__WXMAC__' in wx.Platform:
            # Ugly hack.  Make the window basically invisible on the mac.
            # Unfortunately this is the cleanest way to achieve the desired
            # result at the moment with wxPython.
            self.left = 1
            self.top = 1
            self.width = 1
            self.height = 1
            winstyle = wx.FRAME_TOOL_WINDOW

        # Linux and who knows what else
        else:
            screenDims = wx.ClientDisplayRect()
            self.left = screenDims[0]
            self.top = screenDims[1]
            self.width = screenDims[2] - 2
            self.height = screenDims[3] - 2
            winstyle = wx.MINIMIZE_BOX | wx.CLOSE_BOX | wx.RESIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION      # | wx.MAXIMIZE

        # Now create the Frame for the Menu Bar
        wx.Frame.__init__(self, parent, -1, title, style=winstyle,
                    size=(self.width, self.height), pos=(self.left, self.top))

        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # If no language has been specified, request an initial language
        if TransanaGlobal.configData.language == '':
            initialLanguage = self.GetLanguage(self)

            if initialLanguage == ENGLISH_LABEL:
                TransanaGlobal.configData.language = 'en'
            elif initialLanguage == DANISH_LABEL:
                TransanaGlobal.configData.language = 'da'
            elif initialLanguage == GERMAN_LABEL:
                TransanaGlobal.configData.language = 'de'
#            elif initialLanguage == GREEK_LABEL:
#                TransanaGlobal.configData.language = 'el'
            elif initialLanguage == SPANISH_LABEL:
                TransanaGlobal.configData.language = 'es'
            elif initialLanguage == FINNISH_LABEL:
                TransanaGlobal.configData.language = 'fi'
            elif initialLanguage == FRENCH_LABEL:
                TransanaGlobal.configData.language = 'fr'
            elif initialLanguage == ITALIAN_LABEL:
                TransanaGlobal.configData.language = 'it'
            elif initialLanguage == DUTCH_LABEL:
                TransanaGlobal.configData.language = 'nl'
            elif initialLanguage == NORWEGIAN_BOKMAL_LABEL:
                TransanaGlobal.configData.language = 'nb'
            elif initialLanguage == NORWEGIAN_NYNORSK_LABEL:
                TransanaGlobal.configData.language = 'nn'
            elif initialLanguage == POLISH_LABEL:
                TransanaGlobal.configData.language = 'pl'
            elif initialLanguage == RUSSIAN_LABEL:
                TransanaGlobal.configData.language = 'ru'
                # The single-user version on Windows needs to set the proper encoding for Russian.
                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
                    TransanaGlobal.encoding = 'koi8_r'
            elif initialLanguage == SWEDISH_LABEL:
                TransanaGlobal.configData.language = 'sv'

            # Chinese, Japanese, and Korean are a special circumstance.  We don't have
            # translations for these languages, but want to be able to allow users to
            # work in these languages.  It's not possible on the Mac for now, and it
            # already works on the Windows MU version.  The Windows single-user version
            # should work if we just add the proper encodings!
            #
            # NOTE:  There are multiple possible encodings for these languages.  I've picked
            #        these at random.
            elif initialLanguage == CHINESE_LABEL:
                TransanaGlobal.configData.language = 'zh'
                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
                    TransanaGlobal.encoding = TransanaConstants.chineseEncoding
            elif initialLanguage == EASTEUROPE_LABEL:
                TransanaGlobal.configData.language = 'easteurope'
                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
                    TransanaGlobal.encoding = 'iso8859_2'
            elif initialLanguage == GREEK_LABEL:
                TransanaGlobal.configData.language = 'el'
                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
                    TransanaGlobal.encoding = 'iso8859_7'
            elif initialLanguage == JAPANESE_LABEL:
                TransanaGlobal.configData.language = 'ja'
                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
                    TransanaGlobal.encoding = 'cp932'
            elif initialLanguage == KOREAN_LABEL:
                TransanaGlobal.configData.language = 'ko'
                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
                    TransanaGlobal.encoding = 'cp949'

        # Okay, a few notes on Internationalization (i18n) are called for here.  It gets a little complicated.
        #
        # There is more than one way to skin a cat.  There is the Python way to internationalize, using "gettext",
        # and there is the wxPython way to internationalize, using wx.Locale.  I've experimented with both, and
        # they both seem to work.  However, the Python way allows me to change languages while Transana is running,
        # while the wxPython way does not.  (It used to, but now it raises an exception.)  On the other hand,
        # the Python way does not affect strings provided by wxPython, while the wxPython way does.
        # (I have yet to identify any such strings on Windows, though I think the wx.MessageDialog on Linux
        # might be an example.)  Some error messages are provided by embedded MySQL, and the language for these 
        # messages cannot be changed once the DB in initialized.  A further wrinkle is that common dialogs that 
        # come from the OS rather than from Python or wxPython, such as the File and Print Dialogs, don't respond 
        # to either method of i18n.  They get their language setting from the OS rather than the program.
        #
        # What I have chosen to do is to implement Transana's i18n using gettext, doing it the Python way so that
        # users will be able to change languages on the fly.  In addition, I set the wx.Locale at program start-up
        # to the initial language selected by the user.  In this case, when a user uses the Language Menu to
        # change languages, all of the prompts supplied by Transana should reflect the change.  However, prompts
        # supplied by wxPython and by MySQL will not change until the user restarts Transana.  At the moment, I
        # don't know what any of these prompts that a user would actually see might be.

        # Install gettext.  Once this is done, all strings enclosed in "_()" will automatically be translated.
        # gettext.install('Transana', 'locale', False)
        # Define supported languages for Transana
        self.presLan_en = gettext.translation('Transana', 'locale', languages=['en']) # English
        # Danish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'da', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_da = gettext.translation('Transana', 'locale', languages=['da']) # Danish
        # German
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'de', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_de = gettext.translation('Transana', 'locale', languages=['de']) # German
        # Greek
#        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'el', 'LC_MESSAGES', 'Transana.mo')
#        if os.path.exists(dir):
#            self.presLan_el = gettext.translation('Transana', 'locale', languages=['el']) # Greek
        # Spanish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'es', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_es = gettext.translation('Transana', 'locale', languages=['es']) # Spanish
        # Finnish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'fi', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_fi = gettext.translation('Transana', 'locale', languages=['fi']) # Finnish
        # French
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'fr', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_fr = gettext.translation('Transana', 'locale', languages=['fr']) # French
        # Italian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'it', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_it = gettext.translation('Transana', 'locale', languages=['it']) # Italian
        # Dutch
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nl', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_nl = gettext.translation('Transana', 'locale', languages=['nl']) # Dutch
        # Norwegian Bokmal
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nb', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_nb = gettext.translation('Transana', 'locale', languages=['nb']) # Norwegian Bokmal
        # Norwegian Ny-norsk
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nn', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_nn = gettext.translation('Transana', 'locale', languages=['nn']) # Norwegian Ny-norsk
        # Polish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'pl', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_pl = gettext.translation('Transana', 'locale', languages=['pl']) # Polish
        # Russian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'ru', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_ru = gettext.translation('Transana', 'locale', languages=['ru']) # Russian
        # Swedish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'sv', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_sv = gettext.translation('Transana', 'locale', languages=['sv']) # Swedish

        # Install English as the initial language if no language has been specified
        # NOTE:  Eastern European Encoding, Greek, Japanese, Korean, and Chinese will use English prompts
        if (TransanaGlobal.configData.language in ['', 'en', 'easteurope', 'el', 'ja', 'ko', 'zh']) :
            lang = wx.LANGUAGE_ENGLISH
            self.presLan_en.install()

        # Danish
        elif (TransanaGlobal.configData.language == 'da'):
            lang = wx.LANGUAGE_DANISH
            self.presLan_da.install()

        # German
        elif (TransanaGlobal.configData.language == 'de'):
            lang = wx.LANGUAGE_GERMAN
            self.presLan_de.install()

        # Greek
#        elif (TransanaGlobal.configData.language == 'el'):
#            lang = wx.LANGUAGE_GREEK     # Greek spec causes an error message on my computer
#            self.presLan_el.install()

        # Spanish
        elif (TransanaGlobal.configData.language == 'es'):
            lang = wx.LANGUAGE_SPANISH
            self.presLan_es.install()

        # Finnish
        elif (TransanaGlobal.configData.language == 'fi'):
            lang = wx.LANGUAGE_FINNISH
            self.presLan_fi.install()

        # French
        elif (TransanaGlobal.configData.language == 'fr'):
            lang = wx.LANGUAGE_FRENCH
            self.presLan_fr.install()

        # Italian
        elif (TransanaGlobal.configData.language == 'it'):
            lang = wx.LANGUAGE_ITALIAN
            self.presLan_it.install()

        # Dutch
        elif (TransanaGlobal.configData.language == 'nl'):
            lang = wx.LANGUAGE_DUTCH
            self.presLan_nl.install()

        # Norwegian Bokmal
        elif (TransanaGlobal.configData.language == 'nb'):
            lang = wx.LANGUAGE_NORWEGIAN_BOKMAL
            self.presLan_nb.install()
            
        # Norwegian Ny-norsk
        elif (TransanaGlobal.configData.language == 'nn'):
            lang = wx.LANGUAGE_NORWEGIAN_NYNORSK
            self.presLan_nn.install()

        # Polish
        elif (TransanaGlobal.configData.language == 'pl'):
            lang = wx.LANGUAGE_POLISH    # Polish spec causes an error message on my computer
            self.presLan_pl.install()

        # Russian
        elif (TransanaGlobal.configData.language == 'ru'):
            lang = wx.LANGUAGE_RUSSIAN   # Russian spec causes an error message on my computer
            self.presLan_ru.install()

        # Swedish
        elif (TransanaGlobal.configData.language == 'sv'):
            lang = wx.LANGUAGE_SWEDISH
            self.presLan_sv.install()

        # Due to a problem with wx.Locale on the Mac (It won't load anything but English), I'm disabling 
        # i18n functionality of the wxPython layer on the Mac.  This code accomplishes that.
#        if "__WXMAC__" in wx.PlatformInfo:
#            lang = wx.LANGUAGE_ENGLISH
            
#            if (TransanaGlobal.configData.language != 'en'):
#                print "wxPython language selection over-ridden for the Mac"
            
        # This provides localization for wxPython
        self.locale = wx.Locale(lang, wx.LOCALE_LOAD_DEFAULT | wx.LOCALE_CONV_ENCODING)
        
        # NOTE:  I've commented out the next line as Transana's i18n will be implemented using Python's
        #        "gettext" rather than wxPython's "wx.Locale".
        self.locale.AddCatalog("Transana")

        transanaIcon = wx.Icon("images/Transana.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(transanaIcon)

        # Build the Menu System using the MenuSetup Object
        self.menuBar = MenuSetup.MenuSetup()

        # Define handler for File > New
        wx.EVT_MENU(self, MenuSetup.MENU_FILE_NEW, self.OnFileNew)
        # Define handler for File > New Database
        wx.EVT_MENU(self, MenuSetup.MENU_FILE_NEWDATABASE, self.OnFileNewDatabase)
        # Define handler for File > File Management
        wx.EVT_MENU(self, MenuSetup.MENU_FILE_FILEMANAGEMENT, self.OnFileManagement)
        # Define handler for File > Save Transcript
        wx.EVT_MENU(self, MenuSetup.MENU_FILE_SAVE, self.OnSaveTranscript)
        # Define handler for File > Save Transcript As
        wx.EVT_MENU(self, MenuSetup.MENU_FILE_SAVEAS, self.OnSaveTranscriptAs)
        # Define handler for File > Print Transcript
        wx.EVT_MENU(self, MenuSetup.MENU_FILE_PRINTTRANSCRIPT, self.OnPrintTranscript)
        # Define handler for File > Printer Setup
        wx.EVT_MENU(self, MenuSetup.MENU_FILE_PRINTERSETUP, self.OnPrinterSetup)
        # Define handler for File > Exit
        wx.EVT_MENU(self, MenuSetup.MENU_FILE_EXIT, self.OnFileExit)

        # Define handler for Transcript > Undo
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_EDIT_UNDO, self.OnTranscriptUndo)
        # Define handler for Transcript > Cut
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_EDIT_CUT, self.OnTranscriptCut)
        # Define handler for Transcript > Copy
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_EDIT_COPY, self.OnTranscriptCopy)
        # Define handler for Transcript > Paste
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_EDIT_PASTE, self.OnTranscriptPaste)
        # Define handler for Transcript > Font
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_FONT, self.OnFont)
        # Define handler for Transcript > Print
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_PRINT, self.OnPrintTranscript)
        # Define handler for Transcript > Printer Setup
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_PRINTERSETUP, self.OnPrinterSetup)

        # Define handler for Transcript > Character Map
        # The Character Map is in tough shape. 
        # 1.  It doesn't appear to return Font information along with the Character information, rendering it pretty useless
        # 2.  On some platforms, such as XP, it allows the selection of Unicode characters.  But Transana can't cope
        #     with Unicode at this point.
        # Let's just disable it completely for now.
        # wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_CHARACTERMAP, self.OnCharacterMap)

        # Define handler for Transcript > Adjust Indexes
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_ADJUSTINDEXES, self.OnAdjustIndexes)

        # Define handler for Tools > Notes Browser
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_NOTESBROWSER, self.OnNotesBrowser)
        # Define handler for Tools > File Management
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_FILEMANAGEMENT, self.OnFileManagement)
        # Define handler for Tools > Import Database
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_IMPORT_DATABASE, self.OnImportDatabase)
        # Define handler for Tools > Export Database
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_EXPORT_DATABASE, self.OnExportDatabase)
        # Define handler for Tools > Batch Waveform Generator
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_BATCHWAVEFORM, self.OnBatchWaveformGenerator)
        # Define handler for Tools > Chat Window
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_CHAT, self.OnChat)
        # Define handler for Tools > Record Lock Utility
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_RECORDLOCK, self.OnRecordLock)

        # Define handler for Options > Settings
        wx.EVT_MENU(self, MenuSetup.MENU_OPTIONS_SETTINGS, self.OnOptionsSettings)
        # Define handler for Options > Language changes
        wx.EVT_MENU_RANGE(self, MenuSetup.MENU_OPTIONS_LANGUAGE_EN, MenuSetup.MENU_OPTIONS_LANGUAGE_ZH, self.OnOptionsLanguage)
        # Define handler for Options > Quick Clip Mode
        if 'wxMSW' in wx.PlatformInfo:
            wx.EVT_MENU(self, MenuSetup.MENU_OPTIONS_QUICK_CLIPS, self.OnOptionsQuickClipMode)
        # Define handler for Options > Auto Word-tracking
        wx.EVT_MENU(self, MenuSetup.MENU_OPTIONS_WORDTRACK, self.OnOptionsWordTrack)
        # Define handler for Options > Auto-Arrange
        wx.EVT_MENU(self, MenuSetup.MENU_OPTIONS_AUTOARRANGE, self.OnOptionsAutoArrange)
        # Define handler for Sound > Waveform Quick-load
        wx.EVT_MENU(self, MenuSetup.MENU_OPTIONS_WAVEFORMQUICKLOAD, self.OnOptionsWaveformQuickload)
        # Define handler for Options > Visualization Style changes
        wx.EVT_MENU_RANGE(self, MenuSetup.MENU_OPTIONS_VISUALIZATION_WAVEFORM, MenuSetup.MENU_OPTIONS_VISUALIZATION_HYBRID, self.OnOptionsVisualizationStyle)
        # Define handler for Options > Video Size changes
        wx.EVT_MENU_RANGE(self, MenuSetup.MENU_OPTIONS_VIDEOSIZE_50, MenuSetup.MENU_OPTIONS_VIDEOSIZE_200, self.OnOptionsVideoSize)

        # Define handler for Help > Manual
        wx.EVT_MENU(self, MenuSetup.MENU_HELP_MANUAL, self.OnHelpManual)
        # Define handler for Help > Tutorial
        wx.EVT_MENU(self, MenuSetup.MENU_HELP_TUTORIAL, self.OnHelpTutorial)
        # Define handler for Help > Transcript Notation
        wx.EVT_MENU(self, MenuSetup.MENU_HELP_NOTATION, self.OnHelpNotation)
        # Define handler for Help > www.transana.org
        wx.EVT_MENU(self, MenuSetup.MENU_HELP_WEBSITE, self.OnHelpWebsite)
        # Define handler for Help > Fund Transana
        wx.EVT_MENU(self, MenuSetup.MENU_HELP_FUND, self.OnHelpFund)
        self.SetMenuBar(self.menuBar)
        # Define handler for Help > About
        wx.EVT_MENU(self, MenuSetup.MENU_HELP_ABOUT, self.OnHelpAbout)

        # We need to block moving the Menu Bar.  This should allow that.
        wx.EVT_MOVE(self, self.OnMove)

        # Define a Close Event, so that if the user click the "X" on the Menu Frame, everything
        # will close properly, including the Video Window.
        wx.EVT_CLOSE(self, self.OnCloseWindow)

        # Initialize Menus by setting them to their initial starting point
        self.ClearMenus()

        # We need to know the actual menu height on Windows, as XP has the funky header option that makes the height an unknown.
        if 'wxMSW' in wx.PlatformInfo:
            # The difference between the actual window size and the Client Size is (surprise!) the height of the
            # Header Bar and the Menu.  This is exactly what we need to know.
            TransanaGlobal.menuHeight = self.GetSizeTuple()[1] - self.GetClientSizeTuple()[1]
            

    def OnMove(self, event):
        # We need to block moving the Menu Bar.  This should allow that, except on Linux, where it causes problems in Gnome
        # (and perhaps elsewhere.)
        if not ('wxGTK' in wx.PlatformInfo):
            self.Move(wx.Point(self.left, self.top))
    
    def OnCloseWindow(self, event):
        """ This code forces the Video Window to close when the "X" is used to close the Menu Bar """
        # Prompt for save if transcript was modified
        self.ControlObject.SaveTranscript(1)

        # Check to see if there are Search Results Nodes
        if self.ControlObject.DataWindowHasSearchNodes():
            # If so, prompt the user about if they really want to exit.
            # Define the Message Dialog
#            dlg = wx.MessageDialog(self, _('You have unsaved Search Results.  Are you sure you want to exit Transana without converting them to Collections?'), _('Transana Confirmation'), wx.YES_NO | wx.ICON_QUESTION)
            # Display the Message Dialog and capture the response
#            result = dlg.ShowModal()
            dlg = Dialogs.QuestionDialog(self, _('You have unsaved Search Results.  Are you sure you want to exit Transana without converting them to Collections?'))
            result = dlg.LocalShowModal()
            # Destroy the Message Dialog
            dlg.Destroy()
        else:
            # If no Search Results exist, it's the same as if the user requests to Exit
            result = wx.ID_YES
        # If the user wants to exit (or if there are no Search Results) ...
        if result == wx.ID_YES:
            # unlock the Transcript Record, if it is locked
            if (self.ControlObject.TranscriptWindow != None) and \
               (self.ControlObject.TranscriptWindow.dlg.editor.TranscriptObj != None) and \
               (self.ControlObject.TranscriptWindow.dlg.editor.TranscriptObj.isLocked):
                self.ControlObject.TranscriptWindow.dlg.editor.TranscriptObj.unlock_record()
            # Close the connection to the Database, if one is open
            if DBInterface.is_db_open():
                DBInterface.close_db()
            # If we have the multi-user version ...
            if not TransanaConstants.singleUserVersion:
                # ... stop the Connection Timer.
                TransanaGlobal.connectionTimer.Stop()
            # Save Configuration Data
            TransanaGlobal.configData.SaveConfiguration()
            # We need to force the Video Window to close along with all of the other windows.
            # (The other windows all close automatically.)
            if self.ControlObject != None:
                if self.ControlObject.VideoWindow != None:
                    self.ControlObject.VideoWindow.frame.Close()
            # Terminate MySQL if using the embedded version.
            # (This is slow, so should be done as late as possible, preferably after windows are closed.)
            if TransanaConstants.singleUserVersion:
                DBInterface.EndSingleUserDatabase()
            # Alternately, if we're in the Multi-user version, we need to close the Chat Window, which
            # ends the socket connection to the Transana MessageServer.
            else:
                # If a Chat Window exists ...
                if TransanaGlobal.chatWindow != None:
                    # Closing the form will cause an expected Socket Loss, which should not be reported.
                    TransanaGlobal.chatWindow.reportSocketLoss = False
                    # ... close it ...
                    TransanaGlobal.chatWindow.OnFormClose(event)
                    # ... and destroy the form so Transana won't hang on closing
                    TransanaGlobal.chatWindow.Destroy()
                    # ... and set the pointer to None
                    TransanaGlobal.chatWindow = None
            # Destroy the Menu Window
            self.Destroy()
        # If the user reconsiders exiting...
        else:
            # If the Close Event is triggered from the Menu, the event has no Veto property.
            # We should only call Veto if this is called by pressing the Frame's Close Control [X].
            if event.GetId() != MenuSetup.MENU_FILE_EXIT:
                # ... Veto (block) the Close event
                event.Veto()

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        # If a Control Object is registered, we need to remember it.
        self.ControlObject=ControlObject

    def ClearMenus(self):
        # Set Menus to their initial default values
        self.menuBar.filemenu.Enable(MenuSetup.MENU_FILE_SAVE, False)              
        # Most Transcript Menu items default to disabled until a Transcript
        # is loaded
        self.SetTranscriptOptions(False)
        self.SetTranscriptEditOptions(False)

    def SetTranscriptOptions(self, enable):
        """Enable or disable the menu options that depend on whether or not
        a Transcript is loaded."""
        self.menuBar.filemenu.Enable(MenuSetup.MENU_FILE_SAVEAS, enable)
        self.menuBar.filemenu.Enable(MenuSetup.MENU_FILE_PRINTTRANSCRIPT, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_PRINT, enable)

    def SetTranscriptEditOptions(self, enable):
        self.menuBar.filemenu.Enable(MenuSetup.MENU_FILE_SAVE, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_EDIT_UNDO, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_EDIT_CUT, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_EDIT_COPY, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_EDIT_PASTE, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_FONT, enable)
        # The Character Map is in tough shape. 
        # 1.  It doesn't appear to return Font information along with the Character information, rendering it pretty useless
        # 2.  On some platforms, such as XP, it allows the selection of Unicode characters.  But Transana can't cope
        #     with Unicode at this point.
        # Let's just disable it completely for now.
        # self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_CHARACTERMAP, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_ADJUSTINDEXES, enable)

    def OnFileNew(self, event):
        """ Implements File New menu command """
        # If a Control Object has been defined ...
        if self.ControlObject != None:
            # ... it should know how to clear all the Windows!
            self.ControlObject.ClearAllWindows()

    def OnFileNewDatabase(self, event):
        """ Implements File > New Database menu command """
        # Check to see if there are Search Results Nodes
        if self.ControlObject.DataWindowHasSearchNodes():
            # If so, prompt the user about if they really want to exit.
            # Define the Message Dialog
#            dlg = wx.MessageDialog(self, _('You have unsaved Search Results.  Are you sure you want to close this database without converting them to Collections?'), _('Transana Confirmation'), wx.YES_NO | wx.ICON_QUESTION)
            # Display the Message Dialog and capture the response
#            result = dlg.ShowModal()
            dlg = Dialogs.QuestionDialog(self, _('You have unsaved Search Results.  Are you sure you want to close this database without converting them to Collections?'))
            result = dlg.LocalShowModal()
            # Destroy the Message Dialog
            dlg.Destroy()
        else:
            # If no Search Results exist, it's the same as if the user says "Yes"
            result = wx.ID_YES
        # If the user wants to exit (or if there are no Search Results) ...
        if result == wx.ID_YES:
            # ... tell it to load a new database
            self.ControlObject.GetNewDatabase()
            # if using MU and we're successfully connected to the database, re-connect to the MessageServer.  
            # This is necessary because you've changed databases, so need to share messages with a different user group.  
            if (DBInterface.is_db_open()) and ((not TransanaConstants.singleUserVersion) and (not ChatWindow.ConnectToMessageServer())):
                # If no connection is made, close Transana!
                self.OnCloseWindow(event)

    def OnFileManagement(self, event):
        """ Implements the FileManagement Menu command """
        # Create a File Management Window
        fileManager = FileManagement.FileManagement(self, -1, _("Transana File Management"))
        # Set up, display, and process the File Management Window
        fileManager.Setup()

    def OnPrintTranscript(self, event):
        """ Implements Transcript Printing from the File and Transcript menus """
        # Set the Cursor to the Hourglass while the report is assembled
        self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
        # Get the Transcript Object currently loaded in the Transcript Window
        tempTranscript = self.ControlObject.GetCurrentTranscriptObject()
        # Prepare the Transcript for printing
        (graphic, pageData) = TranscriptPrintoutClass.PrepareData(TransanaGlobal.printData, tempTranscript)
        # Send the results of the PrepareData() call to the MyPrintout object, once for the print preview
        # version and once for the printer version.  
        printout = TranscriptPrintoutClass.MyPrintout('', graphic, pageData)
        printout2 = TranscriptPrintoutClass.MyPrintout('', graphic, pageData)
        # Create the Print Preview Object
        printPreview = wx.PrintPreview(printout, printout2, TransanaGlobal.printData)
        # Check for errors during Print preview construction
        if not printPreview.Ok():
            # Restore Cursor to Arrow
            self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
            errordlg = Dialogs.ErrorDialog(self, "Print Preview Problem")
            errordlg.ShowModal()
            errordlg.Destroy()
        else:
            # Create the Frame for the Print Preview
            theWidth = max(wx.ClientDisplayRect()[2] - 180, 760)
            theHeight = max(wx.ClientDisplayRect()[3] - 200, 560)
            printFrame = wx.PreviewFrame(printPreview, self, _("Print Preview"), size=(theWidth, theHeight))
            printFrame.Centre()
            # Initialize the Frame for the Print Preview
            printFrame.Initialize()
            # Restore Cursor to Arrow
            self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
            # Display the Print Preview Frame
            printFrame.Show(True)

    def OnPrinterSetup(self, event):
        """ Printer Setup method """
        # Let's use PAGE Setup here ('cause you can do Printer Setup from Page Setup.)  It's a better system
        # that allows Landscape on Mac.

        # Get the global Print Data
        self.printData = TransanaGlobal.printData
        # Create a PageSetupDialogData object based on the global printData defined in __init__
        pageSetupDialogData = wx.PageSetupDialogData(self.printData)
        # Calculate the paper size from the paper ID (Obsolete?)
        pageSetupDialogData.CalculatePaperSizeFromId()
        # Create a Page Setup Dialog based on the Page Setup Dialog Data
        pageDialog = wx.PageSetupDialog(self, pageSetupDialogData)
        # Show the Page Dialog box
        pageDialog.ShowModal()
        # Extract the print data from the page dialog
        self.printData = wx.PrintData(pageDialog.GetPageSetupData().GetPrintData())
        # reflect the print data changes globally
        TransanaGlobal.printData = self.printData
        # Destroy the Page Dialog
        pageDialog.Destroy()


    def OnSaveTranscript(self, event):
        """Handler for File > Save Transcript menu command"""
        self.ControlObject.SaveTranscript()

    def OnSaveTranscriptAs(self, event):
        """Handler for File > Save Transcript As menu command"""
        self.ControlObject.SaveTranscriptAs()

    def OnFileExit(self, evt):
        """ Handler for File > Exit menu command """
        #self.ControlObject.CloseAll()
        self.OnCloseWindow(evt)

    def OnTranscriptUndo(self, event):
        self.ControlObject.TranscriptUndo(event)

    def OnTranscriptCut(self, event):
        self.ControlObject.TranscriptCut()

    def OnTranscriptCopy(self, event):
        self.ControlObject.TranscriptCopy()

    def OnTranscriptPaste(self, event):
        self.ControlObject.TranscriptPaste()

    def OnFont(self, event):
        self.ControlObject.TranscriptCallFontDialog()

    def OnCharacterMap(self, event):
        """ Handler for Transcript > Character Map menu command"""
        if wx.Platform == "__WXMSW__":
            import os
            os.system("start charmap.exe")
        elif wx.Platform == "__WXMAC__":
            import platform
            if platform.release() <= '6.8':
                import os
                os.system('Applications/Utilities/Key\ Caps.app/Contents/MacOS/Key\ Caps &')
                # os.system('/System/Library/Components/CharacterPalette.component/Contents/SharedSupport/CharPaletteServer.app/Contents/MacOS/CharPaletteServer &')
            else:
                import Dialogs
                msg = _("Character Map is not available on OS/X 10.3 Panther or greater.")
                errordlg = Dialogs.ErrorDialog(self, msg)
                errordlg.ShowModal()
                errordlg.Destroy()
        else:
            import Dialogs
            msg = _("Character Map not implemented this platform.")
            errordlg = Dialogs.ErrorDialog(self, msg)
            errordlg.ShowModal()
            errordlg.Destroy()

    def OnAdjustIndexes(self, event):
        """ Handler for Transcript > Adjust Indexes menu command """
        msg = _('Please enter the number of seconds by which to adjust the indexes for this transcript.\nValues are accurate to 1/1000 of a second (3 decimal places).')
        dlg = wx.TextEntryDialog(self, msg, _('Adjust Indexes'), '0.000')
        result = dlg.ShowModal()
        if result == wx.ID_OK:
            try:
                adjustValue = float(dlg.GetValue())
                self.ControlObject.AdjustIndexes(adjustValue)
            except:
                import traceback
                traceback.print_exc(file=sys.stdout)
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Error in Adjust Indexes.\n%s\n%s'), 'utf8')
                else:
                    prompt = _('Error in Adjust Indexes.\n%s\n%s')
                errordlg = Dialogs.ErrorDialog(self, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
            
        dlg.Destroy()

    def OnNotesBrowser(self, event):
        """ Notes Browser """
        # Instantiate a Notes Browser window
        notesBrowser = NotesBrowser.NotesBrowser(self, -1, _("Notes Browser"))
        # Register the Control Object with the Notes Browser
        notesBrowser.Register(self.ControlObject)
        # Display the Notes Browser
        notesBrowser.ShowModal()
        # NotesBrowser destroys itself!  No need for notesBrowser.Destroy() here.

    def OnImportDatabase(self, event):
        """ Import Database """
        temp = XMLImport.XMLImport(self, -1, _('Transana XML Import'))
        result = temp.get_input()
        if (result != None) and (result[_("Transana-XML Filename")] != ''):
            temp.Import()
            # If MU, we need to signal other copies that we've imported a database!
            # First, test to see if we're in the Multi-user version.
            if not TransanaConstants.singleUserVersion:
                # Now make sure a Chat Window has been defined
                if TransanaGlobal.chatWindow != None:
                    # Now send the "Import" message
                    TransanaGlobal.chatWindow.SendMessage("I ")
                
        temp.Close()

    def OnExportDatabase(self, event):
        """ Export Database """
        temp = XMLExport.XMLExport(self, -1, _('Transana XML Export'))
        result = temp.get_input()
        if (result != None) and (result[_("Transana-XML Filename")] != ''):
            temp.Export()
        temp.Close()

    def OnChat(self, event):
        """ Chat Window """
        # If a Chat Window has been defined ...
        if TransanaGlobal.chatWindow != None:
            # ... show it!
            TransanaGlobal.chatWindow.Show()
        else:
            dlg = Dialogs.ErrorDialog(None, _("Your connection to the Message Server has been lost.\nYou may have lost your connection to the network, or there may be a problem with the Server.\nPlease quit Transana immediately and resolve the problem."))
            dlg.ShowModal()
            dlg.Destroy()

    def OnRecordLock(self, event):
        """ Record Lock Utility Window """
        # If a Control Object has been defined ...
        if self.ControlObject != None:
            # ... it should know how to clear all the Windows!
            self.ControlObject.ClearAllWindows()
        # Create a Record Lock Utility window
        recordLockWindow = RecordLock.RecordLock(self, -1, _("Transana Record Lock Utility"))
        recordLockWindow.ShowModal()
        recordLockWindow.Destroy()

    def OnBatchWaveformGenerator(self, event):
        """ Batch Waveform Generator """
        temp = BatchWaveformGenerator.BatchWaveformGenerator(self)
        temp.get_input()
        temp.Close()
        temp.Destroy()

    def OnOptionsSettings(self, event):
        """ Handler for Options > Settings """
        # If MU, change Message Servers if necessary.  To do so, let's note what
        # the settings are before the Program Options screen is shown.
        if not TransanaConstants.singleUserVersion:
            messageServer = TransanaGlobal.configData.messageServer
            messageServerPort = TransanaGlobal.configData.messageServerPort
        # Open the Options Settings Dialog Box
        OptionsSettings.OptionsSettings(self)
        # If the video speed was changed ...
        if self.ControlObject.VideoWindow.GetPlayBackSpeed() != TransanaGlobal.configData.videoSpeed/10.0:
            # Change video speed here
            self.ControlObject.VideoWindow.SetPlayBackSpeed(TransanaGlobal.configData.videoSpeed)
        # If MU, if Message Server or Message Server Port is changed, we need to
        # reset the Message Server.
        if not TransanaConstants.singleUserVersion:
            if (messageServer != TransanaGlobal.configData.messageServer) or \
               (messageServerPort != TransanaGlobal.configData.messageServerPort):
                # Attempt to connect to the new MessageServer                    
                if not ChatWindow.ConnectToMessageServer():
                    self.OnCloseWindow(event)
                else:
                    # Now update the Data Window, in case there were changes while we were away.
                    self.ControlObject.DataWindow.DBTab.tree.refresh_tree()

    def OnOptionsLanguage(self, event):
        """ Handler for Options > Language menu selections """
        # English
        if event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_EN:
            TransanaGlobal.configData.language = 'en'
            self.presLan_en.install()
            
        # Danish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_DA:
            TransanaGlobal.configData.language = 'da'
            self.presLan_da.install()
            
        # German
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_DE:
            TransanaGlobal.configData.language = 'de'
            self.presLan_de.install()

        # Greek
#        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_EL:
#            TransanaGlobal.configData.language = 'el'
#            self.presLan_el.install()

        # Spanish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_ES:
            TransanaGlobal.configData.language = 'es'
            self.presLan_es.install()

        # Finnish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_FI:
            TransanaGlobal.configData.language = 'fi'
            self.presLan_fi.install()

        # French
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_FR:
            TransanaGlobal.configData.language = 'fr'
            self.presLan_fr.install()

        # Italian
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_IT:
            TransanaGlobal.configData.language = 'it'
            self.presLan_it.install()

        # Dutch
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_NL:
            TransanaGlobal.configData.language = 'nl'
            self.presLan_nl.install()

        # Norwegian Bokmal
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_NB:
            TransanaGlobal.configData.language = 'nb'
            self.presLan_nb.install()

        # Norwegian Ny-norsk
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_NN:
            TransanaGlobal.configData.language = 'nn'
            self.presLan_nn.install()

        # Polish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_PL:
            TransanaGlobal.configData.language = 'pl'
            self.presLan_pl.install()

        # Russian
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_RU:
            TransanaGlobal.configData.language = 'ru'
            self.presLan_ru.install()

        # Swedish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_SV:
            TransanaGlobal.configData.language = 'sv'
            self.presLan_sv.install()

        # Chinese (English prompts)
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_ZH:
            TransanaGlobal.configData.language = 'zh'
            self.presLan_en.install()

        # Eastern Europe Encoding (English prompts)
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_EASTEUROPE:
            TransanaGlobal.configData.language = 'easteurope'
            self.presLan_en.install()

        # Greek (English prompts)
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_EL:
            TransanaGlobal.configData.language = 'el'
            self.presLan_en.install()

        # Japanese (English prompts)
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_JA:
            TransanaGlobal.configData.language = 'ja'
            self.presLan_en.install()

        # Korean (English prompts)
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_KO:
            TransanaGlobal.configData.language = 'ko'
            self.presLan_en.install()

        else:
            wx.MessageDialog(None, "Unknown Language", "Unknown Language").ShowModal()

            TransanaGlobal.configData.language = 'en'
            self.presLan_en.install()


        infodlg = Dialogs.InfoDialog(None, _("Please note that some prompts cannot be set to the new language until you restart Transana, and\nthat the language of some prompts is determined by your operating system instead of Transana."))
        infodlg.ShowModal()
        infodlg.Destroy()
        
        self.ControlObject.ChangeLanguages()

    def OnOptionsQuickClipMode(self, event):
        """ Handler for Options > Quick Clip Mode """
        # All we need to do is toggle the global value when teh menu option is changed
        TransanaGlobal.configData.quickClipMode = event.IsChecked()

    def OnOptionsAutoArrange(self, event):
        """ Handler for Options > Auto-Arrange """
        # All we need to do is toggle the global value when the menu option is changed
        TransanaGlobal.configData.autoArrange = event.IsChecked()

    def OnOptionsWordTrack(self, event):
        """ Handler for Options > Auto Word-tracking """
        # Set global value appropriately when state changes
        TransanaGlobal.configData.wordTracking = event.IsChecked()

    def OnOptionsWaveformQuickload(self, event):
        """ Handler for Options > Waveform Quick-load menu command """
        TransanaGlobal.configData.waveformQuickLoad = event.IsChecked()

    def OnOptionsVisualizationStyle(self, event):
        """ Handler for Options > Visualization Style menu """
        if event.GetId() == MenuSetup.MENU_OPTIONS_VISUALIZATION_WAVEFORM:
            TransanaGlobal.configData.visualizationStyle = 'Waveform'
        elif event.GetId() == MenuSetup.MENU_OPTIONS_VISUALIZATION_KEYWORD:
            TransanaGlobal.configData.visualizationStyle = 'Keyword'
        if event.GetId() == MenuSetup.MENU_OPTIONS_VISUALIZATION_HYBRID:
            TransanaGlobal.configData.visualizationStyle = 'Hybrid'
        # Change the Visualization to match the new selection
        self.ControlObject.ChangeVisualization()

    def OnOptionsVideoSize(self, event):
        # TODO:  Macintosh needs other options to be explicitly "unchecked" when something here is selected,
        #        and perhaps to have "uncheck current item" blocked.

        # Translate the Menu Selection into the size value Global Variable
        if event.GetId() == MenuSetup.MENU_OPTIONS_VIDEOSIZE_50:
            TransanaGlobal.configData.videoSize = 50
        elif event.GetId() == MenuSetup.MENU_OPTIONS_VIDEOSIZE_66:
            TransanaGlobal.configData.videoSize = 66
        elif event.GetId() == MenuSetup.MENU_OPTIONS_VIDEOSIZE_100:
            TransanaGlobal.configData.videoSize = 100
        elif event.GetId() == MenuSetup.MENU_OPTIONS_VIDEOSIZE_150:
            TransanaGlobal.configData.videoSize = 150
        elif event.GetId() == MenuSetup.MENU_OPTIONS_VIDEOSIZE_200:
            TransanaGlobal.configData.videoSize = 200

        # Trigger the change in size of the Video Component via the Control Object
        self.ControlObject.VideoSizeChange()

    def OnHelpManual(self, evt):
        """ Handler for Help > Manual menu command """
        self.ControlObject.Help('Transana Main Screen')

    def OnHelpTutorial(self, evt):
        """ Handler for Help > Tutorial menu command """
        self.ControlObject.Help('Welcome to the Transana Tutorial')

    def OnHelpNotation(self, evt):
        """ Handler for Help > Notation menu command """
        self.ControlObject.Help('Jeffersonian Transcript Notation')
            
    def OnHelpAbout(self, evt):
        """ Handler for Help > About menu command """
        # Display the About Box
        About.AboutBox()

    def OnHelpWebsite(self, evt):
        """ Handler for Help > www.transana.org menu command """
        # Open the user's browser and display the web site
        webbrowser.open('http://www.transana.org/', new=True)

    def OnHelpFund(self, evt):
        """ Handler for Help > Fund Transana menu command """
        # Open the user's browser and display the funding page
        webbrowser.open('http://www.transana.org/about/funding.htm', new=True)

    def ChangeLanguages(self):
        """ Reset all Menu Labels to reflect a change in selected Language """
        self.menuBar.SetLabelTop(0, _("&File"))
        self.menuBar.filemenu.SetLabel(MenuSetup.MENU_FILE_NEW, _("&New"))
        self.menuBar.filemenu.SetLabel(MenuSetup.MENU_FILE_NEWDATABASE, _("&Change Database"))
        self.menuBar.filemenu.SetLabel(MenuSetup.MENU_FILE_FILEMANAGEMENT, _("File &Management"))
        self.menuBar.filemenu.SetLabel(MenuSetup.MENU_FILE_SAVE, _("&Save Transcript"))
        self.menuBar.filemenu.SetLabel(MenuSetup.MENU_FILE_SAVEAS, _("Save Transcript &As"))
        self.menuBar.filemenu.SetLabel(MenuSetup.MENU_FILE_PRINTTRANSCRIPT, _("&Print Transcript"))
        self.menuBar.filemenu.SetLabel(MenuSetup.MENU_FILE_PRINTERSETUP, _("Printer &Setup"))
        self.menuBar.filemenu.SetLabel(MenuSetup.MENU_FILE_EXIT, _("E&xit"))
        
        self.menuBar.SetLabelTop(1, _("&Transcript"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_EDIT_UNDO, _("&Undo\tCtrl-Z"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_EDIT_CUT, _("Cu&t\tCtrl-X"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_EDIT_COPY, _("&Copy\tCtrl-C"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_EDIT_PASTE, _("&Paste\tCtrl-V"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_FONT, _("&Font"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_PRINT, _("&Print"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_PRINTERSETUP, _("Printer &Setup"))
        # The Character Map is in tough shape. 
        # 1.  It doesn't appear to return Font information along with the Character information, rendering it pretty useless
        # 2.  On some platforms, such as XP, it allows the selection of Unicode characters.  But Transana can't cope
        #     with Unicode at this point.
        # Let's just disable it completely for now.
        # self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_CHARACTERMAP, _("&Character Map"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_ADJUSTINDEXES, _("&Adjust Indexes"))

        self.menuBar.SetLabelTop(2, _("Too&ls"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_NOTESBROWSER, _("&Notes Browser"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_FILEMANAGEMENT, _("&File Management"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_IMPORT_DATABASE, _("&Import Database"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_EXPORT_DATABASE, _("&Export Database"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_BATCHWAVEFORM, _("&Batch Waveform Generator"))
        if not TransanaConstants.singleUserVersion:
            self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_CHAT, _("&Chat Window"))
            self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_RECORDLOCK, _("&Record Lock Utility"))

        self.menuBar.SetLabelTop(3, _("&Options"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_SETTINGS, _("Program &Settings"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE, _("&Language"))
        self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_EN, _("&English"))
        # The Langage menus may not exist, and we should only update them if they do!
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_DA) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_DA, _("&Danish"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_DE) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_DE, _("&German"))
#        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_EL) != None:
#            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_EL, _("Gree&k"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_ES) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_ES, _("&Spanish"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_FI) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_FI, _("Fi&nnish"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_FR) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_FR, _("&French"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_IT) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_IT, _("&Italian"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_NL) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_NL, _("D&utch"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_NB) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_NB, _("Norwegian Bokmal"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_NN) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_NN, _("Norwegian Ny-norsk"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_PL) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_PL, _("&Polish"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_RU) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_RU, _("&Russian"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_SV) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_SV, _("S&wedish"))
        if 'wxMSW' in wx.PlatformInfo:
            self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_QUICK_CLIPS, _("&Quick Clip Mode"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_WORDTRACK, _("Auto &Word-tracking"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_AUTOARRANGE, _("&Auto-Arrange"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_WAVEFORMQUICKLOAD, _("&Waveform Quick-load"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_VISUALIZATION, _("Vi&sualization Style"))
        self.menuBar.optionsvisualizationmenu.SetLabel(MenuSetup.MENU_OPTIONS_VISUALIZATION_WAVEFORM, _("&Waveform"))
        self.menuBar.optionsvisualizationmenu.SetLabel(MenuSetup.MENU_OPTIONS_VISUALIZATION_KEYWORD, _("&Keyword"))
        self.menuBar.optionsvisualizationmenu.SetLabel(MenuSetup.MENU_OPTIONS_VISUALIZATION_HYBRID, _("&Hybrid"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_VIDEOSIZE, _("&Video Size"))
        self.menuBar.optionspresentmenu.SetLabel(MenuSetup.MENU_OPTIONS_PRESENT_ALL, _("&All Windows"))
        self.menuBar.optionspresentmenu.SetLabel(MenuSetup.MENU_OPTIONS_PRESENT_VIDEO, _("&Video Only"))
        self.menuBar.optionspresentmenu.SetLabel(MenuSetup.MENU_OPTIONS_PRESENT_TRANS, _("Video and &Transcript Only"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_PRESENT, _("&Presentation Mode"))

        self.menuBar.SetLabelTop(4, _("&Help"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_MANUAL, _("&Manual"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_TUTORIAL, _("&Tutorial"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_NOTATION, _("Transcript &Notation"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_ABOUT, _("&About"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_WEBSITE, _("&www.transana.org"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_FUND, _("&Fund Transana"))

        # print "Menu Language Changed (%s)" % _("&File")

    def GetLanguage(self, parent):
        """ Determines what languages have been installed and prompts the user to select one. """
        # See what languages are installed on the system and make a list
        languages = [ENGLISH_LABEL]
        # Danish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'da', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(DANISH_LABEL)
        # German
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'de', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(GERMAN_LABEL)
        # Greek
#        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'el', 'LC_MESSAGES', 'Transana.mo')
#        if os.path.exists(dir):
#            languages.append(GREEK_LABEL)
        # Spanish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'es', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(SPANISH_LABEL)
        # Finnish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'fi', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(FINNISH_LABEL)
        # French
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'fr', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(FRENCH_LABEL)
        # Italian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'it', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(ITALIAN_LABEL)
        # Dutch
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nl', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(DUTCH_LABEL)
        # Norwegian Bokmal
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nb', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(NORWEGIAN_BOKMAL_LABEL)
        # Norwegian Ny-norsk
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'nn', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(NORWEGIAN_NYNORSK_LABEL)
        # Polish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'pl', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(POLISH_LABEL)
        # Russian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'ru', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(RUSSIAN_LABEL)
        # Swedish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'sv', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(SWEDISH_LABEL)
        # Easern Europe encoding, Greek, Japanese, Korean, and Chinese
        if ('wxMSW' in wx.PlatformInfo) and TransanaConstants.singleUserVersion:
            languages.append(CHINESE_LABEL)
            languages.append(EASTEUROPE_LABEL)
            languages.append(GREEK_LABEL)
            languages.append(JAPANESE_LABEL)
            # Korean support must be removed due to a bug in wxSTC on Windows.
            # languages.append(KOREAN_LABEL)

        if len(languages) == 1:
            return languages[0]
        else:
        
            dlg = wx.SingleChoiceDialog(parent, "Select a Language:", "Language",
                                        languages,
                                        style= wx.OK | wx.CENTRE)
            dlg.ShowModal()

            result = dlg.GetStringSelection()

            dlg.Destroy()
        
            return result
