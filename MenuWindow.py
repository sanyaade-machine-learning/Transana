# Copyright (C) 2003 - 2012 The Board of Regents of the University of Wisconsin System 
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

# import Python's cStringIO module for fast string processing
import cStringIO
# Import Python os module
import os
# Import Python sys module
import sys
# import Python time module
import time
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
# import the Color Configuration utility
import ColorConfig
# import Batch Waveform Generator, now called the Batch File Processor
import BatchFileProcessor
# Import Transana Options Settings 
import OptionsSettings
# Import Transana's Constants
import TransanaConstants
# Import Transana Globals
import TransanaGlobal
if TransanaConstants.USESRTC:
    import wx.richtext as richtext
    # Import the RTC-based RichTextEditCtrl, needed for printing
    import RichTextEditCtrl_RTC
else:
    # Import the Transcript Printing Module
    import TranscriptPrintoutClass
# ONLY if we're using the Multi-user version ...
if not TransanaConstants.singleUserVersion:
    # ... import Transana's ChatWindow
    import ChatWindow
# import Transana Record Lock Utility
import RecordLock
# import Media Conversion Tool
import MediaConvert

# Language-specific labels for the different languages.  
ENGLISH_LABEL = 'English'
ARABIC_LABEL = 'Arabic'
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
HEBREW_LABEL = 'Hebrew'
ITALIAN_LABEL = 'Italiano'
DUTCH_LABEL = 'Nederlands'
if 'unicode' in wx.PlatformInfo:
    NORWEGIAN_BOKMAL_LABEL = u'Norv\u00e9gien Bokm\u00e5l'
    NORWEGIAN_NYNORSK_LABEL = u'Norv\u00e9gien Ny-norsk'
else:
    NORWEGIAN_BOKMAL_LABEL = 'Norvegien Bokmal'
    NORWEGIAN_NYNORSK_LABEL = 'Norvegien Ny-norsk'
POLISH_LABEL = 'Polish'
PORTUGUESE_LABEL = 'Portuguese'
if 'unicode' in wx.PlatformInfo:
    RUSSIAN_LABEL = u'\u0420\u0443\u0441\u0441\u043a\u0438\u0439'
else:
    RUSSIAN_LABEL = 'Russian'
SWEDISH_LABEL = 'Svenska'
CHINESE_LABEL = 'Chinese - Simplified'
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
            screenDims = wx.Display(0).GetClientArea()  # wx.ClientDisplayRect()
            self.left = screenDims[0]
            self.top = screenDims[1]
            self.width = screenDims[2]
            self.height = screenDims[3]
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
            screenDims = wx.Display(0).GetClientArea()  # wx.ClientDisplayRect()

            print "MenuWindow.__init__():", screenDims, wx.Display.GetCount()
            for x in range(wx.Display.GetCount()):
                d = wx.Display(x)
                print "Display", x, d.GetClientArea(), d.IsPrimary(), d.GetName()
            
            self.left = screenDims[0]
            self.top = screenDims[1]
            self.width = min(screenDims[2], 1440)
            self.height = screenDims[3]
            winstyle = wx.MINIMIZE_BOX | wx.CLOSE_BOX | wx.RESIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION      # | wx.MAXIMIZE

        # Now create the Frame for the Menu Bar
        wx.Frame.__init__(self, parent, -1, title, style=winstyle,
                    size=(self.width, self.height), pos=(self.left, self.top))

        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Initialize Window Layout variables (saved position and size)
        self.menuWindowLayout = None
        self.visualizationWindowLayout = None
        self.videoWindowLayout = None
        self.transcriptWindowLayout = None
        self.dataWindowLayout = None

        # Initialize File Management Window
        self.fileManagementWindow = None

        # Define the Key Down Event Handler
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)

        # If no language has been specified, request an initial language
        if TransanaGlobal.configData.language == '':
            initialLanguage = self.GetLanguage(self)

            if initialLanguage == ENGLISH_LABEL:
                TransanaGlobal.configData.language = 'en'
            elif initialLanguage == ARABIC_LABEL:
                TransanaGlobal.configData.language = 'ar'
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
            elif initialLanguage == HEBREW_LABEL:
                TransanaGlobal.configData.language = 'he'
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
            elif initialLanguage == PORTUGUESE_LABEL:
                TransanaGlobal.configData.language = 'pt'
            elif initialLanguage == RUSSIAN_LABEL:
                TransanaGlobal.configData.language = 'ru'
                # The single-user version on Windows needs to set the proper encoding for Russian.
##                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
##                    TransanaGlobal.encoding = 'koi8_r'
            elif initialLanguage == SWEDISH_LABEL:
                TransanaGlobal.configData.language = 'sv'
            # Chinese
            elif initialLanguage == CHINESE_LABEL:
                TransanaGlobal.configData.language = 'zh'
##                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
##                    TransanaGlobal.encoding = TransanaConstants.chineseEncoding

            # Japanese, and Korean are a special circumstance.  We don't have
            # translations for these languages, but want to be able to allow users to
            # work in these languages.  It's not possible on the Mac for now, and it
            # already works on the Windows MU version.  The Windows single-user version
            # should work if we just add the proper encodings!
            #
            # NOTE:  There are multiple possible encodings for these languages.  I've picked
            #        these at random.
##            elif initialLanguage == EASTEUROPE_LABEL:
##                TransanaGlobal.configData.language = 'easteurope'
##                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
##                    TransanaGlobal.encoding = 'iso8859_2'
##            elif initialLanguage == GREEK_LABEL:
##                TransanaGlobal.configData.language = 'el'
##                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
##                    TransanaGlobal.encoding = 'iso8859_7'
##            elif initialLanguage == JAPANESE_LABEL:
##                TransanaGlobal.configData.language = 'ja'
##                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
##                    TransanaGlobal.encoding = 'cp932'
##            elif initialLanguage == KOREAN_LABEL:
##                TransanaGlobal.configData.language = 'ko'
##                if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
##                    TransanaGlobal.encoding = 'cp949'

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
        # Arabic
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'ar', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_ar = gettext.translation('Transana', 'locale', languages=['ar']) # Arabic
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
        # Hebrew
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'he', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_he = gettext.translation('Transana', 'locale', languages=['he']) # Hebrew
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
        # Portuguese
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'pt', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_pt = gettext.translation('Transana', 'locale', languages=['pt']) # Portuguese
        # Russian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'ru', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_ru = gettext.translation('Transana', 'locale', languages=['ru']) # Russian
        # Swedish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'sv', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_sv = gettext.translation('Transana', 'locale', languages=['sv']) # Swedish
        # Chinese
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'zh', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            self.presLan_zh = gettext.translation('Transana', 'locale', languages=['zh']) # Chinese

        # We're starting to face the situation where not all the translations may be up-to-date.  Let's build some checks.
        # Initialize an empty variable
        outofdateLanguage = ''
        # Set the default prompt, which might get changed
        languageErrorPrompt = "Transana's %s translation is no longer up-to-date.\nMissing prompts will be displayed in English.\n\nIf you are willing to help with this translation,\nplease contact David Woods at dwoods@wcer.wisc.edu." % outofdateLanguage

        # Start exception handling to deal with lost languages
        try:

            # Install English as the initial language if no language has been specified
            # NOTE:  Eastern European Encoding, Greek, Japanese, Korean will use English prompts
            if (TransanaGlobal.configData.language in ['', 'en', 'easteurope', 'el', 'ja', 'ko']) :
                lang = wx.LANGUAGE_ENGLISH
                self.presLan_en.install()

            # Arabic
            elif (TransanaGlobal.configData.language == 'ar'):
                outofdateLanguage = 'Arabic'
                lang = wx.LANGUAGE_ARABIC
                self.presLan_ar.install()

            # Danish
            elif (TransanaGlobal.configData.language == 'da'):
                outofdateLanguage = 'Danish'
                lang = wx.LANGUAGE_DANISH
                self.presLan_da.install()

            # German
            elif (TransanaGlobal.configData.language == 'de'):
                outofdateLanguage = 'German'
                lang = wx.LANGUAGE_GERMAN
                self.presLan_de.install()

            # Greek
#            elif (TransanaGlobal.configData.language == 'el'):
#                outofdateLanguage = 'Greek'
#                lang = wx.LANGUAGE_GREEK
#                self.presLan_el.install()

            # Spanish
            elif (TransanaGlobal.configData.language == 'es'):
                outofdateLanguage = 'Spanish'
                lang = wx.LANGUAGE_SPANISH
                self.presLan_es.install()

            # Finnish
            elif (TransanaGlobal.configData.language == 'fi'):
                outofdateLanguage = 'Finnish'
                lang = wx.LANGUAGE_FINNISH
                self.presLan_fi.install()

            # French
            elif (TransanaGlobal.configData.language == 'fr'):
                outofdateLanguage = 'French'
                lang = wx.LANGUAGE_FRENCH
                self.presLan_fr.install()

            # Hebrew
            elif (TransanaGlobal.configData.language == 'he'):
                outofdateLanguage = 'Hebrew'
                lang = wx.LANGUAGE_HEBREW
                self.presLan_he.install()

            # Italian
            elif (TransanaGlobal.configData.language == 'it'):
                outofdateLanguage = 'Italian'
                lang = wx.LANGUAGE_ITALIAN
                self.presLan_it.install()

            # Dutch
            elif (TransanaGlobal.configData.language == 'nl'):
                outofdateLanguage = 'Dutch'
                lang = wx.LANGUAGE_DUTCH
                self.presLan_nl.install()

            # Norwegian Bokmal
            elif (TransanaGlobal.configData.language == 'nb'):
                outofdateLanguage = 'Norwegian Bokmal'
                # There seems to be a bug in GetText on the Mac when the wxLANGUAGE is set to Bokmal.
                # Setting this to English seems to make little practical difference.
                if 'wxMac' in wx.PlatformInfo:
                    lang = wx.LANGUAGE_ENGLISH
                else:
                    lang = wx.LANGUAGE_NORWEGIAN_BOKMAL
                self.presLan_nb.install()
                
            # Norwegian Ny-norsk
            elif (TransanaGlobal.configData.language == 'nn'):
                outofdateLanguage = 'Norwegian Nynorsk'
                # There seems to be a bug in GetText on the Mac when the wxLANGUAGE is set to Nynorsk.
                # Setting this to English seems to make little practical difference.
                if 'wxMac' in wx.PlatformInfo:
                    lang = wx.LANGUAGE_ENGLISH
                else:
                    lang = wx.LANGUAGE_NORWEGIAN_NYNORSK
                self.presLan_nn.install()

            # Polish
            elif (TransanaGlobal.configData.language == 'pl'):
                outofdateLanguage = 'Polish'
                lang = wx.LANGUAGE_POLISH    # Polish spec causes an error message on my computer
                self.presLan_pl.install()

            # Portuguese
            elif (TransanaGlobal.configData.language == 'pt'):
                outofdateLanguage = 'Portuguese'
                lang = wx.LANGUAGE_PORTUGUESE    # Polish spec causes an error message on my computer
                self.presLan_pt.install()

            # Russian
            elif (TransanaGlobal.configData.language == 'ru'):
                outofdateLanguage = 'Russian'
                lang = wx.LANGUAGE_RUSSIAN   # Russian spec causes an error message on my computer
                self.presLan_ru.install()

            # Swedish
            elif (TransanaGlobal.configData.language == 'sv'):
                outofdateLanguage = 'Swedish'
                lang = wx.LANGUAGE_SWEDISH
                self.presLan_sv.install()

            # Chinese
            elif (TransanaGlobal.configData.language == 'zh'):
                outofdateLanguage = 'Chinese - Simplified'
                lang = wx.LANGUAGE_CHINESE
                self.presLan_zh.install()

        except:
            TransanaGlobal.configData.language = 'en'
            TransanaGlobal.configData.SaveConfiguration()
            lang = wx.LANGUAGE_ENGLISH
            self.presLan_en.install()
            # Change the Language Error Message
            languageErrorPrompt = "Transana's %s translation is no longer up-to-date or available.\nAll prompts will be displayed in English.\n\nIf you are willing to update this translation,\nplease contact David Woods at dwoods@wcer.wisc.edu." % outofdateLanguage

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

        # Check to see if we have a translation, and if it is up-to-date.
        
        # NOTE:  "Fixed-Increment Time Code" works for version 2.42.  "&Media Conversion" works for 2.50.
        # If you update this, also update the phrase
        # below in the OnOptionsLanguage method.)
        
        if (outofdateLanguage != '') and ("&Media Conversion" == _("&Media Conversion")):
            # If not, display an information message.
            dlg = wx.MessageDialog(None, languageErrorPrompt, "Translation update", style=wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()

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
        # Define handler for Transcript > Format Font
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_FONT, self.OnFormatFont)
        # If we're using the Rich Text Control ...
        if TransanaConstants.USESRTC:
            # Define handler for Transcript > Format Paragraph
            wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_PARAGRAPH, self.OnFormatParagraph)
            # Define handler for Transcript > Format Tabs
            wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_TABS, self.OnFormatTabs)
            # Define handler for Transcript > Insert Image
            wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_INSERT_IMAGE, self.OnInsertImage)
        # Define handler for Transcript > Print
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_PRINT, self.OnPrintTranscript)
        # Define handler for Transcript > Printer Setup
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_PRINTERSETUP, self.OnPrinterSetup)

        # Define handler for Transcript > Fixed-Increment Time Codes
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_AUTOTIMECODE, self.OnAutoTimeCode)
        # Define handler for Transcript > Adjust Indexes
        wx.EVT_MENU(self, MenuSetup.MENU_TRANSCRIPT_ADJUSTINDEXES, self.OnAdjustIndexes)

        # Define handler for Tools > Notes Browser
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_NOTESBROWSER, self.OnNotesBrowser)
        # Define handler for Tools > File Management
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_FILEMANAGEMENT, self.OnFileManagement)
        # Define handler for Tools > Media Conversion
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_MEDIACONVERSION, self.OnMediaConversion)
        # Define handler for Tools > Import Database
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_IMPORT_DATABASE, self.OnImportDatabase)
        # Define handler for Tools > Export Database
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_EXPORT_DATABASE, self.OnExportDatabase)
        # Define handler for Tools > Graphics Color Configuration
        wx.EVT_MENU(self, MenuSetup.MENU_TOOLS_COLORCONFIG, self.OnColorConfig)
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
        if TransanaConstants.macDragDrop or (not 'wxMac' in wx.PlatformInfo):
            wx.EVT_MENU(self, MenuSetup.MENU_OPTIONS_QUICK_CLIPS, self.OnOptionsQuickClipMode)
        # Define handler for Options > Show Quick Clip Warning
        wx.EVT_MENU(self, MenuSetup.MENU_OPTIONS_QUICKCLIPWARNING, self.OnOptionsQuickClipWarning)
        # Define handler for Options > Auto Word-tracking
        wx.EVT_MENU(self, MenuSetup.MENU_OPTIONS_WORDTRACK, self.OnOptionsWordTrack)
        # Define handler for Options > Auto-Arrange
        wx.EVT_MENU(self, MenuSetup.MENU_OPTIONS_AUTOARRANGE, self.OnOptionsAutoArrange)
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
            
    def OnKeyDown(self, event):
        """ Handle Key Down Events """
        # See if the ControlObject wants to handle the key that was pressed.
        if self.ControlObject.ProcessCommonKeyCommands(event):
            # If so, we're done here.
            return
        # If we didn't already handle the key ...
        else:
            # ... pass it along to the parent control.  (Leaving this out breaks Menu shortcuts!)
            event.Skip()

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
            dlg = Dialogs.QuestionDialog(self, _('You have unsaved Search Results.  Are you sure you want to exit Transana without converting them to Collections?'))
            # Display the Message Dialog and capture the response
            result = dlg.LocalShowModal()
            # Destroy the Message Dialog
            dlg.Destroy()
        else:
            # If no Search Results exist, it's the same as if the user requests to Exit
            result = wx.ID_YES
        # If the user wants to exit (or if there are no Search Results) ...
        if result == wx.ID_YES:
            # Signal that we're closing Transana.  This allows us to avoid a problem or two on shutdown.
            self.ControlObject.shuttingDown = True
            # See if there's a Notes Browser open
            if self.ControlObject.NotesBrowserWindow != None:
                # If so, close it, which saves anything being edited.
                self.ControlObject.NotesBrowserWindow.Close()
            # unlock the Transcript Records, if any are locked
            for x in range(len(self.ControlObject.TranscriptWindow)):
                # Turn off the Line Number Timer
                self.ControlObject.TranscriptWindow[x].dlg.LineNumTimer.Stop()
                # If the transcript has been modified ...
                if self.ControlObject.TranscriptWindow[x].TranscriptModified():
                    # ... save it!
                    self.ControlObject.SaveTranscript(1, transcriptToSave=x)
                # If the transcript is locked ...
                if (self.ControlObject.TranscriptWindow[x].dlg.editor.TranscriptObj != None) and \
                   (self.ControlObject.TranscriptWindow[x].dlg.editor.TranscriptObj.isLocked):
                    # ... unlock it
                    self.ControlObject.TranscriptWindow[x].dlg.editor.TranscriptObj.unlock_record()
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
                    self.ControlObject.VideoWindow.close()
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

            # If we are running from a BUILD instead of source code ...
            if hasattr(sys, "frozen") or ('wxMac' in wx.PlatformInfo):
                # Put a shutdown indicator in the Error Log
                print "Transana stopped:", time.asctime()
                print

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
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_EDIT_COPY, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_PRINT, enable)

    def SetTranscriptEditOptions(self, enable):
        """Enable or disable the menu options that depend on whether or not
        a Transcript is editable."""
        self.menuBar.filemenu.Enable(MenuSetup.MENU_FILE_SAVE, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_EDIT_UNDO, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_EDIT_CUT, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_EDIT_PASTE, enable)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_FONT, enable)
        if TransanaConstants.USESRTC:
            self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_PARAGRAPH, enable)
            self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_TABS, enable)
            self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_INSERT_IMAGE, enable)
        if enable and self.ControlObject.AutoTimeCodeEnableTest():
            self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_AUTOTIMECODE, enable)
        else:
            self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_AUTOTIMECODE, False)
        self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_ADJUSTINDEXES, enable)

    def OnFileNew(self, event):
        """ Implements File New menu command """
        # If a Control Object has been defined ...
        if self.ControlObject != None:
            # set the active transcript to 0 so multiple transcript will be cleared
            self.ControlObject.activeTranscript = 0
            # ... it should know how to clear all the Windows!
            self.ControlObject.ClearAllWindows()

    def OnFileNewDatabase(self, event):
        """ Implements File > New Database menu command """
        # Check to see if there are Search Results Nodes
        if self.ControlObject.DataWindowHasSearchNodes():
            # If so, prompt the user about if they really want to exit.
            dlg = Dialogs.QuestionDialog(self, _('You have unsaved Search Results.  Are you sure you want to close this database without converting them to Collections?'))
            result = dlg.LocalShowModal()
            # Destroy the Message Dialog
            dlg.Destroy()
        else:
            # If no Search Results exist, it's the same as if the user says "Yes"
            result = wx.ID_YES
        # If the user wants to exit (or if there are no Search Results) ...
        if result == wx.ID_YES:
            # ... if a NotesBrowser window is open ...
            if self.ControlObject.NotesBrowserWindow != None:
                # ... close the Notes Browser Window ...
                self.ControlObject.NotesBrowserWindow.Close()
            # ... tell it to load a new database
            self.ControlObject.GetNewDatabase()
            # if using MU and we're successfully connected to the database, re-connect to the MessageServer.  
            # This is necessary because you've changed databases, so need to share messages with a different user group.  
            if (DBInterface.is_db_open()) and ((not TransanaConstants.singleUserVersion) and (not ChatWindow.ConnectToMessageServer())):
                # If no connection is made, close Transana!
                self.OnCloseWindow(event)

    def OnFileManagement(self, event):
        """ Implements the FileManagement Menu command """
        # if no file management window is defined 
        if self.fileManagementWindow == None:
            # Create a File Management Window
            self.fileManagementWindow = FileManagement.FileManagement(self, -1, _("Transana File Management"))
            # Set up, display, and process the File Management Window
            self.fileManagementWindow.Setup()
        # If a file management window IS defined ...
        else:
            # ... make sure it is visible ...
            self.fileManagementWindow.Show(True)
            # ... and Raise it to the top of other windows
            self.fileManagementWindow.Raise()

    def OnPrintTranscript(self, event):
        """ Implements Transcript Printing from the File and Transcript menus """
        # Get the Transcript Object currently loaded in the Transcript Window
        tempTranscript = self.ControlObject.GetCurrentTranscriptObject()
        # If we're using the RTC ...
        if TransanaConstants.USESRTC:
            # We want to ask the user whether we should include time codes or not.  Create the prompt
            prompt = unicode(_("Do you want to include Transana Time Codes in the printout?"), "utf8")
            # Create a dialog box for the question
            dlg = Dialogs.QuestionDialog(self, prompt)
            # Display the dialog box and get the user response
            result = dlg.LocalShowModal()
            # Destroy the dialog box
            dlg.Destroy()

            # Define the FONT for the Header and Footer text
            headerFooterFont = wx.Font(10, wx.FONTFAMILY_ROMAN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Times New Roman")
            # Get the current date and time
            t = time.localtime()
            # Format the Footer Date as M/D/Y
            footerDate = "%d/%d/%d" % (t.tm_mon, t.tm_mday, t.tm_year)
            # Format the Footer Page Number
            footerPage = _("Page") + " @PAGENUM@ " + _("of") + " @PAGESCNT@"
            # Create a RichTextPrinting Object
            printout = richtext.RichTextPrinting(_("Transana Report"), self)
            # Let the printout know about the default printer settings
            printout.SetPrintData(TransanaGlobal.printData)
            # Specify the Header / Footer Font
            printout.SetHeaderFooterFont(headerFooterFont)
            # Add the Report Title to the top right
            printout.SetHeaderText(tempTranscript.id, location=richtext.RICHTEXT_PAGE_RIGHT)
            # Add the date to the bottom left
            printout.SetFooterText(footerDate, location=richtext.RICHTEXT_PAGE_LEFT)
            # Add the page number to the bottom right
            printout.SetFooterText(footerPage, location=richtext.RICHTEXT_PAGE_RIGHT)
            # Do NOT show Header and Footer on the First Page
            printout.SetShowOnFirstPage(False)
            # print the RTC Buffer Contents
            # Create a Rich Text Buffer object
            buf = richtext.RichTextBuffer()
            # If the user requested that we strip time codes ...
            if result == wx.ID_NO:
                # ... create a hidden RichTextEditCtrl_RTC
                hiddenCtrl = RichTextEditCtrl_RTC.RichTextEditCtrl(self)
                # Use the hidden control to strip the time codes from the Temp Transcript Object's text
                tempTranscript.text = hiddenCtrl.StripTimeCodes(tempTranscript.text)
                # Now get rid of the hidden control.  We're done with it!
                hiddenCtrl.Destroy()
            # Now put the contents of the XML transcript text into the buffer!
            try:
                # Create a IO stream object
                stream = cStringIO.StringIO(tempTranscript.text)
                # Create an XML Handler
                handler = richtext.RichTextXMLHandler()
                # Load the XML text via the XML Handler.
                # Note that for XML, the BUFFER is passed.
                handler.LoadStream(buf, stream)
            # exception handling
            except:
                print "XML Handler Load failed"
                print
                print sys.exc_info()[0], sys.exc_info()[1]
                print traceback.print_exc()
                print
                pass
            # Define the RichTextPrintout Object's Print Buffer
            printout.PrintBuffer(buf)
            # Destroy the RichTextPrinting Object
            printout.Destroy()

        # If we're using the STC ...
        else:
            # Set the Cursor to the Hourglass while the report is assembled
            self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
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

                # Print Preview on the Mac is broken.  Just print the transcript.
                if 'wxMac' in wx.PlatformInfo:
                    printPreview.Print(True)
                else:
                    
                    # Create the Frame for the Print Preview
                    theWidth = max(wx.Display(0).GetClientArea()[2] - 180, 760)  # wx.ClientDisplayRect()
                    theHeight = max(wx.Display(0).GetClientArea()[3] - 200, 560)  # wx.ClientDisplayRect()
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
        """ Handler for Transcript > Undo """
        # Determine the object with the current focus
        tmpObj = wx.Window.FindFocus()
        # If we have an object OTHER THAN a RichTextEditCtrl ...
        # (this is required for wxMac, as otherwise the RTC handles ALL Cut/Copy/Paste requests!)
        if isinstance(tmpObj, RichTextEditCtrl_RTC.RichTextEditCtrl):
            self.ControlObject.TranscriptUndo(event)
        else:
            event.Skip()

    def OnTranscriptCut(self, event):
        """ Handler for Transcript > Cut """
        # Determine the object with the current focus
        tmpObj = wx.Window.FindFocus()
        # If we have an object OTHER THAN a RichTextEditCtrl ...
        # (this is required for wxMac, as otherwise the RTC handles ALL Cut/Copy/Paste requests!)
        if isinstance(tmpObj, RichTextEditCtrl_RTC.RichTextEditCtrl):
            self.ControlObject.TranscriptCut(event)
        else:
            event.Skip()

    def OnTranscriptCopy(self, event):
        """ Handler for Transcript > Copy """
        # Determine the object with the current focus
        tmpObj = wx.Window.FindFocus()
        # If we have an object OTHER THAN a RichTextEditCtrl ...
        # (this is required for wxMac, as otherwise the RTC handles ALL Cut/Copy/Paste requests!)
        if isinstance(tmpObj, RichTextEditCtrl_RTC.RichTextEditCtrl):
            self.ControlObject.TranscriptCopy(event)
        else:
            event.Skip()

    def OnTranscriptPaste(self, event):
        """ Handler for Transcript > Paste """
        # Determine the object with the current focus
        tmpObj = wx.Window.FindFocus()
        # If we have an object OTHER THAN a RichTextEditCtrl ...
        # (this is required for wxMac, as otherwise the RTC handles ALL Cut/Copy/Paste requests!)
        if isinstance(tmpObj, RichTextEditCtrl_RTC.RichTextEditCtrl):
            self.ControlObject.TranscriptPaste(event)
        else:
            event.Skip()

    def OnFormatFont(self, event):
        """ Handler for Transcript > Format Font """
        self.ControlObject.TranscriptCallFormatDialog()

    def OnFormatParagraph(self, event):
        """ Handler for Transcript > Format Paragraph """
        self.ControlObject.TranscriptCallFormatDialog(tabToOpen=1)

    def OnFormatTabs(self, event):
        """ Handler for Transcript > Format Tabs """
        self.ControlObject.TranscriptCallFormatDialog(tabToOpen=2)

    def OnInsertImage(self, event):
        """ Handler for Transcript > Insert Image """
        self.ControlObject.TranscriptInsertImage()

    def OnAutoTimeCode(self, event):
        """ Handler for Transcript > Fixed-Increment Time Codes """
        # Ask the Control Object to process AutoTimeCoding and let us know if it worked.
        if self.ControlObject.AutoTimeCode():
            # If it worked, disable the menu item!
            self.menuBar.transcriptmenu.Enable(MenuSetup.MENU_TRANSCRIPT_AUTOTIMECODE, False)

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
        # If the Notes Browser is NOT already open ...
        if self.ControlObject.NotesBrowserWindow == None:
            # Instantiate a Notes Browser window
            notesBrowser = NotesBrowser.NotesBrowser(self, -1, _("Notes Browser"))
            # Register the Control Object with the Notes Browser
            notesBrowser.Register(self.ControlObject)
            # Display the Notes Browser
            notesBrowser.Show()
        # If the Notes Browswer IS already open ...
        else:
            # ... bring it to the front.
            self.ControlObject.NotesBrowserWindow.Raise()
            # If the window has been minimized ...
            if self.ControlObject.NotesBrowserWindow.IsIconized():
                # ... then restore it to its proper size!
                self.ControlObject.NotesBrowserWindow.Iconize(False)

    def OnMediaConversion(self, event):
        """ Handler for Tools > Media Conversion """
        # Create a Media Convert dialog
        mediaConv = MediaConvert.MediaConvert(self)
        # Show the dialog
        mediaConv.ShowModal()
        # Call Close here to clean up Temp Files in some circumstances
        mediaConv.Close()
        # Destroy the dialog
        mediaConv.Destroy()
        
    def OnImportDatabase(self, event):
        """ Import Database """

         # If the current database is not empty, we need to tell the user.
        if not DBInterface.IsDatabaseEmpty():
            prompt = _('Your current database is not empty.') + '\n\n' + \
                     _('You can only import data into an existing databases if\nthere are no overlapping Series or Collection records.\nIf there is any overlap in these records, the import\nwill fail.') + '\n\n' + \
                     _('If you have overlapping Keywords, the existing Keyword\nis retained and the importing Keyword (including its\ndefinition, which could differ) is discarded.') + '\n\n' + \
                     _('If you have overlapping Core Data records, the existing\nCore Data record is retained and the importing Core\nData record is discarded.')
            dlg = Dialogs.InfoDialog(self, prompt)
            dlg.ShowModal()
            dlg.Destroy()

        # Create an Import Database dialog
        temp = XMLImport.XMLImport(self, -1, _('Transana XML Import'))
        # Get the User Input
        result = temp.get_input()
        # If the user gave a valid response ...
        if (result != None) and (result[_("Transana-XML Filename")] != ''):
            # See if there's a Notes Browser open
            if self.ControlObject.NotesBrowserWindow != None:
                # If so, close it, which saves anything being edited.
                self.ControlObject.NotesBrowserWindow.Close()
            # ... Import the requested data!
            temp.Import()
            # If MU, we need to signal other copies that we've imported a database!
            # First, test to see if we're in the Multi-user version.
            if not TransanaConstants.singleUserVersion:
                # Now make sure a Chat Window has been defined
                if TransanaGlobal.chatWindow != None:
                    # Now send the "Import" message
                    TransanaGlobal.chatWindow.SendMessage("I ")
        # Close the Import Database dialog
        temp.Close()

    def OnExportDatabase(self, event):
        """ Export Database """
        # Create an Export Database dialog
        temp = XMLExport.XMLExport(self, -1, _('Transana XML Export'))
        # Set up the confirmation loop signal variable
        repeat = True
        # While we are in the confirmation loop ...
        while repeat:
            # ... assume we will want to exit the confirmation loop by default
            repeat = False
            # Get the XML Export User Input
            result = temp.get_input()
            # if the user clicked OK ...
            if result != None:
                # ... make sure they entered a file name.
                if result[_('Transana-XML Filename')] == '':
                    # If not, create a prompt to inform the user ...
                    prompt = unicode(_('A file name is required'), 'utf8') + '.'
                    # ... and signal that we need to repeat the file prompt
                    repeat = True
                # If they did ...
                else:
                    # ... error check the file name.  If it does not have a PATH ...
                    if os.path.split(result[_('Transana-XML Filename')])[0] == u'':
                        # ... add the Video Path to the file name
                        fileName = os.path.join(TransanaGlobal.configData.videoPath, result[_('Transana-XML Filename')])
                    # If there is a path, just continue.
                    else:
                        fileName = result[_('Transana-XML Filename')]
                    # If the file does not have a .TRA extension ...
                    if fileName[-4:].lower() != '.tra':
                        # ... add one
                        fileName = fileName + '.tra'
                    # Set the FORM's field value to the modified file name
                    temp.XMLFile.SetValue(fileName)

                    # Check the file name for illegal characters.  First, define illegal characters
                    # (not including PATH characters)
                    illegalChars = '"*?<>|'
                    # For each illegal character ...
                    for char in illegalChars:
                        # ... see if that character appears in the file name the user entered
                        if char in result[_('Transana-XML Filename')]:
                            # If so, create a prompt to inform the user ...
                            prompt = unicode(_('There is an illegal character in the file name.'), 'utf8')
                            # ... and signal that we need to repeat the file prompt ...
                            repeat = True
                            # ... and stop looking.
                            break

                # Was there a file name problem or an illegal character?
                if repeat:
                    # If so, display the prompt to inform the user.
                    dlg2 = Dialogs.ErrorDialog(self, prompt)
                    dlg2.ShowModal()
                    dlg2.Destroy()
                    # Signal that we have not gotten a result.
                    result = None
                # If we get to here, check for a duplicate file name
                elif (os.path.exists(fileName)):
                    # If so, create a prompt to inform the user and ask to overwrite the file.
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('A file named "%s" already exists.  Do you want to replace it?'), 'utf8')
                    else:
                        prompt = _('A file named "%s" already exists.  Do you want to replace it?')
                    # Create the dialog for the prompt for the user
                    dlg2 = Dialogs.QuestionDialog(None, prompt % fileName)
                    # Center the confirmation dialog on screen
                    dlg2.CentreOnScreen()
                    # Show the confirmation dialog and get the user response.  If the user DOES NOT say to overwrite the file, ...
                    if dlg2.LocalShowModal() != wx.ID_YES:
                        # ... nullify the results of the Clip Data Export dialog so the file won't be overwritten ...
                        result = None
                        # ... and signal that the user should be re-prompted.
                        repeat = True
                    # Destroy the confirmation dialog
                    dlg2.Destroy()

        # If the user requests it ...
        if (result != None) and (result[_("Transana-XML Filename")] != ''):
            # ... export the data
            temp.Export()
        # Close the Export Database dialog
        temp.Close()

    def OnColorConfig(self, event):
        """ Graphics Color Configuration """
        # Create a Color Configuration Dialog
        temp = ColorConfig.ColorConfig(self)
        # Show the Dialog
        temp.ShowModal()
        # Destroy the Dialog
        temp.Destroy()

    def OnBatchWaveformGenerator(self, event):
        """ Batch Waveform Generator """
        # Create a Batch Waveform Dialog
        temp = BatchFileProcessor.BatchFileProcessor(self, mode="waveform")
        # Get User input
        temp.get_input()
        # Close the Dialog
        temp.Close()
        # Destroy the Dialog
        temp.Destroy()

    def OnChat(self, event):
        """ Chat Window """
        # If a Chat Window has been defined ...
        if TransanaGlobal.chatWindow != None:
            # ... show it!
            TransanaGlobal.chatWindow.Show()
            # Call Raise to bring the Chat Window to the top.
            TransanaGlobal.chatWindow.Raise()
            # If the Chat Window is minimized ...
            if TransanaGlobal.chatWindow.IsIconized():
                # ... then restore it to its proper size!
                TransanaGlobal.chatWindow.Iconize(False)
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
            # If the Notes Browser Window is open ...
            if self.ControlObject.NotesBrowserWindow != None:
                # ... close it, thus releasing any records that might be locked there.
                self.ControlObject.NotesBrowserWindow.Close()
        # Create a Record Lock Utility window
        recordLockWindow = RecordLock.RecordLock(self, -1, _("Transana Record Lock Utility"))
        recordLockWindow.ShowModal()
        recordLockWindow.Destroy()

    def OnOptionsSettings(self, event):
        """ Handler for Options > Settings """
        # Remember the old Tab Size and Word Wrap values
        oldTabSize = TransanaGlobal.configData.tabSize
        oldWordWrap = TransanaGlobal.configData.wordWrap
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
        # Update the Tab Size and Word Wrap.  First, see if they've changed.
        if (TransanaGlobal.configData.tabSize != oldTabSize) or (TransanaGlobal.configData.wordWrap != oldWordWrap):
            # For each Transcript Window ...
            for trWin in self.ControlObject.TranscriptWindow:
                # ... set the new tab size
                trWin.dlg.editor.SetTabWidth(int(TransanaGlobal.configData.tabSize))
                # ... and set the new Word Wrap
                trWin.dlg.editor.SetWrapMode(TransanaGlobal.configData.wordWrap)
            
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
        # We're starting to face the situation where not all the translations may be up-to-date.  Let's build some checks.
        # Initialize an empty variable
        outofdateLanguage = ''

        # English
        if event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_EN:
            TransanaGlobal.configData.language = 'en'
            self.presLan_en.install()
            
        # Arabic
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_AR:
            outofdateLanguage = 'Arabic'
            TransanaGlobal.configData.language = 'ar'
            self.presLan_ar.install()
            
        # Danish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_DA:
            outofdateLanguage = 'Danish'
            TransanaGlobal.configData.language = 'da'
            self.presLan_da.install()
            
        # German
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_DE:
            outofdateLanguage = 'German'
            TransanaGlobal.configData.language = 'de'
            self.presLan_de.install()

        # Greek
#        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_EL:
#            outofdateLanguage = 'Greek'
#            TransanaGlobal.configData.language = 'el'
#            self.presLan_el.install()

        # Spanish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_ES:
            outofdateLanguage = 'Spanish'
            TransanaGlobal.configData.language = 'es'
            self.presLan_es.install()

        # Finnish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_FI:
            outofdateLanguage = 'Finnish'
            TransanaGlobal.configData.language = 'fi'
            self.presLan_fi.install()

        # French
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_FR:
            outofdateLanguage = 'French'
            TransanaGlobal.configData.language = 'fr'
            self.presLan_fr.install()

        # Hebrew
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_HE:
            outofdateLanguage = 'Hebrew'
            TransanaGlobal.configData.language = 'he'
            self.presLan_he.install()

        # Italian
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_IT:
            outofdateLanguage = 'Italian'
            TransanaGlobal.configData.language = 'it'
            self.presLan_it.install()

        # Dutch
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_NL:
            outofdateLanguage = 'Dutch'
            TransanaGlobal.configData.language = 'nl'
            self.presLan_nl.install()

        # Norwegian Bokmal
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_NB:
            outofdateLanguage = 'Norwegian Bokmal'
            TransanaGlobal.configData.language = 'nb'
            self.presLan_nb.install()

        # Norwegian Ny-norsk
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_NN:
            outofdateLanguage = 'Norwegian Nynorsk'
            TransanaGlobal.configData.language = 'nn'
            self.presLan_nn.install()

        # Polish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_PL:
            outofdateLanguage = 'Polish'
            TransanaGlobal.configData.language = 'pl'
            self.presLan_pl.install()

        # Portuguese
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_PT:
            outofdateLanguage = 'Portuguese'
            TransanaGlobal.configData.language = 'pt'
            self.presLan_pt.install()

        # Russian
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_RU:
            outofdateLanguage = 'Russian'
            TransanaGlobal.configData.language = 'ru'
            self.presLan_ru.install()

        # Swedish
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_SV:
            outofdateLanguage = 'Swedish'
            TransanaGlobal.configData.language = 'sv'
            self.presLan_sv.install()

        # Chinese (English prompts)
        elif  event.GetId() == MenuSetup.MENU_OPTIONS_LANGUAGE_ZH:
            outofdateLanguage = 'Chinese - Simplified'
            TransanaGlobal.configData.language = 'zh'
            self.presLan_zh.install()

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

        # Check to see if we have a translation, and if it is up-to-date.
        
        # NOTE:  "Fixed-Increment Time Code" works for version 2.42.  "&Media Conversion" works for version 2.50.
        # If you update this, also update the phrase above in the __init__ method.)
        
        if (outofdateLanguage != '') and ("&Media Conversion" == _("&Media Conversion")):
            # If not, display an information message.
            prompt = "Transana's %s translation is no longer up-to-date.\nMissing prompts will be displayed in English.\n\nIf you are willing to help with this translation,\nplease contact David Woods at dwoods@wcer.wisc.edu." % outofdateLanguage
            dlg = wx.MessageDialog(None, prompt, "Translation update", style=wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()

        infodlg = Dialogs.InfoDialog(None, _("Please note that some prompts cannot be set to the new language until you restart Transana, and\nthat the language of some prompts is determined by your operating system instead of Transana."))
        infodlg.ShowModal()
        infodlg.Destroy()
        
        self.ControlObject.ChangeLanguages()

    def OnOptionsQuickClipMode(self, event):
        """ Handler for Options > Quick Clip Mode """
        # All we need to do is toggle the global value when teh menu option is changed
        TransanaGlobal.configData.quickClipMode = event.IsChecked()

    def OnOptionsQuickClipWarning(self, event):
        """ Handler for Options > Show Quick Clip Warning """
        TransanaGlobal.configData.quickClipWarning = event.IsChecked()

    def OnOptionsAutoArrange(self, event):
        """ Handler for Options > Auto-Arrange """
        # We need to toggle the global value when the menu option is changed
        TransanaGlobal.configData.autoArrange = event.IsChecked()

        # If we turn Auto-Arrange ON ...
        if event.IsChecked():
            # ... let's restore the windows to the positions saved earlier.
            # Menu Window
            if self.menuWindowLayout != None:
                self.SetPosition(self.menuWindowLayout[0])
                self.SetSize(self.menuWindowLayout[1])
            # Visualization Window
            if self.visualizationWindowLayout != None:
                self.ControlObject.VisualizationWindow.SetPosition(self.visualizationWindowLayout[0])
                self.ControlObject.VisualizationWindow.SetSize(self.visualizationWindowLayout[1])
            # Video Window
            if self.videoWindowLayout != None:
                self.ControlObject.VideoWindow.SetPosition(self.videoWindowLayout[0])
                self.ControlObject.VideoWindow.SetSize(self.videoWindowLayout[1])
                # Try to get the graphic to update
                self.ControlObject.VideoWindow.Refresh()
            # Transcript Window
            if self.transcriptWindowLayout != None:
                self.ControlObject.TranscriptWindow[0].dlg.SetPosition(self.transcriptWindowLayout[0])
                self.ControlObject.TranscriptWindow[0].dlg.SetSize(self.transcriptWindowLayout[1])
                # Auto-arrange additional transcripts
                self.ControlObject.AutoArrangeTranscriptWindows()
            # Data window
            if self.dataWindowLayout != None:
                self.ControlObject.DataWindow.SetPosition(self.dataWindowLayout[0])
                self.ControlObject.DataWindow.SetSize(self.dataWindowLayout[1])
        # If we turn Auto-Arrange OFF ...
        else:
            # ... remember the current window layouts
            self.menuWindowLayout = (self.GetPosition(), self.GetSize())
            self.visualizationWindowLayout = (self.ControlObject.VisualizationWindow.GetPosition(), self.ControlObject.VisualizationWindow.GetSize())
            self.videoWindowLayout = (self.ControlObject.VideoWindow.GetPosition(), self.ControlObject.VideoWindow.GetSize())
            self.transcriptWindowLayout = (self.ControlObject.TranscriptWindow[0].dlg.GetPosition(), self.ControlObject.TranscriptWindow[0].dlg.GetSize())
            self.dataWindowLayout = (self.ControlObject.DataWindow.GetPosition(), self.ControlObject.DataWindow.GetSize())

    def OnOptionsWordTrack(self, event):
        """ Handler for Options > Auto Word-tracking """
        # Set global value appropriately when state changes
        TransanaGlobal.configData.wordTracking = event.IsChecked()

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
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_FONT, _("Format &Font"))
        if TransanaConstants.USESRTC:
            self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_PARAGRAPH, _("Format Paragrap&h"))
            self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_TABS, _("Format Ta&bs"))
            self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_INSERT_IMAGE, _("&Insert Image"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_PRINT, _("&Print"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_PRINTERSETUP, _("Printer &Setup"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_AUTOTIMECODE, _("F&ixed-Increment Time Codes"))
        self.menuBar.transcriptmenu.SetLabel(MenuSetup.MENU_TRANSCRIPT_ADJUSTINDEXES, _("&Adjust Indexes"))

        self.menuBar.SetLabelTop(2, _("Too&ls"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_NOTESBROWSER, _("&Notes Browser"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_FILEMANAGEMENT, _("&File Management"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_MEDIACONVERSION, _("&Media Conversion"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_IMPORT_DATABASE, _("&Import Database"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_EXPORT_DATABASE, _("&Export Database"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_COLORCONFIG, _("&Graphics Color Configuration"))
        self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_BATCHWAVEFORM, _("&Batch Waveform Generator"))
        if not TransanaConstants.singleUserVersion:
            self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_CHAT, _("&Chat Window"))
            self.menuBar.toolsmenu.SetLabel(MenuSetup.MENU_TOOLS_RECORDLOCK, _("&Record Lock Utility"))

        self.menuBar.SetLabelTop(3, _("&Options"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_SETTINGS, _("Program &Settings"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE, _("&Language"))
        self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_EN, _("&English"))
        # The Langage menus may not exist, and we should only update them if they do!
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_AR) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_AR, _("&Arabic"))
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
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_HE) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_HE, _("&Hebrew"))
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
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_PT) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_PT, _("P&ortuguese"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_RU) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_RU, _("&Russian"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_SV) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_SV, _("S&wedish"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_ZH) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_ZH, _("&Chinese - Simplified"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_EASTEUROPE) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_EASTEUROPE, _("English prompts, Eastern European data (ISO-8859-2 encoding)"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_EL) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_EL, _("English prompts, Greek data"))
        if self.menuBar.optionslanguagemenu.FindItemById(MenuSetup.MENU_OPTIONS_LANGUAGE_JA) != None:
            self.menuBar.optionslanguagemenu.SetLabel(MenuSetup.MENU_OPTIONS_LANGUAGE_JA, _("English prompts, Japanese data"))
        if TransanaConstants.macDragDrop or (not 'wxMac' in wx.PlatformInfo):
            self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_QUICK_CLIPS, _("&Quick Clip Mode"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_QUICKCLIPWARNING, _("Show Quick Clip Warning"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_WORDTRACK, _("Auto &Word-tracking"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_AUTOARRANGE, _("&Auto-Arrange"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_VISUALIZATION, _("Vi&sualization Style"))
        self.menuBar.optionsvisualizationmenu.SetLabel(MenuSetup.MENU_OPTIONS_VISUALIZATION_WAVEFORM, _("&Waveform"))
        self.menuBar.optionsvisualizationmenu.SetLabel(MenuSetup.MENU_OPTIONS_VISUALIZATION_KEYWORD, _("&Keyword"))
        self.menuBar.optionsvisualizationmenu.SetLabel(MenuSetup.MENU_OPTIONS_VISUALIZATION_HYBRID, _("&Hybrid"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_VIDEOSIZE, _("&Video Size"))
        self.menuBar.optionspresentmenu.SetLabel(MenuSetup.MENU_OPTIONS_PRESENT_ALL, _("&All Windows"))
        self.menuBar.optionspresentmenu.SetLabel(MenuSetup.MENU_OPTIONS_PRESENT_VIDEO, _("&Video Only"))
        self.menuBar.optionspresentmenu.SetLabel(MenuSetup.MENU_OPTIONS_PRESENT_TRANS, _("Video and &Transcript Only"))
        self.menuBar.optionspresentmenu.SetLabel(MenuSetup.MENU_OPTIONS_PRESENT_AUDIO, _("A&udio and Transcript Only"))
        self.menuBar.optionsmenu.SetLabel(MenuSetup.MENU_OPTIONS_PRESENT, _("&Presentation Mode"))

        if not 'wxMac' in wx.PlatformInfo:
            self.menuBar.SetLabelTop(4, _("&Help"))
        else:
            wx.App_SetMacHelpMenuTitleName(_("&Help"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_MANUAL, _("&Manual"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_TUTORIAL, _("&Tutorial"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_NOTATION, _("Transcript &Notation"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_ABOUT, _("&About"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_WEBSITE, _("&www.transana.org"))
        self.menuBar.helpmenu.SetLabel(MenuSetup.MENU_HELP_FUND, _("&Fund Transana"))

        wx.App_SetMacHelpMenuTitleName(_("&Help"))

        # print "Menu Language Changed (%s)" % _("&File")

    def GetLanguage(self, parent):
        """ Determines what languages have been installed and prompts the user to select one. """
        # See what languages are installed on the system and make a list
        languages = [ENGLISH_LABEL]
        # Arabic
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'ar', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(ARABIC_LABEL)
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
        # Hebrew
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'he', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(HEBREW_LABEL)
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
        # Portuguese
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'pt', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(PORTUGUESE_LABEL)
        # Russian
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'ru', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(RUSSIAN_LABEL)
        # Swedish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'sv', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(SWEDISH_LABEL)
        # Swedish
        dir = os.path.join(TransanaGlobal.programDir, 'locale', 'zh', 'LC_MESSAGES', 'Transana.mo')
        if os.path.exists(dir):
            languages.append(CHINESE_LABEL)
        # Easern Europe encoding, Greek, Japanese, Korean, and Chinese
        if ('wxMSW' in wx.PlatformInfo) and TransanaConstants.singleUserVersion:
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
