import wx

# This module expects i18n.  Enable it here.
__builtins__._ = wx.GetTranslation
import gettext

import os
import sys

import TransanaConstants
import TransanaGlobal
import Clip
import Collection
import ConfigData
import ControlObjectClass
import DatabaseTreeTab
import DBInterface
import Dialogs
import Episode
import KeywordObject
import MenuWindow
import ProcessSearch
import Library
import Snapshot
import TransanaExceptions
import Transcript

MENU_FILE_EXIT = 101

class FormCheck(wx.Frame):
    """ This window displays a variety of GUI Widgets. """
    def __init__(self,parent,id,title):

        self.presLan_en = gettext.translation('Transana', 'locale', languages=['en'])
        self.presLan_en.install()
        lang = wx.LANGUAGE_ENGLISH
        self.locale = wx.Locale(lang)
        self.locale.AddCatalog("Transana")
        
        # Define the Configuration Data
        TransanaGlobal.configData = ConfigData.ConfigData()
        # Create the global transana graphics colors, once the ConfigData object exists.
        TransanaGlobal.transana_graphicsColorList = TransanaGlobal.getColorDefs(TransanaGlobal.configData.colorConfigFilename)
        # Set essential global color manipulation data structures once the ConfigData object exists.
        (TransanaGlobal.transana_colorNameList, TransanaGlobal.transana_colorLookup, TransanaGlobal.keywordMapColourSet) = TransanaGlobal.SetColorVariables()
        TransanaGlobal.configData.videoPath = 'C:\\Users\\DavidWoods\\Videos'
        TransanaGlobal.configData.ssl = False

        TransanaGlobal.menuWindow = MenuWindow.MenuWindow(None, -1, title)

        self.ControlObject = ControlObjectClass.ControlObject()
        self.ControlObject.Register(Menu = TransanaGlobal.menuWindow)
        TransanaGlobal.menuWindow.Register(ControlObject = self.ControlObject)

        if TransanaConstants.DBInstalled in ['MySQLdb-server', 'PyMySQL']:

            dlg = wx.PasswordEntryDialog(None, 'Please enter your database password.', 'Unit Test 1:  Database Connection')
            result = dlg.ShowModal()
            if result == wx.ID_OK:
                password = dlg.GetValue()
            else:
                password = ''
            dlg.Destroy()

        else:
            if TransanaConstants.DBInstalled in ['MySQLdb-embedded']:
                DBInterface.InitializeSingleUserDatabase()
            password = ''
            
        loginInfo = DBLogin('DavidW', password, 'DKW-Linux', 'Transana_UnitTest', '3306')
        dbReference = DBInterface.get_db(loginInfo)
        DBInterface.establish_db_exists()

        wx.Frame.__init__(self,parent,-1, title, size = (800, 900), style=wx.DEFAULT_FRAME_STYLE|wx.NO_FULL_REPAINT_ON_RESIZE)
        self.testsRun = 0
        self.testsSuccessful = 0
        self.testsFailed = 0
       
        self.SetBackgroundColour(wx.RED)

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        self.txtCtrl = wx.TextCtrl(self, -1, "Unit Test 3:  Search\n\n", style=wx.TE_LEFT | wx.TE_MULTILINE)
        self.txtCtrl.AppendText("Transana Version:  %s     singleUserVersion:  %s\n\n" % (TransanaConstants.versionNumber, TransanaConstants.singleUserVersion))
        mainSizer.Add(self.txtCtrl, 1, wx.EXPAND | wx.ALL, 5)

        self.tree = DatabaseTreeTab._DBTreeCtrl(self, -1, wx.DefaultPosition, self.GetSizeTuple(), wx.TR_HAS_BUTTONS | wx.TR_EDIT_LABELS | wx.TR_MULTIPLE)
        mainSizer.Add(self.tree, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()
        self.CenterOnScreen()
        # Status Bar
        self.CreateStatusBar()
        self.SetStatusText("")

        self.Show(True)
        self.RunTests()

        if TransanaConstants.DBInstalled in ['MySQLdb-embedded']:
            DBInterface.EndSingleUserDatabase()

        TransanaGlobal.menuWindow.Destroy()

    def RunTests(self):
        # Tests defined:
        testsNotToSkip = []  # DON'T SKIP TESTS in this file!
        startAtTest = 1  # Should start at 1, not 0!
        endAtTest = 500   # Should be one more than the last test to be run!
        testsToRun = testsNotToSkip + range(startAtTest, endAtTest)

        if (not TransanaConstants.singleUserVersion):
            # Database Connection
            testName = 'Database Connection, set up Database'
            self.SetStatusText(testName)

        if 11 in testsToRun:
            # Keyword Saving
            testName = 'Creating Keywords : A'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                keywordA = KeywordObject.Keyword('SearchDemo', 'A')
                # If the keyword already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the keyword doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                keywordA = KeywordObject.Keyword()
                keywordA.keywordGroup = 'SearchDemo'
                keywordA.keyword = 'A'
                keywordA.definition = 'Created by unit_test_search'
                try:
                    keywordA.db_save()
                    # If we create the keyword, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the keyword, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the keyword, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
                
        if 12 in testsToRun:
            # Keyword Saving
            testName = 'Creating Keywords : B'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                keywordB = KeywordObject.Keyword('SearchDemo', 'B')
                # If the keyword already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the keyword doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                keywordB = KeywordObject.Keyword()
                keywordB.keywordGroup = 'SearchDemo'
                keywordB.keyword = 'B'
                keywordB.definition = 'Created by unit_test_search'
                try:
                    keywordB.db_save()
                    # If we create the keyword, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the keyword, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the keyword, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
                
        if 13 in testsToRun:
            # Keyword Saving
            testName = 'Creating Keywords : C'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                keywordC = KeywordObject.Keyword('SearchDemo', 'C')
                # If the keyword already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the keyword doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                keywordC = KeywordObject.Keyword()
                keywordC.keywordGroup = 'SearchDemo'
                keywordC.keyword = 'C'
                keywordC.definition = 'Created by unit_test_search'
                try:
                    keywordC.db_save()
                    # If we create the keyword, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the keyword, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the keyword, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 21 in testsToRun:
            # Library Saving
            testName = 'Creating Library : SearchDemo'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                series1 = Library.Library('SearchDemo')
                # If the Library already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Library doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                series1 = Library.Library()
                series1.id = 'SearchDemo'
                series1.owner = 'unit_test_search'
                series1.comment = 'Created by unit_test_search'
                series1.keyword_group = 'SearchDemo'
                try:
                    series1.db_save()
                    # If we create the Library, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Library, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Library, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 31 in testsToRun:
            # Episode Saving
            testName = 'Creating Episode : A'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                episodeA = Episode.Episode(series='SearchDemo', episode='A')
                # If the Episode already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Episode doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                episodeA = Episode.Episode()
                episodeA.id = 'A'
                episodeA.comment = 'Created by unit_test_search'
                episodeA.media_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Demo', 'Demo.mpg')
                episodeA.series_id = series1.id
                episodeA.series_num = series1.number
                episodeA.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                try:
                    episodeA.db_save()
                    # If we create the Episode, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Episode, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Episode, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 32 in testsToRun:
            # Episode Saving
            testName = 'Creating Episode : B'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                episodeB = Episode.Episode(series='SearchDemo', episode='B')
                # If the Episode already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Episode doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                episodeB = Episode.Episode()
                episodeB.id = 'B'
                episodeB.comment = 'Created by unit_test_search'
                episodeB.media_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Demo', 'Demo.mpg')
                episodeB.series_id = series1.id
                episodeB.series_num = series1.number
                episodeB.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                try:
                    episodeB.db_save()
                    # If we create the Episode, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Episode, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Episode, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 33 in testsToRun:
            # Episode Saving
            testName = 'Creating Episode : C'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                episodeC = Episode.Episode(series='SearchDemo', episode='C')
                # If the Episode already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Episode doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                episodeC = Episode.Episode()
                episodeC.id = 'C'
                episodeC.comment = 'Created by unit_test_search'
                episodeC.media_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Demo', 'Demo.mpg')
                episodeC.series_id = series1.id
                episodeC.series_num = series1.number
                episodeC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                try:
                    episodeC.db_save()
                    # If we create the Episode, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Episode, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Episode, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 35 in testsToRun:
            # Episode Saving
            testName = 'Creating Episode : AB'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                episodeAB = Episode.Episode(series='SearchDemo', episode='AB')
                # If the Episode already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Episode doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                episodeAB = Episode.Episode()
                episodeAB.id = 'AB'
                episodeAB.comment = 'Created by unit_test_search'
                episodeAB.media_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Demo', 'Demo.mpg')
                episodeAB.series_id = series1.id
                episodeAB.series_num = series1.number
                episodeAB.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                episodeAB.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                try:
                    episodeAB.db_save()
                    # If we create the Episode, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Episode, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Episode, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 36 in testsToRun:
            # Episode Saving
            testName = 'Creating Episode : AC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                episodeAC = Episode.Episode(series='SearchDemo', episode='AC')
                # If the Episode already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Episode doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                episodeAC = Episode.Episode()
                episodeAC.id = 'AC'
                episodeAC.comment = 'Created by unit_test_search'
                episodeAC.media_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Demo', 'Demo.mpg')
                episodeAC.series_id = series1.id
                episodeAC.series_num = series1.number
                episodeAC.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                episodeAC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                try:
                    episodeAC.db_save()
                    # If we create the Episode, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Episode, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Episode, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 37 in testsToRun:
            # Episode Saving
            testName = 'Creating Episode : BC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                episodeBC = Episode.Episode(series='SearchDemo', episode='BC')
                # If the Episode already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Episode doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                episodeBC = Episode.Episode()
                episodeBC.id = 'BC'
                episodeBC.comment = 'Created by unit_test_search'
                episodeBC.media_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Demo', 'Demo.mpg')
                episodeBC.series_id = series1.id
                episodeBC.series_num = series1.number
                episodeBC.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                episodeBC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                try:
                    episodeBC.db_save()
                    # If we create the Episode, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Episode, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Episode, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 38 in testsToRun:
            # Episode Saving
            testName = 'Creating Episode : ABC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                episodeABC = Episode.Episode(series='SearchDemo', episode='ABC')
                # If the Episode already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Episode doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                episodeABC = Episode.Episode()
                episodeABC.id = 'ABC'
                episodeABC.comment = 'Created by unit_test_search'
                episodeABC.media_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Demo', 'Demo.mpg')
                episodeABC.series_id = series1.id
                episodeABC.series_num = series1.number
                episodeABC.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                episodeABC.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                episodeABC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                try:
                    episodeABC.db_save()
                    # If we create the Episode, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Episode, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Episode, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        tempText = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
    <paragraph>
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;0&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">0</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;5000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">5</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;10000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">10</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;15000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">15</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;20000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">20</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;25000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">25</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;30000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">30</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;35000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">35</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;40000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">40</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;45000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">45</text>
    </paragraph>
  </paragraphlayout>
</richtext>
"""

        if 41 in testsToRun:
            # Transcript Saving
            testName = 'Creating Transcript : A'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                transcriptA = Transcript.Transcript('A', ep=episodeA.number)
                # If the Transcript already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Transcript doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                transcriptA = Transcript.Transcript()
                transcriptA.id = 'A'
                transcriptA.comment = 'Created by unit_test_search'
                transcriptA.episode_num = episodeA.number
                transcriptA.transcriber = 'unit_test_search'
                transcriptA.text = tempText
                try:
                    transcriptA.db_save()
                    # If we create the Transcript, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Transcript, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Transcript, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 42 in testsToRun:
            # Transcript Saving
            testName = 'Creating Transcript : B'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                transcriptB = Transcript.Transcript('B', ep=episodeB.number)
                # If the Transcript already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Transcript doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                transcriptB = Transcript.Transcript()
                transcriptB.id = 'B'
                transcriptB.comment = 'Created by unit_test_search'
                transcriptB.episode_num = episodeB.number
                transcriptB.transcriber = 'unit_test_search'
                transcriptB.text = tempText
                try:
                    transcriptB.db_save()
                    # If we create the Transcript, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Transcript, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Transcript, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 43 in testsToRun:
            # Transcript Saving
            testName = 'Creating Transcript : C'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                transcriptC = Transcript.Transcript('C', ep=episodeC.number)
                # If the Transcript already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Transcript doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                transcriptC = Transcript.Transcript()
                transcriptC.id = 'C'
                transcriptC.comment = 'Created by unit_test_search'
                transcriptC.episode_num = episodeC.number
                transcriptC.transcriber = 'unit_test_search'
                transcriptC.text = tempText
                try:
                    transcriptC.db_save()
                    # If we create the Transcript, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Transcript, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Transcript, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 44 in testsToRun:
            # Transcript Saving
            testName = 'Creating Transcript : AB'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                transcriptAB = Transcript.Transcript('AB', ep=episodeAB.number)
                # If the Transcript already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Transcript doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                transcriptAB = Transcript.Transcript()
                transcriptAB.id = 'AB'
                transcriptAB.comment = 'Created by unit_test_search'
                transcriptAB.episode_num = episodeAB.number
                transcriptAB.transcriber = 'unit_test_search'
                transcriptAB.text = tempText
                try:
                    transcriptAB.db_save()
                    # If we create the Transcript, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Transcript, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Transcript, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 45 in testsToRun:
            # Transcript Saving
            testName = 'Creating Transcript : AC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                transcriptAC = Transcript.Transcript('AC', ep=episodeAC.number)
                # If the Transcript already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Transcript doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                transcriptAC = Transcript.Transcript()
                transcriptAC.id = 'AC'
                transcriptAC.comment = 'Created by unit_test_search'
                transcriptAC.episode_num = episodeAC.number
                transcriptAC.transcriber = 'unit_test_search'
                transcriptAC.text = tempText
                try:
                    transcriptAC.db_save()
                    # If we create the Transcript, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Transcript, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Transcript, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 46 in testsToRun:
            # Transcript Saving
            testName = 'Creating Transcript : BC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                transcriptBC = Transcript.Transcript('BC', ep=episodeBC.number)
                # If the Transcript already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Transcript doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                transcriptBC = Transcript.Transcript()
                transcriptBC.id = 'BC'
                transcriptBC.comment = 'Created by unit_test_search'
                transcriptBC.episode_num = episodeBC.number
                transcriptBC.transcriber = 'unit_test_search'
                transcriptBC.text = tempText
                try:
                    transcriptBC.db_save()
                    # If we create the Transcript, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Transcript, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Transcript, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 47 in testsToRun:
            # Transcript Saving
            testName = 'Creating Transcript : ABC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                transcriptABC = Transcript.Transcript('ABC', ep=episodeABC.number)
                # If the Transcript already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Transcript doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                transcriptABC = Transcript.Transcript()
                transcriptABC.id = 'ABC'
                transcriptABC.comment = 'Created by unit_test_search'
                transcriptABC.episode_num = episodeABC.number
                transcriptABC.transcriber = 'unit_test_search'
                transcriptABC.text = tempText
                try:
                    transcriptABC.db_save()
                    # If we create the Transcript, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Transcript, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Transcript, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 51 in testsToRun:
            # Collection Saving
            testName = 'Creating Collection : SearchDemo'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                collection1 = Collection.Collection('SearchDemo')
                # If the Collection already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Collection doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                collection1 = Collection.Collection()
                collection1.id = 'SearchDemo'
                collection1.parent = 0
                collection1.owner = 'unit_test_search'
                collection1.comment = 'Created by unit_test_search'
                collection1.keyword_group = 'SearchDemo'
                try:
                    collection1.db_save()
                    # If we create the Collection, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Collection, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Collection, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        cliptranscriptA = None
        if 61 in testsToRun:
            # Clip Saving
            testName = 'Creating Clip : A'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                clipA = Clip.Clip('A', collection1.id, collection1.parent)
                # If the Clip already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Clip doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                clipA = Clip.Clip()
                clipA.id = 'A'
                clipA.episode_num = episodeA.number
                clipA.clip_start = 5000
                clipA.clip_stop = 10000
                clipA.collection_num = collection1.number
                clipA.collection_id = collection1.id
                clipA.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                clipA.offset = 0
                clipA.media_filename = episodeA.media_filename
                clipA.audio = 1
                clipA.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                # Prepare the Clip Transcript Object
                cliptranscriptA = Transcript.Transcript()
                cliptranscriptA.episode_num = episodeA.number
                cliptranscriptA.source_transcript = transcriptA.number
                cliptranscriptA.clip_start = clipA.clip_start
                cliptranscriptA.clip_stop = clipA.clip_stop
                cliptranscriptA.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">5</text>
    </paragraph>
  </paragraphlayout>
</richtext>"""
                clipA.transcripts.append(cliptranscriptA)
                try:
                    clipA.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Clip, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Clip, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        cliptranscriptB = None
        if 62 in testsToRun:
            # Clip Saving
            testName = 'Creating Clip : B'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                clipB = Clip.Clip('B', collection1.id, collection1.parent)
                # If the Clip already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Clip doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                clipB = Clip.Clip()
                clipB.id = 'B'
                clipB.episode_num = episodeA.number
                clipB.clip_start = 10000
                clipB.clip_stop = 15000
                clipB.collection_num = collection1.number
                clipB.collection_id = collection1.id
                clipB.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                clipB.offset = 0
                clipB.media_filename = episodeA.media_filename
                clipB.audio = 1
                clipB.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                # Prepare the Clip Transcript Object
                cliptranscriptB = Transcript.Transcript()
                cliptranscriptB.episode_num = episodeA.number
                cliptranscriptB.source_transcript = transcriptA.number
                cliptranscriptB.clip_start = clipB.clip_start
                cliptranscriptB.clip_stop = clipB.clip_stop
                cliptranscriptB.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">10</text>
    </paragraph>
  </paragraphlayout>
</richtext>"""
                clipB.transcripts.append(cliptranscriptB)
                try:
                    clipB.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Clip, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Clip, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        cliptranscriptC = None
        if 63 in testsToRun:
            # Clip Saving
            testName = 'Creating Clip : C'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                clipC = Clip.Clip('C', collection1.id, collection1.parent)
                # If the Clip already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Clip doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                clipC = Clip.Clip()
                clipC.id = 'C'
                clipC.episode_num = episodeA.number
                clipC.clip_start = 15000
                clipC.clip_stop = 20000
                clipC.collection_num = collection1.number
                clipC.collection_id = collection1.id
                clipC.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                clipC.offset = 0
                clipC.media_filename = episodeA.media_filename
                clipC.audio = 1
                clipC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                # Prepare the Clip Transcript Object
                cliptranscriptC = Transcript.Transcript()
                cliptranscriptC.episode_num = episodeA.number
                cliptranscriptC.source_transcript = transcriptA.number
                cliptranscriptC.clip_start = clipC.clip_start
                cliptranscriptC.clip_stop = clipC.clip_stop
                cliptranscriptC.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">15</text>
    </paragraph>
  </paragraphlayout>
</richtext>"""
                clipC.transcripts.append(cliptranscriptC)
                try:
                    clipC.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Clip, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Clip, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        cliptranscriptAB = None
        if 64 in testsToRun:
            # Clip Saving
            testName = 'Creating Clip : AB'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                clipAB = Clip.Clip('AB', collection1.id, collection1.parent)
                # If the Clip already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Clip doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                clipAB = Clip.Clip()
                clipAB.id = 'AB'
                clipAB.episode_num = episodeA.number
                clipAB.clip_start = 20000
                clipAB.clip_stop = 25000
                clipAB.collection_num = collection1.number
                clipAB.collection_id = collection1.id
                clipAB.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                clipAB.offset = 0
                clipAB.media_filename = episodeA.media_filename
                clipAB.audio = 1
                clipAB.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                clipAB.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                # Prepare the Clip Transcript Object
                cliptranscriptAB = Transcript.Transcript()
                cliptranscriptAB.episode_num = episodeA.number
                cliptranscriptAB.source_transcript = transcriptA.number
                cliptranscriptAB.clip_start = clipAB.clip_start
                cliptranscriptAB.clip_stop = clipAB.clip_stop
                cliptranscriptAB.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">20</text>
    </paragraph>
  </paragraphlayout>
</richtext>"""
                clipAB.transcripts.append(cliptranscriptAB)
                try:
                    clipAB.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Clip, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Clip, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        cliptranscriptAC = None
        if 65 in testsToRun:
            # Clip Saving
            testName = 'Creating Clip : AC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                clipAC = Clip.Clip('AC', collection1.id, collection1.parent)
                # If the Clip already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Clip doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                clipAC = Clip.Clip()
                clipAC.id = 'AC'
                clipAC.episode_num = episodeA.number
                clipAC.clip_start = 25000
                clipAC.clip_stop = 30000
                clipAC.collection_num = collection1.number
                clipAC.collection_id = collection1.id
                clipAC.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                clipAC.offset = 0
                clipAC.media_filename = episodeA.media_filename
                clipAC.audio = 1
                clipAC.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                clipAC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                # Prepare the Clip Transcript Object
                cliptranscriptAC = Transcript.Transcript()
                cliptranscriptAC.episode_num = episodeA.number
                cliptranscriptAC.source_transcript = transcriptA.number
                cliptranscriptAC.clip_start = clipAC.clip_start
                cliptranscriptAC.clip_stop = clipAC.clip_stop
                cliptranscriptAC.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">25</text>
    </paragraph>
  </paragraphlayout>
</richtext>"""
                clipAC.transcripts.append(cliptranscriptAC)
                try:
                    clipAC.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Clip, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Clip, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        cliptranscriptBC = None
        if 66 in testsToRun:
            # Clip Saving
            testName = 'Creating Clip : BC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                clipBC = Clip.Clip('BC', collection1.id, collection1.parent)
                # If the Clip already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Clip doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                clipBC = Clip.Clip()
                clipBC.id = 'BC'
                clipBC.episode_num = episodeA.number
                clipBC.clip_start = 30000
                clipBC.clip_stop = 35000
                clipBC.collection_num = collection1.number
                clipBC.collection_id = collection1.id
                clipBC.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                clipBC.offset = 0
                clipBC.media_filename = episodeA.media_filename
                clipBC.audio = 1
                clipBC.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                clipBC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                # Prepare the Clip Transcript Object
                cliptranscriptBC = Transcript.Transcript()
                cliptranscriptBC.episode_num = episodeA.number
                cliptranscriptBC.source_transcript = transcriptA.number
                cliptranscriptBC.clip_start = clipBC.clip_start
                cliptranscriptBC.clip_stop = clipBC.clip_stop
                cliptranscriptBC.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">30</text>
    </paragraph>
  </paragraphlayout>
</richtext>"""
                clipBC.transcripts.append(cliptranscriptBC)
                try:
                    clipBC.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Clip, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Clip, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        cliptranscriptABC = None
        if 67 in testsToRun:
            # Clip Saving
            testName = 'Creating Clip : ABC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                clipABC = Clip.Clip('ABC', collection1.id, collection1.parent)
                # If the Clip already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Clip doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                clipABC = Clip.Clip()
                clipABC.id = 'ABC'
                clipABC.episode_num = episodeA.number
                clipABC.clip_start = 35000
                clipABC.clip_stop = 40000
                clipABC.collection_num = collection1.number
                clipABC.collection_id = collection1.id
                clipABC.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                clipABC.offset = 0
                clipABC.media_filename = episodeA.media_filename
                clipABC.audio = 1
                clipABC.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                clipABC.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                clipABC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                # Prepare the Clip Transcript Object
                cliptranscriptABC = Transcript.Transcript()
                cliptranscriptABC.episode_num = episodeA.number
                cliptranscriptABC.source_transcript = transcriptA.number
                cliptranscriptABC.clip_start = clipABC.clip_start
                cliptranscriptABC.clip_stop = clipABC.clip_stop
                cliptranscriptABC.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">35</text>
    </paragraph>
  </paragraphlayout>
</richtext>"""
                clipABC.transcripts.append(cliptranscriptABC)
                try:
                    clipABC.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Clip, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Clip, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 71 in testsToRun:
            # Snapshot Saving
            testName = 'Creating Snapshot : A'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                snapshotA = Snapshot.Snapshot('A', collection1.number)
                # If the Snapshot already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Snapshot doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                snapshotA = Snapshot.Snapshot()
                snapshotA.id = 'A'
                snapshotA.collection_num = collection1.number
                snapshotA.collection_id = collection1.id
                snapshotA.image_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Images', 'ch130214.gif')
                snapshotA.image_scale = 1.19066666667
                snapshotA.image_size = (768, 391)
                snapshotA.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                snapshotA.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                try:
                    snapshotA.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Snapshot, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Snapshot, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            
        if 72 in testsToRun:
            # Snapshot Saving
            testName = 'Creating Snapshot : B'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                snapshotB = Snapshot.Snapshot('B', collection1.number)
                # If the Snapshot already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Snapshot doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                snapshotB = Snapshot.Snapshot()
                snapshotB.id = 'B'
                snapshotB.collection_num = collection1.number
                snapshotB.collection_id = collection1.id
                snapshotB.image_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Images', 'ch130214.gif')
                snapshotB.image_scale = 1.19066666667
                snapshotB.image_size = (768, 391)
                snapshotB.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                snapshotB.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                try:
                    snapshotB.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Snapshot, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Snapshot, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            
        if 73 in testsToRun:
            # Snapshot Saving
            testName = 'Creating Snapshot : C'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                snapshotC = Snapshot.Snapshot('C', collection1.number)
                # If the Snapshot already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Snapshot doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                snapshotC = Snapshot.Snapshot()
                snapshotC.id = 'C'
                snapshotC.collection_num = collection1.number
                snapshotC.collection_id = collection1.id
                snapshotC.image_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Images', 'ch130214.gif')
                snapshotC.image_scale = 1.19066666667
                snapshotC.image_size = (768, 391)
                snapshotC.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                snapshotC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                try:
                    snapshotC.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Snapshot, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Snapshot, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            
        if 74 in testsToRun:
            # Snapshot Saving
            testName = 'Creating Snapshot : AB'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                snapshotAB = Snapshot.Snapshot('AB', collection1.number)
                # If the Snapshot already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Snapshot doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                snapshotAB = Snapshot.Snapshot()
                snapshotAB.id = 'AB'
                snapshotAB.collection_num = collection1.number
                snapshotAB.collection_id = collection1.id
                snapshotAB.image_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Images', 'ch130214.gif')
                snapshotAB.image_scale = 1.19066666667
                snapshotAB.image_size = (768, 391)
                snapshotAB.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                snapshotAB.codingObjects = {0: {'keywordGroup': u'SearchDemo',
                                                 'keyword': u'A',
                                                 'x1': -294L,
                                                 'x2': -150L,
                                                 'y1': 91L,
                                                 'y2': -92L,
                                                 'visible': True},
                                             1: {'keywordGroup': u'SearchDemo',
                                                 'keyword': u'B',
                                                 'x1': -144L,
                                                 'x2': -3L,
                                                 'y1': 91L,
                                                 'y2': -92L,
                                                 'visible': True}}
                snapshotAB.keywordStyles = {(u'SearchDemo', u'A'): {'lineColorDef': u'#ff0000',
                                                                     'drawMode': u'Rectangle',
                                                                     'lineStyle': u'Solid',
                                                                     'lineWidth': '5',
                                                                     'lineColorName': u'Red'},
                                             (u'SearchDemo', u'B'): {'lineColorDef': u'#00ff80',
                                                                     'drawMode': u'Rectangle',
                                                                     'lineStyle': u'LongDash',
                                                                     'lineWidth': '5',
                                                                     'lineColorName': u'Green Blue'}}
                try:
                    snapshotAB.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Snapshot, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Snapshot, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            
        if 75 in testsToRun:
            # Snapshot Saving
            testName = 'Creating Snapshot : AC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                snapshotAC = Snapshot.Snapshot('AC', collection1.number)
                # If the Snapshot already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Snapshot doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                snapshotAC = Snapshot.Snapshot()
                snapshotAC.id = 'AC'
                snapshotAC.collection_num = collection1.number
                snapshotAC.collection_id = collection1.id
                snapshotAC.image_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Images', 'ch130214.gif')
                snapshotAC.image_scale = 1.19066666667
                snapshotAC.image_size = (768, 391)
                snapshotAC.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                snapshotAC.codingObjects = {0: {'keywordGroup': u'SearchDemo',
                                                 'keyword': u'A',
                                                 'x1': -294L,
                                                 'x2': -150L,
                                                 'y1': 91L,
                                                 'y2': -92L,
                                                 'visible': True},
                                             1: {'keywordGroup': u'SearchDemo',
                                                 'keyword': u'C',
                                                 'x1': 5L,
                                                 'x2': 144L,
                                                 'y1': 91L,
                                                 'y2': -92L,
                                                 'visible': True}}
                snapshotAC.keywordStyles = {(u'SearchDemo', u'A'): {'lineColorDef': u'#ff0000',
                                                                     'drawMode': u'Rectangle',
                                                                     'lineStyle': u'Solid',
                                                                     'lineWidth': '5',
                                                                     'lineColorName': u'Red'},
                                             (u'SearchDemo', u'C'): {'lineColorDef': u'#0000ff',
                                                                     'drawMode': u'Rectangle',
                                                                     'lineStyle': u'DotDash',
                                                                     'lineWidth': '5',
                                                                     'lineColorName': u'Blue'}}
                try:
                    snapshotAC.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Snapshot, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Snapshot, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            
        if 76 in testsToRun:
            # Snapshot Saving
            testName = 'Creating Snapshot : BC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                snapshotBC = Snapshot.Snapshot('BC', collection1.number)
                # If the Snapshot already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Snapshot doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                snapshotBC = Snapshot.Snapshot()
                snapshotBC.id = 'BC'
                snapshotBC.collection_num = collection1.number
                snapshotBC.collection_id = collection1.id
                snapshotBC.image_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Images', 'ch130214.gif')
                snapshotBC.image_scale = 1.19066666667
                snapshotBC.image_size = (768, 391)
                snapshotBC.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                snapshotBC.codingObjects = {0: {'keywordGroup': u'SearchDemo',
                                                 'keyword': u'B',
                                                 'x1': -144L,
                                                 'x2': -3L,
                                                 'y1': 91L,
                                                 'y2': -92L,
                                                 'visible': True},
                                            1: {'keywordGroup': u'SearchDemo',
                                                 'keyword': u'C',
                                                 'x1': 5L,
                                                 'x2': 144L,
                                                 'y1': 91L,
                                                 'y2': -92L,
                                                 'visible': True}}
                snapshotBC.keywordStyles = {(u'SearchDemo', u'C'): {'lineColorDef': u'#0000ff',
                                                                     'drawMode': u'Rectangle',
                                                                     'lineStyle': u'DotDash',
                                                                     'lineWidth': '5',
                                                                     'lineColorName': u'Blue'},
                                             (u'SearchDemo', u'B'): {'lineColorDef': u'#00ff80',
                                                                     'drawMode': u'Rectangle',
                                                                     'lineStyle': u'LongDash',
                                                                     'lineWidth': '5',
                                                                     'lineColorName': u'Green Blue'}}
                try:
                    snapshotBC.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Snapshot, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Snapshot, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            
        if 77 in testsToRun:
            # Snapshot Saving
            testName = 'Creating Snapshot : ABC'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                snapshotABC = Snapshot.Snapshot('ABC', collection1.number)
                # If the Snapshot already exists, consider the test passed.
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            # If the Snapshot doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... create it
                snapshotABC = Snapshot.Snapshot()
                snapshotABC.id = 'ABC'
                snapshotABC.collection_num = collection1.number
                snapshotABC.collection_id = collection1.id
                snapshotABC.image_filename = os.path.join('C:\\Users', 'DavidWoods', 'Videos', 'Images', 'ch130214.gif')
                snapshotABC.image_scale = 1.19066666667
                snapshotABC.image_size = (768, 391)
                snapshotABC.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                snapshotABC.add_keyword(keywordA.keywordGroup, keywordA.keyword)
                snapshotABC.add_keyword(keywordB.keywordGroup, keywordB.keyword)
                snapshotABC.add_keyword(keywordC.keywordGroup, keywordC.keyword)
                snapshotABC.codingObjects = {0: {'keywordGroup': u'SearchDemo',
                                                 'keyword': u'A',
                                                 'x1': -294L,
                                                 'x2': -150L,
                                                 'y1': 91L,
                                                 'y2': -92L,
                                                 'visible': True},
                                             1: {'keywordGroup': u'SearchDemo',
                                                 'keyword': u'B',
                                                 'x1': -144L,
                                                 'x2': -3L,
                                                 'y1': 91L,
                                                 'y2': -92L,
                                                 'visible': True},
                                             2: {'keywordGroup': u'SearchDemo',
                                                 'keyword': u'C',
                                                 'x1': 5L,
                                                 'x2': 144L,
                                                 'y1': 91L,
                                                 'y2': -92L,
                                                 'visible': True}}
                snapshotABC.keywordStyles = {(u'SearchDemo', u'A'): {'lineColorDef': u'#ff0000',
                                                                     'drawMode': u'Rectangle',
                                                                     'lineStyle': u'Solid',
                                                                     'lineWidth': '5',
                                                                     'lineColorName': u'Red'},
                                             (u'SearchDemo', u'C'): {'lineColorDef': u'#0000ff',
                                                                     'drawMode': u'Rectangle',
                                                                     'lineStyle': u'DotDash',
                                                                     'lineWidth': '5',
                                                                     'lineColorName': u'Blue'},
                                             (u'SearchDemo', u'B'): {'lineColorDef': u'#00ff80',
                                                                     'drawMode': u'Rectangle',
                                                                     'lineStyle': u'LongDash',
                                                                     'lineWidth': '5',
                                                                     'lineColorName': u'Green Blue'}}
                try:
                    snapshotABC.db_save()
                    # If we create the Clip, consider the test passed.
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # If we can't create the Snapshot, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                # If we can't load or create the Snapshot, consider the test failed
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            

        del(series1)
        del(episodeA)
        del(episodeB)
        del(episodeC)
        del(episodeAB)
        del(episodeAC)
        del(episodeBC)
        del(episodeABC)
        del(transcriptA)
        del(transcriptB)
        del(transcriptC)
        del(transcriptAB)
        del(transcriptAC)
        del(transcriptBC)
        del(transcriptABC)
        del(collection1)
        del(clipA)
        del(cliptranscriptA)
        del(clipB)
        del(cliptranscriptB)
        del(clipC)
        del(cliptranscriptC)
        del(clipAB)
        del(cliptranscriptAB)
        del(clipAC)
        del(cliptranscriptAC)
        del(clipBC)
        del(cliptranscriptBC)
        del(clipABC)
        del(cliptranscriptABC)
        del(keywordA)
        del(keywordB)
        del(keywordC)

        if 101 in testsToRun:
            testName = 'QuickSearch A'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 1, kwg='SearchDemo', kw='A')
            msg = 'SearchDemo : A should show:\n\n  A\n  AB\n  AC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), 'SearchDemo : A'), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 102 in testsToRun:
            testName = 'NOT SearchDemo : A'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 2, searchName = testName, searchTerms=[u'NOT SearchDemo:A'])
            msg = testName + ' should show:\n\n  B\n  C\n  BC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 103 in testsToRun:
            testName = 'QuickSearch B'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 3, kwg='SearchDemo', kw='B')
            msg = 'SearchDemo:B should show:\n\n  B\n  AB\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), 'SearchDemo : B'), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 104 in testsToRun:
            testName = 'NOT SearchDemo : B'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 4, searchName = testName, searchTerms=[u'NOT SearchDemo:B'])
            msg = testName + ' should show:\n\n  A\n  C\n  AC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 105 in testsToRun:
            testName = 'QuickSearch C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 5, kwg='SearchDemo', kw='C')
            msg = 'SearchDemo:C should show:\n\n  C\n  AC\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), 'SearchDemo : C'), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 106 in testsToRun:
            testName = 'NOT SearchDemo:C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 6, searchName = testName, searchTerms=[u'NOT SearchDemo:C'])
            msg = testName + ' should show:\n\n  A\n  B\n  AB\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 110 in testsToRun:
            testName = 'SearchDemo:A AND SearchDemo:B'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 7, searchName = testName, searchTerms=[u'SearchDemo:A AND', u'SearchDemo:B'])
            msg = testName + ' should show:\n\n  AB\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 111 in testsToRun:
            testName = 'SearchDemo:A AND NOT SearchDemo:B'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 8, searchName = testName, searchTerms=[u'SearchDemo:A AND', u'NOT SearchDemo:B'])
            msg = testName + ' should show:\n\n  A\n  AC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 112 in testsToRun:
            testName = 'SearchDemo:A OR SearchDemo:B'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 9, searchName = testName, searchTerms=[u'SearchDemo:A OR', u'SearchDemo:B'])
            msg = testName + ' should show:\n\n  A\n  B\n  AB\n  AC\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 113 in testsToRun:
            testName = 'SearchDemo:A AND SearchDemo:B AND SearchDemo:C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 10, searchName = testName, searchTerms=[u'SearchDemo:A AND', u'SearchDemo:B AND', u'SearchDemo:C'])
            msg = testName + ' should show:\n\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 114 in testsToRun:
            testName = 'SearchDemo:A OR SearchDemo:B OR SearchDemo:C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 11, searchName = testName, searchTerms=[u'SearchDemo:A OR', u'SearchDemo:B OR', u'SearchDemo:C'])
            msg = testName + ' should show:\n\n  A\n  B\n  C\n  AB\n  AC\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 120 in testsToRun:
            testName = 'SearchDemo:A AND SearchDemo:B OR SearchDemo:C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 12, searchName = testName, searchTerms=[u'SearchDemo:A AND', u'SearchDemo:B OR', u'SearchDemo:C'])
            msg = testName + ' should show:\n\n  C\n  AB\n  AC\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 121 in testsToRun:
            testName = '(SearchDemo:A AND SearchDemo:B) OR SearchDemo:C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 13, searchName = testName, searchTerms=[u'(SearchDemo:A AND', u'SearchDemo:B) OR', u'SearchDemo:C'])
            msg = testName + ' should show:\n\n  C\n  AB\n  AC\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 122 in testsToRun:
            testName = 'SearchDemo:A AND (SearchDemo:B OR SearchDemo:C)'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 14, searchName = testName, searchTerms=[u'SearchDemo:A AND', u'(SearchDemo:B OR', u'SearchDemo:C)'])
            msg = testName + ' should show:\n\n  AB\n  AC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)
        
        if 123 in testsToRun:
            testName = 'SearchDemo:A OR SearchDemo:B AND SearchDemo:C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 15, searchName = testName, searchTerms=[u'SearchDemo:A OR', u'SearchDemo:B AND', u'SearchDemo:C'])
            msg = testName + ' should show:\n\n  A\n  AB\n  AC\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 124 in testsToRun:
            testName = 'SearchDemo:A OR (SearchDemo:B AND SearchDemo:C)'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 16, searchName = testName, searchTerms=[u'SearchDemo:A OR', u'(SearchDemo:B AND', u'SearchDemo:C)'])
            msg = testName + ' should show:\n\n  A\n  AB\n  AC\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 125 in testsToRun:
            testName = '(SearchDemo:A OR SearchDemo:B) AND SearchDemo:C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 17, searchName = testName, searchTerms=[u'(SearchDemo:A OR', u'SearchDemo:B) AND', u'SearchDemo:C'])
            msg = testName + ' should show:\n\n  AC\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 130 in testsToRun:
            testName = 'SearchDemo:A AND SearchDemo:B AND NOT SearchDemo:C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 18, searchName = testName, searchTerms=[u'SearchDemo:A AND', u'SearchDemo:B AND', u'NOT SearchDemo:C'])
            msg = testName + ' should show:\n\n  AB\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 131 in testsToRun:
            testName = 'SearchDemo:A OR (SearchDemo:B AND NOT SearchDemo:C)'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 19, searchName = testName, searchTerms=[u'SearchDemo:A OR', u'(SearchDemo:B AND', u'NOT SearchDemo:C)'])
            msg = testName + ' should show:\n\n  A\n  B\n  AB\n  AC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 132 in testsToRun:
            testName = '(SearchDemo:A OR SearchDemo:B) AND NOT SearchDemo:C'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 20, searchName = testName, searchTerms=[u'(SearchDemo:A OR', u'SearchDemo:B) AND', u'NOT SearchDemo:C'])
            msg = testName + ' should show:\n\n  A\n  B\n  AB\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 133 in testsToRun:
            testName = '(SearchDemo:A AND SearchDemo:B) or (SearchDemo:A AND SearchDemo:C)'
            self.SetStatusText(testName)
            search = ProcessSearch.ProcessSearch(self.tree, 21, searchName = testName, searchTerms=[u'(SearchDemo:A AND', u'SearchDemo:B) OR', u'(SearchDemo:A AND', u'SearchDemo:C)'])
            msg = testName + ' should show:\n\n  AB\n  AC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), testName), 'SearchResultsNode')
            self.tree.Collapse(node)

        if 201 in testsToRun:
            testName = 'Search for A in Episodes Only'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 201)
            msg = 'SearchDemo : A should show:\n\n  A\n  AB\n  AC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), 'Search 201'), 'SearchResultsNode')
            if node != None:
                self.tree.Collapse(node)

        if 202 in testsToRun:
            testName = 'Search for B in Clips Only'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 202)
            msg = 'SearchDemo : B should show:\n\n  B\n  AB\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), 'Search 202'), 'SearchResultsNode')
            if node != None:
                self.tree.Collapse(node)

        if 203 in testsToRun:
            testName = 'Search for C in Snapshots Only'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 203)
            msg = 'SearchDemo : C should show:\n\n  C\n  AC\n  BC\n  ABC\n\nDoes it?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), 'Search 203'), 'SearchResultsNode')
            if node != None:
                self.tree.Collapse(node)

        if 204 in testsToRun:
            testName = 'Search for "A AND".  Is the Search Button enabled?\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 204)
            msg = 'Was the Search button disabled?'
            self.ConfirmTest(msg, testName)

        if 205 in testsToRun:
            testName = 'Search for "B OR".  Is the Search Button enabled?\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 205)
            msg = 'Was the Search button disabled?'
            self.ConfirmTest(msg, testName)

        if 206 in testsToRun:
            testName = 'Search for "NOT".  Is the Search Button enabled?\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 206)
            msg = 'Was the Search button disabled?'
            self.ConfirmTest(msg, testName)

        if 207 in testsToRun:
            testName = 'Search for "C AND NOT".  Is the Search Button enabled?\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 207)
            msg = 'Was the Search button disabled?'
            self.ConfirmTest(msg, testName)

        if 208 in testsToRun:
            testName = 'Search for "A OR NOT".  Is the Search Button enabled?\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 208)
            msg = 'Was the Search button disabled?'
            self.ConfirmTest(msg, testName)

        if 209 in testsToRun:
            testName = 'Search for "(B AND C".  Is the Search Button enabled?\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 209)
            msg = 'Was the Search button disabled?'
            self.ConfirmTest(msg, testName)

        if 210 in testsToRun:
            testName = 'Try to close Parentheses that were never opened.\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 207)
            msg = 'Was it impossible to close non-opened parentheses?'
            self.ConfirmTest(msg, testName)

        if 221 in testsToRun:
            testName = 'Specify a multi-step query.  Press the "Undo" button repeatedly.\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 221)
            msg = 'Was "Undo" behavior reasonable?'
            self.ConfirmTest(msg, testName)

        if 222 in testsToRun:
            testName = 'Specify a multi-step query.  Press the "Reset" button.\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 222)
            msg = 'Was the Search Query cleared?'
            self.ConfirmTest(msg, testName)

        if 231 in testsToRun:
            testName = 'Specify a multi-step query.  Press the "Save" button.\n\nPress OK to run the query.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 231)
            msg = 'Did the query run?'
            self.ConfirmTest(msg, testName)

        if 232 in testsToRun:
            testName = 'Press the "Load" button to load the query you just saved.\n\nPress OK to run the query.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 232)
            msg = 'Did the same query run?'
            self.ConfirmTest(msg, testName)
            node = self.tree.select_Node((_('Search'), 'Search 231'), 'SearchResultsNode')
            if node != None:
                self.tree.Collapse(node)
            node = self.tree.select_Node((_('Search'), 'Search 232'), 'SearchResultsNode')
            if node != None:
                self.tree.Collapse(node)

        if 233 in testsToRun:
            testName = 'Delete the saved query.\n\nPress Cancel.'
            self.SetStatusText(testName)
            msg = testName
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 233)
            msg = 'Was the query deleted?'
            self.ConfirmTest(msg, testName)

        if 234 in testsToRun:
            testName = 'Search Query Form > Help'
            self.SetStatusText(testName)
            msg = 'You will see the Search Query Form.'
            msg += '\n\n'
            msg += 'Press HELP, then CANCEL.'
            self.ShowMessage(msg)
            search = ProcessSearch.ProcessSearch(self.tree, 234)
            msg = 'Did you see the Search Help?\n\n'
            self.ConfirmTest(msg, testName)

        self.txtCtrl.AppendText('All tests completed.')
        self.txtCtrl.AppendText('\nFinal Summary:  Total Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
        DBInterface.close_db()

    def ShowMessage(self, msg):
        dlg = Dialogs.InfoDialog(self, msg)
        dlg.ShowModal()
        dlg.Destroy()

    def ConfirmTest(self, msg, testName):
        dlg = Dialogs.QuestionDialog(self, msg, "Unit Test 1: Objects and Forms", noDefault=True)
        result = dlg.LocalShowModal()
        dlg.Destroy()
        self.testsRun += 1
        self.txtCtrl.AppendText('Test "%s" ' % testName)
        if result == wx.ID_YES:
            self.txtCtrl.AppendText('Passed.')
            self.testsSuccessful += 1
        else:
            self.txtCtrl.AppendText('FAILED.')
            self.testsFailed += 1
        self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))


class DBLogin:
    def __init__(self, username, password, dbServer, databaseName, port):
        self.username = username
        self.password = password
        self.dbServer = dbServer
        self.databaseName = databaseName
        self.port = port

class MyApp(wx.App):
   def OnInit(self):
      frame = FormCheck(None, -1, "Unit Test 3: Search")
      self.SetTopWindow(frame)
      return True
      

app = MyApp(0)
app.MainLoop()
