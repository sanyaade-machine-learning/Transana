# -*- coding: cp1252 -*-
import wx

import gettext

# This module expects i18n.  Enable it here.
__builtins__._ = wx.GetTranslation

import os
import sys

import TransanaConstants
import TransanaGlobal
import About
import Clip
import ClipPropertiesForm
import Collection
import CollectionPropertiesForm
import ConfigData
import ControlObjectClass
import CoreData
import DBInterface
import Dialogs
import Episode
import EpisodePropertiesForm
import KeywordObject
import KeywordPropertiesForm
import MenuWindow
import Note
import NotePropertiesForm
import Series
import SeriesPropertiesForm
import Snapshot
import SnapshotPropertiesForm
import TransanaExceptions
import Transcript
import TranscriptPropertiesForm

MENU_FILE_EXIT = 101

class FormCheck(wx.Frame):
    """ This window displays a variety of GUI Widgets. """
    def __init__(self,parent,id,title):

        wx.SetDefaultPyEncoding('utf8')
       
        # Define the Configuration Data
        TransanaGlobal.configData = ConfigData.ConfigData()

        print "Original videoPath:", TransanaGlobal.configData.videoPath
        print "Original language:", TransanaGlobal.configData.language

        # Clear the default language, so that the language used for these tests can be selected.
        TransanaGlobal.configData.language = ''

        # Create the global transana graphics colors, once the ConfigData object exists.
        TransanaGlobal.transana_graphicsColorList = TransanaGlobal.getColorDefs(TransanaGlobal.configData.colorConfigFilename)
        # Set essential global color manipulation data structures once the ConfigData object exists.
        (TransanaGlobal.transana_colorNameList, TransanaGlobal.transana_colorLookup, TransanaGlobal.keywordMapColourSet) = TransanaGlobal.SetColorVariables()
        if 'wxMSW' in wx.PlatformInfo:
            TransanaGlobal.configData.videoPath = 'C:\\Users\\DavidWoods\\Video'
        elif 'wxMac' in wx.PlatformInfo:
            TransanaGlobal.configData.videoPath = u'/Volumes/Vidëo'

        print "Modified videoPath:", TransanaGlobal.configData.videoPath
        
        TransanaGlobal.menuWindow = MenuWindow.MenuWindow(None, -1, title)

        print "Modified language:", TransanaGlobal.configData.language

        self.ControlObject = ControlObjectClass.ControlObject()
        self.ControlObject.Register(Menu = TransanaGlobal.menuWindow)
        TransanaGlobal.menuWindow.Register(ControlObject = self.ControlObject)

        wx.Frame.__init__(self,parent,-1, title, size = (800,600), style=wx.DEFAULT_FRAME_STYLE|wx.NO_FULL_REPAINT_ON_RESIZE)
        self.testsRun = 0
        self.testsSuccessful = 0
        self.testsFailed = 0
       
        self.SetBackgroundColour(wx.RED)

        print "Current videoPath:", TransanaGlobal.configData.videoPath
        
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        self.txtCtrl = wx.TextCtrl(self, -1, "Unit Test:  Objects and Forms\n\n", style=wx.TE_LEFT | wx.TE_MULTILINE)
        self.txtCtrl.AppendText("Transana Version:  %s     singleUserVersion:  %s\n\n" % (TransanaConstants.versionNumber, TransanaConstants.singleUserVersion))

        self.txtCtrl.AppendText("Translation Check (%s):  %s  =  %s\n\n" % (TransanaGlobal.configData.language, 'Downloading . . .', _('Downloading . . .')))

        mainSizer.Add(self.txtCtrl, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()
        self.CenterOnScreen()
        # Status Bar
        self.CreateStatusBar()
        self.SetStatusText("")
        self.Show(True)
        self.RunTests()

    def RunTests(self):
        # Tests defined:
        # 1 - 100     Basic Dialogs
        # 101 - 200    Database Connetion
        # 201 - 300    Keyword
        # 301 - 400    Series
        # 401 - 500    Episode
        # 501 - 600    Transcript
        # 601 - 700    Collection
        # 701 - 800    Clip
        # 801 - 900    Snapshot
        # 901 - 1000   Note
        testsNotToSkip = [102, 201, 213, 301, 401, 501, 601, 603, 701, 801, 901]
        startAtTest = 1  # Should start at 1, not 0!
        endAtTest = 1000   # Should be one more than the last test to be run!
        testsToRun = testsNotToSkip + range(startAtTest, endAtTest)

        if 1 in testsToRun:      
            # InfoDialog
            testName = 'InfoDialog'
            self.SetStatusText(testName)
            msg = "Welcome to Unit Test 1:  Objects and Forms.\n\n"
            msg += 'You are about to see an InfoDialog with the message "This is an InfoDialog!".'
            self.ShowMessage(msg)
            msg = "This is an InfoDialog!"
            dlg = Dialogs.InfoDialog(self, msg, dlgTitle='Unit Test 1: Objects and Forms')
            dlg.ShowModal()
            dlg.Destroy()
            self.ConfirmTest('Did you see the correct InfoDialog?', testName)

        if 2 in testsToRun:
            # ErrorDialog
            testName = 'ErrorDialog'
            self.SetStatusText(testName)
            msg = 'You are about to see a simple ErrorDialog with the message "This is an Error Message!".'
            self.ShowMessage(msg)
            msg = "This is an Error Message!"
            dlg = Dialogs.ErrorDialog(self, msg)
            dlg.ShowModal()
            dlg.Destroy()
            self.ConfirmTest('Did you see the correct ErrorDialog?', testName)

        if 3 in testsToRun:
            # ErrorDialog with SkipCheck
            testName = 'ErrorDialog with SkipCheck'
            self.SetStatusText(testName)
            msg = 'You are about to see an ErrorDialog with the message "This is an Error Message!"\n'
            msg += 'with includeSkipCheck enabled.\n\n'
            msg += 'DO NOT check the box.'
            self.ShowMessage(msg)
            msg = "This is an Error Message!"
            dlg = Dialogs.ErrorDialog(self, msg, includeSkipCheck=True)
            dlg.ShowModal()
            result = dlg.GetSkipCheck()
            dlg.Destroy()
            msg = 'Was SkipCheck enabled in the ErrorDialog?\n\n'
            if result:
                msg += 'The box WAS checked!'
            else:
                msg += 'The box was not checked.'
            self.ConfirmTest(msg, testName)

        if 4 in testsToRun:
            # ErrorDialog with SkipCheck
            testName = 'ErrorDialog with SkipCheck 2'
            self.SetStatusText(testName)
            msg = 'You are about to see an ErrorDialog with the message "This is an Error Message!"\n'
            msg += 'with includeSkipCheck enabled.\n\n'
            msg += 'Check the box.'
            self.ShowMessage(msg)
            msg = "This is an Error Message!"
            dlg = Dialogs.ErrorDialog(self, msg, includeSkipCheck=True)
            dlg.ShowModal()
            result = dlg.GetSkipCheck()
            dlg.Destroy()
            msg = 'Was SkipCheck enabled in the ErrorDialog?\n\n'
            if result:
                msg += 'The box WAS checked!'
            else:
                msg += 'The box was not checked.'
            self.ConfirmTest(msg, testName)

        if 5 in testsToRun:
            # QuestionDialog
            testName = 'QuestionDialog'
            self.SetStatusText(testName)
            msg = 'You are about to see a yes/no QuestionDialog defaulting to Yes with the question "Should you say Yes?".'
            msg += '\n\n'
            msg += 'Say %s.' % _("&Yes")
            self.ShowMessage(msg)
            msg = "Should you say %s?" % _("&Yes")
            dlg = Dialogs.QuestionDialog(self, msg, header="Unit Test 1")
            result = dlg.LocalShowModal()
            dlg.Destroy()
            msg = 'Did you see the QuestionDialog?\n\n'
            if result == wx.ID_YES:
                msg += 'You said Yes!'
            else:
                msg += 'You said No.'
            self.ConfirmTest(msg, testName)

        if 6 in testsToRun:
            # QuestionDialog 2
            testName = 'QuestionDialog 2'
            self.SetStatusText(testName)
            msg = 'You are about to see a yes/no QuestionDialog defaulting to No with the question "Should you say Yes?".'
            msg += '\n\n'
            msg += 'Say NO.'
            self.ShowMessage(msg)
            msg = "Should you say Yes?"
            dlg = Dialogs.QuestionDialog(self, msg, header="Unit Test 1", noDefault=True)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            msg = 'Did you see the QuestionDialog?\n\n'
            if result == wx.ID_NO:
                msg += 'You said No.'
            else:
                msg += 'You said Yes!'
            self.ConfirmTest(msg, testName)

        if 7 in testsToRun:
            # QuestionDialog 3
            testName = 'QuestionDialog 3'
            self.SetStatusText(testName)
            msg = 'You are about to see an ok/cancel QuestionDialog defaulting to OK with the question "Should you say OK?".'
            msg += '\n\n'
            msg += 'Say OK.'
            self.ShowMessage(msg)
            msg = "Should you say OK?"
            dlg = Dialogs.QuestionDialog(self, msg, header="Unit Test 1", useOkCancel=True)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            msg = 'Did you see the QuestionDialog?\n\n'
            if result == wx.ID_OK:
                msg += 'You said OK!'
            else:
                msg += 'You said Cancel.'
            self.ConfirmTest(msg, testName)

        if 8 in testsToRun:
            # QuestionDialog 4
            testName = 'QuestionDialog 4'
            self.SetStatusText(testName)
            msg = 'You are about to see an ok/cancel QuestionDialog defaulting to Cancel with the question "Should you say OK?".'
            msg += '\n\n'
            msg += 'Say CANCEL.'
            self.ShowMessage(msg)
            msg = "Should you say OK?"
            dlg = Dialogs.QuestionDialog(self, msg, header="Unit Test 1", noDefault=True, useOkCancel=True)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            msg = 'Did you see the QuestionDialog?\n\n'
            if result == wx.ID_CANCEL:
                msg += 'You said Cancel.'
            else:
                msg += 'You said OK!'
            self.ConfirmTest(msg, testName)

        if 9 in testsToRun:
            # QuestionDialog 5 
            testName = 'QuestionDialog 5 - YesToAll'
            self.SetStatusText(testName)
            msg = 'You are about to see a Yes/YesToAll/No QuestionDialog with the question "Should you say YES?".'
            msg += '\n\n'
            msg += 'Say Yes.'
            self.ShowMessage(msg)
            msg = "Should you say Yes?"
            dlg = Dialogs.QuestionDialog(self, msg, header="Unit Test 1", yesToAll=True)
            result = dlg.LocalShowModal()
            msg = 'Did you see the QuestionDialog?\n\n'
            if result == wx.ID_NO:
                msg += 'You said No.'
            elif result == dlg.YESTOALLID:
                msg += 'You said Yes to All.'
            else:
                msg += 'You said Yes.'
            dlg.Destroy()
            self.ConfirmTest(msg, testName)

        if 10 in testsToRun:
            # QuestionDialog 6 
            testName = 'QuestionDialog 6 - YesToAll'
            self.SetStatusText(testName)
            msg = 'You are about to see a Yes/YesToAll/No QuestionDialog with the question "Should you say YES TO ALL?".'
            msg += '\n\n'
            msg += 'Say Yes to All.'
            self.ShowMessage(msg)
            msg = "Should you say YES TO ALL?"
            dlg = Dialogs.QuestionDialog(self, msg, header="Unit Test 1", yesToAll=True)
            result = dlg.LocalShowModal()
            msg = 'Did you see the QuestionDialog?\n\n'
            if result == wx.ID_NO:
                msg += 'You said No.'
            elif result == dlg.YESTOALLID:
                msg += 'You said Yes to All.'
            else:
                msg += 'You said Yes.'
            dlg.Destroy()
            self.ConfirmTest(msg, testName)

        if 11 in testsToRun:
            # QuestionDialog 7
            testName = 'QuestionDialog 7 - YesToAll'
            self.SetStatusText(testName)
            msg = 'You are about to see a Yes/YesToAll/No QuestionDialog with the question "Should you say Yes?".'
            msg += '\n\n'
            msg += 'Say No.'
            self.ShowMessage(msg)
            msg = "Should you say Yes?"
            dlg = Dialogs.QuestionDialog(self, msg, header="Unit Test 1", yesToAll=True)
            result = dlg.LocalShowModal()
            msg = 'Did you see the QuestionDialog?\n\n'
            if result == wx.ID_NO:
                msg += 'You said No.'
            elif result == dlg.YESTOALLID:
                msg += 'You said Yes to All.'
            else:
                msg += 'You said Yes.'
            dlg.Destroy()
            self.ConfirmTest(msg, testName)

        if 12 in testsToRun:
            # wx.TextEntryDialog
            testName = 'wx.TextEntryDialog'
            self.SetStatusText(testName)
            msg = 'You are about to see a wx.TextEntryDialog.'
            msg += '\n\n'
            msg += 'Enter a value and Say OK.'
            self.ShowMessage(msg)
            msg = "Should you say OK?"
            dlg = wx.TextEntryDialog(self, _('What is the Host / Server name other computers use to connect to this MySQL Server?'),
                                         _('Transana Message Server connection'), _('localhost'))
            result = dlg.ShowModal()
            value = dlg.GetValue()
            dlg.Destroy()
            msg = 'Did you see the wx.TextEntryDialog?\n\n'
            if result == wx.ID_CANCEL:
                msg += 'You said Cancel.'
            else:
                msg += 'You said OK!\nYou entered "%s"' % value
            self.ConfirmTest(msg, testName)

        if 13 in testsToRun:
            # wx.TextEntryDialog
            testName = 'wx.TextEntryDialog'
            self.SetStatusText(testName)
            msg = 'You are about to see a wx.TextEntryDialog.'
            msg += '\n\n'
            msg += 'Enter a value and Say CANCEL.'
            self.ShowMessage(msg)
            msg = "Should you say OK?"
            dlg = wx.TextEntryDialog(self, 'What is the Host / Server name other computers use to connect to this MySQL Server?',
                                         'Transana Message Server connection', 'localhost')
            result = dlg.ShowModal()
            value = dlg.GetValue()
            dlg.Destroy()
            msg = 'Did you see the wx.TextEntryDialog?\n\n'
            if result == wx.ID_CANCEL:
                msg += 'You said Cancel.'
            else:
                msg += 'You said OK!\nYou entered "%s"' % value
            self.ConfirmTest(msg, testName)

        if 14 in testsToRun:
            # About Box
            testName = 'About Box'
            self.SetStatusText(testName)
            msg = 'You are about to see the About Box.'
            msg += '\n\n'
            msg += 'Say OK.'
            self.ShowMessage(msg)
            About.AboutBox()            
            self.ConfirmTest('Did you see the About Box?', testName)

        if (not TransanaConstants.singleUserVersion):
            # Database Connection
            testName = 'Database Connection, set up Database'
            self.SetStatusText(testName)
            if True or (101 in testsToRun):   # Do this whether it's included in the test to run or not!!
#                import MySQLdb

                dlg = wx.PasswordEntryDialog(self, 'Please enter your database password.', 'Unit Test 1:  Database Connection')
                result = dlg.ShowModal()
                if result == wx.ID_OK:
                    password = dlg.GetValue()
                else:
                    password = ''
                dlg.Destroy()
                if 'wxMSW' in wx.PlatformInfo:
                    dbName = 'Transana_UnitTest_Win'
                elif 'wxMac' in wx.PlatformInfo:
                    dbName = 'Transana_UnitTest_Mac'
                else:
                    dbName = 'Transana_UnitTest_Linux'
                loginInfo = DBLogin('DavidW', password, 'DKW-Linux', dbName, '3306')
                dbReference = DBInterface.get_db(loginInfo)

                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                if dbReference != None:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if (102 in testsToRun) and (not TransanaConstants.singleUserVersion):
            self.txtCtrl.AppendText('Dropping all tables from ' + loginInfo.databaseName + '\n\n')
            if dbReference != None:
                dbCursor = dbReference.cursor()
                dbCursor2 = dbReference.cursor()
                dbCursor.execute('USE ' + loginInfo.databaseName)
                query = 'SHOW TABLES'
                dbCursor.execute(query)
                for result in dbCursor.fetchall():
                    dbCursor2.execute('DROP TABLE ' + result[0])
                dbCursor2.execute('COMMIT')

                self.txtCtrl.AppendText('Creating new tables for ' + loginInfo.databaseName + '\n\n')
                DBInterface.establish_db_exists()

                dbCursor.execute(query)
                for result in dbCursor.fetchall():
                    dbCursor2.execute('SELECT * FROM ' + result[0])
                    self.txtCtrl.AppendText('%s has %d rows.\n' % (result[0], dbCursor2.rowcount))
                self.txtCtrl.AppendText('\n')
                dbCursor2.close()
            else:
                # We can't run more tests if we don't connect to the database!
                testsToRun = []

        keyword1 = None
        if 201 in testsToRun:
            # Keyword Properties Form
            testName = 'Add Keyword > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Add Keyword Form.'
            msg += '\n\n'
            msg += 'Enter distinct values for all fields.'
            self.ShowMessage(msg)
            dlg = KeywordPropertiesForm.AddKeywordDialog(self, -1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Keyword data?\n\n'
            if result is None:
                msg += 'You pressed Cancel.'
            else:
                keyword1 = result
                keyword1.db_save()
                msg += 'You pressed OK!\n\n' + keyword1.__repr__()
            self.ConfirmTest(msg, testName)

        if 202 in testsToRun:
            # Keyword Properties Form
            testName = 'Add Keyword > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Add Keyword Form.'
            msg += '\n\n'
            msg += 'Press CANCEL.'
            self.ShowMessage(msg)
            dlg = KeywordPropertiesForm.AddKeywordDialog(self, -1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this say you pressed Cancel?\n\n'
            if result == None:
                msg += 'You pressed Cancel.\n\n' + keyword1.__repr__()
            else:
                keyword1 = result
                keyword1.db_save()
                msg += 'You pressed OK!\n\n' + keyword1.__repr__()
            self.ConfirmTest(msg, testName)

        if 203 in testsToRun:
            # Keyword Properties Form
            testName = 'Keyword Properties Form > Help'
            self.SetStatusText(testName)
            msg = 'You will see an Add Keyword Form.'
            msg += '\n\n'
            msg += 'Press HELP, then CANCEL.'
            self.ShowMessage(msg)
            dlg = KeywordPropertiesForm.AddKeywordDialog(self, -1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Did you see the Keyword Properties Help?\n\n'
            self.ConfirmTest(msg, testName)

        if 204 in testsToRun:
            # Alter the Keyword so that it appears to have been loaded from the database!
            keyword1.originalKeywordGroup = keyword1.keywordGroup
            keyword1.originalKeyword = keyword1.keyword
            # Keyword Properties Form
            testName = 'EDIT Keyword > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Keyword Form.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press OK.'
            self.ShowMessage(msg)
            keyword1.lock_record()
            dlg = KeywordPropertiesForm.EditKeywordDialog(self, -1, keyword1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the edited Keyword data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.\n\n' + keyword1.__repr__()
            else:
                keyword1 = result
                msg += 'You pressed OK!\n\n' + keyword1.__repr__()
                keyword1.db_save()
            keyword1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 205 in testsToRun:
            # Keyword Properties Form
            testName = 'EDIT Keyword > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Keyword Form.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press CANCEL.'
            self.ShowMessage(msg)
            keyword1.lock_record()
            dlg = KeywordPropertiesForm.EditKeywordDialog(self, -1, keyword1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this indicate you pressed Cancel, having left your keyword unchanged?\n\n'
            if result == None:
                msg += 'You pressed Cancel.\n\n' + keyword1.__repr__()
            else:
                keyword1 = result
                msg += 'You pressed OK!\n\n' + keyword1.__repr__()
                keyword1.db_save()
            keyword1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 206 in testsToRun:
            # Keyword Loading
            testName = 'Keyword.db_load_by_name()'
            self.SetStatusText(testName)
            keyword2 = KeywordObject.Keyword(keyword1.keywordGroup, keyword1.keyword)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(keyword1, keyword2):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Keyword.db_load_by_name():"
                print keyword1
                print keyword2
                print self.Compare(keyword1, keyword2)
                print
                
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(keyword2)
                
        if 207 in testsToRun:
            msg = "You should now see several error messages!"
            dlg = Dialogs.InfoDialog(self, msg, dlgTitle='Unit Test 1: Objects and Forms')
            dlg.ShowModal()
            dlg.Destroy()
            # Keyword Saving
            testName = 'Save Keyword with no Keyword Group and no Keyword'
            self.SetStatusText(testName)
            keyword2 = KeywordObject.Keyword()
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                keyword2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(keyword2)
                
        if 208 in testsToRun:
            # Keyword Saving
            testName = 'Save Keyword with no Keyword Group'
            self.SetStatusText(testName)
            keyword2 = KeywordObject.Keyword()
            keyword2.keyword = 'Test 208'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                keyword2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(keyword2)
                
        if 209 in testsToRun:
            # Keyword Saving
            testName = 'Save Keyword with no Keyword'
            self.SetStatusText(testName)
            keyword2 = KeywordObject.Keyword()
            keyword2.keywordGroup = 'Test 209'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                keyword2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(keyword2)
                
        if 210 in testsToRun:
            # Keyword Saving
            testName = 'Save Keyword with duplicate ID'
            self.SetStatusText(testName)
            keyword2 = KeywordObject.Keyword()
            keyword2.keywordGroup = keyword1.keywordGroup
            keyword2.keyword = keyword1.keyword
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                keyword2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(keyword2)

        if 211 in testsToRun:
            # Keyword Saving
            testName = 'Save a second Keyword'
            self.SetStatusText(testName)
            keyword2 = KeywordObject.Keyword()
            keyword2.keywordGroup = 'A very long keyword group name meant to challenge form layout'
            keyword2.keyword = 'A very long keyword name meant to challenge form layout'

            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                keyword2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 212 in testsToRun:
            # Keyword Properties Form
            testName = 'EDIT very long Keyword > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Keyword Form with an extreme data value.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press OK.'
            self.ShowMessage(msg)
            keyword2.originalKeywordGroup = keyword2.keywordGroup
            keyword2.originalKeyword = keyword2.keyword
            keyword2.lock_record()
            dlg = KeywordPropertiesForm.EditKeywordDialog(self, -1, keyword2)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the edited Keyword data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.\n\n' + keyword2.__repr__()
            else:
                keyword2 = result
                msg += 'You pressed OK!\n\n' + keyword2.__repr__()
                try:
                    keyword2.db_save()
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
            keyword2.unlock_record()
            self.ConfirmTest(msg, testName)
            del(keyword2)

        if 213 in testsToRun:
            # Keyword Saving
            testName = 'Save a third Keyword'
            self.SetStatusText(testName)
            keyword2 = KeywordObject.Keyword()
            keyword2.keywordGroup = keyword1.keywordGroup
            keyword2.keyword = 'This is a very long keyword name meant to challenge form layout'

            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                keyword2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        series1 = None
        if 301 in testsToRun:
            # Series Properties Form
            testName = 'Add Series > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Add Series Form.'
            msg += '\n\n'
            msg += 'Enter distinct values for all fields.'
            self.ShowMessage(msg)
            dlg = SeriesPropertiesForm.AddSeriesDialog(self, -1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Series data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
            else:
                series1 = result
                series1.db_save()
                msg += 'You pressed OK!\n\n' + series1.__repr__()
            self.ConfirmTest(msg, testName)

        if 302 in testsToRun:
            # Series Properties Form
            testName = 'Add Series > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Add Series Form.'
            msg += '\n\n'
            msg += 'Press CANCEL.'
            self.ShowMessage(msg)
            dlg = SeriesPropertiesForm.AddSeriesDialog(self, -1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this say you pressed Cancel?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if series1 != None:
                    msg += '\n\n' + series1.__repr__()
            else:
                series1 = result
                series1.db_save()
                msg += 'You pressed OK!\n\n' + series1.__repr__()
            self.ConfirmTest(msg, testName)

        if 303 in testsToRun:
            # Series Properties Form
            testName = 'Series Properties Form > Help'
            self.SetStatusText(testName)
            msg = 'You will see an Add Series Form.'
            msg += '\n\n'
            msg += 'Press HELP, then CANCEL.'
            self.ShowMessage(msg)
            dlg = SeriesPropertiesForm.AddSeriesDialog(self, -1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Did you see the Series Properties Help?\n\n'
            self.ConfirmTest(msg, testName)

        if 304 in testsToRun:
            # Series Properties Form
            testName = 'Edit Series > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Series Form.'
            msg += '\n\n'
            msg += 'Edit values for all fields and press OK.'
            self.ShowMessage(msg)
            series1.lock_record()
            dlg = SeriesPropertiesForm.EditSeriesDialog(self, -1, series1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Series data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if series1 != None:
                    msg += '\n\n' + series1.__repr__()
            else:
                series1 = result
                series1.db_save()
                msg += 'You pressed OK!\n\n' + series1.__repr__()
            series1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 305 in testsToRun:
            # Series Properties Form
            testName = 'EDIT Series > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Series Form.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press CANCEL.'
            self.ShowMessage(msg)
            series1.lock_record()
            dlg = SeriesPropertiesForm.EditSeriesDialog(self, -1, series1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this indicate you pressed Cancel, having left your Series unchanged?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if series1 != None:
                    msg += '\n\n' + series1.__repr__()
            else:
                series1 = result
                series1.db_save()
                msg += 'You pressed OK!\n\n' + series1.__repr__()
            series1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 306 in testsToRun:
            # Series Loading
            testName = 'Series.db_load_by_num()'
            self.SetStatusText(testName)
            series2 = Series.Series(series1.number)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(series1, series2, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Series.db_load_by_num():"
                print series1
                print series2
                print self.Compare(series1, series2, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(series2)
                
        if 307 in testsToRun:
            # Series Loading
            testName = 'Series.db_load_by_name()'
            self.SetStatusText(testName)
            series2 = Series.Series(series1.id)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(series1, series2, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Series.db_load_by_name():"
                print series1
                print series2
                print self.Compare(series1, series2, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(series2)
                
        if 308 in testsToRun:
            msg = "You should now see several error messages!"
            dlg = Dialogs.InfoDialog(self, msg, dlgTitle='Unit Test 1: Objects and Forms')
            dlg.ShowModal()
            dlg.Destroy()
            # Series Saving
            testName = 'Save Series with no ID'
            self.SetStatusText(testName)
            series2 = Series.Series()
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                series2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(series2)
                
        if 309 in testsToRun:
            # Series Saving
            testName = 'Save Series with duplicate ID'
            self.SetStatusText(testName)
            series2 = Series.Series()
            series2.id = series1.id
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                series2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(series2)

        if 310 in testsToRun:
            # Series Saving
            testName = 'Save a second Series'
            self.SetStatusText(testName)
            series2 = Series.Series()
            series2.id = series1.id + ' - Duplicate'
            series2.comment = 'This Series was created automatically'
            series2.owner = 'Unit Test Form Check'
            
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                series2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(series2)

        episode1 = None
        if 401 in testsToRun:
            # Episode Properties Form
            testName = 'Add Episode > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Add Episode Form.'
            msg += '\n\n'
            msg += 'Enter distinct values for all fields.'
            self.ShowMessage(msg)
            dlg = EpisodePropertiesForm.AddEpisodeDialog(self, -1, series1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Episode data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
            else:
                episode1 = result
                episode1.db_save()
                msg += 'You pressed OK!\n\n' + episode1.__repr__()
            self.ConfirmTest(msg, testName)

        if 402 in testsToRun:
            # Episode Properties Form
            testName = 'Add Episode > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Add Episode Form.'
            msg += '\n\n'
            msg += 'Press CANCEL.'
            self.ShowMessage(msg)
            dlg = EpisodePropertiesForm.AddEpisodeDialog(self, -1, series1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this say you pressed Cancel?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if episode1 != None:
                    episode1.refresh_keywords()
                    msg += '\n\n' + episode1.__repr__()
            else:
                episode1 = result
                episode1.db_save()
                msg += 'You pressed OK!\n\n' + episode1.__repr__()
            self.ConfirmTest(msg, testName)

        if 403 in testsToRun:
            # Episode Properties Form
            testName = 'Episode Properties Form > Help'
            self.SetStatusText(testName)
            msg = 'You will see an Add Episode Form.'
            msg += '\n\n'
            msg += 'Press HELP, then CANCEL.'
            self.ShowMessage(msg)
            dlg = EpisodePropertiesForm.AddEpisodeDialog(self, -1, series1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Did you see the Episode Properties Help?\n\n'
            self.ConfirmTest(msg, testName)

        if 404 in testsToRun:
            # Episode Properties Form
            testName = 'Edit Episode > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Episode Form.'
            msg += '\n\n'
            msg += 'Edit values for all fields and press OK.'
            self.ShowMessage(msg)
            episode1.lock_record()
            dlg = EpisodePropertiesForm.EditEpisodeDialog(self, -1, episode1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Episode data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if episode1 != None:
                    # The Media File can get changed ... let's reload the Episode!
                    # We don't have to worry about this in Transana, as we don't retain the
                    # object for further use once we close the form!
                    episode1 = Episode.Episode(episode1.number)
#                    episode1.refresh_keywords()
                    msg += '\n\n' + episode1.__repr__()
            else:
                episode1 = result
                episode1.db_save()
                msg += 'You pressed OK!\n\n' + episode1.__repr__()
            episode1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 405 in testsToRun:
            # Episode Properties Form
            testName = 'EDIT Episode > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Episode Form.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press CANCEL.'
            self.ShowMessage(msg)
            episode1.lock_record()
            dlg = EpisodePropertiesForm.EditEpisodeDialog(self, -1, episode1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this indicate you pressed Cancel, having left your Episode unchanged?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if episode1 != None:
                    # The Media File can get changed ... let's reload the Episode!
                    episode1 = Episode.Episode(episode1.number)
#                    episode1.refresh_keywords()
                    msg += '\n\n' + episode1.__repr__()
            else:
                episode1 = result
                episode1.db_save()
                msg += 'You pressed OK!\n\n' + episode1.__repr__()
            episode1.unlock_record()
            self.ConfirmTest(msg, testName)

        coredata1 = None
        if 406 in testsToRun:
            # Episode Properties Form
            testName = 'EDIT Episode > Core Data > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Episode Form.'
            msg += '\n\n'
            msg += 'Press the Core Data button, edit the values for all fields, then press OK.'
            self.ShowMessage(msg)
            episode1.lock_record()
            dlg = EpisodePropertiesForm.EditEpisodeDialog(self, -1, episode1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Core Data information?\n\n'
            if result != None:
                episode1 = result
                episode1.db_save()
            episode1.unlock_record()
            (path, coredatafile) = os.path.split(episode1.media_filename)
            try:
                coredata1 = CoreData.CoreData(coredatafile)
                msg += coredata1.__repr__()
            except:
                msg += 'No Core Data Record found!'
            self.ConfirmTest(msg, testName)

        if 407 in testsToRun:
            # Episode Properties Form
            testName = 'EDIT Episode > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Episode Form.'
            msg += '\n\n'
            msg += 'Press the Core Data button, edit the values for all fields, then press CANCEL.'
            self.ShowMessage(msg)
            episode1.lock_record()
            dlg = EpisodePropertiesForm.EditEpisodeDialog(self, -1, episode1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Is your Core Data unchanged?\n\n'
            if result != None:
                episode1 = result
                episode1.db_save()
            episode1.unlock_record()
            try:
                (path, coredatafile) = os.path.split(episode1.media_filename)
                coredata1 = CoreData.CoreData(coredatafile)
                msg += coredata1.__repr__()
            except RecordNotFoundError:
                pass
            self.ConfirmTest(msg, testName)

        if 408 in testsToRun:
            # Episode Loading
            testName = 'Episode.db_load_by_num()'
            self.SetStatusText(testName)
            episode2 = Episode.Episode(episode1.number)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(episode1, episode2, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Episode.db_load_by_num():"
                print episode1
                print episode2
                print self.Compare(episode1, episode2, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(episode2)

        if 409 in testsToRun:
            # Episode Loading
            testName = 'Episode.db_load_by_name()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                episode2 = Episode.Episode(series=episode1.series_id, episode=episode1.id)
                if self.Compare(episode1, episode2, True):
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                    
                    print "Episode.db_load_by_name():"
                    print episode1
                    print episode2
                    print self.Compare(episode1, episode2, True)
                    print

            except:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1])
                errordlg.ShowModal()
                errordlg.Destroy()
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(episode2)
                
        if 410 in testsToRun:
            msg = "You should now see several error messages!"
            dlg = Dialogs.InfoDialog(self, msg, dlgTitle='Unit Test 1: Objects and Forms')
            dlg.ShowModal()
            dlg.Destroy()
            # Episode Saving
            testName = 'Save Episode with no ID'
            self.SetStatusText(testName)
            episode2 = Episode.Episode(episode1.number)
            episode2.id = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                episode2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(episode2)
                
        if 411 in testsToRun:
            # Episode Saving
            testName = 'Save Episode with no Series'
            self.SetStatusText(testName)
            episode2 = Episode.Episode(episode1.number)
            episode2.series_num = 0
            episode2.series_id = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                episode2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(episode2)
                
        if 412 in testsToRun:
            # Episode Saving
            testName = 'Save Episode with no media file'
            self.SetStatusText(testName)
            episode2 = Episode.Episode(episode1.number)
            episode2.media_filename = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                episode2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(episode2)
                
        if 413 in testsToRun:
            # Episode Saving
            testName = 'Save Episode with duplicate ID'
            self.SetStatusText(testName)
            episode2 = Episode.Episode(episode1.number)
            episode2.number = 0
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                episode2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(episode2)

        if 414 in testsToRun:
            # Episode Keyword Manipulation
            testName = 'Add Keyword to Episode'
            self.SetStatusText(testName)
            episode2 = Episode.Episode(episode1.number)
            episode2.clear_keywords()
            numKeywords = len(episode2._kwlist)
            hasKeyword = episode2.has_keyword(keyword2.keywordGroup, keyword2.keyword)
            episode2.add_keyword(keyword2.keywordGroup, keyword2.keyword)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if (not hasKeyword) and (len(episode2._kwlist) == numKeywords + 1):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.  ')
                if hasKeyword:
                    self.txtCtrl.AppendText('Keyword present when it should not have been!')
                else:
                    self.txtCtrl.AppendText('Keyword NOT added correctly!')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 415 in testsToRun:
            # Episode Keyword Manipulation
            testName = 'Episode has Keyword'
            self.SetStatusText(testName)
            hasKeyword = episode2.has_keyword(keyword2.keywordGroup, keyword2.keyword)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if hasKeyword:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.  ')
                self.txtCtrl.AppendText('Episode.has_keyword() failed to detect keyword that should be present!')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 416 in testsToRun:
            # Episode Keyword Manipulation
            testName = 'Remove Keyword from Episode'
            self.SetStatusText(testName)
            numKeywords = len(episode2._kwlist)
            hasKeyword = episode2.has_keyword(keyword2.keywordGroup, keyword2.keyword)
            episode2.remove_keyword(keyword2.keywordGroup, keyword2.keyword)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if (hasKeyword) and (len(episode2._kwlist) == numKeywords - 1):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.  ')
                if not hasKeyword:
                    self.txtCtrl.AppendText('Keyword NOT present when it should not have been!')
                else:
                    self.txtCtrl.AppendText('Keyword NOT removed correctly!')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 417 in testsToRun:
            # Episode Keyword Manipulation
            testName = 'Episode Clear Keywords'
            self.SetStatusText(testName)
            self.testsRun += 1
            episode2.add_keyword(keyword1.keywordGroup, keyword1.keyword)
            episode2.add_keyword(keyword2.keywordGroup, keyword2.keyword)
            hasKeyword = (len(episode2._kwlist) >= 2)
            episode2.clear_keywords()
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if hasKeyword and (len(episode2._kwlist) == 0):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.  ')
                self.txtCtrl.AppendText('Episode.clear_keywords()!')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            del(episode2)

        transcript1 = None
        if 501 in testsToRun:
            # Transcript Properties Form
            testName = 'Add Transcript Properties > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Add Transcript Form.'
            msg += '\n\n'
            msg += 'Enter distinct values for all fields.'
            self.ShowMessage(msg)
            dlg = TranscriptPropertiesForm.AddTranscriptDialog(self, -1, episode1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Transcript data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
            else:
                transcript1 = result
                if transcript1.text == '':
                    transcript1.text = 'txt\nThis is an Episode Transcript.'
                transcript1.db_save()
                msg += 'You pressed OK!\n\n' + transcript1.__repr__()
            self.ConfirmTest(msg, testName)

        if 502 in testsToRun:
            # Transcript Properties Form
            testName = 'Add Transcript Properties > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Add Transcript Form.'
            msg += '\n\n'
            msg += 'Press CANCEL.'
            self.ShowMessage(msg)
            dlg = TranscriptPropertiesForm.AddTranscriptDialog(self, -1, episode1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this say you pressed Cancel?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if transcript1 != None:
                    msg += '\n\n' + transcript1.__repr__()
            else:
                transcript1 = result
                if transcript1.text == '':
                    transcript1.text = 'txt\nThis is an Episode Transcript.'
                transcript1.db_save()
                msg += 'You pressed OK!\n\n' + transcript1.__repr__()
            self.ConfirmTest(msg, testName)

        if 503 in testsToRun:
            # Transcript Properties Form
            testName = 'Transcript Properties Form > Help'
            self.SetStatusText(testName)
            msg = 'You will see an Add Transcript Form.'
            msg += '\n\n'
            msg += 'Press HELP, then CANCEL.'
            self.ShowMessage(msg)
            dlg = TranscriptPropertiesForm.AddTranscriptDialog(self, -1, episode1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Did you see the Transcript Properties Help?\n\n'
            self.ConfirmTest(msg, testName)

        if 504 in testsToRun:
            # Transcript Properties Form
            testName = 'Edit Transcript Properties > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Transcript Properties Form.'
            msg += '\n\n'
            msg += 'Edit values for all fields and press OK.'
            self.ShowMessage(msg)
            transcript1.lock_record()
            dlg = TranscriptPropertiesForm.EditTranscriptDialog(self, -1, transcript1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Transcript data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if transcript1 != None:
                    msg += '\n\n' + transcript1.__repr__()
            else:
                transcript1 = result
                if transcript1.text == '':
                    transcript1.text = 'txt\nThis is an Episode Transcript.'
                transcript1.db_save()
                msg += 'You pressed OK!\n\n' + transcript1.__repr__()
            transcript1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 505 in testsToRun:
            # Transcript Properties Form
            testName = 'EDIT Transcript Properties > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Transcript Properties Form.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press CANCEL.'
            self.ShowMessage(msg)
            transcript1.lock_record()
            dlg = TranscriptPropertiesForm.EditTranscriptDialog(self, -1, transcript1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this indicate you pressed Cancel, having left your Transcript Properties unchanged?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if transcript1 != None:
                    msg += '\n\n' + transcript1.__repr__()
            else:
                transcript1 = result
                if transcript1.text == '':
                    transcript1.text = 'txt\nThis is an Episode Transcript.'
                transcript1.db_save()
                msg += 'You pressed OK!\n\n' + transcript1.__repr__()
            transcript1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 506 in testsToRun:
            # Episode Transcript Loading
            testName = 'Episode Transcript.db_load_by_num()'
            self.SetStatusText(testName)
            transcript2 = Transcript.Transcript(transcript1.number)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(transcript1, transcript2, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Episode Transcript.db_load_by_num():"
                print transcript1
                print transcript2
                print self.Compare(transcript1, transcript2, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(transcript2)

        if 507 in testsToRun:
            # Episode Transcript Loading
            testName = 'Episode Transcript.db_load_by_name()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                transcript2 = Transcript.Transcript(transcript1.id, transcript1.episode_num)
                if self.Compare(transcript1, transcript2, True):
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                    print "Episode Transcript.db_load_by_name():"
                    print transcript1
                    print transcript2
                    print self.Compare(transcript1, transcript2, True)
                    print

            except:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print
                print dir(sys.exc_info()[1])
                print
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                print
                
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].message)
                errordlg.ShowModal()
                errordlg.Destroy()
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(transcript2)
                
        if 508 in testsToRun:
            msg = "You should now see several error messages!"
            dlg = Dialogs.InfoDialog(self, msg, dlgTitle='Unit Test 1: Objects and Forms')
            dlg.ShowModal()
            dlg.Destroy()
            # Episode Transcript Saving
            testName = 'Save Episode Transcript with no ID'
            self.SetStatusText(testName)
            transcript2 = Transcript.Transcript(transcript1.number)
            transcript2.id = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                transcript2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(transcript2)
                
        if 509 in testsToRun:
            # Episode Transcript Saving
            testName = 'Save Episode with Duplicate ID'
            self.SetStatusText(testName)
            transcript2 = Transcript.Transcript(transcript1.number)
            transcript2.number = 0
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                transcript2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(transcript2)
                
        collection1 = None
        if 601 in testsToRun:
            # Collection Properties Form
            testName = 'Add Collection > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Add Collection Form.'
            msg += '\n\n'
            msg += 'Enter distinct values for all fields.'
            self.ShowMessage(msg)
            dlg = CollectionPropertiesForm.AddCollectionDialog(self, -1, 0)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Collection data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
            else:
                collection1 = result
                collection1.db_save()
                msg += 'You pressed OK!\n\n' + collection1.__repr__()
            self.ConfirmTest(msg, testName)

        if 602 in testsToRun:
            # Collection Properties Form
            testName = 'Add Collection > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Add Collection Form.'
            msg += '\n\n'
            msg += 'Press CANCEL.'
            self.ShowMessage(msg)
            dlg = CollectionPropertiesForm.AddCollectionDialog(self, -1, 0)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this say you pressed Cancel?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if collection1 != None:
                    msg += '\n\n' + collection1.__repr__()
            else:
                collection1 = result
                collection1.db_save()
                msg += 'You pressed OK!\n\n' + collection1.__repr__()
            self.ConfirmTest(msg, testName)

        collection2 = None
        if (603 in testsToRun) and (collection1 != None):
            # Nested Collection Properties Form
            testName = 'Add Nested Collection > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Add Nested Collection Form.'
            msg += '\n\n'
            msg += 'Enter distinct values for all fields.'
            self.ShowMessage(msg)
            dlg = CollectionPropertiesForm.AddCollectionDialog(self, -1, collection1.number)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Collection data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
            else:
                collection2 = result
                collection2.db_save()
                msg += 'You pressed OK!\n\n' + collection2.__repr__()
            self.ConfirmTest(msg, testName)

        if (604 in testsToRun) and (collection1 != None):
            # Nested Collection Properties Form
            testName = 'Add Nested Collection > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Add Nested Collection Form.'
            msg += '\n\n'
            msg += 'Press CANCEL.'
            self.ShowMessage(msg)
            dlg = CollectionPropertiesForm.AddCollectionDialog(self, -1, collection1.number)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this say you pressed Cancel?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if collection2 != None:
                    msg += '\n\n' + collection2.__repr__()
            else:
                collection2 = result
                collection2.db_save()
                msg += 'You pressed OK!\n\n' + collection2.__repr__()
            self.ConfirmTest(msg, testName)

        if 605 in testsToRun:
            # Collection Properties Form
            testName = 'Collection Properties Form > Help'
            self.SetStatusText(testName)
            msg = 'You will see an Add Collection Form.'
            msg += '\n\n'
            msg += 'Press HELP, then CANCEL.'
            self.ShowMessage(msg)
            dlg = CollectionPropertiesForm.AddCollectionDialog(self, -1, 0)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Did you see the Collection Properties Help?\n\n'
            self.ConfirmTest(msg, testName)

        if 606 in testsToRun:
            # Collection Properties Form
            testName = 'Edit Collection > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Collection Form.'
            msg += '\n\n'
            msg += 'Edit values for all fields and press OK.'
            self.ShowMessage(msg)
            collection1.lock_record()
            dlg = CollectionPropertiesForm.EditCollectionDialog(self, -1, collection1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Collection data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if collection1 != None:
                    msg += '\n\n' + collection1.__repr__()
            else:
                collection1 = result
                collection1.db_save()
                msg += 'You pressed OK!\n\n' + collection1.__repr__()
            collection1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 607 in testsToRun:
            # Collection Properties Form
            testName = 'EDIT Collection > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Collection Form.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press CANCEL.'
            self.ShowMessage(msg)
            collection1.lock_record()
            dlg = CollectionPropertiesForm.EditCollectionDialog(self, -1, collection1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this indicate you pressed Cancel, having left your Collection unchanged?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if collection1 != None:
                    msg += '\n\n' + collection1.__repr__()
            else:
                collection1 = result
                collection1.db_save()
                msg += 'You pressed OK!\n\n' + collection1.__repr__()
            collection1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 608 in testsToRun:
            # Collection Loading
            testName = 'Collection.db_load_by_num()'
            self.SetStatusText(testName)
            collection3 = Collection.Collection(collection1.number)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(collection1, collection3, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Collection.db_load_by_num():"
                print collection1
                print collection3
                print self.Compare(collection1, collection3, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)
                
        if 609 in testsToRun:
            # Collection Loading
            testName = 'Collection.db_load_by_name()'
            self.SetStatusText(testName)
            collection3 = Collection.Collection(collection1.id, collection1.parent)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(collection1, collection3, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Collection.db_load_by_num():"
                print collection1
                print collection3
                print self.Compare(collection1, collection3, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)
                
        if 610 in testsToRun:
            # Collection Loading
            testName = 'Nested Collection.db_load_by_num()'
            self.SetStatusText(testName)
            collection3 = Collection.Collection(collection2.number)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(collection2, collection3, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Collection.db_load_by_num():"
                print collection2
                print collection3
                print self.Compare(collection2, collection3, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)
                
        if 611 in testsToRun:
            # Collection Loading
            testName = 'Nested Collection.db_load_by_name()'
            self.SetStatusText(testName)
            collection3 = Collection.Collection(collection2.id, collection2.parent)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(collection2, collection3, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Collection.db_load_by_num():"
                print collection2
                print collection3
                print self.Compare(collection2, collection3, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)
                
        if 612 in testsToRun:
            msg = "You should now see several error messages!"
            dlg = Dialogs.InfoDialog(self, msg, dlgTitle='Unit Test 1: Objects and Forms')
            dlg.ShowModal()
            dlg.Destroy()
            # Collection Saving
            testName = 'Save Collection with no ID'
            self.SetStatusText(testName)
            collection3 = Collection.Collection(collection1.number)
            collection3.id = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                collection3.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)
                
        if 613 in testsToRun:
            # Collection Saving
            testName = 'Save Collection with duplicate ID'
            self.SetStatusText(testName)
            collection3 = Collection.Collection()
            collection3.id = collection1.id
            collection3.parent = collection1.parent
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                collection3.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)

        if 614 in testsToRun:
            # Collection Saving
            testName = 'Save Nested Collection with no ID'
            self.SetStatusText(testName)
            collection3 = Collection.Collection(collection2.number)
            collection3.id = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                collection3.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)
                
        if 615 in testsToRun:
            # Collection Saving
            testName = 'Save Nested Collection with duplicate ID'
            self.SetStatusText(testName)
            collection3 = Collection.Collection()
            collection3.id = collection2.id
            collection3.parent = collection2.parent
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                collection3.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)

        if 616 in testsToRun:
            # Collection Saving
            testName = 'Save a second Collection'
            self.SetStatusText(testName)
            collection3 = Collection.Collection(collection1.number)
            collection3.number = 0
            collection3.id = collection3.id + ' - Duplicate'
            collection3.comment = 'This Collection was created automatically'
            collection3.owner = 'Unit Test Form Check'

            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                collection3.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)

        if 617 in testsToRun:
            # Collection Saving
            testName = 'Save a second Nested Collection'
            self.SetStatusText(testName)
            collection3 = Collection.Collection(collection2.number)
            collection3.number = 0
            collection3.id = collection3.id + ' - Duplicate'
            collection3.comment = 'This Collection was created automatically'
            collection3.owner = 'Unit Test Form Check'

            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                collection3.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)

        if 701 in testsToRun:
            # Prepare the Clip Object, which would be passed in
            clip1 = Clip.Clip()
            clip1.episode_num = episode1.number
            clip1.clip_start = 5000
            clip1.clip_stop = 10000
            clip1.collection_num = collection1.number
            clip1.collection_id = collection1.id
            clip1.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
            clip1.offset = 0
            clip1.media_filename = episode1.media_filename
            clip1.audio = 1
            for clipKeyword in episode1.keyword_list:
                clip1.add_keyword(clipKeyword.keywordGroup, clipKeyword.keyword)

            # Prepare the Clip Transcript Object, which would be passed in
            cliptranscript1 = Transcript.Transcript()
            cliptranscript1.episode_num = episode1.number
            cliptranscript1.source_transcript = transcript1.number
            cliptranscript1.clip_start = clip1.clip_start
            cliptranscript1.clip_stop = clip1.clip_stop
            cliptranscript1.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" fontsize="8" fontstyle="90" fontweight="90" fontunderlined="0" fontface="MS Shell Dlg 2" alignment="1" parspacingafter="10" parspacingbefore="0" linespacing="10">
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="443" rightindent="0" parspacingafter="62" parspacingbefore="0" linespacing="10" tabs="317,444">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">This is the Clip Transcript</text>
      <text></text>
    </paragraph>
  </paragraphlayout>
</richtext>"""

            clip1.transcripts.append(cliptranscript1)

            clip2 = clip1.duplicate()

        if 701 in testsToRun:
            # Clip Properties Form
            testName = 'Add Clip > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Add Clip Form.'
            msg += '\n\n'
            msg += 'Enter distinct values for all fields.'
            self.ShowMessage(msg)
            dlg = ClipPropertiesForm.AddClipDialog(self, -1, clip1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Clip data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
            else:
                clip1 = result
                clip1.db_save()
                msg += 'You pressed OK!\n\n' + clip1.__repr__()
            self.ConfirmTest(msg, testName)

        if 702 in testsToRun:
            # Clip Properties Form
            testName = 'Add Clip > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Add Clip Form.'
            msg += '\n\n'
            msg += 'Press CANCEL.'
            self.ShowMessage(msg)
            dlg = ClipPropertiesForm.AddClipDialog(self, -1, clip2)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this say you pressed Cancel?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                msg += '\n\n' + clip1.__repr__()
            else:
                clip2 = result
                clip2.db_save()
                msg += 'You pressed OK!\n\n' + clip2.__repr__()
            self.ConfirmTest(msg, testName)

        if 703 in testsToRun:
            # Clip Properties Form
            testName = 'Clip Properties Form > Help'
            self.SetStatusText(testName)
            msg = 'You will see an Add Clip Form.'
            msg += '\n\n'
            msg += 'Press HELP, then CANCEL.'
            self.ShowMessage(msg)
            dlg = ClipPropertiesForm.AddClipDialog(self, -1, clip2)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Did you see the Clip Properties Help?\n\n'
            self.ConfirmTest(msg, testName)
        del(clip2)

        if 704 in testsToRun:
            # Clip Properties Form
            testName = 'Edit Clip > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Clip Form.'
            msg += '\n\n'
            msg += 'Edit values for all fields and press OK.'
            self.ShowMessage(msg)
            clip1.lock_record()
            dlg = ClipPropertiesForm.EditClipDialog(self, -1, clip1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Clip data?\n\n'
            if result == None:
                clip1 = Clip.Clip(clip1.number)
#                clip1.refresh_keywords()
                msg += 'You pressed Cancel.'
                msg += '\n\n' + clip1.__repr__()
            else:
                clip1 = result
                clip1.db_save()
                msg += 'You pressed OK!\n\n' + clip1.__repr__()
            clip1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 705 in testsToRun:
            # Clip Properties Form
            testName = 'EDIT Clip > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Clip Form.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press CANCEL.'
            self.ShowMessage(msg)
            clip1.lock_record()
            dlg = ClipPropertiesForm.EditClipDialog(self, -1, clip1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this indicate you pressed Cancel, having left your Clip unchanged?\n\n'
            if result == None:
                clip1 = Clip.Clip(clip1.number)
#                clip1.refresh_keywords()
                msg += 'You pressed Cancel.'
                msg += '\n\n' + clip1.__repr__()
            else:
                clip1 = result
                clip1.db_save()
                msg += 'You pressed OK!\n\n' + clip1.__repr__()
            clip1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 706 in testsToRun:
            # Clip Loading
            testName = 'Clip.db_load_by_num()'
            self.SetStatusText(testName)
            clip2 = Clip.Clip(clip1.number)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(clip1, clip2, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Clip.db_load_by_num():"
                print clip1
                print clip2
                print self.Compare(clip1, clip2, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(clip2)

        if 707 in testsToRun:
            # Clip Loading
            testName = 'Clip.db_load_by_name()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                collection3 = Collection.Collection(clip1.collection_num)
                clip2 = Clip.Clip(clip1.id, clip1.collection_id, collection3.parent)
                if self.Compare(clip1, clip2, True):
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED (comparison).')
                    self.testsFailed += 1

                    print "Clip.db_load_by_name():"
                    print clip1
                    print clip2
                    print self.Compare(clip1, clip2, True)
                    print

            except:
                self.txtCtrl.AppendText('FAILED (exception).')
                self.testsFailed += 1
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(collection3)
            del(clip2)
                
        if 708 in testsToRun:
            msg = "You should now see several error messages!"
            dlg = Dialogs.InfoDialog(self, msg, dlgTitle='Unit Test 1: Objects and Forms')
            dlg.ShowModal()
            dlg.Destroy()
            # Clip Saving
            testName = 'Save Clip with no ID'
            self.SetStatusText(testName)
            clip2 = Clip.Clip(clip1.number)
            clip2.id = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                clip2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(clip2)

        if 709 in testsToRun:
            # Clip Saving
            testName = 'Save Clip without Collection'
            self.SetStatusText(testName)
            clip2 = Clip.Clip(clip1.number)
            clip2.collection_num = 0
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                clip2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(clip2)

        if 710 in testsToRun:
            # Clip Saving
            testName = 'Save Clip without media file'
            self.SetStatusText(testName)
            clip2 = Clip.Clip(clip1.number)
            clip2.media_filename = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                clip2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(clip2)

        if 711 in testsToRun:
            # Clip Saving
            testName = 'Save Clip with negative start time'
            self.SetStatusText(testName)
            clip2 = Clip.Clip(clip1.number)
            clip2.clip_start = -1
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                clip2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(clip2)

        if 712 in testsToRun:
            # Clip Saving
            testName = 'Save Clip with Duplicate ID'
            self.SetStatusText(testName)
            clip2 = Clip.Clip(clip1.number)
            clip2.number = 0
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                clip2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(clip2)

        if 713 in testsToRun:
            # Clip Transcript Loading
            testName = 'Clip Transcript.db_load_by_clipnum()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                transcript2 = Transcript.Transcript(clip=clip1.number)
                if self.Compare(clip1.transcripts[0], transcript2, True):
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                    print "Clip Transcript.db_load_by_clipnum():"
                    print clip1.transcripts[0]
                    print transcript2
                    print self.Compare(clip1.transcripts[0], transcript2, True)
                    print

            except:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(transcript2)
                
        snapshot1 = Snapshot.Snapshot()
        snapshot1.collection_num = collection1.number
        snapshot2 = Snapshot.Snapshot()
        snapshot2.collection_num = collection1.number
        if 801 in testsToRun:
            # Snapshot Properties Form
            testName = 'Add Snapshot > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Add Snapshot Form.'
            msg += '\n\n'
            msg += 'Enter distinct values for all fields.'
            self.ShowMessage(msg)
            
            dlg = SnapshotPropertiesForm.AddSnapshotDialog(self, -1, snapshot1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Snapshot data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
            else:
                snapshot1 = result
                snapshot1.db_save()
                msg += 'You pressed OK!\n\n' + snapshot1.__repr__()
            self.ConfirmTest(msg, testName)

        if 802 in testsToRun:
            # Snapshot Properties Form
            testName = 'Add Snapshot > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Add Snapshot Form.'
            msg += '\n\n'
            msg += 'Press CANCEL.'
            self.ShowMessage(msg)
            dlg = SnapshotPropertiesForm.AddSnapshotDialog(self, -1, snapshot2)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this say you pressed Cancel?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                msg += '\n\n' + snapshot1.__repr__()
            else:
                snapshot2 = result
                snapshot2.db_save()
                msg += 'You pressed OK!\n\n' + snapshot2.__repr__()
            self.ConfirmTest(msg, testName)

        if 803 in testsToRun:
            # Snapshot Properties Form
            testName = 'Snapshot Properties Form > Help'
            self.SetStatusText(testName)
            msg = 'You will see an Add Snapshot Form.'
            msg += '\n\n'
            msg += 'Press HELP, then CANCEL.'
            self.ShowMessage(msg)
            dlg = SnapshotPropertiesForm.AddSnapshotDialog(self, -1, snapshot2)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Did you see the Snapshot Properties Help?\n\n'
            self.ConfirmTest(msg, testName)
        del(snapshot2)

        if 804 in testsToRun:
            # Snapshot Properties Form
            testName = 'Edit Snapshot > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Snapshot Form.'
            msg += '\n\n'
            msg += 'Edit values for all fields and press OK.'
            self.ShowMessage(msg)
            snapshot1.lock_record()
            dlg = SnapshotPropertiesForm.EditSnapshotDialog(self, -1, snapshot1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Snapshot data?\n\n'
            if result == None:
                snapshot1 = Snapshot.Snapshot(snapshot1.number)
#                snapshot1.refresh_keywords()
                msg += 'You pressed Cancel.'
                msg += '\n\n' + snapshot1.__repr__()
            else:
                snapshot1 = result
                snapshot1.db_save()
                msg += 'You pressed OK!\n\n' + snapshot1.__repr__()
            snapshot1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 805 in testsToRun:
            # Snapshot Properties Form
            testName = 'EDIT Snapshot > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Snapshot Form.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press CANCEL.'
            self.ShowMessage(msg)
            snapshot1.lock_record()
            dlg = SnapshotPropertiesForm.EditSnapshotDialog(self, -1, snapshot1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this indicate you pressed Cancel, having left your Snapshot unchanged?\n\n'
            if result == None:
                snapshot1 = Snapshot.Snapshot(snapshot1.number)
#                snapshot1.refresh_keywords()
                msg += 'You pressed Cancel.'
                msg += '\n\n' + snapshot1.__repr__()
            else:
                snapshot1 = result
                snapshot1.db_save()
                msg += 'You pressed OK!\n\n' + snapshot1.__repr__()
            snapshot1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 806 in testsToRun:
            # Snapshot Loading
            testName = 'Snapshot.db_load_by_num()'
            self.SetStatusText(testName)
            snapshot2 = Snapshot.Snapshot(snapshot1.number)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(snapshot1, snapshot2, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED (comparison).')
                self.testsFailed += 1

                print "Snapshot.db_load_by_num():"
                print snapshot1
                print snapshot2
                print self.Compare(snapshot1, snapshot2, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(snapshot2)

        if 807 in testsToRun:
            # Snapshot Loading
            testName = 'Snapshot.db_load_by_name()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                snapshot2 = Snapshot.Snapshot(snapshot1.id, snapshot1.collection_num)
                if self.Compare(snapshot1, snapshot2, True):
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED (comparison).')
                    self.testsFailed += 1

                    print "Snapshot.db_load_by_name():"
                    print snapshot1
                    print snapshot2
                    print self.Compare(snapshot1, snapshot2, True)
                    print

            except:
                self.txtCtrl.AppendText('FAILED (exception).')
                self.testsFailed += 1
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            try:
                del(snapshot2)
            except:
                pass

        if 808 in testsToRun:
            msg = "You should now see several error messages!"
            dlg = Dialogs.InfoDialog(self, msg, dlgTitle='Unit Test 1: Objects and Forms')
            dlg.ShowModal()
            dlg.Destroy()
            # Snapshot Saving
            testName = 'Save Snapshot with no ID'
            self.SetStatusText(testName)
            snapshot2 = Snapshot.Snapshot(snapshot1.number)
            snapshot2.id = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                snapshot2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(snapshot2)

        if 809 in testsToRun:
            # Snapshot Saving
            testName = 'Save Snapshot without Collection'
            self.SetStatusText(testName)
            snapshot2 = Snapshot.Snapshot(snapshot1.number)
            snapshot2.collection_num = 0
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                snapshot2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(snapshot2)

        if 810 in testsToRun:
            # Snapshot Saving
            testName = 'Save Snapshot without image file'
            self.SetStatusText(testName)
            snapshot2 = Snapshot.Snapshot(snapshot1.number)
            snapshot2.image_filename = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                snapshot2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(snapshot2)

        if 811 in testsToRun:
            # Snapshot Saving
            testName = 'Save Snapshot with negative episode start'
            self.SetStatusText(testName)
            snapshot2 = Snapshot.Snapshot(snapshot1.number)
            snapshot2.episode_start = -1
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                snapshot2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(snapshot2)

        if 812 in testsToRun:
            # Snapshot Saving
            testName = 'Save Snapshot with Duplicate ID'
            self.SetStatusText(testName)
            snapshot2 = Snapshot.Snapshot(snapshot1.number)
            snapshot2.number = 0
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                snapshot2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(snapshot2)

        note1 = None
        if 901 in testsToRun:
            # Note Properties Form
            testName = 'Add Series Note Properties > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Add Series Note Properties Form.'
            msg += '\n\n'
            msg += 'Enter distinct values for all fields.'
            self.ShowMessage(msg)
            dlg = NotePropertiesForm.AddNoteDialog(self, -1, series1.number, 0, 0, 0, 0, 0)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Note data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
            else:
                note1 = result
                note1.db_save()
                msg += 'You pressed OK!\n\n' + note1.__repr__()
            self.ConfirmTest(msg, testName)

        if 902 in testsToRun:
            # Note Properties Form
            testName = 'Add Series Note Properties > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Add Series Note Properties Form.'
            msg += '\n\n'
            msg += 'Press CANCEL.'
            self.ShowMessage(msg)
            dlg = NotePropertiesForm.AddNoteDialog(self, -1, series1.number, 0, 0, 0, 0, 0)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this say you pressed Cancel?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if note1 != None:
                    msg += '\n\n' + note1.__repr__()
            else:
                note1 = result
                note1.db_save()
                msg += 'You pressed OK!\n\n' + note1.__repr__()
            self.ConfirmTest(msg, testName)

        if 903 in testsToRun:
            # Note Properties Form
            testName = 'Note Properties Form > Help'
            self.SetStatusText(testName)
            msg = 'You will see an Add Note Form.'
            msg += '\n\n'
            msg += 'Press HELP, then CANCEL.'
            self.ShowMessage(msg)
            dlg = NotePropertiesForm.AddNoteDialog(self, -1, series1.number, 0, 0, 0, 0, 0)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Did you see the Note Properties Help?\n\n'
            self.ConfirmTest(msg, testName)

        if 904 in testsToRun:
            # Note Properties Form
            testName = 'Edit Series Note Properties > OK'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Series Note Properties Form.'
            msg += '\n\n'
            msg += 'Edit values for all fields and press OK.'
            self.ShowMessage(msg)
            note1.lock_record()
            dlg = NotePropertiesForm.EditNoteDialog(self, -1, note1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Do you see the correct Note data?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if note11 != None:
                    msg += '\n\n' + note1.__repr__()
            else:
                note1 = result
                note1.db_save()
                msg += 'You pressed OK!\n\n' + note1.__repr__()
            note1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 905 in testsToRun:
            # Note Properties Form
            testName = 'EDIT Series Note Properties > Cancel'
            self.SetStatusText(testName)
            msg = 'You will see an Edit Series Note Properties Form.'
            msg += '\n\n'
            msg += 'Edit the values for all fields, then press CANCEL.'
            self.ShowMessage(msg)
            note1.lock_record()
            dlg = NotePropertiesForm.EditNoteDialog(self, -1, note1)
            result = dlg.get_input()
            dlg.Destroy()
            msg = 'Does this indicate you pressed Cancel, having left your Note unchanged?\n\n'
            if result == None:
                msg += 'You pressed Cancel.'
                if note1 != None:
                    msg += '\n\n' + note1.__repr__()
            else:
                note1 = result
                note1.db_save()
                msg += 'You pressed OK!\n\n' + note1.__repr__()
            note1.unlock_record()
            self.ConfirmTest(msg, testName)

        if 906 in testsToRun:
            # Note Loading
            testName = 'Note.db_load_by_num()'
            self.SetStatusText(testName)
            note2 = Note.Note(note1.number)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(note1, note2, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Note.db_load_by_num():"
                print note1
                print note2
                print self.Compare(note1, note2, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 907 in testsToRun:
            # Note Loading
            testName = 'Note.db_load_by_name()'
            self.SetStatusText(testName)
            note2 = Note.Note(note1.id, Series=note1.series_num)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            if self.Compare(note1, note2, True):
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

                print "Note.db_load_by_name():"
                print note1
                print note2
                print self.Compare(note1, note2, True)
                print

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 908 in testsToRun:
            msg = "You should now see several error messages!"
            dlg = Dialogs.InfoDialog(self, msg, dlgTitle='Unit Test 1: Objects and Forms')
            dlg.ShowModal()
            dlg.Destroy()
            # Note Saving
            testName = 'Save Series Note with no ID'
            self.SetStatusText(testName)
            note2 = Note.Note(note1.number)
            note2.id = ''
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 909 in testsToRun:
            # Note Saving
            testName = 'Save Note not attached to any object'
            self.SetStatusText(testName)
            note2 = Note.Note(note1.number)
            note2.series_num = 0
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 910 in testsToRun:
            # Note Saving
            testName = 'Save Series Note with duplicate ID'
            self.SetStatusText(testName)
            note2 = Note.Note()
            note2.id = note1.id
            note2.series_num = note1.series_num
            note2.text = 'This is a duplicate Series Note'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 911 in testsToRun:
            # Note Saving
            testName = 'Save another Series Note'
            self.SetStatusText(testName)
            note2 = Note.Note()
            if note1.id == 'Series Note':
                note2.id = note1.id + ' - Duplicate'
            else:
                note2.id = 'Series Note'
            note2.series_num = series1.number
            note2.author = 'unit_test_form_check'
            note2.text = 'This is a Series Note generated by unit_test_form_check.'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 912 in testsToRun:
            # Note Saving
            testName = 'Save an Episode Note'
            self.SetStatusText(testName)
            note2 = Note.Note()
            note2.id = 'Episode Note'
            note2.episode_num = episode1.number
            note2.author = 'unit_test_form_check'
            note2.text = 'This is an Episode Note generated by unit_test_form_check.'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 913 in testsToRun:
            # Note Saving
            testName = 'Save a Transcript Note'
            self.SetStatusText(testName)
            note2 = Note.Note()
            note2.id = 'Transcript Note'
            note2.transcript_num = transcript1.number
            note2.author = 'unit_test_form_check'
            note2.text = 'This is a Transcript Note generated by unit_test_form_check.'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 914 in testsToRun:
            # Note Saving
            testName = 'Save a Collection Note'
            self.SetStatusText(testName)
            note2 = Note.Note()
            note2.id = 'Collection Note'
            note2.collection_num = collection1.number
            note2.author = 'unit_test_form_check'
            note2.text = 'This is a Collection Note generated by unit_test_form_check.'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 915 in testsToRun:
            # Note Saving
            testName = 'Save a Collection Note'
            self.SetStatusText(testName)
            note2 = Note.Note()
            note2.id = 'Nested Collection Note'
            note2.collection_num = collection2.number
            note2.author = 'unit_test_form_check'
            note2.text = 'This is a Nested Collection Note generated by unit_test_form_check.'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 916 in testsToRun:
            # Note Saving
            testName = 'Save a Clip Note'
            self.SetStatusText(testName)
            note2 = Note.Note()
            note2.id = 'Clip Note'
            note2.clip_num = clip1.number
            note2.author = 'unit_test_form_check'
            note2.text = 'This is a Clip Note generated by unit_test_form_check.'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        if 917 in testsToRun:
            # Note Saving
            testName = 'Save a Snapshot Note'
            self.SetStatusText(testName)
            note2 = Note.Note()
            note2.id = 'Snapshot Note'
            note2.snapshot_num = snapshot1.number
            note2.author = 'unit_test_form_check'
            note2.text = 'This is a Snapshot Note generated by unit_test_form_check.'
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            try:
                # A SaveError should be raised!
                note2.db_save()
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
            del(note2)

        self.txtCtrl.AppendText('All tests completed.')
        self.txtCtrl.AppendText('\nFinal Summary:  Total Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
        DBInterface.close_db()

    def Compare(self, obj1, obj2, incNumber=False):
        if type(obj1) != type(obj2):
            return False
        result = True
        objType = None
        if isinstance(obj1, Series.Series):
            objType = 'Series'
        elif isinstance(obj1, Episode.Episode):
            objType = 'Episode'
        elif isinstance(obj1, Transcript.Transcript):
            objType = 'Transcript'
        elif isinstance(obj1, Collection.Collection):
            objType = 'Collection'
        elif isinstance(obj1, Clip.Clip):
            objType = 'Clip'
        elif isinstance(obj1, Snapshot.Snapshot):
            objType = 'Snapshot'
        elif isinstance(obj1, KeywordObject.Keyword):
            objType = 'Keyword'
        elif isinstance(obj1, Note.Note):
            objType = 'Note'
            
        if objType in ['Series', 'Episode', 'Transcript', 'Collection', 'Clip', 'Snapshot', 'Note']:
            if incNumber:
                result = result and (obj1.number == obj2.number)
            result = result and (obj1.id == obj2.id)
            result = result and (obj1.comment == obj2.comment)
        if objType in ['Series', 'Collection']:
            result = result and (obj1.owner == obj2.owner)
            result = result and (obj1.keyword_group == obj2.keyword_group)
        if objType in ['Episode']:
            result = result and (obj1.tape_length == obj2.tape_length)
            result = result and (obj1.episode_length() == obj2.episode_length())
            result = result and (obj1.tape_date == obj2.tape_date)
            if len(obj1._kwlist) == len(obj2._kwlist):
                for x in range(len(obj1._kwlist)):
                    result = result and (obj1._kwlist[x].keywordPair == obj2._kwlist[x].keywordPair)
            else:
                result = False
        if objType in ['Episode', 'Clip']:
            result = result and (obj1.media_filename == obj2.media_filename)
            result = result and (obj1.offset == obj2.offset)
            if len(obj1.additional_media_files) == len(obj2.additional_media_files):
                for x in range(len(obj1.additional_media_files)):
                    result = result and (obj1.additional_media_files[x]['filename'] == obj2.additional_media_files[x]['filename'])
                    result = result and (obj1.additional_media_files[x]['offset'] == obj2.additional_media_files[x]['offset'])
                    result = result and (obj1.additional_media_files[x]['length'] == obj2.additional_media_files[x]['length'])
                    result = result and (obj1.additional_media_files[x]['audio'] == obj2.additional_media_files[x]['audio'])
            else:
                result = False
        if objType in ['Episode', 'Snapshot']:
            result = result and (obj1.series_id == obj2.series_id)
        if objType in ['Episode', 'Snapshot', 'Note']:
            result = result and (obj1.series_num == obj2.series_num)
        if objType in ['Transcript']:
            result = result and (obj1.source_transcript == obj2.source_transcript)
            result = result and (obj1.transcriber == obj2.transcriber)
            result = result and (obj1.clip_start == obj2.clip_start)
            result = result and (obj1.clip_stop == obj2.clip_stop)
            result = result and (obj1.minTranscriptWidth == obj2.minTranscriptWidth)
            result = result and (obj1.text == obj2.text)
        if objType in ['Transcript', 'Snapshot']:
            result = result and (obj1.episode_num == obj2.episode_num)
        if objType in ['Transcript', 'Note']:
            result = result and (obj1.clip_num == obj2.clip_num)
        if objType in ['Collection']:
            result = result and (obj1.parent == obj2.parent)
        if objType in ['Clip', 'Note']:
            result = result and (obj1.episode_num == obj2.episode_num)
        if objType in ['Clip']:
            result = result and (obj1.clip_start == obj2.clip_start)
            result = result and (obj1.clip_stop == obj2.clip_stop)
            result = result and (obj1.audio == obj2.audio)
            if len(obj1.transcripts) == len(obj2.transcripts):
                for x in range(len(obj1.transcripts)):
                    result = result and self.Compare(obj1.transcripts[x], obj2.transcripts[x])
            else:
                result = False
        if objType in ['Clip', 'Snapshot', 'Note']:
            result = result and (obj1.collection_num == obj2.collection_num)
        if objType in ['Clip', 'Snapshot']:
            result = result and (obj1.collection_id == obj2.collection_id)
            if len(obj1.keyword_list) == len(obj2.keyword_list):
                for x in range(len(obj1.keyword_list)):
                    result = result and (obj1.keyword_list[x].keywordPair == obj2.keyword_list[x].keywordPair)
            else:
                result = False
        if objType in ['Transcript', 'Clip', 'Snapshot']:
            result = result and (obj1.sort_order == obj2.sort_order)
        if objType in ['Snapshot']:
            result = result and (obj1.image_filename == obj2.image_filename)
            result = result and (obj1.image_scale == obj2.image_scale)
            result = result and (obj1.image_coords == obj2.image_coords)
            result = result and (obj1.image_size == obj2.image_size)
            result = result and (obj1.episode_id == obj2.episode_id)
            if obj1.transcript_num > 0:
                result = result and (obj1.transcript_id == obj2.transcript_id)
            result = result and (obj1.episode_start == obj2.episode_start)
            result = result and (obj1.episode_duration == obj2.episode_duration)
            if len(obj1.codingObjects) == len(obj2.codingObjects):
                for x in range(len(obj1.codingObjects)):
                    result = result and (obj1.codingObjects[x]['x1'] == obj2.codingObjects[x]['x1'])
                    result = result and (obj1.codingObjects[x]['y1'] == obj2.codingObjects[x]['y1'])
                    result = result and (obj1.codingObjects[x]['x2'] == obj2.codingObjects[x]['x2'])
                    result = result and (obj1.codingObjects[x]['y2'] == obj2.codingObjects[x]['y2'])
                    result = result and (obj1.codingObjects[x]['keywordGroup'] == obj2.codingObjects[x]['keywordGroup'])
                    result = result and (obj1.codingObjects[x]['keyword'] == obj2.codingObjects[x]['keyword'])
                    result = result and (obj1.codingObjects[x]['visible'] == obj2.codingObjects[x]['visible'])
            else:
                result = False
            if len(obj1.keywordStyles) == len(obj2.keywordStyles):
                for x in obj1.keywordStyles.keys():
                    result = result and (obj1.keywordStyles[x]['drawMode'] == obj2.keywordStyles[x]['drawMode'])
                    result = result and (obj1.keywordStyles[x]['lineColorName'] == obj2.keywordStyles[x]['lineColorName'])
                    result = result and (obj1.keywordStyles[x]['lineColorDef'] == obj2.keywordStyles[x]['lineColorDef'])
                    result = result and (obj1.keywordStyles[x]['lineWidth'] == obj2.keywordStyles[x]['lineWidth'])
                    result = result and (obj1.keywordStyles[x]['lineStyle'] == obj2.keywordStyles[x]['lineStyle'])
            else:
                result = False
        if objType in ['Snapshot', 'Note']:
            result = result and (obj1.transcript_num == obj2.transcript_num)
        if objType in ['Keyword']:
            result = result and (obj1.keywordGroup == obj2.keywordGroup)
            result = result and (obj1.originalKeywordGroup == obj2.originalKeywordGroup)
            result = result and (obj1.keyword == obj2.keyword)
            result = result and (obj1.originalKeyword == obj2.originalKeyword)
            result = result and (obj1.definition == obj2.definition)
            result = result and (obj1.lineColorName == obj2.lineColorName)
            result = result and (obj1.lineColorDef == obj2.lineColorDef)
            result = result and (obj1.drawMode == obj2.drawMode)
            result = result and (obj1.lineWidth == obj2.lineWidth)
            result = result and (obj1.lineStyle == obj2.lineStyle)
        if objType in ['Note']:
            result = result and (obj1.notetype == obj2.notetype)
            result = result and (obj1.snapshot_num == obj2.snapshot_num)
            result = result and (obj1.author == obj2.author)
            result = result and (obj1.text == obj2.text)
        return result

    def ShowMessage(self, msg):
        dlg = Dialogs.InfoDialog(self, msg)
        dlg.SetBackgroundColour(wx.Colour(166, 253, 251))
        dlg.ShowModal()
        dlg.Destroy()

    def ConfirmTest(self, msg, testName):
        dlg = Dialogs.QuestionDialog(self, msg, "Unit Test 1: Objects and Forms", noDefault=True)
        dlg.SetBackgroundColour(wx.Colour(166, 253, 168))
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
      frame = FormCheck(None, -1, "Unit Test 1: Objects and Forms")
      self.SetTopWindow(frame)
      return True
      

app = MyApp(0)
app.MainLoop()
