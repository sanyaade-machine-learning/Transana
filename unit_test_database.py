# -*- coding: cp1252 -*-
import datetime
import os
import sys
import wx
# import traceback

import gettext

# This module expects i18n.  Enable it here.
__builtins__._ = wx.GetTranslation


import TransanaConstants

SINGLE_USER = TransanaConstants.singleUserVersion

import TransanaGlobal
import Clip
import Collection
import CoreData
import ConfigData
import ControlObjectClass
import DBInterface
import Dialogs
import Document
import Episode
import KeywordObject
import MenuWindow
import Misc
import Note
import Quote
import Library
import Snapshot
import TransanaExceptions
import Transcript

#if TransanaConstants.DBInstalled in ['sqlite3']:
import sqlite3


MENU_FILE_EXIT = 101

class DBTest(wx.Frame):
    """ This window displays a variety of GUI Widgets. """
    def __init__(self,parent,id,title):

        wx.SetDefaultPyEncoding('utf8')

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
        if 'wxMSW' in wx.PlatformInfo:
            TransanaGlobal.configData.videoPath = 'C:\\Users\\DavidWoods\\Videos\\'
        elif 'wxMac' in wx.PlatformInfo:
            TransanaGlobal.configData.videoPath = '/Volumes/Vidëos/'
        TransanaGlobal.configData.ssl = False

        TransanaGlobal.menuWindow = MenuWindow.MenuWindow(None, -1, title)

        self.ControlObject = ControlObjectClass.ControlObject()
        self.ControlObject.Register(Menu = TransanaGlobal.menuWindow)
        TransanaGlobal.menuWindow.Register(ControlObject = self.ControlObject)

        if TransanaConstants.DBInstalled in ['MySQLdb-embedded']:
            DBInterface.InitializeSingleUserDatabase()
            password = ''
        elif TransanaConstants.DBInstalled in ['MySQLdb-server', 'PyMySQL']:
            dlg = wx.PasswordEntryDialog(None, 'Please enter your database password.', 'Unit Test 1:  Database Connection')
            result = dlg.ShowModal()
            if result == wx.ID_OK:
                password = dlg.GetValue()
            else:
                password = ''
            dlg.Destroy()
        elif TransanaConstants.DBInstalled in ['sqlite3']:
            password = ''

        loginInfo = DBLogin('DavidW', password, 'DKW-Linux', 'Transana_UnitTest', '3306')
        dbReference = DBInterface.get_db(loginInfo)
        DBInterface.establish_db_exists()

        wx.Frame.__init__(self,parent,-1, title, size = (800,600), style=wx.DEFAULT_FRAME_STYLE|wx.NO_FULL_REPAINT_ON_RESIZE)
        self.testsRun = 0
        self.testsSuccessful = 0
        self.testsFailed = 0

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        self.txtCtrl = wx.TextCtrl(self, -1, "Unit Test:  Database\n\n", style=wx.TE_LEFT | wx.TE_MULTILINE)
        self.txtCtrl.AppendText("Transana Version:  %s     DBInstalled:  %s     singleUserVersion:  %s\n\n" % (TransanaConstants.versionNumber, TransanaConstants.DBInstalled, TransanaConstants.singleUserVersion))

        mainSizer.Add(self.txtCtrl, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()
        self.CenterOnScreen()
        # Status Bar
        self.CreateStatusBar()
        self.SetStatusText("")
        self.Show(True)
        try:
            self.RunTests()
        except:
            print
            print
            print sys.exc_info()[0]
            print sys.exc_info()[1]
            import traceback
            traceback.print_exc(file=sys.stdout)
            

        if TransanaConstants.DBInstalled in ['MySQLdb-embedded']:
            DBInterface.EndSingleUserDatabase()
            
        TransanaGlobal.menuWindow.Destroy()

    def RunTests(self):
        # Tests defined:
        testsNotToSkip = []
        startAtTest = 1  # Should start at 1, not 0!
        endAtTest = 1000   # Should be one more than the last test to be run!
        testsToRun = testsNotToSkip + range(startAtTest, endAtTest)

        numTestQuotes = 10

        t = datetime.datetime.now()

        db = DBInterface.get_db()
        dbCursor = db.cursor()

        if 10 in testsToRun:
            # Is Database Open?
            testName = 'Database Open'
            self.SetStatusText(testName)
            result = DBInterface.is_db_open()
            self.testsRun += 1

            if result:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 20 in testsToRun:
            testName = 'Delete All Tables'
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            self.testsRun += 1

            if TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
                query = "DROP TABLE ConfigInfo, Series2, Episodes2, AdditionalVids2, Documents2, Transcripts2, Collections2, "
                query += "Clips2, Snapshots2, Quotes2, QuotePositions2, Notes2, Keywords2, ClipKeywords2, SnapshotKeywords2, "
                query += "SnapshotKeywordStyles2, CoreData2, Filters2"

                dbCursor.execute(query)
            elif TransanaConstants.DBInstalled in ['sqlite3']:
                tableNames = ["ConfigInfo", "Series2", "Episodes2", "AdditionalVids2", "Documents2", "Transcripts2",
                              "Collections2", "Clips2", "Snapshots2", "Quotes2", "QuotePositions2", "Notes2", "Keywords2", 
                              "ClipKeywords2", "SnapshotKeywords2", "SnapshotKeywordStyles2", "CoreData2", "Filters2"]
                query = "DROP TABLE IF EXISTS %s"
                for table in tableNames:
                    dbCursor.execute(query % table)


            self.testsSuccessful += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))


        if (30 in testsToRun):
            testName = 'DBInterface.establish_db_exists to create new tables'
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            
            DBInterface.establish_db_exists()

            if TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
                query = "SHOW TABLES"
            elif TransanaConstants.DBInstalled in ['sqlite3']:
                query = "SELECT * FROM sqlite_master WHERE type='table'"
            dbCursor.execute(query)
            if TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
                n = dbCursor.rowcount
                results = dbCursor.fetchall()
            elif TransanaConstants.DBInstalled in ['sqlite3']:
                tableResults = dbCursor.fetchall()
                n = len(tableResults) - 1
                results = []
                for x in range(1, n + 1):
                    results.append((tableResults[x][1],))

            self.testsRun += 1
            if n == 18:
                self.txtCtrl.AppendText('Passed.\n')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.\n')
                self.testsFailed += 1

            self.CheckTable('ConfigInfo', results)                    #
            self.CheckTable('Series2', results)                       #
            self.CheckTable('Episodes2', results)                     #
            self.CheckTable('AdditionalVids2', results)
            self.CheckTable('Documents2', results)                    #
            self.CheckTable('Transcripts2', results)                  #
            self.CheckTable('Collections2', results)                  #
            self.CheckTable('Clips2', results)                        #
            self.CheckTable('Snapshots2', results)                    #
            self.CheckTable('Quotes2', results)                       #
            self.CheckTable('QuotePositions2', results)               #
            self.CheckTable('Notes2', results)
            self.CheckTable('Keywords2', results)                     #
            self.CheckTable('ClipKeywords2', results)                 #
            self.CheckTable('SnapshotKeywords2', results)             #
            self.CheckTable('SnapshotKeywordStyles2', results)        #
            self.CheckTable('CoreData2', results)
            self.CheckTable('Filters2', results)

            self.txtCtrl.AppendText('Total Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testKeywordPairs = [(u'English Keyword Group 1',
                             u'English Keyword 1',
                             u'English Definition'),
                            (u"Apostrophe's Test Keyword Group 1",
                             u"Apostrophe's Test Keyword 1",
                             u"Apostrophe's Test Definition"),
                            (unicode('Fran\xc3\xa7ais Cat\xc3\xa9gorie 1', 'utf8'),
                             unicode('Fran\xc3\xa7ais Mot-cl\xc3\xa9 1', 'utf8'),
                             unicode('Fran\xc3\xa7ais D\xc3\xa9finition', 'utf8')),
                            (unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe5\x85\xb3\xe9\x94\xae\xe8\xaf\x8d\xe7\xbb\x84 1', 'utf8'),
                             unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe5\x85\xb3\xe9\x94\xae\xe8\xaf\x8d 1', 'utf8'),
                             unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93  \xe5\xae\x9a\xe4\xb9\x89', 'utf8')),
                            (unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x81\xd8\xa6\xd8\xa9 1', 'utf8'),
                             unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x83\xd9\x84\xd9\x85\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x81\xd8\xaa\xd8\xa7\xd8\xad\xd9\x8a\xd8\xa9 1', 'utf8'),
                             unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xaa\xd8\xb9\xd8\xb1\xd9\x8a\xd9\x81', 'utf8')),
                            (u'Testing Keyword Group 1',
                             u'Testing Keyword 1',
                             u'Testing Definition')]

        for (keywordGroupName, keywordName, definition) in testKeywordPairs:
##            if 40 in testsToRun:
##                try:
##
##                    DBInterface.delete_keyword(keywordGroupName, keywordName)
##                    
##                # If the keyword doesn't exist ...
##                except TransanaExceptions.RecordNotFoundError:
##                    # this is the expected result.
##                    pass
##                except sqlite3.OperationalError:
##                    # This is the desired result!
##                    pass
##                except:
##
##                    print "Exception in Test 40 -- Could not delete Keyword %s : %s." % (keywordGroupName.encode('utf8'), keywordName.encode('utf8'))
##                    print sys.exc_info()[0]
##                    print sys.exc_info()[1]
##                    traceback.print_exc(file=sys.stdout)
##                    print
##
            if 50 in testsToRun:
                # Keyword Saving
                testName = 'Creating Keywords : %s : %s' % (keywordGroupName, keywordName)
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a Keyword
                    keyword1 = KeywordObject.Keyword()
                    keyword1.keywordGroup = keywordGroupName
                    keyword1.keyword = keywordName
                    keyword1.definition = definition
                    keyword1.lineColorName = 'Light Blue'
                    keyword1.lineColorDef = '#0080ff'
                    keyword1.drawMode = 'Rectangle'
                    keyword1.lineWidth = 3
                    keyword1.lineStyle = 'DotDash'
                    try:
                        keyword1.db_save()
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
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    # If we can't load or create the keyword, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 60 in testsToRun:
                # Compare Created with Saved Keywords
                testName = 'Comparing Keywords'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                keyword2 = KeywordObject.Keyword(keywordGroupName, keywordName)

                # A created keyword is missing the "Original" keywordgroup and keyword values.
                # Add them here, or the comparison will definitely fail!
                keyword1.originalKeywordGroup = keyword2.originalKeywordGroup
                keyword1.originalKeyword = keyword2.originalKeyword

                if keyword1 == keyword2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                    self.txtCtrl.AppendText("\n%s\n\n%s\n\n" % (keyword1, keyword2))
                    
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testClipNames = [('English',
                          unicode('English Clip 1', 'utf8'),
                          unicode('English Collection 1', 'utf8')),
                         ('French',
                          unicode('Fran\xc3\xa7ais Extrait 1', 'utf8'),
                          unicode('Fran\xc3\xa7ais Collection 1', 'utf8')),
                         ('Chinese',
                          unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe5\x89\xaa\xe8\xbe\x91 1', 'utf8'),
                          unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe9\x9b\x86\xe5\x90\x88 1', 'utf8')),
                         ('Arabic',
                          unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x82\xd8\xb7\xd8\xb9 1', 'utf8'),
                          unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xac\xd9\x85\xd9\x88\xd8\xb9\xd8\xa9 1', 'utf8')),
                         ("Apostrophe's Test",
                          "Apostrophe's Test Clip 1",
                          "Apostrophe's Test Collection 1"),
                         ('SearchDemo',
                          'A', 'SearchDemo'),
                         ('SearchDemo',
                          'B', 'SearchDemo'),
                         ('SearchDemo',
                          'C', 'SearchDemo'),
                         ('SearchDemo',
                          'AB', 'SearchDemo'),
                         ('SearchDemo',
                          'AC', 'SearchDemo'),
                         ('SearchDemo',
                          'BC', 'SearchDemo'),
                         ('SearchDemo',
                          'ABC', 'SearchDemo')]

        for (language, clipName, collectionName) in testClipNames:
            if 65 in testsToRun:
                self.testsRun += 1
                try:
                    # Try to load the Test Clip
                    testClip1 = Clip.Clip(clipName, collectionName, 0)
                    # Delete the Test Clip
                    testClip1.db_delete()
                    self.testsSuccessful += 1
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    self.testsSuccessful += 1
                except:

                    print "Exception in Test 65 -- Could not delete Clip %s." % clipName.encode('utf8')
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    print
                    
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testQuoteNames = [('English',
                           unicode('English Quote 1', 'utf8'),
                           unicode('English Collection 1', 'utf8')),
                          ('French',
                           unicode('Fran\xc3\xa7ais Quote 1', 'utf8'),
                           unicode('Fran\xc3\xa7ais Collection 1', 'utf8')),
                          ('Chinese',
                           unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Quote 1', 'utf8'),
                           unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe9\x9b\x86\xe5\x90\x88 1', 'utf8')),
                          ('Arabic',
                           unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Quote 1', 'utf8'),
                           unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xac\xd9\x85\xd9\x88\xd8\xb9\xd8\xa9 1', 'utf8')),
                          ("Apostrophe's Test",
                           "Apostrophe's Test Quote 1",
                           "Apostrophe's Test Collection 1")]
 
##        for (language, quoteName, collectionName) in testQuoteNames:
##            if 65 in testsToRun:
##                self.testsRun += 1
##                try:
##                    # Try to load the Test Quote
##                    testQuote1 = Quote.Quote(quoteID=quoteName, collectionID=collectionName, collectionParent=0)
##                    # Delete the Test Quote
##                    testQuote1.db_delete()
##                    self.testsSuccessful += 1
##                except TransanaExceptions.RecordNotFoundError:
##                    # This is the desired result!
##                    self.testsSuccessful += 1
##                except:
##
##                    print "Exception in Test 65 -- Could not delete Quote %s." % quoteName.encode('utf8')
##                    print sys.exc_info()[0]
##                    print sys.exc_info()[1]
##                    traceback.print_exc(file=sys.stdout)
##                    print
##                    
##                    self.txtCtrl.AppendText('FAILED.')
##                    self.testsFailed += 1
##                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
##
##        # Delete Testing Collection
##        testName = 'Deleting Collection : %s' % 'Testing'
##        self.SetStatusText(testName)
##        self.testsRun += 1
##        self.txtCtrl.AppendText('Test "%s" ' % testName)
##        try:
##            # Try to load the Test Collection
##            testCollection1 = Collection.Collection('Testing')
##            # Delete the Test Collection
##            testCollection1.db_delete()
##            self.testsSuccessful += 1
##        except TransanaExceptions.RecordNotFoundError:
##            # This is the desired result!
##            self.testsSuccessful += 1
##        except:
##
##            print "Exception in Test ??? -- Could not delete Collection %s." % 'Testing'
##            print sys.exc_info()[0]
##            print sys.exc_info()[1]
##            print
##            self.txtCtrl.AppendText('FAILED.')
##            self.testsFailed += 1
##        self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
##
##        # Delete SearchDemo Collection
##        testName = 'Deleting Collection : %s' % 'SearchDemo'
##        self.SetStatusText(testName)
##        self.testsRun += 1
##        self.txtCtrl.AppendText('Test "%s" ' % testName)
##        try:
##            # Try to load the Test Collection
##            testCollection1 = Collection.Collection('SearchDemo')
##            # Delete the Test Collection
##            testCollection1.db_delete()
##            self.testsSuccessful += 1
##        except TransanaExceptions.RecordNotFoundError:
##            # This is the desired result!
##            self.testsSuccessful += 1
##        except:
##
##            print "Exception in Test 135 -- Could not delete Collection %s." % 'SearchDemo'
##            print sys.exc_info()[0]
##            print sys.exc_info()[1]
##            print
##            self.txtCtrl.AppendText('FAILED.')
##            self.testsFailed += 1
##        self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
##
##        # Delete Testing Library
##        testName = 'Deleting Library : %s' % 'Testing'
##        self.SetStatusText(testName)
##        self.testsRun += 1
##        self.txtCtrl.AppendText('Test "%s" ' % testName)
##        try:
##            # Try to load the Test Library
##            testSeries1 = Library.Library('Testing')
##            # Delete the Test Library
##            testSeries1.db_delete()
##            self.testsSuccessful += 1
##        except TransanaExceptions.RecordNotFoundError:
##            # This is the desired result!
##            self.testsSuccessful += 1
##        except sqlite3.OperationalError:
##            # This is the desired result!
##            self.testsSuccessful += 1
##        except:
##
##            print "Exception in Test 70 -- Could not delete Library %s." % 'Testing'
##            print sys.exc_info()[0]
##            print sys.exc_info()[1]
##            traceback.print_exc(file=sys.stdout)
##            print
##            self.txtCtrl.AppendText('FAILED.')
##            self.testsFailed += 1
##        self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
##
##        # Delete SearchDemo Library
##        testName = 'Deleting Library : %s' % 'SearchDemo'
##        self.SetStatusText(testName)
##        self.testsRun += 1
##        self.txtCtrl.AppendText('Test "%s" ' % testName)
##        try:
##            # Try to load the Test Library
##            testSeries1 = Library.Library('SearchDemo')
##            # Delete the Test Library
##            testSeries1.db_delete()
##            self.testsSuccessful += 1
##        except TransanaExceptions.RecordNotFoundError:
##            # This is the desired result!
##            self.testsSuccessful += 1
##        except sqlite3.OperationalError:
##            # This is the desired result!
##            self.testsSuccessful += 1
##        except:
##
##            print "Exception in Test 70 -- Could not delete Library %s." % 'SearchDemo'
##            print sys.exc_info()[0]
##            print sys.exc_info()[1]
##            traceback.print_exc(file=sys.stdout)
##            print
##            self.txtCtrl.AppendText('FAILED.')
##            self.testsFailed += 1
##        self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
##
        testSeriesNames = [('English',
                            unicode('English Library 1', 'utf8'),
                            unicode('English Library Owner 1', 'utf8'),
                            unicode('English Library Comment 1', 'utf8'),
                            testKeywordPairs[0][0]),
                            ("Apostrophe's Test",
                             unicode("Apostrophe's Test Library 1", 'utf8'),
                             unicode("Apostrophe's Test Library Owner 1", 'utf8'),
                             unicode("Apostrophe's Test Library Comment 1", 'utf8'),
                            testKeywordPairs[1][0]),
                            ('French',
                             unicode('Fran\xc3\xa7ais S\xc3\xa9ries 1', 'utf8'),
                             unicode('Fran\xc3\xa7ais S\xc3\xa9ries D\xc3\xa9tenteur 1', 'utf8'),
                             unicode('Fran\xc3\xa7ais S\xc3\xa9ries Commentaire 1', 'utf8'),
                            testKeywordPairs[2][0]),
                            ('Chinese',
                             unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe5\xba\x8f\xe5\x88\x97 1', 'utf8'),
                             unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe5\xba\x8f\xe5\x88\x97 \xe6\x89\x80\xe6\x9c\x89\xe8\x80\x85 1', 'utf8'),
                             unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe5\xba\x8f\xe5\x88\x97 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                            testKeywordPairs[3][0]),
                            ('Arabic',
                             unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xb3\xd9\x84\xd8\xb3\xd9\x84\xd8\xa7\xd8\xaa 1', 'utf8'),
                             unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xb3\xd9\x84\xd8\xb3\xd9\x84\xd8\xa7\xd8\xaa \xd8\xa7\xd9\x84\xd9\x85\xd8\xa4\xd9\x84\xd9\x81 1', 'utf8'),
                             unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xb3\xd9\x84\xd8\xb3\xd9\x84\xd8\xa7\xd8\xaa \xd8\xaa\xd8\xb9\xd9\x84\xd9\x8a\xd9\x82 1', 'utf8'),
                            testKeywordPairs[4][0]),
                           ('Testing',
                            unicode('Testing', 'utf8'),
                            unicode('Testing Library Owner 1', 'utf8'),
                            unicode('Testing Library Comment 1', 'utf8'),
                            testKeywordPairs[5][0])]
        testSeriesData = {}

        for (language, seriesName, seriesOwner, seriesComment, seriesKeywordGroup) in testSeriesNames:
            if 70 in testsToRun:
                self.testsRun += 1
                try:
                    # Try to load the Test Library
                    testSeries1 = Library.Library(seriesName)
                    # Delete the Test Library
                    testSeries1.db_delete()
                    self.testsSuccessful += 1
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    self.testsSuccessful += 1
#                except sqlite3.OperationalError:
#                    # This is the desired result!
#                    self.testsSuccessful += 1
                except:

                    print "Exception in Test 70 -- Could not delete Library %s." % seriesName.encode('utf8')
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    print
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 80 in testsToRun:
                # Library Saving
                testName = 'Creating Library : %s' % seriesName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Library
                    series1 = Library.Library()
                    # Populate it with data
                    series1.id = seriesName
                    series1.owner = seriesOwner
                    series1.comment = seriesComment
                    series1.keyword_group = seriesKeywordGroup
                    try:
                        series1.db_save()
                        # If we create the Library, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testSeriesData[language] = (series1.number, series1.id)
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
                    traceback.print_exc(file=sys.stdout)
                    
                    # If we can't load or create the Library, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 90 in testsToRun:
                # Compare Created with Saved Library
                testName = 'Comparing Library by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                series2 = Library.Library(series1.number)

                if series1 == series2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 95 in testsToRun:
                # Compare Created with Saved Library
                testName = 'Comparing Library by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                series2 = Library.Library(seriesName)

                if series1 == series2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
                
        if sys.platform == 'win32':
            mediaFileEnglish    = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Demo.mpg', 'utf8')
            mediaFileApostrophe = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Demo2.mpg', 'utf8')
            mediaFileFrench     = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Test \xc3\xa2\xc3\xab\xc3\xac\xc3\xb3\xc3\xbb 5.mpg', 'utf8')
            mediaFileChinese    = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\\xe4\xba\xb2\xe4\xba\xb3 \xe4\xba\xb2\\\xe4\xba\xb2\xe4\xba\xb3\xe4\xba\xb2 5.mpg', 'utf8')
            mediaFileArabic     = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4\\\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4.mpg', 'utf8')
        
        elif sys.platform == 'darwin':
            mediaFileEnglish    = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\x96d\xc3\xaa/Demo.mpg', 'utf8')
            mediaFileApostrophe = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\x96d\xc3\xaa/Demo2.mpg', 'utf8')
            mediaFileFrench     = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\xb6d\xc3\xaa/Test \xc3\xa2\xc3\xab\xc3\xac\xc3\xb3\xc3\xbb 5.mpg', 'utf8')
            mediaFileChinese    = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\xb6d\xc3\xaa/\xe4\xba\xb2\xe4\xba\xb3 \xe4\xba\xb2/\xe4\xba\xb2\xe4\xba\xb3\xe4\xba\xb2 5.mpg', 'utf8')
            mediaFileArabic     = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\xb6d\xc3\xaa/\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4/\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4.mpg', 'utf8')

        testEpisodeNames = [('English',
                             unicode('English Episode 1', 'utf8'),
                             unicode('English Episode Comment 1', 'utf8'),
                             mediaFileEnglish,
                             testSeriesData['English'][0],
                             testSeriesData['English'][1],
                             testKeywordPairs[0][0],
                             testKeywordPairs[0][1]),
                            ("Apostrophe's Test",
                             unicode("Apostrophe's Test Episode 1", 'utf8'),
                             unicode("Apostrophe's Test Episode Comment 1", 'utf8'),
                             mediaFileApostrophe,
                             testSeriesData["Apostrophe's Test"][0],
                             testSeriesData["Apostrophe's Test"][1],
                             testKeywordPairs[1][0],
                             testKeywordPairs[1][1]),
                            ('French',
                             unicode('Fran\xc3\xa7ais \xc3\x89pisode 1', 'utf8'),
                             unicode('Fran\xc3\xa7ais \xc3\x89pisode Commentaire 1', 'utf8'),
                             mediaFileFrench,
                             testSeriesData['French'][0],
                             testSeriesData['French'][1],
                             testKeywordPairs[2][0],
                             testKeywordPairs[2][1]),
                            ('Chinese',
                             unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe4\xba\x8b\xe4\xbb\xb6 1', 'utf8'),
                             unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe4\xba\x8b\xe4\xbb\xb6 \xe4\xba\x8b\xe4\xbb\xb6 1', 'utf8'),
                             mediaFileChinese,
                             testSeriesData['Chinese'][0],
                             testSeriesData['Chinese'][1],
                             testKeywordPairs[3][0],
                             testKeywordPairs[3][1]),
                            ('Arabic',
                             unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xad\xd9\x84\xd9\x82\xd8\xa9 1', 'utf8'),
                             unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xad\xd9\x84\xd9\x82\xd8\xa9 \xd8\xaa\xd8\xb9\xd9\x84\xd9\x8a\xd9\x82 1', 'utf8'),
                             mediaFileArabic,
                             testSeriesData['Arabic'][0],
                             testSeriesData['Arabic'][1],
                             testKeywordPairs[4][0],
                             testKeywordPairs[4][1])]

        testEpisodeData = {}

        for (language, episodeName, episodeComment, episodeMediaFile, seriesNum, seriesName, keywordGroup, keyword) \
            in testEpisodeNames:

            if 100 in testsToRun:
                # Episode Saving
                testName = 'Creating Episode : %s' % episodeName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    episode1 = Episode.Episode()
                    episode1.id = episodeName
                    episode1.comment = episodeComment
                    episode1.media_filename = episodeMediaFile
                    episode1.tape_length = 3600000L
                    episode1.tape_date = datetime.date(2014, 6, 9)
                    episode1.series_id = seriesName
                    episode1.series_num = seriesNum
                    episode1.add_keyword(keywordGroup, keyword)
                    try:
                        episode1.db_save()
                        # If we create the Episode, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testEpisodeData[language] = (episode1.number, episode1.id)
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

            if 110 in testsToRun:
                # Compare Created with Saved Episode
                testName = 'Comparing Episode by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                episode2 = Episode.Episode(num=episode1.number)

                # A created Episode is missing the "Keyword Episode Number" from keywords!
                # Add them here, or the comparison will definitely fail!
                for obj in episode1._kwlist:
                    obj.episodeNum = episode1.number
                episode1.useVideoRoot = episode2.useVideoRoot

                if episode1 == episode2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 115 in testsToRun:
                # Compare Created with Saved Episode
                testName = 'Comparing Episode by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                episode2 = Episode.Episode(series=seriesName, episode=episodeName)

                # A created Episode is missing the "Keyword Episode Number" from keywords!
                # Add them here, or the comparison will definitely fail!
                for obj in episode1._kwlist:
                    obj.episodeNum = episode1.number
                episode1.useVideoRoot = episode2.useVideoRoot

                if episode1 == episode2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testTranscriptNames = [('English',
                                unicode('English Transcript 1', 'utf8'),
                                unicode('English Transcript Transcriber 1', 'utf8'),
                                unicode('English Transcript Comment 1', 'utf8'),
                                testEpisodeData['English'][0],
                                testEpisodeData['English'][1],
                                testSeriesData['English'][1]),
                               ("Apostrophe's Test",
                                unicode("Apostrophe's Test Transcript 1", 'utf8'),
                                unicode("Apostrophe's Test Transcript Transcriber 1", 'utf8'),
                                unicode("Apostrophe's Test Transcript Comment 1", 'utf8'),
                                testEpisodeData["Apostrophe's Test"][0],
                                testEpisodeData["Apostrophe's Test"][1],
                                testSeriesData["Apostrophe's Test"][1]),
                               ('French',
                                unicode('Fran\xc3\xa7ais Transcription 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Transcription Transcripteur 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Transcription Commentaire 1', 'utf8'),
                                testEpisodeData['French'][0],
                                testEpisodeData['French'][1],
                                testSeriesData['French'][1]),
                               ('Chinese',
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xbd\xac\xe5\xbd\x95 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xbd\xac\xe5\xbd\x95 \xe8\xbd\xac\xe5\xbd\x95\xe4\xba\xba 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xbd\xac\xe5\xbd\x95 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                                testEpisodeData['Chinese'][0],
                                testEpisodeData['Chinese'][1],
                                testSeriesData['Chinese'][1]),
                               ('Arabic',
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xaa\xd8\xaf\xd9\x88\xd9\x8a\xd9\x86\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xaa\xd8\xaf\xd9\x88\xd9\x8a\xd9\x86\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xaf\xd9\x88\xd9\x86 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xaa\xd8\xaf\xd9\x88\xd9\x8a\xd9\x86\xd8\xa9 \xd8\xaa\xd8\xb9\xd9\x84\xd9\x8a\xd9\x82 1', 'utf8'),
                                testEpisodeData['Arabic'][0],
                                testEpisodeData['Arabic'][1],
                                testSeriesData['Arabic'][1])]

        testTranscriptData = {}

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
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">5 English Test</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;10000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">10 Apostrophe's Test</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;15000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">15 Fran\xc3\xa7ais</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;20000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">20 \xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#FF0000" bgcolor="#FFFFFF" fontpointsize="14" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">&#164;</text>
      <text textcolor="#FFFFFF" bgcolor="#FFFFFF" fontpointsize="1" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Times New Roman">"&lt;25000&gt; "</text>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">25 \xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9</text>
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
        for (language, transcriptName, transcriptTranscriber, transcriptComment, episodeNum, episodeName, seriesName) \
            in testTranscriptNames:

            if 120 in testsToRun:
                # Episode Transcript
                testName = 'Creating Transcript : %s' % transcriptName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    transcript1 = Transcript.Transcript()
                    transcript1.id = transcriptName
                    transcript1.transcriber = transcriptTranscriber
                    transcript1.comment = transcriptComment
                    transcript1.episode_num = episodeNum
                    transcript1.text = tempText
                    try:
                        transcript1.db_save()
                        # If we create the Episode, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testTranscriptData[language] = (transcript1.number, transcript1.id)

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

            if 130 in testsToRun:
                # Compare Created with Saved Transcript
                testName = 'Comparing Transcript by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                transcript2 = Transcript.Transcript(transcript1.number)

                # A created Transcript is missing a few key values!
                # Add them here, or the comparison will definitely fail!
#                transcript1.series_id = seriesName
#                transcript1.episode_id = episodeName
                transcript1.changed = False

                if transcript1 == transcript2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 135 in testsToRun:
                # Compare Created with Saved Transcript
                testName = 'Comparing Transcript by name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                transcript2 = Transcript.Transcript(transcriptName, episodeNum)

                # A created Transcript is missing a few key values!
                # Add them here, or the comparison will definitely fail!
                transcript1.series_id = seriesName
                transcript1.episode_id = episodeName
                transcript1.changed = False

                if transcript1 == transcript2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                    break
                
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if sys.platform == 'win32':
            documentFileEnglish    = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Demo.rtf', 'utf8')
            documentFileApostrophe = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Demo2.rtf', 'utf8')
            documentFileFrench     = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Test \xc3\xa2\xc3\xab\xc3\xac\xc3\xb3\xc3\xbb 5.rtf', 'utf8')
            documentFileChinese    = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\\xe4\xba\xb2\xe4\xba\xb3 \xe4\xba\xb2\\\xe4\xba\xb2\xe4\xba\xb3\xe4\xba\xb2 5.rtf', 'utf8')
            documentFileArabic     = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4\\\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4.rtf', 'utf8')
            documentFileTesting    = unicode('C:\\Users\\DavidWoods\\Videos\\Harrie\\7000_No_TC.xml', 'utf8')
        
        elif sys.platform == 'darwin':
            documentFileEnglish    = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\x96d\xc3\xaa/Demo.rtf', 'utf8')
            documentFileApostrophe = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\x96d\xc3\xaa/Demo2.rtf', 'utf8')
            documentFileFrench     = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\xb6d\xc3\xaa/Test \xc3\xa2\xc3\xab\xc3\xac\xc3\xb3\xc3\xbb 5.rtf', 'utf8')
            documentFileChinese    = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\xb6d\xc3\xaa/\xe4\xba\xb2\xe4\xba\xb3 \xe4\xba\xb2/\xe4\xba\xb2\xe4\xba\xb3\xe4\xba\xb2 5.rtf', 'utf8')
            documentFileArabic     = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\xb6d\xc3\xaa/\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4/\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4.rtf', 'utf8')
            documentFileTesting    = unicode('/Volumes/Vid\xc3\xabo/Harrie/7000_No_TC.xml', 'utf8')

        testDocumentNames = [('English',
                              unicode('English Document 1', 'utf8'),
                              unicode('English Document Author 1', 'utf8'),
                              unicode('English Document Comment 1', 'utf8'),
                              testSeriesData['English'][0],
                              testSeriesData['English'][1],
                              documentFileEnglish,
                              testKeywordPairs[0][0],
                              testKeywordPairs[0][1]),
                             ("Apostrophe's Test",
                              unicode("Apostrophe's Test Document 1", 'utf8'),
                              unicode("Apostrophe's Test Document Author 1", 'utf8'),
                              unicode("Apostrophe's Test Document Comment 1", 'utf8'),
                              testSeriesData["Apostrophe's Test"][0],
                              testSeriesData["Apostrophe's Test"][1],
                              documentFileApostrophe,
                              testKeywordPairs[1][0],
                              testKeywordPairs[1][1]),
                             ('French',
                              unicode('Fran\xc3\xa7ais Document 1', 'utf8'),
                              unicode('Fran\xc3\xa7ais Document Author 1', 'utf8'),
                              unicode('Fran\xc3\xa7ais Document Commentaire 1', 'utf8'),
                              testSeriesData['French'][0],
                              testSeriesData['French'][1],
                              documentFileFrench,
                              testKeywordPairs[2][0],
                              testKeywordPairs[2][1]),
                             ('Chinese',
                              unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Document 1', 'utf8'),
                              unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Document Author 1', 'utf8'),
                              unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Document \xe4\xba\x8b\xe4\xbb\xb6 1', 'utf8'),
                              testSeriesData['Chinese'][0],
                              testSeriesData['Chinese'][1],
                              documentFileChinese,
                              testKeywordPairs[3][0],
                              testKeywordPairs[3][1]),
                             ('Arabic',
                              unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Document 1', 'utf8'),
                              unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Document Author 1', 'utf8'),
                              unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Document \xd8\xaa\xd8\xb9\xd9\x84\xd9\x8a\xd9\x82 1', 'utf8'),
                              testSeriesData['Arabic'][0],
                              testSeriesData['Arabic'][1],
                              documentFileArabic,
                              testKeywordPairs[4][0],
                              testKeywordPairs[4][1]),
                             ('Testing',
                              unicode('Testing', 'utf8'),
                              unicode('Testing Document Author 1', 'utf8'),
                              unicode('Testing Document Comment 1', 'utf8'),
                              testSeriesData['Testing'][0],
                              testSeriesData['Testing'][1],
                              documentFileTesting,
                              testKeywordPairs[5][0],
                              testKeywordPairs[5][1])]

        tempText = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
    <paragraph>
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">0</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">5 English Test.</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">10 Apostrophe's Test.</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">15 Fran\xc3\xa7ais.</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">20 \xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93.</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">25 \xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9.</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">30</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">35</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">40</text>
    </paragraph>
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">45</text>
    </paragraph>
  </paragraphlayout>
</richtext>
"""
        testDocumentData = {}

        for (language, documentName, documentAuthor, documentComment, libraryNum, libraryName, importFileName, keywordGroup, keyword) \
            in testDocumentNames:

            if 136 in testsToRun:
                # Document Saving
                testName = 'Creating Document : %s' % documentName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    document1 = Document.Document()
                    document1.id = documentName
                    document1.author = documentAuthor
                    document1.comment = documentComment
                    document1.library_num = libraryNum
                    document1.imported_file = importFileName
                    document1.add_keyword(keywordGroup, keyword)
                    if language == 'Testing':
#                        document1.text = documentFileTesting

                        # Open the file
                        f = open(documentFileTesting, "r")
                        # Read the file straight into the Transcript Text
                        document1.text = f.read()
                        # if the text does NOT have an RTF or XML header ...
                        if (document1.text[:5].lower() != '{\\rtf') and (document1.text[:5].lower() != '<?xml'):
                            # ... add "txt" to the start of the file to signal that it's probably a text file
                            document1.text = 'txt\n' + document1.text
                        # Close the file
                        f.close()

                    else:
                        document1.text = tempText
                    try:
                        document1.db_save()
                        # If we create the Document, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testDocumentData[language] = (document1.number, document1.id)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Document, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Document, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 137 in testsToRun:
                # Compare Created with Saved Document
                testName = 'Comparing Document by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)

                document2 = Document.Document(num=document1.number)

                # A created Document is missing the "Keyword Document Number" from keywords!
                # Add them here, or the comparison will definitely fail!
                for obj in document1._kwlist:
                    obj.documentNum = document1.number
                document1.changed = False

                if document1 == document2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                    print "Document by Number:"
                    print document1
                    print
                    print document2
                    print

                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 138 in testsToRun:
                # Compare Created with Saved Document
                testName = 'Comparing Document by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                document2 = Document.Document(libraryID=libraryName, documentID=documentName)

                if document1 == document2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testCollectionNames = [('English',
                                unicode('English Collection 1', 'utf8'),
                                unicode('English Collection Owner 1', 'utf8'),
                                unicode('English Collection Comment 1', 'utf8'),
                                testKeywordPairs[0][0]),
                               ("Apostrophe's Test",
                                unicode("Apostrophe's Test Collection 1", 'utf8'),
                                unicode("Apostrophe's Test Collection Owner 1", 'utf8'),
                                unicode("Apostrophe's Test Collection Comment 1", 'utf8'),
                                testKeywordPairs[1][0]),
                               ('French',
                                unicode('Fran\xc3\xa7ais Collection 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Collection D\xc3\xa9tenteur 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Collection Commentaire 1', 'utf8'),
                                testKeywordPairs[2][0]),
                               ('Chinese',
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe9\x9b\x86\xe5\x90\x88 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe9\x9b\x86\xe5\x90\x88 \xe6\x89\x80\xe6\x9c\x89\xe8\x80\x85 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe9\x9b\x86\xe5\x90\x88 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                                testKeywordPairs[3][0]),
                               ('Arabic',
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xac\xd9\x85\xd9\x88\xd8\xb9\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xac\xd9\x85\xd9\x88\xd8\xb9\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xa4\xd9\x84\xd9\x81 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xac\xd9\x85\xd9\x88\xd8\xb9\xd8\xa9 \xd8\xaa\xd8\xb9\xd9\x84\xd9\x8a\xd9\x82 1', 'utf8'),
                                testKeywordPairs[4][0]),
                               ('Testing',
                                unicode('Testing', 'utf8'),
                                unicode('Testing Collection Owner 1', 'utf8'),
                                unicode('Testing Collection Comment 1', 'utf8'),
                                testKeywordPairs[5][0])]

        testCollectionData = {}

        for (language, collectionName, collectionOwner, collectionComment, collectionKeywordGroup) in testCollectionNames:
            if 140 in testsToRun:
                try:
                    # Try to load the Test Collection
                    testCollection1 = Collection.Collection(collectionName, 0)
                    # Delete the Test Collection
                    testCollection1.db_delete()
                except TransanaExceptions.RecordNotFoundError:

#                    print u"Collection %s not found in test 140" % collectionName.decode('utf8')
                    
                    # This is the desired result!
                    pass
                except:

                    print u"Exception in Test 140 -- Could not delete Collection %s." % collectionName.decode('utf8')
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    import traceback
                    traceback.print_exc()
                    print

            if 150 in testsToRun:
                # Collection Saving
                testName = 'Creating Collection : %s' % collectionName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Collection
                    collection1 = Collection.Collection()
                    # Populate it with data
                    collection1.id = collectionName
                    collection1.parent = 0
                    collection1.owner = collectionOwner
                    collection1.comment = collectionComment
                    collection1.keyword_group = collectionKeywordGroup
                    try:
                        collection1.db_save()
                        # If we create the Collection, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testCollectionData[language] = (collection1.number, collection1.id)
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

            if 160 in testsToRun:
                # Compare Created with Saved Collection
                testName = 'Comparing Collections by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                collection2 = Collection.Collection(collection1.number)

                if collection1 == collection2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 165 in testsToRun:
                # Compare Created with Saved Collection
                testName = 'Comparing Collections by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                collection2 = Collection.Collection(collectionName, 0)

                if collection1 == collection2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if sys.platform == 'win32':
            mediaFileEnglish    = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Demo.mpg', 'utf8')
            mediaFileApostrophe = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Demo2.mpg', 'utf8')
            mediaFileFrench     = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Test \xc3\xa2\xc3\xab\xc3\xac\xc3\xb3\xc3\xbb 5.mpg', 'utf8')
            mediaFileChinese    = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\\xe4\xba\xb2\xe4\xba\xb3 \xe4\xba\xb2\\\xe4\xba\xb2\xe4\xba\xb3\xe4\xba\xb2 5.mpg', 'utf8')
            mediaFileArabic     = unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4\\\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4.mpg', 'utf8')
        
        elif sys.platform == 'darwin':
            mediaFileEnglish    = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\x96d\xc3\xaa/Demo.mpg', 'utf8')
            mediaFileApostrophe = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\x96d\xc3\xaa/Demo2.mpg', 'utf8')
            mediaFileFrench     = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\xb6d\xc3\xaa/Test \xc3\xa2\xc3\xab\xc3\xac\xc3\xb3\xc3\xbb 5.mpg', 'utf8')
            mediaFileChinese    = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\xb6d\xc3\xaa/\xe4\xba\xb2\xe4\xba\xb3 \xe4\xba\xb2/\xe4\xba\xb2\xe4\xba\xb3\xe4\xba\xb2 5.mpg', 'utf8')
            mediaFileArabic     = unicode('/Volumes/Vid\xc3\xabo/Unic\xc3\xb6d\xc3\xaa/\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4/\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4.mpg', 'utf8')

        testClipNames = [('English',
                          unicode('English Clip 1', 'utf8'),
                          unicode('English Clip Comment 1', 'utf8'),
                          mediaFileEnglish,
                          testEpisodeData['English'][0],
                          testTranscriptData['English'][0],
                          testCollectionData['English'][0],
                          testCollectionData['English'][1],
                          5000,
                          10000,
                          testKeywordPairs[0][0],
                          testKeywordPairs[0][1]),
                         ("Apostrophe's Test",
                          unicode("Apostrophe's Test Clip 1", 'utf8'),
                          unicode("Apostrophe's Test Clip Comment 1", 'utf8'),
                          mediaFileApostrophe,
                          testEpisodeData["Apostrophe's Test"][0],
                          testTranscriptData["Apostrophe's Test"][0],
                          testCollectionData["Apostrophe's Test"][0],
                          testCollectionData["Apostrophe's Test"][1],
                          10000,
                          15000,
                          testKeywordPairs[1][0],
                          testKeywordPairs[1][1]),
                         ('French',
                          unicode('Fran\xc3\xa7ais Extrait 1', 'utf8'),
                          unicode('Fran\xc3\xa7ais Extrait Commentaire 1', 'utf8'),
                          mediaFileFrench,
                          testEpisodeData['French'][0],
                          testTranscriptData['French'][0],
                          testCollectionData['French'][0],
                          testCollectionData['French'][1],
                          15000,
                          20000,
                          testKeywordPairs[2][0],
                          testKeywordPairs[2][1]),
                         ('Chinese',
                          unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe5\x89\xaa\xe8\xbe\x91 1', 'utf8'),
                          unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe5\x89\xaa\xe8\xbe\x91 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                          mediaFileChinese,
                          testEpisodeData['Chinese'][0],
                          testTranscriptData['Chinese'][0],
                          testCollectionData['Chinese'][0],
                          testCollectionData['Chinese'][1],
                          20000,
                          25000,
                          testKeywordPairs[3][0],
                          testKeywordPairs[3][1]),
                         ('Arabic',
                          unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x82\xd8\xb7\xd8\xb9 1', 'utf8'),
                          unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x82\xd8\xb7\xd8\xb9 \xd8\xaa\xd8\xb9\xd9\x84\xd9\x8a\xd9\x82 1', 'utf8'),
                          mediaFileArabic,
                          testEpisodeData['Arabic'][0],
                          testTranscriptData['Arabic'][0],
                          testCollectionData['Arabic'][0],
                          testCollectionData['Arabic'][1],
                          25000,
                          30000,
                          testKeywordPairs[4][0],
                          testKeywordPairs[4][1])]

        testClipTranscripts = {
              'English' :
"""<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
<paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
<paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
<text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">5 English Test</text>
</paragraph>
</paragraphlayout>
</richtext>""",
              "Apostrophe's Test" :
"""<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
<paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
<paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
<text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">10 Apostrophe's Test</text>
</paragraph>
</paragraphlayout>
</richtext>""",
              'French' :
"""<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
<paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
<paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
<text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">15 Fran\xc3\xa7ais Test</text>
</paragraph>
</paragraphlayout>
</richtext>""",
              'Chinese' :
"""<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
<paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
<paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
<text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">20 \xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Test</text>
</paragraph>
</paragraphlayout>
</richtext>""",
              'Arabic' :
"""<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
<paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10" tabs="">
<paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">
<text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">25 \xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Test</text>
</paragraph>
</paragraphlayout>
</richtext>"""
            }

        testClipData = {}

        for (language, clipName, clipComment, clipMediaFile, episodeNum, transcriptNum, collectionNum, collectionName, \
             clipStart, clipStop, keywordGroup, keyword) \
            in testClipNames:

            if 170 in testsToRun:
                # Clip Saving
                testName = 'Creating Clip : %s' % clipName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    clip1 = Clip.Clip()
                    clip1.id = clipName
                    clip1.comment = clipComment
                    clip1.episode_num = episodeNum
                    clip1.clip_start = clipStart
                    clip1.clip_stop = clipStop
                    clip1.collection_num = collectionNum
                    clip1.collection_id = collectionName
                    clip1.sort_order = DBInterface.getMaxSortOrder(collectionNum) + 1
                    clip1.offset = 0
                    clip1.media_filename = clipMediaFile
                    clip1.audio = 1
                    clip1.add_keyword(keywordGroup, keyword)
                    # Prepare the Clip Transcript Object
                    cliptranscript1 = Transcript.Transcript()
                    cliptranscript1.episode_num = episodeNum
                    cliptranscript1.source_transcript = transcriptNum
                    cliptranscript1.clip_start = clip1.clip_start
                    cliptranscript1.clip_stop = clip1.clip_stop
                    cliptranscript1.text = testClipTranscripts[language]
                    clip1.transcripts.append(cliptranscript1)
                    try:
                        clip1.db_save()
                        # If we create the Clip, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testClipData[language] = (clip1.number, clip1.id, cliptranscript1.number, cliptranscript1.source_transcript, clip1.episode_num, cliptranscript1.sort_order)
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
                    traceback.print_exc(file=sys.stdout)
                    # If we can't load or create the Clip, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 180 in testsToRun:
                # Compare Created with Saved Clip
                testName = 'Comparing Clips by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)

                # Add the "changed" property to the created clip
                for tr in clip1.transcripts:
                    tr.changed = False
                for ckw in clip1._kwlist:
                    ckw.clipNum = clip1.number
                
                clip2 = Clip.Clip(clip1.number)

                clip1.useVideoRoot = clip2.useVideoRoot

                if clip1 == clip2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 185 in testsToRun:
                # Compare Created with Saved Clip
                testName = 'Comparing Clips by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)

                # Add the "changed" property to the created clip
                for tr in clip1.transcripts:
                    tr.changed = False
                for ckw in clip1._kwlist:
                    ckw.clipNum = clip1.number
                
                clip2 = Clip.Clip(clipName, collectionName, 0)

                clip1.useVideoRoot = clip2.useVideoRoot

                if clip1 == clip2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testQuoteNames = [('English',
                           unicode('English Quote 1', 'utf8'),
                           testCollectionData['English'][0],
                           testCollectionData['English'][1],
                           testDocumentData['English'][0],
                           4,
                           16,
                           unicode('English Quote Comment 1', 'utf8'),
                           testKeywordPairs[0][0],
                           testKeywordPairs[0][1]),
                          ("Apostrophe's Test",
                           unicode("Apostrophe's Test Quote 1", 'utf8'),
                           testCollectionData["Apostrophe's Test"][0],
                           testCollectionData["Apostrophe's Test"][1],
                           testDocumentData["Apostrophe's Test"][0],
                           21,
                           38,
                           unicode("Apostrophe's Test Quote Comment 1", 'utf8'),
                           testKeywordPairs[1][0],
                           testKeywordPairs[1][1]),
                          ('French',
                           unicode('Fran\xc3\xa7ais Quote 1', 'utf8'),
                           testCollectionData['French'][0],
                           testCollectionData['French'][1],
                           testDocumentData['French'][0],
                           43,
                           51,
                           unicode('Fran\xc3\xa7ais Quote Commentaire 1', 'utf8'),
                           testKeywordPairs[2][0],
                           testKeywordPairs[2][1]),
                          ('Chinese',
                           unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Quote 1', 'utf8'),
                           testCollectionData['Chinese'][0],
                           testCollectionData['Chinese'][1],
                           testDocumentData['Chinese'][0],
                           56,
                           61,
                           unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Quote \xe4\xba\x8b\xe4\xbb\xb6 1', 'utf8'),
                           testKeywordPairs[3][0],
                           testKeywordPairs[3][1]),
                          ('Arabic',
                           unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Quote 1', 'utf8'),
                           testCollectionData['Arabic'][0],
                           testCollectionData['Arabic'][1],
                           testDocumentData['Arabic'][0],
                           66,
                           73,
                           unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Quote \xd8\xaa\xd8\xb9\xd9\x84\xd9\x8a\xd9\x82 1', 'utf8'),
                           testKeywordPairs[4][0],
                           testKeywordPairs[4][1])]

        testingQuoteText = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
  <paragraphlayout textcolor="#000000" fontpointsize="9" fontfamily="70" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Segoe UI" alignment="1" parspacingafter="10" parspacingbefore="0" linespacing="10" margin-left="5,4098" margin-right="5,4098" margin-top="5,4098" margin-bottom="5,4098">
    <paragraph textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontfamily="70" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" parspacingafter="10" parspacingbefore="0" linespacing="10">
      <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">ABCD %04d</text>
    </paragraph>
  </paragraphlayout>
</richtext>"""

        # Add the Testing Quotes
        for x in range(1, numTestQuotes + 1):
            tmpRec = ('Testing',
                           unicode('Testing Quote %04d' % x, 'utf8'),
                           testCollectionData['Testing'][0],
                           testCollectionData['Testing'][1],
                           testDocumentData['Testing'][0],
                           (x - 1) * 72 + 7,
                           (x - 1) * 72 + 33,
                           unicode('Testing Quote Comment %04d' % x, 'utf8'),
                           testKeywordPairs[5][0],
                           testKeywordPairs[5][1])
            testQuoteNames += (tmpRec,)

        testQuoteData = {}
        tmpCnt = 0

        for (language, quoteName, collectionNum, collectionName, sourceDocumentNum, startChar, endChar, quoteComment, \
             keywordGroup, keyword) in testQuoteNames:

            if 175 in testsToRun:
                # Quote Saving
                testName = 'Creating Quote : %s' % quoteName
                if language != 'Testing':
                    self.SetStatusText(testName)
                    self.testsRun += 1
                    self.txtCtrl.AppendText('Test "%s" ' % testName)
                else:
                    tmpCnt += 1
                    if (tmpCnt % 100 == 1) or (tmpCnt == numTestQuotes):
                        self.txtCtrl.AppendText("Creating %s of %s\n" % (quoteName, numTestQuotes))
                        self.Update()
                try:
                    quote1 = Quote.Quote()
                    quote1.id = quoteName
                    quote1.collection_num = collectionNum
                    quote1.source_document_num = sourceDocumentNum
                    quote1.sort_order = DBInterface.getMaxSortOrder(collectionNum) + 1
                    quote1.start_char = startChar
                    if language != 'Testing':
                        # We can use the Clip Transcripts as Quote Text too!
                        quote1.text = testClipTranscripts[language]
                    else:
                        quote1.text = testingQuoteText % (tmpCnt)
                    quote1.end_char = endChar
                    quote1.comment = quoteComment
                    quote1.add_keyword(keywordGroup, keyword)
                    try:
                        quote1.db_save()
                        if language != 'Testing':
                            # If we create the Quote, consider the test passed.
                            self.txtCtrl.AppendText('Passed.')
                            self.testsSuccessful += 1
                            testQuoteData[language] = (quote1.number, quote1.id, quote1.collection_num, quote1.source_document_num)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Quote, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Quote, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                if language != 'Testing':
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if language != 'Testing':
                if 176 in testsToRun:
                    # Compare Created with Saved Quote
                    testName = 'Comparing Quotes by Number'
                    self.SetStatusText(testName)
                    self.testsRun += 1
                    self.txtCtrl.AppendText('Test "%s" ' % testName)

                    # A created Quote is missing the "Keyword Document Number" from keywords!
                    # Add them here, or the comparison will definitely fail!
                    for obj in quote1._kwlist:
                        obj.quoteNum = quote1.number
                    quote1.changed = False

                    quote2 = Quote.Quote(quote1.number)

                    # Add the missing CollectionID text to Quote1, which doesn't get it when created.
                    quote1.collection_id = quote2.collection_id

                    if quote1 == quote2:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')

                        print "Failed:", testName
                        print quote1
                        print
                        print quote2
                        print
                        
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

                if 177 in testsToRun:
                    # Compare Created with Saved Clip
                    testName = 'Comparing Quotes by Name'
                    self.SetStatusText(testName)
                    self.testsRun += 1
                    self.txtCtrl.AppendText('Test "%s" ' % testName)

                    quote2 = Quote.Quote(quoteID = quoteName, collectionID = collectionName, collectionParent = 0)

                    if quote1 == quote2:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if sys.platform == 'win32':
            mediaFile1 = 'C:\\Users\\DavidWoods\\Videos'
        
        elif sys.platform == 'darwin':
            mediaFile1 = unicode('/Volumes/Vid\xc3\xabo', 'utf8')

        testSnapshotNames = [('English',
                              unicode('English Snapshot 1', 'utf8'),
                              unicode('English Snapshot Comment 1', 'utf8'),
                              os.path.join(mediaFile1, 'Images', 'ch130214.gif'),
                              testSeriesData['English'][0],
                              testSeriesData['English'][1],
                              testEpisodeData['English'][0],
                              testEpisodeData['English'][1],
                              testTranscriptData['English'][0],
                              testTranscriptData['English'][1],
                              testCollectionData['English'][0],
                              testCollectionData['English'][1],
                              12500, 5000,
                              testKeywordPairs[0][0],
                              testKeywordPairs[0][1]),
                             ("Apostrophe's Test",
                              unicode("Apostrophe's Test Snapshot 1", 'utf8'),
                              unicode("Apostrophe's Test Snapshot Comment 1", 'utf8'),
                              os.path.join(mediaFile1, 'Images', 'ch130214.gif'),
                              testSeriesData["Apostrophe's Test"][0],
                              testSeriesData["Apostrophe's Test"][1],
                              testEpisodeData["Apostrophe's Test"][0],
                              testEpisodeData["Apostrophe's Test"][1],
                              testTranscriptData["Apostrophe's Test"][0],
                              testTranscriptData["Apostrophe's Test"][1],
                              testCollectionData["Apostrophe's Test"][0],
                              testCollectionData["Apostrophe's Test"][1],
                              17500, 5000,
                              testKeywordPairs[1][0],
                              testKeywordPairs[1][1]),
                             ('French',
                              unicode('Fran\xc3\xa7ais Capture instantan\xc3\xa9e 1', 'utf8'),
                              unicode('Fran\xc3\xa7ais Capture instantan\xc3\xa9e Commentaire 1', 'utf8'),
                              os.path.join(mediaFile1, 'Images', 'ch130214.gif'), # unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\Test \xc3\xa2\xc3\xab\xc3\xac\xc3\xb3\xc3\xbb 5.mpg', 'utf8'),
                              testSeriesData['French'][0],
                              testSeriesData['French'][1],
                              testEpisodeData['French'][0],
                              testEpisodeData['French'][1],
                              testTranscriptData['French'][0],
                              testTranscriptData['French'][1],
                              testCollectionData['French'][0],
                              testCollectionData['French'][1],
                              22500, 5000,
                              testKeywordPairs[2][0],
                              testKeywordPairs[2][1]),
                             ('Chinese',
                              unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe6\x88\xaa\xe5\x9b\xbe 1', 'utf8'),
                              unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe6\x88\xaa\xe5\x9b\xbe \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                              os.path.join(mediaFile1, 'Images', 'ch130214.gif'), # unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\\xe4\xba\xb2\xe4\xba\xb3 \xe4\xba\xb2\\\xe4\xba\xb2\xe4\xba\xb3\xe4\xba\xb2 5.mpg', 'utf8'),
                              testSeriesData['Chinese'][0],
                              testSeriesData['Chinese'][1],
                              testEpisodeData['Chinese'][0],
                              testEpisodeData['Chinese'][1],
                              testTranscriptData['Chinese'][0],
                              testTranscriptData['Chinese'][1],
                              testCollectionData['Chinese'][0],
                              testCollectionData['Chinese'][1],
                              27500, 5000,
                              testKeywordPairs[3][0],
                              testKeywordPairs[3][1]),
                             ('Arabic',
                              unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x84\xd9\x82\xd8\xb7\xd8\xa9 1', 'utf8'),
                              unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x84\xd9\x82\xd8\xb7\xd8\xa9 \xd8\xaa\xd8\xb9\xd9\x84\xd9\x8a\xd9\x82 1', 'utf8'),
                              os.path.join(mediaFile1, 'Images', 'ch130214.gif'), # unicode('C:\\Users\\DavidWoods\\Videos\\Unic\xc3\xb6d\xc3\xaa\\\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4\\\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4.mpg', 'utf8'),
                              testSeriesData['Arabic'][0],
                              testSeriesData['Arabic'][1],
                              testEpisodeData['Arabic'][0],
                              testEpisodeData['Arabic'][1],
                              testTranscriptData['Arabic'][0],
                              testTranscriptData['Arabic'][1],
                              testCollectionData['Arabic'][0],
                              testCollectionData['Arabic'][1],
                              32500, 5000,
                              testKeywordPairs[4][0],
                              testKeywordPairs[4][1])]

        testSnapshotData = {}

        for (language, snapshotName, snapshotComment, snapshotMediaFile, seriesNum, seriesName, episodeNum, episodeName,
             transcriptNum, transcriptName, collectionNum, collectionName, startTime, duration, keywordGroup, keyword) \
            in testSnapshotNames:

            if 190 in testsToRun:
                # Snapshot Saving
                testName = 'Creating Snapshot : %s' % snapshotName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    snapshot1 = Snapshot.Snapshot()
                    snapshot1.id = snapshotName
                    snapshot1.comment = snapshotComment
                    snapshot1.collection_num = collectionNum
                    snapshot1.collection_id = collectionName
                    snapshot1.image_filename = snapshotMediaFile
                    snapshot1.image_scale = 1.19066666667
                    snapshot1.image_size = (768, 391)
                    snapshot1.series_num = seriesNum
                    snapshot1.series_id = seriesName
                    snapshot1.episode_num = episodeNum
                    snapshot1.episode_id = episodeName
                    snapshot1.transcript_num = transcriptNum
                    snapshot1.transcript_id = transcriptName
                    snapshot1.episode_start = startTime
                    snapshot1.episode_duration = duration
                    snapshot1.sort_order = DBInterface.getMaxSortOrder(collection1.number) + 1
                    snapshot1.add_keyword(keywordGroup, keyword)
                    snapshot1.codingObjects = {0: {'keywordGroup': keywordGroup,
                                                   'keyword': keyword,
                                                   'x1': -294L,
                                                   'x2': -150L,
                                                   'y1': 91L,
                                                   'y2': -92L,
                                                   'visible': True}}
                    snapshot1.keywordStyles = {(keywordGroup, keyword): {'lineColorDef': u'#ff0000',
                                                                         'drawMode': u'Rectangle',
                                                                         'lineStyle': u'Solid',
                                                                         'lineWidth': '5',
                                                                         'lineColorName': u'Red'}}
                    try:
                        snapshot1.db_save()
                        # If we create the Snapshot, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testSnapshotData[language] = (snapshot1.number, snapshot1.id)
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
                    traceback.print_exc(file=sys.stdout)

                    # If we can't load or create the Snapshot, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 200 in testsToRun:
                # Compare Created with Saved Snapshot
                testName = 'Comparing Snapshots by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)

                # Add the snapshot number to the keyword object or the comparison will fail
                snapshot1._kwlist[0].snapshotNum = snapshot1.number
                
                snapshot2 = Snapshot.Snapshot(snapshot1.number)

                snapshot1.useVideoRoot = snapshot2.useVideoRoot

                if snapshot1 == snapshot2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 205 in testsToRun:
                # Compare Created with Saved Snapshot
                testName = 'Comparing Snapshots by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)

                # Add the snapshot number to the keyword object or the comparison will fail
                snapshot1._kwlist[0].snapshotNum = snapshot1.number
                
                snapshot2 = Snapshot.Snapshot(snapshotName, collectionNum)

                snapshot1.useVideoRoot = snapshot2.useVideoRoot

                if snapshot1 == snapshot2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1

                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testSeriesNoteNames = [('English',
                                unicode('English Library Note 1', 'utf8'),
                                unicode('English Library Note Comment 1', 'utf8'),
                                unicode('English Library Note NoteTaker 1', 'utf8'),
                                unicode('English Library Note Text 1', 'utf8'),
                                testSeriesData['English'][0]),
                               ("Apostrophe's Test",
                                unicode("Apostrophe's Test Library Note 1", 'utf8'),
                                unicode("Apostrophe's Test Library Note Comment 1", 'utf8'),
                                unicode("Apostrophe's Test Library Note NoteTaker 1", 'utf8'),
                                unicode("Apostrophe's Test Library Note Text 1", 'utf8'),
                                testSeriesData["Apostrophe's Test"][0]),
                               ('French',
                                unicode('Fran\xc3\xa7ais S\xc3\xa9ries Note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais S\xc3\xa9ries Note Commentaire 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais S\xc3\xa9ries Auteur-e de la note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais S\xc3\xa9ries Contenu de la note 1', 'utf8'),
                                testSeriesData['French'][0]),
                               ('Chinese',
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe5\xba\x8f\xe5\x88\x97 \xe5\xa4\x87\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe5\xba\x8f\xe5\x88\x97 \xe5\xa4\x87\xe6\xb3\xa8 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe5\xba\x8f\xe5\x88\x97 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x8f\x90\xe5\x8f\x96\xe5\x99\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe5\xba\x8f\xe5\x88\x97 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x96\x87\xe6\x9c\xac\xef\xbc\x9a 1', 'utf8'),
                                testSeriesData['Chinese'][0]),
                               ('Arabic',
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xb3\xd9\x84\xd8\xb3\xd9\x84\xd8\xa7\xd8\xaa \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xb3\xd9\x84\xd8\xb3\xd9\x84\xd8\xa7\xd8\xaa \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9  1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xb3\xd9\x84\xd8\xb3\xd9\x84\xd8\xa7\xd8\xaa \xd9\x85\xd8\xa4\xd9\x84\xd9\x81 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xb3\xd9\x84\xd8\xb3\xd9\x84\xd8\xa7\xd8\xaa \xd9\x86\xd8\xb5 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                testSeriesData['Arabic'][0])]

        testNoteData = {}

        for (language, noteName, noteComment, noteTaker, noteText, seriesNum) in testSeriesNoteNames:
            testNoteData[language] = {'Libraries' : [], 'Episode' : [], 'Transcript' : [],
                                      'Collection' : [], 'Clip' : [], 'Snapshot' : [],
                                      'Document' : [], 'Quote' : []}
            if 210 in testsToRun:
                try:
                    # Try to load the Test Note
                    testNote1 = Note.Note(noteName, Library=seriesNum)
                    # Delete the Test Note
                    testNote1.db_delete()
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    pass
                except:

                    print "Exception in Test 210 -- Could not delete Note %s." % noteName.encode('utf8')
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    print

            if 220 in testsToRun:
                # Note Saving
                testName = 'Creating Library Note : %s' % noteName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Note
                    note1 = Note.Note()
                    # Populate it with data
                    note1.id = noteName
#                    note1.comment = noteComment
                    note1.author = noteTaker
                    note1.text = noteText
                    note1.series_num = seriesNum
                    try:
                        note1.db_save()
                        # If we create the Note, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testNoteData[language]['Libraries'].append((note1.number, note1.id),)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Note, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Note, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 230 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(note1.number)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 235 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(noteName, Library=seriesNum)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testEpisodeNoteNames = [('English',
                                unicode('English Episode Note 1', 'utf8'),
                                unicode('English Episode Note Comment 1', 'utf8'),
                                unicode('English Episode Note NoteTaker 1', 'utf8'),
                                unicode('English Episode Note Text 1', 'utf8'),
                                testEpisodeData['English'][0]),
                               ("Apostrophe's Test",
                                unicode("Apostrophe's Test Episode Note 1", 'utf8'),
                                unicode("Apostrophe's Test Episode Note Comment 1", 'utf8'),
                                unicode("Apostrophe's Test Episode Note NoteTaker 1", 'utf8'),
                                unicode("Apostrophe's Test Episode Note Text 1", 'utf8'),
                                testEpisodeData["Apostrophe's Test"][0]),
                               ('French',
                                unicode('Fran\xc3\xa7ais \xc3\x89pisode Note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais \xc3\x89pisode Note Commentaire 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais \xc3\x89pisode Auteur-e de la note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais \xc3\x89pisode Contenu de la note 1', 'utf8'),
                                testEpisodeData['French'][0]),
                               ('Chinese',
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe4\xba\x8b\xe4\xbb\xb6 \xe5\xa4\x87\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe4\xba\x8b\xe4\xbb\xb6 \xe5\xa4\x87\xe6\xb3\xa8 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe4\xba\x8b\xe4\xbb\xb6 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x8f\x90\xe5\x8f\x96\xe5\x99\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe4\xba\x8b\xe4\xbb\xb6 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x96\x87\xe6\x9c\xac\xef\xbc\x9a 1', 'utf8'),
                                testEpisodeData['Chinese'][0]),
                               ('Arabic',
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xad\xd9\x84\xd9\x82\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xad\xd9\x84\xd9\x82\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9  1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xad\xd9\x84\xd9\x82\xd8\xa9 \xd9\x85\xd8\xa4\xd9\x84\xd9\x81 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xad\xd9\x84\xd9\x82\xd8\xa9 \xd9\x86\xd8\xb5 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                testEpisodeData['Arabic'][0])]

        for (language, noteName, noteComment, noteTaker, noteText, episodeNum) in testEpisodeNoteNames:
            if 210 in testsToRun:
                try:
                    # Try to load the Test Note
                    testNote1 = Note.Note(noteName, Episode=episodeNum)
                    # Delete the Test Note
                    testNote1.db_delete()
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    pass
                except sqlite3.OperationalError:
                    # This is the desired result!
                    pass
                except:

                    print "Exception in Test 210 -- Could not delete Note %s." % noteName
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    print

            if 220 in testsToRun:
                # Note Saving
                testName = 'Creating Episode Note : %s' % noteName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Note
                    note1 = Note.Note()
                    # Populate it with data
                    note1.id = noteName
#                    note1.comment = noteComment
                    note1.author = noteTaker
                    note1.text = noteText
                    note1.episode_num = episodeNum
                    try:
                        note1.db_save()
                        # If we create the Note, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testNoteData[language]['Episode'].append((note1.number, note1.id),)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Note, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Note, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 230 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(note1.number)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 235 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(noteName, Episode=episodeNum)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testTranscriptNoteNames = [('English',
                                unicode('English Transcript Note 1', 'utf8'),
                                unicode('English Transcript Note Comment 1', 'utf8'),
                                unicode('English Transcript Note NoteTaker 1', 'utf8'),
                                unicode('English Transcript Note Text 1', 'utf8'),
                                testTranscriptData['English'][0]),
                               ("Apostrophe's Test",
                                unicode("Apostrophe's Test Transcript Note 1", 'utf8'),
                                unicode("Apostrophe's Test Transcript Note Comment 1", 'utf8'),
                                unicode("Apostrophe's Test Transcript Note NoteTaker 1", 'utf8'),
                                unicode("Apostrophe's Test Transcript Note Text 1", 'utf8'),
                                testTranscriptData["Apostrophe's Test"][0]),
                               ('French',
                                unicode('Fran\xc3\xa7ais Transcription Note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Transcription Note Commentaire 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Transcription Auteur-e de la note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Transcription Contenu de la note 1', 'utf8'),
                                testTranscriptData['French'][0]),
                               ('Chinese',
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xbd\xac\xe5\xbd\x95 \xe5\xa4\x87\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xbd\xac\xe5\xbd\x95 \xe5\xa4\x87\xe6\xb3\xa8 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xbd\xac\xe5\xbd\x95 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x8f\x90\xe5\x8f\x96\xe5\x99\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xbd\xac\xe5\xbd\x95 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x96\x87\xe6\x9c\xac\xef\xbc\x9a 1', 'utf8'),
                                testTranscriptData['Chinese'][0]),
                               ('Arabic',
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xaa\xd8\xaf\xd9\x88\xd9\x8a\xd9\x86\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xaa\xd8\xaf\xd9\x88\xd9\x8a\xd9\x86\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9  1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xaa\xd8\xaf\xd9\x88\xd9\x8a\xd9\x86\xd8\xa9 \xd9\x85\xd8\xa4\xd9\x84\xd9\x81 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xaa\xd8\xaf\xd9\x88\xd9\x8a\xd9\x86\xd8\xa9 \xd9\x86\xd8\xb5 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                testTranscriptData['Arabic'][0])]

        for (language, noteName, noteComment, noteTaker, noteText, transcriptNum) in testTranscriptNoteNames:
            if 210 in testsToRun:
                try:
                    # Try to load the Test Note
                    testNote1 = Note.Note(noteName, Transcript=transcriptNum)
                    # Delete the Test Note
                    testNote1.db_delete()
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    pass
                except sqlite3.OperationalError:
                    # This is the desired result!
                    pass
                except:

                    print "Exception in Test 210 -- Could not delete Note %s." % noteName
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    print

            if 220 in testsToRun:
                # Note Saving
                testName = 'Creating Transcript Note : %s' % noteName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Note
                    note1 = Note.Note()
                    # Populate it with data
                    note1.id = noteName
#                    note1.comment = noteComment
                    note1.author = noteTaker
                    note1.text = noteText
                    note1.transcript_num = transcriptNum
                    try:
                        note1.db_save()
                        # If we create the Note, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testNoteData[language]['Transcript'].append((note1.number, note1.id),)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Note, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Note, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 230 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(note1.number)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 235 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(noteName, Transcript=transcriptNum)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testCollectionNoteNames = [('English',
                                unicode('English Collection Note 1', 'utf8'),
                                unicode('English Collection Note Comment 1', 'utf8'),
                                unicode('English Collection Note NoteTaker 1', 'utf8'),
                                unicode('English Collection Note Text 1', 'utf8'),
                                testCollectionData['English'][0]),
                               ("Apostrophe's Test",
                                unicode("Apostrophe's Test Collection Note 1", 'utf8'),
                                unicode("Apostrophe's Test Collection Note Comment 1", 'utf8'),
                                unicode("Apostrophe's Test Collection Note NoteTaker 1", 'utf8'),
                                unicode("Apostrophe's Test Collection Note Text 1", 'utf8'),
                                testCollectionData["Apostrophe's Test"][0]),
                               ('French',
                                unicode('Fran\xc3\xa7ais Collection Note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Collection Note Commentaire 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Collection Auteur-e de la note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Collection Contenu de la note 1', 'utf8'),
                                testCollectionData['French'][0]),
                               ('Chinese',
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe9\x9b\x86\xe5\x90\x88 \xe5\xa4\x87\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe9\x9b\x86\xe5\x90\x88 \xe5\xa4\x87\xe6\xb3\xa8 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe9\x9b\x86\xe5\x90\x88 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x8f\x90\xe5\x8f\x96\xe5\x99\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe8\xa7\x86\xe9\xa2\x91\xe9\x9b\x86\xe5\x90\x88 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x96\x87\xe6\x9c\xac\xef\xbc\x9a 1', 'utf8'),
                                testCollectionData['Chinese'][0]),
                               ('Arabic',
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xac\xd9\x85\xd9\x88\xd8\xb9\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xac\xd9\x85\xd9\x88\xd8\xb9\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9  1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xac\xd9\x85\xd9\x88\xd8\xb9\xd8\xa9 \xd9\x85\xd8\xa4\xd9\x84\xd9\x81 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xac\xd9\x85\xd9\x88\xd8\xb9\xd8\xa9 \xd9\x86\xd8\xb5 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                testCollectionData['Arabic'][0])]

        for (language, noteName, noteComment, noteTaker, noteText, collectionNum) in testCollectionNoteNames:
            if 210 in testsToRun:
                try:
                    # Try to load the Test Note
                    testNote1 = Note.Note(noteName, Collection=collectionNum)
                    # Delete the Test Note
                    testNote1.db_delete()
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    pass
                except sqlite3.OperationalError:
                    # This is the desired result!
                    pass
                except:

                    print "Exception in Test 210 -- Could not delete Note %s." % noteName
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    print

            if 220 in testsToRun:
                # Note Saving
                testName = 'Creating Collection Note : %s' % noteName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Note
                    note1 = Note.Note()
                    # Populate it with data
                    note1.id = noteName
#                    note1.comment = noteComment
                    note1.author = noteTaker
                    note1.text = noteText
                    note1.collection_num = collectionNum
                    try:
                        note1.db_save()
                        # If we create the Note, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testNoteData[language]['Collection'].append((note1.number, note1.id),)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Note, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Note, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 230 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(note1.number)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 235 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(noteName, Collection=collectionNum)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testClipNoteNames = [('English',
                                unicode('English Clip Note 1', 'utf8'),
                                unicode('English Clip Note Comment 1', 'utf8'),
                                unicode('English Clip Note NoteTaker 1', 'utf8'),
                                unicode('English Clip Note Text 1', 'utf8'),
                                testClipData['English'][0]),
                               ("Apostrophe's Test",
                                unicode("Apostrophe's Test Clip Note 1", 'utf8'),
                                unicode("Apostrophe's Test Clip Note Comment 1", 'utf8'),
                                unicode("Apostrophe's Test Clip Note NoteTaker 1", 'utf8'),
                                unicode("Apostrophe's Test Clip Note Text 1", 'utf8'),
                                testClipData["Apostrophe's Test"][0]),
                               ('French',
                                unicode('Fran\xc3\xa7ais Extrait Note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Extrait Note Commentaire 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Extrait Auteur-e de la note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Extrait Contenu de la note 1', 'utf8'),
                                testClipData['French'][0]),
                               ('Chinese',
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe5\x89\xaa\xe8\xbe\x91 \xe5\xa4\x87\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe5\x89\xaa\xe8\xbe\x91 \xe5\xa4\x87\xe6\xb3\xa8 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe5\x89\xaa\xe8\xbe\x91 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x8f\x90\xe5\x8f\x96\xe5\x99\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe5\x89\xaa\xe8\xbe\x91 \xe5\xa4\x87\xe6\xb3\xa8\xe6\x96\x87\xe6\x9c\xac\xef\xbc\x9a 1', 'utf8'),
                                testClipData['Chinese'][0]),
                               ('Arabic',
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x82\xd8\xb7\xd8\xb9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x82\xd8\xb7\xd8\xb9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9  1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x82\xd8\xb7\xd8\xb9 \xd9\x85\xd8\xa4\xd9\x84\xd9\x81 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x82\xd8\xb7\xd8\xb9 \xd9\x86\xd8\xb5 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                testClipData['Arabic'][0])]

        for (language, noteName, noteComment, noteTaker, noteText, clipNum) in testClipNoteNames:
            if 210 in testsToRun:
                try:
                    # Try to load the Test Note
                    testNote1 = Note.Note(noteName, Clip=clipNum)
                    # Delete the Test Note
                    testNote1.db_delete()
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    pass
                except sqlite3.OperationalError:
                    # This is the desired result!
                    pass
                except:

                    print "Exception in Test 210 -- Could not delete Note %s." % noteName
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    print

            if 220 in testsToRun:
                # Note Saving
                testName = 'Creating Clip Note : %s' % noteName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Note
                    note1 = Note.Note()
                    # Populate it with data
                    note1.id = noteName
#                    note1.comment = noteComment
                    note1.author = noteTaker
                    note1.text = noteText
                    note1.clip_num = clipNum
                    try:
                        note1.db_save()
                        # If we create the Note, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testNoteData[language]['Clip'].append((note1.number, note1.id),)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Note, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Note, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 230 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(note1.number)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 235 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(noteName, Clip=clipNum)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testSnapshotNoteNames = [('English',
                                unicode('English Snapshot Note 1', 'utf8'),
                                unicode('English Snapshot Note Comment 1', 'utf8'),
                                unicode('English Snapshot Note NoteTaker 1', 'utf8'),
                                unicode('English Snapshot Note Text 1', 'utf8'),
                                testSnapshotData['English'][0]),
                               ("Apostrophe's Test",
                                unicode("Apostrophe's Test Snapshot Note 1", 'utf8'),
                                unicode("Apostrophe's Test Snapshot Note Comment 1", 'utf8'),
                                unicode("Apostrophe's Test Snapshot Note NoteTaker 1", 'utf8'),
                                unicode("Apostrophe's Test Snapshot Note Text 1", 'utf8'),
                                testSnapshotData["Apostrophe's Test"][0]),
                               ('French',
                                unicode('Fran\xc3\xa7ais Capture instantan\xc3\xa9e Note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Capture instantan\xc3\xa9e Note Commentaire 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Capture instantan\xc3\xa9e Auteur-e de la note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Capture instantan\xc3\xa9e Contenu de la note 1', 'utf8'),
                                testSnapshotData['French'][0]),
                               ('Chinese',
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe6\x88\xaa\xe5\x9b\xbe \xe5\xa4\x87\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe6\x88\xaa\xe5\x9b\xbe \xe5\xa4\x87\xe6\xb3\xa8 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe6\x88\xaa\xe5\x9b\xbe \xe5\xa4\x87\xe6\xb3\xa8\xe6\x8f\x90\xe5\x8f\x96\xe5\x99\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 \xe6\x88\xaa\xe5\x9b\xbe \xe5\xa4\x87\xe6\xb3\xa8\xe6\x96\x87\xe6\x9c\xac\xef\xbc\x9a 1', 'utf8'),
                                testSnapshotData['Chinese'][0]),
                               ('Arabic',
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x84\xd9\x82\xd8\xb7\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x84\xd9\x82\xd8\xb7\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9  1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x84\xd9\x82\xd8\xb7\xd8\xa9 \xd9\x85\xd8\xa4\xd9\x84\xd9\x81 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x84\xd9\x82\xd8\xb7\xd8\xa9 \xd9\x86\xd8\xb5 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                testSnapshotData['Arabic'][0])]

        for (language, noteName, noteComment, noteTaker, noteText, snapshotNum) in testSnapshotNoteNames:
            if 210 in testsToRun:
                try:
                    # Try to load the Test Note
                    testNote1 = Note.Note(noteName, Snapshot=snapshotNum)
                    # Delete the Test Note
                    testNote1.db_delete()
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    pass
                except sqlite3.OperationalError:
                    # This is the desired result!
                    pass
                except:

                    print "Exception in Test 210 -- Could not delete Note %s." % noteName
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    print

            if 220 in testsToRun:
                # Note Saving
                testName = 'Creating Snapshot Note : %s' % noteName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Note
                    note1 = Note.Note()
                    # Populate it with data
                    note1.id = noteName
#                    note1.comment = noteComment
                    note1.author = noteTaker
                    note1.text = noteText
                    note1.snapshot_num = snapshotNum
                    try:
                        note1.db_save()
                        # If we create the Note, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testNoteData[language]['Snapshot'].append((note1.number, note1.id),)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Note, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Note, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 230 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(note1.number)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 235 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(noteName, Snapshot=snapshotNum)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testDocumentNoteNames = [('English',
                                unicode('English Document Note 1', 'utf8'),
                                unicode('English Document Note Comment 1', 'utf8'),
                                unicode('English Document Note NoteTaker 1', 'utf8'),
                                unicode('English Document Note Text 1', 'utf8'),
                                testDocumentData['English'][0]),
                               ("Apostrophe's Test",
                                unicode("Apostrophe's Test Document Note 1", 'utf8'),
                                unicode("Apostrophe's Test Document Note Comment 1", 'utf8'),
                                unicode("Apostrophe's Test Document Note NoteTaker 1", 'utf8'),
                                unicode("Apostrophe's Test Document Note Text 1", 'utf8'),
                                testDocumentData["Apostrophe's Test"][0]),
                               ('French',
                                unicode('Fran\xc3\xa7ais Document Note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Document Note Commentaire 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Document Auteur-e de la note 1', 'utf8'),
                                unicode('Fran\xc3\xa7ais Document Contenu de la note 1', 'utf8'),
                                testDocumentData['French'][0]),
                               ('Chinese',
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Document \xe5\xa4\x87\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Document \xe5\xa4\x87\xe6\xb3\xa8 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Document \xe5\xa4\x87\xe6\xb3\xa8\xe6\x8f\x90\xe5\x8f\x96\xe5\x99\xa8 1', 'utf8'),
                                unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Document \xe5\xa4\x87\xe6\xb3\xa8\xe6\x96\x87\xe6\x9c\xac\xef\xbc\x9a 1', 'utf8'),
                                testDocumentData['Chinese'][0]),
                               ('Arabic',
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Document \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Document \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9  1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Document \xd9\x85\xd8\xa4\xd9\x84\xd9\x81 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Document \xd9\x86\xd8\xb5 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                                testDocumentData['Arabic'][0])]

        for (language, noteName, noteComment, noteTaker, noteText, documentNum) in testDocumentNoteNames:
            if 210 in testsToRun:
                try:
                    # Try to load the Test Note
                    testNote1 = Note.Note(noteName, Document=documentNum)
                    # Delete the Test Note
                    testNote1.db_delete()
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    pass
                except sqlite3.OperationalError:
                    # This is the desired result!
                    pass
                except:

                    print "Exception in Test 210 -- Could not delete Note %s." % noteName
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    print

            if 220 in testsToRun:
                # Note Saving
                testName = 'Creating Document Note : %s' % noteName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Note
                    note1 = Note.Note()
                    # Populate it with data
                    note1.id = noteName
#                    note1.comment = noteComment
                    note1.author = noteTaker
                    note1.text = noteText
                    note1.document_num = documentNum
                    try:
                        note1.db_save()
                        # If we create the Note, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testNoteData[language]['Document'].append((note1.number, note1.id),)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Note, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Note, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 230 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(note1.number)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 235 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(noteName, Document=documentNum)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testQuoteNoteNames = [('English',
                               unicode('English Quote Note 1', 'utf8'),
                               unicode('English Quote Note Comment 1', 'utf8'),
                               unicode('English Quote Note NoteTaker 1', 'utf8'),
                               unicode('English Quote Note Text 1', 'utf8'),
                               testQuoteData['English'][0]),
                              ("Apostrophe's Test",
                               unicode("Apostrophe's Test Quote Note 1", 'utf8'),
                               unicode("Apostrophe's Test Quote Note Comment 1", 'utf8'),
                               unicode("Apostrophe's Test Quote Note NoteTaker 1", 'utf8'),
                               unicode("Apostrophe's Test Quote Note Text 1", 'utf8'),
                               testQuoteData["Apostrophe's Test"][0]),
                              ('French',
                               unicode('Fran\xc3\xa7ais Quote Note 1', 'utf8'),
                               unicode('Fran\xc3\xa7ais Quote Note Commentaire 1', 'utf8'),
                               unicode('Fran\xc3\xa7ais Quote Auteur-e de la note 1', 'utf8'),
                               unicode('Fran\xc3\xa7ais Quote Contenu de la note 1', 'utf8'),
                               testQuoteData['French'][0]),
                              ('Chinese',
                               unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Quote \xe5\xa4\x87\xe6\xb3\xa8 1', 'utf8'),
                               unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Quote \xe5\xa4\x87\xe6\xb3\xa8 \xe6\x89\xb9\xe6\xb3\xa8 1', 'utf8'),
                               unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Quote \xe5\xa4\x87\xe6\xb3\xa8\xe6\x8f\x90\xe5\x8f\x96\xe5\x99\xa8 1', 'utf8'),
                               unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93 Quote \xe5\xa4\x87\xe6\xb3\xa8\xe6\x96\x87\xe6\x9c\xac\xef\xbc\x9a 1', 'utf8'),
                               testQuoteData['Chinese'][0]),
                              ('Arabic',
                               unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Quote \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                               unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Quote \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9  1', 'utf8'),
                               unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Quote \xd9\x85\xd8\xa4\xd9\x84\xd9\x81 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                               unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Quote \xd9\x86\xd8\xb5 \xd8\xa7\xd9\x84\xd9\x85\xd9\x84\xd8\xa7\xd8\xad\xd8\xb8\xd8\xa9 1', 'utf8'),
                               testQuoteData['Arabic'][0])]

        for (language, noteName, noteComment, noteTaker, noteText, quoteNum) in testQuoteNoteNames:
            if 210 in testsToRun:
                try:
                    # Try to load the Test Note
                    testNote1 = Note.Note(noteName, Quote=quoteNum)
                    # Delete the Test Note
                    testNote1.db_delete()
                except TransanaExceptions.RecordNotFoundError:
                    # This is the desired result!
                    pass
                except sqlite3.OperationalError:
                    # This is the desired result!
                    pass
                except:

                    print "Exception in Test 210 -- Could not delete Note %s." % noteName
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    print

            if 220 in testsToRun:
                # Note Saving
                testName = 'Creating Quote Note : %s' % noteName
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                try:
                    # ... create a new Note
                    note1 = Note.Note()
                    # Populate it with data
                    note1.id = noteName
#                    note1.comment = noteComment
                    note1.author = noteTaker
                    note1.text = noteText
                    note1.quote_num = quoteNum
                    try:
                        note1.db_save()
                        # If we create the Note, consider the test passed.
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                        testNoteData[language]['Quote'].append((note1.number, note1.id),)
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If we can't create the Note, consider the test failed
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                except:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    # If we can't load or create the Note, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 230 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Number'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(note1.number)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

            if 235 in testsToRun:
                # Compare Created with Saved Note
                testName = 'Comparing Notes by Name'
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)
                
                note2 = Note.Note(noteName, Quote=quoteNum)

                if note1 == note2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testCoreDataNames = [('English',
                              unicode('Demo.mpg', 'utf8'),
                              'English Title',
                              'English Creator',
                              'English Subject',
                              'English Description',
                              'English Publisher',
                              'English Contributor',
                              '06/29/2014',
                              'Image',
                              'mpeg',
                              'English Source',
                              'English Language',
                              'English Relation',
                              'English Coverage',
                              'English Rights'),
                             ("Apostrophe's Test",
                              unicode('Demo2.mpg', 'utf8'),
                              "Apostrophe's Test Title",
                              "Apostrophe's Test Creator",
                              "Apostrophe's Test Subject",
                              "Apostrophe's Test Description",
                              "Apostrophe's Test Publisher",
                              "Apostrophe's Test Contributor",
                              '06/29/2014',
                              "Apostrophe's Media Type",
                              "Apostrophe's Format",
                              "Apostrophe's Test Source",
                              "Apostrophe's Language",
                              "Apostrophe's Test Relation",
                              "Apostrophe's Test Coverage",
                              "Apostrophe's Test Rights"),
                             (unicode("Fran\xc3\xa7ais Test", 'utf8'),
                              unicode('Test \xc3\xa2\xc3\xab\xc3\xac\xc3\xb3\xc3\xbb 5.mpg', 'utf8'),
                              unicode("Fran\xc3\xa7ais Titre", 'utf8'),
                              unicode("Fran\xc3\xa7ais Cr\xc3\xa9ateur", 'utf8'),
                              unicode("Fran\xc3\xa7ais Sujet", 'utf8'),
                              unicode("Fran\xc3\xa7ais Description", 'utf8'),
                              unicode("Fran\xc3\xa7ais \xc3\x89diteur", 'utf8'),
                              unicode("Fran\xc3\xa7ais Partenaires", 'utf8'),
                              '06/29/2014',
                              unicode('Fran\xc3\xa7ais Type de fichier audio ou vid\xc3\xa9o', 'utf8'),
                              unicode('Fran\xc3\xa7ais Format', 'utf8'),
                              unicode("Fran\xc3\xa7ais Source", 'utf8'),
                              unicode("Fran\xc3\xa7ais Langue", 'utf8'),
                              unicode("Fran\xc3\xa7ais Relation", 'utf8'),
                              unicode("Fran\xc3\xa7ais \xc3\x89tendue", 'utf8'),
                              unicode("Fran\xc3\xa7ais Droits", 'utf8')),
                             (unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 Test", 'utf8'),
                              unicode('\xe4\xba\xb2\xe4\xba\xb3\xe4\xba\xb2 5.mpg', 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe6\xa0\x87\xe9\xa2\x98", 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe5\x88\x9b\xe5\xbb\xba\xe8\x80\x85", 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe4\xb8\xbb\xe9\xa2\x98", 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe6\x8f\x8f\xe8\xbf\xb0", 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe5\x8f\x91\xe5\xb8\x83\xe8\x80\x85", 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe6\x8a\x95\xe7\xa8\xbf\xe4\xba\xba", 'utf8'),
                              '06/29/2014',
                              unicode('\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe5\xaa\x92\xe4\xbd\x93\xe7\xb1\xbb\xe5\x9e\x8b', 'utf8'),
                              unicode('\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe6\xa0\xbc\xe5\xbc\x8f', 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe6\x9d\xa5\xe6\xba\x90", 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe8\xaf\xad\xe8\xa8\x80", 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe5\x85\xb3\xe7\xb3\xbb", 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe5\xad\xa6\xe7\xa7\x91\xe8\x8c\x83\xe5\x9b\xb4", 'utf8'),
                              unicode("\xe4\xb8\xad\xe6\x96\x87-\xe7\xae\x80\xe4\xbd\x93 \xe7\x89\x88\xe6\x9d\x83\xe5\xa3\xb0\xe6\x98\x8e", 'utf8')),
                             (unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 Test", 'utf8'),
                              unicode('\xd8\xb4\xd9\x82\xd8\xb4\xd9\x84\xd8\xa7\xd9\x87\xd8\xa4.mpg', 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xb9\xd9\x86\xd9\x88\xd8\xa7\xd9\x86", 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xa4\xd9\x84\xd9\x81", 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd9\x88\xd8\xb6\xd9\x88\xd8\xb9", 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x88\xd8\xb5\xd9\x81", 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x86\xd8\xa7\xd8\xb4\xd8\xb1", 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xb3\xd8\xa7\xd9\x87\xd9\x85\xd9\x88\xd9\x86", 'utf8'),
                              '06/29/2014',
                              unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd9\x86\xd9\x88\xd8\xb9 \xd9\x85\xd9\x84\xd9\x81 \xd8\xa7\xd9\x84\xd9\x88\xd8\xb3\xd8\xa7\xd8\xa6\xd8\xb7', 'utf8'),
                              unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xaa\xd9\x86\xd8\xb3\xd9\x8a\xd9\x82', 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x85\xd8\xb5\xd8\xaf\xd8\xb1", 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd9\x84\xd8\xba\xd8\xa9", 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xb9\xd9\x84\xd8\xa7\xd9\x82\xd8\xa9", 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd9\x85\xd8\xaf\xd9\x89 \xd8\xa7\xd9\x84\xd8\xaa\xd8\xba\xd8\xb7\xd9\x8a\xd8\xa9", 'utf8'),
                              unicode("\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9 \xd8\xa7\xd9\x84\xd8\xad\xd9\x82\xd9\x88\xd9\x82", 'utf8'))]

        for (language, mediaFilename, title, creator, subject, description, publisher, contributor, dc_date, dc_type, format,
             source, language, relation, coverage, rights) in testCoreDataNames:
            
##            if 270 in testsToRun:
##                testName = 'Testing Core Data -- Deleting'
##                self.SetStatusText(testName)
##                self.testsRun += 1
##                self.txtCtrl.AppendText('Test "%s" ' % testName)
##
##                print "Deleting Core Data: ", language.encode('utf8'), mediaFilename.encode('utf8')
##
##                try:
##                    # Try to load the Test Core Data Record
##                    testCoreData1 = CoreData.CoreData(mediaFilename)
##                    # Delete the Test Core Data Record
##                    testCoreData1.db_delete()
##                    self.testsSuccessful += 1
##                except TransanaExceptions.RecordNotFoundError:
##                    # This is the desired result!
##                    self.testsSuccessful += 1
##                except:
##
##                    print sys.exc_info()[0]
##                    print sys.exc_info()[1]
##                    traceback.print_exc(file=sys.stdout)
##                    print
##                    print "Exception in Test 270 -- Could not delete CoreData %s." % mediaFilename.encode('utf8')
##                    print
##                    self.txtCtrl.AppendText('FAILED.')
##                    self.testsFailed += 1
##                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))
##
            if 280 in testsToRun:
                # Core Data Saving
                testName = 'Creating Core Data : %s' % mediaFilename.encode('utf8')
                self.SetStatusText(testName)
                self.testsRun += 1
                self.txtCtrl.AppendText('Test "%s" ' % testName)

                try:
                    # ... create a new Core Data record
                    coreData1 = CoreData.CoreData()
                    # Populate it with data
                    coreData1.id = mediaFilename
                    coreData1.title = title
                    coreData1.creator = creator
                    coreData1.subject = subject
                    coreData1.description = description
                    coreData1.publisher = publisher
                    coreData1.contributor = contributor
                    coreData1.dc_date = dc_date
                    coreData1.dc_type = dc_type
                    coreData1.format = format
                    coreData1.source = source
                    coreData1.language = language
                    coreData1.relation = relation
                    coreData1.coverage = coverage
                    coreData1.rights = rights
                    try:
                        coreData1.db_save()
                        # If we create the Core Data record, consider the test passed.
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
                    traceback.print_exc(file=sys.stdout)
                    
                    # If we can't load or create the Library, consider the test failed
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

##            if 290 in testsToRun:
##                # Compare Created with Saved Library
##                testName = 'Comparing Library by Number'
##                self.SetStatusText(testName)
##                self.testsRun += 1
##                self.txtCtrl.AppendText('Test "%s" ' % testName)
##                
##                series2 = Library.Library(series1.number)
##
##                if series1 == series2:
##                    self.txtCtrl.AppendText('Passed.')
##                    self.testsSuccessful += 1
##                else:
##                    self.txtCtrl.AppendText('FAILED.')
##                    self.testsFailed += 1
##                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 300 in testsToRun:
            testName = 'Testing DBInterface.list_of_series()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_series()

            comparison = []
            for key in testSeriesData.keys():
                comparison.append(testSeriesData[key])

            results.sort()
            comparison.sort()
            
            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 302 in testsToRun:
            testName = 'Testing DBInterface.list_of_documents()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_documents()

            comparison = []
            for key in testDocumentData.keys():
                data = testDocumentData[key] + (testSeriesData[key][0],)
                comparison.append(data)

            results.sort()
            comparison.sort()

            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 304 in testsToRun:

            testName = 'Testing DBInterface.list_of_documents() with libraryNum'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testSeriesData.keys():
                if key != 'Testing':

                    self.testsRun += 1

                    results = DBInterface.list_of_documents(int(testSeriesData[key][0]))

                    comparison = []

                    data = testDocumentData[key] + (testSeriesData[key][0],)
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 306 in testsToRun:
            testName = 'Testing DBInterface.dictionary_of_documents_and_episodes()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.dictionary_of_documents_and_episodes()

            comp = {}
            for key in testDocumentData.keys():
                comp[(testDocumentData[key][1], testDocumentData[key][0])] = ('Document', testSeriesData[key][0], testDocumentData[key][0])

            for key in testEpisodeData.keys():
                comp[(testEpisodeData[key][1], testEpisodeData[key][0])] = ('Episode', testSeriesData[key][0], testEpisodeData[key][0])

            resKeys = results.keys()
            resKeys.sort()
            compKeys = comp.keys()
            compKeys.sort()

            passed = (resKeys == compKeys)
            try:
                for key in resKeys:
                    if results[key] != comp[key]:
                        passed = (passed and False)
            except:
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                print
                import traceback
                traceback.print_exc(file=sys.stdout)
                print
                passed = False

            if passed:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 308 in testsToRun:

            testName = 'Testing DBInterface.dictionary_of_documents_and_episodes() with library'
            self.SetStatusText(testName)

            for key in testSeriesData.keys():
                if key != 'Testing':

                    self.testsRun += 1
                    self.txtCtrl.AppendText('Test "%s" ' % testName)

                    results = DBInterface.dictionary_of_documents_and_episodes(Library.Library(int(testSeriesData[key][0])))

                    comp = {}
                        
                    comp[(testDocumentData[key][1], testDocumentData[key][0])] = ('Document', testSeriesData[key][0], testDocumentData[key][0])

                    comp[(testEpisodeData[key][1], testEpisodeData[key][0])] = ('Episode', testSeriesData[key][0], testEpisodeData[key][0])

                    resKeys = results.keys()
                    resKeys.sort()
                    compKeys = comp.keys()
                    compKeys.sort()

                    passed = (resKeys == compKeys)
                    try:
                        for key in resKeys:
                            if results[key] != comp[key]:
                                passed = (passed and False)
                    except:
                        print sys.exc_info()[0]
                        print sys.exc_info()[1]
                        print
                        import traceback
                        traceback.print_exc(file=sys.stdout)
                        print
                        passed = False

                    if passed:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 310 in testsToRun:
            testName = 'Testing DBInterface.list_of_episodes()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_episodes()

            comparison = []
            for key in testEpisodeData.keys():
                data = testEpisodeData[key] + (testSeriesData[key][0],)
                comparison.append(data)

            results.sort()
            comparison.sort()

            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 320 in testsToRun:

            testName = 'Testing DBInterface.list_of_episodes_for_series()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testSeriesData.keys():
                if key != 'Testing':

                    self.testsRun += 1

                    results = DBInterface.list_of_episodes_for_series(testSeriesData[key][1])

                    comparison = []

                    data = testEpisodeData[key] + (testSeriesData[key][0],)
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 330 in testsToRun:

            testName = 'Testing DBInterface.list_of_episode_transcripts()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_episode_transcripts()

            comparison = []

            for key in testTranscriptData.keys():
                data = testTranscriptData[key] + (testEpisodeData[key][0],)
                comparison.append(data)

            results.sort()
            comparison.sort()

            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 340 in testsToRun:

            testName = 'Testing DBInterface.list_transcripts()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testEpisodeData.keys():

                self.testsRun += 1

                results = DBInterface.list_transcripts(testSeriesData[key][1], testEpisodeData[key][1])

                comparison = []

                data = testTranscriptData[key] + (testEpisodeData[key][0],)
                comparison.append(data)

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 350 in testsToRun:

            testName = 'Testing DBInterface.list_clip_transcripts()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testClipData.keys():

                self.testsRun += 1
                
                results = DBInterface.list_clip_transcripts(testClipData[key][0])

                comparison = []

                data = (testClipData[key][2:4] + (testClipData[key][-1],))

                comparison.append(data)

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 360 in testsToRun:
            testName = 'Testing DBInterface.list_of_collections()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_collections(0)

            comparison = []
            for key in testCollectionData.keys():
                comparison.append(testCollectionData[key] + (0,))

            results.sort()
            comparison.sort()
            
            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 370 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_collections()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_collections()

            comparison = []
            for key in testCollectionData.keys():
                comparison.append(testCollectionData[key] + (0,))

            results.sort()
            comparison.sort()
            
            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 373 in testsToRun:
            testName = 'Testing DBInterface.list_of_quotes()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_quotes()

            # Strip Testing Quotes from Results!
            results2 = []
            for res in results:
                if not u'Testing Quote' in res[1]:
                    results2.append(res)

            comparison = []
            for key in testQuoteData.keys():

                if key != 'Testing':
                    comparison.append(testQuoteData[key][ : 2] + (testCollectionData[key][0], testQuoteData[key][3], 2L))
                
            results2.sort()
            comparison.sort()
            
            if results2 == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:

                print
                print testName, key
                print
                print results2
                print
                print comparison
                print
                
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 375 in testsToRun:

            testName = 'Testing DBInterface.list_of_quotes_by_document()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testDocumentData.keys():

                if key != 'Testing':
                    self.testsRun += 1

                    results = DBInterface.list_of_quotes_by_document(testDocumentData[key][0])

                    comparison1 = []
                    
                    for tmpQuote in results:
                        data = (tmpQuote['QuoteNum'], tmpQuote['QuoteID'], tmpQuote['CollectNum'])
                        comparison1.append(data)

                    comparison2 = []

                    data = testQuoteData[key][ : 2] + (testCollectionData[key][0],)
                    comparison2.append(data)

                    comparison1.sort()
                    comparison2.sort()

                    if comparison1 == comparison2:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:

                        print
                        print testName, key
                        print
                        print comparison1
                        print
                        print comparison2
                        print
                        
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 377 in testsToRun:

            testName = 'Testing DBInterface.list_of_quotes_by_collectionnum()'
            self.SetStatusText(testName)
#            self.txtCtrl.AppendText('Test "%s:" \n' % testName)

            for key in testCollectionData.keys():

                if key != 'Testing':

                    self.txtCtrl.AppendText('Test "%s: %s" ' % (testName, key))
                    self.testsRun += 1

                    results = DBInterface.list_of_quotes_by_collectionnum(testCollectionData[key][0])

                    comparison = []

                    data = testQuoteData[key][ : 2] + (testCollectionData[key][0],)
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:

                        print
                        print testName, key
                        print
                        print results
                        print
                        print comparison
                        print
                        
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 380 in testsToRun:
            testName = 'Testing DBInterface.list_of_clips()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_clips()

            comparison = []

            for key in testClipData.keys():
                comparison.append(testClipData[key][ : 2] + (testCollectionData[key][0], testClipData[key][4], 1L))

            results.sort()
            comparison.sort()
            
            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:

                print
                print testName
                print
                print results
                print
                print comparison
                print
                
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 390 in testsToRun:

            testName = 'Testing DBInterface.list_of_clips_by_collection()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testCollectionData.keys():

                if key != 'Testing':

                    self.testsRun += 1

                    results = DBInterface.list_of_clips_by_collection(testCollectionData[key][1], 0)

                    comparison = []

                    data = testClipData[key][ : 2] + (testCollectionData[key][0],)
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 400 in testsToRun:

            testName = 'Testing DBInterface.list_of_clips_by_collectionnum()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testCollectionData.keys():

                if key != 'Testing':

                    self.testsRun += 1

                    results = DBInterface.list_of_clips_by_collectionnum(testCollectionData[key][0])

                    comparison = []

                    data = testClipData[key][ : 2] + (testCollectionData[key][0],)
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 410 in testsToRun:

            testName = 'Testing DBInterface.list_of_clips_by_episode()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testEpisodeData.keys():

                self.testsRun += 1

                results = DBInterface.list_of_clips_by_episode(testEpisodeData[key][0])

                comparison1 = []
                
                for tmpClip in results:
                    data = (tmpClip['ClipNum'], tmpClip['ClipID'], tmpClip['CollectNum'])
                    comparison1.append(data)

                comparison2 = []

                data = testClipData[key][ : 2] + (testCollectionData[key][0],)
                comparison2.append(data)

                comparison1.sort()
                comparison2.sort()

                if comparison1 == comparison2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 420 in testsToRun:

            testName = 'Testing DBInterface.list_of_clips_by_transcriptnum()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testTranscriptData.keys():

                self.testsRun += 1

                results = DBInterface.list_of_clips_by_transcriptnum(testTranscriptData[key][0])

                comparison1 = []
                
                for tmpClip in results:
                    data = (tmpClip['ClipNum'], tmpClip['ClipID'], tmpClip['CollectNum'], tmpClip['TranscriptNum'])
                    comparison1.append(data)

                comparison2 = []

                data = testClipData[key][ : 2] + (testCollectionData[key][0],) + (testTranscriptData[key][0],)
                comparison2.append(data)

                comparison1.sort()
                comparison2.sort()

                if comparison1 == comparison2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 430 in testsToRun:
            testName = 'Testing DBInterface.list_of_snapshots()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_snapshots()

            comparison = []
            for key in testSnapshotData.keys():

                comparison.append(testSnapshotData[key] + (testCollectionData[key][0], numTestQuotes + 1))

            results.sort()
            comparison.sort()

            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:

                print
                print testName
                print
                print results
                print
                print comparison
                print
                
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 440 in testsToRun:

            testName = 'Testing DBInterface.list_of_snapshots_by_episode()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testEpisodeData.keys():

                self.testsRun += 1

                results = DBInterface.list_of_snapshots_by_episode(testEpisodeData[key][0])

                comparison1 = []
                
                for tmpSnapshot in results:
                    data = (tmpSnapshot['SnapshotNum'], tmpSnapshot['SnapshotID'], tmpSnapshot['CollectNum'])
                    comparison1.append(data)

                comparison2 = []

                data = testSnapshotData[key][ : 2] + (testCollectionData[key][0],)
                comparison2.append(data)

                comparison1.sort()
                comparison2.sort()

                if comparison1 == comparison2:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 450 in testsToRun:

            testName = 'Testing DBInterface.list_of_snapshots_by_transcriptnum()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testTranscriptData.keys():

                self.testsRun += 1

                results = DBInterface.list_of_snapshots_by_transcriptnum(testTranscriptData[key][0])

                comparison = []

                data = testSnapshotData[key][ : 2] + (testCollectionData[key][0],)
                comparison.append(data)

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 460 in testsToRun:

            testName = 'Testing DBInterface.list_of_snapshots_by_collectionnum()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testCollectionData.keys():

                if key != 'Testing':

                    self.testsRun += 1

                    results = DBInterface.list_of_snapshots_by_collectionnum(testCollectionData[key][0])

                    comparison = []

                    data = testSnapshotData[key][ : 2] + (testCollectionData[key][0],)
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 470 in testsToRun:
            testName = 'Testing DBInterface.list_of_notes(Library)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testSeriesData.keys():

                if key != 'Testing':
                    self.testsRun += 1

                    results = DBInterface.list_of_notes(Library=testSeriesData[key][0], includeNumber=True)

                    comparison = []

                    data = testNoteData[key]['Libraries'][0]
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 472 in testsToRun:
            testName = 'Testing DBInterface.list_of_notes(Episode)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testEpisodeData.keys():

                self.testsRun += 1

                results = DBInterface.list_of_notes(Episode=testEpisodeData[key][0], includeNumber=True)

                comparison = []

                data = testNoteData[key]['Episode'][0]
                comparison.append(data)

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 474 in testsToRun:
            testName = 'Testing DBInterface.list_of_notes(Transcript)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testTranscriptData.keys():

                self.testsRun += 1

                results = DBInterface.list_of_notes(Transcript=testTranscriptData[key][0], includeNumber=True)

                comparison = []

                data = testNoteData[key]['Transcript'][0]
                comparison.append(data)

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 476 in testsToRun:
            testName = 'Testing DBInterface.list_of_notes(Collection)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testCollectionData.keys():

                if key != 'Testing':

                    self.testsRun += 1

                    results = DBInterface.list_of_notes(Collection=testCollectionData[key][0], includeNumber=True)

                    comparison = []

                    data = testNoteData[key]['Collection'][0]
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 478 in testsToRun:
            testName = 'Testing DBInterface.list_of_notes(Clip)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testClipData.keys():

                self.testsRun += 1

                results = DBInterface.list_of_notes(Clip=testClipData[key][0], includeNumber=True)

                comparison = []

                data = testNoteData[key]['Clip'][0]
                comparison.append(data)

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 480 in testsToRun:
            testName = 'Testing DBInterface.list_of_notes(Snapshot)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testSnapshotData.keys():

                self.testsRun += 1

                results = DBInterface.list_of_notes(Snapshot=testSnapshotData[key][0], includeNumber=True)

                comparison = []

                data = testNoteData[key]['Snapshot'][0]
                comparison.append(data)

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 485 in testsToRun:
            testName = 'Testing DBInterface.list_of_notes(Document)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testDocumentData.keys():

                if key != 'Testing':

                    self.testsRun += 1

                    results = DBInterface.list_of_notes(Document=testDocumentData[key][0], includeNumber=True)

                    comparison = []

                    data = testNoteData[key]['Document'][0]
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 486 in testsToRun:
            testName = 'Testing DBInterface.list_of_notes(Quote)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testQuoteData.keys():

                if key != 'Testing':

                    self.testsRun += 1

                    results = DBInterface.list_of_notes(Quote=testQuoteData[key][0], includeNumber=True)

                    comparison = []

                    data = testNoteData[key]['Quote'][0]
                    comparison.append(data)

                    results.sort()
                    comparison.sort()

                    if results == comparison:
                        self.txtCtrl.AppendText('Passed.')
                        self.testsSuccessful += 1
                    else:
                        self.txtCtrl.AppendText('FAILED.')
                        self.testsFailed += 1
                    self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 490 in testsToRun:
            testName = 'Testing DBInterface.list_of_node_notes(LibraryNode=True)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_node_notes(LibraryNode=True)

            comparison = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Libraries'][0] + (testSeriesData[language][0], 0, 0, 0, 0, 0, 0, 0)
                comparison.append(data)
                data = testNoteData[language]['Episode'][0] + (0, testEpisodeData[language][0], 0, 0, 0, 0, 0, 0)
                comparison.append(data)
                data = testNoteData[language]['Transcript'][0] + (0, 0, testTranscriptData[language][0], 0, 0, 0, 0, 0)
                comparison.append(data)
                data = testNoteData[language]['Document'][0] + (0, 0, 0, 0, 0, 0, testDocumentData[language][0], 0)
                comparison.append(data)

            results.sort()
            comparison.sort()

            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 500 in testsToRun:
            testName = 'Testing DBInterface.list_of_node_notes(CollectionNode=True)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_node_notes(CollectionNode=True)

            comparison = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Collection'][0] + (0, 0, 0, testCollectionData[language][0], 0, 0, 0, 0)
                comparison.append(data)
                data = testNoteData[language]['Clip'][0] + (0, 0, 0, 0, testClipData[language][0], 0, 0, 0)
                comparison.append(data)
                data = testNoteData[language]['Snapshot'][0] + (0, 0, 0, 0, 0, testSnapshotData[language][0], 0, 0)
                comparison.append(data)
                data = testNoteData[language]['Quote'][0] + (0, 0, 0, 0, 0, 0, 0, testQuoteData[language][0])
                comparison.append(data)

            results.sort()
            comparison.sort()

            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 510 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(No Scope, No Text Search)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes()

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Libraries'][0] + (testSeriesData[language][0], 0, 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Episode'][0] + (0, testEpisodeData[language][0], 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Transcript'][0] + (0, 0, testTranscriptData[language][0], 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Collection'][0] + (0, 0, 0, testCollectionData[language][0], 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Clip'][0] + (0, 0, 0, 0, testClipData[language][0], 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Snapshot'][0] + (0, 0, 0, 0, 0, testSnapshotData[language][0], 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Document'][0] + (0, 0, 0, 0, 0, 0, testDocumentData[language][0], 0)
                comparison2.append(data)
                data = testNoteData[language]['Quote'][0] + (0, 0, 0, 0, 0, 0, 0, testQuoteData[language][0])
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 520 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(Library Scope, No Text Search)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(reportType='LibraryNode')

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Libraries'][0] + (testSeriesData[language][0], 0, 0, 0, 0, 0, 0, 0)
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 530 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(Episode Scope, No Text Search)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(reportType='EpisodeNode')

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Episode'][0] + (0, testEpisodeData[language][0], 0, 0, 0, 0, 0, 0)
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 540 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(Transcript Scope, No Text Search)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(reportType='TranscriptNode')

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Transcript'][0] + (0, 0, testTranscriptData[language][0], 0, 0, 0, 0, 0)
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 550 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(Collection Scope, No Text Search)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(reportType='CollectionNode')

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Collection'][0] + (0, 0, 0, testCollectionData[language][0], 0, 0, 0, 0)
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 560 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(Clip Scope, No Text Search)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(reportType='ClipNode')

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Clip'][0] + (0, 0, 0, 0, testClipData[language][0], 0, 0, 0)
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 570 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(Snapshot Scope, No Text Search)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(reportType='SnapshotNode')

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Snapshot'][0] + (0, 0, 0, 0, 0, testSnapshotData[language][0], 0, 0)
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 573 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(Document Scope, No Text Search)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(reportType='DocumentNode')

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Document'][0] + (0, 0, 0, 0, 0, 0, testDocumentData[language][0], 0)
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 574 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(Quote Scope, No Text Search)'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(reportType='QuoteNode')

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English', "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Quote'][0] + (0, 0, 0, 0, 0, 0, 0, testQuoteData[language][0])
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 580 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_notes(No Scope, Text Search for "English")'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(searchText='English')

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['English']:  # , "Apostrophe's Test", 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Libraries'][0] + (testSeriesData[language][0], 0, 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Episode'][0] + (0, testEpisodeData[language][0], 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Transcript'][0] + (0, 0, testTranscriptData[language][0], 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Collection'][0] + (0, 0, 0, testCollectionData[language][0], 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Clip'][0] + (0, 0, 0, 0, testClipData[language][0], 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Snapshot'][0] + (0, 0, 0, 0, 0, testSnapshotData[language][0], 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Document'][0] + (0, 0, 0, 0, 0, 0, testDocumentData[language][0], 0)
                comparison2.append(data)
                data = testNoteData[language]['Quote'][0] + (0, 0, 0, 0, 0, 0, 0, testQuoteData[language][0])
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:

                print '------------------------------------------'
                print comparison1
                print
                print comparison2
                print '------------------------------------------'
                for x in range(len(comparison1)):
                    print x, comparison1[x] == comparison2[x]
                    if not comparison1[x] == comparison2[x]:
                        print comparison1[x]
                        print comparison2[x]
                
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 590 in testsToRun:

            searchText = unicode('Fran\xc3\xa7ais', 'utf8')

            testName = u'Testing DBInterface.list_of_all_notes(No Scope, Text Search for "%s")' % searchText
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(searchText=searchText)

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['French']:  #  ['English', 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Libraries'][0] + (testSeriesData[language][0], 0, 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Episode'][0] + (0, testEpisodeData[language][0], 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Transcript'][0] + (0, 0, testTranscriptData[language][0], 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Collection'][0] + (0, 0, 0, testCollectionData[language][0], 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Clip'][0] + (0, 0, 0, 0, testClipData[language][0], 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Snapshot'][0] + (0, 0, 0, 0, 0, testSnapshotData[language][0], 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Document'][0] + (0, 0, 0, 0, 0, 0, testDocumentData[language][0], 0)
                comparison2.append(data)
                data = testNoteData[language]['Quote'][0] + (0, 0, 0, 0, 0, 0, 0, testQuoteData[language][0])
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 600 in testsToRun:

            searchText = unicode('\xe4\xb8\xad\xe6\x96\x87\x2d\xe7\xae\x80\xe4\xbd\x93', 'utf8')

            testName = u'Testing DBInterface.list_of_all_notes(No Scope, Text Search for "%s")' % searchText
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(searchText=searchText)

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['Chinese']:  #  ['English', 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Libraries'][0] + (testSeriesData[language][0], 0, 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Episode'][0] + (0, testEpisodeData[language][0], 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Transcript'][0] + (0, 0, testTranscriptData[language][0], 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Collection'][0] + (0, 0, 0, testCollectionData[language][0], 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Clip'][0] + (0, 0, 0, 0, testClipData[language][0], 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Snapshot'][0] + (0, 0, 0, 0, 0, testSnapshotData[language][0], 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Document'][0] + (0, 0, 0, 0, 0, 0, testDocumentData[language][0], 0)
                comparison2.append(data)
                data = testNoteData[language]['Quote'][0] + (0, 0, 0, 0, 0, 0, 0, testQuoteData[language][0])
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 610 in testsToRun:

            searchText = unicode('\xd8\xa7\xd9\x84\xd8\xb9\xd8\xb1\xd8\xa8\xd9\x8a\xd8\xa9', 'utf8')

            testName = u'Testing DBInterface.list_of_all_notes(No Scope, Text Search for "%s")' % searchText
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_notes(searchText=searchText)

            comparison1 = []
            for rec in results:
                data = (rec['NoteNum'], rec['NoteID'], rec['SeriesNum'], rec['EpisodeNum'], rec['TranscriptNum'], rec['CollectNum'], rec['ClipNum'], rec['SnapshotNum'], rec['DocumentNum'], rec['QuoteNum'])
                comparison1.append(data)

            comparison2 = []

            for language in ['Arabic']:  #  ['English', 'French', 'Chinese', 'Arabic']:

                data = testNoteData[language]['Libraries'][0] + (testSeriesData[language][0], 0, 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Episode'][0] + (0, testEpisodeData[language][0], 0, 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Transcript'][0] + (0, 0, testTranscriptData[language][0], 0, 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Collection'][0] + (0, 0, 0, testCollectionData[language][0], 0, 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Clip'][0] + (0, 0, 0, 0, testClipData[language][0], 0, 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Snapshot'][0] + (0, 0, 0, 0, 0, testSnapshotData[language][0], 0, 0)
                comparison2.append(data)
                data = testNoteData[language]['Document'][0] + (0, 0, 0, 0, 0, 0, testDocumentData[language][0], 0)
                comparison2.append(data)
                data = testNoteData[language]['Quote'][0] + (0, 0, 0, 0, 0, 0, 0, testQuoteData[language][0])
                comparison2.append(data)

            comparison1.sort()
            comparison2.sort()

            if comparison1 == comparison2:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 620 in testsToRun:
            testName = 'Testing DBInterface.list_of_keyword_groups()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_keyword_groups()

            comparison = []
            for (kwg, kw, kwdef) in testKeywordPairs:
                comparison.append(kwg)

            results.sort()
            comparison.sort()

            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 630 in testsToRun:
            testName = 'Testing DBInterface.list_of_keywords_by_group()'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for (kwg, kw, kwdef) in testKeywordPairs:

                self.testsRun += 1
                
                results = DBInterface.list_of_keywords_by_group(kwg)

                comparison = []
                comparison.append(kw)

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 640 in testsToRun:
            testName = 'Testing DBInterface.list_of_all_keywords()'
            self.SetStatusText(testName)
            self.testsRun += 1
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            results = DBInterface.list_of_all_keywords()

            comparison = []
            for (kwg, kw, kwdef) in testKeywordPairs:
                comparison.append((kwg, kw),)

            results.sort()
            comparison.sort()

            if results == comparison:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1
            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        testKeywordNames = {}
        testKeywordNames['English'] = testKeywordPairs[0]
        testKeywordNames["Apostrophe's Test"] = testKeywordPairs[1]
        testKeywordNames['French'] = testKeywordPairs[2]
        testKeywordNames['Chinese'] = testKeywordPairs[3]
        testKeywordNames['Arabic'] = testKeywordPairs[4]
        
        if 650 in testsToRun:
            testName = 'Testing DBInterface.list_of_keywords(Document)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testEpisodeData.keys():

                self.testsRun += 1
                
                results = DBInterface.list_of_keywords(Document=testDocumentData[key][0])

                comparison = []
                comparison.append(testKeywordNames[key][:2] + (u'0',))

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 655 in testsToRun:
            testName = 'Testing DBInterface.list_of_keywords(Episode)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testEpisodeData.keys():

                self.testsRun += 1
                
                results = DBInterface.list_of_keywords(Episode=testEpisodeData[key][0])

                comparison = []
                comparison.append(testKeywordNames[key][:2] + (u'0',))

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 660 in testsToRun:
            testName = 'Testing DBInterface.list_of_keywords(Quote)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testClipData.keys():

                self.testsRun += 1
                
                results = DBInterface.list_of_keywords(Quote=testQuoteData[key][0])

                comparison = []
                comparison.append(testKeywordNames[key][:2] + (u'0',))

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 665 in testsToRun:
            testName = 'Testing DBInterface.list_of_keywords(Clip)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testClipData.keys():

                self.testsRun += 1
                
                results = DBInterface.list_of_keywords(Clip=testClipData[key][0])

                comparison = []
                comparison.append(testKeywordNames[key][:2] + (u'0',))

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 670 in testsToRun:
            testName = 'Testing DBInterface.list_of_keywords(Snapshot)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testSnapshotData.keys():

                self.testsRun += 1
                
                results = DBInterface.list_of_keywords(Snapshot=testSnapshotData[key][0])

                comparison = []
                comparison.append(testKeywordNames[key][:2] + (u'0',))

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 680 in testsToRun:
            testName = 'Testing DBInterface.list_of_snapshot_detail_keywords(Snapshot)'
            self.SetStatusText(testName)
            self.txtCtrl.AppendText('Test "%s" ' % testName)

            for key in testSnapshotData.keys():

                self.testsRun += 1
                
                results = DBInterface.list_of_snapshot_detail_keywords(Snapshot=testSnapshotData[key][0])

                comparison = []
                comparison.append(testKeywordNames[key][:2])

                results.sort()
                comparison.sort()

                if results == comparison:
                    self.txtCtrl.AppendText('Passed.')
                    self.testsSuccessful += 1
                else:
                    self.txtCtrl.AppendText('FAILED.')
                    self.testsFailed += 1
                self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))


        self.txtCtrl.AppendText('All tests completed.')
        self.txtCtrl.AppendText('\nFinal Summary:  Total Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        dbCursor.close()
        db.close()
        db = None

    def ShowMessage(self, msg):
        dlg = Dialogs.InfoDialog(self, msg)
        dlg.ShowModal()
        dlg.Destroy()

    def ConfirmTest(self, msg, testName):
        dlg = Dialogs.QuestionDialog(self, msg, "Unit Test 2: Database", noDefault=True)
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

    def CheckTable(self, table, results):
        self.testsRun += 1
        self.txtCtrl.AppendText('%s exists:  ' % table)
        if ((table,) in results) or ((table.lower(),) in results):
            self.txtCtrl.AppendText('Passed.\n')
            self.testsSuccessful += 1
        else:
            self.txtCtrl.AppendText('FAILED.\n')
            self.testsFailed += 1


class DBLogin:
    def __init__(self, username, password, dbServer, databaseName, port):
        self.username = username
        self.password = password
        self.dbServer = dbServer
        self.databaseName = databaseName
        self.port = port
        self.ssl = True
        if 'wxMSW' in wx.PlatformInfo:
            self.sslClientCert = 'C:\\Users\\DavidWoods\\Documents\\SSL\\DKW-Linux\\DKWLinux-client-cert.pem'
            self.sslClientKey = 'C:\\Users\\DavidWoods\\Documents\\SSL\\DKW-Linux\\DKWLinux-client-key.pem'
            self.sslMsgSrvCert = 'C:\\Users\\DavidWoods\\Documents\\SSL\\DKW-Linux\\DKWLinux-TransanaMessageServer-cert.pem'
        elif 'wxMac' in wx.PlatformInfo:
            self.sslClientCert = '/Users/davidwoods/Transana 2/SSL/DKW-Linux/DKWLinux-client-cert.pem'
            self.sslClientKey = '/Users/davidwoods/Transana 2/SSL/DKW-Linux/DKWLinux-client-key.pem'
            self.sslMsgSrvCert = '/Users/davidwoods/Transana 2/SSL/DKW-Linux/DKWLinux-TransanaMessageServer-cert.pem'

class MyApp(wx.App):
   def OnInit(self):
      frame = DBTest(None, -1, "Unit Test 2: Database")
      self.SetTopWindow(frame)
      return True
      

app = MyApp(0)
app.MainLoop()
