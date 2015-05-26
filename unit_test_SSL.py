import datetime
import os
import sys
import wx
import traceback

import gettext

# This module expects i18n.  Enable it here.
__builtins__._ = wx.GetTranslation


import TransanaConstants

SINGLE_USER = TransanaConstants.singleUserVersion

import ConfigData
import ControlObjectClass
import DBInterface
import MenuWindow
import TransanaGlobal

if TransanaConstants.DBInstalled in ['sqlite3']:
    import sqlite3


MENU_FILE_EXIT = 101

class UnitTest(wx.Frame):
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
        TransanaGlobal.configData.videoPath = 'C:\\Users\\DavidWoods\\Videos'

        TransanaGlobal.menuWindow = MenuWindow.MenuWindow(None, -1, title)

        self.ControlObject = ControlObjectClass.ControlObject()
        self.ControlObject.Register(Menu = TransanaGlobal.menuWindow)
        TransanaGlobal.menuWindow.Register(ControlObject = self.ControlObject)

        wx.Frame.__init__(self,parent,-1, title, size = (800,600), style=wx.DEFAULT_FRAME_STYLE|wx.NO_FULL_REPAINT_ON_RESIZE)
        self.testsRun = 0
        self.testsSuccessful = 0
        self.testsFailed = 0

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        self.txtCtrl = wx.TextCtrl(self, -1, "Unit Test:  SSL\n\n", style=wx.TE_LEFT | wx.TE_MULTILINE)
        self.txtCtrl.AppendText("Transana Version:  %s\nsingleUserVersion:  %s\n" % (TransanaConstants.versionNumber, TransanaConstants.singleUserVersion))

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
            traceback.print_exc(file=sys.stdout)
            

        if TransanaConstants.DBInstalled in ['MySQLdb-embedded']:
            DBInterface.EndSingleUserDatabase()
            
        TransanaGlobal.menuWindow.Destroy()

    def RunTests(self):

        # Basic Information
        self.txtCtrl.AppendText('Python Version:  %s\n' % sys.version[:5])
        self.txtCtrl.AppendText('wxPython Version:  %s\n' % wx.VERSION_STRING)
        self.txtCtrl.AppendText('DBInstalled:  %s\n' % TransanaConstants.DBInstalled)
        if TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server']:
            import MySQLdb
            str = 'MySQL for Python:  %s\n' % (MySQLdb.__version__)
        elif TransanaConstants.DBInstalled in ['PyMySQL']:
            import pymysql
            str = 'PyMySQL:  %s\n' % (pymysql.version_info)
        elif TransanaConstants.DBInstalled in ['sqlite3']:
            import sqlite3
            str = 'sqlite:  %s\n' % (sqlite3.version)
        else:
            str = 'Unknown Database:  Unknown Version\n'
        self.txtCtrl.AppendText(str)

        self.txtCtrl.AppendText('\n')

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
        
        # Tests defined:
        testsNotToSkip = []
        startAtTest = 1  # Should start at 1, not 0!
        endAtTest = 1000   # Should be one more than the last test to be run!
        testsToRun = testsNotToSkip + range(startAtTest, endAtTest)

        t = datetime.datetime.now()

        if 10 in testsToRun:
            # Is Database Open, NO SSL
            
            TransanaGlobal.configData.ssl = False

            loginInfo = DBLogin('DavidW', password, 'DKW-Linux', 'Transana_UnitTest', '3306', TransanaGlobal.configData.ssl)
            
            dbReference = DBInterface.get_db(loginInfo)
            DBInterface.establish_db_exists()

            testName = 'Open Database - No SSL'
            self.txtCtrl.AppendText('Test "%s" ' % testName)
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
            # Close Database, NO SSL

            testName = 'Close Database - No SSL'
            self.txtCtrl.AppendText('Test "%s" ' % testName)
            self.SetStatusText(testName)

            DBInterface.close_db()
            result = DBInterface.is_db_open()
            self.testsRun += 1

            if not result:
                self.txtCtrl.AppendText('Passed.')
                self.testsSuccessful += 1
            else:
                self.txtCtrl.AppendText('FAILED.')
                self.testsFailed += 1

            self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))

        if 30 in testsToRun:
            # Is Database Open, SSL
            
            TransanaGlobal.configData.ssl = True

            loginInfo = DBLogin('DavidW', password, 'DKW-Linux', 'Transana_UnitTest', '3306', TransanaGlobal.configData.ssl)
            dbReference = DBInterface.get_db(loginInfo)
            DBInterface.establish_db_exists()

            testName = 'Open Database - SSL'
            self.txtCtrl.AppendText('Test "%s" ' % testName)
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








        self.txtCtrl.AppendText('All tests completed.')
        self.txtCtrl.AppendText('\nFinal Summary:  Total Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))


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


class DBLogin:
    def __init__(self, username, password, dbServer, databaseName, port, ssl):
        self.username = username
        self.password = password
        self.dbServer = dbServer
        self.databaseName = databaseName
        self.port = port
        self.ssl = ssl
        if ssl:
            if sys.platform == 'win32':
                self.sslClientCert = 'C:\\Users\\DavidWoods\\Documents\\Transana 2\\SSL\\DKW-Linux\\DKWLinux-client-cert.pem'
                self.sslClientKey = 'C:\\Users\\DavidWoods\\Documents\\Transana 2\\SSL\\DKW-Linux\\DKWLinux-client-key.pem'
                self.sslMsgSrvCert = 'C:\\Users\DavidWoods\\Documents\\Transana 2\\SSL\\DKW-Linux\\DKWLinux-MessageServer-Cert.pem'
            elif sys.platform == 'darwin':
                self.sslClientCert = '/Users/davidwoods/Transana 2/SSL/DKW-Linux/DKWLinux-client-cert.pem'
                self.sslClientKey = '/Users/davidwoods/Transana 2/SSL/DKW-Linux/DKWLinux-client-key.pem'
                self.sslMsgSrvCert = ''
            else:

                print '*******************************************'
                print '*          SSL Files not defined          *'
                print '*******************************************'
                
        else:
            self.sslClientCert = ''
            self.sslClientKey = ''
            self.sslMsgSrvCert = ''
            

class MyApp(wx.App):
   def OnInit(self):
      frame = UnitTest(None, -1, "Unit Test: SSL")
      self.SetTopWindow(frame)
      return True
      

app = MyApp(0)
app.MainLoop()
