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

"""This file implements the Transana class, which is the main Transana application
definition."""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

import sys                          # import Python's sys module
import wx                           # import wxPython's wxWindows implementation
import os
import gettext                      # localization module
if __name__ == '__main__':
    # Define the "_" method, pointing it to wxPython's GetTranslation method
    __builtins__._ = wx.GetTranslation
from TransanaExceptions import *    # import all exception classes
import Dialogs                      # import Transana Error Dialog
import TransanaConstants            # import the Transana Constants
import TransanaGlobal               # import Transana's Global Variables
from ControlObjectClass import ControlObject   # import the Transana Control Object
if "__WXMAC__" in wx.PlatformInfo:
    import MacOS

import DBInterface                  # import the Database Interface module
import time                         # import the time module (Python)
# import Transana's ConfigData module
import ConfigData
import pickle                       # import Python's pickle module

DEBUG = False
if DEBUG:
    print "Transana DEBUG is ON!!"
    print
    print "wxPython version loaded: ", wx.VERSION_STRING,
    if 'unicode' in wx.PlatformInfo:
        print "- unicode"
    else:
        print "- ansi"
    print


class Transana(wx.App):
    """This class contains the main Transana application definition and the 
    logic that instantiates all other objects."""
    
    def OnInit(self):
        """ Initialize the application """
        # If we have a Workshop Version, see if the version is still valid, report to the user and exit if not.
        # NOTE:  This code just provides a user message and clean exit.  Transana will still be disabled if this
        # code is removed.  (This code is user-accessible on OS X!)
        if TransanaConstants.workshopVersion:
            import datetime
            t1 = TransanaConstants.startdate
            t2 = TransanaConstants.expirationdate
            t3 = time.localtime()
            t4 = datetime.datetime(t3[0], t3[1], t3[2], t3[3], t3[4])
            if (t1 > t4) or (t2 < t4):
                dlg = Dialogs.ErrorDialog(None, "This copy of Transana is no longer valid.")
                dlg.ShowModal()
                dlg.Destroy
                return False
        # In wxPython, you used to be able to initialize the ConfigData object in the TransanaGlobal module, even though the
        # wx.App() hadn't been initialized yet.  Moving from wxPython 2.8.1 to wxPython 2.8.9, this stopped being true
        # at least on the Mac.  Therefore, we moved creation of the ConfigData object to here in the code.
        # However, there are MANY references to ConfigData being in the TransanaGlobal module, so that's where we'll keep it.
        TransanaGlobal.configData = ConfigData.ConfigData()
        # If we are running from a BUILD instead of source code ...
        if hasattr(sys, "frozen") or ('wxMac' in wx.PlatformInfo):
            # See if the Default Profile Path exists.  If not ...
            if not os.path.exists(TransanaGlobal.configData.GetDefaultProfilePath()):
                # ... then create it (recursively).
                os.makedirs(TransanaGlobal.configData.GetDefaultProfilePath())
            # Build the path for the error log
            path = os.path.join(TransanaGlobal.configData.GetDefaultProfilePath(), 'Transana_Error.log')
            # redirect output to the error log
            self.RedirectStdio(filename=path)
            # Put a startup indicator in the Error Log
            print "Transana started:", time.asctime()

        # If no Language is defined ...
        if TransanaGlobal.configData.language == '':
            # ... then we know this is the first startup for this user profile.  Remember that!
            firstStartup = True
        # If language is known ...
        else:
            # ... then it's NOT first time startup.
            firstStartup = False

        # Now that we've loaded the Configuration Data, we can see if we need to alter the default encoding
        # If we're on Windows, single-user, using Russian, use KOI8r encoding instead of Latin-1,
        # Chinese uses big5, Japanese uses cp932, and Korean uses cp949
##        if ('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion):
##            if (TransanaGlobal.configData.language == 'ru'):
##                TransanaGlobal.encoding = 'koi8_r'
##            elif (TransanaGlobal.configData.language == 'zh'):
##                TransanaGlobal.encoding = TransanaConstants.chineseEncoding
##            elif (TransanaGlobal.configData.language == 'el'):
##                TransanaGlobal.encoding = 'iso8859_7'
##            elif (TransanaGlobal.configData.language == 'ja'):
##                TransanaGlobal.encoding = 'cp932'
##            elif (TransanaGlobal.configData.language == 'ko'):
##                TransanaGlobal.encoding = 'cp949'

        # Use UTF-8 Encoding throughout Transana to allow maximum internationalization
        if ('unicode' in wx.PlatformInfo) and (wx.VERSION_STRING >= '2.6'):
            wx.SetDefaultPyEncoding('utf_8')

        # On OS/X, change the working directory to the directory the script is 
        # running from, this is necessary for running from a bundle.
        if "__WXMAC__" in wx.PlatformInfo:
            if TransanaGlobal.programDir != '':
                os.chdir(TransanaGlobal.programDir)

        import MenuWindow                      # import Menu Window Object

        sys.excepthook = transana_excepthook        # Define the system exception handler

        # Create the global transana graphics colors, once the ConfigData object exists.
        TransanaGlobal.transana_graphicsColorList = TransanaGlobal.getColorDefs(TransanaGlobal.configData.colorConfigFilename)
        # Set essential global color manipulation data structures once the ConfigData object exists.
        (TransanaGlobal.transana_colorNameList, TransanaGlobal.transana_colorLookup, TransanaGlobal.keywordMapColourSet) = TransanaGlobal.SetColorVariables()
        
        # Add the RTF modules to the Python module search path.  This allows
        # us to import from a directory other than the standard search paths
        # and the current directory/subdirectories.
        sys.path.append("rtf")
        
        wx.InitAllImageHandlers()                       # Required on Mac to enable display of Splash Screen

        bitmap = wx.Bitmap("images/splash.gif", wx.BITMAP_TYPE_GIF)  # Load the Splash Screen graphic

        # We need to draw the Version Number onto the Splash Screen Graphic.
        # First, create a Memory DC
        memoryDC = wx.MemoryDC()
        # Select the bitmap into the Memory DC
        memoryDC.SelectObject(bitmap)
        # Determine the size of the version text
        (verWidth, verHeight) = memoryDC.GetTextExtent(_("Version %s") % TransanaConstants.versionNumber)
        # Add the Version Number text to the Memory DC (and therefore the bitmap)
        memoryDC.DrawText(_("Version %s") % TransanaConstants.versionNumber, 370 - verWidth, 156)
        # Clear the bitmap from the Memory DC, thus freeing it to be displayed!
        memoryDC.SelectObject(wx.EmptyBitmap(10, 10))
        # If the Splash Screen Graphic exists, display the Splash Screen for 4 seconds.
        # If not, raise an exception.
        if bitmap:
            # Mac requires a different style, as "STAY_ON_TOP" adds a header to the Splash Screen
            if "__WXMAC__" in wx.PlatformInfo:
                splashStyle = wx.SIMPLE_BORDER
            else:
                splashStyle = wx.SIMPLE_BORDER | wx.STAY_ON_TOP
            splash = wx.SplashScreen(bitmap,
                        wx.SPLASH_CENTRE_ON_SCREEN | wx.SPLASH_TIMEOUT,
                        4000, None, -1, wx.DefaultPosition, wx.DefaultSize, splashStyle)
        else:
            raise ImageLoadError, \
                    _("Unable to load Transana's splash screen image.  Installation error?")

        wx.Yield()

        import DataWindow                      # import Data Window Object
        import VideoWindow                     # import Video Window Object
        # import Transcript Window Object
        if TransanaConstants.USESRTC:
            import TranscriptionUI_RTC as TranscriptionUI
        else:
            import TranscriptionUI
        import VisualizationWindow             # import Visualization Window Object
        import exceptions                      # import exception handler (Python)
        # if we're running the multi-user version of Transana ...
        if not TransanaConstants.singleUserVersion:
            # ... import the Transana ChatWindow module
            import ChatWindow
        
        # Initialize all main application Window Objects

        # First, determine the program name that should be displayed, single or multi-user
        if TransanaConstants.singleUserVersion:
            programTitle = "Transana"
        else:
            programTitle = "Transana-MU"
        # Ammend the program title for the Demo version if appropriate
        if TransanaConstants.demoVersion:
            programTitle += " - Demonstration Version"

        # Create the Menu Window
        TransanaGlobal.menuWindow = MenuWindow.MenuWindow(None, -1, programTitle)

        # If we're running the Lab version OR we're on the Mac ...
        if TransanaConstants.labVersion or ('wxMac' in wx.PlatformInfo):
            # ... then pausing for 4 seconds delays the appearance of the Lab initial configuration dialog
            # or the Mac version login / database dialog until the Splash screen closes.
            time.sleep(4)
        # If we are running the Lab version ...
        if TransanaConstants.labVersion:
            # ... we want an initial configuration screen.  Start by importing Transana's Option Settings dialog
            import OptionsSettings
            # Initialize all paths to BLANK for the lab version
            TransanaGlobal.configData.videoPath = ''
            TransanaGlobal.configData.visualizationPath = ''
            TransanaGlobal.configData.databaseDir = ''
            # Create the config dialog for the Lab initial configuration
            options = OptionsSettings.OptionsSettings(TransanaGlobal.menuWindow, lab=True)
            options.Destroy()
            wx.Yield()
            # If the databaseDir is blank, user pressed CANCEL ...
            if (TransanaGlobal.configData.databaseDir == ''):
                # ... and we should quit immediately, signalling failure
                return False

        # initialze a variable indicating database connection to False (assume the worst.)
        connectionEstablished = False

        # Let's trap the situation where the database folder is not available.
        try:
            # Start MySQL if using the embedded version
            if TransanaConstants.singleUserVersion:
                DBInterface.InitializeSingleUserDatabase()
            # If we get here, we've been successful!  (NOTE that MU merely changes our default from False to True!)
            connectionEstablished = True
        except:
            if DEBUG:
                import traceback
                print sys.exc_info()[0], sys.exc_info()[1]
                traceback.print_exc(file=sys.stdout)
                
            msg = _('Transana is unable to access any Database at "%s".\nPlease check to see if this path is available.')
            if not TransanaConstants.labVersion:
                msg += '\n' + _('Would you like to restore the default Database path?')
            if ('unicode' in wx.PlatformInfo) and isinstance(msg, str):
                msg = unicode(msg, 'utf8')
            msg = msg % TransanaGlobal.configData.databaseDir
            if TransanaConstants.labVersion:
                dlg = Dialogs.ErrorDialog(None, msg)
                dlg.ShowModal()
                dlg.Destroy()
                return False

        # If we're on the Single-user version for Windows ...
        if TransanaConstants.singleUserVersion and ('wxMSW' in wx.PlatformInfo):
            # ... determine the file name for the data conversion information pickle file
            fs = os.path.join(TransanaGlobal.configData.databaseDir, '242_250_Convert.pkl')
            # If there is data in mid-conversion ...
            if os.path.exists(fs):
                # ... get the conversion data from the Pickle File
                f = file(fs, 'r')
                # exportedDBs is a dictionary containing the Transana-XML file name and the encoding of each DB to be imported
                self.exportedDBs = pickle.load(f)
                # Close the pickle file
                f.close()

                # Prompt the user about importing the converted data
                prompt = unicode(_("Transana has detected one or more databases ready for conversion.\nDo you want to convert those databases now?"), 'utf8')
                tmpDlg = Dialogs.QuestionDialog(TransanaGlobal.menuWindow, prompt, _("Database Conversion"))
                result = tmpDlg.LocalShowModal()
                tmpDlg.Destroy()

                # If the user wants to do the conversion now ...
                if result == wx.ID_YES:
                    # ... import Transana's XML Import code
                    import XMLImport
                    # Before we go further, let's make sure none of the conversion files have ALREADY been imported!
                    # The 2.42 to 2.50 Data Conversion Utility shouldn't allow this, but it doesn't hurt to check.
                    # Iterate through the conversion data
                    for key in self.exportedDBs:
                        # Determine the file path for the converted database
                        newDBPath = os.path.join(TransanaGlobal.configData.databaseDir, key + '_Converted')
                        # If the converted database ALREADY EXISTS ...
                        if os.path.exists(newDBPath):
                            # ... create an error message
                            prompt = unicode(_('Database "%s" already exists.\nDatabase "%s" cannot be converted at this time.'), 'utf8')
                            # ... display an error message
                            tmpDlg = Dialogs.ErrorDialog(None, prompt % (key + '_Converted', key))
                            tmpDlg.ShowModal()
                            tmpDlg.Destroy()
                        # If the converted database does NOT exist ...
                        else:
                            # Create the Import Database, passing the database name so the user won't be prompted for one.
                            DBInterface.establish_db_exists(dbToOpen = key + '_Converted')

                            # Import the database.
                            # First, create an Import Database dialog, but don't SHOW it.
                            temp = XMLImport.XMLImport(TransanaGlobal.menuWindow, -1, _('Transana XML Import'), importData=self.exportedDBs[key])
                            # ... Import the requested data!
                            temp.Import()
                            # Close the Import Database dialog
                            temp.Close()

                            # Close the database that was just imported
                            DBInterface.close_db()
                            # Clear the current database name from the Config data
                            TransanaGlobal.configData.database = ''
                            # If we do NOT have a localhost key (which Lab version does not!) ...
                            if not TransanaGlobal.configData.databaseList.has_key('localhost'):
                                # ... let's create one so we can pass the converted database name!
                                TransanaGlobal.configData.databaseList['localhost'] = {}
                                TransanaGlobal.configData.databaseList['localhost']['dbList'] = []
                            # Update the database name
                            TransanaGlobal.configData.database = key + '_Converted'
                            # Add the new (converted) database name to the database list
                            TransanaGlobal.configData.databaseList['localhost']['dbList'].append(key + '_Converted')
                            # Start exception handling
                            try:
                                # If we're NOT in the lab version of Transana ...
                                if not TransanaConstants.labVersion:
                                    # Add the unconverted database's PATH values to the CONVERTED database's configuration!
                                    TransanaGlobal.configData.pathsByDB[('', 'localhost', key + '_Converted')] = \
                                        {'visualizationPath' : TransanaGlobal.configData.pathsByDB[('', 'localhost', key)]['visualizationPath'],
                                         'videoPath' : TransanaGlobal.configData.pathsByDB[('', 'localhost', key)]['videoPath']}
                                # Save the altered configuration data
                                TransanaGlobal.configData.SaveConfiguration()
                            # The Computer Lab version sometimes throws a KeyError
                            except exceptions.KeyError:
                                # If this comes up, we can ignore it.
                                pass

                    # Delete the Import File Information
                    os.remove(fs)

        # We can only continue if we initialized the database OR are running MU.
        if connectionEstablished:
            # If a new database login fails three times, we need to close the program.
            # Initialize a counter to track that.
            logonCount = 1
            # Flag if Logon succeeds
            loggedOn = False
            # Keep trying for three tries or until successful
            while (logonCount <= 3) and (not loggedOn):
                logonCount += 1
                # Confirm the existence of the DB Tables, creating them if needed.
                # This method also calls the Username and Password Dialog if needed.
                # NOTE:  The Menu Window must be created first to server as a parent for the Username and Password Dialog
                #        called up by DBInterface.
                if DBInterface.establish_db_exists():
        
                    # Create the Data Window
                    # Data Window creation causes Username and Password Dialog to be displayed,
                    # so it should be created before the Video Window
                    self.dataWindow = DataWindow.DataWindow(TransanaGlobal.menuWindow)   
                    # Create the Video Window
                    self.videoWindow = VideoWindow.VideoWindow(TransanaGlobal.menuWindow)
                    # Create the Transcript Window.  If on the Mac, include the Close button.
                    self.transcriptWindow = TranscriptionUI.TranscriptionUI(TransanaGlobal.menuWindow, includeClose = ('wxMac' in wx.PlatformInfo))
                    # Create the Visualization Window
                    self.visualizationWindow = VisualizationWindow.VisualizationWindow(TransanaGlobal.menuWindow)

                    # Create the Control Object and register all objects to be controlled with it
                    self.controlObject = ControlObject()
                    self.controlObject.Register(Menu = TransanaGlobal.menuWindow,
                                                Video = self.videoWindow,
                                                Transcript=self.transcriptWindow,
                                                Data=self.dataWindow,
                                                Visualization=self.visualizationWindow)
                    # Set the active transcript
                    self.controlObject.activeTranscript = 0

                    # Register the ControlObject with all other objects to be controlled
                    TransanaGlobal.menuWindow.Register(ControlObject=self.controlObject)
                    self.dataWindow.Register(ControlObject=self.controlObject)
                    self.videoWindow.Register(ControlObject=self.controlObject)
                    self.transcriptWindow.Register(ControlObject=self.controlObject)
                    self.visualizationWindow.Register(ControlObject=self.controlObject)

                    # Set the Application Top Window to the Menu Window (wxPython)
                    self.SetTopWindow(TransanaGlobal.menuWindow)

                    # Show all Windows.
                    self.videoWindow.Show()
                    self.transcriptWindow.Show()
                    self.dataWindow.Show()
                    self.visualizationWindow.Show()
                    TransanaGlobal.menuWindow.Show(True)     # Show this last so it will have focus when the screen is displayed

                    loggedOn = True
                # If logon fails, inform user and offer to try again twice.
                elif logonCount <= 3:
                    dlg = Dialogs.QuestionDialog(TransanaGlobal.menuWindow, _('Transana was unable to connect to the database.\nWould you like to try again?'),
                                             _('Transana Database Connection'))
                    # If the user does not want to try again, set the counter to 4, which will cause the program to exit
                    if dlg.LocalShowModal() == wx.ID_NO:
                        logonCount = 4
                    # Clean up the Dialog Box
                    dlg.Destroy()

            # if we're running the multi-user version of Transana and successfully connected to a database ...
            if not TransanaConstants.singleUserVersion and loggedOn:
                # ... connect to the Message Server Here
                TransanaGlobal.socketConnection = ChatWindow.ConnectToMessageServer()
                # If the connections fails ...
                if TransanaGlobal.socketConnection == None:
                    # ... signal that Transana should NOT start up!
                    loggedOn = False
                else:
                    # If Transana MU sits idle too long (30 - 60 minutes), people would sometimes get a
                    # "Connection to Database Lost" error message even though MySQL was set to maintain the
                    # connection for 8 hours.  To try to address this, we will set up a Timer that will run
                    # a simple query every 10 minutes to maintain the connection to the database.

                    # Create the Connection Timer
                    TransanaGlobal.connectionTimer = wx.Timer(self)
                    # Bind the timer to its event
                    self.Bind(wx.EVT_TIMER, self.OnTimer)
                    # Tell the timer to fire every 10 minutes.
                    # NOTE:  If changing this value, it also needs to be changed in the ControlObjectClass.GetNewDatabase() method.
                    TransanaGlobal.connectionTimer.Start(600000)

            # If we're on OS X ...
            if 'wxMac' in wx.PlatformInfo:
                # ... create a list of the FFmpeg files that should be in /usr/local/lib
                ffmpegFileList = ['libmp4ff.a', 'libfaad.la', 'libfaad.a', 'libfaad.2.0.0.dylib', 'libmp3lame.la', 'libmp3lame.a',
                                  'libmp3lame.0.0.0.dylib']
                # Create a list of the Links that should be in /usr/local/lib, including what they should link to
                ffmpegLinkList = [('libfaad.2.0.0.dylib', 'libfaad.2.dylib'),
                                  ('libfaad.2.0.0.dylib', 'libfaad.dylib'),
                                  ('libmp3lame.0.0.0.dylib', 'libmp3lame.0.dylib'),
                                  ('libmp3lame.0.0.0.dylib', 'libmp3lame.dylib')]
                # Assume the files exist
                filesExist = True
                # For each file that should be in /usr/local/lib ...
                for filename in ffmpegFileList:
                    # ... if that file does not exist ...
                    if not os.path.exists('/usr/local/lib/%s' % filename):
                        # ... not that NOT ALL FILES exist ...
                        filesExist = False
                        # ... and stop looking
                        break
                # See if the FFMpeg libraries are installed correctly.  If not ...
                if not filesExist:
                    # ... get the program directory, removing the final directory which points to Transana.py
                    sourceDir = os.path.split(TransanaGlobal.programDir)[0]
                    # For each file in the FFmpeg file list ...
                    for filename in ffmpegFileList:
                        # ... if the file is not already in /usr/local/lib ...
                        if not os.path.exists('/usr/local/lib/%s' % filename):
                            # ... copy the file to /usr/local/lib
                            os.system("cp %s/Frameworks/%s /usr/local/lib/%s" % (sourceDir, filename, filename))
                    # For each file LINK in the FFmpeg List List ...
                    for (filename, linkname) in ffmpegLinkList:
                        # ... create the appropriate symbolic link
                        os.system("ln -fs /usr/local/lib/%s /usr/local/lib/%s" % (filename, linkname))

            # if this is the first time this user profile has used Transana ...
            if firstStartup:
                # ... create a prompt about looking at the Tutorial
                prompt = _('If this is your first time using Transana, the Transana Tutorial can help you learn how to use the program.')
                prompt += '\n\n' + _('Would you like to see the Transana Tutorial now?')
                # Display the Tutorial prompt in a Yes / No dialog
                tmpDlg = Dialogs.QuestionDialog(TransanaGlobal.menuWindow, prompt)
                # If the user says Yes ...
                if tmpDlg.LocalShowModal() == wx.ID_YES:
                    # ... start the Tutorial
                    self.controlObject.Help('Welcome to the Transana Tutorial')

        else:
            loggedOn = False
            dlg = Dialogs.QuestionDialog(TransanaGlobal.menuWindow, msg, _('Transana Database Connection'))
            if dlg.LocalShowModal() == wx.ID_YES:
                TransanaGlobal.configData.databaseDir = os.path.join(TransanaGlobal.configData.GetDefaultProfilePath(), 'databases')
                TransanaGlobal.configData.SaveConfiguration()
            # Clean up the Dialog Box
            dlg.Destroy()
        return loggedOn


    def OnTimer(self, event):
        """ To prevent a "Lost Database" message, we periocially run a very simple query to maintain our
            connection to the database. """
        # Get the database connection
        db = DBInterface.get_db()
        if db != None:
            # Get a DB Cursor
            dbCursor = db.cursor()
            # This is the simplest query I can think of
            query = "SHOW TABLES like 'Series%'"
            # Execute the query
            dbCursor.execute(query)


def transana_excepthook(extype, value, trace):
    """Custom global exception handler for Transana.  This is called when
    an unhandled exception occurs, or other errors that are otherwise not
    explicitly caught."""
    # First, do the regular behavior so we get traceback info in the
    # error log

#    if not(hasattr(sys, "frozen")):

    print
    print "Transana Error: ", time.asctime()
    print extype
    print value

    import traceback
    traceback.print_tb(trace, file=sys.stdout)

    sys.__excepthook__(extype, value, trace)
    # Now accomodate for the GUI
    msg = _("An unhandled %s exception occured") % extype
    try:
        msg = msg + ": " + str(value)
    except exceptions.AttributeError, e:
        # Exception doesn't support 'to string' via .args attribute
        msg = msg + "."

    dlg = Dialogs.ErrorDialog(None, msg)
    dlg.ShowModal()
    dlg.Destroy()


if __name__ == "__main__":

    # Main Application definition and execution call (wxPython)
    app = Transana(redirect=False)
    # Run the application main loop
    app.MainLoop()    
