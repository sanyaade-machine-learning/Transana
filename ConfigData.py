# Copyright (C) 2003-2006 The Board of Regents of the University of Wisconsin System 
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

"""This module contains Transana's Configuration Information."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare'

DEBUG = False
if DEBUG:
    print "ConfigData DEBUG is ON!      NOTE:  THIS MUST BE RUN THROUGH IDLE for Unicode support."

# Import wxPython
import wx
# import Python os module
import os
# Import Transana's Constants
import TransanaConstants
# Import Transana's Globals
import TransanaGlobal
# import Python's pickle module
import pickle

class ConfigData(object):
    """ This module handles Transana Configuration Data, including loading and saving this data. """
    def __init__(self):
        """ Initialize the ConfigData Object """
        # Load the data that has been stored or get default values if not stored
        self.LoadConfiguration()

        # Set default values for parameters that are not saved
        # Video Speed of 10 is equal to normal playback speed
        self.videoSpeed = 10
        # Auto Arrange is enabled by default
        self.autoArrange = True
        # Waveform Quickload is enabled by default
        self.waveformQuickLoad = True
    
    def __repr__(self):
        """ String Representation of the data in this object. """
        str = 'ConfigData Object:\n'
        str = str + 'host = %s\n' % self.host
        str = str + 'database = %s\n' % self.database
        str = str + 'visualizationPath = %s\n' % self.visualizationPath
        str = str + 'videoPath = %s\n' % self.videoPath
        str = str + 'transcriptionSetback = %s\n' % self.transcriptionSetback
        str = str + 'videoSpeed = %s\n' % self.videoSpeed
        str = str + 'videoSize = %s\n' % self.videoSize
        str = str + 'wordTracking = %s\n' % self.wordTracking
        str = str + 'autoArrange = %s\n' % self.autoArrange
        str = str + 'QuickLoad = %s\n' % self.waveformQuickLoad
        str = str + 'messageServer = %s\n' % self.messageServer
        str = str + 'messageServerPort = %s\n' % self.messageServerPort
        str = str + 'language = %s\n\n' % self.language
        str = str + 'databaseList = %s\n\n' % self.databaseList
        str = str + 'defaultFontFace = %s\n\n' % self.defaultFontFace
        str = str + 'defaultFontSize = %s\n\n' % self.defaultFontSize
        return str

    def LoadConfiguration(self):
        """ Load Configuration Data from Registry or Config File """
        # Define the default VisualizationPath as the Transana Data Path's 'waveforms' subfolder
        defaultVisualizationPath = os.path.join(self.GetDefaultProfilePath(), 'waveforms')
        # Define the default DatabasePath as the Transana Data Path's 'databases' subfolder
        defaultDatabaseDir = os.path.join(self.GetDefaultProfilePath(), 'databases')
        # Embedded MySQL can only work with Latin-1 compatible paths.  The following
        # code checks to make sure the path will work, and tries a couple of
        # substitutions if necessary.
        try:
            # First, let's see if the Database Dir is Latin-1 compatible
            temp = defaultDatabaseDir.encode('latin1')
        except UnicodeError:
            # If not, try the Program Dir + 'database'
            defaultDatabaseDir = os.path.join(TransanaGlobal.programDir, 'database')
            # Unfortunately, the Program Dir might not be Latin-1 compatible.  Check that.
            try:
                # See if the new Database Dir is Latin-1 compatible
                temp = defaultDatabaseDir.encode('latin1')
            except UnicodeError:
                # If we're still in trouble, let's build the path from scratch
                if 'wxMSW' in wx.PlatformInfo:
                    defaultDatabaseDir = os.path.join('C:', 'Transana 2', 'database')
                else:
                    # Actually, I have no idea if this will work.  I'd bet permissions issues will prevent it.
                    # But I have no way to pursue the issue further for the Mac until I find a user who's willing to help.
                    defaultDatabaseDir = os.path.join('Transana 2', 'database')

        # Define the Default Database Host
        if TransanaConstants.singleUserVersion:
            defaultHost = 'localhost'
        else:
            defaultHost = ''
        # Default Font Face is Courier New, a fixed-width font
        self.defaultFontFace = "Courier New"
        # Default Font Size is 10, reasoning that smaller is better
        self.defaultFontSize = 10
        
        # Load the Config Data.  wxConfig automatically uses the Registry on Windows and the appropriate file on Mac.
        # Program Name is Transana, Vendor Name is Verception to remain compatible with Transana 1.0.
        config = wx.Config('Transana', 'Verception')
        # See if a version 2.0 Configuration exists, and use it if so
        if config.Exists('/2.0'):
            # Load Host
            if TransanaConstants.singleUserVersion:
                self.host = config.Read('/2.0/host', defaultHost)
            else:
                self.host = config.Read('/2.0/hostMU', defaultHost)
            # Load Database and Database Directory(single user version only)
            if TransanaConstants.singleUserVersion:
                self.database = config.Read('/2.0/database', '')
                self.databaseDir = config.Read('/2.0/Directories/databaseDir', defaultDatabaseDir)
            else:
                self.database = config.Read('/2.0/databaseMU', '')
                self.databaseDir = config.Read('/2.0/Directories/databaseDir', defaultDatabaseDir)
            # Load Visualization Path
            self.visualizationPath = config.Read('/2.0/Directories/visualizationPath', defaultVisualizationPath)
            # Load Video Root Path
            self.videoPath = config.Read('/2.0/Directories/videoPath', '')
            # Load Transcriber Setback setting
            self.transcriptionSetback = config.ReadInt('/2.0/TranscriptionSetback', 2)
            # Load Video Size setting
            self.videoSize = config.ReadInt('/2.0/VideoSize', 100)
            # Load Auto Word-Tracking setting
            self.wordTracking = config.ReadInt('/2.0/WordTracking', True)
            # Load Message Server Host Setting
            self.messageServer = config.Read('/2.0/MessageHost', '')
            # Load Message Server Port Setting
            self.messageServerPort = config.ReadInt('/2.0/MessagePort', 17595)
            # Load Language Setting
            self.language = config.Read('/2.0/Language', '')
            # Load Default Font Face Setting
            self.defaultFontFace = config.Read('/2.0/FontFace', self.defaultFontFace)
            # Load Default Font Size Setting
            self.defaultFontSize = config.ReadInt('/2.0/FontSize', self.defaultFontSize)

        # If no version 2.0 Config File exists, ...
        else:

            # See if a verion 1 config file exists, and use it's data if it does.
            if config.Exists('/1.0/'):
                # Load Visualization Path
                self.visualizationPath = config.Read('/1.0/Directories/Waveforms', defaultVisualizationPath)
                # Set default database directory
                self.databaseDir = defaultDatabaseDir
                # Load Message Server Host Setting
                self.messageServer = config.Read('/1.0/MessageHost', '')
                # Load Message Server Port Setting
                self.messageServerPort = config.ReadInt('/1.0/MessagePort', 17595)

            # If no Config File exists, use default settings
            else:
                # Set Default Visualization Path
                self.visualizationPath = defaultVisualizationPath
                # Set Default Database directory
                self.databaseDir = defaultDatabaseDir
                # There is no default Message Server Host
                self.messageServer = ''
                # Set Default Message Server Port
                self.messageServerPort = 17595

            # Default Database Host
            self.host = defaultHost
            # Default Database
            self.database = ''
            # Set Default Video Root Folder
            self.videoPath = ''
            # Default Transcriber Setback is 2 seconds
            self.transcriptionSetback = 2
            # Default Video Size is 100%
            self.videoSize = 100
            # Auto Word Tracking is enabled by default
            self.wordTracking = True
            # Language setting
            self.language = ''

        # Load the databaseList, if it exists
        # NOTE:  if using Unicode, this MUST be a String object!
        if TransanaConstants.singleUserVersion:
            dbList = str(config.Read('/2.0/DatabaseListSU', ''))
        else:
            dbList = str(config.Read('/2.0/DatabaseListMU', ''))

        if DEBUG:
            print "ConfigData.LoadConfiguration():  dbList = '%s'" % dbList

        # The early versions (Alpha and Beta releases) used a different system.  Check to see
        # if we need to get the data from there!

        # If dbList is empty, let's see if a pickle file exists of the database host and database combinations
        if dbList == '':
            if TransanaConstants.singleUserVersion:
                # Name of the file for pickling the single-user Database Information
                dbFile = 'TransanaDBs.pkl'
            else:
                # Name of the file for pickling the multi-user Database Information
                dbFile = 'TransanaMUDBs.pkl'

            # Add the data path to the pickle file name
            dbFile = os.path.join(self.databaseDir, dbFile)
            # Initialize the database list to an empty dictionary
            self.databaseList = {}
            
            # See if Database Definitions have been pickled on this computer.
            if os.path.exists(dbFile):
               try:
                 # If so, try to open it 
                 file = open(dbFile, 'r')
                 # Load the Databases structure from the pickle file
                 self.databaseList = pickle.load(file)
               except:
                   print "Exception in ConfigData loading Databases."
                   self.Databases = {}

        # if the dbList from the Config files is not empty, we need to unpickle the string to restore the dictionary object
        else:
            self.databaseList = pickle.loads(dbList)

        if DEBUG:
            print "ConfigData.LoadConfiguration():  self.databaseList ="
            for h in self.databaseList.keys():
                for d in self.databaseList[h]:
                    print h, d

        # Embedded MySQL can only work with Latin-1 compatible paths.  The following
        # code checks to make sure the path will work.  This final check handles the situation where
        # an improper path is saved in the configuration file, which is entirely possible, as the selection
        # browser does no tests.
        try:
            # First, let's see if the Database Dir is Latin-1 compatible
            temp = self.databaseDir.encode('latin1')
        except UnicodeError:
            # NOTE:  Can't use the Dialogs.ErrorDialog here.  The wxApp object hasn't been created yet.
            msg = _("Illegal Database Directory specification.\nCurrent Directory replaced with\n%s.") % (defaultDatabaseDir)
            print msg
            self.databaseDir = defaultDatabaseDir

    def SaveConfiguration(self):
        """ Save Configuration Data to the Registry or a Config File. """
        # Save the Config Data.  wxConfig automatically uses the Registry on Windows and the appropriate file on Mac.
        # Program Name is Transana, Vendor Name is Verception to remain compatible with Transana 1.0.
        config = wx.Config('Transana', 'Verception')
        # Save the Host
        if TransanaConstants.singleUserVersion:
            config.Write('/2.0/host', self.host)
        else:
            config.Write('/2.0/hostMU', self.host)
        # Save the Database
        if TransanaConstants.singleUserVersion:
            config.Write('/2.0/database', self.database)
            config.Write('/2.0/Directories/databaseDir', self.databaseDir)
        else:
            config.Write('/2.0/databaseMU', self.database)
            config.Write('/2.0/Directories/databaseDir', self.databaseDir)
        # Save the Visualization Path
        config.Write('/2.0/Directories/visualizationPath', self.visualizationPath)
        # Save the Video Root Folder
        config.Write('/2.0/Directories/videoPath', self.videoPath)
        # Save the Transcriber Setback setting
        config.WriteInt('/2.0/TranscriptionSetback', self.transcriptionSetback)
        # Save the Video Size setting
        config.WriteInt('/2.0/VideoSize', self.videoSize)
        # Save the Auto Word Tracking setting
        config.WriteInt('/2.0/WordTracking', self.wordTracking)
        # Save the Message Server Host
        config.Write('/2.0/MessageHost', self.messageServer)
        # Save the Message Server Port
        config.WriteInt('/2.0/MessagePort', self.messageServerPort)
        # Save the Language setting
        config.Write('/2.0/Language', self.language)
        # NOTE:  Video Speed, Auto-Arrange, and Waveform Quickload are NOT saved to the config file.
        #        We decided it was better to have them reset to default values when the program is restarted.

        if DEBUG:
            print "ConfigData.SaveConfiguration():  self.databaseList ="
            for h in self.databaseList.keys():
                for d in self.databaseList[h]:
                    print h, d, type(h), type(d)

        # Save the list of Databases this user has used as a string by pickling it
        tmpDbList = pickle.dumps(self.databaseList)

        if DEBUG:
            print "ConfigData.SaveConfiguration():  tmpDbList = '%s'" % tmpDbList

        if TransanaConstants.singleUserVersion:
            config.Write('/2.0/DatabaseListSU', tmpDbList)
        else:
            config.Write('/2.0/DatabaseListMU', tmpDbList)
        # Save Default Font Face Setting
        config.Write('/2.0/FontFace', self.defaultFontFace)
        # Save Default Font Size Setting
        config.WriteInt('/2.0/FontSize', self.defaultFontSize)


    def GetDefaultProfilePath(self):
        """ Query the operating system and get the default path for user data. """
        # Initialize the default Profile Path to None
        defaultProfilePath = None
        # Try to get the proper Profile Path from the OS.
        if "__WXMSW__" in wx.Platform:
            # Windows 2K and XP should use APPDATA, which generally points to
            # C:\Documents and Settings\USERNAME\Application Data
            defaultProfilePath = os.getenv("APPDATA")
        elif "__WXMAC__" in wx.Platform:
            # Mac OS/X should use HOME, which generally points to
            # /Users/USERNAME
            defaultProfilePath = os.getenv("HOME")
        else: # Assuming that getenv("HOME") returns something useful
            defaultProfilePath = os.getenv("HOME")

        # defaultProfilePath = u'E:\\Video\\\u4eb2\u4eb3\u4eb2'
        # print "ConfigData.GetDefaultProfilePath() overridden"

        # I think the above fails for Windows 98 and Windows Me.  So if we don't get
        # something from the above, let's fall back to using the Program Directory here.
        if defaultProfilePath == None:
            defaultProfilePath = TransanaGlobal.programDir
        else:
            defaultProfilePath = os.path.join(defaultProfilePath, 'Transana 2')
        
        return defaultProfilePath
