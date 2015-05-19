# Copyright (C) 2003 - 2009 The Board of Regents of the University of Wisconsin System 
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

"""This module contains functions for encapsulating access to the database
for Transana."""

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>, Nathaniel Case, Rajas Sambhare'

DEBUG = False
if DEBUG:
    print "DBInterface DEBUG is ON!"
    
# import wxPython
import wx
# import MySQL for Python
import MySQLdb

if DEBUG:
    print "MySQLdb version =", MySQLdb.version_info, MySQLdb.__version__
    
# import Python's exceptions module
from exceptions import *
# We also need the MySQL exceptions!
import _mysql_exceptions
# import Python's array module
import array
# import Python's os module
import os
# import Python's sys module
import sys
# import Python's string module
import string
# import Transana's Clip object
import Clip
# Import Transana's Dialog Boxes
import Dialogs
# import Transana's Episode Object
import Episode
# import Transana's Constants
import TransanaConstants
# import Transana's Global Variables
import TransanaGlobal
# import Transana's Exceptions
import TransanaExceptions

_dbref = None

def InitializeSingleUserDatabase():
    """ For single-user Transana only, this initializes (starts) the embedded MySQL Server. """
    # See if the "databases" path exists off the Transana Program folder
    if "__WXMSW__" in wx.Platform:
        databasePath = TransanaGlobal.configData.databaseDir
    elif "__WXMAC__" in wx.Platform:
        databasePath = TransanaGlobal.configData.databaseDir
    else:
        databasePath = TransanaGlobal.configData.databaseDir
    
    if not os.path.exists(databasePath):
        # If not, create it
        os.makedirs(databasePath)

    # Start the embedded MySQL Server, using the "databases" folder off the Transana Program
    # folder for the root Data folder.
    datadir = "--datadir=" + databasePath
    # Default to English
    lang = '--language=./share/english'
    # Change the language if it is supported
    # Danish
    if (TransanaGlobal.configData.language == 'da'):
        lang = '--language=./share/danish'
    # German
    elif (TransanaGlobal.configData.language == 'de'):
        lang = '--language=./share/german'
    # Greek
    elif (TransanaGlobal.configData.language == 'el'):
        lang = '--language=./share/greek'
    # Spanish
    elif (TransanaGlobal.configData.language == 'es'):
        lang = '--language=./share/spanish'
    # French
    elif (TransanaGlobal.configData.language == 'fr'):
        lang = '--language=./share/french'
    # Italian
    elif (TransanaGlobal.configData.language == 'it'):
        lang = '--language=./share/italian'
    # Dutch
    elif (TransanaGlobal.configData.language == 'nl'):
        lang = '--language=./share/dutch'
    # Norwegian Bokmal
    elif (TransanaGlobal.configData.language == 'nb'):
        lang = '--language=./share/norwegian'
    # Norwegian Ny-norsk
    elif (TransanaGlobal.configData.language == 'nn'):
        lang = '--language=./share/norwegian-ny'
    # Polish
    elif (TransanaGlobal.configData.language == 'pl'):
        lang = '--language=./share/polish'
    # 
    elif (TransanaGlobal.configData.language == 'ru'):
        lang = '--language=./share/russian'
    # Swedish
    elif (TransanaGlobal.configData.language == 'sv'):
        lang = '--language=./share/swedish'

    # MySQLdb.server_init parameters MUST be strings, but paths with encoded characters aren't.
    if isinstance(datadir, unicode):
        # On Windows ...
        if 'wxMSW' in wx.PlatformInfo:
            # ... we seem to be able to fix this by using "CP1250" encoding
            datadir = datadir.encode('cp1250')
        # On non-Windows platforms, maybe we can use UTF8.
        else:
            datadir = datadir.encode('utf8')

    MySQLdb.server_init(args=['Transana', datadir, '--basedir=.', lang])
    
def EndSingleUserDatabase():
    """ For single-user Transana only, this ends (exits) the embedded MySQL Server. """
    # End the embedded MySQL Server
    MySQLdb.server_end()

def SetTableType(hasInnoDB, query):
    """ Set Table Type and Character Set Information for the database as appropriate """
    # NOTE:  Default Table Type switched from BDB to InnoDB for version 2.10.  BDB tables were having trouble importing
    #        a German Transana XML database someone sent to me, and switching databases was the only way I could find to
    #        fix the problem.  Therefore, the BDB tables will probably never be used, as I think they're always present
    #        if the BDB tables are.
    if TransanaGlobal.DBVersion >= u'5.0':
        if hasInnoDB:
            query = query + 'ENGINE=InnoDB'
        else:
            query = query + 'ENGINE=BDB'
    else:
        if hasInnoDB:
            query = query + 'TYPE=InnoDB'
        else:
            query = query + 'TYPE=BDB'

    if TransanaGlobal.DBVersion >= u'4.1':
        # Add the Character Set specification
        query += '  CHARACTER SET %s' % TransanaGlobal.encoding
    return query

def is_db_open():
    """ Quick and dirty test to see if the database is currently open """
    return _dbref != None

def establish_db_exists():
    """Check for the existence of all database tables and create them
    if necessary."""

    # NOTE:  Syntax for updating tables from MySQL 4.0 to MySQL 4.1 with Unicode UTF8 Characters Set:
    #          ALTER TABLE xxxx2 default character set utf8

    # Obtain a Database
    db = get_db()
    # If this fails, return "False" to indicate failure
    if db == None:
        return False

    # Obtain a Database Cursor
    dbCursor = db.cursor()

    # MySQL for Python 1.2.0 and later defaults to turning off AUTOCOMMIT.  We want AutoCommit to be ON.
    # query = "SET AUTOCOMMIT = 1"
    # Execute the Query
    # dbCursor.execute(query)
    db.autocommit(1)
    
    # MySQLdb 1.2.2 displays Warnings if the tables already exist as they are created.  We don't want this!
    if not MySQLdb.version_info in [(1, 2, 0, 'final', 1)]:
        dbCursor._defer_warnings = True

    # Initialize BDB and InnoDB Table Flags to false
    hasBDB = False
    hasInnoDB = False
    # First, let's find out if the InnoDB and BDB Tables are supported on the MySQL Instance we are using.
    # Define a "SHOW VARIABLES" Query
    query = "SHOW VARIABLES LIKE 'have%'"
    # Execute the Query
    dbCursor.execute(query)
    # Look at the Results Set
    for pair in dbCursor.fetchall():
        # If there is a pair in the Results that indicates that the value for "have_bdb" is "YES",
        # set the DBD Flag to True
        if pair[0] == 'have_bdb':
            if type(pair[1]).__name__ == 'array':
                p1 = pair[1].tostring()
            else:
                p1 = pair[1]
            if p1 == 'YES':
                hasBDB = True
        # If there is a pair in the Results that indicates that the value for "have_innodb" is "YES",
        # set the InnoDB Flag to True
        if pair[0] == 'have_innodb':
            if type(pair[1]).__name__ == 'array':
                p1 = pair[1].tostring()
            else:
                p1 = pair[1]
            if p1 == 'YES':
                hasInnoDB = True

    # If neither BDB nor InnoDB are supported, display an error message.
    if not (hasBDB or hasInnoDB):
        dlg = Dialogs.ErrorDialog(None, _("This MySQL Server is not configured to use BDB or InnoDB Tables.  Transana requires a MySQL-max Server."))
        dlg.ShowModal()
        dlg.Destroy
    # If either DBD or InnoDB is supported ...
    else:

        # Series Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Series2
                    (SeriesNum            INTEGER auto_increment, 
                     SeriesID             VARCHAR(100), 
                     SeriesComment        VARCHAR(255), 
                     SeriesOwner          VARCHAR(100), 
                     DefaultKeywordGroup  VARCHAR(50), 
                     RecordLock           VARCHAR(25),
                     LockTime             DATETIME, 
                     PRIMARY KEY (SeriesNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Episode Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Episodes2
                    (EpisodeNum     INTEGER auto_increment, 
                     EpisodeID      VARCHAR(100), 
                     SeriesNum      INTEGER, 
                     TapingDate     DATE, 
                     MediaFile      VARCHAR(255), 
                     EpLength       INTEGER, 
                     EpComment      VARCHAR(255), 
                     RecordLock     VARCHAR(25), 
                     LockTime       DATETIME, 
                     PRIMARY KEY (EpisodeNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Transcript Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Transcripts2
                    (TranscriptNum        INTEGER auto_increment, 
                     TranscriptID         VARCHAR(100), 
                     EpisodeNum           INTEGER,
                     SourceTranscriptNum  INTEGER,
                     ClipNum              INTEGER,
                     SortOrder            INTEGER,
                     Transcriber          VARCHAR(100),
                     ClipStart            INTEGER,
                     ClipStop             INTEGER,
                     Comment              VARCHAR(255), 
                     RTFText              LONGBLOB, 
                     RecordLock           VARCHAR(25), 
                     LockTime             DATETIME, 
                     LastSaveTime         DATETIME, 
                     PRIMARY KEY (TranscriptNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Now we need to check the Transcripts table to see if it has the SourceTranscriptNum and SortOrder
        # fields, added for Transana 2.30 to allow multi-transcript clips.

        # First, let's look at the Transcripts table structure.
        # Define the appropriate query
        query = "SHOW CREATE TABLE Transcripts2"
        # Execute the Query
        dbCursor.execute(query)
        # now let's look at the data returned from the database
        for data in dbCursor.fetchall():
            # Check for "array" data and convert if needed
            if type(data[1]).__name__ == 'array':
                d1 = data[1].tostring()
            else:
                d1 = data[1]
            # if a "SourceTranscriptNum" field is present, the table has already been altered.
            if not u"sourcetranscriptnum" in d1.lower():
                # If not, build a message to the user about upgrading the database
                msg = _("Transana has detected that this database needs to be upgraded.") + "\n" + \
                      _("Once you upgrade your database, you should not use it with older versions of Transana.") + "\n\n" + \
                      _("Do you want to upgrade this database at this time?")
                # Provide an extra warning for MU users
                if not TransanaConstants.singleUserVersion:
                    msg += "\n\n" + _("NOTE:  Upgrading the database before the Message Server has been upgraded\ncan lead to serious problems.  Do not upgrade unless you are SURE your\nMessage Server has already been upgraded.")
                # Create the dialog box
                dlg = Dialogs.QuestionDialog(None, msg, noDefault=True)
                # Display the error message
                result = dlg.LocalShowModal()
                # Destroy the dialog
                dlg.Destroy()
                # If the user answered "YES" to upgrading ...
                if result == wx.ID_YES:
                    # If not, we need to alter the table to add the SourceTranscriptNum field
                    query = """ ALTER TABLE Transcripts2
                                  ADD COLUMN
                                    SourceTranscriptNum  INTEGER AFTER EpisodeNum """
                    dbCursor2 = db.cursor()
                    dbCursor2.execute(query)
                    # ... then we need to alter the table to add the SortOrder field
                    query = """ ALTER TABLE Transcripts2
                                  ADD COLUMN
                                    SortOrder  INTEGER AFTER ClipNum """
                    dbCursor2.execute(query)
                    # ... then we need to alter the table to add the ClipStart field
                    query = """ ALTER TABLE Transcripts2
                                  ADD COLUMN
                                    ClipStart  INTEGER AFTER Transcriber """
                    dbCursor2.execute(query)
                    # ... then we need to alter the table to add the ClipStop field
                    query = """ ALTER TABLE Transcripts2
                                  ADD COLUMN
                                    ClipStop  INTEGER AFTER ClipStart """
                    dbCursor2.execute(query)
                    # Now find all the Clip Transcripts
                    query = " SELECT ClipNum, TranscriptNum, ClipStart, ClipStop FROM Clips2 "
                    dbCursor2.execute(query)
                    # Get the list of clips from the query
                    clipList = dbCursor2.fetchall()
                    # Create a progress dialog (This isn't even seen on small databases, but might be nice for large ones.)
                    progDlg = wx.ProgressDialog(_("Transana"), _("Database upgrade in progress"), maximum = len(clipList),
                                                style = wx.PD_APP_MODAL | wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME)
                    # We need another database cursor
                    dbCursor3 = db.cursor()
                    # Define a query that will move the Source Transcript data from the Clip data we've already got to
                    # the Transcript table
                    query = "UPDATE Transcripts2 SET SourceTranscriptNum = %d, SortOrder = 0, ClipStart = %d, ClipStop = %d WHERE ClipNum = %d"
                    # Initialize a counter for the Progress Dialog
                    counter = 0
                    # Iterate through the list of known clips ...
                    for (clipNum, transcriptNum, clipStart, clipStop) in clipList:
                        # ... update the Progress Dialog ...
                        (cont, skip) = progDlg.Update(counter)
                        # ... move the Source Transcript data from the Clip table to the Transcript table
                        dbCursor3.execute(query % (transcriptNum, clipStart, clipStop, clipNum))
                        # ... and increment the progress counter
                        counter += 1
                    # We can now close the Progress Dialog
                    progDlg.Destroy()
                    # Finally, drop TranscriptNum from the Clip table, now that the data's beem moved to the Transcript table.
                    query = """ ALTER TABLE Clips2
                                  DROP COLUMN TranscriptNum """
                    dbCursor2.execute(query)
                # If the user says "NO" to upgrading ...
                else:
                    # Close the database before any changes get made!
                    close_db()
                    # signal failure to connect to the database
                    return False

        # Collection Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Collections2
                    (CollectNum            INTEGER auto_increment, 
                     CollectID             VARCHAR(100), 
                     ParentCollectNum      INTEGER, 
                     CollectComment        VARCHAR(255), 
                     CollectOwner          VARCHAR(100), 
                     DefaultKeywordGroup   VARCHAR(50), 
                     RecordLock          VARCHAR(25), 
                     LockTime            DATETIME, 
                     PRIMARY KEY (CollectNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        #  Clip Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Clips2
                    (ClipNum        INTEGER auto_increment, 
                     ClipID         VARCHAR(100), 
                     CollectNum     INTEGER, 
                     EpisodeNum     INTEGER, 
                     MediaFile      VARCHAR(255), 
                     ClipStart      INTEGER, 
                     ClipStop       INTEGER,
                     ClipOffset     INTEGER,
                     Audio          INTEGER,
                     ClipComment    VARCHAR(255), 
                     SortOrder      INTEGER, 
                     RecordLock     VARCHAR(25), 
                     LockTime       DATETIME, 
                     PRIMARY KEY (ClipNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Now, let's look at the Clips table structure.
        # Define the appropriate query
        query = "SHOW CREATE TABLE Clips2"
        # Execute the Query
        dbCursor.execute(query)
        # now let's look at the data returned from the database
        for data in dbCursor.fetchall():
            # Check for "array" data and convert if needed
            if type(data[1]).__name__ == 'array':
                d1 = data[1].tostring()
            else:
                d1 = data[1]
            # if a "ClipOffset" field is present, the table has already been altered.
            if not u"clipoffset" in d1.lower():
                # If not, we need to alter the table to add the ClipOffset field
                query = """ ALTER TABLE Clips2
                              ADD COLUMN
                                ClipOffset  INTEGER
                              AFTER ClipStop """
                dbCursor2 = db.cursor()
                dbCursor2.execute(query)
                # Define a query that will set ClipOffset to the default value of 0
                query = "UPDATE Clips2 SET ClipOffset = 0"
                dbCursor2.execute(query)
            # Check to see if an "Audio" field is present.
            if not u"audio" in d1.lower():
                # If not, we need to alter the table to add the Audio field
                query = """ ALTER TABLE Clips2
                              ADD COLUMN
                                Audio  INTEGER
                              AFTER ClipOffset """
                dbCursor2 = db.cursor()
                dbCursor2.execute(query)
                # Define a query that will set Audio to the default value of 1
                query = "UPDATE Clips2 SET Audio = 1"
                dbCursor2.execute(query)

        # Notes Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Notes2
                    (NoteNum        INTEGER auto_increment, 
                     NoteID         VARCHAR(100), 
                     SeriesNum      INTEGER, 
                     EpisodeNum     INTEGER, 
                     CollectNum     INTEGER, 
                     ClipNum        INTEGER, 
                     TranscriptNum  INTEGER, 
                     NoteTaker      VARCHAR(100), 
                     NoteText       LONGBLOB, 
                     RecordLock     VARCHAR(25), 
                     LockTime       DATETIME, 
                     PRIMARY KEY (NoteNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Now we need to check the Notes table to see if the NoteText field is a BLOB or a LONGBLOB,
        # and we need to increase the size of the field if it's a BLOB.

        # I originally used SHOW COLUMNS to detect this, but that caused a WEAK REFERENCE failure on
        # changing databases against my linux box running MySQL 5.0.27. (Not sure what the relevant
        # factor was in that.  Just know it fails.)  So I switched to SHOW CREATE TABLE.

        # query = "SHOW COLUMNS FROM Notes2 LIKE 'NoteText'"
        query = "SHOW CREATE TABLE Notes2"
        # Execute the Query
        dbCursor.execute(query)
        # now let's look at the data returned from the database
        for data in dbCursor.fetchall():
            # Check for "array" data and convert if needed
            if type(data[1]).__name__ == 'array':
                d1 = data[1].tostring()
            else:
                d1 = data[1]
            # if a LONGBLOB is present, we can skip this.     # If the Field is "NoteText" and the Type is "Blob" ...
            if not u"longblob" in d1.lower():            # (data[0].lower() == u'notetext') and (data[1].lower() == u'blob'):
                # ... then we need to alter the table to change the data type to LONGBLOB.
                query = "ALTER TABLE Notes2 MODIFY NoteText LONGBLOB"
                dbCursor2 = db.cursor()
                dbCursor2.execute(query)

        # Keywords Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Keywords2
                    (KeywordGroup  VARCHAR(50) NOT NULL, 
                     Keyword       VARCHAR(85) NOT NULL, 
                     Definition    LONGBLOB, 
                     RecordLock    VARCHAR(25), 
                     LockTime      DATETIME, 
                     PRIMARY KEY (KeywordGroup, Keyword))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Now we need to check the Keywords table to see if the Definition field is a VARCHAR(255) or a LONGBLOB,
        # and we need to increase the size of the field if it's a VARCHAR.

        query = "SHOW CREATE TABLE Keywords2"
        # Execute the Query
        dbCursor.execute(query)
        # now let's look at the data returned from the database
        for data in dbCursor.fetchall():
            # Check for "array" data and convert if needed
            if type(data[1]).__name__ == 'array':
                d1 = data[1].tostring()
            else:
                d1 = data[1]
            # if a LONGBLOB is present, we can skip this.
            if not u"longblob" in d1.lower():
                # ... then we need to alter the table to change the data type to LONGBLOB.
                query = "ALTER TABLE Keywords2 MODIFY Definition LONGBLOB"
                dbCursor2 = db.cursor()
                dbCursor2.execute(query)

        # ClipKeywords Table: Test for existence and create if needed
        # MySQL Primary Keys cannot contain NULL values, and either EpisodeNum
        # or ClipNum will always be NULL!  Therefore, use a UNIQUE KEY rather
        # than a PRIMARY KEY for this table.
        query = """
                  CREATE TABLE IF NOT EXISTS ClipKeywords2
                    (EpisodeNum    INTEGER, 
                     ClipNum       INTEGER, 
                     KeywordGroup  VARCHAR(50), 
                     Keyword       VARCHAR(85), 
                     Example       CHAR(1), 
                     UNIQUE KEY (EpisodeNum, ClipNum, KeywordGroup, Keyword))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Core Data Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS CoreData2
                    (CoreDataNum    INTEGER auto_increment, 
                     Identifier     VARCHAR(255), 
                     Title          VARCHAR(255), 
                     Creator        VARCHAR(255), 
                     Subject        VARCHAR(255), 
                     Description    VARCHAR(255), 
                     Publisher      VARCHAR(255), 
                     Contributor    VARCHAR(255), 
                     DCDate         DATE, 
                     DCType         VARCHAR(50), 
                     Format         VARCHAR(100), 
                     Source         VARCHAR(255), 
                     Language       VARCHAR(25), 
                     Relation       VARCHAR(255), 
                     Coverage       VARCHAR(255), 
                     Rights         VARCHAR(255), 
                     RecordLock     VARCHAR(25), 
                     LockTime       DATETIME, 
                     PRIMARY KEY (CoreDataNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Filters Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Filters2
                    (ReportType      INTEGER, 
                     ReportScope     INTEGER, 
                     ConfigName      VARCHAR(100),
                     FilterDataType  INTEGER,
                     FilterData      LONGBLOB,
                     PRIMARY KEY (ReportType, ReportScope, ConfigName, FilterDataType))
                """
        # See FilterDialog.py for a list of ReportTypes adn FilterDataTypes
        
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Now we need to check the Filters table to see if the FilterData field is a BLOB or a LONGBLOB,
        # and we need to increase the size of the field if it's a BLOB.

        # I originally used SHOW COLUMNS to detect this, but that caused a WEAK REFERENCE failure on
        # changing databases against my linux box running MySQL 5.0.27. (Not sure what the relevant
        # factor was in that.  Just know it fails.)  So I switched to SHOW CREATE TABLE.

        # query = "SHOW COLUMNS FROM Filters2 LIKE 'FilterData'"
        query = "SHOW CREATE TABLE Filters2"
        # Execute the Query
        dbCursor.execute(query)
        # now let's look at the data returned from the database
        for data in dbCursor.fetchall():
            # Check for "array" data and convert if needed
            if type(data[1]).__name__ == 'array':
                d1 = data[1].tostring()
            else:
                d1 = data[1]
            # if a LONGBLOB is present, we can skip this.     # If the Field is "FilterData" and the Type is "Blob" ...
            if not u"longblob" in d1.lower():            # (data[0].lower() == u'filterdata') and (data[1].lower() == u'blob'):
                # ... then we need to alter the table to change the data type to LONGBLOB.
                query = "ALTER TABLE Filters2 MODIFY FilterData LONGBLOB"
                dbCursor2 = db.cursor()
                dbCursor2.execute(query)

        # We need to detect the upgrade to version 2.40.  This should do that.
        query = "SHOW TABLES LIKE 'AdditionalVids2'"
        dbCursor.execute(query)
        if dbCursor.rowcount == 0:
            UpdateTranscriptRecsfor240(None)

        # AdditionalVids (Additional Videos) Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS AdditionalVids2
                    (AddVidNum      INTEGER auto_increment,
                     EpisodeNum     INTEGER,
                     ClipNum        INTEGER,
                     MediaFile      VARCHAR(255), 
                     VidLength      INTEGER,
                     Offset         INTEGER,
                     Audio          INTEGER,
                     PRIMARY KEY (AddVidNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasInnoDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # See if this (username, server, database) combination has defined paths.
        if TransanaGlobal.configData.pathsByDB.has_key((TransanaGlobal.userName.encode('utf8'), TransanaGlobal.configData.host.encode('utf8'), TransanaGlobal.configData.database.encode('utf8'))):
            # If so, load the video root and visualization paths.
            TransanaGlobal.configData.videoPath = TransanaGlobal.configData.pathsByDB[(TransanaGlobal.userName.encode('utf8'), TransanaGlobal.configData.host.encode('utf8'), TransanaGlobal.configData.database.encode('utf8'))]['videoPath']
            if isinstance(TransanaGlobal.configData.videoPath, str):
                TransanaGlobal.configData.videoPath = TransanaGlobal.configData.videoPath.decode('utf8')
            TransanaGlobal.configData.visualizationPath = TransanaGlobal.configData.pathsByDB[(TransanaGlobal.userName.encode('utf8'), TransanaGlobal.configData.host.encode('utf8'), TransanaGlobal.configData.database.encode('utf8'))]['visualizationPath']
            if isinstance(TransanaGlobal.configData.visualizationPath, str):
                TransanaGlobal.configData.visualizationPath = TransanaGlobal.configData.visualizationPath.decode('utf8')
        # If not ...
        else:
            # ... use the current video root and visualization paths to initialize the configuration values.
            TransanaGlobal.configData.pathsByDB[(TransanaGlobal.userName.encode('utf8'), TransanaGlobal.configData.host.encode('utf8'), TransanaGlobal.configData.database.encode('utf8'))] = \
                    {'videoPath' : TransanaGlobal.configData.videoPath.encode('utf8'),
                     'visualizationPath' : TransanaGlobal.configData.visualizationPath.encode('utf8')}

        # If we've gotten this far, return "true" to indicate success.
        return True


def UpdateTranscriptRecsfor240(self):
    """ For release 2.40, I changed the way clips created from other clips track their source transcript.
        Instead of remembering the Clip Transcript they were taken from, which often may get deleted,
        they need to point to the original Episode Transcript, which rarely gets deleted.
        This routine looks for example and corrects them if possible. """
    # Obtain a Database
    db = get_db()
    # Obtain Database Cursors
    dbCursor = db.cursor()
    dbCursor2 = db.cursor()

    # We need to create a Clip List of all Clip Transcripts drawn from other Clip Transcripts.
    # The version of MySQL used for the single-user Windows version of Transana doesn't support subqueries.
    
    # Determine if that's the version we've got.
    if ('wxMSW' in wx.PlatformInfo) and TransanaConstants.singleUserVersion:
        # If so, initialize the Clip List
        clipList = []
        # Start with the subquery term, Creating a query to gather a list of all clip transcript numbers
        query = "SELECT TranscriptNum FROM Transcripts2 WHERE ClipNum > 0"
        # Execute the query
        dbCursor.execute(query)
        # Store the results as ClipList1.  The format is [(ClipNum,), (ClipNum,) ... ]
        clipList1 = dbCursor.fetchall()
        # Now create a query for the outer query, a list of all transcripts
        query = "SELECT TranscriptNum, SourceTranscriptNum, ClipNum FROM Transcripts2 T "
        # Execute the query.
        dbCursor2.execute(query)
        # Iterate through the outer query's results
        for (trNum, sTrNum, cNum) in dbCursor2.fetchall():
            # If the Source Transcript number is a CLIP Transcript ...
            if (sTrNum,) in clipList1:
                # ... then append this record to the Clip List
                clipList.append((trNum, sTrNum, cNum))
    # If we're not on single-user Windows Transana, we can use nested queries.            
    else:
        # Create a query to get the clips that need updating
        query = "SELECT TranscriptNum, SourceTranscriptNum, ClipNum FROM Transcripts2 T where SourceTranscriptNum in (SELECT TranscriptNum FROM Transcripts2 WHERE ClipNum > 0)"
        # Execute the query
        dbCursor.execute(query)
        # The Clip List is the query results.
        clipList = dbCursor.fetchall()

    # Iterate through the Clip List
    for (TranscriptNum, SourceTranscriptNum, ClipNum) in clipList:
        # Set the initial values from the larger query
        TNum2 = TranscriptNum
        STNum2 = SourceTranscriptNum
        CN2 = ClipNum
        # While the current transcript is a Clip transcript ...
        while CN2 > 0:
            # Create a secondary query to get the current transcript's Source Transcript
            query = "SELECT TranscriptNum, SourceTranscriptNum, ClipNum FROM Transcripts2 WHERE TranscriptNum = %d"
            # Execute the secondary query
            dbCursor2.execute(query % STNum2)
            # If data is returned ...
            if dbCursor2.rowcount >= 1:
                # ... get the data
                (TNum2, STNum2, CN2) = dbCursor2.fetchone()
            # If no data is returned ...
            else:
                # ... then indicate that with all 0 values
                (TNum2, STNum2, CN2) = (0, 0, 0)
        # Build a query to update the data in the database
        query = 'UPDATE Transcripts2 SET SourceTranscriptNum = %d WHERE TranscriptNum = %d'
        # Execute the query
        dbCursor2.execute(query % (TNum2, TranscriptNum))

def get_db():
    """Get a connection object reference to the database.  If a connection
    has not yet been established, then create the connection."""
    global _dbref
    # If a database reference is not defined ...
    if (_dbref == None):
        # import the Username and Password Dialog.
        # (This dialog requests Username, Password, dbServer, and Database Name for the multi-user version
        # of Transana, Database Name only for the single-user version)
        from UsernameandPasswordClass import UsernameandPassword
        # Create the Dialog Box
        UsernameForm = UsernameandPassword(TransanaGlobal.menuWindow)
        # Get the Data Entered in the Dialog
        (userName, password, dbServer, databaseName, port) = UsernameForm.GetValues()
        # Destroy the form now that we're done with it.
        UsernameForm.Destroy()
        # Check for the validity of the data.
        # The single-user version of Transana needs only the Database Name.  The multi-user version of
        # Transana requires all four values.
        if (databaseName == '') or \
           ((not TransanaConstants.singleUserVersion) and ((userName == '') or (password == '') or (dbServer == '') or (port == ''))):
            # If the Username and Password form is not filled out completely, the get_db method fails and returns "None"
            return None
        # Otherwise, all data was provided by the user.
        else:
            try:

                if DEBUG:
                    print "Establishing Database Server connection."
                    
                # Assign the username to a global variable
                TransanaGlobal.userName = userName  

                # Establish a connection to the Database Server.
                # NOTE:  This does not yet support Unicode.  We'll delay the Database Name so it can be unicode.
                if TransanaConstants.singleUserVersion:
                    if 'unicode' in wx.PlatformInfo:
                        # The single-user version requires no parameters
                        _dbref = MySQLdb.connect(use_unicode=True)
                    else:
                        # The single-user version requires no parameters
                        _dbref = MySQLdb.connect()
                else:
                    if 'unicode' in wx.PlatformInfo:
                        _dbref = MySQLdb.connect(host=dbServer, user=userName, passwd=password, port=int(port), use_unicode=True)
                    else:
                        # The multi-user version requires all information to connect to the database server
                        _dbref = MySQLdb.connect(host=dbServer, user=userName, passwd=password, port=int(port))
                    # Put the Host name in the Configuration Data
                    # so that the same connection can be the default for the next logon
                    TransanaGlobal.configData.host = dbServer
                    TransanaGlobal.configData.dbport = port

            except MySQLdb.OperationalError:
                if DEBUG:
                    print "DBInterface.get_db():  ", sys.exc_info()[1]

                _dbref = None

            # If the Database Connection fails, an exception is raised.
            except:

                if DEBUG:

                    print "DBInterface.get_db():  Exception 1"
                    
                    print sys.exc_info()[0], sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                _dbref = None

            # Okay, this is a little weird.
            # By default, MySQL limits the size of a record to 1 MB.  This has started causing some problems with large transcripts.
            # So we need to change the value of the "max_allowed_packet" to fix this.  However, we apparently need to break our
            # database connection and re-establish it for the parameter change to go into effect.  That's what the following block
            # of code does.

            # Initialize a value, as I suppose it's possible it won't be found.
            max_allowed_packet = 0
            
            # If we were able to make a connection to the database...
            if _dbref != None:
                # Let's find out the current setting of max_allowed_packets
                # Query the DB for the current value
                query = "SHOW VARIABLES LIKE 'max_allowed_packe%'"
                dbCursor = _dbref.cursor()
                # Execute the Query
                dbCursor.execute(query)
                # Look at the Results Set
                for pair in dbCursor.fetchall():
                    # Find the max_allowed_packet variable
                    if pair[0] == 'max_allowed_packet':
                        # Its value comes in different forms depending on what version of MySQL and MySQL for Python we're using.
                        if type(pair[1]) == array.array:
                            max_allowed_packet = pair[1].tostring()
                        else:
                            max_allowed_packet = pair[1]
                        
                # We need to increase the size of the maximum allowed "packet" from 1MB (default) to 8MB.

                if DEBUG:
                    print "DBInterface.get_db():  max_allowed_packet:", int(max_allowed_packet), '<', 8388608, '=', (int(max_allowed_packet) < 8388608)
                
                if int(max_allowed_packet) < 8388608:
                    dbCursor.execute('SET GLOBAL max_allowed_packet=8388608')
                    # If we had to change this, we need to shut down our connection and re-establish it for the change to "take".
                    close_db()
                    # Re-establish a connection to the Database Server.
                    if TransanaConstants.singleUserVersion:
                        if 'unicode' in wx.PlatformInfo:
                            # The single-user version requires no parameters
                            _dbref = MySQLdb.connect(use_unicode=True)
                        else:
                            # The single-user version requires no parameters
                            _dbref = MySQLdb.connect()
                    else:
                        if 'unicode' in wx.PlatformInfo:
                            _dbref = MySQLdb.connect(host=dbServer, user=userName, passwd=password, port=int(port), use_unicode=True)
                        else:
                            # The multi-user version requires all information to connect to the database server
                            _dbref = MySQLdb.connect(host=dbServer, user=userName, passwd=password, port=int(port))

                # We need to know the MySQL version we're dealing with to know if UTF-8 is supported.
                # Get a Database Cursor
                dbCursor = _dbref.cursor()
                # Query the Database about what Database Names have been defined
                dbCursor.execute('SELECT VERSION()')
                vs = dbCursor.fetchall()
                for v in vs:
                    TransanaGlobal.DBVersion = v[0][:3]

                    if DEBUG:
                        print "MySQL Version =", TransanaGlobal.DBVersion, type(TransanaGlobal.DBVersion)
                        
            try:
                # If we made a connection to MySQL...
                if _dbref != None:
                    # If we have MySQL 4.1 or later, we have UTF-8 support and should use it.
                    if TransanaGlobal.DBVersion >= u'4.1':
                        # Get a Database Cursor
                        dbCursor = _dbref.cursor()
                        # Set Character Encoding settings
                        dbCursor.execute('SET CHARACTER SET utf8')
                        dbCursor.execute('SET character_set_connection = utf8')
                        dbCursor.execute('SET character_set_client = utf8')
                        dbCursor.execute('SET character_set_server = utf8')
                        dbCursor.execute('SET character_set_database = utf8')
                        dbCursor.execute('SET character_set_results = utf8')
                        
                        dbCursor.execute('USE %s' % databaseName.encode('utf8'))
                        # Set the global character encoding to UTF-8
                        TransanaGlobal.encoding = 'utf8'
                    # If we're using MySQL 4.0 or earlier, we lack UTF-8 support, so should use 
                    # another language-appropriate encoding, or Latin-1 encoding as a fall-back
                    else:
                        # If we're in Russian, change the encoding to KOI8r
                        if TransanaGlobal.configData.language == 'ru':
                            TransanaGlobal.encoding = 'koi8_r'
                        # If we're in Chinese, change the encoding to the appropriate Chinese encoding
                        elif TransanaGlobal.configData.language == 'zh':
                            TransanaGlobal.encoding = TransanaConstants.chineseEncoding
                        # If we're in Eastern Europe Encoding, change the encoding to 'iso8859_2'
                        elif TransanaGlobal.configData.language == 'easteurope':
                            TransanaGlobal.encoding = 'iso8859_2'
                        # If we're in Greek, change the encoding to 'iso8859_7'
                        elif TransanaGlobal.configData.language == 'el':
                            TransanaGlobal.encoding = 'iso8859_7'
                        # If we're in Japanese, change the encoding to cp932
                        elif TransanaGlobal.configData.language == 'ja':
                            TransanaGlobal.encoding = 'cp932'
                        # If we're in Korean, change the encoding to cp949
                        elif TransanaGlobal.configData.language == 'ko':
                            TransanaGlobal.encoding = 'cp949'
                        # Otherwise, fall back to utf8??
                        else:
                            TransanaGlobal.encoding = 'utf8'  # 'latin1'


                        dbCursor.execute('USE %s' % databaseName.encode(TransanaGlobal.encoding))

                    if TransanaConstants.demoVersion:
                        # Get a Database Cursor
                        dbCursor.execute('SELECT COUNT(EpisodeNum) from Episodes2')
                        # Determine the number of Episode records
                        epCount = dbCursor.fetchone()[0]
                        # Determine the number of Episode Transcript records (exclude Clip Transcripts)
                        dbCursor.execute('SELECT COUNT(TranscriptNum) FROM Transcripts2 WHERE ClipNum = 0')
                        trCount = dbCursor.fetchone()[0]
                        # Determine the number of Clip records
                        dbCursor.execute('SELECT COUNT(ClipNum) from Clips2')
                        clCount = dbCursor.fetchone()[0]
                        # Determine the number of Keyword records
                        dbCursor.execute('SELECT COUNT(Keyword) from Keywords2')
                        kwCount = dbCursor.fetchone()[0]
                        # Check to see that the Demo limits are not exceeded
                        if (epCount > TransanaConstants.maxEpisodes) or \
                           (trCount > TransanaConstants.maxEpisodeTranscripts) or \
                           (clCount > TransanaConstants.maxClips) or \
                           (kwCount > TransanaConstants.maxKeywords):
                            # If they are, display an error message
                            errormsg = _("The data in this database exceeds what is allowed in the Transana Demonstration.")
                            errordlg = Dialogs.ErrorDialog(None, errormsg)
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # Close the Database Cursor
                            dbCursor.close()
                            # Close the Database Connection
                            _dbref.close()
                            # If database limits have been exceeded, block the database open.
                            _dbref = None

                    TransanaGlobal.configData.database = databaseName

            except MySQLdb.OperationalError:
                if DEBUG:
                    print "DBInterface.get_db():  Unknown Database!"

                # Skip the Database Creation message if we're in Demonstartion Mode
                if not TransanaConstants.demoVersion:
                    # If the Database Name was not found, prompt the user to see if they want to create a new Database.
                    # First, create the Prompt Dialog
                    # NOTE:  This does not use Dialogs.ErrorDialog because it requires a Yes/No reponse
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Database "%s" does not exist.  Would you like to create it?\n(If you do not have rights to create a database, see your system administrator.)'), 'utf8')
                    else:
                        prompt = _('Database "%s" does not exist.  Would you like to create it?\n(If you do not have rights to create a database, see your system administrator.)')
                    dlg = Dialogs.QuestionDialog(None, prompt % databaseName)
                    # Display the Dialog
                    result = dlg.LocalShowModal()
                    # Clean up after the Dialog
                    dlg.Destroy()
                else:
                    result = wx.ID_YES
                # If the user wants to create a new Database ...
                if result == wx.ID_YES:
                    try:
                        if 'unicode' in wx.PlatformInfo:
                            tempDatabaseName = databaseName.encode(TransanaGlobal.encoding)
                        else:
                            tempDatabaseName = databaseName
                        if TransanaGlobal.DBVersion >= u'4.1':
                            query = 'CREATE DATABASE IF NOT EXISTS %s CHARACTER SET utf8' % tempDatabaseName
                        else:
                            query = 'CREATE DATABASE IF NOT EXISTS %s' % tempDatabaseName
                        # ... create the Database ...
                        dbCursor.execute(query)
                        # ... specify that the new database should be used ...
                        dbCursor.execute('USE %s' % tempDatabaseName)
                        TransanaGlobal.configData.database = databaseName
                        # Close the Database Cursor
                        dbCursor.close()
                    # If the Create fails ...
                    except:
                        if DEBUG:
                            print sys.exc_info()[0], sys.exc_info()[1]
                            import traceback
                            traceback.print_exc(file=sys.stdout)

                        # ... the user probably lacks CREATE parmission in the Database Rights structure.
                        # Create an error message Dialog
                        dlg = Dialogs.ErrorDialog(None, _('Database Creation Error.\nYou specified an illegal database name, or do not have rights to create a database.\nTry again with a simple database name (with no punctuation or spaces), or see your system administrator.'))
                        # Display the Error Message.
                        dlg.ShowModal()
                        # Clean up the Error Message
                        dlg.Destroy()
                        # Close the Database Cursor
                        dbCursor.close()
                        # Close the Database Connection
                        _dbref.close()
                        _dbref = None
                else:
                    # Close the Database Cursor
                    dbCursor.close()
                    # Close the Database Connection
                    _dbref.close()
                    _dbref = None

            except UnicodeEncodeError, e:

                if DEBUG:

                    print "DBInterface.get_db():  Exception:", TransanaGlobal.configData.language, TransanaGlobal.encoding

                    print sys.exc_info()[0], sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                # ... The only time I've seen an error here has to do with encoding failures.
                # Create an error message Dialog
                dlg = Dialogs.ErrorDialog(None, _("Unicode Error opening the database.\nIs it possible your current language setting doesn't match the database's language?\nPlease open a different database, change your language setting, then try this database again."))
                # Display the Error Message.
                dlg.ShowModal()
                # Clean up the Error Message
                dlg.Destroy()

                # Close the Database Cursor
                dbCursor.close()
                # Close the Database Connection
                _dbref.close()
                _dbref = None

            except _mysql_exceptions.ProgrammingError, e:

                if DEBUG:

                    print "DBInterface.get_db():  Exception:", TransanaGlobal.configData.language, TransanaGlobal.encoding

                    print sys.exc_info()[0], sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                # ... The only time I've seen an error here has to do with encoding failures.
                # Create an error message Dialog
                dlg = Dialogs.ErrorDialog(None, _("MySQL Error opening the database.\nTry again with a simple database name (with no punctuation or spaces.)\nAlso try entering a database name in English."))
                # Display the Error Message.
                dlg.ShowModal()
                # Clean up the Error Message
                dlg.Destroy()

                # Close the Database Cursor
                dbCursor.close()
                # Close the Database Connection
                _dbref.close()
                _dbref = None

            # If the Database Connection fails, an exception is raised.
            except:

                if DEBUG:

                    print "DBInterface.get_db():  Exception 2"

                    print sys.exc_info()[0], sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                errormsg = '%s' % sys.exc_info()[1]
                errordlg = Dialogs.ErrorDialog(None, errormsg)
                errordlg.ShowModal()
                errordlg.Destroy()

                # Close the Database Cursor
                dbCursor.close()
                # Close the Database Connection
                _dbref.close()
                _dbref = None

    return _dbref

def close_db():
    """ This method flushes all database tables (saving data to disk) and closes the Database Connection. """
    # obtain the Database
    db = get_db()

    if db != None:
        # Close the Database itself
        db.close()

    global _dbref
    # Remove all reference to the database
    _dbref = None


def get_username():
    """Get the name of the current database user."""
    return TransanaGlobal.userName


def list_of_series():
    """Get a list of all Series record names."""
    l = []
    query = "SELECT SeriesNum, SeriesID FROM Series2 ORDER BY SeriesID\n"
    DBCursor = get_db().cursor()
    DBCursor.execute(query)
    for row in fetchall_named(DBCursor):
        id = row['SeriesID']
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        l.append((row['SeriesNum'], id))
    DBCursor.close()
    return l

def list_of_episodes():
    """ Get a list of all Episode records. """
    # Create an empty list to hold results
    l = []
    # Define the Query
    query = "SELECT EpisodeNum, EpisodeID, SeriesNum FROM Episodes2 ORDER BY EpisodeID"
    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Execute the Query
    DBCursor.execute(query)
    # Iterate through the Results set
    for row in fetchall_named(DBCursor):
        # Get the Episode ID
        id = row['EpisodeID']
        # If we're using Unicode ...
        if 'unicode' in wx.PlatformInfo:
            # ... we need to handle the Unicode decoding
            id = ProcessDBDataForUTF8Encoding(id)
        # Add the results to the list
        l.append((row['EpisodeNum'], id, row['SeriesNum']))
    # Close the Database Cursor
    DBCursor.close()
    # Return the list as the function result
    return l

def list_of_episodes_for_series(SeriesName):
    """Get a list of all Episodes contained within a named Series."""
    if 'unicode' in wx.PlatformInfo:
        SeriesName = SeriesName.encode(TransanaGlobal.encoding)
    l = []
    query = """
    SELECT EpisodeNum, EpisodeID, a.SeriesNum FROM Episodes2 a, Series2 b
        WHERE a.SeriesNum = b.SeriesNum AND
              b.SeriesID = %s
        ORDER BY EpisodeID
    """
    DBCursor = get_db().cursor()
    DBCursor.execute(query, SeriesName)
    # Records returned contain EpisodeNum, EpisodeID, and parent Series Num
    for row in fetchall_named(DBCursor):
        id = row['EpisodeID']
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        l.append((row['EpisodeNum'], id, row['SeriesNum']))
    DBCursor.close()
    return l

def list_of_episode_transcripts():
    """ Get a list of all Episode Transcript records. """
    # Create an empty list
    l = []
    # Define the Query.  We only want Episode Transcripts, not Clip Transcripts.
    query = "SELECT TranscriptNum, TranscriptID, EpisodeNum FROM Transcripts2 WHERE ClipNum = 0 ORDER BY TranscriptID"
    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Execute the Query
    DBCursor.execute(query)
    # Iterate through the Results Set
    for row in fetchall_named(DBCursor):
        # Get the Transcript ID
        id = row['TranscriptID']
        # If we're using Unicode ...
        if 'unicode' in wx.PlatformInfo:
            # ... then we need to decode the Transcript ID
            id = ProcessDBDataForUTF8Encoding(id)
        # Add the results to the list
        l.append((row['TranscriptNum'], id, row['EpisodeNum']))
    # Close the Database Cursor
    DBCursor.close()
    # Return the list as the function results
    return l
    
def list_transcripts(SeriesName, EpisodeName):
    """Get a list of all Transcripts for the named Episode within the
    named Series."""
    if 'unicode' in wx.PlatformInfo:
        SeriesName = SeriesName.encode(TransanaGlobal.encoding)
        EpisodeName = EpisodeName.encode(TransanaGlobal.encoding)
    l = []
    query = """
    SELECT TranscriptNum, TranscriptID, a.EpisodeNum FROM Transcripts2 a, Episodes2 b, Series2 c
        WHERE   EpisodeID = %s AND
                a.EpisodeNum = b.EpisodeNum AND
                b.SeriesNum = c.SeriesNum AND
                SeriesID = %s AND
                ClipNum = %s
        ORDER BY TranscriptID
    """
    DBCursor = get_db().cursor()
    DBCursor.execute(query, (EpisodeName, SeriesName, 0))
    for row in fetchall_named(DBCursor):
        id = row['TranscriptID']
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        l.append((row['TranscriptNum'], id, row['EpisodeNum']))
    DBCursor.close()
    return l

def list_clip_transcripts(clipNum):
    """ Return a list of all Clip Transcripts for the given Clip """
    l = []
    query = """ SELECT TranscriptNum, SourceTranscriptNum, SortOrder
                  FROM Transcripts2
                  WHERE ClipNum = %d
                  ORDER BY SortOrder """
    DBCursor = get_db().cursor()
    DBCursor.execute(query % clipNum)
    for row in fetchall_named(DBCursor):
        l.append((row['TranscriptNum'], row['SourceTranscriptNum'], row['SortOrder']))
    DBCursor.close()
    return l

def list_of_collections(ParentNum=0):
    """Get a list of all collections for under the given parent record.  By
    default, the root parent (0) record is used."""
    l = []

    DBCursor = get_db().cursor()
    
    if ParentNum:
        query = """
        SELECT CollectNum, CollectID, ParentCollectNum FROM Collections2
            WHERE   ParentCollectNum = %s
        ORDER BY CollectID
        """
        DBCursor.execute(query, ParentNum)
    else:
        query = """
        SELECT CollectNum, CollectID, ParentCollectNum FROM Collections2
            WHERE   (ParentCollectNum = %s OR ParentCollectNum = %s)
        ORDER BY CollectID
        """
        DBCursor.execute(query, (0, None))

    # This method returns Collection Number, Collection ID, and Parent Colletion Number
    for row in fetchall_named(DBCursor):
        id = row['CollectID']
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        l.append((row['CollectNum'], id, row['ParentCollectNum']))
    DBCursor.close()
    return l

def list_of_all_collections():
    """Get a list of all collections."""
    # Create an empty list
    l = []
    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Define the Query
    query = """ SELECT CollectNum, CollectID, ParentCollectNum FROM Collections2
                  ORDER BY ParentCollectNum, CollectID """
    # Execute the Query
    DBCursor.execute(query)
    # The results set returns Collection Number, Collection ID, and Parent Collection Number.
    # Iterate through the Results Set
    for row in fetchall_named(DBCursor):
        # Get the Collection ID
        id = row['CollectID']
        # If we're using Unicode ...
        if 'unicode' in wx.PlatformInfo:
            # ... we need to decode the Collection ID
            id = ProcessDBDataForUTF8Encoding(id)
        # Add the results to the list
        l.append((row['CollectNum'], id, row['ParentCollectNum']))
    # Close the Database Cursor
    DBCursor.close()
    # Return the List as the function result
    return l

def locate_quick_clips_collection():
    """ Determine the collection number of the Quick Clips Collection, creating it if necessary. """
    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Create a query to get the Collection Number for the QuickClips Collection (Beware of nested copies of
    # the QuickClips Collection, which might've been created by saving search results!)
    query = "SELECT CollectNum from Collections2 where CollectID = %s AND ParentCollectNum = 0"
    # Determine the appropriate name for the QuickClips Collection
    if 'unicode' in wx.PlatformInfo:
        collectionName = unicode(_("Quick Clips"), 'utf8')
    else:
        collectionName = _("Quick Clips")
    if not TransanaConstants.singleUserVersion:
        collectionName += " - %s" % get_username()
    # Execute the query
    DBCursor.execute(query, collectionName.encode(TransanaGlobal.encoding))
    # See if the Quick Clips Collection already exists.  If so, return the Collection Number and False to indicate we didn't create
    # a new collection.
    if DBCursor.rowcount == 1:
        return (DBCursor.fetchone()[0], collectionName, False)
    # If not, we need to create it!
    else:
        import Collection
        tempCollection = Collection.Collection()
        tempCollection.id = collectionName
        tempCollection.parent = 0
        if 'unicode' in wx.PlatformInfo:
            tempCollection.comment = unicode(_('This collection was created automatically to accept Quick Clips.'), 'utf8')
        else:
            tempCollection.comment = _('This collection was created automatically to accept Quick Clips.')
        if not TransanaConstants.singleUserVersion:
            tempCollection.owner = get_username()
        tempCollection.db_save()
        return (tempCollection.number, collectionName, True)    

def list_of_clips():
    """ Get a list of all Clips, regardless of collection. """
    # Create an empty list
    l = []
    # Define the Query
    query = """ SELECT ClipNum, ClipID, CollectNum FROM Clips2
                  ORDER BY SortOrder, ClipID """
    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Execute the Query
    DBCursor.execute(query)
    # Iterate through the Results
    for row in fetchall_named(DBCursor):
        # Get the Clip ID
        id = row['ClipID']
        # If we're using Unicode ...
        if 'unicode' in wx.PlatformInfo:
            # ... we need to decode the ClipID
            id = ProcessDBDataForUTF8Encoding(id)
        # Add the results to the list
        l.append((row['ClipNum'], id, row['CollectNum']))
    # Close the Database Cursor
    DBCursor.close()
    # Return the list as the funtion results
    return l

def list_of_clips_by_collection(CollectionID, ParentNum):
    """Get a list of all Clips for a named Collection."""
    if 'unicode' in wx.PlatformInfo:
        CollectionID = CollectionID.encode(TransanaGlobal.encoding)
    l = []
    if ParentNum:
        subquery = "        b.ParentCollectNum = %s"
        values = (CollectionID, ParentNum)
    else:
        subquery = "    (b.ParentCollectNum = %s OR b.ParentCollectNum = %s)"
        values = (CollectionID, 0, None)
    query = """
    SELECT ClipNum, ClipID, a.CollectNum FROM Clips2 a, Collections2 b
        WHERE a.CollectNum = b.CollectNum AND
              b.CollectID = %%s AND
              %s
        ORDER BY SortOrder, ClipID
    """ % subquery
    DBCursor = get_db().cursor()
    # FIXME: Need try/except block here?
    DBCursor.execute(query, values)
    # This method will return the Clip's Record Number, its ID, and
    # the Record Number for the Parent Collection for each record
    for row in fetchall_named(DBCursor):
        id = row['ClipID']
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        l.append((row['ClipNum'], id, row['CollectNum']))
    DBCursor.close()
    return l

def list_of_clips_by_collectionnum(collectionNum):
    clipList = []
    query = """ SELECT ClipNum, ClipID, CollectNum
                FROM Clips2
                WHERE CollectNum = %s
                ORDER BY SortOrder, ClipID """
    cursor = get_db().cursor()
    cursor.execute(query, collectionNum)
    for (clipNum, clipID, collectNum) in cursor.fetchall():
        id = clipID
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        clipList.append((clipNum, id, collectNum))
    cursor.close()
    return clipList

def list_of_clips_by_episode(EpisodeNum, TimeCode=None):
    """Get a list of all Clips that have been created from a given Collection
    Number.  Optionally restrict list to contain only a given timecode."""
    l = []
    if TimeCode == None:
        query = """
                  SELECT a.ClipNum, a.ClipID, a.CollectNum, a.ClipStart, a.ClipStop, b.CollectID, b.ParentCollectNum, a.ClipComment
                    FROM Clips2 a, Collections2 b
                    WHERE a.CollectNum = b.CollectNum AND
                          a.EpisodeNum = %s
                    ORDER BY a.ClipStart, b.CollectID, a.ClipID
                """
        args = (EpisodeNum)
    else:
        query = """
                  SELECT a.ClipNum, a.ClipID, a.CollectNum, a.ClipStart, a.ClipStop, b.CollectID, b.ParentCollectNum, a.ClipComment
                    FROM Clips2 a, Collections2 b
                    WHERE a.CollectNum = b.CollectNum AND
                          a.EpisodeNum = %s AND 
                          ClipStart <= %s AND 
                          ClipStop > %s
                    ORDER BY a.ClipStart, b.CollectID, a.ClipID
                """
        args = (EpisodeNum, TimeCode, TimeCode)
    DBCursor = get_db().cursor()
    DBCursor.execute(query, args)
    for row in fetchall_named(DBCursor):
        ClipNum = row['ClipNum']
        ClipID = row['ClipID']
        CollectID = row['CollectID']
        # Convert the ID values to the proper UTF-8 representation if needed
        if 'unicode' in wx.PlatformInfo:
            ClipID = ProcessDBDataForUTF8Encoding(ClipID)
            CollectID = ProcessDBDataForUTF8Encoding(CollectID)
        # Add a dictionary object to the results list that spells out the clip data
        l.append({'ClipNum' : ClipNum, 'ClipID' : ClipID, 'ClipStart' : row['ClipStart'], 'ClipStop' : row['ClipStop'], 'CollectID' : CollectID, 'CollectNum' : row['CollectNum'], 'ParentCollectNum' : row['ParentCollectNum'], 'Comment' : row['ClipComment']})

    DBCursor.close()
    return l

def list_of_clips_by_transcriptnum(TranscriptNum):
    """  Get a list of all Clips that have been created from a given Transcript Number.  """
    # Initialize an empty list.
    l = []
    # Define the query that gets clips based on Transcript Number the clips was created from
    query = """ SELECT a.ClipNum, a.ClipID, a.CollectNum, a.ClipStart, a.ClipStop, b.CollectID, b.ParentCollectNum, a.ClipComment, c.SourceTranscriptNum
                  FROM Clips2 a, Collections2 b, Transcripts2 c
                  WHERE a.CollectNum = b.CollectNum AND
                        a.ClipNum = c.ClipNum AND
                        c.SourceTranscriptNum = %s
                  ORDER BY a.ClipStart, b.CollectID, a.ClipID """
    # Set up the query parameters
    args = (TranscriptNum, )
    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Execute the query
    DBCursor.execute(query, args)
    # For each row returned from the database ...
    for row in fetchall_named(DBCursor):
        # Isolate some of the crucial values into variables
        ClipNum = row['ClipNum']
        ClipID = row['ClipID']
        CollectID = row['CollectID']
        # Convert the ID values to the proper UTF-8 representation if needed
        if 'unicode' in wx.PlatformInfo:
            ClipID = ProcessDBDataForUTF8Encoding(ClipID)
            CollectID = ProcessDBDataForUTF8Encoding(CollectID)
        # Add a dictionary object to the results list that spells out the clip data
        l.append({'ClipNum' : ClipNum, 'ClipID' : ClipID, 'ClipStart' : row['ClipStart'], 'ClipStop' : row['ClipStop'], 'CollectID' : CollectID, 'CollectNum' : row['CollectNum'], 'ParentCollectNum' : row['ParentCollectNum'], 'Comment' : row['ClipComment'], 'TranscriptNum' : row['SourceTranscriptNum']})
    # Close the Database Cursor
    DBCursor.close()
    # Return the list of Dictionary Objects
    return l

def list_of_clip_copies(clipID, sourceTranscriptNum, clipStart, clipStop):
    """ Return a list of clips that match the ClipID, source Transcript number, start, and stop times submitted """
    # Create an empty list to hold data
    clipList = []
    # Define the SQL query
    query = """ SELECT c.ClipNum, CollectNum, ClipID, TranscriptNum
                  FROM Clips2 c, Transcripts2 t
                  WHERE ClipID = %s AND
                        SourceTranscriptNum = %s AND
                        c.ClipNum = t.ClipNum AND
                        t.ClipStart = %s AND
                        t.ClipStop = %s"""
    # Define the data to get plugged into the SQL query
    data = (clipID.encode('utf8'), sourceTranscriptNum, clipStart, clipStop)
    # Get a database cursor
    cursor = get_db().cursor()
    # Execute the SQL query
    cursor.execute(query, data)
    # Fetch the data and iterate through it.
    for (clipNum, collectNum, clipID, transcriptNum) in cursor.fetchall():
        # Get the Clip ID
        id = clipID
        # Convert it for UTF-8 if needed
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        # Add the data to the data list
        clipList.append((clipNum, collectNum, id, transcriptNum))
    # Close the database cursor
    cursor.close()
    # Return the data list to the calling routine
    return clipList

def CheckForDuplicateQuickClip(collectNum, episodeNum, transcriptNum, clipStart, clipStop):
    """ Check to see if there is already a Quick Clip for this video segment. """
    # Get a database cursor
    DBCursor = get_db().cursor()
    # If we have a single-transcript Quick Clip, transcriptNum will be a longint value.
    if type(transcriptNum) in [int, long]:
        # Design a query to identify single-transcript Quick Clips which match the data passed in
        query = """SELECT a.ClipNum FROM Clips2 a, Transcripts2 b
                     WHERE CollectNum = %s AND
                           a.EpisodeNum = %s AND
                           b.SourceTranscriptNum = %s AND
                           a.ClipStart = %s AND
                           a.ClipStop = %s AND
                           a.ClipNum = b.ClipNum """
        # Put the data passed in into a compatible data structure
        data = (collectNum, episodeNum, transcriptNum, clipStart, clipStop)
        # Execute the query
        DBCursor.execute(query, data)
        # If no rows are returned ...
        if DBCursor.rowcount == 0:
            # ... close the database cursor ...
            DBCursor.close()
            # ... and return -1 to indicate that no duplicate clips were found
            return -1
        # If duplicate clip(s) are found ...
        else:
            # ... get the Clip Number of the first one ...
            clipNum = DBCursor.fetchone()[0]
            # ... close the database cursor ...
            DBCursor.close()
            # ... and return the Clip Number of the offending clip.
            return clipNum
    # If we have a multi-transcript Quick Clip ...
    else:
        # ... then transcriptNum is a LIST of source transcript numbers.  First, let's sort that list to make comparisons easier.
        transcriptNum.sort()
        # Design a query to identify multi-transcript Quick Clips which match the data passed in.
        # This query returns a row for each transcript for all clips in the same collection, from the same episide with the same
        # starting and stopping points.  If there are different Transcript configurations, these MAY NOT be duplicates!
        query = """SELECT a.ClipNum, SourceTranscriptNum
                     FROM Clips2 a, Transcripts2 b
                     WHERE CollectNum = %s AND
                           a.EpisodeNum = %s AND
                           a.ClipStart = %s AND
                           a.ClipStop = %s AND
                           a.ClipNum = b.ClipNum
                     ORDER BY a.ClipNum, SourceTranscriptNum"""
        # Put the data passed in into a compatible data structure
        data = (collectNum, episodeNum, clipStart, clipStop)

        if DEBUG:
            tmpDlg = wx.MessageDialog(None, query % data, "Hello!")
            tmpDlg.ShowModal()
            tmpDlg.Destroy()
        
        # Execute the query
        DBCursor.execute(query, data)
        # If no rows are returned ...
        if DBCursor.rowcount == 0:
            # ... close the database cursor ...
            DBCursor.close()
            # ... and return -1 to indicate that no duplicate clips were found
            return -1
        # If duplicate clip(s) are found, we have to check to see if they have the same transcript constellation.
        # My methods here are imperfect, but I can't figure out a better way of doing it.
        else:
            # let's initialize a dictionary
            queryTranscripts = {}
            # Let's initialize a variable for holding the clip number we're looking at
            currentClipNum = 0
            # ... get the query results ...
            queryResults = DBCursor.fetchall()
            # Iterate through the query results
            for (cl, tr) in queryResults:
                # If we're looking at a new Clip number ...
                if cl != currentClipNum:
                    # ... update the current clip number ...
                    currentClipNum = cl
                    # ... and create a list for the source transcripts associated with THIS clip.
                    queryTranscripts[cl] = []
                # Add the source transcripts to the clip's Source Transcripts list.
                queryTranscripts[cl].append(tr)
            # Assume we will find no matches, so default the return Clip Number to signal no matches.
            clipNum = -1
            # Iterate through the dictionary's keys
            for key in queryTranscripts.keys():
                # If the transcript configuration for a particular clip matches what was sent in ...
                if queryTranscripts[key] == transcriptNum:
                    # ... then we've found our duplicate clips ...
                    clipNum = key
                    # ... and we're done.
                    break
            # ... close the database cursor ...
            DBCursor.close()
            # ... and return the Clip Number of the offending clip.
            return clipNum

def FindAdjacentClips(episodeNum, startTime, endTime, trNums, trFiles):
    """ Find clips that are adjacent to the time codes sent in.
        Parameters:  episodeNum of the chosen clip
                     startTime of the chosen clip
                     endTime of the chosen clip
                     trNums - source transcript numbers from the selected clip's one or more transcripts """

    # NOTE:  This is a bit complex.  To be considered adjacent, the clip needs to be from the same Episode as clip X,
    #        end where X starts or start where X ends,
    #        and have the same set of transcript sources, ie. same number of transcripts from the same source transcripts.
    #        I can't figure out how to determine this last part through SQL, so I'm doing it programatically.
    #        and have the same set of media files, ie. the same media file and audio inclusion info.
    #        I can't figure out how to determine this through SQL either, so am doing it programmatically.
    
    # Initialize a results set
    results = []
    # Get a database cursor
    DBCursor = get_db().cursor()
    # Define a query
    query = """ SELECT c.ClipNum, ClipID, c.CollectNum, CollectID, MIN(t.ClipStart) ClipStart, MIN(t.ClipStop) ClipStop,
                       COUNT(TranscriptNum) cnt
                  FROM Clips2 c, Transcripts2 t, Collections2 co
                  WHERE (c.EpisodeNum = %s) AND
                        ((t.ClipStop = %s) OR (t.ClipStart = %s)) AND
                        (c.ClipNum = t.ClipNum) AND
                        (c.CollectNum = co.CollectNum)
                  GROUP BY ClipNum
                  ORDER BY ClipStart, ClipID """
    # Define the data for the query
    data = (episodeNum, startTime, endTime)
    # Execute the query
    DBCursor.execute(query, data)
    # Iterate through the query results
    for clipData in DBCursor.fetchall():
        # Actually load the clip found by the query.
        tempClip = Clip.Clip(clipData[0])
        # Assume it's identical until proven otherwise
        identical = True
        # See if the number of transcripts is the same.
        if len(trNums) == len(tempClip.transcripts):
            # Iterate through the sequence of source transcripts
            for x in range(len(trNums)):
                # If the indexed source transcript DON'T match ...
                if trNums[x] != tempClip.transcripts[x].source_transcript:
                    # ... then the clips are not identical ...
                    identical = False
                    # ... and we can stop looking
                    break
        # If not ...
        else:
            # ... FAIL.
            identical = False

        # If we pass the Transcripts test, we need to do the Media Files test
        if identical:
            # See if the number of media files is identical, and
            # the first media files names are identical, and
            # the first audio inclusion flags are identical ...
            if (len(trFiles) == len(tempClip.additional_media_files) + 1) and \
               (trFiles[0][0] == tempClip.media_filename) and \
               (trFiles[0][1] == tempClip.audio):
                # Initialize the counter to 1, as we've already checked element 0!
                cnt = 1
                # Iterate through the additional media files ...
                for addFile in tempClip.additional_media_files:
                    # ... and compare them to the data passed in from the original clip
                    if (trFiles[cnt][0] != addFile['filename']) or (trFiles[cnt][1] != addFile['audio']):
                        # If they differ, the comparison FAILS ...
                        identical = False
                        # ... and we can stop looking at additional media files
                        break
                    # Increment the comparison counter
                    cnt += 1
            # If not ...
            else:
                # ... FAIL
                identical = False

        # If no transcript sources or media file sources were NOT identical ...
        if identical:
            # ... then include the query data in the results set.  Note that we expand the collection name to
            # include the full collection path.
            results.append((clipData[0],) + (ProcessDBDataForUTF8Encoding(clipData[1]),) + (clipData[2:3] + (tempClip.GetNodeString(includeClip=False),) + clipData[4:]))
    # Return the results set
    return results

def getMaxSortOrder(collNum):
    """Get the largest Sort Order value for all the Clips in a Collection."""
    DBCursor = get_db().cursor()
    query = "SELECT MAX(SortOrder) FROM Clips2 WHERE CollectNum = %s" 
    DBCursor.execute(query, collNum)
    if DBCursor.rowcount >= 1:
        maxSortOrder = DBCursor.fetchone()[0]
        # Dropping a clip into a new collection produces a maxSortOrder of None, rather than rowcount being 0!
        if maxSortOrder == None:
            maxSortOrder = 0
    else:
        maxSortOrder = 0
    DBCursor.close()
    return maxSortOrder


def list_of_notes(** kwargs):
    """Get a list of Note IDs for the given Series, Episode, Collection,
    or Clip numbers.  Parameters are passed as keyword
    arguments, where valid parameters are Transcript, Episode, Series, Clip, and
    Collection.
    
    Examples: list_of_notes(Series=12)
              list_of_notes(Episode=14)

    The optional parameter "includeNumber=True" causes (note number, note id) tuples
    to be returned rather than just a list of note ids. """
    # Initialize the note list to hold the query results
    notelist = []
    # Build the query
    query = "SELECT NoteNum, NoteID From Notes2"
    if kwargs.has_key("Series"):
        query += " WHERE   SeriesNum = %s"
        values = (kwargs['Series'],)
    elif kwargs.has_key("Episode"):
        query += " WHERE   EpisodeNum = %s"
        values = (kwargs['Episode'],)
    elif kwargs.has_key("Transcript"):
        query += " WHERE   TranscriptNum = %s"
        values = (kwargs['Transcript'],)
    elif kwargs.has_key("Collection"):
        query += " WHERE   CollectNum = %s"
        values = (kwargs['Collection'],)
    elif kwargs.has_key("Clip"):
        query += " WHERE   ClipNum = %s"
        values = (kwargs['Clip'],)
    else:
        return []   # Should we raise an exception?
    query = query + " ORDER BY NoteID"
    # Get a database
    db = get_db()
    # Get a database cursor
    DBCursor = db.cursor()
    # Execute our query
    DBCursor.execute(query % values)
    # Get the results set
    r = DBCursor.fetchall()
    # Iterate through the records in the query results
    for tup in r:
        # Identify the Note ID
        id = tup[1]
        # Convert for Unicode, if needed
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        # If we want both the Note Number and Note ID ...
        if kwargs.has_key("includeNumber"):
            # ... add both elements to the Note List in a tuple
            notelist.append((tup[0], id))
        # Otherwise ...
        else:
            # ... just add the Note ID to the Note List
            notelist.append(id)
    # Close the database cursor
    DBCursor.close()
    # Return the Note List
    return notelist

def list_of_node_notes(** kwargs):
    """ Get a list of all Notes for the given Series or Collection node, including sub-nodes.
        Valid parameters are SeriesNode=True or CollectionNode=True."""
    # Create an empty list
    notelist = []
    # Start building the Query
    query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum
                 FROM Notes2 """
    # If we're looking for Series Node Notes ...
    if kwargs.has_key("SeriesNode"):
        # ... we need to build a query for Series, Episode, or Transcript Notes
        query += """WHERE   SeriesNum <> 0 OR
                            EpisodeNum <> 0 OR
                            TranscriptNum <> 0"""
    # If we're looking for Collection Node Notes ...
    elif kwargs.has_key("CollectionNode"):
        # ... we need to build a query for Collection or Clip Notes
        query += """WHERE   CollectNum <> 0 OR
                            ClipNum <> 0"""
    # If neither SeriesNode nor CollectionNode is defined, we've got a programming error.
    else:
        return []   # Should we raise an exception?
    # Finish the Query
    query = query + "   ORDER BY NoteID\n"
    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Execute the Query
    DBCursor.execute(query)
    # Get the results set
    r = DBCursor.fetchall()
    # Iterate through the Results Set
    for tup in r:
        # Get the Note ID
        id = tup[1]
        # If we're working with Unicode ...
        if 'unicode' in wx.PlatformInfo:
            # ... we need to decode the Note ID
            id = ProcessDBDataForUTF8Encoding(id)
        # Add the results to the list
        notelist.append((tup[0], id) + tup[2:])
    # Close the Database Cursor
    DBCursor.close()
    # Return the list as the function results
    return notelist

def list_of_all_notes(reportType=None, searchText=None):
    """ Get a list of all Notes for the Notes Browser """
    # initialize the Notes List as empty
    notelist = []

    # We want to display all the Notes in each section in alphabetical order.

    # Query for ALL Notes in order of NoteID.
    query = """ SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, NoteTaker
                FROM Notes2 N"""
    # If we want the Series report, limit the query to Series notes
    if reportType == 'SeriesNode':
        query += " WHERE SeriesNum <> 0"
    # If we want the Episode report, limit the query to Episode notes
    elif reportType == 'EpisodeNode':
        query += " WHERE EpisodeNum <> 0"
    # If we want the Transcript report, limit the query to Transcript notes
    elif reportType == 'TranscriptNode':
        query += " WHERE TranscriptNum <> 0"
    # If we want the Collection report, limit the query to Collection notes
    elif reportType == 'CollectionNode':
        query += " WHERE CollectNum <> 0"
    # If we want the Clip report, limit the query to Clip notes
    elif reportType == 'ClipNode':
        query += " WHERE ClipNum <> 0"
    # If searchText is passed in, we want to limit the results to notes containing that text.
    # We need to add that to our Query
    if searchText != None:
        if (reportType != None):
            query += " AND "
        else:
            query += " WHERE "
        query += "LOWER(CAST(NoteText AS CHAR)) like '%%%s%%'" % searchText.lower().encode(TransanaGlobal.encoding)
        
    # We always want to sort by NoteID
    query += " ORDER BY NoteID"
    # Make sure we have a Database connection
    db = get_db()
    # Get a database cursor
    DBCursor = db.cursor()
    # Execute the query
    DBCursor.execute(query)
    # Get the Results Set
    results = DBCursor.fetchall()
    # For each row in the results set ...
    for row in results:
        # Pull out the elements that need to be encoded
        ID = row[1]
        noteTaker = row[7]
        # Encode the elements, if needed
        if 'unicode' in wx.PlatformInfo:
            ID = ProcessDBDataForUTF8Encoding(ID)
            noteTaker = ProcessDBDataForUTF8Encoding(noteTaker)
        # Create a Dictionary Object to be added to the Notes List
        notelist.append({'NoteNum' : row[0],
                         'NoteID' : ID,
                         'SeriesNum' : row[2],
                         'EpisodeNum' : row[3],
                         'TranscriptNum' : row[4],
                         'CollectionNum' : row[5],
                         'ClipNum' : row[6],
                         'NoteTaker' : noteTaker})
    # Close the Database Cursor
    DBCursor.close()
    # Return the Note List as the Function Result
    return notelist

def list_of_keyword_groups():
    """Get a list of all keyword groups."""
    l = []
    query = "SELECT KeywordGroup FROM Keywords2 GROUP BY KeywordGroup\n"
    DBCursor = get_db().cursor()
    DBCursor.execute(query)
    for row in fetchall_named(DBCursor):
        id = row['KeywordGroup']
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        l.append(id)
    DBCursor.close()
    return l

def list_of_keywords_by_group(KeywordGroup):
    """Get a list of all keywords for the named Keyword group."""
    if 'unicode' in wx.PlatformInfo:
        KeywordGroup = KeywordGroup.encode(TransanaGlobal.encoding)
    l = []
    query = \
    "SELECT Keyword FROM Keywords2 WHERE KeywordGroup = %s ORDER BY Keyword\n"
    DBCursor = get_db().cursor()
    DBCursor.execute(query, KeywordGroup)
    for row in fetchall_named(DBCursor):
        id = row['Keyword']
        if 'unicode' in wx.PlatformInfo:
            id = ProcessDBDataForUTF8Encoding(id)
        l.append(id)
    DBCursor.close()
    return l

def list_of_all_keywords():
    """Get a list of all keywords in the Transana database."""
    # Create an empty list
    l = []
    # Define the Query
    query = "SELECT KeywordGroup, Keyword FROM Keywords2 ORDER BY KeywordGroup, Keyword"
    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Execute the Query
    DBCursor.execute(query)
    # Iterate through the results
    for row in fetchall_named(DBCursor):
        # Get the Keyword Group
        kwg = row['KeywordGroup']
        # Get the Keyword
        kw = row['Keyword']
        # If we're using Unicode ...
        if 'unicode' in wx.PlatformInfo:
            # ... we need to decode the Keyword Group and Keyword
            kwg = ProcessDBDataForUTF8Encoding(kwg)
            kw = ProcessDBDataForUTF8Encoding(kw)
        # Add the results to the list
        l.append((kwg, kw))
    # Close the database cursor
    DBCursor.close()
    # return the list as the function results
    return l
   
def list_of_keywords(** kwargs):
    """Get a list of all keywordgroup/keyword pairs for the specified
    qualifiers (Episode, Clip numbers).  Result is a list of tuples,
    where the first element in the tuple is the keyword group, 
    the second element is the keyword itself, and the third element
    indicates whether the keyword is an example or not.

    examples: list_of_keywords(Episode=5)
              list_of_keywords(Clip=1)
              list_of_keywords(Episode=5, Clip=1)
    """
    
    count = len(kwargs)
    i = 1
    query = "SELECT * FROM ClipKeywords2\n"
    for obj in kwargs:
        query = query + "   WHERE %sNum = %%s" % (obj)
        if i != count:      # not last item
            query = query + " AND \n"
        else:
            query = query + "\n"
        i += 1
    query = query + "    ORDER BY KeywordGroup, Keyword\n"
    DBCursor = get_db().cursor()
    if len(kwargs) > 0:
        DBCursor.execute(query, kwargs.values())
    else:
        DBCursor.execute(query)
    r = DBCursor.fetchall()
    kwlist = []
    # Current ClipKeywords table row format used:
    # EpNum, ClipNum, KWGroup, Keyword, Example
    for tup in r:
        if 'unicode' in wx.PlatformInfo:
            kwlist.append((ProcessDBDataForUTF8Encoding(tup[2]), \
                           ProcessDBDataForUTF8Encoding(tup[3]), \
                           ProcessDBDataForUTF8Encoding(tup[4])))
        else:
            kwlist.append((tup[2], tup[3], tup[4]))
    DBCursor.close()
    return kwlist

def list_of_keyword_examples():
    """Get a list of all Keyword Examples from the ClipKeywords table."""
    
    query = "SELECT * FROM ClipKeywords2 WHERE Example = 1"
    dbCursor = get_db().cursor()
    dbCursor.execute(query)
    results = dbCursor.fetchall()
    keywordExampleList = []
    # Current ClipKeywords table row format used:
    # EpNum, ClipNum, KWGroup, Keyword, Example
    for dbRowData in results:
        kwg = dbRowData[2]
        kw = dbRowData[3]
        if 'unicode' in wx.PlatformInfo:
            kwg = ProcessDBDataForUTF8Encoding(kwg)
            kw = ProcessDBDataForUTF8Encoding(kw)
        keywordExampleList.append((dbRowData[0], dbRowData[1], kwg, kw, dbRowData[4]))
    dbCursor.close()
    return keywordExampleList

def SetKeywordExampleStatus(kwg, kw, clipNum, exampleValue):
    """ The SetKeywordExampleStatus(kwg, kw, clipNum, exampleValue) method sets the
        Example value in the ClipKeywords table for the appropriate KWG, KW, Clip
        combination.  Set Example to 1 to specify a Keyword Example, 0 to remove
        it from being an example without deleting the keyword for the Clip. """
    if 'unicode' in wx.PlatformInfo:
        tempkwg = kwg.encode(TransanaGlobal.encoding)
        tempkw = kw.encode(TransanaGlobal.encoding)
    query = """UPDATE ClipKeywords2
                 SET Example = %s
                 WHERE KeywordGroup = %s AND
                       Keyword = %s AND
                       ClipNum = %s"""
    dbCursor = get_db().cursor()
    dbCursor.execute(query, (exampleValue, tempkwg, tempkw, clipNum))

    # If rowcount == 0, no rows were affected by the update.  That is, the Keyword
    # had not previously been assigned to this Clip, so there was no record to update.
    # In this case, we have to ADD the record!
    if dbCursor.rowcount == 0:
        insert_clip_keyword(0, clipNum, kwg, kw, 1)
    dbCursor.close()


def check_username_as_keyword():
    """ Determine if the username is already a keyword, creating it if necessary. """
    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Create a query to get the Collection Number for the QuickClips Collection
    query = "SELECT * from Keywords2 where KeywordGroup = %s AND Keyword = %s"
    # Determine the appropriate Keyword Group and Keyword
    data = (_("Transana Users"), get_username())
    # Execute the query
    DBCursor.execute(query, data)
    # See if the keyword already exists.  If not, we need to create it.
    if DBCursor.rowcount == 0:
        import Keyword
        tempKeyword = Keyword.Keyword()
        if 'unicode' in wx.PlatformInfo:
            tempKeyword.keywordGroup = unicode(_("Transana Users"), 'utf8')
        else:
            tempKeyword.keywordGroup = _("Transana Users")
        tempKeyword.keyword = get_username()
        tempKeyword.db_save()
        # Return True to indicate that a keyword was created
        return True
    else:
        # Return False to indicate that no keyword was created
        return False


def VideoFilePaths(filePath, update=False):
    """ This method returns the number of Collections and Clips that would be affected by
        implementing the Video Root Path, and optionally makes the changes. """

    # Okay, this wasn't simple.  It's convoluted, but the simple way doesn't work.
    #
    # First, I can't make the query "SELECT * FROM Episodes2 WHERE MediaFile LIKE 'V:\Demo\%'" work.
    # Something about the "%" character breaks it, and I tried many things that didn't help.
    #
    # Second, files appear to be stored in the database inconsistently, sometimes with a single
    # separator character and sometimes with a pair.  This is probably a difference between how
    # Delphi and Python store the path separators as part of file names.  But the end result is that
    # the SQL "LIKE" comparison fails sometimes when it should succeed because "V:\Demo" isn't
    # LIKE "V:\\Demo".
    #
    # The solution I've come up with is to pull ALL the file names and do the comparisons manually.
    # You are welcome to optimize this code if you can figure out a better way.
    #
    # David Woods
    # 1/27/2004
    # If the filePath is empty, just return.  This happens when the user deletes the Video Path.
    if filePath == '':
        return (0, 0)

    # Get a Database Cursor
    dbCursor = get_db().cursor()
    # If update is True, but some records are locked, we will need to know that the Database Transaction
    # needs to be rolled back.  If it fails early, we may skip some of the processing that becomes clearly
    # pointless.  Here, we declare a variable to track whether we should continue with the Transaction.
    # It needs to be declared regardless of update status.
    transactionStatus = True
    # If we are updating records, begin a Transaction so that everything can be undone
    # if we run into problems.
    if update:
        dbCursor.execute("BEGIN")

    # Initialize the Episode Counter        
    episodeCount = 0
    # Create the Query for the Episode Table
    query = "SELECT EpisodeNum, MediaFile FROM Episodes2 "
    # Execute the Query
    dbCursor.execute(query)
    # Fetch all the Database Results, and process them row by row
    for (episodeNum, mediafile) in dbCursor.fetchall():
        # If we're using Unicode ...
        if 'unicode' in wx.PlatformInfo:
            # ... then encode the file names appropriately
            mediafile = ProcessDBDataForUTF8Encoding(mediafile)
        # Some Filenames have doubled backslashes, though not all do.  Let's eliminate them if they are present.
        # (Python requires a double backslash in a string to represent a single backslash, so this replaces double
        # backslashes ('\\') with single ones ('\') even though it looks like it replaces quadruples with doubles.)
        mediafile = string.replace(mediafile, '\\\\', '\\')
        # Now replace the backslash with the more universal slash character in both the file name and the filePath
        mediafile = string.replace(mediafile, '\\', '/')
        filePath = string.replace(filePath, '\\', '/')
        # Compare the Video Root filePath passed in with the front portion of the File Name from the Database.
        if filePath == mediafile[:len(filePath)]:
            # If they are the same, increment the Episode Counter
            episodeCount += 1
            # If update is True, we should update the records we find.
            if update:
                # Load the Episode using the Episode Number.
                tempEpisode = Episode.Episode(episodeNum)
                # We need a "try .. except" block to catch record lock exceptions
                try:
                    # Try to lock the Episode
                    tempEpisode.lock_record()
                    # Remove the Video Root from the File Name
                    tempEpisode.media_filename = mediafile[len(filePath):]
                    # Save the Episode
                    tempEpisode.db_save()
                    # Unlock the Episode
                    tempEpisode.unlock_record()
                # Catch failed record locks or Saves
                except:
                    # TODO:  Detect exception type and customize the error message below.
                    print sys.exc_info()[0], sys.exc_info()[1]

                    # If it fails, set the transactionStatus Flag to False
                    transactionStatus = False
                    # and stop processing records.
                    break

    # If we've already failed, there's no reason to continue trying.  Make sure we should continue.
    if transactionStatus:
        # Initialize the Clip Counter
        clipCount = 0
        # Create the Query for the Episode Table
        query = "SELECT ClipNum, MediaFile FROM Clips2 "
        # Execute the Query
        dbCursor.execute(query)
        # Fetch all the Database Results, and process them row by row
        for (clipNum, mediafile) in dbCursor.fetchall():
            # If we're using Unicode ...
            if 'unicode' in wx.PlatformInfo:
                # ... then encode the file names appropriately
                mediafile = ProcessDBDataForUTF8Encoding(mediafile)
            # Some Filenames have doubled backslashes, though not all do.  Let's eliminate them if they are present.
            # (Python requires a double backslash in a string to represent a single backslash, so this replaces double
            # backslashes ('\\') with single ones ('\') even though it looks like it replaces quadruples with doubles.)
            mediafile = string.replace(mediafile, '\\\\', '\\')
            # Now replace the backslash with the more universal slash character in both the file name and the filePath
            mediafile = string.replace(mediafile, '\\', '/')
            filePath = string.replace(filePath, '\\', '/')
            # Compare the Video Root filePath passed in with the front portion of the File Name from the Database.
            if filePath == mediafile[:len(filePath)]:
                # If they are the same, increment the Clip Counter
                clipCount += 1
                # If update is True, we should update the record we find.
                if update:
                    # Load the Clip using the Clip Number.
                    tempClip = Clip.Clip(clipNum)
                    # We need a "try .. except" block to catch record lock exceptions
                    try:
                        # Try to lock the Episode
                        tempClip.lock_record()
                        # Remove the Video Root from the File Name
                        tempClip.media_filename = mediafile[len(filePath):]
                        # Save the Clip
                        tempClip.db_save()
                        # Unlock the Clip
                        tempClip.unlock_record()
                    # Catch failed record locks or Saves
                    except:
                        # TODO:  Detect exception type and customize the error message below.
                        print sys.exc_info()[0], sys.exc_info()[1]
                        import traceback
                        traceback.print_exc(file=sys.stdout)

                        # If it fails, set the transactionStatus Flag to False
                        transactionStatus = False
                        # and stop processing records.
                        break

    if transactionStatus:
        # Create the Query for the Episode Table
        query = "SELECT EpisodeNum, ClipNum, MediaFile FROM AdditionalVids2 "
        # Execute the Query
        dbCursor.execute(query)
        # Fetch all the Database Results, and process them row by row
        for (episodeNum, clipNum, mediafile) in dbCursor.fetchall():
            # If we're using Unicode ...
            if 'unicode' in wx.PlatformInfo:
                # ... then encode the file names appropriately
                mediafile = ProcessDBDataForUTF8Encoding(mediafile)
            # Some Filenames have doubled backslashes, though not all do.  Let's eliminate them if they are present.
            # (Python requires a double backslash in a string to represent a single backslash, so this replaces double
            # backslashes ('\\') with single ones ('\') even though it looks like it replaces quadruples with doubles.)
            mediafile = string.replace(mediafile, '\\\\', '\\')
            # Now replace the backslash with the more universal slash character in both the file name and the filePath
            mediafile = string.replace(mediafile, '\\', '/')
            filePath = string.replace(filePath, '\\', '/')
            # Compare the Video Root filePath passed in with the front portion of the File Name from the Database.
            if filePath == mediafile[:len(filePath)]:
                # If they are the same, increment the appropriate Counter
                if clipNum > 0:
                    clipCount += 1
                else:
                    episodeCount += 1
                # If update is True, we should update the record we find.
                if update:
                    if clipNum > 0:
                        # Load the Clip using the Clip Number.
                        tempClip = Clip.Clip(clipNum)
                        # We need a "try .. except" block to catch record lock exceptions
                        try:
                            # Try to lock the Episode
                            tempClip.lock_record()
                            # Remove the Video Root from the File Name
                            # Start by initializing a file counter to 0
                            filCount = 0
                            # Iterate through the Additional Media Files ...
                            for fil in tempClip.additional_media_files:
                                # ... if the data from the database matches the data passed in by parameter ...
                                if (mediafile.upper().replace('/', os.sep) == fil['filename'].upper()):
                                    # ... then rename the Clip's Additional Media File
                                    tempClip.additional_media_files[filCount]['filename'] = mediafile[len(filePath):]
                                    # Iterate the file counter
                                    filCount += 1
                            # Save the Clip
                            tempClip.db_save()
                            # Unlock the Clip
                            tempClip.unlock_record()
                        # Catch failed record locks or Saves
                        except:
                            # TODO:  Detect exception type and customize the error message below.
                            print sys.exc_info()[0], sys.exc_info()[1]
                            import traceback
                            traceback.print_exc(file=sys.stdout)

                            # If it fails, set the transactionStatus Flag to False
                            transactionStatus = False
                            # and stop processing records.
                            break
                    else:
                       # Load the Episode using the Episode Number.
                        tempEpisode = Episode.Episode(episodeNum)
                        # We need a "try .. except" block to catch record lock exceptions
                        try:
                            # Try to lock the Episode
                            tempEpisode.lock_record()
                            # Remove the Video Root from the File Name
                            # Start by initializing a file counter to 0
                            filCount = 0
                            # Iterate through the Additional Media Files ...
                            for fil in tempEpisode.additional_media_files:
                                # ... if the data from the database matches the data passed in by parameter ...
                                if (mediafile.upper().replace('/', os.sep) == fil['filename'].upper()):
                                    # ... then rename the Clip's Additional Media File
                                    tempEpisode.additional_media_files[filCount]['filename'] = mediafile[len(filePath):]
                                    # Iterate the file counter
                                    filCount += 1
                            # Save the Episode
                            tempEpisode.db_save()
                            # Unlock the Episode
                            tempEpisode.unlock_record()
                        # Catch failed record locks or Saves
                        except:
                            # TODO:  Detect exception type and customize the error message below.
                            print sys.exc_info()[0], sys.exc_info()[1]

                            # If it fails, set the transactionStatus Flag to False
                            transactionStatus = False
                            # and stop processing records.
                            break

    # If we are updating data...
    if update:
        # Check to see if we have been successful.
        if transactionStatus:
            # If so, commit the changes to the Database.
            dbCursor.execute('COMMIT')
        else:
            dlg = Dialogs.ErrorDialog(TransanaGlobal.menuWindow, _('An error occurred.  Most likely, a record was locked by another user.\nValues in the Database were not updated.'))
            dlg.ShowModal()
            dlg.Destroy()
            # If not, roll back the database changes.
            dbCursor.execute('ROLLBACK')
            # If we fail, no records will have been changed, so zero out both counters.
            episodeCount = 0
            clipCount = 0
    # Close the Database Cursor
    dbCursor.close()
    # Return the number of records found or updated.
    return (episodeCount, clipCount)

def IsDatabaseEmpty():
    """ Returns True if the database is empty, False if there are ANY data records """
    result = True
    db = get_db()
    dbCursor = db.cursor()
    tables = ('Series2', 'Collections2', 'CoreData2', 'Keywords2')
    SQLText = "SELECT * FROM %s"
    for table in tables:
        dbCursor.execute(SQLText % table)
        if dbCursor.rowcount > 0:
            result = False
            break
    dbCursor.close()
    return result
        
def fetch_named(cursor, row_result=None):
    """Fetch a row result from the cursor object, but return it as a dictionary
    including the database field names.  Optionally specify an already-fetched
    row result by passing the optional `row_result' as something other than
    None."""
    d = cursor.description
    if row_result == None:
        row_result = cursor.fetchone()
    dict = {}
    if not d:
        return dict
    for c, r in map(None, d, row_result):
        dict[c[0]] = r

    return dict

def fetchall_named(cursor):
    """Fetch all row results from the cursor object, and return it as a 
    sequence of dictionaries including the database field names."""
    d = cursor.description
    rows = cursor.fetchall()
    if not d:
        return ()
    l = []
    for row in rows:
        dict = {}
        for c, r in map(None, d, row):
            dict[c[0]] = r
        l.append(dict)
    return l

def list_all_keyword_examples_for_all_clips_in_a_collection(collectionNum):
    """ Lists all Keyword Examples for all Clips in the specified Collection and all
        nested Collections recursively """
    # Start with an empty list
    kwExamples = []
    # Get a list of all Clips in the Collection
    for (clipNum, clipID, collectNum) in list_of_clips_by_collectionnum(collectionNum):
        # Get a list of all Keyword Examples for each Clip
        kwExamples = kwExamples + list_all_keyword_examples_for_a_clip(clipNum)
    # Now get a list of Nested Collections
    collectionList = list_of_collections(collectionNum)
    # For each Nested Collection ...
    for collection in collectionList:
        # ... add its Clip Keyword Examples to the Keyword Example List recursively
        kwExamples = kwExamples + list_all_keyword_examples_for_all_clips_in_a_collection(collection[0])
    # Return the Keyword Example List to the calling routine
    return kwExamples


def list_all_keyword_examples_for_a_clip(clipnum):
    """ Lists all Keyword Examples for the specified Clip. """
    # Start with an empty list
    kwExamples = []
    # If clipnum = 0, all Episode Keywords would be deleted.  This would be BAD.
    if clipnum != 0:
        # Create a query to find all Keyword Examples for the specified Clip
        query = """ SELECT KeywordGroup, Keyword, CK.ClipNum, ClipID
                    FROM ClipKeywords2 CK, Clips2 C
                    WHERE CK.ClipNum = %s AND
                          Example = 1 AND
                          CK.ClipNum = C.ClipNum """
        # Get a Database Cursor
        cursor = get_db().cursor()
        # Execute the query
        cursor.execute(query, clipnum)
        # Iterate through the cursor's result set and put all Keyword Group : Keyword pairs in the Keyword List
        for (kwg, kw, clipNumber, clipID) in cursor.fetchall():
            kwg = ProcessDBDataForUTF8Encoding(kwg)
            kw = ProcessDBDataForUTF8Encoding(kw)
            clipID = ProcessDBDataForUTF8Encoding(clipID)
            kwExamples.append((kwg, kw, clipNumber, clipID))
        # Close the Database Cursor
        cursor.close()
    # Return the list of Keyword Examples as the function result
    return kwExamples

def delete_all_keywords_for_a_group(epnum, clipnum):
    """Given an Episode or a Clip number, delete the keywordgroup/word
    pairs."""
    if (epnum == 0) and (clipnum == 0):
        raise Exception, _("All keywords would have been deleted!")
    
    if epnum != 0:
        specifier = "EpisodeNum"
        num = epnum
    else:
        # clipnum must be non-zero
        specifier = "ClipNum"
        num = clipnum
        

    query = """
    DELETE FROM ClipKeywords2
        WHERE %s = %%s
    """ % (specifier)
    DBCursor = get_db().cursor()
    DBCursor.execute(query, num)
    DBCursor.close()

def insert_clip_keyword(ep_num, clip_num, kw_group, kw, exampleValue=0):
    """Insert a new record in the Clip Keywords table."""
    if 'unicode' in wx.PlatformInfo:
        kw_group = kw_group.encode(TransanaGlobal.encoding)
        kw = kw.encode(TransanaGlobal.encoding)
    query = """
    INSERT INTO ClipKeywords2
        (EpisodeNum, ClipNum, KeywordGroup, Keyword, Example)
        VALUES
        (%s,%s,%s,%s, %s)
    """
    DBCursor = get_db().cursor()
    DBCursor.execute(query, (ep_num, clip_num, kw_group, kw, exampleValue))
    DBCursor.close()

def add_keyword(group, kw_name):
    """Add a keyword to the database."""
    DBCursor = get_db().cursor()
    query = """INSERT INTO Keywords2
        (KeywordGroup, Keyword)
        VALUES (%s,%s)
    """
    DBCursor.execute(query, (group, kw_name))
    DBCursor.close()

def delete_keyword_group(name):
    """Delete a Keyword Group from the database, including all associated
    keywords."""

    if 'unicode' in wx.PlatformInfo:
        kwg = name.encode(TransanaGlobal.encoding)

    DBCursor = get_db().cursor()
    DBCursor.execute("BEGIN")
    t = ""
    query = """SELECT KeywordGroup, Keyword, RecordLock
        FROM Keywords2
        WHERE   KeywordGroup = %s AND
                RecordLock <> %s
    """

    DBCursor.execute(query, (kwg, ""))
    for row in fetchall_named(DBCursor):
        tempkwg = row['KeywordGroup']
        tempkw = row['Keyword']
        temprl = row['RecordLock']
        if 'unicode' in wx.PlatformInfo:
            tempkwg = ProcessDBDataForUTF8Encoding(tempkwg)
            tempkw = ProcessDBDataForUTF8Encoding(tempkw)
            temprl = ProcessDBDataForUTF8Encoding(temprl)
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('%s  Keyword "%s : %s" is locked by %s\n'), 'utf8')
        else:
            prompt = _('%s  Keyword "%s : %s" is locked by %s\n')
        t = prompt % (t, tempkwg, tempkw, temprl)

    query = """SELECT a.KeywordGroup, a.Keyword, b.EpisodeID, b.RecordLock
        FROM ClipKeywords2 a, Episodes2 b
        WHERE   a.KeywordGroup = %s AND
                a.EpisodeNum <> %s AND
                a.EpisodeNum = b.EpisodeNum AND
                b.RecordLock <> %s
    """

    DBCursor.execute(query, (kwg, 0, ""))
    for row in fetchall_named(DBCursor):
        tempepid = row['EpisodeID']
        temprl = row['RecordLock']
        if 'unicode' in wx.PlatformInfo:
            tempepid = ProcessDBDataForUTF8Encoding(tempepid)
            temprl = ProcessDBDataForUTF8Encoding(temprl)
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('%s  Episode "%s" is locked by %s\n'), 'utf8')
        else:
            prompt = _('%s  Episode "%s" is locked by %s\n')
        t = prompt % (t, tempepid, temprl)

    # Clips next
    query = """SELECT a.KeywordGroup, a.Keyword, c.ClipID, c.RecordLock
        FROM ClipKeywords2 a, Clips2 c
        WHERE   a.KeywordGroup = %s AND
                a.ClipNum <> %s AND
                a.ClipNum = c.ClipNum AND
                c.RecordLock <> %s
    """
    DBCursor.execute(query, (kwg, 0, ""))
    for row in fetchall_named(DBCursor):
        tempclid = row['ClipID']
        temprl = row['RecordLock']
        if 'unicode' in wx.PlatformInfo:
            tempclid = ProcessDBDataForUTF8Encoding(tempclid)
            temprl = ProcessDBDataForUTF8Encoding(temprl)
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('%s  Clip "%s" is locked by %s\n'), 'utf8')
        else:
            prompt = _('%s  Clip "%s" is locked by %s\n')
        t = prompt % (t, tempclid, temprl)

    if t == "":
        # Delphi Transana had a confirmation dialog here, but we won't do
        # that here (do it before calling this function).

        # Build and execute the SQL to delete a word from the list.
        query = """DELETE FROM Keywords2
            WHERE   KeywordGroup = %s"""
        DBCursor.execute(query, kwg)
        # Now delete all instances of this keywordgroup/keyword combo in the
        # Clipkeywords file
        query = """DELETE FROM ClipKeywords2
            WHERE   KeywordGroup = %s"""
        DBCursor.execute(query, (kwg))

        # Finish the transaction
        DBCursor.execute("COMMIT")
    else:
        DBCursor.execute("ROLLBACK")
        DBCursor.close()
        msg = _('Unable to delete keyword group "%s".\n%s')
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            msg = unicode(msg, 'utf8')
        msg = msg  % (name, t)
        raise TransanaExceptions.GeneralError, msg
    DBCursor.close()

def delete_keyword(group, kw_name):
    """Delete a Keyword from the database."""
    if 'unicode' in wx.PlatformInfo:
        kwg = group.encode(TransanaGlobal.encoding)
        kw = kw_name.encode(TransanaGlobal.encoding)
    else:
        kwg = group
        kw = kw_name
    DBCursor = get_db().cursor()
    DBCursor.execute("BEGIN")
    t = ""
    query = """SELECT KeywordGroup, Keyword, RecordLock
        FROM Keywords2
        WHERE   KeywordGroup = %s AND
                Keyword = %s AND
                RecordLock <> %s
    """

    DBCursor.execute(query, (kwg, kw, ""))
    for row in fetchall_named(DBCursor):
        msg = _('%s  Keyword "%s : %s" is locked by %s\n')
        tempkwg = row['KeywordGroup']
        tempkw = row['Keyword']
        temprl = row['RecordLock']
        if 'unicode' in wx.PlatformInfo:
            msg = unicode(msg, 'utf8')
            tempkwg = ProcessDBDataForUTF8Encoding(tempkwg)
            tempkw = ProcessDBDataForUTF8Encoding(tempkw)
            temprl = ProcessDBDataForUTF8Encoding(temprl)
        t =  msg % (t, tempkwg, tempkw, temprl)

    query = """SELECT a.KeywordGroup, a.Keyword, b.EpisodeID, b.RecordLock
        FROM ClipKeywords2 a, Episodes2 b
        WHERE   a.KeywordGroup = %s AND
                a.Keyword = %s AND
                a.EpisodeNum <> %s AND
                a.EpisodeNum = b.EpisodeNum AND
                b.RecordLock <> %s
    """

    DBCursor.execute(query, (kwg, kw, 0, ""))
    for row in fetchall_named(DBCursor):
        msg = _('%s  Episode "%s" is locked by %s\n')
        tempepid = row['EpisodeID']
        temprl = row['RecordLock']
        if 'unicode' in wx.PlatformInfo:
            msg = unicode(msg, 'utf8')
            tempepid = ProcessDBDataForUTF8Encoding(tempepid)
            temprl = ProcessDBDataForUTF8Encoding(temprl)
        t = msg % (t, tempepid, temprl)

    # Clips next
    query = """SELECT a.KeywordGroup, a.Keyword, c.ClipID, c.RecordLock
        FROM ClipKeywords2 a, Clips2 c
        WHERE   a.KeywordGroup = %s AND
                a.Keyword = %s AND
                a.ClipNum <> %s AND
                a.ClipNum = c.ClipNum AND
                c.RecordLock <> %s
    """
    DBCursor.execute(query, (kwg, kw, 0, ""))
    for row in fetchall_named(DBCursor):
        msg = _('%s  Clip "%s" is locked by %s\n')
        tempclid = row['ClipID']
        temprl = row['RecordLock']
        if 'unicode' in wx.PlatformInfo:
            msg = unicode(msg, 'utf8')
            tempclid = ProcessDBDataForUTF8Encoding(tempclid)
            temprl = ProcessDBDataForUTF8Encoding(temprl)
        t = msg % (t, tempclid, temprl)

    if t == "":
        # Delphi Transana had a confirmation dialog here, but we won't do
        # that here (do it before calling this function).

        # Build and execute the SQL to delete a word from the list.
        query = """DELETE FROM Keywords2
            WHERE   KeywordGroup = %s AND
                    KeyWord = %s
        """
        DBCursor.execute(query, (kwg, kw))
        # Now delete all instances of this keywordgroup/keyword combo in the
        # Clipkeywords file
        query = """DELETE FROM ClipKeywords2
            WHERE   KeywordGroup = %s AND
                    KeyWord = %s
        """
        DBCursor.execute(query, (kwg, kw))

        # Finish the transaction
        DBCursor.execute("COMMIT")
    else:
        DBCursor.execute("ROLLBACK")
        DBCursor.close()
        msg = _('Unable to delete keyword "%s : %s".\n%s')
        if 'unicode' in wx.PlatformInfo:
            msg = unicode(msg, 'utf8')
        msg = msg % (group, kw_name, t)
        raise TransanaExceptions.GeneralError, msg
    DBCursor.close()

def ClearSourceTranscriptRecords(transcriptNum):
    """ When an Episode Transcript is deleted, it must be removed as a SourceTranscript from Clip Transcript records. """

    # NOTE:  This routine is not perfect.  If a Clip Transcript record is locked by another user, the record WILL be changed
    #        here but that change will be wiped out when the user with the record lock saves (thus restoring the
    #        SourceTranscriptNum value).  However, Transana still knows how to handle it when this value exists but
    #        cannot be found, so I'm not too worried about this rare case.  Blocking the delete seems too extreme here.
    
    # Get a Database cursor
    DBCursor = get_db().cursor()
    # Define a query to delete the appropriate records
    query = "UPDATE Transcripts2 SET SourceTranscriptNum = 0 where SourceTranscriptNum = %s"
    # Execute the query
    DBCursor.execute(query, transcriptNum)
    # Close the Database Cursor
    DBCursor.close()

def delete_filter_records(reportType, reportScope):
    """ Delete Filter Configuration records of a given reportType with a given reportScope """
    # Get a Database cursor
    DBCursor = get_db().cursor()
    # Define a query to delete the appropriate records
    query = "DELETE FROM Filters2 WHERE ReportType = %s AND ReportScope = %s"
    # Execute the query
    DBCursor.execute(query, (reportType, reportScope))
    # Close the Database Cursor
    DBCursor.close()

def record_match_count(table, field_names, field_values):
    """Find number of records in the given table where the given fields
    contain the given values.  If the field name begins with the `!'
    character, then it will match only if the value does NOT equal the given
    field value."""
    DBCursor = get_db().cursor()
    query = "SELECT * FROM %s\n   WHERE" % table
    for field in field_names:
        if field[0] == "!":
            cmp_op = "<>"
            field = field[1:]
        else:
            cmp_op = "="
        query = "%s    %s %s %%s AND\n" % (query, field, cmp_op)
    query = query[:-5]

    DBCursor.execute(query, field_values)
    num = DBCursor.rowcount
    DBCursor.close()
    return num

def ProcessDBDataForUTF8Encoding(text):
    """ MySQL's UTF8 Encoding isn't straight-forward because of variable character length.  For example, the
        Chinese character 4EB0 is stored as \xE4\xBA\xB0 .  Therefore, we need to do some translation
        of the data read from the database to get it into the format that wxPython wants. """
    if not 'unicode' in wx.PlatformInfo:
        return text
    else:
        result = unicode('', TransanaGlobal.encoding)
        # Because some Unicode characters are more than one byte wide, but the STC doesn't recognize this,
        # we will need to skip the processing of the later parts of multi-byte characters.  skipNext allows this.
        skipNext = 0
        # process each character in the StyledText.  The GetStyledText() call has returned a string
        # with the data in two-character chunks, the text char and the styling char.  This for loop
        # allows up to process these character pairs.
        try:
            for x in range(len(text)):
                # If we are looking at the second character of a Unicode character pair, we can skip
                # this processing, as it has already been handled.
                if skipNext > 0:
                    # We need to reset the skipNext flag so we won't skip too many characters.
                    skipNext -= 1
                else:
                    # Check for a Unicode character pair by looking to see if the first character is above 128
                    if ord(text[x]) > 127:
                            
                        # UTF-8 characters are variable length.  We need to figure out the correct number of bytes.
                        # Note the current position
                        pos = x
                        # Initialize the final character variable
                        c = ''

                        # Begin processing of unicode characters, continue until we have a legal character.
                        while (pos < len(text)):
                            # Add the current character to the character variable
                            c += chr(ord(text[pos]))  # "Un-Unicode" the character ????
                            # Try to encode the character.
                            try:
                                # See if we have a legal UTF-8 character yet.
                                d = unicode(c, TransanaGlobal.encoding)
                                # If so, break out of the while loop
                                break
                            # If we don't have a legal UTF-8 character, we'll get a UnicodeDecodeError exception
                            except UnicodeDecodeError:
                                # We need to signal the need to skip a charater in overall processing
                                skipNext += 1
                                # We need to update the current position and keep processing until we have a legal UTF-8 character
                                pos += 1

                        result += unicode(c, TransanaGlobal.encoding)
                    else:
                        c = text[x]
                        result += c

        except TypeError:
            result = text
        except UnicodeDecodeError:
            # If we are reading Unicode text from Transana 2.05 or earlier, the line above that reads:
            # result += unicode(c, TransanaGlobal.encoding)
            # throws a UnicodeDecodeError when it can't interpret Latin-1 encoded characters using UTF-8.
            # When that happens, we need to use Latin-1 encoding instead of UTF-8.

            # The text doesn't need to be encoded in this circumstance.
            result = text
            # If we're in Russian, change the encoding to KOI8r
            if TransanaGlobal.configData.language == 'ru':
                TransanaGlobal.encoding = 'koi8_r'
            # If we're in Chinese, change the encoding to the appropriate Chinese encoding
            elif TransanaGlobal.configData.language == 'zh':
                TransanaGlobal.encoding = TransanaConstants.chineseEncoding
            # If we're in Eastern European Encoding, change the encoding to 'iso8859_2'
            elif TransanaGlobal.configData.language == 'easteurope':
                TransanaGlobal.encoding = 'iso8859_2'
            # If we're in Greek, change the encoding to 'iso8859_7'
            elif TransanaGlobal.configData.language == 'el':
                TransanaGlobal.encoding = 'iso8859_7'
            # If we're in Japanese, change the encoding to cp932
            elif TransanaGlobal.configData.language == 'ja':
                TransanaGlobal.encoding = 'cp932'
            # If we're in Korean, change the encoding to cp949
            elif TransanaGlobal.configData.language == 'ko':
                TransanaGlobal.encoding = 'cp949'
            # Otherwise, fall back to Latin-1
            else:
                TransanaGlobal.encoding = 'latin1'

        return result


def UpdateDBFilenames(parent, filePath, fileList):
    # To start with, let's make sure the filePath ends with the appropriate Seperator
    if filePath[-1] != os.sep:
        filePath = filePath + os.sep
    # If the filePath is on the Default Video Path, extract the Default Video Path from the filePath
    if filePath.find(TransanaGlobal.configData.videoPath) == 0:
        filePath = filePath[len(TransanaGlobal.configData.videoPath):]

    # Initialize a Success flag, assuming we will succeed.
    success = True

    parent.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))

    # Get a Database Cursor
    DBCursor = get_db().cursor()
    # Begin a Database Transaction
    DBCursor.execute("BEGIN")

    # Define Queries for Episode and Clip Records.  Because some records might be locked by other users,
    # we can't simply use this:
    #
    #   query = "UPDATE Episodes2 SET MediaFile = %s WHERE MediaFile LIKE %s"
    #
    # Instead, we have to load each record, lock it, change it, save it, and unlock it.
    episodeQuery = "SELECT EpisodeNum FROM Episodes2 WHERE MediaFile LIKE %s"
    clipQuery = "SELECT ClipNum FROM Clips2 WHERE MediaFile LIKE %s"
    additionalQuery = "SELECT AddVidNum, EpisodeNum, ClipNum FROM AdditionalVids2 WHERE MediaFile LIKE %s"

    # Let's count the number of records changed
    episodeCounter = 0
    clipCounter = 0

    # Go through the fileList and run the query repeatedly
    for fileName in fileList:
        if 'unicode' in wx.PlatformInfo:
            # NOTE:  Hmmmm.  How do we know what encoding the filenames are in?  It could be UTF-8 or it could be Latin1.
            #        The Global encoding may not be correct.
            queryFileName = fileName.encode(TransanaGlobal.encoding)
        else:
            queryFileName = fileName
        
        # Add a "%" character to the beginning of the File Name so that the "LIKE" operator will work
        # Execute the Episode Query
        DBCursor.execute(episodeQuery, '%' + queryFileName)
        # Iterate through the records returned from the Database
        for (episodeNum, ) in DBCursor.fetchall():
            # Load the Episode
            tempEpisode = Episode.Episode(episodeNum)
            # Be ready to catch exceptions
            try:
                # Lock the Record
                tempEpisode.lock_record()
                # Make sure the file names match, that we don't have a subset name.
                # ('mens group.mov' was substituted for 'womens group.mov', for instance.)
                if os.path.split(tempEpisode.media_filename)[1].upper() == fileName.upper():
                    # Update the Media Filename
                    tempEpisode.media_filename = filePath + fileName
                # Save the Record
                tempEpisode.db_save()
                # Unlock the Record
                tempEpisode.unlock_record()
                # Increment the Counter
                episodeCounter += 1
            # If an exception is raised, catch it
            except:
                if DEBUG:
                    (exctype, excvalue) = sys.exc_info()[:2]
                    print "DBInterface.UpdateDBFilenames() Exception: \n%s\n%s" % (exctype, excvalue)
                # Indicate that we have failed.
                success = False
                # Don't bother to contine processing DB Records
                break
        # If we failed during Episodes, there's no reason to look at Clips
        if not success:
            break
        
        # Execute the Clip Query
        DBCursor.execute(clipQuery, '%' + queryFileName)

        # Iterate through the records returned from the Database
        for (clipNum, ) in DBCursor.fetchall():
            # Load the Clip
            tempClip = Clip.Clip(clipNum)
            # Be ready to catch exceptions
            try:
                # Lock the Record
                tempClip.lock_record()
                # Make sure the file names match, that we don't have a subset name.
                # ('mens group.mov' was substituted for 'womens group.mov', for instance.)
                if os.path.split(tempClip.media_filename)[1].upper() == fileName.upper():
                    # Update the Media Filename
                    tempClip.media_filename = filePath + fileName
                # Save the Record
                tempClip.db_save()
                # Unlock the Record
                tempClip.unlock_record()
                # Increment the Counter
                clipCounter += 1
            # If an exception is raised, catch it
            except:
                if DEBUG:
                    (exctype, excvalue) = sys.exc_info()[:2]
                    print "DBInterface.UpdateDBFilenames() Exception: \n%s\n%s" % (exctype, excvalue)
                # Indicate that we have failed.
                success = False
                # Don't bother to contine processing DB Records
                break

        # If we failed during Clips, there's no reason to look at Additional Media Files
        if not success:
            break

        # Execute the Additional Query
        DBCursor.execute(additionalQuery, '%' + queryFileName)

        # Iterate through the records returned from the Database
        for (addVidNum, episodeNum, clipNum) in DBCursor.fetchall():
            # If we're dealing with an Episode ...
            if episodeNum > 0:
                # Load the Episode
                tempEpisode = Episode.Episode(episodeNum)
                # Be ready to catch exceptions
                try:
                    # Lock the Record
                    tempEpisode.lock_record()
                    # Remember the Additional Media Files
                    additionalMediaFiles = tempEpisode.additional_media_files
                    # Clear the Additional Media Files from the Episode
                    del(tempEpisode.additional_media_files)
                    # Update the Additional Media Files.  Iterate through the list of files
                    for vid in additionalMediaFiles:
                        # Extract the video filename
                        fn = vid['filename']
                        # If the file matches the name being updated ...
                        if os.path.split(fn)[1].upper() == fileName.upper():
                            # ... then update the additional filename entry
                            tempEpisode.additional_media_files = {'filename' : filePath + fileName,
                                                                  'length'   : vid['length'],
                                                                  'offset'   : vid['offset'],
                                                                  'audio'    : vid['audio']}
                        # If there's NOT a match ...
                        else:
                            # ... then continue to use the existing file name
                            tempEpisode.additional_media_files = vid
                    # Save the Record
                    tempEpisode.db_save()
                    # Unlock the Record
                    tempEpisode.unlock_record()
                    # Increment the Counter
                    episodeCounter += 1
                # If an exception is raised, catch it
                except:
                    if DEBUG:
                        (exctype, excvalue) = sys.exc_info()[:2]
                        print "DBInterface.UpdateDBFilenames() Exception: \n%s\n%s" % (exctype, excvalue)
                    # Indicate that we have failed.
                    success = False
                    # Don't bother to contine processing DB Records
                    break

            elif clipNum > 0:
                # Load the Clip
                tempClip = Clip.Clip(clipNum)
                # Be ready to catch exceptions
                try:
                    # Lock the Record
                    tempClip.lock_record()
                    # Remember the Additional Media Files
                    additionalMediaFiles = tempClip.additional_media_files
                    # Clear the Additional Media Files from the Clip
                    del(tempClip.additional_media_files)
                    # Update the Additional Media Files.  Iterate through the list of files
                    for fn in additionalMediaFiles:
                        # Extract the video filename
                        fn = vid['filename']
                        # If the file matches the name being updated ...
                        if os.path.split(fn)[1].upper() == fileName.upper():
                            # ... then update the additional filename entry
                            tempClip.additional_media_files = {'filename' : filePath + fileName,
                                                               'length'   : vid['length'],
                                                               'offset'   : vid['offset'],
                                                               'audio'    : vid['audio']}
                        # If there's NOT a match ...
                        else:
                            # ... then continue to use the existing file name
                            tempClip.additional_media_files = vid
                    # Save the Record
                    tempClip.db_save()
                    # Unlock the Record
                    tempClip.unlock_record()
                    # Increment the Counter
                    clipCounter += 1
                # If an exception is raised, catch it
                except:
                    if DEBUG:
                        (exctype, excvalue) = sys.exc_info()[:2]
                        print "DBInterface.UpdateDBFilenames() Exception: \n%s\n%s" % (exctype, excvalue)
                    # Indicate that we have failed.
                    success = False
                    # Don't bother to contine processing DB Records
                    break

    # If there have been no problems, Commit the Transaction to the Database
    if success:
        DBCursor.execute("COMMIT")
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("%s Episode Records have been updated.\n%s Clip Records have been updated."), 'utf8')
        else:
            prompt = _("%s Episode Records have been updated.\n%s Clip Records have been updated.")
        infodlg = Dialogs.InfoDialog(None, prompt % (episodeCounter, clipCounter))
        infodlg.ShowModal()
        infodlg.Destroy()
    # Otherwise, Roll the Transaction Back.
    else:
        DBCursor.execute("ROLLBACK")
    # Close the Database Cursor
    DBCursor.close()

    parent.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
    
    # Let the calling routine know if we were successful or not
    return success

def DeleteDatabase(username, password, server, database, port):
    """ Delete an entire database """
    # Assume failure
    res = False
    try:
        # Connect to the MySQL Server without specifying a Database
        if TransanaConstants.singleUserVersion:
            # The single-user version requires no parameters
            dbConn = MySQLdb.connect()
        else:
            try:
                # The multi-user version requires all information except the database name
                dbConn = MySQLdb.connect(host=server, user=username, passwd=password, port=port)
                # If the connection fails here, it's not the database name that was the problem.
            except:
                dbConn = None
        # If we made a connection to MySQL...
        if dbConn != None:
            try:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Are you sure you want to delete Database "%s"?\nAll data in this database will be permanently deleted and cannot be recovered.'), 'utf8')
                else:
                    prompt = _('Are you sure you want to delete Database "%s"?\nAll data in this database will be permanently deleted and cannot be recovered.')
                dlg = Dialogs.QuestionDialog(None, prompt % database, _('Delete Database'), noDefault=True)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    # Give the user another chance to back out!
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Are you ABSOLUTELY SURE you want to delete Database "%s"?\nAll data in this database will be permanently deleted and cannot be recovered.'), 'utf8')
                    else:
                        prompt = _('Are you ABSOLUTELY SURE you want to delete Database "%s"?\nAll data in this database will be permanently deleted and cannot be recovered.')
                    dlg = Dialogs.QuestionDialog(None, prompt % database, _('Delete Database'), noDefault=True)
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                    if result == wx.ID_YES:
                        # Get a Database Cursor
                        dbCursor = dbConn.cursor()
                        # Query the Database to delete the desired database
                        dbCursor.execute('DROP DATABASE IF EXISTS %s' % database.encode(TransanaGlobal.encoding))
                        # Close the Database Cursor
                        dbCursor.close()
                        # If we get this far, return True rather than False
                        res = True

                    # Okay, this is weird.  If you drop a database, immediately quit transana, then
                    # shut down the database server (in multi-user mode), then the server 
                    # won't be able to start up again.  
                    # (To get the server going again, you need to restore the missing database's folder!)
            finally:
                # Close the Database Connection
                dbConn.close()
        return res
    except:
        return res

def ReportRecordLocks(parent):
    """ Query the database for Record Locks and build a string that holds the report data. """
    # Initialize the Report Results string
    if 'unicode' in wx.PlatformInfo:
        resMessage = u''
    else:
        resMessage = ''
    # We also need to return a list of the users with Record Locks.  Initialize that list.
    userList = []
    # Change to the Wait Cursor
    parent.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))

    # Get a Database Connection
    dbConn = get_db()
    # Get a Database Cursor
    DBCursor = dbConn.cursor()

    # NOTE:  Originally, I had a clever structure that looped through a list of files and pulled the correct records.
    #        I abandoned this approach because I wanted to be able to provide more information about each locked
    #        record, requiring more complex and customized queries and report strings.   It ain't pretty, but it works.

    # Define the SERIES Query
    lockQuery = """SELECT SeriesID, RecordLock
                   FROM Series2
                   WHERE RecordLock <> ''"""
    # run the query
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        resMessage += unicode(_('Series records:\n'), 'utf8')
    else:
        resMessage += _('Series records:\n')
    # Execute the Series Query
    DBCursor.execute(lockQuery)
    # Iterate through the records returned from the Database
    for recs in fetchall_named(DBCursor):
        # Get the DB Values
        tempSeriesID = recs['SeriesID']
        tempRecordLock = recs['RecordLock']
        # If we're in Unicode mode, format the strings appropriately
        if 'unicode' in wx.PlatformInfo:
            tempSeriesID = ProcessDBDataForUTF8Encoding(tempSeriesID)
            tempRecordLock = ProcessDBDataForUTF8Encoding(tempRecordLock)
        # Add the data to the Report Results string
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Series "%s" is locked by %s\n'), 'utf8')
        else:
            prompt = _('Series "%s" is locked by %s\n')
        resMessage += prompt % (tempSeriesID, tempRecordLock)
        # If the user holding the lock isn't already in the list, add him/her.
        if not (tempRecordLock in userList):
            userList.append(tempRecordLock)
    resMessage += '\n'

    # Define the EPISODE Query
    lockQuery = """SELECT EpisodeID, SeriesID, e.RecordLock
                   FROM Episodes2 e, Series2 s
                   WHERE e.RecordLock <> '' AND
                         e.SeriesNum = s.SeriesNum""" 
    # run the query
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        resMessage += unicode(_('Episode records:\n'), 'utf8')
    else:
        resMessage += _('Episode records:\n')
    # Execute the Episode Query
    DBCursor.execute(lockQuery)
    # Iterate through the records returned from the Database
    for recs in fetchall_named(DBCursor):
        # Get the DB Values
        tempEpisodeID = recs['EpisodeID']
        tempSeriesID = recs['SeriesID']
        tempRecordLock = recs['RecordLock']
        # If we're in Unicode mode, format the strings appropriately
        if 'unicode' in wx.PlatformInfo:
            tempEpisodeID = ProcessDBDataForUTF8Encoding(tempEpisodeID)
            tempSeriesID = ProcessDBDataForUTF8Encoding(tempSeriesID)
            tempRecordLock = ProcessDBDataForUTF8Encoding(tempRecordLock)
        # Add the data to the Report Results string
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Episode "%s" in Series "%s" is locked by %s\n'), 'utf8')
        else:
            prompt = _('Episode "%s" in Series "%s" is locked by %s\n')
        resMessage += prompt % (tempEpisodeID, tempSeriesID, tempRecordLock)
        # If the user holding the lock isn't already in the list, add him/her.
        if not (tempRecordLock in userList):
            userList.append(tempRecordLock)
    # Add a blank line to clearly delineate the report sections
    resMessage += '\n'

    # Define the EPISODE TRANSCRIPT Query
    lockQuery = """SELECT TranscriptID, EpisodeID, SeriesID, t.RecordLock
                   FROM Transcripts2 t, Episodes2 e, Series2 s
                   WHERE t.RecordLock <> '' AND
                         t.EpisodeNum > 0 AND
                         t.ClipNum = 0 AND
                         t.EpisodeNum = e.EpisodeNum AND
                         e.SeriesNum = s.SeriesNum""" 
    # run the query
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        resMessage += unicode(_('Episode Transcript records:\n'), 'utf8')
    else:
        resMessage += _('Episode Transcript records:\n')
    # Execute the Episode Query
    DBCursor.execute(lockQuery)
    # Iterate through the records returned from the Database
    for recs in fetchall_named(DBCursor):
        # Get the DB Values
        tempTranscriptID = recs['TranscriptID']
        tempEpisodeID = recs['EpisodeID']
        tempSeriesID = recs['SeriesID']
        tempRecordLock = recs['RecordLock']
        # If we're in Unicode mode, format the strings appropriately
        if 'unicode' in wx.PlatformInfo:
            tempTranscriptID = ProcessDBDataForUTF8Encoding(tempTranscriptID)
            tempEpisodeID = ProcessDBDataForUTF8Encoding(tempEpisodeID)
            tempSeriesID = ProcessDBDataForUTF8Encoding(tempSeriesID)
            tempRecordLock = ProcessDBDataForUTF8Encoding(tempRecordLock)
        # Add the data to the Report Results string
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Transcript "%s" in Episode "%s" in Series "%s" is locked by %s\n'), 'utf8')
        else:
            prompt = _('Transcript "%s" in Episode "%s" in Series "%s" is locked by %s\n')
        resMessage += prompt % (tempTranscriptID, tempEpisodeID, tempSeriesID, tempRecordLock)
        # If the user holding the lock isn't already in the list, add him/her.
        if not (tempRecordLock in userList):
            userList.append(tempRecordLock)
    # Add a blank line to clearly delineate the report sections
    resMessage += '\n'

    # Define the COLLECTION Query
    lockQuery = """SELECT CollectID, ParentCollectNum, RecordLock
                   FROM Collections2
                   WHERE RecordLock <> ''"""
    # run the query
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        resMessage += unicode(_('Collection records:\n'), 'utf8')
    else:
        resMessage += _('Collection records:\n')
    # Execute the Collection Query
    DBCursor.execute(lockQuery)
    # Iterate through the records returned from the Database
    for recs in fetchall_named(DBCursor):
        # Get the DB Values
        tempCollectID = recs['CollectID']
        tempRecordLock = recs['RecordLock']
        # If we're in Unicode mode, format the strings appropriately
        if 'unicode' in wx.PlatformInfo:
            tempCollectID = ProcessDBDataForUTF8Encoding(tempCollectID)
            tempRecordLock = ProcessDBDataForUTF8Encoding(tempRecordLock)
        # Add the data to the Report Results string
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Collection "%s" '), 'utf8')
        else:
            prompt = _('Collection "%s" ')
        resMessage += prompt % tempCollectID
        # Note the collection Parent, so we can list Collection Nesting
        collPar = recs['ParentCollectNum']
        # While we're looking at a nested Collection ...
        while collPar > 0L:
            # ... build a query to get the parent collection ...
            subQ = """ SELECT CollectID, ParentCollectNum
                       FROM Collections2
                       WHERE CollectNum = %d """
            # ... get a second database cursor ...
            DBCursor2 = dbConn.cursor()
            # ... execute the parent collection query ...
            DBCursor2.execute(subQ % collPar)
            # ... get the parent collection data ...
            rec2 = DBCursor2.fetchone()
            # ... note the collection's parent ...
            collPar = rec2[1]
            # Get the DB Value
            tempCollectID = rec2[0]
            # If we're in Unicode mode, format the strings appropriately
            if 'unicode' in wx.PlatformInfo:
                tempCollectID = ProcessDBDataForUTF8Encoding(tempCollectID)
            # ... add the parent collection to the report Results String ...
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('nested in "%s" '), 'utf8')
            else:
                prompt = _('nested in "%s" ')
            resMessage += prompt % tempCollectID
            # ... close the second database cursor ...
            DBCursor2.close()
        # ... and complete the Report Results string.
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('is locked by %s\n'), 'utf8')
        else:
            prompt = _('is locked by %s\n')
        resMessage += prompt % tempRecordLock
        # If the user holding the lock isn't already in the list, add him/her.
        if not (tempRecordLock in userList):
            userList.append(tempRecordLock)
    # Add a blank line to clearly delineate the report sections
    resMessage += '\n'

    # Define the Clip Query
    lockQuery = """SELECT ClipID, CollectID, c.ParentCollectNum, cl.RecordLock
                   FROM Clips2 cl, Collections2 c
                   WHERE cl.RecordLock <> '' AND
                         cl.CollectNum = c.CollectNum""" 
    # run the query
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        resMessage += unicode(_('Clip records:\n'), 'utf8')
    else:
        resMessage += _('Clip records:\n')
    # Execute the Clip Query
    DBCursor.execute(lockQuery)
    # Iterate through the records returned from the Database
    for recs in fetchall_named(DBCursor):
        # Get the DB Values
        tempClipID = recs['ClipID']
        tempCollectID = recs['CollectID']
        tempRecordLock = recs['RecordLock']
        # If we're in Unicode mode, format the strings appropriately
        if 'unicode' in wx.PlatformInfo:
            tempClipID = ProcessDBDataForUTF8Encoding(tempClipID)
            tempCollectID = ProcessDBDataForUTF8Encoding(tempCollectID)
            tempRecordLock = ProcessDBDataForUTF8Encoding(tempRecordLock)
        # Add the data to the Report Results string
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Clip "%s" in Collection "%s" '), 'utf8')
        else:
            prompt = _('Clip "%s" in Collection "%s" ')
        resMessage += prompt % (tempClipID, tempCollectID)
        # Note the collection Parent, so we can list Collection Nesting
        collPar = recs['ParentCollectNum']
        # While we're looking at a nested Collection ...
        while collPar > 0L:
            # ... build a query to get the parent collection ...
            subQ = """ SELECT CollectID, ParentCollectNum
                       FROM Collections2
                       WHERE CollectNum = %d """
            # ... get a second database cursor ...
            DBCursor2 = dbConn.cursor()
            # ... execute the parent collection query ...
            DBCursor2.execute(subQ % collPar)
            # ... get the parent collection data ...
            rec2 = DBCursor2.fetchone()
            # ... note the collection's parent ...
            collPar = rec2[1]
            # Get the DB Value
            tempCollectID = rec2[0]
            # If we're in Unicode mode, format the strings appropriately
            if 'unicode' in wx.PlatformInfo:
                tempCollectID = ProcessDBDataForUTF8Encoding(tempCollectID)
            # ... add the parent collection to the report Results String ...
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('nested in "%s" '), 'utf8')
            else:
                prompt = _('nested in "%s" ')
            resMessage += prompt % tempCollectID
            # ... close the second database cursor ...
            DBCursor2.close()
        # ... and complete the Report Results string.
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('is locked by %s\n'), 'utf8')
        else:
            prompt = _('is locked by %s\n')
        resMessage += prompt % tempRecordLock
        # If the user holding the lock isn't already in the list, add him/her.
        if not (tempRecordLock in userList):
            userList.append(tempRecordLock)
    # Add a blank line to clearly delineate the report sections
    resMessage += '\n'

    # Define the CLIP TRANSCRIPT Query
    lockQuery = """SELECT TranscriptID, ClipID, CollectID, ParentCollectNum, t.RecordLock
                   FROM Transcripts2 t, Clips2 cl, Collections2 c
                   WHERE t.RecordLock <> '' AND
                         t.ClipNum > 0 AND
                         t.ClipNum = cl.ClipNum AND
                         cl.CollectNum = c.CollectNum""" 
    # run the query
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        resMessage += unicode(_('Clip Transcript records:\n'), 'utf8')
    else:
        resMessage += _('Clip Transcript records:\n')
    # Execute the Episode Query
    DBCursor.execute(lockQuery)
    # Iterate through the records returned from the Database
    for recs in fetchall_named(DBCursor):
        # Get the DB Values
        tempClipID = recs['ClipID']
        tempCollectID = recs['CollectID']
        tempRecordLock = recs['RecordLock']
        # If we're in Unicode mode, format the strings appropriately
        if 'unicode' in wx.PlatformInfo:
            tempClipID = ProcessDBDataForUTF8Encoding(tempClipID)
            tempCollectID = ProcessDBDataForUTF8Encoding(tempCollectID)
            tempRecordLock = ProcessDBDataForUTF8Encoding(tempRecordLock)
        # Add the data to the Report Results string
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('The Clip Transcript for Clip "%s" in Collection "%s" '), 'utf8')
        else:
            prompt = _('The Clip Transcript for Clip "%s" in Collection "%s" ')
        resMessage += prompt % (tempClipID, tempCollectID)
        # Note the collection Parent, so we can list Collection Nesting
        collPar = recs['ParentCollectNum']
        # While we're looking at a nested Collection ...
        while collPar > 0L:
            # ... build a query to get the parent collection ...
            subQ = """ SELECT CollectID, ParentCollectNum
                       FROM Collections2
                       WHERE CollectNum = %d """
            # ... get a second database cursor ...
            DBCursor2 = dbConn.cursor()
            # ... execute the parent collection query ...
            DBCursor2.execute(subQ % collPar)
            # ... get the parent collection data ...
            rec2 = DBCursor2.fetchone()
            # ... note the collection's parent ...
            collPar = rec2[1]
            # Get the DB Value
            tempCollectID = rec2[0]
            # If we're in Unicode mode, format the strings appropriately
            if 'unicode' in wx.PlatformInfo:
                tempCollectID = ProcessDBDataForUTF8Encoding(tempCollectID)
            # ... add the parent collection to the report Results String ...
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('nested in "%s" '), 'utf8')
            else:
                prompt = _('nested in "%s" ')
            resMessage += prompt % tempCollectID
            # ... close the second database cursor ...
            DBCursor2.close()
        # ... and complete the Report Results string.
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('is locked by %s\n'), 'utf8')
        else:
            prompt = _('is locked by %s\n')
        resMessage += prompt % tempRecordLock
        # If the user holding the lock isn't already in the list, add him/her.
        if not (tempRecordLock in userList):
            userList.append(tempRecordLock)
    # Add a blank line to clearly delineate the report sections
    resMessage += '\n'

    # Define the NOTES Query
    lockQuery = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, RecordLock
                   FROM Notes2
                   WHERE RecordLock <> ''"""
    # run the query
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        resMessage += unicode(_('Notes records:\n'), 'utf8')
    else:
        resMessage += _('Notes records:\n')
    # Execute the Notes Query
    DBCursor.execute(lockQuery)
    # Iterate through the records returned from the Database
    for recs in fetchall_named(DBCursor):
        # Determine what type of Note we're looking at, and add that to the Report Results string
        if recs['SeriesNum'] > 0:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                resMessage += unicode(_('Series') + ' ', 'utf8')
            else:
                resMessage += _('Series') + ' '
        elif recs['EpisodeNum'] > 0:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                resMessage += unicode(_('Episode') + ' ', 'utf8')
            else:
                resMessage += _('Episode') + ' '
        elif recs['TranscriptNum'] > 0:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                resMessage += unicode(_('Transcript') + ' ', 'utf8')
            else:
                resMessage += _('Transcript') + ' '
        elif recs['CollectNum'] > 0:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                resMessage += unicode(_('Collection') + ' ', 'utf8')
            else:
                resMessage += _('Collection') + ' '
        elif recs['ClipNum'] > 0:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                resMessage += unicode(_('Clip') + ' ', 'utf8')
            else:
                resMessage += _('Clip') + ' '
        # Get the DB Values
        tempNoteID = recs['NoteID']
        tempRecordLock = recs['RecordLock']
        # If we're in Unicode mode, format the strings appropriately
        if 'unicode' in wx.PlatformInfo:
            tempNoteID = ProcessDBDataForUTF8Encoding(tempNoteID)
            tempRecordLock = ProcessDBDataForUTF8Encoding(tempRecordLock)
        # Add the data to the Report Results string
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Note "%s" is locked by %s\n'), 'utf8')
        else:
            prompt = _('Note "%s" is locked by %s\n')
        resMessage += prompt % (tempNoteID, tempRecordLock)
        # If the user holding the lock isn't already in the list, add him/her.
        if not (tempRecordLock in userList):
            userList.append(tempRecordLock)
    # Add a blank line to clearly delineate the report sections
    resMessage += '\n'

    # Define the KEYWORDS Query
    lockQuery = """SELECT KeywordGroup, Keyword, RecordLock
                   FROM Keywords2
                   WHERE RecordLock <> ''"""
    # run the query
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        resMessage += unicode(_('Keyword records:\n'), 'utf8')
    else:
        resMessage += _('Keyword records:\n')
    # Execute the Keyword Query
    DBCursor.execute(lockQuery)
    # Iterate through the records returned from the Database
    for recs in fetchall_named(DBCursor):
        # Get the DB Values
        tempKWG = recs['KeywordGroup']
        tempKW = recs['Keyword']
        tempRecordLock = recs['RecordLock']
        # If we're in Unicode mode, format the strings appropriately
        if 'unicode' in wx.PlatformInfo:
            tempKWG = ProcessDBDataForUTF8Encoding(tempKWG)
            tempKW = ProcessDBDataForUTF8Encoding(tempKW)
            tempRecordLock = ProcessDBDataForUTF8Encoding(tempRecordLock)
        # Add the data to the Report Results string
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Keyword "%s : %s" is locked by %s\n'), 'utf8')
        else:
            prompt = _('Keyword "%s : %s" is locked by %s\n')
        resMessage += prompt % (tempKWG, tempKW, tempRecordLock)
        # If the user holding the lock isn't already in the list, add him/her.
        if not (tempRecordLock in userList):
            userList.append(tempRecordLock)
    # Add a blank line to clearly delineate the report sections
    resMessage += '\n'

    # Define the CORE DATA Query
    lockQuery = """SELECT Identifier, RecordLock
                   FROM CoreData2
                   WHERE RecordLock <> ''"""
    # run the query
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        resMessage += unicode(_('Core Data records:\n'), 'utf8')
    else:
        resMessage += _('Core Data records:\n')
    # Execute the Core Data Query
    DBCursor.execute(lockQuery)
    # Iterate through the records returned from the Database
    for recs in fetchall_named(DBCursor):
        # Get the DB Values
        tempID = recs['Identifier']
        tempRecordLock = recs['RecordLock']
        # If we're in Unicode mode, format the strings appropriately
        if 'unicode' in wx.PlatformInfo:
            tempID = ProcessDBDataForUTF8Encoding(tempID)
            tempRecordLock = ProcessDBDataForUTF8Encoding(tempRecordLock)
        # Add the data to the Report Results string
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Core Data for "%s" is locked by %s\n'), 'utf8')
        else:
            prompt = _('Core Data for "%s" is locked by %s\n')
        resMessage += prompt % (tempID, tempRecordLock)
        # If the user holding the lock isn't already in the list, add him/her.
        if not (tempRecordLock in userList):
            userList.append(tempRecordLock)
    # Add a blank line to clearly delineate the report sections
    resMessage += '\n'

    # Close the Database Cursor
    DBCursor.close()

    # Reset the Cursor now that the Report data is generated
    parent.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
    
    # return the message and the list of users with Locks
    return (resMessage, userList)

def UnlockRecords(parent, userName):
    """ Unlock all locked records by the named User """
    # Change the Cursor to the Wait Cursor
    parent.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
    # Create a list of tables to iterate through unlocking records.
    tableList = ['Series2', 'Episodes2', 'Transcripts2', 'Collections2', 'Clips2', 'Notes2', 'Keywords2', 'CoreData2']
    # Get a Database Connection
    dbConn = get_db()
    # Get a Database Cursor
    DBCursor = dbConn.cursor()

    # Create the Unlock query
    unlockQuery = """ UPDATE %s
                        SET RecordLock = '',
                            LockTime = NULL
                        WHERE RecordLock = '%s' """

    # Iterate through the list of tables ...
    for table in tableList:
        # ... and execute the unlock query for each table.
        DBCursor.execute(unlockQuery % (table, userName))

    # Close the Database Cursor
    DBCursor.close()
    # Reset the Cursor now that the unlock is complete.
    parent.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

def ServerDateTime():
    """ Returns the SERVER's current Date and Time for record lock comparisons """
    # NOTE:  When comparing lock time to current time to see if the lock has expired, you can't just look
    #        at the time on the local computer, which might not be in the same time zone as the locking
    #        computer.  Further, what if someone's computer has the wrong date or time setting?
    #        Therefore, all comparisons should be made with the DB Server's current date and time.
    #        This function returns that value.
    
    # Get a Database Connection
    dbConn = get_db()
    # Get a Database Cursor
    DBCursor = dbConn.cursor()
    # Get the DB Server's current Date and Time
    DBCursor.execute('SELECT CURRENT_TIMESTAMP()')
    # Get the result
    serverDateTime = DBCursor.fetchall()[0][0]
    # Close the Database Cursor
    DBCursor.close()
    # Return the value retrieved from the server
    return serverDateTime
