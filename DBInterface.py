# Copyright (C) 2003 - 2005 The Board of Regents of the University of Wisconsin System 
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


__author__ = 'Nathaniel Case, David K. Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare'


# import wxPython
import wx
# import MySQL for Python
import MySQLdb
# import Python's exceptions module
from exceptions import *
# import Python's os module
import os
# import Python's string module
import string
# Import Transana's Dialog Boxes
import Dialogs
# import Transana's Constants
import TransanaConstants
# import Transana's Global Variables
import TransanaGlobal

_dbref = None
_user = ""


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
    # TODO:  Internationalize the "--language" parameter based on selected language.
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
    # Polish
    elif (TransanaGlobal.configData.language == 'pl'):
        lang = '--language=./share/polish'
    # Russian
    elif (TransanaGlobal.configData.language == 'ru'):
        lang = '--language=./share/russian'
    # Swedish
    elif (TransanaGlobal.configData.language == 'sv'):
        lang = '--language=./share/swedish'
    MySQLdb.server_init(args=['Transana', datadir, '--basedir=.', lang])
    
def EndSingleUserDatabase():
    """ For single-user Transana only, this ends (exits) the embedded MySQL Server. """
    # End the embedded MySQL Server
    MySQLdb.server_end()

def SetTableType(hasBDB, query):
    if hasBDB:
        query = query + 'TYPE=BDB'
    else:
        query = query + 'TYPE=InnoDB'
    return query

def is_db_open():
    """ Quick and dirty test to see if the database is currently open """
    return _dbref != None

def establish_db_exists():
    """Check for the existence of all database tables and create them
    if necessary."""
    # Obtain a Database
    db = get_db()
    # If this fails, return "False" to indicate failure
    if db == None:
        
        return False

    # Obtain a Database Cursor
    dbCursor = db.cursor()
    # Initialize BDB and InnoDB Table Flags to false
    hasBDB = False
    hasInnoDB = False
    # First, let's find out if the BDB Tables are supported on the MySQL Instance we are using.
    # Define a "SHOW VARIABLES" Query
    query = "SHOW VARIABLES"
    # Execute the Query
    dbCursor.execute(query)
    # Look at the Results Set
    for pair in dbCursor.fetchall():
        # If there is a pair in the Results that indicates that the value for "have_bdb" is "YES",
        # set the DBD Flag to True
        if pair[0] == 'have_bdb':
            if pair[1] == 'YES':
                hasBDB = True
        # If there is a pair in the Results that indicates that the value for "have_innodb" is "YES",
        # set the InnoDB Flag to True
        if pair[0] == 'have_innodb':
            if pair[1] == 'YES':
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
                     DefaultKeywordGroup  VARCHAR(100), 
                     RecordLock           VARCHAR(25),
                     LockTime             DATETIME, 
                     PRIMARY KEY (SeriesNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasBDB, query)
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
        query = SetTableType(hasBDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Transcript Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Transcripts2
                    (TranscriptNum   INTEGER auto_increment, 
                     TranscriptID    VARCHAR(100), 
                     EpisodeNum      INTEGER, 
                     ClipNum         INTEGER, 
                     Transcriber     VARCHAR(100), 
                     Comment         VARCHAR(255), 
                     RTFText         LONGBLOB, 
                     RecordLock      VARCHAR(25), 
                     LockTime        DATETIME, 
                     LastSaveTime    DATETIME, 
                     PRIMARY KEY (TranscriptNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasBDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Collection Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Collections2
                    (CollectNum            INTEGER auto_increment, 
                     CollectID             VARCHAR(100), 
                     ParentCollectNum      INTEGER, 
                     CollectComment        VARCHAR(255), 
                     CollectOwner          VARCHAR(100), 
                     DefaultKeywordGroup   VARCHAR(100), 
                     RecordLock          VARCHAR(25), 
                     LockTime            DATETIME, 
                     PRIMARY KEY (CollectNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasBDB, query)
        # Execute the Query
        dbCursor.execute(query)

        #  Clip Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Clips2
                    (ClipNum        INTEGER auto_increment, 
                     ClipID         VARCHAR(100), 
                     CollectNum     INTEGER, 
                     EpisodeNum     INTEGER, 
                     TranscriptNum  INTEGER, 
                     MediaFile      VARCHAR(255), 
                     ClipStart      INTEGER, 
                     ClipStop       INTEGER, 
                     ClipComment    VARCHAR(255), 
                     SortOrder      INTEGER, 
                     RecordLock     VARCHAR(25), 
                     LockTime       DATETIME, 
                     PRIMARY KEY (ClipNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasBDB, query)
        # Execute the Query
        dbCursor.execute(query)

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
                     NoteText       BLOB, 
                     RecordLock     VARCHAR(25), 
                     LockTime       DATETIME, 
                     PRIMARY KEY (NoteNum))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasBDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # Keywords Table: Test for existence and create if needed
        query = """
                  CREATE TABLE IF NOT EXISTS Keywords2
                    (KeywordGroup  VARCHAR(100) NOT NULL, 
                     Keyword       VARCHAR(255) NOT NULL, 
                     Definition    VARCHAR(255), 
                     RecordLock    VARCHAR(25), 
                     LockTime      DATETIME, 
                     PRIMARY KEY (KeywordGroup, Keyword))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasBDB, query)
        # Execute the Query
        dbCursor.execute(query)

        # ClipKeywords Table: Test for existence and create if needed
        # MySQL Primary Keys cannot contain NULL values, and either EpisodeNum
        # or ClipNum will always be NULL!  Therefore, use a UNIQUE KEY rather
        # than a PRIMARY KEY for this table.
        query = """
                  CREATE TABLE IF NOT EXISTS ClipKeywords2
                    (EpisodeNum    INTEGER, 
                     ClipNum       INTEGER, 
                     KeywordGroup  VARCHAR(100), 
                     Keyword       VARCHAR(255), 
                     Example       CHAR(1), 
                     UNIQUE KEY (EpisodeNum, ClipNum, KeywordGroup, Keyword))
                """
        # Add the appropriate Table Type to the CREATE Query
        query = SetTableType(hasBDB, query)
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
        query = SetTableType(hasBDB, query)
        # Execute the Query
        dbCursor.execute(query)
        # If we've gotten this far, return "true" to indicate success.
        return True


def get_db():
    """Get a connection object reference to the database.  If a connection
    has not yet been established, then create the connection."""
    global _user
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
        (userName, password, dbServer, databaseName) = UsernameForm.GetValues()
        # Check for the validity of the data.
        # The single-user version of Transana needs only the Database Name.  The multi-user version of
        # Transana requires all four values.
        if (databaseName == '') or \
           ((not TransanaConstants.singleUserVersion) and ((userName == '') or (password == '') or (dbServer == ''))):
            # If the Username and Password form is not filled out completely, the get_db method fails and returns "None"
            return None
        # Otherwise, all data was provided by the user.
        else:
            try:
                # Assign the username to a global variable
                _user = userName
                # try to Connect to the identified Database
                if TransanaConstants.singleUserVersion:
                    # The single-user version requires only the Database name
                    _dbref = MySQLdb.connect(db=databaseName)
                    TransanaGlobal.configData.database = databaseName
                else:
                    # The multi-user version requires all information to connect to the database
                    _dbref = MySQLdb.connect(host=dbServer, user=userName, passwd=password, db=databaseName)
                    # Put the Host and Databse names in the Configuration Data
                    # so that the same connection can be the default for the next logon
                    TransanaGlobal.configData.host = dbServer
                    TransanaGlobal.configData.database = databaseName
            # If the Database Connection fails, an exception is raised.
            except:
                # Set the Database Reference to "None" to indicate failure
                _dbref = None

                # Let's query the database to find out what databases have been defined, to see if the user
                # is attempting to connect to a NEW database.
                # Connect to the MySQL Server without specifying a Database
                if TransanaConstants.singleUserVersion:
                    # The single-user version requires no parameters
                    dbConn = MySQLdb.connect()
                else:
                    try:
                        # The multi-user version requires all information except the database name
                        dbConn = MySQLdb.connect(host=dbServer, user=userName, passwd=password)
                    # If the connection fails here, it's not the database name that was the problem.
                    except:
                        dbConn = None
                # If we made a connection to MySQL...
                if dbConn != None:
                    # Get a Database Cursor
                    dbCursor = dbConn.cursor()
                    # Query the Database about what Database Names have been defined
                    dbCursor.execute('SHOW DATABASES')
                    # Get all of the defined Database Names from the Database Cursor
                    dbs = dbCursor.fetchall()
                    # Initialize a flag that says the desired Database Name has not been found
                    dbFound = False
                    # Iterate through the list of databases to see if the desired Database Name exists.
                    # (We can't use "in" because we need a case insensitive search.)
                    for (db,) in dbs:
                        # Compare the Database Name with the item in the list returned from MySQL
                        if string.upper(databaseName) == string.upper(db):
                            # If they match, indicate that the Database was Found
                            dbFound = True
                            # Put the Host and Database names in the Configuration Data
                            # so that the same connection can be the default for the next logon
                            TransanaGlobal.configData.host = dbServer
                            TransanaGlobal.configData.database = databaseName
                            # and we can stop looking!
                            break
                    # If the Database Name was not found ...
                    if not dbFound:
                        # ... prompt the user to see if they want to create a new Database.
                        # First, create the Prompt Dialog
                        # NOTE:  This does not use Dialogs.ErrorDialog because it requires a Yes/No reponse
                        dlg = wx.MessageDialog(None, _('Database "%s" does not exist.  Would you like to create it?\n(If you do not have rights to create a database, see your system administrator.)') % databaseName, _('Transana Database Error'), wx.YES_NO | wx.ICON_ERROR)
                        # Display the Dialog
                        result = dlg.ShowModal()
                        # Clean up after the Dialog
                        dlg.Destroy()
                        # If the user wants to create a new Database ...
                        if result == wx.ID_YES:
                            try:
                                # ... create the Database ...
                                dbCursor.execute('CREATE DATABASE %s' % databaseName)
                                # ... specify that the new database should be used ...
                                dbCursor.execute('USE %s' % databaseName)
                                # and set the database reference to this database connection.
                                _dbref = dbConn
                                # Close the Database Cursor.
                                dbCursor.close()
                                # Put the Host and Database names in the Configuration Data
                                # so that the same connection can be the default for the next logon
                                TransanaGlobal.configData.host = dbServer
                                TransanaGlobal.configData.database = databaseName
                            # If the Create fails ...
                            except:
                                # ... the user probably lacks CREATE parmission in the Database Rights structure.
                                # Create an error message Dialog
                                dlg = Dialogs.ErrorDialog(None, _('Database Creation Error.\nYou specified an illegal database name, or do not have rights to create a database.\nTry again with a simple database name (with no punctuation or spaces), or see your system administrator.'))
                                # Display the Error Message.
                                result = dlg.ShowModal()
                                # Clean up the Error Message
                                dlg.Destroy()
                                # Close the Database Cursor
                                dbCursor.close()
                                # Close the Database Connection
                                dbConn.close()
                        # If the user indicates s/he does not want to create a new database ...
                        else:
                            # Close the Database Cursor
                            dbCursor.close()
                            # Close the Database Connection
                            dbConn.close()

    return _dbref

def close_db():
    """ This method flushes all database tables (saving data to disk) and closes the Database Connection. """
    # obtain the Database
    db = get_db()

    if db != None:
        # Obtain a Database Cursor
        dbCursor = db.cursor()

        # Let's FLUSH all tables, ensuring that data gets saved to the DB.
        query = "FLUSH TABLE Series2"
        dbCursor.execute(query)
        query = "FLUSH TABLE Episodes2"
        dbCursor.execute(query)
        query = "FLUSH TABLE Collections2"
        dbCursor.execute(query)
        query = "FLUSH TABLE Clips2"
        dbCursor.execute(query)
        query = "FLUSH TABLE Transcripts2"
        dbCursor.execute(query)
        query = "FLUSH TABLE Notes2"
        dbCursor.execute(query)
        query = "FLUSH TABLE Keywords2"
        dbCursor.execute(query)
        query = "FLUSH TABLE ClipKeywords2"
        dbCursor.execute(query)
        query = "FLUSH TABLE CoreData2"
        dbCursor.execute(query)

        # Close the Database Cursor
        dbCursor.close()
        
        # Close the Database itself
        db.close()

    global _dbref
    # Remove all reference to the database
    _dbref = None
    

   
def get_username():
    """Get the name of the current database user."""
    return _user

#def get_collection_num(Name):
#    """Look up a named Collection's Collection Number without loading the
#    full object."""

def list_of_series():
    """Get a list of all Series record names."""
    l = []
    query = "SELECT SeriesNum, SeriesID FROM Series2 ORDER BY SeriesID\n"
    DBCursor = get_db().cursor()
    DBCursor.execute(query)
    for row in fetchall_named(DBCursor):
        l.append((row['SeriesNum'], row['SeriesID']))
    DBCursor.close()
    return l
    

def list_of_episodes_for_series(SeriesName):
    """Get a list of all Episodes contained within a named Series."""
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
        l.append((row['EpisodeNum'], row['EpisodeID'], row['SeriesNum']))
    DBCursor.close()
    return l

def list_transcripts(SeriesName, EpisodeName):
    """Get a list of all Transcripts for the named Episode within the
    named Series."""
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
        l.append((row['TranscriptNum'], row['TranscriptID'], row['EpisodeNum']))
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
        l.append((row['CollectNum'], row['CollectID'], row['ParentCollectNum']))
    DBCursor.close()
    return l

def list_of_clips_by_collection(CollectionID, ParentNum):
    """Get a list of all Clips for a named Collection."""
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
        l.append((row['ClipNum'], row['ClipID'], row['CollectNum']))
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
        clipList.append((clipNum, clipID, collectNum))
    cursor.close()
    return clipList

def list_of_clips_by_episode(EpisodeNum, TimeCode=None):
    """Get a list of all Clips that have been created from a given Collection
    Number.  Optionally restrict list to contain only a given timecode."""
    l = []
    if TimeCode == None:
        query = """
                  SELECT a.ClipID, a.ClipStart, a.ClipStop, b.CollectID, b.ParentCollectNum FROM Clips2 a, Collections2 b
                    WHERE a.CollectNum = b.CollectNum AND
                          a.EpisodeNum = %s
                    ORDER BY a.ClipStart, b.CollectID, a.ClipID
                """
        args = (EpisodeNum)
    else:
        query = """
                  SELECT a.ClipID, a.ClipStart, a.ClipStop, b.CollectID, b.ParentCollectNum FROM Clips2 a, Collections2 b
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
        l.append(row)
    DBCursor.close()
    return l


def list_of_notes(** kwargs):
    """Get a list of Note IDs for the given Series, Episode, Collection,
    or Clip numbers.  Parameters are passed as keyword
    arguments, where valid parameters are Transcript, Episode, Series, Clip, and
    Collection.
    
    Examples: list_of_notes(Series=12)
              list_of_notes(Episode=14)"""
    notelist = []
   
    if kwargs.has_key("Series"):
        query = """
        SELECT NoteID From Notes2
            WHERE   SeriesNum = %s
        """
        values = (kwargs['Series'],)
    elif kwargs.has_key("Episode"):
        query = """
        SELECT NoteID FROM Notes2
            WHERE   EpisodeNum = %s
        """
        values = (kwargs['Episode'],)
    elif kwargs.has_key("Transcript"):
        query = """
        SELECT NoteID FROM Notes2
            WHERE   TranscriptNum = %s
        """
        values = (kwargs['Transcript'],)
    elif kwargs.has_key("Collection"):
        query = """
        SELECT NoteID FROM Notes2
            WHERE   CollectNum = %s
        """
        values = (kwargs['Collection'],)
    elif kwargs.has_key("Clip"):
        query = """
        SELECT NoteID FROM Notes2
            WHERE   ClipNum = %s
        """
        values = (kwargs['Clip'])
    else:
        return []   # Should we raise an exception?

    query = query + "   ORDER BY NoteID\n"
    db = get_db()
    DBCursor = db.cursor()
    DBCursor.execute(query, values)
    r = DBCursor.fetchall()    # return array of tuples with results
    for tup in r:
        notelist.append(tup[0])
    DBCursor.close()
    return notelist
   
    
# I need to carefully reconsider this interface.  Originally it would have
# been passed strings.  However, it made it cleaner to just pass an object
# reference and let the method just figure out what it is.  In this state,
# it seems like this function would be more appropriate as a method for
# the individual objects.
# So I'll either move this function into the Series/Episode/Collection/Clip
# objects as methods, or I'll change it to the original strings-based
# interface.  I first need to see how this method is used to see what's best.
# - Nate
# UPDATE: I decided to keep the original "strings" interface, but modified
# syntatically to be more Python-esque using keyword dictionary arguments.
# so this function is now obsolete but left here for now
#def __obs_list_of_notes(object):
#    """Get a list of all notes for the given Series, Episode, Collection,
#    or Clip object."""
#    notelist = []
   
    # this little trick let's us use the type definitions in this method
    # without getting in trouble for circular module dependencies
#    from Episode import Episode
#    from Series import Series
#    from Clip import Clip
#    from Collection import Collection
#    t = type(object)

#    if t == Episode:
        # Get the Series name for the episode
#        s = Series()
#        s.db_load_by_num(object.series_number)
#        query = "SELECT NoteID FROM Notes2 a, Series2 b, Episodes2 c\n" + \
#                "   WHERE b.SeriesNum = c.SeriesNum and\n" + \
#                "         a.EpisodeNum = c.EpisodeNum and\n" + \
#                "         SeriesID = '" + s.id + "' and\n" + \
#                "         EpisodeID = '" + object.id + "'\n" 
#        del s
#    elif t == Series:
#        query = "SELECT NoteID FROM Notes2 a, Series2 b\n" + \
#                "   WHERE a.SeriesNum = b.SeriesNum and\n" + \
#                "       SeriesID = '" + object.id + "'\n"
#    elif t == Clip:
        # Get the Collection name for the Clip
#        c = Collection()
#        c.db_load_by_num(object.collection_num)
#        query = "SELECT NoteID FROM Notes2 a, Collections2 b, Clips2 c\n" + \
#                "   WHERE b.CollectNum = c.CollectNum and \n" + \
#                "         a.ClipNum = c.ClipNum and\n" + \
#                "         CollectID = '" + c.id + "' and\n" + \
#                "         ClipID = '" + object.id + "'\n"
#        del c
#    elif t == Collection:
#        query = "SELECT NoteID FROM Notes2 a, Collections2 b\n" + \
#                "   WHERE a.CollectNum = b.CollectNum and\n" + \
#                "         CollectID = '" + object.id + "'\n" 
#    else:           # unknown object passed
#        return []   # FIXME: Exception?

#    query = query + "   ORDER BY NoteID\n"
#    db = get_db()
#    DBCursor = db.cursor()
#    DBCursor.execute(query)
#    r = DBCursor.fetchall()    # return array of tuples with results
#    for tup in r:
#        notelist.append(tup[0])
#    DBCursor.close()
#    return notelist

def list_of_keyword_groups():
    """Get a list of all keyword groups."""
    l = []
    query = "SELECT KeywordGroup FROM Keywords2 GROUP BY KeywordGroup\n"
    DBCursor = get_db().cursor()
    DBCursor.execute(query)
    for row in fetchall_named(DBCursor):
        l.append(row['KeywordGroup'])
    DBCursor.close()
    return l


def list_of_keywords_by_group(KeywordGroup):
    """Get a list of all keywords for the named Keyword group."""
    l = []
    query = \
    "SELECT Keyword FROM Keywords2 WHERE KeywordGroup = %s ORDER BY Keyword\n"
    DBCursor = get_db().cursor()
    DBCursor.execute(query, KeywordGroup)
    for row in fetchall_named(DBCursor):
        l.append(row['Keyword'])
    DBCursor.close()
    return l

def list_of_all_keywords():
    """Get a list of all keywords in the Transana database."""
    l = []
    query = \
    "SELECT Keyword FROM Keywords2 ORDER BY Keyword\n"
    DBCursor = get_db().cursor()
    DBCursor.execute(query)
    for row in fetchall_named(DBCursor):
        l.append(row['Keyword'])
    DBCursor.close()
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
        keywordExampleList.append(dbRowData)
    dbCursor.close()
    return keywordExampleList

def SetKeywordExampleStatus(kwg, kw, clipNum, exampleValue):
    """ The SetKeywordExampleStatus(kwg, kw, clipNum, exampleValue) method sets the
        Example value in the ClipKeywords table for the appropriate KWG, KW, Clip
        combination.  Set Example to 1 to specify a Keyword Example, 0 to remove
        it from being an example without deleting the keyword for the Clip. """
    query = """UPDATE ClipKeywords2
                 SET Example = %s
                 WHERE KeywordGroup = %s AND
                       Keyword = %s AND
                       ClipNum = %s"""

    # print query % (exampleValue, kwg, kw, clipNum)

    dbCursor = get_db().cursor()
    dbCursor.execute(query, (exampleValue, kwg, kw, clipNum))

    # If rowcount == 0, no rows were affected by the update.  That is, the Keyword
    # had not previously been assigned to this Clip, so there was no record to update.
    # In this case, we have to ADD the record!
    if dbCursor.rowcount == 0:
        insert_clip_keyword(0, clipNum, kwg, kw, 1)
    dbCursor.close()


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
    query = "SELECT * FROM Episodes2 "
    # Execute the Query
    dbCursor.execute(query)
    # Fetch all the Database Results, and process them row by row
    for row in dbCursor.fetchall():
        # The Media Filename is Element 4
        file = row[4]
        # Some Filenames have doubled backslashes, though not all do.  Let's eliminate them if they are present.
        # (Python requires a double backslash in a string to represent a single backslash, so this replaces double
        # backslashes ('\\') with single ones ('\') even though it looks like it replaces quadruples with doubles.)
        file = string.replace(file, '\\\\', '\\')
        # Compare the Video Root filePath passed in with the front portion of the File Name from the Database.
        if filePath == file[:len(filePath)]:
            # If they are the same, increment the Episode Counter
            episodeCount += 1
            # If update is True, we should update the records we find.
            if update:
                # Import Transana's Episode definition
                import Episode
                # Load the Episode using the Episode Number (Element 0 in the Row).
                tempEpisode = Episode.Episode(row[0])
                # We need a "try .. except" block to catch record lock exceptions
                try:
                    # Try to lock the Episode
                    tempEpisode.lock_record()
                    # Remove the Video Root from the File Name
                    tempEpisode.media_filename = file[len(filePath):]
                    # Save the Episode
                    tempEpisode.db_save()
                    # Unlock the Episode
                    tempEpisode.unlock_record()
                # Catch failed record locks or Saves
                except:
                    # TODO:  Detect exception type and customize the error message below.
                    import sys
                    print sys.exc_info()

                    # If it fails, set the transactionStatus Flag to False
                    transactionStatus = False
                    # and stop processing records.
                    break

    # If we've already failed, there's no reason to continue trying.  Make sure we should continue.
    if transactionStatus:
        # Initialize the Clip Counter
        clipCount = 0
        # Create the Query for the Episode Table
        query = "SELECT * FROM Clips2 "
        # Execute the Query
        dbCursor.execute(query)
        # Fetch all the Database Results, and process them row by row
        for row in dbCursor.fetchall():
            # The Media Filename is Element 5
            file = row[5]
            # Some Filenames have doubled backslashes, though not all do.  Let's eliminate them if they are present.
            # (Python requires a double backslash in a string to represent a single backslash, so this replaces double
            # backslashes ('\\') with single ones ('\') even though it looks like it replaces quadruples with doubles.)
            file = string.replace(file, '\\\\', '\\')
            # Compare the Video Root filePath passed in with the front portion of the File Name from the Database.
            if filePath == file[:len(filePath)]:
                # If they are the same, increment the Clip Counter
                clipCount += 1
                # If update is True, we should update the record we find.
                if update:
                    # Import Transana's Clip definition
                    import Clip
                    # Load the Clip using the Clip Number (Element 0 in the Row).
                    tempClip = Clip.Clip(row[0])
                    # We need a "try .. except" block to catch record lock exceptions
                    try:
                        # Try to lock the Episode
                        tempClip.lock_record()
                        # Remove the Video Root from the File Name
                        tempClip.media_filename = file[len(filePath):]
                        # Save the Clip
                        tempClip.db_save()
                        # Unlock the Clip
                        tempClip.unlock_record()
                    # Catch failed record locks or Saves
                    except:
                        # TODO:  Detect exception type and customize the error message below.
                        import sys
                        print sys.exc_info()
                        
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
    DBCursor = get_db().cursor()
    for table in ("Keywords2", "ClipKeywords2"):
        query = """DELETE FROM %s
        WHERE KeywordGroup = %%s
        """ % table
        DBCursor.execute(query, name)

    DBCursor.close()

def delete_keyword(group, kw_name):
    """Delete a Keyword from the database."""

    DBCursor = get_db().cursor()
    DBCursor.execute("BEGIN")
    query = """SELECT a.KeywordGroup, a.Keyword, b.EpisodeID, b.RecordLock
        FROM ClipKeywords2 a, Episodes2 b
        WHERE   a.KeywordGroup = %s AND
                a.Keyword = %s AND
                a.EpisodeNum <> %s AND
                a.EpisodeNum = b.EpisodeNum AND
                b.RecordLock <> %s
    """

    DBCursor.execute(query, (group, kw_name, 0, ""))
    t = ""
    for row in fetchall_named(DBCursor):
        t = "%s  Episode %s is locked by %s\n" % \
                    (t, row['EpisodeID'], row['RecordLock'])

    # Clips next
    query = """SELECT a.KeywordGroup, a.Keyword, c.ClipID, c.RecordLock
        FROM ClipKeywords2 a, Clips2 c
        WHERE   a.KeywordGroup = %s AND
                a.Keyword = %s AND
                a.ClipNum <> %s AND
                a.ClipNum = c.ClipNum AND
                c.RecordLock <> %s
    """
    DBCursor.execute(query, (group, kw_name, 0, ""))
    for row in fetchall_named(DBCursor):
        t = "%s Clip %s is locked by %s\n" % \
                    (t, row['ClipID'], row['RecordLock'])

    if t == "":
        # Delphi Transana had a confirmation dialog here, but we won't do
        # that here (do it before calling this function).

        # Build and execute the SQL to delete a word from the list.
        query = """DELETE FROM Keywords2
            WHERE   KeywordGroup = %s AND
                    KeyWord = %s
        """
        DBCursor.execute(query, (group, kw_name))
        # Now delete all instances of this keywordgroup/keyword combo in the
        # Clipkeywords file
        query = """DELETE FROM ClipKeywords2
            WHERE   KeywordGroup = %s AND
                    KeyWord = %s
        """
        DBCursor.execute(query, (group, kw_name))

        # Finish the transaction
        DBCursor.execute("COMMIT")
    else:
        DBCursor.execute("ROLLBACK")
        DBCursor.close()
        raise Exception, _("Unable to delete keyword.\n") + t
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

def UpdateDBFilenames(parent, filePath, fileList):
    # import Transana's Episode Object
    import Episode
    # import Transana's Clip Object
    import Clip
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

    # Let's count the number of records changed
    episodeCounter = 0
    clipCounter = 0

    # Go through the fileList and run the query repeatedly
    for fileName in fileList:
        # Add a "%" character to the beginning of the File Name so that the "LIKE" operator will work

        # Execute the Episode Query
        DBCursor.execute(episodeQuery, '%' + fileName)

        # Iterate through the records returned from the Database
        for (episodeNum, ) in DBCursor.fetchall():
            # Load the Episode
            tempEpisode = Episode.Episode(episodeNum)
            # Be ready to catch exceptions
            try:
                # Lock the Record
                tempEpisode.lock_record()
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
                # Indicate that we have failed.
                success = False
                # Don't bother to contine processing DB Records
                break
        # If we failed during Episodes, there's no reason to look at Clips
        if not success:
            break
        
        # Execute the Clip Query
        DBCursor.execute(clipQuery, '%' + fileName)

        # Iterate through the records returned from the Database
        for (clipNum, ) in DBCursor.fetchall():
            # Load the Clip
            tempClip = Clip.Clip(clipNum)
            # Be ready to catch exceptions
            try:
                # Lock the Record
                tempClip.lock_record()
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
                # Indicate that we have failed.
                success = False
                # Don't bother to contine processing DB Records
                break
    # If there have been no problems, Commit the Transaction to the Database
    if success:
        DBCursor.execute("COMMIT")
        msg = _("%s Episode Records have been updated.\n%s Clip Records have been updated.") % (episodeCounter, clipCounter)
        infodlg = Dialogs.InfoDialog(None, msg)
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

def DeleteDatabase(username, password, server, database):
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
                dbConn = MySQLdb.connect(host=server, user=username, passwd=password)
                # If the connection fails here, it's not the database name that was the problem.
            except:
                dbConn = None
        # If we made a connection to MySQL...
        if dbConn != None:
            try:
                dlg = wx.MessageDialog(None, _('Are you sure you want to delete Database "%s"?\nAll data in this database will be permanently deleted and cannot be recovered.') % database, _('Delete Database'), wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
                result = dlg.ShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    # Give the user another chance to back out!
                    dlg = wx.MessageDialog(None, _('Are you ABSOLUTELY SURE you want to delete Database "%s"?\nAll data in this database will be permanently deleted and cannot be recovered.') % database, _('Delete Database'), wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
                    result = dlg.ShowModal()
                    dlg.Destroy()
                    if result == wx.ID_YES:
                        # Get a Database Cursor
                        dbCursor = dbConn.cursor()
                        # Query the Database to delete the desired database
                        dbCursor.execute('DROP DATABASE IF EXISTS %s' % database)
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
