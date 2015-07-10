# Copyright (C) 2003 - 2015 The Board of Regents of the University of Wisconsin System 
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

""" This module implements the Snapshot class as part of the Data Objects. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx
# import Python's os module
import os
# import Python's types module
import types
# import Transana's Clip Keyword Object
import ClipKeywordObject
# import Transana's Collection Object
import Collection
# Import the Transana base Data Object
import DataObject
# Import the Transana Database Interface
import DBInterface
# import Transana's Dialogs
import Dialogs
# Import Transana's Episode Object
import Episode
# import Transana's Miscellaneous Functions
import Misc
# import Transana's Note Object
import Note
# import Transana's Library Object
import Library
# Import Transana's exceptions
from TransanaExceptions import *
# Import the Transana Constants
import TransanaConstants
# Import Transana's Global Variables
import TransanaGlobal
# import Transana's Transcript Object
import Transcript

class Snapshot(DataObject.DataObject):
    """ This class defines the structure for a snapshot object.  A snapshot object
        describes a still image file that can be coded. """

    def __init__(self, num_or_id=0, collNum = None, suppressEpisodeError = False):
        """ Initialize a Snapshot object.  The suppressEpisodeError parameter indicates that the Snapshot is
            intended for deletion, so we can ignore the Context Episode error message if that arises. """

        # Remember if we're opening this snapshot just for deletion
        self.suppressEpisodeError = suppressEpisodeError

        # Create a Data Object
        DataObject.DataObject.__init__(self)
        # By default, use the Video Root folder if one has been defined
        self.useVideoRoot = (TransanaGlobal.configData.videoPath != '')

        if (isinstance(num_or_id, long) or isinstance(num_or_id, int)) and (num_or_id != 0):
            self.db_load(num_or_id)
        elif isinstance(num_or_id, unicode) or isinstance(num_or_id, str):
            self.db_load_by_name(num_or_id, collNum)
        else:
            self.number = 0
            self.id = ''
            self.collection_num = 0
            self.collection_id = ''
            self.image_filename = ""
            self.image_scale = 0.0
            self.image_coords = (0.0, 0.0)
            self.image_size = (0, 0)
            self.series_num = 0
            self.series_id = ''
            self.episode_num = 0
            self.episode_id = ''
            self.transcript_num = 0
            self.transcript_id = ''
            self.episode_start = 0
            self.episode_duration = 0
            self.comment = ''
            self.sort_order = 0
            self.codingObjects = {}
            self.keywordStyles = {}

    def __repr__(self):
        str = 'Snapshot Object Definition:\n'
        str += "  number           = %s\n" % self.number
        str += "  id               = %s\n" % self.id
        str += "  collection_num   = %s\n" % self.collection_num 
        str += "  collection_id    = %s\n" % self.collection_id
        str += "  image_filename   = %s\n" % self.image_filename
        str += "  image_scale      = %15.9f\n" % self.image_scale
        str += "  image_coords     = (%15.5f, %15.5f)\n" % self.image_coords
        str += "  image_size       = (%d, %d)\n" % self.image_size
        str += "  Library_num       = %s\n" % self.series_num
        str += "  Library_id        = %s\n" % self.series_id
        str += "  episode_num      = %s\n" % self.episode_num
        str += "  episode_id       = %s\n" % self.episode_id
        str += "  transcript_num   = %s\n" % self.transcript_num
        if self.transcript_num > 0:
            str += "  transcript_id    = %s\n" % self.transcript_id
        str += "  episode_start    = %s (%s)\n" % (self.episode_start, Misc.time_in_ms_to_str(self.episode_start))
        str += "  episode_duration = %s (%s)\n" % (self.episode_duration, Misc.time_in_ms_to_str(self.episode_duration))
        str += "  comment          = %s\n" % self.comment
        str += "  sort_order       = %s\n" % self.sort_order
        str += "isLocked = %s\n" % self._isLocked
        str += "recordlock = %s\n" % self.recordlock
        str += "locktime = %s\n" % self.locktime
        str += "  lastsavetime     = %s\n" % self.lastsavetime
        for kws in self.keyword_list:
            str = str + "\nKeyword:  %s" % kws.keywordPair
        str += "\ncodingObjects:"
        for x in self.codingObjects.keys():
            str += "\n  %d" % x
            for y in self.codingObjects[x].keys():
                str += "\n    %s  %s" % (y, self.codingObjects[x][y])
            str += '\n'
        str += "\nkeywordStyles:"
        for x in self.keywordStyles.keys():
            str += "\n  (%s, %s)" % x
            for y in self.keywordStyles[x].keys():
                str += "\n    %s  %s" % (y, self.keywordStyles[x][y])
            str += '\n'
        str = str + '\n'
        return str.encode('utf8')
        
    def __eq__(self, other):
        """ Object Equality function """

#        print "Snapshot.__eq__():", len(self.__dict__.keys()), len(other.__dict__.keys())
#        for key in self.__dict__.keys():
#            print key, self.__dict__[key] == other.__dict__[key]
#        print

        if other == None:
            return False
        else:
            return self.__dict__ == other.__dict__

# Public methods

    def db_load(self, num):
        """Load a record by record number."""
        self.clear()
        # Get a database Connection
        db = DBInterface.get_db()
        # Craft a query to get the Clip data
        query = """
        SELECT *
          FROM Snapshots2 a
          WHERE a.SnapshotNum = %s
        """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, (num, ))
        # rowcount doesn't work for sqlite!
        if TransanaConstants.DBInstalled == 'sqlite3':
            n = 1
        else:
            n = c.rowcount
        # If we don't get exactly one result ...
        if (n != 1):
            # ... close the database cursor ...
            c.close()
            # ... clear the current Snapshot object ...
            self.clear()
            # Raise an exception indicating the data was not found
            raise RecordNotFoundError, (num, n)
        # If we get exactly one result ...
        else:
            # ... get the data from the cursor
            r = DBInterface.fetch_named(c)
            # If sqlite and no results ...
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}):
                # ... close the database cursor ...
                c.close()
                # ... clear the current Snapshot object ...
                self.clear()
                # ... and raise an exception
                raise RecordNotFoundError, (num, 0)
            # ... load the data into the Snapshot Object
            self._load_row(r)
            # Refresh the Keywords
            self.refresh_keywords()
        # Create a Query to get the codingObjects
        query = """ SELECT SnapshotNum, KeywordGroup, Keyword, x1, y1, x2, y2, visible
                      FROM SnapshotKeywords2
                      WHERE SnapshotNum = %s """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, (num, ))
        # Get the query results
        results = c.fetchall()
        # Initialize the counter
        counter = 0
        # Put the results into the Snapshot Object
        for (SnapshotNum, KeywordGroup, Keyword, x1, y1, x2, y2, visible) in results:
            self.codingObjects[counter] = {'x1'             :  x1,
                                         'y1'             :  y1,
                                         'x2'             :  x2,
                                         'y2'             :  y2,
                                         'keywordGroup'   :  DBInterface.ProcessDBDataForUTF8Encoding(KeywordGroup),
                                         'keyword'        :  DBInterface.ProcessDBDataForUTF8Encoding(Keyword),
                                         'visible'        :  visible == '1'}
            counter += 1
        # Create a Query to get the keywordStyles
        query = """ SELECT SnapshotNum, KeywordGroup, Keyword, DrawMode, LineColorName, LineColorDef, LineWidth, LineStyle
                      FROM SnapshotKeywordStyles2
                      WHERE SnapshotNum = %s """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, (num, ))
        # Get the query results
        results = c.fetchall()
        # Put the results into the Snapshot Object
        for (SnapshotNum, KeywordGroup, Keyword, DrawMode, LineColorName, LineColorDef, LineWidth, LineStyle) in results:
            self.keywordStyles[(DBInterface.ProcessDBDataForUTF8Encoding(KeywordGroup), DBInterface.ProcessDBDataForUTF8Encoding(Keyword))] = \
                { 'drawMode'       :  DrawMode,
                  'lineColorName'  :  DBInterface.ProcessDBDataForUTF8Encoding(LineColorName),
                  'lineColorDef'   :  LineColorDef,
                  'lineWidth'      :  "%d" % LineWidth,
                  'lineStyle'      :  LineStyle  }

        # Close the database cursor
        c.close()
        self._sync_snapshot()

##        tmpDlg = Dialogs.InfoDialog(None, self.__repr__(), "Snapshot.db_load_by_num()")
##        tmpDlg.ShowModal()
##        tmpDlg.Destroy()

    def db_load_by_name(self, name, collNum):
        """Load a record by Name and Parent Collection Number."""
        self.clear()
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            name = name.encode(TransanaGlobal.encoding)
        # Get a database Connection
        db = DBInterface.get_db()
        # Craft a query to get the Clip data
        query = """
        SELECT *
          FROM Snapshots2 a
          WHERE a.SnapshotID = %s AND
                a.CollectNum = %s
        """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, (name, collNum))
        # Get the number of rows returned
        # rowcount doesn't work for sqlite!
        if TransanaConstants.DBInstalled == 'sqlite3':
            n = 1
        else:
            n = c.rowcount
        # If we don't get exactly one result ...
        if (n != 1):
            # ... close the database cursor ...
            c.close()
            # ... clear the current Snapshot object ...
            self.clear()
            # Raise an exception indicating the data was not found
            raise RecordNotFoundError, (name, n)
        # If we get exactly one result ...
        else:
            # ... get the data from the cursor
            r = DBInterface.fetch_named(c)
            # if sqlite and no data in the cursor ...
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}):
                # ... close the database cursor ...
                c.close()
                # ... clear the current Snapshot object ...
                self.clear()
                # Raise an exception
                raise RecordNotFoundError, (name, 0)
            # ... load the data into the Snapshot Object
            self._load_row(r)
            # Refresh the Keywords
            self.refresh_keywords()
        # Create a Query to get the codingObjects
        query = """ SELECT SnapshotNum, KeywordGroup, Keyword, x1, y1, x2, y2, visible
                      FROM SnapshotKeywords2
                      WHERE SnapshotNum = %s """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, (self.number, ))
        # Get the query results
        results = c.fetchall()
        # Initialize the counter
        counter = 0
        # Put the results into the Snapshot Object
        for (SnapshotNum, KeywordGroup, Keyword, x1, y1, x2, y2, visible) in results:
            self.codingObjects[counter] = {'x1'             :  x1,
                                         'y1'             :  y1,
                                         'x2'             :  x2,
                                         'y2'             :  y2,
                                         'keywordGroup'   :  DBInterface.ProcessDBDataForUTF8Encoding(KeywordGroup),
                                         'keyword'        :  DBInterface.ProcessDBDataForUTF8Encoding(Keyword),
                                         'visible'        :  visible == '1'}
            counter += 1
        # Create a Query to get the keywordStyles
        query = """ SELECT SnapshotNum, KeywordGroup, Keyword, DrawMode, LineColorName, LineColorDef, LineWidth, LineStyle
                      FROM SnapshotKeywordStyles2
                      WHERE SnapshotNum = %s """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, (self.number, ))
        # Get the query results
        results = c.fetchall()
        # Put the results into the Snapshot Object
        for (SnapshotNum, KeywordGroup, Keyword, DrawMode, LineColorName, LineColorDef, LineWidth, LineStyle) in results:
            self.keywordStyles[(DBInterface.ProcessDBDataForUTF8Encoding(KeywordGroup), DBInterface.ProcessDBDataForUTF8Encoding(Keyword))] = \
                { 'drawMode'       :  DrawMode,
                  'lineColorName'  :  DBInterface.ProcessDBDataForUTF8Encoding(LineColorName),
                  'lineColorDef'   :  LineColorDef,
                  'lineWidth'      :  "%d" % LineWidth,
                  'lineStyle'      :  LineStyle  }

        # Close the database cursor
        c.close()
        self._sync_snapshot()

    def db_save(self, use_transactions=True):
        """Save the record to the database using Insert or Update as appropriate."""

        # Define and implement Demo Version limits
        if TransanaConstants.demoVersion and (self.number == 0):
            # Get a DB Cursor
            c = DBInterface.get_db().cursor()
            # Find out how many records exist
            c.execute('SELECT COUNT(SnapshotNum) from Snapshots2')
            res = c.fetchone()
            c.close()
            # Define the maximum number of records allowed
            maxSnapshots = TransanaConstants.maxSnapshots
            # Compare
            if res[0] >= maxSnapshots:
                # If the limit is exceeded, create and display the error using a SaveError exception
                prompt = _('The Transana Demonstration limits you to %d Snapshot records.\nPlease cancel the "Add Snapshot" dialog to continue.')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                raise SaveError, prompt % maxSnapshots

        # Sanity checks
        if self.id == "":
            raise SaveError, _("Snapshot ID is required.")
        if (self.collection_num == 0):
            raise SaveError, _("Parent Collection number is required.")
        elif self.image_filename == "":
            raise SaveError, _("Image Filename is required.")
        # If a user Adjusts Indexes, it's possible to have a clip that starts BEFORE the media file.
        elif self.episode_start < 0.0:
            raise SaveError, _("Snapshot location in the Episode cannot be before the media file begins.")
        else:
            # videoPath probably has the OS.sep character, but we need the generic "/" character here.
            videoPath = TransanaGlobal.configData.videoPath
            # Determine if we are supposed to extract the Video Root Path from the Media Filename and extract it if appropriate
            if self.useVideoRoot and (videoPath == self.image_filename[:len(videoPath)]):
                tempImageFilename = self.image_filename[len(videoPath):]
            else:
                tempImageFilename = self.image_filename

            # Substitute the generic OS seperator "/" for the Windows "\".
            tempImageFilename = tempImageFilename.replace('\\', '/')
            # If we're not in Unicode mode ...
            if 'ansi' in wx.PlatformInfo:
                # ... we don't need to encode the string values, but we still need to copy them to our local variables.
                id = self.id
                comment = self.comment
            # If we're in Unicode mode ...
            else:
                # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
                # variables for the data.  We don't want to change the underlying object values.  Also, this way,
                # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
                id = self.id.encode(TransanaGlobal.encoding)
                tempImageFilename = tempImageFilename.encode(TransanaGlobal.encoding)
                videoPath = videoPath.encode(TransanaGlobal.encoding)
                comment = self.comment.encode(TransanaGlobal.encoding)

#        self._sync_snapshot()

        values = (id, self.collection_num, \
                      tempImageFilename, \
                      self.image_scale, self.image_coords[0], self.image_coords[1], \
                      self.image_size[0], self.image_size[1], \
                      self.episode_num, self.transcript_num, self.episode_start, self.episode_duration, \
                      comment, \
                      self.sort_order)
        if (self._db_start_save() == 0):

            if DBInterface.record_match_count("Snapshots2", \
                                ("SnapshotID", "CollectNum"), \
                                (id, self.collection_num) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Snapshot named "%s" already exists in this Collection.\nPlease enter a different Snapshot ID.'), 'utf8') % self.id
                else:
                    prompt = _('A Snapshot named "%s" already exists in this Collection.\nPlease enter a different Snapshot ID.') % self.id

                raise SaveError, prompt
            # insert the new record
            query = """
            INSERT INTO Snapshots2
                (SnapshotID, CollectNum,
                 ImageFile, ImageScale, ImageCoordsX, ImageCoordsY, ImageSizeW, ImageSizeH,
                 EpisodeNum, TranscriptNum, SnapshotTimeCode, SnapshotDuration, SnapshotComment,
                 SortOrder, LastSaveTime)
                VALUES
                (%s, %s, %s , %s, %s , %s, %s, %s, %s, %s, %s, %s, %s, %s, CURRENT_TIMESTAMP)
            """
        else:
            if DBInterface.record_match_count("Snapshots2", \
                            ("SnapshotID", "CollectNum", "!SnapshotNum"), \
                            (id, self.collection_num, self.number)) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Snapshot named "%s" already exists in this Collection.\nPlease enter a different Snapshot ID.'), 'utf8') % self.id
                else:
                    prompt = _('A Snapshot named "%s" already exists in this Collection.\nPlease enter a different Snapshot ID.') % self.id
                raise SaveError, prompt

            # update the record
            query = """
            UPDATE Snapshots2
                SET SnapshotID = %s,
                    CollectNum = %s,
                    ImageFile = %s,
                    ImageScale = %s,
                    ImageCoordsX = %s,
                    ImageCoordsY = %s,
                    ImageSizeW = %s,
                    ImageSizeH = %s,
                    EpisodeNum = %s,
                    TranscriptNum = %s,
                    SnapshotTimeCode = %s,
                    SnapshotDuration = %s,
                    SnapshotComment = %s,
                    SortOrder = %s,
                    LastSaveTime = CURRENT_TIMESTAMP
                WHERE SnapshotNum = %s
            """
            values = values + (self.number,)
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        c = DBInterface.get_db().cursor()
        if use_transactions:
            c.execute('BEGIN')
        # Execute the query that puts the data in the database
        c.execute(query, values)
        # if our object number is 0, we have a NEW Snapshot
        if self.number == 0:
            # ... then flag that we've change the object number
            numberChanged = True
            # If we are dealing with a brand new Snapshot, it does not yet know its
            # record number.  It HAS a record number, but it is not known yet.
            # The following query should produce the correct record number.
            query = """
                      SELECT SnapshotNum, LastSaveTime FROM Snapshots2
                      WHERE SnapshotID = %s AND
                            CollectNum = %s
                    """
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Execute the query
            c.execute(query, (id, self.collection_num))
            # Get the query's results.
            data = c.fetchall()
            # If we have exactly one record ...
            if len(data) == 1:
                # ... get the object number
                self.number = data[0][0]
                # ... and update the LastSaveTime
                self.lastsavetime = data[0][1]
            # If we have something other than exactly one record ...
            else:
                # if we're using Transactions ...
                if use_transaction:
                    # ... roll back the transaction ...
                    c.execute('ROLLBACK')
                # ... and raise an exception
                raise RecordNotFoundError, (self.id, len(data))
        # If we've updated an existing record ...
        else:
            # ... then we haven't changed the object's number 
            numberChanged = False
            # If we are dealing with an existing Snapshot, delete all the Keywords
            # in anticipation of putting them all back in
            DBInterface.delete_all_keywords_for_a_group(0, 0, 0, 0, self.number)
            # Define the query for deleting Snapshot Keywords
            query = "DELETE FROM SnapshotKeywords2 WHERE SnapshotNum = %s"
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # execute the query
            c.execute(query, (self.number, ))
            # Define the query for deleting Snapshot Keyword Styles
            query = "DELETE FROM SnapshotKeywordStyles2 WHERE SnapshotNum = %s"
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Execute the query
            c.execute(query, (self.number, ))

        # Initialize a blank error prompt
        prompt = ''
        # Add the Snapshot keywords back.  Iterate through the Keyword List
        for kws in self._kwlist:
            # Try to add the Snapshot Keyword record.  If it is NOT added, the keyword has been changed by another user!
            if not DBInterface.insert_clip_keyword(0, 0, 0, 0, self.number, kws.keywordGroup, kws.keyword, kws.example):
                # if the prompt isn't blank ...
                if prompt != '':
                    # ... add a couple of line breaks to it
                    prompt += u'\n\n'
                # Add the current keyword to the error prompt
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt += unicode(_('Keyword "%s : %s" cannot be added to Snapshot "%s".\nAnother user must have edited the keyword while you were adding it.'), 'utf8') % (kws.keywordGroup, kws.keyword, self.id)

        # If there is an error prompt ...
        if prompt != '':
            # If the Episode Number was changed ...
            if numberChanged:
                # ... change it back to zero!!
                self.number = 0
            # If we're using Transactions ...
            if use_transactions:
                # Undo the database save transaction
                c.execute('ROLLBACK')
            # Close the Database Cursor
            c.close()
            # Complete the error prompt
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt += u'\n\n' + unicode(_('Please remove, and possibly replace, this keyword.'), 'utf8')
            # ... raise a SaveError exception using the error prompt
            raise SaveError, prompt
        # If there's no error prompt ...
        else:

            # Let's get a temporary list of all existing keywords
            tmpKeywordList = DBInterface.list_of_all_keywords()
            # Let's also build a temporary list of all the Detail Codes that are used in this Snapshot.
            # Initialize it first.
            tmpCodeList = []
            # Define a query for inserting the Snapshot Keywords
            query = """ INSERT INTO SnapshotKeywords2
                          (SnapshotNum, KeywordGroup, Keyword, x1, y1, x2, y2, visible)
                          VALUES
                          (%s, %s, %s, %s, %s, %s, %s, %s) """
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Get the Coding Objects' Key values ...
            keys = self.codingObjects.keys()
            # ... and sort them
            keys.sort()
            # Now iterate through these Key values ...
            for x in keys:
                # See if the keyword is in the list of defined keywords, which it should be.
                if (self.codingObjects[x]['keywordGroup'], self.codingObjects[x]['keyword']) in tmpKeywordList:
                    # If this keyword isn't already part of the list of known keywords ...
                    if not((self.codingObjects[x]['keywordGroup'], self.codingObjects[x]['keyword']) in tmpCodeList):
                        # ... add it to the list of keywords used in this Snapshot
                        tmpCodeList.append((self.codingObjects[x]['keywordGroup'], self.codingObjects[x]['keyword']))
                    # If we're using Unicode (and we always are) ...
                    if 'unicode' in wx.PlatformInfo:
                        # ... encode the Keyword Group and Keyword
                        keywordGroup = self.codingObjects[x]['keywordGroup'].encode(TransanaGlobal.encoding)
                        keyword = self.codingObjects[x]['keyword'].encode(TransanaGlobal.encoding)
                    # Assemble the data values for the Insert query
                    values = (self.number, keywordGroup, keyword,
                              self.codingObjects[x]['x1'], self.codingObjects[x]['y1'], self.codingObjects[x]['x2'], self.codingObjects[x]['y2'])
                    # Encode the Visible property as '0' or '1' for the database
                    if self.codingObjects[x]['visible']:
                        values += ('1',)
                    else:
                        values += ('0',)
                    # Insert the data into the database
                    c.execute(query, values)
                # If the keyword isn't in the list of existing keywords, some other user must have changed it.  Inform the user.
                else:
                    # if the prompt isn't blank ...
                    if prompt != '':
                        # ... add a couple of line breaks to it
                        prompt += u'\n\n'
                    # Add the current keyword to the error prompt
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt += unicode(_('Keyword "%s : %s" cannot be added to Snapshot "%s".\nAnother user must have edited the keyword while you were adding it.'), 'utf8') % (self.codingObjects[x]['keywordGroup'], self.codingObjects[x]['keyword'], self.id)

            # If no error prompt is defines yet ...
            if prompt == '':
                # ... iterate through the Keyword Styles ....
                for (keywordGroup, keyword) in self.keywordStyles.keys():
                    # If the keyword is in the Code List ...
                    if (keywordGroup, keyword) in tmpCodeList:
                        # ... and the keyword is ALSO in the Temporary Keyword List ...
                        if (keywordGroup, keyword) in tmpKeywordList:
                            # Define the query for inserting the Keyword Style into the database
                            query = """ INSERT INTO SnapshotKeywordStyles2
                                          (SnapshotNum, KeywordGroup, Keyword, DrawMode, LineColorName, LineColorDef, LineWidth, LineStyle)
                                          VALUES
                                          (%s, %s, %s, %s, %s, %s, %s, %s) """
                            # Get the styles for this Keyword ...
                            row = self.keywordStyles[(keywordGroup, keyword)]
                            # Encode the style values that aren't part of "row" ...
                            if 'unicode' in wx.PlatformInfo:
                                keywordGroup = keywordGroup.encode(TransanaGlobal.encoding)
                                keyword = keyword.encode(TransanaGlobal.encoding)
                                lineColorName = row['lineColorName'].encode(TransanaGlobal.encoding)
                            # Gather the data values for the query, encoding values from "row" while we're at it.
                            values = (self.number, keywordGroup, keyword,
                                      row['drawMode'].encode('utf8'),
                                      lineColorName,
                                      row['lineColorDef'].encode('utf8'),
                                      row['lineWidth'].encode('utf8'),
                                      row['lineStyle'].encode('utf8'))
                            # Adjust the query for sqlite if needed
                            query = DBInterface.FixQuery(query)
                            # Execute the query
                            c.execute(query, values)
                        # If the keyword is NOT part of the Temporary Keyword List ...
                        else:
                            # if the prompt isn't blank ...
                            if prompt != '':
                                # ... add a couple of line breaks to it
                                prompt += u'\n\n'
                            # Add the current keyword to the error prompt
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt += unicode(_('Keyword "%s : %s" cannot be added to Snapshot "%s".\nAnother user must have edited the keyword while you were adding it.'), 'utf8') % (kws.keywordGroup, kws.keyword, self.id)

                    # If a Keyword Style does not appear in the Coding List ...
                    else:
                        # ... remove it from the Keyword Style List
                        del self.keywordStyles[(keywordGroup, keyword)]

            # If there is an error prompt ...
            if prompt != '':
                # If the Episode Number was changed ...
                if numberChanged:
                    # ... change it back to zero!!
                    self.number = 0
                # If we're using Transactions ...
                if use_transactions:
                    # Undo the database save transaction
                    c.execute('ROLLBACK')
                # Close the Database Cursor
                c.close()
                # Complete the error prompt
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt += u'\n\n' + unicode(_('Please remove, and possibly replace, this keyword.'), 'utf8')
                # ... raise a SaveError exception using the error prompt
                raise SaveError, prompt
            else:
                # If we're using Transactions ...
                if use_transactions:
                    # ... Commit the database transaction
                    c.execute('COMMIT')
                # Close the Database Cursor
                c.close()

    def db_delete(self, use_transactions=True):
        """ Delete this object record from the database.  Parameter indicates if we should use DB Transactions """
        result = 1
        try:
            # Initialize delete operation, begin transaction if necessary
            (db, c) = self._db_start_delete(use_transactions)

            # Detect, Load, and Delete all Snapshot Notes.
            notes = self.get_note_nums()
            for note_num in notes:
                note = Note.Note(note_num)
                result = result and note.db_delete(0)
                del note
            del notes

            # Delete all related references in the ClipKeywords table as well as the Snapshot Keywords and SnapshotKeyword
            # Styles tables.
            if result:
                # Delete Clip Keywords
                DBInterface.delete_all_keywords_for_a_group(0, 0, 0, 0, self.number)
                # Create the query to delete Snapshot Keywords
                query = "DELETE FROM SnapshotKeywords2 WHERE SnapshotNum = %s"
                # Adjust the query for sqlite if needed
                query = DBInterface.FixQuery(query)
                # Execute the query
                c.execute(query, (self.number, ))
                # Create the query to delete Snapshot Keyword Styles
                query = "DELETE FROM SnapshotKeywordStyles2 WHERE SnapshotNum = %s"
                # Adjust the query for sqlite if needed
                query = DBInterface.FixQuery(query)
                # Execute the query
                c.execute(query, (self.number, ))

            # Delete the actual Snapshot record.
            self._db_do_delete(use_transactions, c, result)

            # Cleanup
            c.close()
            self.clear()
        except RecordLockedError, e:
            # if a sub-record is locked, we may need to unlock the Snapshot record (after rolling back the Transaction)
            if self.isLocked:
                # c (the database cursor) only exists if the record lock was obtained!
                # We must roll back the transaction before we unlock the record.
                c.execute("ROLLBACK")
                c.close()
                self.unlock_record()
            raise e
        except:
            raise
        return result

    def lock_record(self):
        """ Override the DataObject Lock Method """
        # If we're using the multi-user version of Transana, we need to ...
        if not TransanaConstants.singleUserVersion:
            # ... confirm that the Snapshot has not been altered by another user since it was loaded.
            # To do this, first pull the LastSaveTime for this record from the database.
            query = """
                      SELECT LastSaveTime FROM Snapshots2
                      WHERE SnapshotNum = %s
                    """
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Get a database cursor
            tempDBCursor = DBInterface.get_db().cursor()
            # Execute the query
            tempDBCursor.execute(query, (self.number, ))
            # Get the results from the query
            data = tempDBCursor.fetchall()
            # If one record is returned ...
            if len(data) == 1:
                # ... get tehe last save time
                newLastSaveTime = data[0][0]
            # Otherwise ...
            else:
                # ... raise an exception
                raise RecordNotFoundError, (self.id, len(data))
            # Close the database cursor
            tempDBCursor.close()
            # If the object has a different LastSaveTime ...
            if newLastSaveTime != self.lastsavetime:
                # ... it's been edited elsewhere, so we need to re-load it!
                self.db_load(self.number)
        
        # ... lock the Transcript Record
        DataObject.DataObject.lock_record(self)

    def clear_keywords(self):
        """Clear the keyword list."""
        self._kwlist = []
        
    def refresh_keywords(self):
        """Clear the keyword list and refresh it from the database."""
        self._kwlist = []
        kwpairs = DBInterface.list_of_keywords(Snapshot=self.number)
        for data in kwpairs:
            tempClipKeyword = ClipKeywordObject.ClipKeyword(data[0], data[1], snapshotNum=self.number)
            self._kwlist.append(tempClipKeyword)
        
    def add_keyword(self, kwg, kw):
        """ Add a keyword to the keyword list.  By default, it is NOT a keyword example. """
        # We need to check to see if the keyword is already in the keyword list
        keywordFound = False
        # Iterate through the list
        for snapshotKeyword in self._kwlist:
            # If we find a match, set the flag and quit looking.
            if (snapshotKeyword.keywordGroup == kwg) and (snapshotKeyword.keyword == kw):
                keywordFound = True
                break

        # If the keyword is not found, add it.  (If it's already there, we don't need to do anything!)
        if not keywordFound:
            # Create an appropriate ClipKeyword Object
            tempClipKeyword = ClipKeywordObject.ClipKeyword(kwg, kw, snapshotNum=self.number)
            # Add it to the Keyword List
            self._kwlist.append(tempClipKeyword)

    def remove_keyword(self, kwg, kw):
        """Remove a keyword from the keyword list.  The value returned by this function can be:
             0  Keyword NOT deleted.  (probably overridden by the user)
             1  Keyword deleted, but it was NOT a Keyword Example
             2  Keyword deleted, and it WAS a Keyword Example. """
        # We need different return codes for failure, success of a Non-Example, and success of an Example.
        # If it's an example, we need to remove the Node on the Database Tree Tab

        # Let's assume the Delete will fail (or be refused by the user) until it actually happens.
        delResult = 0

        # We need to find the keyword in the keyword list
        # Iterate through the keyword list
        for index in range(len(self._kwlist)):

            # Look for the entry to be deleted
            if (self._kwlist[index].keywordGroup == kwg) and (self._kwlist[index].keyword == kw):
                # If the entry is found, delete it and stop looking
                del self._kwlist[index]
                # Signal that the delete was successful and was NOT an Example.
                delResult = 1
                # Once the entry has been found, stop looking for it
                break
            
        # Signal whether the delete was successful
        return delResult

    def has_keyword(self, kwg, kw):
        """ Determines if the Snapshot has a given keyword assigned """
        # Assume the result will be false
        res = False
        # Iterate through the keyword list
        for keyword in self.keyword_list:
            # See if the text passed in matches the strings in the keyword objects in the keyword list
            if (kwg == keyword.keywordGroup) and (kw == keyword.keyword):
                # If so, signal that it HAS been found
                res = True
                # If found, we don't need to look any more!
                break
        # Return the results
        return res

    def GetNodeData(self, includeSnapshot=True):
        """ Returns the Node Data list (list of parent collections) needed for Database Tree Manipulation """
        # Load the Snapshot's collection
        tempCollection = Collection.Collection(self.collection_num)
        # If we're including the Snapshot in the Node Data ...
        if includeSnapshot:
            # ... get the Collection's node data and tack the Snapshot's ID on the end.
            return tempCollection.GetNodeData() + (self.id,)
        # If we're NOT including the Clip in the Node data ...
        else:
            # ... just return the Collection's Node Data.
            return tempCollection.GetNodeData()

    def GetNodeString(self, includeSnapshot=True):
        """ Returns a string that delineates the full nested collection structure for the present Snapshot.
            if includeSnapshot=False, only the Collection path will be returned. """
        # Load the Clip's Collection
        tempCollection = Collection.Collection(self.collection_num)
        # If we're including the Snapshot in the Node String ...
        if includeSnapshot:
            # ... get the Collection's node string and tack the Snapshot's ID on the end.
            return tempCollection.GetNodeString() + ' > ' + self.id
        # If we're NOT including the Clip in the Node String ...
        else:
            # ... just return the Collection's Node String.
            return tempCollection.GetNodeString()
    
# Private methods    

    def _load_row(self, r):
        self.number = r['SnapshotNum']
        self.id = r['SnapshotID']
        self.collection_num = r['CollectNum']
        self.image_filename = r['ImageFile']
        self.image_scale = r['ImageScale']
        self.image_coords = (r['ImageCoordsX'], r['ImageCoordsY'])
        self.image_size = (r['ImageSizeW'], r['ImageSizeH'])
        self.episode_num = r['EpisodeNum']
        self.transcript_num = r['TranscriptNum']
        self.episode_start = r['SnapshotTimeCode']
        self.episode_duration = r['SnapshotDuration']
        self.comment = r['SnapshotComment']
        self.sort_order = r['SortOrder']
        self.lastsavetime = r['LastSaveTime']

        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.image_filename = DBInterface.ProcessDBDataForUTF8Encoding(self.image_filename)
            self.comment = DBInterface.ProcessDBDataForUTF8Encoding(self.comment)

        # Remember whether the ImageFile uses the VideoRoot Folder or not.
        # Detection of the use of the Video Root Path is platform-dependent.
        if wx.Platform == "__WXMSW__":
            # On Windows, check for a colon in the position, which signals the presence or absence of a drive letter
            self.useVideoRoot = (self.image_filename[1] != ':') and (self.image_filename[:2] != '\\\\')
        else:
            # On Mac OS-X and *nix, check for a slash in the first position for the root folder designation
            self.useVideoRoot = (self.image_filename[0] != '/')
        # If we are using the Video Root Path, add it to the Filename
        if self.useVideoRoot:
            self.image_filename = TransanaGlobal.configData.videoPath.replace('\\', '/') + self.image_filename

    def _sync_snapshot(self):
        """Synchronize the Snapshot's Collection, Episode, and Transcript properties, if needed."""
        # If there is a Collection Number ...
        if self.collection_num > 0:
            # ... load the Collection
            tempCollection = Collection.Collection(self.collection_num)
            # ... and grab the Collection ID
            self.collection_id = tempCollection.id
        # If there is an Episode Number ...
        if self.episode_num > 0:
            # Start Exception Handling
            try:
                # Load the Episode
                tempEpisode = Episode.Episode(self.episode_num)
                # Get the Episode ID and the Library Information
                self.episode_id = tempEpisode.id
                self.series_num = tempEpisode.series_num
                self.series_id = tempEpisode.series_id
            # If there is a RecordNotFound exception ...
            except RecordNotFoundError:
                # ... and if we're NOT deleting this Snapshot ...
                if not self.suppressEpisodeError:
                    # ... create the Episode Not Found error message
                    msg = unicode(_('Context Episode for Snapshot "%s" cannot be loaded.'), 'utf8')
                    # Display the Error Message
                    errordlg = Dialogs.ErrorDialog(None, msg % self.id)
                    errordlg.ShowModal()
                    errordlg.Destroy()

##                print "Snapshot._sync_snapshot(): Context Episode cannot be loaded error", self.episode_num
##                import sys
##                print sys.exc_info()[0]
##                print sys.exc_info()[1]
##                import traceback
##                traceback.print_exc(file=sys.stdout)
##                print

                # Eliminate the Episode and Transcript information
                self.episode_num = 0
                self.episode_id = ''
                self.transcript_num = 0
                self.transcript_id = ''
                self.episode_start = 0
                self.episode_duration = 0
        # If there is no Episode Number ...
        else:
            # ... then there's no Episode ID ...
            self.episode_id = ''
            # ... and there's no Library number or ID
            self.series_num = 0
            self.series_id = ''
            # There can't be Transcript information either
#            self.transcript_num = 0
#            self.transcript_id = ''
            # There can't be start and duration information either
            self.episode_start = 0
            self.episode_duration = 0

        # If there's an Episode AND a Transcript (avoids problems if Episode cause exception) ...
        if (self.episode_num > 0) and (self.transcript_num > 0):
            # Start Exception Handling
            try:
                # Try to load the Transcript.  (Don't need text)
                tempTranscript = Transcript.Transcript(self.transcript_num, skipText=True)
                # If loaded, get the ID
                self.transcript_id = tempTranscript.id
            # If the Transcript cannot be loaded ...
            except RecordNotFoundError:
                # ... prepare the Transcript Not Found error message
                msg = unicode(_('Context Transcript for Snapshot "%s" cannot be loaded.'), 'utf8')
                # Display the Error Message
                errordlg = Dialogs.ErrorDialog(None, msg % self.id)
                errordlg.ShowModal()
                errordlg.Destroy()
                
                # Clear the Transcript Number
                self.transcript_num = 0

    def _get_col_num(self):
        return self._col_num
    def _set_col_num(self, num):
        self._col_num = num
    def _del_col_num(self):
        self._col_num = 0

    def _get_fname(self):
        return self._fname.replace('/', os.sep)
    def _set_fname(self, fname):
        self._fname = fname.replace('\\', '/')
    def _del_fname(self):
        self._fname = ""

    def _get_scale(self):
        return self._scale
    def _set_scale(self, scale):
        self._scale = scale
    def _del_scale(self):
        self._scale = 0.0

    def _get_coords(self):
        return self._coords
    def _set_coords(self, coords):
        self._coords = coords
    def _del_coords(self):
        self._coords = (0.0, 0.0)

    def _get_size(self):
        return self._size
    def _set_size(self, size):
        self._size = size
    def _del_size(self):
        self._size = (0, 0)

    def _get_ep_num(self):
        return self._ep_num
    def _set_ep_num(self, num):
        self._ep_num = num
        self._sync_snapshot()
    def _del_ep_num(self):
        self._ep_num = 0

    def _get_tr_num(self):
        return self._tr_num
    def _set_tr_num(self, num):
        self._tr_num = num
        self._sync_snapshot()
    def _del_tr_num(self):
        self._tr_num = 0

    def _get_start(self):
        return self._start
    def _set_start(self, start):
        self._start = start
    def _del_start(self):
        self._start = 0.0

    def _get_duration(self):
        return self._duration
    def _set_duration(self, duration):
        self._duration = duration
    def _del_duration(self):
        self._duration = 0.0

    def _get_sort_order(self):
        return self._sort_order
    def _set_sort_order(self, so):
        self._sort_order = so
    def _del_sort_order(self):
        self._sort_order = 0

    def _get_kwlist(self):
        return self._kwlist
    def _set_kwlist(self, kwlist):
        self._kwlist = kwlist
    def _del_kwlist(self):
        self._kwlist = []

    def _get_codingObjects(self):
        return self._codingObjects
    def _set_codingObjects(self, codingObjects):
        self._codingObjects = codingObjects
    def _del_codingObjects(self):
        self._codingObjects = {}

    def _get_keywordSyles(self):
        return self._keywordSyles
    def _set_keywordSyles(self, keywordSyles):
        self._keywordSyles = keywordSyles
    def _del_keywordSyles(self):
        self._keywordSyles = {}

    # Implementation for LastSaveTime Property
    def _get_lastsavetime(self):
        return self._lastsavetime
    def _set_lastsavetime(self, lst):
        self._lastsavetime = lst
    def _del_lastsavetime(self):
        self._lastsavetime = None

# Public properties
    collection_num = property(_get_col_num, _set_col_num, _del_col_num,
                        """Collection number to which the clip belongs.""")
    image_filename = property(_get_fname, _set_fname, _del_fname,
                        """The name (including path) of the media file.""")
    image_scale = property(_get_scale, _set_scale, _del_scale, """Scaling Factor for the Image""")
    image_coords = property(_get_coords, _set_coords, _del_coords, """Coordinates""")
    image_size = property(_get_size, _set_size, _del_size, """Size""")
    episode_num = property(_get_ep_num, _set_ep_num, _del_ep_num,
                        """Number of episode from which this Snapshot was taken or to which it was assigned.""")
    transcript_num = property(_get_tr_num, _set_tr_num, _del_tr_num,
                        """Number of episode transcript from which this Snapshot was taken or to which it was assigned.""")
    episode_start = property(_get_start, _set_start, _del_start, """Episode Context Position""")
    episode_duration = property(_get_duration, _set_duration, _del_duration, """Episode Context Duration""")
    sort_order = property(_get_sort_order, _set_sort_order, _del_sort_order,
                        """Sort Order position within the parent Collection.""")
    keyword_list = property(_get_kwlist, _set_kwlist, _del_kwlist,
                        """The list of keywords that have been applied to the Clip.""")
    codingObjects = property(_get_codingObjects, _set_codingObjects, _del_codingObjects,
                           """ The coding objects drawn on the Snapshot """)
    keywordStyles = property(_get_keywordSyles, _set_keywordSyles, _del_keywordSyles,
                           """ The style for coding objects drawn on the Snapshot """)
    lastsavetime = property(_get_lastsavetime, _set_lastsavetime, _del_lastsavetime,
                        """The timestamp of the last save (MU only).""")
