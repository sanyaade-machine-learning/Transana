# Copyright (C) 2003 - 2010 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Episode class as part of the Data Objects."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

DEBUG = False
if DEBUG:
    print "Episode DEBUG is ON!"

# import wxPython
import wx
# import MySQLdb (for Date Formatting)
import MySQLdb
# Import Python's os module
import os
# import Python's sys module
import sys
# import Python's types module
import types
# import Transana's Clip Keyword Object
import ClipKeywordObject
# import Transana's base Data Object
import DataObject
# import Transana's Database Interface
import DBInterface
# import Transana's Note Object
import Note
# import Transana's Exceptions
from TransanaExceptions import *
# import the Transana Constants
import TransanaConstants
# import Transana's Globals
import TransanaGlobal
# import Transana's Transcript Object
import Transcript


class Episode(DataObject.DataObject):
    """This class defines the structure for a episode object.  A episode object
    describes a video (or other media) file."""

    def __init__(self, num=None, series=None, episode=None):
        """Initialize an Episode object."""
        DataObject.DataObject.__init__(self)
        # By default, use the Video Root folder if one has been defined
        self.useVideoRoot = (TransanaGlobal.configData.videoPath != '')
        
        if num != None:
            self.db_load_by_num(num)
        elif (series != None and episode != None):
            self.db_load_by_name(series, episode)

        
# Public methods
    def __repr__(self):
        str = 'Episode Object:\n'
        str = str + "Number = %s\n" % self.number
        str = str + "id = %s\n" % self.id
        str = str + "comment = %s\n" % self.comment
        str = str + "media file = %s\n" % self.media_filename
        str = str + "Additional media file:\n"
        for addFile in self.additional_media_files:
            str += '  %s  %s %s %s\n' % (addFile['filename'], addFile['offset'], addFile['length'], addFile['audio'])
        str += "Offset = %s\n" % self.offset
        str = str + "Length = %s\n" % self.tape_length
        str = str + "Date = %s\n" % self.tape_date
        str = str + "Series ID = %s\n" % self.series_id
        str = str + "Series Num = %s\n\n" % self.series_num
        return str.encode('utf8')


    def db_load_by_name(self, series, episode):
        """Load a record by ID / Name."""
        # If we're in Unicode mode, we need to encode the parameters so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            series = series.encode(TransanaGlobal.encoding)
            episode = episode.encode(TransanaGlobal.encoding)
        # Get a database connection
        db = DBInterface.get_db()
        # Craft a query to get Episode data
        query = """SELECT * FROM Episodes2 a, Series2 b
            WHERE   EpisodeID = %s AND
                    a.SeriesNum = b.SeriesNum AND
                    b.SeriesID = %s
        """
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, (episode, series))
        # Get the number of rows returned
        n = c.rowcount
        # If we don't get exactly one result ...
        if (n != 1):
            # Close the cursor
            c.close()
            # Clear the current Episode object
            self.clear()
            # Raise an exception saying the data is not found
            raise RecordNotFoundError, (episode, n)
        # If we get exactly one result ...
        else:
            # Get the data from the cursor
            r = DBInterface.fetch_named(c)
            # Load the data into the Episode object
            self._load_row(r)
            # Load Additional Media Files, which aren't handled in the "old" code
            self.load_additional_vids()
            # Refresh the Keywords
            self.refresh_keywords()
        # Close the Database cursor
        c.close()

    def db_load_by_num(self, num):
        """Load a record by record number."""
        # Get a database connection
        db = DBInterface.get_db()
        # Craft a query to get Episode data
        query = """SELECT * FROM Episodes2 a, Series2 b
            WHERE   EpisodeNum = %s AND
                    a.SeriesNum = b.SeriesNum
        """
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, num)
        # Get the number of rows returned
        n = c.rowcount
        # If we don't get exactly one result ...
        if (n != 1):
            # Close the cursor
            c.close()
            # Clear the current Episode object
            self.clear()
            # Raise an exception saying the data is not found
            raise RecordNotFoundError, (num, n)
        # If we get exactly one result ...
        else:
            # Get the data from the cursor
            r = DBInterface.fetch_named(c)
            # Load the data into the Episode object
            self._load_row(r)
            # Load Additional Media Files, which aren't handled in the "old" code
            self.load_additional_vids()
            # Refresh the Keywords
            self.refresh_keywords()
        # Close the Database cursor
        c.close()

    def db_save(self):
        """Save the record to the database using Insert or Update as
        appropriate."""

        # Define and implement Demo Version limits
        if TransanaConstants.demoVersion and (self.number == 0):
            # Get a DB Cursor
            c = DBInterface.get_db().cursor()
            # Find out how many records exist
            c.execute('SELECT COUNT(EpisodeNum) from Episodes2')
            res = c.fetchone()
            c.close()
            # Define the maximum number of recors allowed
            maxEpisodes = TransanaConstants.maxEpisodes
            # Compare
            if res[0] >= maxEpisodes:
                # If the limit is exceeded, create and display the error using a SaveError exception
                prompt = _('The Transana Demonstration limits you to %d Episode records.\nPlease cancel the "Add Episode" dialog to continue.')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                raise SaveError, prompt % maxEpisodes
            
        # Sanity checks
        if self.id == "":
            raise SaveError, _("Episode ID is required.")
        elif self.series_num == 0:
            raise SaveError, _("This Episode is not associated properly with a Series.")
        elif self.media_filename == "":
            raise SaveError, _("Media Filename is required.")
        else:
            # videoPath probably has the OS.sep character, but we need the generic "/" character here.
            videoPath = TransanaGlobal.configData.videoPath

            # Determine if we are supposed to extract the Video Root Path from the Media Filename and extract it if appropriate
            if self.useVideoRoot and (videoPath == self.media_filename[:len(videoPath)]):
                tempMediaFilename = self.media_filename[len(videoPath):]
            else:
                tempMediaFilename = self.media_filename

            # Substitute the generic OS seperator "/" for the Windows "\".
            tempMediaFilename = tempMediaFilename.replace('\\', '/')
            # if we're not in Unicode mode ...
            if 'ansi' in wx.PlatformInfo:
                # ... we don't need to encode the string values, but we still need to copy them to our local variables.
                id = self.id
                comment = self.comment
                series_id = self.series_id
            # If we're in Unicode mode ...
            else:
                # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
                # variables for the data.  We don't want to change the underlying object values.  Also, this way,
                # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
                id = self.id.encode(TransanaGlobal.encoding)
                tempMediaFilename = tempMediaFilename.encode(TransanaGlobal.encoding)
                videoPath = videoPath.encode(TransanaGlobal.encoding)
                comment = self.comment.encode(TransanaGlobal.encoding)
                series_id = self.series_id.encode(TransanaGlobal.encoding)

        self._sync_series()

        fields = ("EpisodeID", "SeriesNum", "MediaFile", "EpLength", \
                        "TapingDate", "EpComment")
        
        values = (id, self.series_num, tempMediaFilename, \
                    self.tape_length, self.tape_date_db, comment)

        if (self._db_start_save() == 0):
            # Duplicate Episode IDs within a Series are not allowed.
            if DBInterface.record_match_count("Episodes2", \
                    ("EpisodeID", "SeriesNum"),
                    (id, self.series_num)) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('An Episode named "%s" already exists in Series "%s".\nPlease enter a different Episode ID.'), 'utf8')
                else:
                    prompt = _('An Episode named "%s" already exists in Series "%s".\nPlease enter a different Episode ID.')
                raise SaveError, prompt % (self.id, self .series_id)

            # insert the new record
            # Could be some issues with how TapingDate is stored (YYYY/MM/DD ?)
            query = "INSERT INTO Episodes2\n("
            for field in fields:
                query = "%s%s," % (query, field)
            query = query[:-1] + ')'
            query = "%s\nVALUES\n(" % query
            for value in values:
                query = "%s%%s," % query
            query = query[:-1] + ')'
        else:
            # check for dupes
            if DBInterface.record_match_count("Episodes2", \
                    ("EpisodeID", "SeriesNum", "!EpisodeNum"),
                    (id, self.series_num, self.number) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('An Episode named "%s" already exists in Series "%s".\nPlease enter a different Episode ID.'), 'utf8')
                else:
                    prompt = _('An Episode named "%s" already exists in Series "%s".\nPlease enter a different Episode ID.')
                raise SaveError, prompt % (self.id, self.series_id)
            
            # OK to update the episode record
            query = """UPDATE Episodes2
                SET EpisodeID = %s,
                    SeriesNum = %s,
                    MediaFile = %s,
                    EpLength = %s,
                    TapingDate = %s,
                    EpComment = %s
                WHERE EpisodeNum = %s
            """
            values = values + (self.number,)

        c = DBInterface.get_db().cursor()
        c.execute(query, values)

        if self.number == 0:
            # If we are dealing with a brand new Episode, it does not yet know its
            # record number.  It HAS a record number, but it is not known yet.
            # The following query should produce the correct record number.
            query = """
                      SELECT EpisodeNum FROM Episodes2
                      WHERE EpisodeID = %s AND
                            SeriesNum = %s
                    """
            tempDBCursor = DBInterface.get_db().cursor()
            tempDBCursor.execute(query, (id, self.series_num))
            if tempDBCursor.rowcount == 1:
                self.number = tempDBCursor.fetchone()[0]
            else:
                raise RecordNotFoundError, (self.id, tempDBCursor.rowcount)
            tempDBCursor.close()
        else:
            # If we are dealing with an existing Episode, delete all the Keywords
            # in anticipation of putting them all back later
            DBInterface.delete_all_keywords_for_a_group(self.number, 0)

        # To save the additional video file names, we must first delete them from the database!
        # Craft a query to remove all existing Additonal Videos
        query = "DELETE FROM AdditionalVids2 WHERE EpisodeNum = %d" % self.number
        # Execute the query
        c.execute(query)
        # Define the query to insert the additional media files into the databse
        query = "INSERT INTO AdditionalVids2 (EpisodeNum, ClipNum, MediaFile, VidLength, Offset, Audio) VALUES (%s, %s, %s, %s, %s, %s)"
        # For each additional media file ...
        for vid in self.additional_media_files:
            # Encode the filename
            tmpFilename = vid['filename'].encode(TransanaGlobal.encoding)
            # Determine if we are supposed to extract the Video Root Path from the Media Filename and extract it if appropriate
            if self.useVideoRoot and (videoPath == tmpFilename[:len(videoPath)]):
                tmpFilename = tmpFilename[len(videoPath):]
            # Substitute the generic OS seperator "/" for the Windows "\".
            tmpFilename = tmpFilename.replace('\\', '/')
            # Get the data for each insert query
            data = (self.number, 0, tmpFilename, vid['length'], vid['offset'], vid['audio'])
            # Execute the query
            c.execute(query, data)

        # Add the Episode keywords back
        for kws in self._kwlist:
            DBInterface.insert_clip_keyword(self.number, 0, kws.keywordGroup, kws.keyword, kws.example)

        c.close()
            
            
    def db_delete(self, use_transactions=1):
        """Delete this object record from the database."""
        result = 1
        try:
            # Initialize delete operation, begin transaction if necessary
            (db, c) = self._db_start_delete(use_transactions)

            # Delete all Episode-based Filter Configurations
            #   Delete Keyword Map records
            DBInterface.delete_filter_records(1, self.number)
            #   Delete Keyword Visualization records
            DBInterface.delete_filter_records(2, self.number)
            #   Delete Episode Clip Data Export records
            DBInterface.delete_filter_records(3, self.number)
            #   Delete Episode Clip Data Coder Reliability Export records (Kathleen Liston's code)
            DBInterface.delete_filter_records(8, self.number)
            #   Delete Episode Report records
            DBInterface.delete_filter_records(11, self.number)

            # Detect, Load, and Delete all Clip Notes.
            notes = self.get_note_nums()
            for note_num in notes:
                note = Note.Note(note_num)
                result = result and note.db_delete(0)
                del note
            del notes

            for (transcriptNo, transcriptID, transcriptEpisodeNo) in DBInterface.list_transcripts(self.series_id, self.id):
                trans = Transcript.Transcript(transcriptNo)
                # if transcript delete fails, rollback clip delete
                result = result and trans.db_delete(0)
                del trans

            # Delete all related references in the ClipKeywords table
            if result:
                DBInterface.delete_all_keywords_for_a_group(self.number, 0)

            if result:
                # Craft a query to remove all existing Additonal Videos
                query = "DELETE FROM AdditionalVids2 WHERE EpisodeNum = %d" % self.number
                # Execute the query
                c.execute(query)

            # Delete the actual record.
            self._db_do_delete(use_transactions, c, result)

            # Cleanup
            c.close()
            self.clear()
        except RecordLockedError, e:
            # if a sub-record is locked, we may need to unlock the Episode record (after rolling back the Transaction)
            if self.isLocked:
                # c (the database cursor) only exists if the record lock was obtained!
                # We must roll back the transaction before we unlock the record.
                c.execute("ROLLBACK")
                c.close()
                self.unlock_record()
            raise e
        # Handle the DeleteError Exception
        except DeleteError, e:
            # If the record is locked ...
            if self.isLocked:
                # ... unlock it ...
                self.unlock_record()
            # ... then pass the exception on.
            raise e
        except:
            raise

        return result

    def clear_keywords(self):
        """Clear the keyword list."""
        self._kwlist = []
        
    def refresh_keywords(self):
        """Clear the keyword list and refresh it from the database."""
        self._kwlist = []
        kwpairs = DBInterface.list_of_keywords(Episode=self.number)
        for data in kwpairs:
            tempClipKeyword = ClipKeywordObject.ClipKeyword(data[0], data[1], episodeNum=self.number, example=data[2])
            self._kwlist.append(tempClipKeyword)
        
    def add_keyword(self, kwg, kw):
        """Add a keyword to the keyword list."""
        # We need to check to see if the keyword is already in the keyword list
        keywordFound = False
        # Iterate through the list
        for episodeKeyword in self._kwlist:
            # If we find a match, set the flag and quit looking.
            if (episodeKeyword.keywordGroup == kwg) and (episodeKeyword.keyword == kw):
                keywordFound = True
                break

        # If the keyword is not found, add it.  (If it's already there, we don't need to do anything!)
        if not keywordFound:
            # Create an appropriate ClipKeyword Object
            tempClipKeyword = ClipKeywordObject.ClipKeyword(kwg, kw, episodeNum=self.number)
            # Add it to the Keyword List
            self._kwlist.append(tempClipKeyword)

    def remove_keyword(self, kwg, kw):
        """Remove a keyword from the keyword list."""
        # Let's assume the Delete will fail (or be refused by the user) until it actually happens
        delResult = False

        # We need to find the keyword in the keyword list
        # Iterate through the keyword list 
        for index in range(len(self._kwlist)):
            # Look for the entry to be deleted
            if (self._kwlist[index].keywordGroup == kwg) and (self._kwlist[index].keyword == kw):
                # If the entry is found, delete it and stop looking
                del self._kwlist[index]
                delResult = True
                break
        return delResult

    def has_keyword(self, kwg, kw):
        """ Determines if the Episode has a given keyword assigned """
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

    def load_additional_vids(self):
        """Load additional media file names from the database."""
        # Get a database connection
        db = DBInterface.get_db()
        # Create a database cursor
        c = db.cursor()
        # Define the DB Query
        query = "SELECT MediaFile, VidLength, Offset, Audio FROM AdditionalVids2 WHERE EpisodeNum = %s ORDER BY AddVidNum"
        # Execute the query
        c.execute(query, self.number)
        # For each video in the query results ...
        for (vidFilename, vidLength, vidOffset, audio) in c.fetchall():
            # Detection of the use of the Video Root Path is platform-dependent and must be done for EACH filename!
            if wx.Platform == "__WXMSW__":
                # On Windows, check for a colon in the position, which signals the presence or absence of a drive letter
                useVideoRoot = (vidFilename[1] != ':')
            else:
                # On Mac OS-X and *nix, check for a slash in the first position for the root folder designation
                useVideoRoot = (vidFilename[0] != '/')
            # If we are using the Video Root Path, add it to the Filename
            if useVideoRoot:
                video = TransanaGlobal.configData.videoPath.replace('\\', '/') + DBInterface.ProcessDBDataForUTF8Encoding(vidFilename)
            else:
                video = DBInterface.ProcessDBDataForUTF8Encoding(vidFilename)
            # Add the video to the additional media files list
            self.additional_media_files = {'filename' : video,
                                           'length'   : vidLength,
                                           'offset'   : vidOffset,
                                           'audio'    : audio}
            # If the video offset is less than 0 and is smaller than the current smallest (most negative) offset ...
            if (vidOffset < 0) and (vidOffset < -self.offset):
                # ... then use this video offset as the global offset
                self.offset = abs(vidOffset)
        # Close the database cursor
        c.close()

    def remove_an_additional_vid(self, indx):
        """ remove ONE additional media file from the list of additional media files """
        del(self._additional_media[indx])

    def tape_length_str(self):
        """Return a string representation (HH:MM:SS) of tape length."""
        secs = int(round(self._tl / 1000.0))    # total # seconds
        hours = secs / 3600                     # num full hours
        mins = (secs / 60) - hours*60
        secs = secs - mins*60 - hours*3600
        return "%02d:%02d:%02d" % (hours, mins, secs)

    
# Private methods

    def _sync_series(self):
        """Synchronize the Series ID property to reflect the current state
        of the Series Number property."""
        from Series import Series
        s = Series(self.series_num)
        self.series_id = s.id

    def _load_row(self, r):
        self.number = r['EpisodeNum']
        self.id = r['EpisodeID']
        self.comment = r['EpComment']
        self.media_filename = r['MediaFile']
        self.tape_length = r['EpLength']
        self.tape_date = r['TapingDate']
        
        # These come from the Series record
        self.series_id = r['SeriesID']
        self.series_num = r['SeriesNum']
        
        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.comment = DBInterface.ProcessDBDataForUTF8Encoding(self.comment)
            self.media_filename = DBInterface.ProcessDBDataForUTF8Encoding(self.media_filename)
            self.series_id = DBInterface.ProcessDBDataForUTF8Encoding(self.series_id)

        # Remember whether the MediaFile uses the VideoRoot Folder or not.
        # Detection of the use of the Video Root Path is platform-dependent.
        if wx.Platform == "__WXMSW__":
            # On Windows, check for a colon in the position, which signals the presence or absence of a drive letter
            self.useVideoRoot = (self.media_filename[1] != ':')
        else:
            # On Mac OS-X and *nix, check for a slash in the first position for the root folder designation
            self.useVideoRoot = (self.media_filename[0] != '/')
        # If we are using the Video Root Path, add it to the Filename
        if self.useVideoRoot:
            self.media_filename = TransanaGlobal.configData.videoPath.replace('\\', '/') + self.media_filename
            

    def _get_ser_num(self):
        return self._ser_num
    def _set_ser_num(self, num):
        self._ser_num = num
    def _del_ser_num(self):
        self._ser_num = 0

    def _get_ser_id(self):
        return self._ser_id
    def _set_ser_id(self, id):
        self._ser_id = id
    def _del_ser_id(self):
        self._ser_id = ""

    # This is a DateTime object
    def _get_td(self):
        # Although _dc_date is a MySQLdb.DateTime object, Date will be presented as a 'MM/DD/YYYY' string!
        if self._td == None:
            return ''
        else:
            return "%02d/%02d/%s" % (self._td.month, self._td.day, self._td.year)
    def _set_td(self, td):
        # Although _dc_date is a MySQLdb.DateTime object, Date will be presented for ingestion as a 'MM/DD/YYYY' string!
        # A None Object or a blank string should set _dc_date to None
        if (td == '') or (td == None) or (td == u'  /  /    '):
            self._td = None
        else:
            try:
                # A 'MM/DD/YYYY' string needs to be parsed.
                if isinstance(td, types.StringTypes):
                    (month, day, year) = td.split('/')
                    # The wxMaskedTextCtrl returns '  /  /    ' if left empty.
                    # We need to convert the Month into an Integer, setting it to 1 if empty
                    if month.strip() == '':
                        month = 1
                    else:
                        month = int(month)
                    # We need to convert the Day into an Integer, setting it to 1 if empty
                    if day.strip() == '':
                        day = 1
                    else:
                        day = int(day)
                    # We need to convert the Year into an Integer, setting it to 0 if empty
                    if year.strip() == '':
                        year = 0
                    else:
                        year = int(year)
                    # if all values were empty or bogus, set _dc_date to None, but if we have SOMETHING to save,
                    # save it.
                    if (year != 0) or (month != 0) or (day != 0):
                        self._td = MySQLdb.Date(int(year), int(month), int(day))
                    else:
                        self._td = None
                # If we get a non-string argument, just save it.  
                else:
                    self._td = td
            except:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Date Format Error.  The Date must be in the format MM/DD/YYYY.\n%s'), 'utf8')
                else:
                    prompt = _('Date Format Error.  The Date must be in the format MM/DD/YYYY.\n%s')
                raise SaveError, prompt % sys.exc_info()[1]
    def _del_td(self):
        self._td = None
        
    def _get_td_db(self):
        """ Read-only version of the Date in a MySQL Friendly format """
        # Although _td is a MySQLdb.DateTime object, the database needs the Date to be
        # presented as a 'YYYY-MM-DD' string!
        if self._td == None:
            return None
        else:
            return "%04d/%02d/%02d" % (self._td.year, self._td.month, self._td.day)

    def _get_offset(self):
        return self._offset
    def _set_offset(self, offset):
        self._offset = offset
    def _del_offset(self):
        self._offset = 0

    def _get_tl(self):
        return self._tl
    def _set_tl(self, tl):
        self._tl = tl
    def _del_tl(self):
        self._tl = 0

    def _get_fname(self):
        return self._fname.replace('/', os.sep)
    def _set_fname(self, fname):
        self._fname = fname.replace('\\', '/')
    def _del_fname(self):
        self._fname = ""

    def _get_additional_media(self):
        temp_additional_media = []
        for vid in self._additional_media:
            vid['filename'] = vid['filename'].replace('/', os.sep)
            temp_additional_media.append(vid)
        return temp_additional_media
    def _set_additional_media(self, vidDict):
        # If we receive a Dictionary Object ...
        if isinstance(vidDict, dict):
            # ... replace the backslashes in the filename item
            vidDict['filename'] = vidDict['filename'].replace('\\', '/')
            # Add the dictionary object to the media files list
            self._additional_media.append(vidDict)
        # If we receive something else (a list or a tuple, most likely) ...
        else:
            # ... for each element (which should be a dictionary object) ...
            for vid in vidDict:
                # ... replace the backslashes in the filename item
                vid['filename'] = vid['filename'].replace('\\', '/')
                # Add the dictionary object to the media files list
                self._additional_media.append(vid)

    def _del_additional_media(self):
        self._additional_media = []

    def _get_kwlist(self):
        return self._kwlist
    def _set_kwlist(self, kwlist):
        self._kwlist = kwlist
    def _del_kwlist(self):
        self._kwlist = []

    

# Public properties
    series_num = property(_get_ser_num, _set_ser_num, _del_ser_num,
                        """The number of the series to which the episode belongs.""")
    series_id = property(_get_ser_id, _set_ser_id, _del_ser_id,
                        """The name of the series to which the episode belongs.""")
    media_filename = property(_get_fname, _set_fname, _del_fname,
                        """The name (including path) of the media file.""")
    additional_media_files = property(_get_additional_media, _set_additional_media, _del_additional_media,
                        """A list of additional media files (including path).""")
    offset = property(_get_offset, _set_offset, _del_offset,
                      """The offset indicates how much later than the earliest multiple video the FIRST video starts.""")
    tape_length = property(_get_tl, _set_tl, _del_tl,
                        """The length of the file in milliseconds.""")
    tape_date = property(_get_td, _set_td, _del_td,
                        """The date the media was recorded.""")
    tape_date_db = property(_get_td_db, None, None, 'tape_date in the format needed for the MySQL database')
    keyword_list = property(_get_kwlist, _set_kwlist, _del_kwlist,
                        """The list of keywords that have been applied to the Episode.""")
