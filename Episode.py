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

"""This module implements the Episode class as part of the Data Objects."""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx
from DataObject import DataObject
import DBInterface
import ClipKeywordObject
import Note
import Transcript
from TransanaExceptions import *
# import the Transana Constants
import TransanaConstants
# import Transana's Globals
import TransanaGlobal


class Episode(DataObject):
    """This class defines the structure for a episode object.  A episode object
    describes a video (or other media) file."""

    def __init__(self, num=None, series=None, episode=None):
        """Initialize an Episode object."""
        DataObject.__init__(self)
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
        str = str + "Length = %s\n" % self.tape_length
        str = str + "Date = %s\n" % self.tape_date
        str = str + "Series ID = %s\n" % self.series_id
        str = str + "Series Num = %s\n\n" % self.series_num
        return str


    def db_load_by_name(self, series, episode):
        """Load a record by ID / Name."""
        db = DBInterface.get_db()
        query = """SELECT * FROM Episodes2 a, Series2 b
            WHERE   EpisodeID = %s AND
                    a.SeriesNum = b.SeriesNum AND
                    b.SeriesID = %s
        """
        c = db.cursor()
        c.execute(query, (episode, series))
        n = c.rowcount
        if (n != 1):
            c.close()
            self.clear()
            raise RecordNotFoundError, (episode, n)
        else:
            r = DBInterface.fetch_named(c)
            self._load_row(r)
            self.refresh_keywords()
        
        c.close()
    
            
    def db_load_by_num(self, num):
        """Load a record by record number."""
        db = DBInterface.get_db()
        query = """SELECT * FROM Episodes2 a, Series2 b
            WHERE   EpisodeNum = %s AND
                    a.SeriesNum = b.SeriesNum
        """
        c = db.cursor()
        c.execute(query, num)
        n = c.rowcount
        if (n != 1):
            c.close()
            self.clear()
            raise RecordNotFoundError, (num, n)
        else:
            r = DBInterface.fetch_named(c)
            self._load_row(r)
            self.refresh_keywords()
        
        c.close()
        
    def db_save(self):
        """Save the record to the database using Insert or Update as
        appropriate."""

        # Sanity checks
        if self.id == "":
            raise SaveError, _("Episode ID is required.")
        elif self.series_num == 0:
            raise SaveError, _("This Episode is not associated properly with a Series.")
        elif self.media_filename == "":
            raise SaveError, _("Media Filename is required.")
        else:
            # Create a string of legal characters for the file names
            allowedChars = TransanaConstants.legalFilenameCharacters
            # check each character in the file name string
            for char in self.media_filename:
                # If the character is illegal ...
                if allowedChars.find(char) == -1:
                    msg = _('There is an unsupported character in the Media File Name.\n\n"%s" includes the "%s" character, which Transana does not allow at this time.\nPlease rename your folders and files so that they do not include characters that are not part of US English.\nWe apologize for this inconvenience.') % (self.media_filename, char)
                    raise SaveError, msg

        self._sync_series()

        fields = ("EpisodeID", "SeriesNum", "MediaFile", "EpLength", \
                        "TapingDate", "EpComment")
        
        # Determine if we are supposed to extract the Video Root Path from the Media Filename and extract it if appropriate
        if self.useVideoRoot and (TransanaGlobal.configData.videoPath == self.media_filename[:len(TransanaGlobal.configData.videoPath)]):
            tempMediaFilename = self.media_filename[len(TransanaGlobal.configData.videoPath):]
        else:
            tempMediaFilename = self.media_filename

        values = (self.id, self.series_num, tempMediaFilename, \
                    self.tape_length, self.tape_date, self.comment)

        if (self._db_start_save() == 0):
            # Duplicate Episode IDs within a Series are not allowed.
            if DBInterface.record_match_count("Episodes2", \
                    ("EpisodeID", "SeriesNum"),
                    (self.id, self.series_num)) > 0:
                raise SaveError, _('An Episode named "%s" already exists in Series "%s".\nPlease enter a different Episode ID.') % (self.id, self.series_id)

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
                    (self.id, self.series_num, self.number) ) > 0:
                raise SaveError, _('An Episode named "%s" already exists in Series "%s".\nPlease enter a different Episode ID.') % (self.id, self.series_id)
            
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

        # Delete all episode keywords
        if self.number > 0:
            DBInterface.delete_all_keywords_for_a_group(self.number, 0)

        # Add the Episode keywords back
        for kws in self._kwlist:
            DBInterface.insert_clip_keyword(self.number, 0, kws.keywordGroup, kws.keyword, kws.example)

        if (self.number == 0):
            # Load the auto-assigned new number record
            self.db_load_by_name(self.series_id, self.id)
            query = """
            UPDATE ClipKeywords2
                SET EpisodeNum = %s
                WHERE EpisodeNum = 0 AND
                        ClipNum = 0
            """
            c.execute(query, self.number)
            self.db_load_by_name(self.series_id, self.id)
            
            
        c.close()
            
            
    def db_delete(self, use_transactions=1):
        """Delete this object record from the database."""
        # Initialize delete operation, begin transaction if necessary
        (db, c) = self._db_start_delete(use_transactions)
        result = 1
        
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

        # Delete the actual record.
        self._db_do_delete(use_transactions, c, result)

        # Cleanup
        c.close()
        self.clear()

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
            self.media_filename = TransanaGlobal.configData.videoPath + self.media_filename
        self.tape_length = r['EpLength']
        self.tape_date = r['TapingDate']
        
        # These come from the Series record
        self.series_id = r['SeriesID']
        self.series_num = r['SeriesNum']
        

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
        return self._td
    def _set_td(self, td):
        self._td = td
    def _del_td(self):
        self._td = None

    def _get_tl(self):
        return self._tl
    def _set_tl(self, tl):
        self._tl = tl
    def _del_tl(self):
        self._tl = 0

    def _get_fname(self):
        return self._fname
    def _set_fname(self, fname):
        self._fname = fname
    def _del_fname(self):
        self._fname = ""


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
    tape_length = property(_get_tl, _set_tl, _del_tl,
                        """The length of the file in milliseconds.""")
    tape_date = property(_get_td, _set_td, _del_td,
                        """The date the media was recorded.""")
    keyword_list = property(_get_kwlist, _set_kwlist, _del_kwlist,
                        """The list of keywords that have been applied to the Episode.""")
