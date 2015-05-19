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

"""This module implements the Series class as part of the Data Objects."""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

# Based on code/ideas/logic from USeriesObject Delphi unit by DKW

DEBUG = False
if DEBUG:
    print "Series DEBUG is ON!"

import wx
from DataObject import DataObject
from Note import Note
from Episode import Episode
from TransanaExceptions import *
# import Transana's Globals
import TransanaGlobal
import DBInterface
import types
# import Transana Dialogs
import Dialogs

class Series(DataObject):
    """This class defines the structure for a series object.  A series object
    holds information about a group (e.g., a Series) of video files
    (e.g., Episodes)."""

    def __init__(self, id_or_num=None):
        """Initialize an Series object.  If a record ID number or Series ID
        is given, load it from the database."""
        DataObject.__init__(self)
        if type(id_or_num) in (int, long):
            self.db_load_by_num(id_or_num)
        elif isinstance(id_or_num, types.StringTypes):
            self.db_load_by_name(id_or_num)

    def __repr__(self):
        """ String Representation of a Series Object """
        str = 'Series Object:\n'
        str += 'number = %s\n' % self.number
        str += 'id = %s\n' % self.id
        str += 'comment = %s\n' % self.comment
        str += 'owner = %s\n' % self.owner
        str += 'keyword_group = %s\n\n' % self.keyword_group
        return str
        

# Public methods

    def db_load_by_name(self, name):
        """Load a record by ID / Name.  Raise a RecordNotFound exception
        if record is not found in database."""
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            name = name.encode(TransanaGlobal.encoding)
        db = DBInterface.get_db()
        query = """
        SELECT * FROM Series2
            WHERE SeriesID = %s
        """
        c = db.cursor()
        c.execute(query, name)
        n = c.rowcount
        if (n != 1):
            c.close()
            self.clear()
            raise RecordNotFoundError, (name, n)
        else:
            r = DBInterface.fetch_named(c)
            self._load_row(r)

        c.close()

    def db_load_by_num(self, num):
        """Load a record by record number."""
        db = DBInterface.get_db()
        query = """
        SELECT * FROM Series2
            WHERE SeriesNum = %s
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
        c.close()
        
    def db_save(self):
        """Save the record to the database using Insert or Update as
        appropriate."""
        # Sanity checks
        if self.id == "":
            raise SaveError, _("Series ID is required.")

        # If we're in Unicode mode, ...
        if 'unicode' in wx.PlatformInfo:
            # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
            # variables for the data.  We don't want to change the underlying object values.  Also, this way,
            # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
            id = self.id.encode(TransanaGlobal.encoding)
            comment = self.comment.encode(TransanaGlobal.encoding)
            owner = self.owner.encode(TransanaGlobal.encoding)
            keyword_group = self.keyword_group.encode(TransanaGlobal.encoding)
        else:
            # If we don't need to encode the string values, we still need to copy them to our local variables.
            id = self.id
            comment = self.comment
            owner = self.owner
            keyword_group = self.keyword_group
        
        values = (id, comment, owner, keyword_group)
        if (self._db_start_save() == 0):

            # duplicate Series IDs are not allowed
            if DBInterface.record_match_count("Series2", \
                            ("SeriesID",), \
                            (id,) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Series named "%s" already exists.\nPlease enter a different Series ID.'), 'utf8')
                else:
                    prompt = _('A Series named "%s" already exists.\nPlease enter a different Series ID.')
                raise SaveError, prompt % self.id

            # insert the new series
            query = """
            INSERT INTO Series2
                (SeriesID, SeriesComment, SeriesOwner, DefaultKeywordGroup)
                VALUES
                (%s,%s,%s,%s)
            """
        else:
            # check for dupes
            if DBInterface.record_match_count("Series2", \
                            ("SeriesID", "!SeriesNum"), \
                            (id, self.number) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Series named "%s" already exists.\nPlease enter a different Series ID.'), 'utf8')
                else:
                    prompt = _('A Series named "%s" already exists.\nPlease enter a different Series ID.')
                raise SaveError, prompt % self.id

            # update the record
            query = """
            UPDATE Series2
                SET SeriesID = %s,
                    SeriesComment = %s,
                    SeriesOwner = %s,
                    DefaultKeywordGroup = %s
                WHERE SeriesNum = %s
            """
            values = values + (self.number,)

        c = DBInterface.get_db().cursor()
        c.execute(query, values)
        c.close()
        # if we saved a new series, NUM was auto-assigned so our
        # 'local' data is out of date.  re-sync
        if (self.number == 0):
            self.db_load_by_name(self.id)

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database.  Raises
        RecordLockedError exception if record is locked and unable to be
        deleted."""

        result = 1
        try:
            # Initialize delete operation, begin transaction if necessary.
            (db, c) = self._db_start_delete(use_transactions)
            if (db == None):
                return      # Abort delete without even trying

            # Detect, Load, and Delete all Series Notes.
            notes = self.get_note_nums()
            for note_num in notes:
                note = Note(note_num)
                result = result and note.db_delete(0)
                del note
            del notes

            # Deletes Episodes, which in turn will delete Episode Transcripts,
            # Episode Notes, and Episode Keywords
            episodes = self.get_episode_nums()
            for episode_num in episodes:
                episode = Episode(episode_num)
                # Store the result so we can rollback the transaction on failure
                result = result and episode.db_delete(0)
                del episode
            del episodes

            # Delete the actual record.
            self._db_do_delete(use_transactions, c, result)

            # Cleanup
            c.close()
            self.clear()
        except RecordLockedError, e:

            if DEBUG:
                print "Series: RecordLocked Error", e

            # if a sub-record is locked, we may need to unlock the Series record (after rolling back the Transaction)
            if self.isLocked:
                # c (the database cursor) only exists if the record lock was obtained!
                # We must roll back the transaction before we unlock the record.
                c.execute("ROLLBACK")

                if DEBUG:
                    print "Series: roll back Transaction"
            
                c.close()

                self.unlock_record()

                if DEBUG:
                    print "Series: unlocking record"

            raise e    
        except:

            if DEBUG:
                print "Series: Exception"
            
            raise
        return result


# Private methods

    def _load_row(self, r):
        self.number = r['SeriesNum']
        self.id = r['SeriesID']
        self.comment = r['SeriesComment']
        self.owner = r['SeriesOwner']
        self.keyword_group = r['DefaultKeywordGroup']
        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.comment = DBInterface.ProcessDBDataForUTF8Encoding(self.comment)
            self.owner = DBInterface.ProcessDBDataForUTF8Encoding(self.owner)
            self.keyword_group = DBInterface.ProcessDBDataForUTF8Encoding(self.keyword_group)

    def _get_owner(self):
        return self._owner
    def _set_owner(self, owner):
        self._owner = owner
    def _del_owner(self):
        self._owner = ""
    
    def _get_kg(self):
        return self._kg
    def _set_kg(self, kg):
        self._kg = kg
    def _del_kg(self):
        self._kg = ""

# Public properties
    owner = property(_get_owner, _set_owner, _del_owner,
                        """The person responsible for creating or maintaining the Series.""")
    keyword_group = property(_get_kg, _set_kg, _del_kg,
                        """The default keyword group to be suggested for all new Episodes.""")
