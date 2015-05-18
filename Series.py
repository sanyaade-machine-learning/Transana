# Copyright (C) 2003 The Board of Regents of the University of Wisconsin System 
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

__author__ = 'Nathaniel Case <nacase@wisc.edu>, David Woods <dwoods@wcer.wisc.edu>'

# Based on code/ideas/logic from USeriesObject Delphi unit by DKW

from DataObject import DataObject
from Note import Note
from Episode import Episode
from TransanaExceptions import *
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

        values = (self.id, self.comment, self.owner, self.keyword_group)
        if (self._db_start_save() == 0):
            # duplicate Series IDs are not allowed
            if DBInterface.record_match_count("Series2", \
                            ("SeriesID",), \
                            (self.id,) ) > 0:
                raise SaveError, _('A Series named "%s" already exists.\nPlease enter a different Series ID.') % self.id

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
                            (self.id, self.number) ) > 0:
                raise SaveError, _('A Series named "%s" already exists.\nPlease enter a different Series ID.') % self.id

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
            # Delphi version unlocked the record here.  I'm tentatively
            # deciding to do that outside this function since this method
            # didn't grab a lock.  May change my mind later.
            # Personally I can't figure out why it didn't get a lock in the
            # first place

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database.  Raises
        RecordLockedError exception if record is locked and unable to be
        deleted."""

        # Initialize delete operation, begin transaction if necessary.
        (db, c) = self._db_start_delete(use_transactions)
        if (db == None):
            return      # Abort delete without even trying
       
        result = 1
        
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
        return result


# Private methods

    def _load_row(self, r):
        self.number = r['SeriesNum']
        self.id = r['SeriesID']
        self.comment = r['SeriesComment']
        self.owner = r['SeriesOwner']
        self.keyword_group = r['DefaultKeywordGroup']
        # FIXME: Old Transana would ensure "DefaultKeywordGroup" field
        # exists and add it to the table if necessary

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
