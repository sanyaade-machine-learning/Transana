# Copyright (C) 2003-2015 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Library class (formerly the Series class) as part of the Data Objects."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

DEBUG = False
if DEBUG:
    print "Library DEBUG is ON!"

# Import wxPython
import wx
# import Python's types module
import types
# Import Transana's base Data Object
import DataObject
# import Transana's Database Interface
import DBInterface
# import Transana Dialogs
import Dialogs
# import Transana Document object
import Document
# import Transana's Episode object
import Episode
# Import Transana's Note object
import Note
# import Transana's Constants
import TransanaConstants
# import Transana's Exceptions
from TransanaExceptions import *
# import Transana's Globals
import TransanaGlobal

class Library(DataObject.DataObject):
    """This class defines the structure for a Library object.  A Library object
    holds information about a group (e.g., a Library) of data files
    (e.g., Documents and Episodes)."""

    def __init__(self, id_or_num=None):
        """Initialize an Library object.  If a record ID number or Library ID
        is given, load it from the database."""
        DataObject.DataObject.__init__(self)
        if type(id_or_num) in (int, long):
            self.db_load_by_num(id_or_num)
        elif isinstance(id_or_num, types.StringTypes):
            self.db_load_by_name(id_or_num)

    def __repr__(self):
        """ String Representation of a Library Object """
        str = 'Library Object:\n'
        str += 'number = %s\n' % self.number
        str += 'id = %s (%s)\n' % (self.id.encode('utf8'), type(self.id))
        str += 'comment = %s\n' % self.comment.encode('utf8')
        str += 'owner = %s\n' % self.owner.encode('utf8')
        str += 'keyword_group = %s\n\n' % self.keyword_group.encode('utf8')
#        str += "isLocked = %s\n" % self._isLocked
#        str += "recordlock = %s\n" % self.recordlock
#        str += "locktime = %s\n" % self.locktime
        return str
        
    def __eq__(self, other):
        """ Object equality check """
        if other == None:
            return False
        else:
            return self.__dict__ == other.__dict__

# Public methods

    def db_load_by_name(self, name):
        """Load a record by ID / Name.  Raise a RecordNotFound exception
        if record is not found in database."""
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            name = name.encode(TransanaGlobal.encoding)
        # Get the database connection
        db = DBInterface.get_db()
        # Define the load query
        query = """
        SELECT * FROM Series2
            WHERE SeriesID = %s
        """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, (name, ))
        # rowcount doesn't work for sqlite!
        if TransanaConstants.DBInstalled == 'sqlite3':
            # ... so assume one row return for now
            n = 1
        # If not sqlite ...
        else:
            # ... we can use rowcount
            n = c.rowcount
        # If not one row ...
        if (n != 1):
            # ... close the database cursor ...
            c.close()
            # ... clear the current series ...
            self.clear()
            # ... and raise an exception
            raise RecordNotFoundError, (name, n)
        # If exactly one row, or sqlite ...
        else:
            # ... get the query results ...
            r = DBInterface.fetch_named(c)
            # If sqlite and no results ...
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}):
                # ... close the database cursor ...
                c.close()
                # ... clear the current series ...
                self.clear()
                # ... and raise an exception
                raise RecordNotFoundError, (name, 0)
            # Load the database results into the Series object
            self._load_row(r)
        # Close the database cursor ...
        c.close()

    def db_load_by_num(self, num):
        """Load a record by record number."""
        # Get the database connection
        db = DBInterface.get_db()
        # Define the load query
        query = """
        SELECT * FROM Series2
            WHERE SeriesNum = %s
        """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, (num, ))
        # rowcount doesn't work for sqlite!
        if TransanaConstants.DBInstalled == 'sqlite3':
            # ... so assume one result for now
            n = 1
        # If not sqlite ...
        else:
            # ... we can use rowcount
            n = c.rowcount
        # If something other than one record is returned ...
        if (n != 1):
            # ... close the database cursor ...
            c.close()
            # ... clear the current Library object ...
            self.clear()
            # ... and raise an exception
            raise RecordNotFoundError, (num, n)
        # If exactly one row is returned, or we're using sqlite ...
        else:
            # ... load the query results ...
            r = DBInterface.fetch_named(c)
            # If sqlite and no data is loaded ...
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}):
                # ... close the database cursor ...
                c.close()
                # ... clear the current Library object ...
                self.clear()
                # ... and raise an exception
                raise RecordNotFoundError, (num, 0)
            # Place the loaded data in the object
            self._load_row(r)
        # Close the database cursor ...
        c.close()
        
    def db_save(self, use_transactions=True):
        """Save the record to the database using Insert or Update as
        appropriate."""

        # Sanity checks
        if self.id == "":
            raise SaveError, _("Library ID is required.")

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

            # duplicate Library IDs are not allowed
            if DBInterface.record_match_count("Series2", \
                            ("SeriesID",), \
                            (id,) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Library named "%s" already exists.\nPlease enter a different Library ID.'), 'utf8')
                else:
                    prompt = _('A Library named "%s" already exists.\nPlease enter a different Library ID.')
                raise SaveError, prompt % self.id

            # insert the new Library
            query = """
            INSERT INTO Series2
                (SeriesID, SeriesComment, SeriesOwner, DefaultKeywordGroup)
                VALUES
                (%s, %s, %s, %s)
            """
        else:
            # check for dupes
            if DBInterface.record_match_count("Series2", \
                            ("SeriesID", "!SeriesNum"), \
                            (id, self.number) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Library named "%s" already exists.\nPlease enter a different Library ID.'), 'utf8')
                else:
                    prompt = _('A Library named "%s" already exists.\nPlease enter a different Library ID.')
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
        # Get a database cursor
        c = DBInterface.get_db().cursor()
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, values)
        # Close the database cursor
        c.close()
        # if we saved a new Library, NUM was auto-assigned so our
        # 'local' data is out of date.  re-sync
        if (self.number == 0):
            self.db_load_by_name(self.id)

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database.  Raises
        RecordLockedError exception if record is locked and unable to be
        deleted."""
        # Assume success
        result = 1
        try:
            # Initialize delete operation, begin transaction if necessary.
            (db, c) = self._db_start_delete(use_transactions)
            if (db == None):
                return      # Abort delete without even trying

            # Delete all Library-based Filter Configurations
            #   Delete Library Keyword Sequence Map records
            DBInterface.delete_filter_records(5, self.number)
            #   Delete Library Keyword Bar Graph records
            DBInterface.delete_filter_records(6, self.number)
            #   Delete Library Keyword Percentage Map records
            DBInterface.delete_filter_records(7, self.number)
            #   Delete Library Report records
            DBInterface.delete_filter_records(10, self.number)
            #   Delete Library Clip Data Export records
            DBInterface.delete_filter_records(14, self.number)

            # Detect, Load, and Delete all Library Notes.
            notes = self.get_note_nums()
            for note_num in notes:
                note = Note.Note(note_num)
                result = result and note.db_delete(0)
                del note
            del notes

            # Deletes Episodes, which in turn will delete Episode Transcripts,
            # Episode Notes, and Episode Keywords
            episodes = DBInterface.list_of_episodes_for_series(self.id)
            for (episode_num, episode_id, series_num) in episodes:
                episode = Episode.Episode(episode_num)
                # Store the result so we can rollback the transaction on failure
                result = result and episode.db_delete(0)
                del episode
            del episodes

            # Deletes Documents, which in turn will delete Document Notes and Document Keywords
            documents = DBInterface.list_of_documents(self.number)
            for (document_num, document_id, library_num) in documents:
                document = Document.Document(document_num)
                # Store the result so we can rollback the transaction on failure
                result = result and document.db_delete(0)
                del document
            del documents

            # Delete the actual record.
            self._db_do_delete(use_transactions, c, result)

            # Cleanup
            c.close()
            self.clear()
        except RecordLockedError, e:
            # if a sub-record is locked, we may need to unlock the Library record (after rolling back the Transaction)
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
                # ... then unlock it ...
                self.unlock_record()
            # ... and pass on the exception.
            raise e
        except:
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
                        """The person responsible for creating or maintaining the Library.""")
    keyword_group = property(_get_kg, _set_kg, _del_kg,
                        """The default keyword group to be suggested for all new Episodes.""")
