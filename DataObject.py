# Copyright (C) 2003 - 2008 The Board of Regents of the University of Wisconsin System 
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

DEBUG = False
if DEBUG:
    print "DataObject DEBUG is ON!"

"""This module implements the DataObject class as part of the Data Objects.  This is the
   base class for Series, Episode, CoreData, Transcript, Collection, Clip, and Note Objects.
   It provides the following public methods:

    clear()
    duplicate()
    lock_record()
    unlock_record()
    get_note_nums()
"""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

import DBInterface
import inspect
import copy
import Misc
from MySQLdb import *
from TransanaExceptions import *
import TransanaGlobal

"""Concerning MySQL literal values:

In the Delphi version of Transana, we had to use a string escaping function
(named DoubleApostrophes) to handle characters that would result in a SQL
error or it being stored incorrectly.

Now, all the messy details of SQL literal values are handled in the MySQLdb
module.  When we use the execute() method on a DB connection object, anything
in the query with %s with a corresponding item in the argument list will
get translated properly.  Be careful with this, because it's NOT the same
as using C-printf-style strings.

This means we can do, for example:

c.execute("SELECT * From Clips2 a, Collections2 b" + \
          "  WHERE ClipID = %s AND " + \
          "        a.CollectNum = %s AND" + \
          "        b.CollectID = %s", \
          ("some clip ID", 2, "some collection ID"))

And it will 'just work'.  When a string is passed in the argument list, it
will automatically get the single quotes around it.  Integers will just
be converted to a string without the quotes as needed.  This also works for
other Python data types.  The short explanation is that we no
longer need to care about how to format literal values when writing MySQL
queries.  Just put %s in the query format string and then pass the Python
variable to the argument list.

The actual query sent to MySQL for the above example would be:

SELECT * From Clips2 a, Collections2 b
  WHERE ClipID = 'some clip ID' AND
        a.CollectNum = 2 AND
        b.CollectID = 'some collection ID'
"""

class DataObject(object):
    """This class defines the features common among all classes in the
    Data Objects component group.  The Data Object classes will inherit
    from this base class."""

    def __init__(self):
        """Initialize an DataObject object."""
        self.clear()
        # In Transana-MU, we need to track whether an object is locked or not.
        self._isLocked = False
        

# Public methods
    def clear(self):
        """Clear all properties, resetting them to default values."""
        attrdesc = inspect.classify_class_attrs(self.__class__)
        for attr in attrdesc:
            if attr[1] == "property":
                try:
                    delattr(self, attr[0])
                except AttributeError, e:
                    pass # probably a non-deletable attribute
        
    def duplicate(self):
        """Return a copy of the object with only the record number changed."""
        # copy other attrs with setattr() or something
        # then manually do id, comment, rec*, etc because dir() won't
        # see the properties from this base class.

        # alternatively use copy.*
        newobj = copy.copy(self)
        newobj.number = 0
        return newobj
        
    def lock_record(self):
        """Lock a record.  If the lock is unable to be obtained, a
        RecordLockedError exception is raised with the username of the lock
        holder passed."""

        if self.number == 0:    # no record or new record not yet in database loaded
            return
            
        tablename = self._table()
        numname = self._num()
        
        db = DBInterface.get_db()
        c = db.cursor()
        lq = self._get_db_fields(('RecordLock', 'LockTime'), c)

        if (lq[1] == None) or (lq[0] == "") or ((DBInterface.ServerDateTime() - lq[1]).days > 1):
            # Lock the record
            self._set_db_fields(    ('RecordLock', 'LockTime'),
                                    (DBInterface.get_username(),
                                    str(DBInterface.ServerDateTime())[:-3]), c)
            c.close()
            # Indicate that the object was successfully locked
            self._isLocked = True

            if DEBUG:
                print "DataObject.lock_record(): Record '%s' locked by '%s'" % (self.number, DBInterface.get_username())
        else:
            # We just raise an exception here since GUI code isn't appropriate.
            c.close()

            if DEBUG:
                print "DataObject.lock_record(): Record %s locked by %s raises exception" % (self.id, lq[0])
                
            raise RecordLockedError, lq[0]  # Pass name of person

        
    def unlock_record(self):
        """Unlock a record."""

        self._set_db_fields(    ('RecordLock', 'LockTime'),
                                ('', None), None)
        # Indicate that the object was successfully unlocked
        self._isLocked = False

        if DEBUG:
            print "DataObject.unlock_record(): Record '%s' has been unlocked" % (self.id,)


    def get_note_nums(self):
        """Get a list of Note numbers that belong to this Object."""
        notelist = []
        
        t = self._prefix()

        query = """
        SELECT NoteNum FROM Notes2
            WHERE %sNum = %%s
            ORDER BY NoteID
        """ % (t,)

        db = DBInterface.get_db()
        c = db.cursor()
        c.execute(query, (self.number,))
        r = c.fetchall()    # return array of tuples with results
        for tup in r:
            notelist.append(tup[0])
        c.close()
        return notelist

    # TODO:  Eliminate this.  This belongs in the Series object, not in the DataObject superclass
    def get_episode_nums(self):
        """Get a list of Episode numbers that belong to this Series."""
        notelist = []
        t = self._prefix()
        table = self._table()
       
        query = """
        SELECT EpisodeNum FROM Episodes2 a, %s b
            WHERE a.%sNum = b.%sNum and
                    %sID = %%s
            ORDER BY EpisodeID
        """ % (table, t, t, t)

        if type(self.id).__name__ == 'unicode':
            id = self.id.encode(TransanaGlobal.encoding)
        else:
            id = self.id

        c = DBInterface.get_db().cursor()
        c.execute(query, (id,))
        r = c.fetchall()    # return array of tuples with results
        for tup in r:
            notelist.append(tup[0])
        c.close()
        return notelist


# Private methods

    def _table(self):
        """Return the SQL table name."""
        # general case
        t = type(self).__name__
        # NOTE: the CoreData table does not follow the pattern of having an 's' near the end of the table name
        if (t[-1] != "s") and (t != 'CoreData'):
            t = t + "s"
        t = t + "2"
        return t

    def _prefix(self):
        """Return the SQL field prefix for the object type."""
        # FIXME: Can we determine what the auto-increment field is to just
        # handle this in all cases easily?
        # General case
        name = type(self).__name__

        # Exceptions
        if type(self).__name__ == "Collection":
            name = "Collect"

        return name

    def _num(self):
        """Return the SQL field for table record number."""
        # FIXME: Can we determine what the auto-increment field is to just
        # handle this in all cases easily?
        
        # General case
        numname = type(self).__name__ + "Num"        # e.g. "SeriesNum"
        # Exceptions
        if type(self).__name__ == "Collection":
            numname = "CollectNum"
        return numname
 
    def _get_db_fields(self, fieldlist, c=None):
        """Get the values of fields from the database for the currently
        loaded record.  Use existing cursor if it exists, otherwise create
        a new one.  Return a tuple containing the values obtained."""
        
        if self.number == 0:    # no record loaded?
            return ()
        
        tablename = self._table()
        numname = self._num()
        
        close_c = 0
        if (c == None):
            close_c = 1
            db = DBInterface.get_db()
            c = db.cursor()

        fields = ""
        for field in fieldlist:
            fields = fields + field + ", "
        fields = fields[:-2]

        query = "SELECT " + fields + " FROM " + tablename + "\n" + \
                "  WHERE " + numname + " = %s\n"
        c.execute(query, self.number)
        
        qr = c.fetchone()       # get query row results

        if DEBUG:
            print "DataObject._get_db_fields():\n", query, qr
            print
        
        if (close_c):
            c.close()
        return qr

    def _set_db_fields(self, fields, values, c=None):
        """Set the values of fields in the database for the currently loaded
        record.  Use existing cursor if it exists, otherwise create a new
        one."""

        if self.number == 0:    # no record loaded?
            return
 
        tablename = self._table()
        numname = self._num()

        close_c = 0
        if (c == None):
            close_c = 1
            db = DBInterface.get_db()
            c = db.cursor()

        fv = ""
        for f, v in map(None, fields, values):
            fv = fv + f + " = " + "%s,\n\t\t"
        fv = fv[:-4]
        
        query = "UPDATE " + tablename + "\n  SET " + fv + "\n  WHERE " + numname + " = %s\n"

        c.execute(query, values + (self.number,))
        if (close_c):
            c.close()

    def _db_start_save(self):
        """Return 0 if creating new record, 1 if updating an existing one."""
        tname = type(self).__name__
        # You can save a Clip Transcript with a blank Transcript ID!
        if (self.id == "") and (tname != 'Transcript'):
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Cannot save a %s with a blank %s ID"), 'utf8')
            else:
                prompt = _("Cannot save a %s with a blank %s ID")
            raise SaveError, prompt % (tname, tname)
        else:
            # Verify record lock is still good
            if (self.number == 0) or \
                ((self.record_lock == DBInterface.get_username()) and
                ((DBInterface.ServerDateTime() - self.lock_time).days <= 1)):
                # If record num is 0, this is a NEW record and needs to be
                # INSERTed.  Otherwise, it is an existing record to be UPDATEd.
                if (self.number == 0):
                    return 0
                else:
                    return 1
            else:
                raise SaveError, _("Record lock no longer valid.")
                
                        
    def _db_start_delete(self, use_transactions):
        """Initialize delete operation and begin transaction if necessary.
        This is a helper method intended for sub-class db_delete() methods."""
        if (self.number == 0):
            self.clear()
            raise DeleteError, _("Invalid record number (0)")

        self.lock_record()

        if DEBUG:
            print "Record '%s' locked" % self.id

        db = DBInterface.get_db()
        c = db.cursor()
 
        if use_transactions:
            query = "BEGIN"   # Begin a transaction
            c.execute(query)

            if DEBUG:
                print "Begin Delete Transaction"

        return (db, c)


    def _db_do_delete(self, use_transactions, c, result):
        """Do the actual record delete and handle the transaction as needed.
        This is a helper method intended for sub-class db_delete() methods."""
        tablename = self._table()
        numname = self._num()

        query = "DELETE FROM " + tablename + "\n" + \
                "   WHERE " + numname + " = %s\n"
        c.execute(query, self.number)

        if (use_transactions):
            # Commit the transaction
            if (result):
                c.execute("COMMIT")

                if DEBUG:
                    print "Transaction committed"
                    
            else:
                # Rollback transaction because some part failed
                c.execute("ROLLBACK")

                if DEBUG:
                    print "Transaction rolled back"
                    
                if (self.number != 0):
                    self.unlock_record()
                    
                    if DEBUG:
                        print "Record '%s' unlocked" % self.id
        
        return

    def _get_number(self):
        return self._number
    def _set_number(self, number):
        self._number = number
    def _del_number(self):
        self._number = 0

    def _get_id(self):
        return self._id
    def _set_id(self, id):
        self._id = Misc.unistrip(id)
    def _del_id(self):
        self._id = ""

    def _get_comment(self):
        return self._comment
    def _set_comment(self, comment):
        self._comment = comment
    def _del_comment(self):
        self._comment = ""

    def _get_rl(self):
        return self._get_db_fields(('RecordLock',))[0]

    def _get_lt(self):
        lt = self._get_db_fields(('LockTime',))
        if len(lt) > 0:
            return lt[0]
        else:
            return DBInterface.ServerDateTime()

    def _get_isLocked(self):
        return self._isLocked
    def _set_isLocked(self, lock):
        self._isLocked = lock
    def _del_isLocked(self):
        self._isLocked = False

# Public properties
    number = property(_get_number, _set_number, _del_number,
                        """Record number (auto-incremented database field).""")
    id = property(_get_id, _set_id, _del_id,
                        """ID or Name (required).""")
    comment = property(_get_comment, _set_comment, _del_comment,
                        """Description of the Object.""")
    record_lock = property(_get_rl, None, None,
                        """Username of person who has locked the record
                        (Read only).""")
    lock_time = property(_get_lt, None, None,
                        """Time of the last record lock (Read only).""")
    isLocked = property(_get_isLocked, _set_isLocked, _del_isLocked,
                        """Object Is Locked?""")
