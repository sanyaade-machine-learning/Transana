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
import TransanaConstants
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
        # Adjust the query for sqlite if needed
        db = DBInterface.get_db()
        c = db.cursor()
        query = DBInterface.FixQuery(query)
        c.execute(query, (self.number,))
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
        # NOTE:  Series has been renamed Library, but we didn't change the table name
        if t == 'Library':
            t = 'Series'
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
        if type(self).__name__ == 'Library':
            name = 'Series'
        elif type(self).__name__ == "Collection":
            name = "Collect"

        return name

    def _num(self):
        """Return the SQL field for table record number."""
        # FIXME: Can we determine what the auto-increment field is to just
        # handle this in all cases easily?
        
        # General case
        numname = type(self).__name__ + "Num"        # e.g. "EpisodeNum"
        # Exceptions
        if type(self).__name__ == 'Library':
            numname = 'SeriesNum'
        elif type(self).__name__ == "Collection":
            numname = "CollectNum"
        return numname
 
    def _get_db_fields(self, fieldlist, c=None):
        """Get the values of fields from the database for the currently
        loaded record.  Use existing cursor if it exists, otherwise create
        a new one.  Return a tuple containing the values obtained."""
        # If the object number is 0, there's no record in the database ...
        if self.number == 0:
            # ... so just return an empty tuple
            return ()
        
        # Get the table name and the name of the Number property
        tablename = self._table()
        numname = self._num()
        
        # Create a flag tht indicates whether the cursor was passed in
        close_c = False
        # If no cursor was passed in ...
        if (c == None):
            # ... update the flag ...
            close_c = True
            # ... get a database reference ...
            db = DBInterface.get_db()
            # ... and create a database cursor
            c = db.cursor()

        # Determine the field values needed for the query
        fields = ""
        for field in fieldlist:
            fields = fields + field + ", "
        fields = fields[:-2]

        # Formulate the query based on the fields
        query = "SELECT " + fields + " FROM " + tablename + \
                "  WHERE " + numname + " = %s"
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, (self.number, ))
        # Get query row results
        qr = c.fetchone()

        if DEBUG:
            print "DataObject._get_db_fields():\n", query, qr
            print
        
        # If we created the cursor locally (as flagged) ...
        if (close_c):
            # ... close the database cursor
            c.close()
        # Return the Query Results
        return qr

    def _set_db_fields(self, fields, values, c=None):
        """Set the values of fields in the database for the currently loaded
        record.  Use existing cursor if it exists, otherwise create a new
        one."""
        # If the object number is 0, there's no record in the database ...
        if self.number == 0:
            # ... so just return
            return

        # Get the table name and the name of the Number property
        tablename = self._table()
        numname = self._num()

        # Create a flag tht indicates whether the cursor was passed in
        close_c = False
        # If no cursor was passed in ...
        if (c == None):
            # ... update the flag ...
            close_c = True
            # ... get a database reference ...
            db = DBInterface.get_db()
            # ... and create a database cursor
            c = db.cursor()

        # Determine the field values needed for the query
        fv = ""
        for f, v in map(None, fields, values):
            fv = fv + f + " = " + "%s, "
        fv = fv[:-2]

        # Formulate the query based on the fields
        query = "UPDATE " + tablename + " SET " + fv + " WHERE " + numname + " = %s"
        # Modify the values by adding the object number on the end
        values = values + (self.number,)
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, values)
        # If we created the cursor locally (as flagged) ...
        if (close_c):
            # ... close the database cursor
            c.close()

    def _db_start_save(self):
        """Return 0 if creating new record, 1 if updating an existing one."""
        tname = _(type(self).__name__)
        # You can save a Clip Transcript with a blank Transcript ID!
        if (self.id == "") and (tname != _('Transcript')):
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Cannot save a %s with a blank %s ID"), 'utf8')
            else:
                prompt = _("Cannot save a %s with a blank %s ID")
            raise SaveError, prompt % (tname.decode('utf8'), tname.decode('utf8'))
        else:
            # Verify record lock is still good
            if (self.number == 0) or \
                ((self.record_lock == DBInterface.get_username()) and
                ((self.lock_time == None) or
                 ((DBInterface.ServerDateTime() - self.lock_time).days <= 1))):
                # If record num is 0, this is a NEW record and needs to be
                # INSERTed.  Otherwise, it is an existing record to be UPDATEd.
                if (self.number == 0):
                    return 0
                else:
                    return 1
            else:
                raise SaveError, _("Record lock no longer valid.\nYour changes cannot be saved.")
                
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
        # Define the Delete query
        query = "DELETE FROM " + tablename + \
                "   WHERE " + numname + " = %s"
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, (self.number, ))
        # If we're using Transactions ...
        if (use_transactions):
            # ... and the result exists ...
            if (result):
                # Commit the transaction
                c.execute("COMMIT")

                if DEBUG:
                    print "Transaction committed"
            # ... and the result does NOT exist (failed) ....
            else:
                # Rollback transaction because some part failed
                c.execute("ROLLBACK")

                if DEBUG:
                    print "Transaction rolled back"

                # if the object has a number (and therefore existed before) ...   
                if (self.number != 0):
                    # ... release the record lock when the delete fails
                    self.unlock_record()
                    
                    if DEBUG:
                        print "Record '%s' unlocked" % self.id

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
        # Get the Record Lock Time from the Database
        lt = self._get_db_fields(('LockTime',))
        # If a Lock Time has been specified ...
        if (len(lt) > 0) and (lt[0] is not None):
            # ... If we're using sqlite, we get a string and need to convert it to a datetime object
            if TransanaConstants.DBInstalled in ['sqlite3']:
                import datetime
                tempDate = datetime.datetime.strptime(lt[0], '%Y-%m-%d %H:%M:%S.%f')
                return tempDate
            # ... If we're using MySQL, we get a MySQL DateTime value
            else:
                return lt[0]
        # If we don't get a Lock Time ...
        else:
            # ... return the current Server Time
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
