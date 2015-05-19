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

"""This module implements the Collection class as part of the Data Objects."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

DEBUG = False
if DEBUG:
    print "Collection DEBUG is ON!"

# import wxPython
import wx
# import Python's types module
import types
# import Transana's Clip Object
import Clip
# import Transana's base Data Object
import DataObject
# import Transana's Database Interface
import DBInterface
# import Transana's Note Object
import Note
# import Transana's Exceptions
from TransanaExceptions import *
# import Transana's Globals
import TransanaGlobal

class Collection(DataObject.DataObject):
    """This class defines the structure for a collection object.  A collection
    holds information about a group of video clips."""

    def __init__(self, id_or_num=None, parent=0):
        """Initialize an Collection object.  If a record ID number or
        Collection ID is given, load it from the Database."""
        DataObject.DataObject.__init__(self)
        if type(id_or_num) in (int, long):
            self.db_load_by_num(id_or_num)
        elif isinstance(id_or_num, types.StringTypes):
            self.db_load_by_name(id_or_num, parent)
             
        self._parentName = ""  # This property is looked up and loaded when requested.  It needs to be initialized, however, to a blank string.
        

# Public methods
    def __repr__(self):
        str = 'Collection Object:\n'
        str = str + "Number = %s\n" % self.number
        str = str + "id = %s\n" % self.id.encode('utf8')
        str = str + "parent = %s\n" % self.parent
        str = str + "comment = %s\n" % self.comment.encode('utf8')
        str = str + "owner = %s\n" % self.owner.encode('utf8')
        str = str + "Default KWG = %s\n\n" % self.keyword_group.encode('utf8')
        return str

    def db_load_by_name(self, name, parent_num=0):
        """Load a record by ID / Name.  Raise a RecordNotFound exception
        if record is not found in database."""
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            name = name.encode(TransanaGlobal.encoding)
        db = DBInterface.get_db()

        c = db.cursor()
        
        if parent_num:
            query = """
            SELECT * FROM Collections2
                WHERE   CollectID = %s AND
                        ParentCollectNum = %s
            """
            c.execute(query, (name, parent_num))
        else:
            query = """
            SELECT * FROM Collections2
                WHERE   CollectID = %s AND
                        (ParentCollectNum = %s OR ParentCollectNum = %s)
            """
            c.execute(query, (name, 0, None))

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
        """Load a record by record number. Raise a RecordNotFound exception
        if record is not found in database."""
        db = DBInterface.get_db()
        query = """
        SELECT * FROM Collections2
            WHERE CollectNum = %s
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
            raise SaveError, _("Collection ID is required.")
        
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
        
        fields = ("CollectID", "ParentCollectNum", "CollectComment",
                    "CollectOwner", "DefaultKeywordGroup")
        values = (id, self.parent, comment, owner, keyword_group)
        if (self._db_start_save() == 0):        # Add new collection
            # Duplicate Collection IDs are not allowed within a collection
            if DBInterface.record_match_count("Collections2", \
                            ("CollectID", "ParentCollectNum"),
                            (id, self.parent) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Collection named "%s" already exists.\nPlease enter a different Collection ID.'), 'utf8')
                else:
                    prompt = _('A Collection named "%s" already exists.\nPlease enter a different Collection ID.')
                raise SaveError, prompt % self.id

            # insert the new collection
            query = """
            INSERT INTO Collections2
                (%s,%s,%s,%s,%s)
                VALUES
                (%%s,%%s,%%s,%%s,%%s)
            """ % fields
        else:               # Update existing collection
            
            # check for dupes
            if DBInterface.record_match_count("Collections2", \
                            ("CollectID", "ParentCollectNum", "!CollectNum"),
                            (id, self.parent, self.number) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Collection named "%s" already exists.\nPlease enter a different Collection ID.'), 'utf8')
                else:
                    prompt = _('A Collection named "%s" already exists.\nPlease enter a different Collection ID.')
                raise SaveError, prompt % self.id
            # update the record
            query = """
            UPDATE Collections2
                SET """
            for field in fields:
                query = "%s%s = %%s,\n\t" % (query, field)
            query = query[:-3] + "\n"
            query = query + "    WHERE CollectNum = %s"
            values = values + (self.number,)

        c = DBInterface.get_db().cursor()
        c.execute(query, values)
        c.close()
        # if new collection, Number was auto assigned, so resync.
        if (self.number == 0):
            self.db_load_by_name(self.id, self.parent)
        

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database.  Raises
        RecordLockedError exception if the record is locked and unable to
        be deleted."""
        result = 1
        try:
            # Initialize delete operation, begin transaction if necessary
            (db, c) = self._db_start_delete(use_transactions)
            if (db == None):
                return      # Abort delete

            # Delete all Collection-based Filter Configurations
            #   Delete Collection Clip Data Export records
            DBInterface.delete_filter_records(4, self.number)
            #   Delete Collection Report records
            DBInterface.delete_filter_records(12, self.number)

            # Detect, Load, and Delete all Collection Notes
            notes = self.get_note_nums()
            for note_num in notes:
                note = Note.Note(note_num)
                result = result and note.db_delete(0)
                del note
            del notes

            # Delete Clips, which in turn will delete Clip transcripts/notes/kws
            clips = DBInterface.list_of_clips_by_collection(self.id, self.parent)
            for (clipNo, clip_id, collNo) in clips:
                clip = Clip.Clip(clipNo)
                result = result and clip.db_delete(0)
                del clip
            del clips

            # Delete all Nested Collections
            for (collNo, collID, parentCollNo) in DBInterface.list_of_collections(self.number):
                tempCollection = Collection(collNo)
                result = result and tempCollection.db_delete(0)
                del tempCollection

            # Delete the actual record
            self._db_do_delete(use_transactions, c, result)
            
            # Cleanup
            c.close()
            self.clear()
        except RecordLockedError, e:

            if DEBUG:
                print "Collection: RecordLocked Error", e

            # if a sub-record is locked, we may need to unlock the Collection record (after rolling back the Transaction)
            if self.isLocked:
                # c (the database cursor) only exists if the record lock was obtained!
                # We must roll back the transaction before we unlock the record.
                c.execute("ROLLBACK")

                if DEBUG:
                    print "Collection: roll back Transaction"
            
                c.close()

                self.unlock_record()

                if DEBUG:
                    print "Collection: unlocking record"

            raise e    
        except:

            if DEBUG:
                print "Collection: Exception"
            
            raise

        if DEBUG:
            print
        return result

    def GetNodeData(self):
        """ Returns the Node Data list (list of parent collections) needed for Database Tree Manipulation """
        # Initialize the nodeData structure
        nodeData = ()
        # If this is a nested collection (parent != 0), we have to load the full nesting structure here
        if self.parent != 0:
            # Load the parent collection
            parentColl = Collection(self.parent)
            # add the parent's name to the data structure
            nodeData = (parentColl.id,) + nodeData
            # repeat until we get to the root, where the parent is 0
            while parentColl.parent != 0:
                # Load the parent collection
                parentColl = Collection(parentColl.parent)
                # add the parent's name to the data structure
                nodeData = (parentColl.id,) + nodeData
        # Complete the nodeData structure by the Collection Name to the end
        nodeData = nodeData + (self.id,)
        return nodeData

    def GetNodeString(self):
        """ Returns a string that delineates the full nested collection structure for the present collection """
        # Initialize a string variable
        st = ''
        # Get the collection's Node Data
        nodeData = self.GetNodeData()
        # For each node in the Node Data ...
        for node in nodeData:
            # ... if this isn't the first node, ...
            if st != '':
                # ... add a ">" character to indicate we're moving to a new nesting level
                st += ' > '
            # ... and append the node text onto the string
            st += node
        # Return the string
        return st
        
        
# Private methods    

    def _load_row(self, r):
        self.number = r['CollectNum']
        self.id = r['CollectID']
        self.parent = r['ParentCollectNum']
        self.comment = r['CollectComment']
        self.owner = r['CollectOwner']
        if r.has_key('DefaultKeywordGroup'):
            self.keyword_group = r['DefaultKeywordGroup']
        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.comment = DBInterface.ProcessDBDataForUTF8Encoding(self.comment)
            self.owner = DBInterface.ProcessDBDataForUTF8Encoding(self.owner)
            self.keyword_group = DBInterface.ProcessDBDataForUTF8Encoding(self.keyword_group)

    def _get_parent(self):
        return self._parent
    def _set_parent(self, parent):
        self._parent = parent
    def _del_parent(self):
        self._parent = 0
    
    def _get_parentName(self):
        # If there is no parent, return blank name
        if self.parent == 0:
            return ""
        # ELSE if Parent Name is not known, look it up
        elif self._parentName == "":
            tempColl = Collection(self.parent)
            self._parentName = tempColl.id
            return self._parentName
        # ELSE if Parent Name IS known, return it
        else:
            return self._parentName
    
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
    parent = property(_get_parent, _set_parent, _del_parent,
                        "Parent Collection number for nested collections.""")
    parentName = property(_get_parentName, None, None,                           # Read Only property
                        "Parent Collection Name for nested collections.""")
    owner = property(_get_owner, _set_owner, _del_owner,
                        """Person responsible for creating or maintaining the Collection.""")
    keyword_group = property(_get_kg, _set_kg, _del_kg,
                        """Default keyword group to be suggested for all new clips.""")
