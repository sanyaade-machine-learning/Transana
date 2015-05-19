# Copyright (C) 2003 - 2012 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Note class as part of the Data Objects."""

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

# import wxPython
import wx
# import Python's types module
import types
# import Transana's base Data Object
import DataObject
# import Transana's Database Interface
import DBInterface
# import Transana's Exceptions
from TransanaExceptions import *
# import Transana's Globals
import TransanaGlobal

class Note(DataObject.DataObject):
    """This class defines the structure for a note object.  A note object
    holds a note that can be attached to various objects."""

    def __init__(self, id_or_num=None, **kwargs):
        """Initialize an Note object."""
        DataObject.DataObject.__init__(self)
        if type(id_or_num) in (int, long):
            self.db_load_by_num(id_or_num)
        elif isinstance(id_or_num, types.StringTypes):
            self.db_load_by_name(id_or_num, **kwargs)

    def __repr__(self):
        str = "Note Object:\n"
        str = str + "number = %s\n" % self.number
        str = str + 'id = %s\n' % self.id
        str = str + 'notetype = %s\n' % self.notetype
        str = str + 'comment = %s\n' % self.comment
        str = str + "series_num = %s\n" % self.series_num
        str = str + "episode_num = %s\n" % self.episode_num
        str = str + "transcript_num = %s\n" % self.transcript_num
        str = str + "collection_num = %s\n" % self.collection_num
        str = str + "clip_num = %s\n" % self.clip_num
        str = str + "author = %s\n" % self.author
        str = str + "text = %s\n\n" % self.text
        return str.encode('utf8')


# Public methods
    def db_load_by_num(self, num):
        """Load a record by record number."""
        
        db = DBInterface.get_db()
        query = """SELECT * FROM Notes2
                   WHERE NoteNum = %s"""
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

       
    def db_load_by_name(self, note_id, **kwargs):
        """Load a Note by Note ID and the Series, Episode, Collection,
        Clip, or Transcript number the Note belongs to.  Record numbers
        are passed for one of the parameters after Note ID.

        Example: db_load_by_name("My note", Collection=1)
        """
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            note_id = note_id.encode(TransanaGlobal.encoding)
        db = DBInterface.get_db()
        if kwargs.has_key("Series"):
            q = "SeriesNum"
        elif kwargs.has_key("Episode"):
            q = "EpisodeNum"
        elif kwargs.has_key("Collection"):
            q = "CollectNum"
        elif kwargs.has_key("Clip"):
            q = "ClipNum"
        elif kwargs.has_key("Transcript"):
            q = "TranscriptNum"

        num = kwargs.values()[0]
        
        if type(num) != int and type(num) != long:
            raise ProgrammingError, _("Integer record number required.")
            
        query = """SELECT * FROM Notes2
                   WHERE NoteID = %%s AND
                   %s = %%s""" % q
        c = db.cursor()
        c.execute(query, (note_id, num))
        n = c.rowcount
        if (n != 1):
            c.close()
            self.clear()
            raise RecordNotFoundError, (note_id, n)
        else:
            r = DBInterface.fetch_named(c)
            self._load_row(r)

        c.close()
        

    def db_save(self):
        """Save the record to the database using Insert or Update as appropriate."""

        # Sanity Checks
        if ((self.series_num == 0) or (self.series_num == None)) and \
           ((self.episode_num == 0) or (self.episode_num == None)) and \
           ((self.collection_num == 0) or (self.collection_num == None)) and \
           ((self.clip_num == 0) or (self.clip_num == None)) and \
           ((self.transcript_num == 0) or (self.transcript_num == None)):
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Note %s is not assigned to any object."), 'utf8')
            else:
                prompt = _("Note %s is not assigned to any object.")
            raise SaveError, prompt % self.id

        # If we're in Unicode mode, ...
        if 'unicode' in wx.PlatformInfo:
            # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
            # variables for the data.  We don't want to change the underlying object values.  Also, this way,
            # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
            id = self.id.encode(TransanaGlobal.encoding)
            author = self.author.encode(TransanaGlobal.encoding)
            text = self.text.encode(TransanaGlobal.encoding)
        else:
            # If we don't need to encode the string values, we still need to copy them to our local variables.
            id = self.id
            author = self.author
            text = self.text

        values = (id, self.series_num, self.episode_num, self.collection_num, self.clip_num, self.transcript_num, \
                  author, text)

        # Determine if we are creating a new record or saving an existing one
        if (self._db_start_save() == 0):  # Creating new record
            # Check to see that no identical record exists
            if DBInterface.record_match_count('Notes2', \
                                              ("NoteID", "SeriesNum", "EpisodeNum", "CollectNum", "ClipNum", "TranscriptNum"), \
                                              (id, self.series_num, self.episode_num, self.collection_num, self.clip_num, self.transcript_num) ) > 0:
                targetObject = _('object')
                if (self.series_num != 0) and (self.series_num != None):
                    targetObject = _('Series')
                elif (self.episode_num != 0) and (self.episode_num != None):
                    targetObject = _('Episode')
                elif (self.transcript_num != 0) and (self.transcript_num != None):
                    targetObject = _('Transcript')
                elif (self.collection_num != 0) and (self.collection_num != None):
                    targetObject = _('Collection')
                elif (self.clip_num != 0) and (self.clip_num != None):
                    targetObject = _('Clip')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Note named "%s" already exists for this %s.'), 'utf8')
                    targetObject = unicode(targetObject, 'utf8')
                else:
                    prompt = _('A Note named "%s" already exists for this %s.')
                raise SaveError, prompt % (self.id, targetObject)

            # insert a new record
            query = """ INSERT INTO Notes2
                            (NoteID, SeriesNum, EpisodeNum, CollectNum, ClipNum, TranscriptNum,
                             NoteTaker, NoteText)
                          VALUES
                            (%s, %s, %s, %s, %s, %s, %s, %s) """
            
        else:  # Saving an existing record
            # Check to see that no identical record with a different number exists (!NoteNum specifies "Not same note number")
            if DBInterface.record_match_count('Notes2', \
                                              ("NoteID", "SeriesNum", "EpisodeNum", "CollectNum", "ClipNum", "TranscriptNum", "!NoteNum"), \
                                              (id, self.series_num, self.episode_num, self.collection_num, self.clip_num, self.transcript_num, self.number) ) > 0:
                targetObject = _('object')
                if (self.series_num != 0) and (self.series_num != None):
                    targetObject = _('Series')
                elif (self.episode_num != 0) and (self.episode_num != None):
                    targetObject = _('Episode')
                elif (self.transcript_num != 0) and (self.transcript_num != None):
                    targetObject = _('Transcript')
                elif (self.collection_num != 0) and (self.collection_num != None):
                    targetObject = _('Collection')
                elif (self.clip_num != 0) and (self.clip_num != None):
                    targetObject = _('Clip')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Note named "%s" already exists for this %s.'), 'utf8')
                    targetObject = unicode(targetObject, 'utf8')
                else:
                    prompt = _('A Note named "%s" already exists for this %s.')                    
                raise SaveError, prompt % (self.id, targetObject)

            # Update the existing record
            query = """ UPDATE Notes2
                          SET NoteID = %s,
                              SeriesNum = %s,
                              EpisodeNum = %s,
                              CollectNum = %s,
                              ClipNum = %s,
                              TranscriptNum = %s,
                              NoteTaker = %s,
                              NoteText = %s
                          WHERE NoteNum = %s """ 
            values = values + (self.number,)

        c = DBInterface.get_db().cursor()
        c.execute(query, values)

        if self.number == 0:
            # If we are dealing with a brand new Clip, it does not yet know its
            # record number.  It HAS a record number, but it is not known yet.
            # The following query should produce the correct record number.
            query = """ SELECT NoteNum FROM Notes2
                          WHERE NoteID = %s AND
                                SeriesNum = %s AND
                                EpisodeNum = %s AND
                                CollectNum = %s AND
                                ClipNum = %s AND
                                TranscriptNum = %s """
            tempDBCursor = DBInterface.get_db().cursor()
            tempDBCursor.execute(query, (id, self.series_num, self.episode_num, self.collection_num, self.clip_num, self.transcript_num))
            if tempDBCursor.rowcount == 1:
                self.number = tempDBCursor.fetchone()[0]
            else:
                raise RecordNotFoundError, (self.id, tempDBCursor.rowcount)
            tempDBCursor.close()

        c.close()

        
    def db_delete(self, use_transactions=1):
        """Delete this object record from the database."""
        # Initialize delete operation.  Initiating a Transaction is never necessary here.
        (db, c) = self._db_start_delete(use_transactions)
        result = 1
        
        # Delete the actual record.
        self._db_do_delete(use_transactions, c, result)

        # Cleanup
        c.close()
        self.clear()

        return result

    def duplicate(self):
        """ Duplicate a Note """
        # Inherit the Dataobject duplicate method, which duplicates the note but strips the note number
        newNote = DataObject.DataObject.duplicate(self)
        # Return the duplicate note
        return newNote
    
# Private methods

    def _load_row(self, r):
        self.number = r['NoteNum']
        self.id = r['NoteID']
        self.comment = 'This is a note object'
        self.series_num = r['SeriesNum']
        self.episode_num = r['EpisodeNum']
        self.collection_num = r['CollectNum']
        self.clip_num = r['ClipNum']
        self.transcript_num = r['TranscriptNum']
        self.author = r['NoteTaker']

        # self.text = r['NoteText']
        # Okay, this isn't so straight-forward any more.
        # With MySQL for Python 0.9.x, r['NoteText'] is of type str.
        # With MySQL for Python 1.2.0, r['NoteText'] is of type array.  It could then either be a
        # character string (typecode == 'c') or a unicode string (typecode == 'u'), which then
        # need to be interpreted differently.
        if type(r['NoteText']).__name__ == 'array':
            if r['NoteText'].typecode == 'u':
                self.text = r['NoteText'].tounicode()
            else:
                self.text = r['NoteText'].tostring()
        else:
            self.text = r['NoteText']
        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.comment = DBInterface.ProcessDBDataForUTF8Encoding(self.comment)
            self.author = DBInterface.ProcessDBDataForUTF8Encoding(self.author)
            self.text = DBInterface.ProcessDBDataForUTF8Encoding(self.text)

    def _set_series(self, num):
        self._series = num
    def _get_series(self):
        return self._series
    def _del_series(self):
        self._series = 0

    def _set_episode(self, num):
        self._episode = num
    def _get_episode(self):
        return self._episode
    def _del_episode(self):
        self._episode = 0

    def _set_collection(self, num):
        self._collection = num
    def _get_collection(self):
        return self._collection
    def _del_collection(self):
        self._collection = 0


    def _set_clip(self, num):
        self._clip = num
    def _get_clip(self):
        return self._clip
    def _del_clip(self):
        self._clip = 0

    def _set_transcript(self, num):
        self._transcript = num
    def _get_transcript(self):
        return self._transcript
    def _del_transcript(self):
        self._transcript = 0

    def _get_notetype(self):
        notetype = None
        if self.transcript_num > 0:
            notetype = 'Transcript'
        elif self.episode_num > 0:
            notetype = 'Episode'
        elif self.series_num > 0:
            notetype = 'Series'
        elif self.clip_num > 0:
            notetype = 'Clip'
        elif self.collection_num > 0:
            notetype = 'Collection'
        return notetype

    def _set_author(self, name):
        self._author = name
    def _get_author(self):
        return self._author
    def _del_author(self):
        self._author = ""

    def _set_text(self, t):
        self._text = t
    def _get_text(self):
        return self._text
    def _del_text(self):
        if 'unicode' in wx.PlatformInfo:
            self._text = u""
        else:
            self._text = ""


# Public properties
    series_num = property(_get_series, _set_series, _del_series,
                        """Series number attached to (if applicable)""")
    episode_num = property(_get_episode, _set_episode, _del_episode,
                        """Episode number attached to (if applicable)""")
    collection_num = property(_get_collection, _set_collection, _del_collection,
                        """Collection number to which the note belongs.""")
    clip_num = property(_get_clip, _set_clip, _del_clip,
                        """Clip number attached to (if applicable)""")
    transcript_num = property(_get_transcript, _set_transcript, _del_transcript,
                        """Number of the transcript from which this Note was taken.""")
    notetype = property(_get_notetype, doc=""" Type of Note (read-only) """)
    author = property(_get_author, _set_author, _del_author,
                        """Person responsible for creating the Note.""")
    text = property(_get_text, _set_text, _del_text,
                        """Text of the note, stored in the database as a BLOB.""")
