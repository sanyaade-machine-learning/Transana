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

"""This module implements the Transcript class as part of the Data Objects."""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

import wx
from mx.DateTime import *
import TransanaGlobal
from DataObject import DataObject
import DBInterface
import Note
from TransanaExceptions import *
import types

TIMECODE_CHAR = "\\'a4"   # Note that this differs from the TIMECODE_CHAR in TranscriptEditor.py
                          # because this is for RTF text and that is for parsed text.

class Transcript(DataObject):
    """This class defines the structure for a transcript object.  A transcript
    object describes a transcript document for Episodes or Clips."""

    def __init__(self, id_or_num=None, ep=None, clip=None):
        """Initialize an Transcript object."""
        # Transcripts can be loaded in 3 ways:
        #   Transcript Number can be provided                   (Loading any Transcript)
        #   Transcript Name and Episode Number can be provided  (Loading an Episode Transcript)
        #   Clip Number can be provided                         (Loading a Clip transcript)
        DataObject.__init__(self)
        if (id_or_num == None) and (clip != None):
            self.db_load_by_clipnum(clip)
        elif type(id_or_num) in (int, long):
            self.db_load_by_num(id_or_num)
        elif isinstance(id_or_num, types.StringTypes) and (type(ep) in (int, long)):
            self.db_load_by_name(id_or_num, ep)

# Public methods

    def __repr__(self):
        str = 'Transcript Object:\n'
        str = str + "Number = %s\n" % self.number
        str = str + "ID = %s\n" % self.id
        str = str + "Episode Number = %s\n" % self.episode_num
        str = str + "Clip Number = %s\n" % self.clip_num
        str = str + "Transcriber = %s\n" % self.transcriber
        str = str + "Comment = %s\n" % self.comment
        if len(self.text) > 250:
            str = str + "text not displayed due to length.\n\n"
        else:
            str = str + "text = %s\n\n" % self.text
        return str

    def GetTranscriptWithoutTimeCodes(self):
        """ Returns a copy of the Transcript Text with the Time Code information removed. """
        newText = self.text
            
        while True:
            timeCodeStart = newText.find(TIMECODE_CHAR)
            if timeCodeStart == -1:
                break
            timeCodeEnd = newText.find('>', timeCodeStart)
            newText = newText[:timeCodeStart] + newText[timeCodeEnd + 1:]

        # We should also replace TAB characters with spaces
        while True:
            tabStart = newText.find('\\tab', 0)
            if tabStart == -1:
                break
            newText = newText[:tabStart] + '    ' + newText[tabStart + 4:]
            # if the RTF delimiter for the \tab marker was a space, the space should be removed too!
            if newText[tabStart] == ' ':
                newText = newText[:tabStart] + newText[tabStart + 1:]
                

        return newText

    def db_load_by_name(self, name, episode):
        """Load a record by ID / Name."""
        db = DBInterface.get_db()
        query = """SELECT a.*, b.EpisodeID, c.SeriesID FROM Transcripts2 a, Episodes2 b, Series2 c
            WHERE   TranscriptID = %s AND
                    a.EpisodeNum = b.EpisodeNum AND
                    b.EpisodeNum = %s AND
                    b.SeriesNum = c.SeriesNum
        """
        c = db.cursor()
        c.execute(query, (name, episode))
        n = c.rowcount
        if (n != 1):
            c.close()
            self.clear()
            raise RecordNotFoundError(name, n)
        else:
            r = DBInterface.fetch_named(c)
            self._load_row(r)
        
        c.close()
    
    def db_load_by_num(self, num):
        """Load a record by record number."""
        db = DBInterface.get_db()
        query = """SELECT a.*, b.EpisodeID, c.SeriesID FROM Transcripts2 a, Episodes2 b, Series2 c
            WHERE   TranscriptNum = %s AND
                    a.EpisodeNum = b.EpisodeNum AND
                    b.SeriesNum = c.SeriesNum"""
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

    def db_load_by_clipnum(self, clip):
        """ Load a Transcript Record based on Clip Number """
        db = DBInterface.get_db()
        query = """SELECT * FROM Transcripts2 a
            WHERE   ClipNum = %s """
        c = db.cursor()
        c.execute(query, clip)
        n = c.rowcount
        if (n != 1):
            c.close()
            self.clear()
            raise RecordNotFoundError, (clip, n)
        else:
            r = DBInterface.fetch_named(c)
            self._load_row(r)
        
        c.close()

    def db_save(self):
        """Save the record to the database using Insert or Update as
        appropriate."""
        # Sanity checks
        if (self.id == "") and (self.clip_num == 0):
            raise SaveError, _("Transcript ID is required.")
        
        fields = ("TranscriptID", "EpisodeNum", "ClipNum", "Transcriber", \
                        "RTFText", "Comment", "LastSaveTime")
        values = (self.id, self.episode_num, self.clip_num, self.transcriber, \
                    self.text, self.comment, str(now())[:-3])

        if (self._db_start_save() == 0):
            # Duplicate Transcript IDs within an Episode are not allowed.
            if DBInterface.record_match_count("Transcripts2", \
                    ("TranscriptID", "EpisodeNum", "ClipNum"),
                    (self.id, self.episode_num, self.clip_num)) > 0:
                raise SaveError, _('A Transcript named "%s" already exists in this Episode.\nPlease enter a different Transcript ID.') % self.id

            # insert the new record
            query = "INSERT INTO Transcripts2\n("
            for field in fields:
                query = "%s%s," % (query, field)
            query = query[:-1] + ')'
            query = "%s\nVALUES\n(" % query
            for value in values:
                query = "%s%%s," % query
            query = query[:-1] + ')'
        else:
            # Duplicate Transcript IDs within an Episode are not allowed.
            if DBInterface.record_match_count("Transcripts2", \
                    ("TranscriptID", "!TranscriptNum", "EpisodeNum", "ClipNum"),
                    (self.id, self.number, self.episode_num, self.clip_num)) > 0:
                raise SaveError, _('A Transcript named "%s" already exists in this Episode.\nPlease enter a different Transcript ID.') % self.id
            
            # OK to update the episode record
            query = """UPDATE Transcripts2
                SET TranscriptID = %s,
                    EpisodeNum = %s,
                    ClipNum = %s,
                    Transcriber = %s,
                    RTFText = %s,
                    Comment = %s,
                    LastSaveTime = %s
                WHERE TranscriptNum = %s
            """
            values = values + (self.number,)

        c = DBInterface.get_db().cursor()
        c.execute(query, values)
                              
        if (self.number == 0):
            # Load the auto-assigned new number record
            # This method is no good.  Clip Transcripts don't have an id and the episode num is not good enough.
            # self.db_load_by_name(self.id, self.episode_num)

            # If we are dealing with a brand new Transcript, it does not yet know its
            # record number.  It HAS a record number, but it is not known yet.
            # The following query should produce the correct record number for both
            # Episode and Clip Transcripts.
            query = """
                      SELECT TranscriptNum FROM Transcripts2
                      WHERE TranscriptID = %s AND
                            EpisodeNum = %s AND
                            ClipNum = %s
                    """
            tempDBCursor = DBInterface.get_db().cursor()
            tempDBCursor.execute(query, (self.id, self.episode_num, self.clip_num))
            if tempDBCursor.rowcount == 1:
                self.number = tempDBCursor.fetchone()[0]
            else:
                raise RecordNotFoundError, (self.id, tempDBCursor.rowcount)
            tempDBCursor.close()

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database."""
        # Initialize delete operation, begin transaction if necessary
        (db, c) = self._db_start_delete(use_transactions)
        result = 1
        
        # Detect, Load, and Delete all Transcript Notes.
        notes = self.get_note_nums()
        for note_num in notes:
            note = Note.Note(note_num)
            result = result and note.db_delete(0)
            del note
        del notes

        # Delete the actual record.
        self._db_do_delete(use_transactions, c, result)

        # Cleanup
        c.close()
        self.clear()

        return result


# Private methods
    # Implementation for Episode Number Property
    def _get_ep_num(self):
        return self._ep_num
    def _set_ep_num(self, num):
        self._ep_num = num
    def _del_ep_num(self):
        self._ep_num = 0

    # Implementation for Clip Number Property
    def _get_cl_num(self):
        return self._cl_num
    def _set_cl_num(self, num):
        self._cl_num = num
    def _del_cl_num(self):
        self._cl_num = 0

    # Implementation for Transcriber Property
    def _get_transcriber(self):
        return self._transcriber
    def _set_transcriber(self, name):
        self._transcriber = name
    def _del_transcriber(self):
        self._transcriber = ''

    # Implementation for Text Property
    def _get_text(self):
        return self._text
    def _set_text(self, txt):
        self._text = txt
    def _del_text(self):
        self._text = ''

    # Implementation for Has_Changed Property
    def _get_changed(self):
        return self._has_changed
    def _set_changed(self, changed):
        self._has_changed = changed
    def _del_changed(self):
        self._has_changed = False

# Public properties
    episode_num = property(_get_ep_num, _set_ep_num, _del_ep_num,
                        """The Episode number, if associated with one.""")
    clip_num = property(_get_cl_num, _set_cl_num, _del_cl_num,
                        """The clip number, if associated with one.""")
    transcriber = property(_get_transcriber, _set_transcriber, _del_transcriber,
                        """The person who created the Transcript.""")
    text = property(_get_text, _set_text, _del_text,
                        """Text of the transcript, stored in the database as a BLOB.""")
    locked_by_me = property(None, None, None,
                        """Determines if this instance owns the Transcript lock.""")
    has_changed = property(_get_changed, _set_changed, _del_changed,
                        """Indicates whether the Transcript has been modified.""")
    last_save_time = property(None, None, None,
                        """The timestamp of the last save (MU only).""")
    
    
    def _load_row(self, r):
        self.number = r['TranscriptNum']
        self.id = r['TranscriptID']
        self.episode_num = r['EpisodeNum']
        self.clip_num = r['ClipNum']
        self.transcriber = r['Transcriber']
        self.text = r['RTFText']
        self.comment = r['Comment']
        self.changed = False

        # These values come from the Series and Episode tables for Episode Transcripts, but do not exist for
        # Clip Transcripts.
        if r.has_key('SeriesID'):
            self.series_id = r['SeriesID']
        if r.has_key('EpisodeID'):
            self.episode_id = r['EpisodeID']
