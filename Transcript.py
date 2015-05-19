# Copyright (C) 2003 - 2006 The Board of Regents of the University of Wisconsin System 
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

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>, Jonathan Beavers <jonathan.beavers@gmail.com>'


DEBUG = False
if DEBUG:
    print "Transcript DEBUG is ON!"

import wx
import TransanaConstants
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
        str = str + "LastSaveTime = %s\n" % self.lastsavetime
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
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            name = name.encode(TransanaGlobal.encoding)
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
# This query doesn't work for loading a Clip Transcript, of course!
#        query = """SELECT a.*, b.EpisodeID, c.SeriesID FROM Transcripts2 a, Episodes2 b, Series2 c
#            WHERE   TranscriptNum = %s AND
#                    a.EpisodeNum = b.EpisodeNum AND
#                    b.SeriesNum = c.SeriesNum"""
        query = """SELECT * FROM Transcripts2 WHERE   TranscriptNum = %s"""
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

        if DEBUG:
            print "Transcript.db_save():  %s\n%s" % (self.text, type(self.text))
        
        # If we're in Unicode mode, ...
        if 'unicode' in wx.PlatformInfo:
            # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
            # variables for the data.  We don't want to change the underlying object values.  Also, this way,
            # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
            id = self.id.encode(TransanaGlobal.encoding)
            # If the transcriber is None in the database, we can't encode that!  Otherwise, we should encode.
            if self.transcriber != None:
                transcriber = self.transcriber.encode(TransanaGlobal.encoding)
            else:
                transcriber = self.transcriber
            # If the comment is None in the database, we can't encode that!  Otherwise, we should encode.
            if self.comment != None:
                comment = self.comment.encode(TransanaGlobal.encoding)
            else:
                comment = self.comment
        else:
            # If we don't need to encode the string values, we still need to copy them to our local variables.
            id = self.id
            transcriber = self.transcriber
            comment = self.comment

        fields = ("TranscriptID", "EpisodeNum", "ClipNum", "Transcriber", \
                        "RTFText", "Comment", "LastSaveTime")
        values = (id, self.episode_num, self.clip_num, transcriber, \
                    self.text, comment)

        if (self._db_start_save() == 0):
            # Duplicate Transcript IDs within an Episode are not allowed.
            if DBInterface.record_match_count("Transcripts2", \
                    ("TranscriptID", "EpisodeNum", "ClipNum"),
                    (id, self.episode_num, self.clip_num)) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Transcript named "%s" already exists in this Episode.\nPlease enter a different Transcript ID.'), 'utf8')
                else:
                    prompt = _('A Transcript named "%s" already exists in this Episode.\nPlease enter a different Transcript ID.')
                raise SaveError, prompt % self.id

            # insert the new record
            query = "INSERT INTO Transcripts2\n("
            for field in fields:
                query = "%s%s," % (query, field)
            query = query[:-1] + ')'
            query = "%s\nVALUES\n(" % query
            for value in values:
                query = "%s%%s," % query
            # The last data value should be the SERVER's time stamp because we don't know if the clients are synchronized.
            # Even a couple minutes difference can cause problems, but with time zones, the different could be hours!
            query += 'CURRENT_TIMESTAMP)'
        else:
            # Duplicate Transcript IDs within an Episode are not allowed.
            if DBInterface.record_match_count("Transcripts2", \
                    ("TranscriptID", "!TranscriptNum", "EpisodeNum", "ClipNum"),
                    (id, self.number, self.episode_num, self.clip_num)) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Transcript named "%s" already exists in this Episode.\nPlease enter a different Transcript ID.'), 'utf8')
                else:
                    prompt = _('A Transcript named "%s" already exists in this Episode.\nPlease enter a different Transcript ID.')
                raise SaveError, prompt % self.id
            
            # OK to update the episode record
            query = """UPDATE Transcripts2
                SET TranscriptID = %s,
                    EpisodeNum = %s,
                    ClipNum = %s,
                    Transcriber = %s,
                    RTFText = %s,
                    Comment = %s,
                    LastSaveTime = CURRENT_TIMESTAMP
                WHERE TranscriptNum = %s
            """
            values = values + (self.number,)

        if DEBUG:
            import Dialogs
            msg = query % values
            dlg = Dialogs.InfoDialog(None, msg)
            dlg.ShowModal()
            dlg.Destroy()

        c = DBInterface.get_db().cursor()
        c.execute(query, values)
                              
        # Load the auto-assigned new number record if necessary and the saved time.
        query = """
                  SELECT TranscriptNum, LastSaveTime FROM Transcripts2
                  WHERE TranscriptID = %s AND
                        EpisodeNum = %s AND
                        ClipNum = %s
                """
        tempDBCursor = DBInterface.get_db().cursor()
        tempDBCursor.execute(query, (id, self.episode_num, self.clip_num))
        if tempDBCursor.rowcount == 1:
            recs = tempDBCursor.fetchone()
            if (self.number == 0):
                self.number = recs[0]
            self.lastsavetime = recs[1]
        else:
            raise RecordNotFoundError, (self.id, tempDBCursor.rowcount)
        tempDBCursor.close()

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database."""
        result = 1
        try:
            # Initialize delete operation, begin transaction if necessary
            (db, c) = self._db_start_delete(use_transactions)
            
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
        except RecordLockedError, e:
            # if a sub-record is locked, we may need to unlock the Transcript record (after rolling back the Transaction)
            if self.isLocked:
                # c (the database cursor) only exists if the record lock was obtained!
                # We must roll back the transaction before we unlock the record.
                c.execute("ROLLBACK")
                c.close()
                self.unlock_record()
            raise e    
        except:
            raise

        return result

    def lock_record(self):
        """ Override the DataObject Lock Method """
        # If we're using the single-user version of Transana, we just need to ...
        if not TransanaConstants.singleUserVersion:
            # ... confirm that the transcript has not been altered by another user since it was loaded.
            # To do this, first pull the LastSaveTime for this record from the database.
            query = """
                      SELECT LastSaveTime FROM Transcripts2
                      WHERE TranscriptNum = %s
                    """
            tempDBCursor = DBInterface.get_db().cursor()
            tempDBCursor.execute(query % self.number)
            if tempDBCursor.rowcount == 1:
                newLastSaveTime = tempDBCursor.fetchone()[0]
            else:
                raise RecordNotFoundError, (self.id, tempDBCursor.rowcount)
            tempDBCursor.close()
            if newLastSaveTime != self.lastsavetime:
                self.db_load_by_num(self.number)
        
        # ... lock the Transcript Record
        DataObject.lock_record(self)
            
    def unlock_record(self):
        """ Override the DataObject Unlock Method """
        # Unlock the Transcript Record
        DataObject.unlock_record(self)


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

    # Implementation for LastSaveTime Property
    def _get_lastsavetime(self):
        return self._lastsavetime
    def _set_lastsavetime(self, lst):
        self._lastsavetime = lst
    def _del_lastsavetime(self):
        self._lastsavetime = None

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
    lastsavetime = property(_get_lastsavetime, _set_lastsavetime, _del_lastsavetime,
                        """The timestamp of the last save (MU only).""")
        
    def _load_row(self, row):
    	self.number = row['TranscriptNum']
        self.id = row['TranscriptID']
        self.episode_num = row['EpisodeNum']
        self.clip_num = row['ClipNum']
        self.transcriber = row['Transcriber']

        # Can I get away with assuming Unicode?
        # Here's the plan:
        #   test for rtf in here, if you find rtf, process normally
        #   if you don't find it, pass data off to some weirdo method in TranscriptEditor.py

        # 1 - Determine encoding, adjust if needed
        # 2 - enact the plan above

        # determine encoding, fix if needed
        if type(row['RTFText']).__name__ == 'array':

            if DEBUG:
                print "Transcript._load_row(): 2", row['RTFText'].typecode
            
            if row['RTFText'].typecode == 'u':
                self.text = row['RTFText'].tounicode()
            else:
                self.text = row['RTFText'].tostring()
        else:
            self.text = row['RTFText']

        if 'unicode' in wx.PlatformInfo:
            if DEBUG:
                print "debug here"
		
            if type(self.text).__name__ == 'str':
                temp = self.text[2:5]

                # check to see if we're working with RTF
                try:
                    if temp == u'rtf':
                        # convert the data to unicode just to be safe.
                        self.text = unicode(self.text, 'utf-8')
                
                except UnicodeDecodeError:
                    # This would sometimes get called while I was using cPickle instead of Pickle.
                    # You could probably remove the exception handling stuff and be okay, but it's
                    # not hurting anything like it is.
                    self.dlg.editor.load_transcript(transcriptObj, 'pickle')

        # self.text gets set to be our data
        # then load_transcript is called, from transcriptionui.LoadTranscript()

        self.comment = row['Comment']
        self.lastsavetime = row['LastSaveTime']
        self.changed = False

        # These values come from the Series and Episode tables for Episode Transcripts, but do not exist for
        # Clip Transcripts.
        if row.has_key('SeriesID'):
            self.series_id = row['SeriesID']
        if row.has_key('EpisodeID'):
            self.episode_id = row['EpisodeID']

        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.transcriber = DBInterface.ProcessDBDataForUTF8Encoding(self.transcriber)
            self.comment = DBInterface.ProcessDBDataForUTF8Encoding(self.comment)
            if row.has_key('SeriesID'):
                self.series_id = DBInterface.ProcessDBDataForUTF8Encoding(self.series_id)
            if row.has_key('EpisodeID'):
                self.episode_id = DBInterface.ProcessDBDataForUTF8Encoding(self.episode_id)

