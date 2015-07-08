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

"""This module implements the Transcript class as part of the Data Objects."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case, Jonathan Beavers <jonathan.beavers@gmail.com>'

DEBUG = False
if DEBUG:
    print "Transcript DEBUG is ON!"

# import wxPython
import wx
# import Python's os module
import os
# import Python's types module
import types
# import Transana's Dialogs, required for an error message
import DataObject
# import Transana's Database Interface
import DBInterface
# import Transana's Dialogs
import Dialogs
# import Transana's Miscellaneous Functions
import Misc
# import Transana's Note Object
import Note
# import Transana's Constants
import TransanaConstants
# import Transana's Exceptions
from TransanaExceptions import *
# import Transana's Globals
import TransanaGlobal

TIMECODE_CHAR = "\\'a4"   # Note that this differs from the TIMECODE_CHAR in TranscriptEditor.py
                          # because this is for RTF text and that is for parsed text.

class Transcript(DataObject.DataObject):
    """This class defines the structure for a transcript object.  A transcript
    object describes a transcript document for Episodes or Clips."""

    def __init__(self, id_or_num=None, ep=None, clip=None, skipText=False):
        """Initialize an Transcript object."""
        # Transcripts can be loaded in 3 ways:
        #   Transcript Number can be provided                   (Loading any Transcript)
        #   Transcript Name and Episode Number can be provided  (Loading an Episode Transcript)
        #   Clip Number can be provided                         (Loading a Clip transcript)
        #   skipText indicates that the RTFText can be left off.  This leads to significantly faster transcript loading
        #     particularly when we start having large transcripts with embedded images.
        DataObject.DataObject.__init__(self)
        # Remember if we're supposed to skip the RTF Text
        self.skipText = skipText
        if (id_or_num == None) and (clip != None):
            self.db_load_by_clipnum(clip)
        elif type(id_or_num) in (int, long):
            self.db_load_by_num(id_or_num)
        elif isinstance(id_or_num, types.StringTypes) and (type(ep) in (int, long)):
            self.db_load_by_name(id_or_num, ep)
        # For Partial Transcript Editing, create a data structure for storing the transcript information by LINE
        self.lines = []
        # Initialize a paragraph counter
        self.paragraphs = 0
        # Create a data structure for tracking very large transcripts by section
        self.paragraphPointers = {}
        # If we have text in the transcript ...
        if id_or_num != None:
            # ... set up data structures needed for editing large paragraphs
            self.UpdateParagraphs()

# Public methods

    def __repr__(self):
        str = 'Transcript Object:\n'
        str = str + "Number = %s\n" % self.number
        str = str + "ID = %s\n" % self.id
        str = str + "Episode Number = %s\n" % self.episode_num
        str = str + "Source Transcript = %s\n" % self.source_transcript
        str = str + "Clip Number = %s\n" % self.clip_num
        str = str + "Sort Order = %s\n" % self.sort_order
        str = str + "Transcriber = %s\n" % self.transcriber
        str = str + "Clip Start = %s\n" % Misc.time_in_ms_to_str(self.clip_start)
        str = str + "Clip Stop = %s\n" % Misc.time_in_ms_to_str(self.clip_stop)
        str = str + "Comment = %s\n" % self.comment
        str += "MinTranscriptWidth = %s\n" % self.minTranscriptWidth
        str = str + "LastSaveTime = %s\n" % self.lastsavetime
        str += "isLocked = %s\n" % self._isLocked
        str += "recordlock = %s\n" % self.recordlock
        str += "locktime = %s\n" % self.locktime
        if len(self.text) > 250:
            str = str + "text not displayed due to length.\n\n"
        else:
            str = str + "text = %s\n\n" % self.text
        return str.encode('utf8')

    def __eq__(self, other):
        """ Object Equality function """
        if other == None:
            return False
        else:

##            print "Transcript.__eq__():", len(self.__dict__.keys()), len(other.__dict__.keys())
##
##            print self.__dict__.keys()
##            print other.__dict__.keys()
##            print
##            
##            for key in self.__dict__.keys():
##                print key, self.__dict__[key] == other.__dict__[key],
##                if self.__dict__[key] != other.__dict__[key]:
##                    print self.__dict__[key], other.__dict__[key]
##                else:
##                    print
##            print
            
            return self.__dict__ == other.__dict__

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
        # Get the database connection
        db = DBInterface.get_db()
        # If we're skipping the RTFText ...
        if self.skipText:
            # Define the query to load a Clip Transcript without text
            query = """SELECT a.TranscriptNum, a.TranscriptID, a.EpisodeNum, a.SourceTranscriptNum,
                              a.ClipNum, a.SortOrder, a.Transcriber, a.ClipStart, a.ClipStop, a.Comment,
                              a.MinTranscriptWidth, a.RecordLock, a.LockTime, a.LastSaveTime,
                              b.EpisodeID, c.SeriesID FROM Transcripts2 a, Episodes2 b, Series2 c
                WHERE   TranscriptID = %s AND
                        a.EpisodeNum = b.EpisodeNum AND
                        b.EpisodeNum = %s AND
                        b.SeriesNum = c.SeriesNum
            """
        # If we're NOT skipping the RTF Text
        else:
            # Define the query to load a Clip Transcript with everything
            query = """SELECT a.*, b.EpisodeID, c.SeriesID FROM Transcripts2 a, Episodes2 b, Series2 c
                WHERE   TranscriptID = %s AND
                        a.EpisodeNum = b.EpisodeNum AND
                        b.EpisodeNum = %s AND
                        b.SeriesNum = c.SeriesNum
            """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = db.cursor()
        # Execute the Load query
        c.execute(query, (name, episode))
        # rowcount doesn't work for sqlite!
        if TransanaConstants.DBInstalled == 'sqlite3':
            n = 1
        else:
            n = c.rowcount
        # If we don't have exactly ONE record ...
        if (n != 1):
            # ... close the database cursor ...
            c.close()
            # ... clear the current Transcript Object ...
            self.clear()
            # ... and raise an exception
            raise RecordNotFoundError(name, n)
        else:
            # ... Fetch the query results ...
            r = DBInterface.fetch_named(c)
            # If sqlite and not results are found ...
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}): 
                # ... close the database cursor ...
                c.close()
                # ... clear the current Transcript object ...
                self.clear()
                # ... and raise an exception
                raise RecordNotFoundError, (name, 0)
            # Load the data into the Transcript Object
            self._load_row(r)
        # Close the database cursor
        c.close()
    
    def db_load_by_num(self, num):
        """Load a record by record number."""
        # Get the database connection
        db = DBInterface.get_db()
        # If we're skipping the RTF Text ...
        if self.skipText:
            # Define the query to load a Clip Transcript without text
            query = """SELECT TranscriptNum, TranscriptID, EpisodeNum, SourceTranscriptNum,
                              ClipNum, SortOrder, Transcriber, ClipStart, ClipStop, Comment,
                              MinTranscriptWidth, RecordLock, LockTime, LastSaveTime
                         FROM Transcripts2 WHERE   TranscriptNum = %s
                    """
        # If we're NOT skipping the RTF Text ...
        else:
            # Define the query to load a Clip Transcript with everything
            query = """SELECT * FROM Transcripts2 WHERE   TranscriptNum = %s"""
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = db.cursor()
        # Execute the Load query
        c.execute(query, (num, ))
        # rowcount doesn't work for sqlite!
        if TransanaConstants.DBInstalled == 'sqlite3':
            n = 1
        else:
            n = c.rowcount
        # If we don't have exactly ONE record ...
        if (n != 1):
            # ... close the database cursor ...
            c.close()
            # ... clear the current Transcript Object ...
            self.clear()
            # ... and raise an exception
            raise RecordNotFoundError, (num, n)
        # If we DO have exactly one record (or use sqlite) ...
        else:
            # ... Fetch the query results ...
            r = DBInterface.fetch_named(c)
            # If sqlite and not results are found ...
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}): 
                # ... close the database cursor ...
                c.close()
                # ... clear the current Transcript object ...
                self.clear()
                # ... and raise an exception
                raise RecordNotFoundError, (num, 0)
            # Load the data into the Transcript Object
            self._load_row(r)
        # Close the database cursor
        c.close()

    def db_load_by_clipnum(self, clip):
        """ Load a Transcript Record based on Clip Number """
        # Get the database connection
        db = DBInterface.get_db()
        # Define the query to load a Clip Transcript
        query = """SELECT * FROM Transcripts2 a
            WHERE   ClipNum = %s """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = db.cursor()
        # Execute the Load query
        c.execute(query, (clip, ))
        # rowcount doesn't work for sqlite!
        if TransanaConstants.DBInstalled == 'sqlite3':
            n = 1
        else:
            n = c.rowcount
        # If we don't have exactly ONE record ...
        if (n != 1):
            # ... close the database cursor ...
            c.close()
            # ... clear the current Transcript Object ...
            self.clear()
            # ... and raise an exception
            raise RecordNotFoundError, (clip, n)
        # If we DO have exactly one record (or use sqlite) ...
        else:
            # ... Fetch the query results ...
            r = DBInterface.fetch_named(c)
            # If sqlite and not results are found ...
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}):
                # ... close the database cursor ...
                c.close()
                # ... clear the current Transcript object ...
                self.clear()
                # ... and raise an exception
                raise RecordNotFoundError, (num, 0)
            # Load the data into the Transcript Object
            self._load_row(r)
        # Close the database cursor
        c.close()

    def UpdateParagraphs(self):
        """ This method divides XML text up into paragraphs, needed for editing LONG transcripts """
        # Initialize (or re-initialize) the paragraph pointers dictionary
        self.paragraphPointers = {}
        # If there's a defined (saved) transcript object ...
        if (self.id != 0) and (self.text != None):
            
#            print "Transcript.UpdateParagraphs()"

            # ... and if the text is in XML form ...
            if self.text[:5] == '<?xml':
                # ... divide the transcript text into individual LINES
                self.lines = self.text.split('\n')
                # Initialize the paragraph counter
                self.paragraphs = 0
                # Iterate through the lines
                for x in range(0, len(self.lines)):
                    # Every 100 paragraphs ...
                    if self.paragraphs % 100 == 0:
                        # ... set a pointer to the line that starts the paragraph
                        self.paragraphPointers[self.paragraphs] = x
                    # # If the line contains the Paragraph XML tag, ...
                    if ("<paragraph " in self.lines[x]) or ("<paragraph>" in self.lines[x]):
                        # ... increment the Paragraph Counter
                        self.paragraphs += 1

#                print "characters:", len(self.text)
#                print "lines:", len(self.lines)
#                print "paragraphs:", self.paragraphs
#            print

    def db_save(self, use_transactions=True):
        """Save the record to the database using Insert or Update as
        appropriate."""

        # Define and implement Demo Version limits
        if TransanaConstants.demoVersion and (self.number == 0) and (self.clip_num == 0):
            # Get a DB Cursor
            c = DBInterface.get_db().cursor()
            # Find out how many Episode Transcript records exist (not counting Clip Transcripts)
            c.execute('SELECT COUNT(TranscriptNum) FROM Transcripts2 WHERE ClipNum = 0')
            res = c.fetchone()
            c.close()
            # Define the maximum number of recors allowed
            maxEpisodeTranscripts = TransanaConstants.maxEpisodeTranscripts
            # Compare
            if res[0] >= maxEpisodeTranscripts:
                # If the limit is exceeded, create and display the error using a SaveError exception
                prompt = _('The Transana Demonstration limits you to %d Transcript records.\nPlease cancel the "Add Transcript" dialog to continue.')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                raise SaveError, prompt % maxEpisodeTranscripts

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

        # If we have a NEW transcript and it HAS an ID and it is EMPTY ...
        if (self.number == 0) and (self.id != "") and (len(self.text) == 0):
            # ... build the Video Root default.xml file name.
            fname = os.path.join(TransanaGlobal.configData.videoPath, 'default.xml')
            # If the XML default doesn't exist ...
            if not os.path.exists(fname):
                # ... build the Video Root default.rtf file name.
                fname = os.path.join(TransanaGlobal.configData.videoPath, 'default.rtf')

            # If the default file exists ...
            if os.path.exists(fname):
                # ... start exception handling ...
                try:
                    # Open the file
                    f = open(fname, "r")
                    # Read the file straight into the Transcript Text
                    self.text = f.read()
                    # if the text does NOT have an RTF or XML header ...
                    if (self.text[:5].lower() != '{\\rtf') and (self.text[:5].lower() != '<?xml'):
                        # ... then it's NOT a legal default transcript and we shouldn't use it.
                        self.obj.text = ''
                    # Close the file
                    f.close()
                # If exceptions are raised ...
                except:
                    # ... we don't need to do anything here.  (Error message??)
                    # The consequence is probably that the Transcript Text will be blank.
                    pass

        if (len(self.text) > TransanaGlobal.max_allowed_packet):   # 8388000
            raise SaveError, _("This transcript is too large for the database.  Please shorten it, split it into two parts\nor if you are importing an RTF document, remove some unnecessary RTF encoding.")

        # Make a minor adjustment to the data, if needed.  (This prevents an error in Database Import.)
        if self.source_transcript == 'None':
            self.source_transcript = 0

        fields = ("TranscriptID", "EpisodeNum", "SourceTranscriptNum", "ClipNum", "SortOrder", "Transcriber", \
                        "ClipStart", "ClipStop", "RTFText", "Comment", "MinTranscriptWidth", "LastSaveTime")
        values = (id, self.episode_num, self.source_transcript, self.clip_num, self.sort_order, transcriber, \
                    self.clip_start, self.clip_stop, self.text, comment, self.minTranscriptWidth)

        if (self._db_start_save() == 0):
            # Duplicate Transcript IDs within an Episode are not allowed.
            if DBInterface.record_match_count("Transcripts2", \
                    ("TranscriptID", "EpisodeNum", "ClipNum", "SortOrder"),
                    (id, self.episode_num, self.clip_num, self.sort_order)) > 0:
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
                    ("TranscriptID", "!TranscriptNum", "EpisodeNum", "ClipNum", "SortOrder"),
                    (id, self.number, self.episode_num, self.clip_num, self.sort_order)) > 0:
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
                    SourceTranscriptNum = %s, 
                    ClipNum = %s,
                    SortOrder = %s,
                    Transcriber = %s,
                    ClipStart = %s,
                    ClipStop = %s,
                    RTFText = %s,
                    Comment = %s,
                    MinTranscriptWidth = %s,
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
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = DBInterface.get_db().cursor()
        # Execure the Save query
        c.execute(query, values)
        # If the object number is 0, we have a new object
        if self.number == 0:
            # Load the auto-assigned new number record if necessary and the saved time.
            query = """
                      SELECT TranscriptNum, LastSaveTime FROM Transcripts2
                      WHERE TranscriptID = %s AND
                            EpisodeNum = %s AND
                            ClipNum = %s
                    """
            # Assemble the arguments for the query
            args = (id, self.episode_num, self.clip_num)
            # If we've got a sort-order ...
            if self.sort_order != None:
                # ... include that as well!
                query += " AND SortOrder = %s"
                args += (self.sort_order,)
            # Get a temporarly database cursor
            tempDBCursor = DBInterface.get_db().cursor()
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Execute the query
            tempDBCursor.execute(query, args)
            # Get the results from the query
            recs = tempDBCursor.fetchall()
            # If we have exactly one record ...
            if len(recs) == 1:
                # ... load the record number ...
                self.number = recs[0][0]
                # ... and update the LastSaveTime
                self.lastsavetime = recs[0][1]
            # If we don't have a single record ...
            else:
                # ... raise an exception
                raise RecordNotFoundError, (self.id, len(recs))
            # Close the temporary database cursor
            tempDBCursor.close()
        # If we're updating an EXISTING record ... 
        else:
            # Load the NEW last saved time.
            query = """
                      SELECT LastSaveTime FROM Transcripts2
                      WHERE TranscriptNum = %s
                    """
            # Assemble the arguments for the query
            args = (self.number, )
            # Get a temporarly database cursor
            tempDBCursor = DBInterface.get_db().cursor()
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Execute the query
            tempDBCursor.execute(query, args)
            # Get the results from the query
            recs = tempDBCursor.fetchall()
            # If we have exactly one record ...
            if len(recs) == 1:
                # ... load the LastSaveTime
                self.lastsavetime = recs[0][0]
            # If we don't have a single record ...
            else:
                # ... raise an exception
                raise RecordNotFoundError, (self.id, len(recs))
            # Close the temporary database cursor
            tempDBCursor.close()

        # For Partial Transcript Editing, update the Paragraph Information for long transcripts
        self.UpdateParagraphs()

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database."""
        result = 1
        try:
            # Initialize delete operation, begin transaction if necessary
            (db, c) = self._db_start_delete(use_transactions)
            # Determine if any Clips have been created from THIS Transcript.
            clips = DBInterface.list_of_clips_by_transcriptnum(self.number)
            # If there are Clips in the list ...
            if len(clips) > 0:
                # ... build a prompt for the warning dialog box
                prompt = _('Clips have been created from Transcript "%s".\nThese clips will become orphaned if you delete the transcript.\nDo you want to delete this transcript anyway?')
                # Adjust the prompt for Unicode if needed
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                # Display the prompt for the user.  We do not want "Yes" as the default!
                tempDlg = Dialogs.QuestionDialog(TransanaGlobal.menuWindow, prompt % self.id, noDefault=True)
                result = tempDlg.LocalShowModal()
                tempDlg.Destroy()
                # If the user indicated they did NOT want to delete the Transcript ...
                if result == wx.ID_NO:
                    # ... build the appropriate prompt ...
                    prompt = _('The delete has been cancelled.  Clips have been created from Transcript "%s".')
                    # ... encode the prompt if needed ...
                    if 'unicode' in wx.PlatformInfo:
                        prompt = unicode(prompt, 'utf8')
                    # ... and use a DeleteError Exception to interupt the deletion.
                    raise DeleteError(prompt % self.id)



            # We need a user confirmation when Snapshot will be orphaned.



            # Determine if any Snapshots are linked to THIS Transcript
            # snapshots = DBInterface.list_of_snapshots_by_transcriptnum(self.number)
            
            
            # Detect, Load, and Delete all Transcript Notes.
            notes = self.get_note_nums()
            for note_num in notes:
                note = Note.Note(note_num)
                result = result and note.db_delete(0)
                del note
            del notes

            # Episode Transcripts with this transcript's number need to be cleared from Clip Transcripts
            # SourceTranscriptNumber records.  Snapshots with this Transcript's number  need to be cleared
            # too.  This must be done LAST, after everything else.
            if result and (self.clip_num == 0):
                DBInterface.ClearSourceTranscriptRecords(self.number)

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
        # Handle the DeleteError Exception
        except DeleteError, e:
            # if a sub-record is locked, we may need to unlock the Transcript record (after rolling back the Transaction)
            if self.isLocked:
                # c (the database cursor) only exists if the record lock was obtained!
                # We must roll back the transaction before we unlock the record.
                c.execute("ROLLBACK")
                # Close the database cursor
                c.close()
                # unlock the record
                self.unlock_record()
            # Pass on the exception
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
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Get a database cursor
            tempDBCursor = DBInterface.get_db().cursor()
            # Execute the Lock query
            tempDBCursor.execute(query, (self.number, ))
            # Get the query results
            rows = tempDBCursor.fetchall()
            # If we get exactly one row (and we should) ...
            if len(rows) == 1:
                # ... get the LastSaveTime
                newLastSaveTime = rows[0]
            # Otherwise ...
            else:
                # ... raise an exception
                raise RecordNotFoundError, (self.id, len(rows))
            # Close the database cursor
            tempDBCursor.close()
            # if the LastSaveTime has changed, some other user has altered the record since we loaded it.
            if newLastSaveTime != self.lastsavetime:
                # ... so we need to re-load it!
                self.db_load_by_num(self.number)
        
        # ... lock the Transcript Record
        DataObject.DataObject.lock_record(self)
            
    def unlock_record(self):
        """ Override the DataObject Unlock Method """
        # Unlock the Transcript Record
        DataObject.DataObject.unlock_record(self)


# Private methods
    # Implementation for Episode Number Property
    def _get_ep_num(self):
        return self._ep_num
    def _set_ep_num(self, num):
        self._ep_num = num
    def _del_ep_num(self):
        self._ep_num = 0

    # Implementation of the Source Transcript Property
    def _get_source_transcript(self):
        return self._source_transcript
    def _set_source_transcript(self, num):
        self._source_transcript = num
    def _del_source_transcript(self):
        self._source_transcript = 0

    # Implementation for Clip Number Property
    def _get_cl_num(self):
        return self._cl_num
    def _set_cl_num(self, num):
        self._cl_num = num
    def _del_cl_num(self):
        self._cl_num = 0

    # Implementation of the Sort Order Property
    def _get_sort_order(self):
        return self._sort_order
    def _set_sort_order(self, num):
        self._sort_order = num
    def _del_sort_order(self):
        self._sort_order = 0

    # Implementation for Transcriber Property
    def _get_transcriber(self):
        return self._transcriber
    def _set_transcriber(self, name):
        self._transcriber = name
    def _del_transcriber(self):
        self._transcriber = ''

    # Implementation for Clip Start Property
    def _get_clip_start(self):
        return self._clip_start
    def _set_clip_start(self, clipStart):
        self._clip_start = clipStart
    def _del_clip_start(self):
        self._clip_start = 0

    # Implementation for Clip Stop Property
    def _get_clip_stop(self):
        return self._clip_stop
    def _set_clip_stop(self, clipStop):
        self._clip_stop = clipStop
    def _del_clip_stop(self):
        self._clip_stop = 0

    # Implementation for Minimum Transcript Width
    def _get_min_transcript_width(self):
        return self._minTranscriptWidth
    def _set_min_transcript_width(self, minTrWidth):
        if minTrWidth == None:
            self._minTranscriptWidth = 0
        else:
            self._minTranscriptWidth = minTrWidth
    def _del_min_transcript_width(self):
        self._minTranscriptWidth = 0

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
    source_transcript = property(_get_source_transcript, _set_source_transcript, _del_source_transcript,
                                 """The Episode Transcript number a Clip Transcript was taken from""")
    clip_num = property(_get_cl_num, _set_cl_num, _del_cl_num,
                        """The clip number, if associated with one.""")
    sort_order = property(_get_sort_order, _set_sort_order, _del_sort_order,
                                 """The Sort Order for Transcript""")
    transcriber = property(_get_transcriber, _set_transcriber, _del_transcriber,
                        """The person who created the Transcript.""")
    clip_start = property(_get_clip_start, _set_clip_start, _del_clip_start,
                          """Clip Start Time, only used for multi-transcript clips for propagating Transcript changes""")
    clip_stop = property(_get_clip_stop, _set_clip_stop, _del_clip_stop,
                          """Clip Stop Time, only used for multi-transcript clips for propagating Transcript changes""")
    minTranscriptWidth = property(_get_min_transcript_width, _set_min_transcript_width, _del_min_transcript_width,
                                  """ Minimim Transcript Display Width """)
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
        self.source_transcript = row['SourceTranscriptNum']
        self.clip_num = row['ClipNum']
        self.sort_order = row['SortOrder']
        self.transcriber = row['Transcriber']
        self.clip_start = row['ClipStart']
        self.clip_stop = row['ClipStop']

        # If we're NOT skipping the RTF Text ...
        if not self.skipText:
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
                if type(self.text).__name__ == 'str':
                    temp = self.text[2:5]

                    # check to see if we're working with RTF
                    try:
                        if temp.encode('utf8') == u'rtf':
                            # convert the data to unicode just to be safe.
                            self.text = unicode(self.text, 'utf-8')
                    
                    except UnicodeDecodeError:
                        # This would sometimes get called while I was using cPickle instead of Pickle.
                        # You could probably remove the exception handling stuff and be okay, but it's
                        # not hurting anything like it is.
                        # self.dlg.editor.load_transcript(transcriptObj, 'pickle')

                        # NOPE.  There is no self.dlg.editor here!
                        pass

            # self.text gets set to be our data
            # then load_transcript is called, from transcriptionui.LoadTranscript()
            
        # If we ARE skipping the text ...
        else:
            # set the text to None
            self.text = None

        self.comment = row['Comment']
        self.minTranscriptWidth = row['MinTranscriptWidth']
        self.recordlock = row['RecordLock']
        if self.recordlock != '':
            self._isLocked = True
        self.locktime = row['LockTime']
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
            self.recordlock = DBInterface.ProcessDBDataForUTF8Encoding(self.recordlock)
