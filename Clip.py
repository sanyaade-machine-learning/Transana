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

"""This module implements the Clip class as part of the Data Objects."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

DEBUG = False
if DEBUG:
    print "Clip DEBUG is ON!"

# import wxPython
import wx
# import Python's os module
import os
# import Python's types module
import types
# import Transana's Clip Keyword object
import ClipKeywordObject
# import Transana's Collection Object
import Collection
# Import the Transana base Data Object
import DataObject
# Import the Transana Database Interface
import DBInterface
# import Transana's Dialogs
import Dialogs
# Import Transana's Episode Object
import Episode
# import Transana's Miscellaneous Functions
import Misc
# import Transana's Note Object
import Note
# Import Transana's exceptions
from TransanaExceptions import *
# Import the Transana Constants
import TransanaConstants
# Import Transana's Global Variables
import TransanaGlobal
# import Transana's Transcript object
import Transcript

TIMECODE_CHAR = "\\'a4"   # Note that this differs from the TIMECODE_CHAR in TranscriptEditor.py
                          # because this is for RTF text and that is for parsed text.

class Clip(DataObject.DataObject):
    """This class defines the structure for a clip object.  A clip object
    describes a portion of a video (or other media) file."""

    def __init__(self, id_or_num=None, collection_name=None, collection_parent=0):
        """Initialize an Clip object."""
        DataObject.DataObject.__init__(self)
        # By default, use the Video Root folder if one has been defined
        self.useVideoRoot = (TransanaGlobal.configData.videoPath != '')

        if type(id_or_num) in (int, long):
            self.db_load_by_num(id_or_num)
        elif isinstance(id_or_num, types.StringTypes):
            self.db_load_by_name(id_or_num, collection_name, collection_parent)
        else:
            self.number = 0
            self.id = ''
            self.comment = ''
            self.collection_num = 0
            self.collection_id = ''
            self.episode_num = 0
            # Keep a list of transcript objects associated with this clip
            self.transcripts = []
            self.media_filename = ""
            # Initialize an empty list for additional media files
            self.additional_media_files = []
            self.clip_start = 0
            self.clip_stop = 0
            # With multiple videos, it's possible the clip time values need an offset
            # (if the first Episode video isn't included in the Clip.)
            self.offset = 0
            # With multiple videos, it's possible to not want the AUDIO for the first video track,
            # but default to True
            self.audio = 1
            self.sort_order = 0
            
        # Create empty placeholders for Series and Episode IDs.  These only get populated if the
        # values are needed, and cannot be implemented in the regular LOADs because the Series/
        # Episode may no longer exist.
        self._series_id = ""
        self._episode_id = ""



# Public methods
    def __repr__(self):
        str = 'Clip Object Definition:\n'
        str = str + "number = %s\n" % self.number
        str = str + "id = %s\n" % self.id
        str = str + "comment = %s\n" % self.comment
        str = str + "collection_num = %s\n" % self.collection_num 
        str = str + "collection_id = %s\n" % self.collection_id
        str = str + "episode_num = %s\n" % self.episode_num
        str = str + "media_filename = %s\n" % self.media_filename 
        str = str + "Additional media file:\n"
        for addFile in self.additional_media_files:
            str += '  %s  %s %s %s\n' % (addFile['filename'], addFile['offset'], addFile['length'], addFile['audio'])
        str = str + "clip_start = %s (%s)\n" % (self.clip_start, Misc.time_in_ms_to_str(self.clip_start))
        str = str + "clip_stop = %s (%s)\n" % (self.clip_stop, Misc.time_in_ms_to_str(self.clip_stop))
        str += "offset = %s\n" % Misc.time_in_ms_to_str(self.offset)
        str += "audio = %s\n" % self.audio
        str = str + "sort_order = %s\n" % self.sort_order
        # Iterate through transcript objects
        for tr in self.transcripts:
            str = str + '\nClip Transcript: %s from %s\n' % (tr.number, tr.source_transcript)
            str = str + '%s\n' % tr
        for kws in self.keyword_list:
            str = str + "\nKeyword:  %s" % kws
        str = str + '\n'
        return str.encode('utf8')
        
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
            tabStart = newText.find(chr(wx.WXK_TAB), 0)
            if tabStart == -1:
                break
            newText = newText[:tabStart] + '  ' + newText[tabStart + 1:]

        return newText

    def db_load_by_name(self, clip_name, collection_name, collection_parent=0):
        """Load a record by ID / Name."""
        self.clear()
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            clip_name = clip_name.encode(TransanaGlobal.encoding)
            collection_name = collection_name.encode(TransanaGlobal.encoding)
        # Get a database connection
        db = DBInterface.get_db()
        # craft a query to get Clip data
        query = """ SELECT a.*, b.*
          FROM Clips2 a, Collections2 b
          WHERE ClipID = %s AND
                a.CollectNum = b.CollectNum AND
                b.CollectID = %s AND
                b.ParentCollectNum = %s """
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, (clip_name, collection_name, collection_parent))
        # Get the number of rows returned
        n = c.rowcount
        # if we don't get exactly one result ...
        if (n != 1):
            # close the cursor
            c.close()
            # clear the current Clip
            self.clear()
            # Raise an exception saying the record is not found
            raise RecordNotFoundError, (collection_name + ", " + clip_name, n)
        # If we get exactly one result ...
        else:
            # get the data from the cursor
            r = DBInterface.fetch_named(c)
            # Load the data into the Clip object
            self._load_row(r)
            # Load additional Media Files, which aren't handled in the "old" code
            self.load_additional_vids()
            # Refresh the Keywords
            self.refresh_keywords()
        # Close the Database Cursor
        c.close()
        
    def db_load_by_num(self, num):
        """Load a record by record number."""
        self.clear()
        # Get a database Connection
        db = DBInterface.get_db()
        # Craft a query to get the Clip data
        query = """
        SELECT a.*, b.*
          FROM Clips2 a, Collections2 b
          WHERE a.ClipNum = %s AND
                a.CollectNum = b.CollectNum
        """
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, (num,))
        # Get the number of rows returned
        n = c.rowcount
        # If we don't get exactly one result ...
        if (n != 1):
            # close the database cursor
            c.close()
            # clear the current Clip object
            self.clear()
            # Raise an exception indicating the data was not found
            raise RecordNotFoundError, (num, n)
        # If we get exactly one result ...
        else:
            # ... get the data from the cursor
            r = DBInterface.fetch_named(c)
            # ... load the data into the Clip Object
            self._load_row(r)
            # Load Additional Media Files, which aren't handled in the "old" code
            self.load_additional_vids()
            # Refresh the Keywords
            self.refresh_keywords()
        # Close the database cursor
        c.close()

    def db_save(self):
        """Save the record to the database using Insert or Update as appropriate."""
        # Define and implement Demo Version limits
        if TransanaConstants.demoVersion and (self.number == 0):
            # Get a DB Cursor
            c = DBInterface.get_db().cursor()
            # Find out how many records exist
            c.execute('SELECT COUNT(ClipNum) from Clips2')
            res = c.fetchone()
            c.close()
            # Define the maximum number of records allowed
            maxClips = TransanaConstants.maxClips
            # Compare
            if res[0] >= maxClips:
                # If the limit is exceeded, create and display the error using a SaveError exception
                prompt = _('The Transana Demonstration limits you to %d Clip records.\nPlease cancel the "Add Clip" dialog to continue.')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                raise SaveError, prompt % maxClips

        # Sanity checks
        if self.id == "":
            raise SaveError, _("Clip ID is required.")
        if (self.collection_num == 0):
            raise SaveError, _("Parent Collection number is required.")
        elif self.media_filename == "":
            raise SaveError, _("Media Filename is required.")
        # If a user Adjusts Indexes, it's possible to have a clip that starts BEFORE the media file.
        elif self.clip_start < 0.0:
            raise SaveError, _("Clip cannot start before media file begins.")
        else:
            # videoPath probably has the OS.sep character, but we need the generic "/" character here.
            videoPath = TransanaGlobal.configData.videoPath
            # Determine if we are supposed to extract the Video Root Path from the Media Filename and extract it if appropriate
            if self.useVideoRoot and (videoPath == self.media_filename[:len(videoPath)]):
                tempMediaFilename = self.media_filename[len(videoPath):]
            else:
                tempMediaFilename = self.media_filename

            # Substitute the generic OS seperator "/" for the Windows "\".
            tempMediaFilename = tempMediaFilename.replace('\\', '/')
            # If we're not in Unicode mode ...
            if 'ansi' in wx.PlatformInfo:
                # ... we don't need to encode the string values, but we still need to copy them to our local variables.
                id = self.id
                comment = self.comment
            # If we're in Unicode mode ...
            else:
                # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
                # variables for the data.  We don't want to change the underlying object values.  Also, this way,
                # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
                id = self.id.encode(TransanaGlobal.encoding)
                tempMediaFilename = tempMediaFilename.encode(TransanaGlobal.encoding)
                videoPath = videoPath.encode(TransanaGlobal.encoding)
                comment = self.comment.encode(TransanaGlobal.encoding)

        self._sync_collection()

        values = (id, self.collection_num, self.episode_num, \
                      tempMediaFilename, \
                      self.clip_start, self.clip_stop, self.offset, self.audio, comment, \
                      self.sort_order)
        if (self._db_start_save() == 0):
            if DBInterface.record_match_count("Clips2", \
                                ("ClipID", "CollectNum"), \
                                (id, self.collection_num) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Clip named "%s" already exists in this Collection.\nPlease enter a different Clip ID.'), 'utf8') % self.id
                else:
                    prompt = _('A Clip named "%s" already exists in this Collection.\nPlease enter a different Clip ID.') % self.id
                raise SaveError, prompt
            # insert the new record
            query = """
            INSERT INTO Clips2
                (ClipID, CollectNum, EpisodeNum, 
                 MediaFile, ClipStart, ClipStop, ClipOffset, Audio, ClipComment,
                 SortOrder)
                VALUES
                (%s,%s,%s,%s,%s,%s,%s,%s,%s, %s)
            """
        else:
            if DBInterface.record_match_count("Clips2", \
                            ("ClipID", "CollectNum", "!ClipNum"), \
                            (id, self.collection_num, self.number)) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Clip named "%s" already exists in this Collection.\nPlease enter a different Clip ID.'), 'utf8') % self.id
                else:
                    prompt = _('A Clip named "%s" already exists in this Collection.\nPlease enter a different Clip ID.') % self.id
                raise SaveError, prompt

            # update the record
            query = """
            UPDATE Clips2
                SET ClipID = %s,
                    CollectNum = %s,
                    EpisodeNum = %s,
                    MediaFile = %s,
                    ClipStart = %s,
                    ClipStop = %s,
                    ClipOffset = %s,
                    Audio = %s,
                    ClipComment = %s,
                    SortOrder = %s
                WHERE ClipNum = %s
            """
            values = values + (self.number,)

        c = DBInterface.get_db().cursor()
        c.execute(query, values)
        if self.number == 0:
            # If we are dealing with a brand new Clip, it does not yet know its
            # record number.  It HAS a record number, but it is not known yet.
            # The following query should produce the correct record number.
            query = """
                      SELECT ClipNum FROM Clips2
                      WHERE ClipID = %s AND
                            CollectNum = %s
                    """
            tempDBCursor = DBInterface.get_db().cursor()
            tempDBCursor.execute(query, (id, self.collection_num))
            if tempDBCursor.rowcount == 1:
                self.number = tempDBCursor.fetchone()[0]
            else:
                raise RecordNotFoundError, (self.id, tempDBCursor.rowcount)
            tempDBCursor.close()
        else:
            # If we are dealing with an existing Clip, delete all the Keywords
            # in anticipation of putting them all back in after we deal with the
            # Clip Transcript
            DBInterface.delete_all_keywords_for_a_group(0, self.number)
            
        # Now let's deal with the Clip's Transcripts

        # For each transcript in the list of clip transcripts ...
        for tr in self.transcripts:
            # Assign the clip's number as the transcript's clip number
            tr.clip_num = self.number
            # There is a spot (in XML Import) where we signal that we DON'T want the Clip
            # Transcript saved by setting its number to -1.
            if tr.number != -1:
                # save the transcript
                tr.db_save()

        # To save the additional video file names, we must first delete them from the database!
        # Craft a query to remove all existing Additonal Videos
        query = "DELETE FROM AdditionalVids2 WHERE ClipNum = %d" % self.number
        # Execute the query
        c.execute(query)
        # Define the query to insert the additional media files into the databse
        query = "INSERT INTO AdditionalVids2 (EpisodeNum, ClipNum, MediaFile, VidLength, Offset, Audio) VALUES (%s, %s, %s, %s, %s, %s)"
        # For each additional media file ...
        for vid in self.additional_media_files:
            # Encode the filename
            tmpFilename = vid['filename'].encode(TransanaGlobal.encoding)
            # Determine if we are supposed to extract the Video Root Path from the Media Filename and extract it if appropriate
            if self.useVideoRoot and (videoPath == tmpFilename[:len(videoPath)]):
                tmpFilename = tmpFilename[len(videoPath):]
            # Substitute the generic OS seperator "/" for the Windows "\".
            tmpFilename = tmpFilename.replace('\\', '/')
            # In Windows single-user, audio of "True" isn't being stored correctly.  It seems to be OK in MU.  (Haven't checked Mac-SU.)
            if vid['audio'] == True:
                vid['audio'] = 1
            # Get the data for each insert query
            data = (0, self.number, tmpFilename, vid['length'], vid['offset'], vid['audio'])
            # Execute the query
            c.execute(query, data)

        # Add the Episode keywords back
        for kws in self._kwlist:
            DBInterface.insert_clip_keyword(0, self.number, kws.keywordGroup, kws.keyword, kws.example)

        c.close()

    def db_delete(self, use_transactions=True, examplesPrompt=True):
        """ Delete this object record from the database.  Parameters indicate if we should use DB Transactions
            and if we should prompt about the deletion of Keyword Examples.  (in Clip Merging, for example, we
            don't want to prompt.) """
        result = 1
        try:
            # Initialize delete operation, begin transaction if necessary
            (db, c) = self._db_start_delete(use_transactions)
            # If this clip serves as a Keyword Example, we should prompt the user about
            # whether it should really be deleted
            kwExampleList = DBInterface.list_all_keyword_examples_for_a_clip(self.number)
            if len(kwExampleList) > 0:
                if len(kwExampleList) == 1:
                    prompt = _('Clip "%s" has been defined as a Keyword Example for Keyword "%s : %s".')
                    data = (self.id, kwExampleList[0][0], kwExampleList[0][1])
                else:
                    prompt = _('Clip "%s" has been defined as a Keyword Example for multiple Keywords.')
                    data = self.id
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8') % data
                    prompt = prompt + unicode(_('\nAre you sure you want to delete it?'), 'utf8')
                else:
                    prompt = prompt % data
                    prompt = prompt + _('\nAre you sure you want to delete it?')
                dlg = Dialogs.QuestionDialog(None, prompt, _('Delete Clip'))
                # if we should prompt about examples, then show the dialog.  If the user responds "No" ...
                if examplesPrompt and (dlg.LocalShowModal() == wx.ID_NO):
                    # ... destroy the dialog box.
                    dlg.Destroy()
                    # A Transcaction was started and the record was locked in _db_start_delete().  Unlock it here if the
                    # user cancels the delete (after rolling back the Transaction)!
                    if self.isLocked:
                        # c (the database cursor) only exists if the record lock was obtained!
                        # if we are using Transactions locally ...
                        if use_transactions:
                            # We must roll back the transaction before we unlock the record.
                            c.execute("ROLLBACK")
                        # Close the database cursor
                        c.close()
                        # unlock the Clip record
                        self.unlock_record()
                    # Exit the Delete method, indicating the delete did not succeed.
                    return 0
                # If we don't show the dialog at all, or if the users doesn't say "No" ...
                else:
                    # ... then just destroy the dialog box.
                    dlg.Destroy()

            # Detect, Load, and Delete all Clip Notes.
            notes = self.get_note_nums()
            for note_num in notes:
                note = Note.Note(note_num)
                result = result and note.db_delete(0)
                del note
            del notes

            # Okay, theoretically we have a lock on this clip's Transcripts, in self.transcripts.  However, that lock
            # prevents us from deleting it!!  Oops.  Therefore, IF we have a legit lock on it, unlock it for
            # the delete.

            # For each transcript in the list of clip transcripts ...
            for tr in self.transcripts:
                # ... if the transcript is locked ...
                if tr.isLocked:
                    # ... then we need to unlock it before we can delete it.
                    tr.unlock_record()
                # Now try to delete the transcript
                result = result and tr.db_delete(0)

            # NOTE:  It is important for the calling routine to delete references to the Keyword Examples
            #        from the screen.  However, that code does not belong in the Clip Object, but in the
            #        user interface.  That is why it is not included here as part of the result.

            # Delete all related references in the ClipKeywords table
            if result:
                DBInterface.delete_all_keywords_for_a_group(0, self.number)

            if result:
                # Craft a query to remove all existing Additonal Videos
                query = "DELETE FROM AdditionalVids2 WHERE ClipNum = %d" % self.number
                # Execute the query
                c.execute(query)

            # Delete the actual Clip record.
            self._db_do_delete(use_transactions, c, result)

            # Cleanup
            c.close()
            self.clear()
        except RecordLockedError, e:
            # if a sub-record is locked, we may need to unlock the Clip record (after rolling back the Transaction)
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
        # Also lock the Clip Transcript records

        # We can run into trouble if one of a multiple transcript records is locked.  Specifically, if a later
        # transcript is already locked, trying to lock it below will raise an exception, but the earlier
        # transcripts will have gotten locked now and will remain so improperly.
        # Therefore, let's check to see if the transcripts are already locked before trying to lock them.

        # Iterate through the transcripts ...
        for tr in self.transcripts:
            # If a transcript is locked ...
            if tr.isLocked or ((tr.record_lock != '') and (tr.record_lock != None)):
                # Raise an exception before locking any transcripts!
                raise RecordLockedError(user=tr.record_lock)

        # For each transcript in the clip transcripts list ...
        for tr in self.transcripts:
            # ... lock the transcript
            tr.lock_record()
        
        # Lock the Clip Record.  Call this second so the Clip is not identified as locked if the
        # Clip Transcript record lock fails.
        DataObject.DataObject.lock_record(self)
            

    def unlock_record(self):
        """ Override the DataObject Unlock Method """
        # Unlock the Clip Record
        DataObject.DataObject.unlock_record(self)
        # Also unlock the Clip Transcript records

        # For each transcript in the clip transcript list ...
        for tr in self.transcripts:
            # ... unlock the transcript
            tr.unlock_record()

    def duplicate(self):
        # Inherit duplicate method
        # BUG:  This duplicate() call wipes out the Keyword Example numbers of BOTH copies of the Clips!!
        #       The keyword_list is clearly a pointer, not a copy of the list object.
        # newClip = DataObject.duplicate(self)
        # INSTEAD:  Let's just create a new clip, loading the existing data from the database!
        # If we know the Clip Number, the Clip is in the database ...
        if self.number != 0:
            # ... so we can just load it
            newClip = Clip(self.number)
        # If we have a clips that's not currently in the database ...
        else:
            # ... we can use the old duplicate method.  If needed, we may want to copy all the data manually here.
            newClip = DataObject.DataObject.duplicate(self)
        # Eliminate the new clip's object number so it will get a new on when saved.
        newClip.number = 0
        # A new Clip should get a new Clip Transcripts!
        for tr in newClip.transcripts:
            tr.number = 0
        # Sort Order should not be duplicated!
        newClip.sort_order = 0
        # Copying a Clip should not cause additional Keyword Examples to be created.
        # We need to strip the "example" status for all keywords in the new clip.
        for clipKeyword in newClip.keyword_list:
            clipKeyword.example = 0

        return newClip
        
    def clear_keywords(self):
        """Clear the keyword list."""
        self._kwlist = []
        
    def refresh_keywords(self):
        """Clear the keyword list and refresh it from the database."""
        self._kwlist = []
        kwpairs = DBInterface.list_of_keywords(Clip=self.number)
        for data in kwpairs:
            tempClipKeyword = ClipKeywordObject.ClipKeyword(data[0], data[1], clipNum=self.number, example=data[2])
            self._kwlist.append(tempClipKeyword)
        
    def add_keyword(self, kwg, kw, example=0):
        """ Add a keyword to the keyword list.  By default, it is NOT a keyword example. """
        # We need to check to see if the keyword is already in the keyword list
        keywordFound = False
        # Iterate through the list
        for clipKeyword in self._kwlist:
            # If we find a match, set the flag and quit looking.
            if (clipKeyword.keywordGroup == kwg) and (clipKeyword.keyword == kw):
                keywordFound = True
                break

        # If the keyword is not found, add it.  (If it's already there, we don't need to do anything!)
        if not keywordFound:
            # Create an appropriate ClipKeyword Object
            tempClipKeyword = ClipKeywordObject.ClipKeyword(kwg, kw, clipNum=self.number, example=example)
            # Add it to the Keyword List
            self._kwlist.append(tempClipKeyword)

    def remove_keyword(self, kwg, kw):
        """Remove a keyword from the keyword list.  The value returned by this function can be:
             0  Keyword NOT deleted.  (probably overridden by the user)
             1  Keyword deleted, but it was NOT a Keyword Example
             2  Keyword deleted, and it WAS a Keyword Example. """
        # We need different return codes for failure, success of a Non-Example, and success of an Example.
        # If it's an example, we need to remove the Node on the Database Tree Tab

        # Let's assume the Delete will fail (or be refused by the user) until it actually happens.
        delResult = 0

        # We need to find the keyword in the keyword list
        # Iterate through the keyword list
        for index in range(len(self._kwlist)):

            # Look for the entry to be deleted
            if (self._kwlist[index].keywordGroup == kwg) and (self._kwlist[index].keyword == kw):
                # If it's a Keyword Example ...
                if self._kwlist[index].example == 1:
                    # ... build and encode the prompt ...
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Clip "%s" has been designated as an example of Keyword "%s : %s".\nRemoving this Keyword from the Clip will also remove the Clip as a Keyword Example.\n\nDo you want to remove Clip "%s" as an example of Keyword "%s : %s"?'), 'utf8')
                    else:
                        prompt = _('Clip "%s" has been designated as an example of Keyword "%s : %s".\nRemoving this Keyword from the Clip will also remove the Clip as a Keyword Example.\n\nDo you want to remove Clip "%s" as an example of Keyword "%s : %s"?')
                    # ... ask the user if they really want to remove the example keyword ...
                    dlg = Dialogs.QuestionDialog(TransanaGlobal.menuWindow, prompt % (self.id, kwg, kw, self.id, kwg, kw))
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                    # ... if the user sayd "Yes" ...
                    if result == wx.ID_YES:
                        # If the entry is found and the user confirms, delete it
                        del self._kwlist[index]
                        # Signal that a Keyword Example was deleted, so the GUI can update.
                        delResult = 2
                else:
                    # If the entry is found, delete it and stop looking
                    del self._kwlist[index]
                    # Signal that the delete was successful and was NOT an Example.
                    delResult = 1
                # Once the entry has been found, stop looking for it
                break
            
        # Signal whether the delete was successful
        return delResult

    def has_keyword(self, kwg, kw):
        """ Determines if the Episode has a given keyword assigned """
        # Assume the result will be false
        res = False
        # Iterate through the keyword list
        for keyword in self.keyword_list:
            # See if the text passed in matches the strings in the keyword objects in the keyword list
            if (kwg == keyword.keywordGroup) and (kw == keyword.keyword):
                # If so, signal that it HAS been found
                res = True
                # If found, we don't need to look any more!
                break
        # Return the results
        return res

    def load_additional_vids(self):
        """Load additional media file names from the database."""
        # Get a database connection
        db = DBInterface.get_db()
        # Create a database cursor
        c = db.cursor()
        # Define the DB Query
        query = "SELECT MediaFile, VidLength, Offset, Audio FROM AdditionalVids2 WHERE ClipNum = %s ORDER BY AddVidNum"
        # Execute the query
        c.execute(query, self.number)
        # For each video in the query results ...
        for (vidFilename, vidLength, vidOffset, audio) in c.fetchall():
            # Detection of the use of the Video Root Path is platform-dependent and must be done for EACH filename!
            if wx.Platform == "__WXMSW__":
                # On Windows, check for a colon in the position, which signals the presence or absence of a drive letter
                useVideoRoot = (vidFilename[1] != ':')
            else:
                # On Mac OS-X and *nix, check for a slash in the first position for the root folder designation
                useVideoRoot = (vidFilename[0] != '/')
            # If we are using the Video Root Path, add it to the Filename
            if useVideoRoot:
                video = TransanaGlobal.configData.videoPath.replace('\\', '/') + DBInterface.ProcessDBDataForUTF8Encoding(vidFilename)
            else:
                video = DBInterface.ProcessDBDataForUTF8Encoding(vidFilename)
            # Add the video to the additional media files list
            self.additional_media_files = {'filename' : video,
                                           'length'   : vidLength,
                                           'offset'   : vidOffset,
                                           'audio'    : audio}
        # Close the database cursor
        c.close()

    def remove_an_additional_vid(self, indx):
        """ remove ONE additional media file from the list of additional media files """
        del(self._additional_media[indx])

    def GetNodeData(self, includeClip=True):
        """ Returns the Node Data list (list of parent collections) needed for Database Tree Manipulation """
        # Load the Clip's collection
        tempCollection = Collection.Collection(self.collection_num)
        # If we're including the Clip in the Node Data ...
        if includeClip:
            # ... get the Collection's node data and tack the Clip's ID on the end.
            return tempCollection.GetNodeData() + (self.id,)
        # If we're NOT including the Clip in the Node data ...
        else:
            # ... just return the Collection's Node Data.
            return tempCollection.GetNodeData()

    def GetNodeString(self, includeClip=True):
        """ Returns a string that delineates the full nested collection structure for the present clip.
            if includeClip=False, only the Collection path will be returned. """
        # Load the Clip's Collection
        tempCollection = Collection.Collection(self.collection_num)
        # If we're including the Clip in the Node String ...
        if includeClip:
            # ... get the Collection's node string and tack the Clip's ID on the end.
            return tempCollection.GetNodeString() + ' > ' + self.id
        # If we're NOT including the Clip in the Node String ...
        else:
            # ... just return the Collection's Node String.
            return tempCollection.GetNodeString()
    
# Private methods    

    def _load_row(self, r):
        self.number = r['ClipNum']
        self.id = r['ClipID']
        self.comment = r['ClipComment']
        self.collection_num = r['CollectNum']
        self.collection_id = r['CollectID']
        self.episode_num = r['EpisodeNum']
        self.media_filename = r['MediaFile']
        self.clip_start = r['ClipStart']
        self.clip_stop = r['ClipStop']
        self.offset = r['ClipOffset']
        self.audio = r['Audio']
        self.sort_order = r['SortOrder']
        # I've seen once instance of incomplete updating of the database on 2.40 upgrade database modification.  The following
        # fixes problems caused there.
        # If the offset is NULL in the database, it would come up as None here....
        if self.offset == None:
            # If so, convert it to 0 to prevent problems loading clips.  (It should have been converted.)
            self.offset = 0
        # If self.audio is None, it's NULL in the database.
        if self.audio == None:
            # It should have been converted to 1.
            self.audio = 1
        # Initialize a list of Transcript objects
        self.transcripts = []
        # Load the Clip Transcripts.  Get the list of clip transcripts from the database and interate ...
        for tr in DBInterface.list_clip_transcripts(self.number):
            # Create a Transcript Object, passing each Transcript's Transcript Number (parameter 0)
            tempTranscript = Transcript.Transcript(tr[0])
            # Append the transcript object to the Clip's Transcript List
            self.transcripts.append(tempTranscript)
            
        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.comment = DBInterface.ProcessDBDataForUTF8Encoding(self.comment)
            self.collection_id = DBInterface.ProcessDBDataForUTF8Encoding(self.collection_id)
            self.media_filename = DBInterface.ProcessDBDataForUTF8Encoding(self.media_filename)

        # Remember whether the MediaFile uses the VideoRoot Folder or not.
        # Detection of the use of the Video Root Path is platform-dependent.
        if wx.Platform == "__WXMSW__":
            # On Windows, check for a colon in the position, which signals the presence or absence of a drive letter
            self.useVideoRoot = (self.media_filename[1] != ':')
        else:
            # On Mac OS-X and *nix, check for a slash in the first position for the root folder designation
            self.useVideoRoot = (self.media_filename[0] != '/')
        # If we are using the Video Root Path, add it to the Filename
        if self.useVideoRoot:
            self.media_filename = TransanaGlobal.configData.videoPath.replace('\\', '/') + self.media_filename

    def _sync_collection(self):
        """Synchronize the Collection ID property to reflect the current state
        of the Collection Number property."""
        tempCollection = Collection.Collection(self.collection_num)
        self.collection_id = tempCollection.id
        
    def _get_col_num(self):
        return self._col_num
    def _set_col_num(self, num):
        self._col_num = num
    def _del_col_num(self):
        self._col_num = 0

    def _get_ep_num(self):
        return self._ep_num
    def _set_ep_num(self, num):
        self._ep_num = num
    def _del_ep_num(self):
        self._ep_num = 0

    def _get_t_num(self):
        return self._t_num
    def _set_t_num(self, num):
        self._t_num = num
    def _del_t_num(self):
        self._t_num = 0

    def _get_clip_transcript_nums(self):
        # Initialize a list object
        tempList = []
        # Iterate through the transcripts ...
        for tr in self.transcripts:
            # ... and append their numbers to the list.
            tempList.append(tr.number)
        # Return the list
        return tempList

    def _get_fname(self):
        return self._fname.replace('/', os.sep)
    def _set_fname(self, fname):
        self._fname = fname.replace('\\', '/')
    def _del_fname(self):
        self._fname = ""

    def _get_additional_media(self):
        temp_additional_media = []
        for vid in self._additional_media:
            vid['filename'] = vid['filename'].replace('/', os.sep)
            temp_additional_media.append(vid)
        return temp_additional_media
    def _set_additional_media(self, vidDict):
        # If we receive a Dictionary Object ...
        if isinstance(vidDict, dict):
            # ... replace the backslashes in the filename item
            vidDict['filename'] = vidDict['filename'].replace('\\', '/')
            # Add the dictionary object to the media files list
            self._additional_media.append(vidDict)
        # If we receive something else (a list or a tuple, most likely) ...
        else:
            # ... for each element (which should be a dictionary object) ...
            for vid in vidDict:
                # ... replace the backslashes in the filename item
                vid['filename'] = vid['filename'].replace('\\', '/')
                # Add the dictionary object to the media files list
                self._additional_media.append(vid)

    def _del_additional_media(self):
        self._additional_media = []

    def _get_clip_start(self):
        return self._clip_start
    def _set_clip_start(self, cs):
        self._clip_start = cs
    def _del_clip_start(self):
        self._clip_start = -1

    def _get_clip_stop(self):
        return self._clip_stop
    def _set_clip_stop(self, cs):
        self._clip_stop = cs
    def _del_clip_stop(self):
        self._clip_stop = -1

    def _get_offset(self):
        return self._offset
    def _set_offset(self, offset):
        self._offset = offset
    def _del_offset(self):
        self._offset = 0

    def _get_audio(self):
        return self._audio
    def _set_audio(self, val):
        self._audio = val
    def _del_audio(self):
        self._audio = 1  # Default to True

    def _get_sort_order(self):
        return self._sort_order
    def _set_sort_order(self, so):
        self._sort_order = so
    def _del_sort_order(self):
        self._sort_order = 0

    def _get_kwlist(self):
        return self._kwlist
    def _set_kwlist(self, kwlist):
        self._kwlist = kwlist
    def _del_kwlist(self):
        self._kwlist = []

    # Clips only know originating Episode Number, which can be used to find the Series ID and Episode ID.
    # For the sake of efficiency, whichever is called first loads both values.
    def _get_series_id(self):
        if self._series_id == "":
            try:
                tempEpisode = Episode.Episode(self.episode_num)
                self._series_id = tempEpisode.series_id
                self._episode_id = tempEpisode.id
            except:
                pass
            return self._series_id
        else:
            return self._series_id
        
    # Clips only know originating Episode Number, which can be used to find the Series ID and Episode ID.
    # For the sake of efficiency, whichever is called first loads both values.
    def _get_episode_id(self):
        if self._episode_id == "":
            try:
                tempEpisode = Episode.Episode(self.episode_num)
                self._series_id = tempEpisode.series_id
                self._episode_id = tempEpisode.id
            except:
                pass
            return self._episode_id
        else:
            return self._episode_id

# Public properties
    collection_num = property(_get_col_num, _set_col_num, _del_col_num,
                        """Collection number to which the clip belongs.""")
    episode_num = property(_get_ep_num, _set_ep_num, _del_ep_num,
                        """Number of episode from which this Clip was taken.""")
    # TranscriptNum is the Transcript Number the Clip was created FROM, not the number of the Clip Transcript!
    transcript_num = property(_get_t_num, _set_t_num, _del_t_num,
                        """Number of the transcript from which this Clip was taken.""")
    # This read-only property provides a LIST of the numbers of Transcripts.
    clip_transcript_nums = property(_get_clip_transcript_nums, None, None,
                        """Number of the Clip's transcript record in the Transcript Table.""")
    media_filename = property(_get_fname, _set_fname, _del_fname,
                        """The name (including path) of the media file.""")
    additional_media_files = property(_get_additional_media, _set_additional_media, _del_additional_media,
                        """A list of additional media files (including path).""")
    clip_start = property(_get_clip_start, _set_clip_start, _del_clip_start,
                        """Starting position of the Clip in the media file.""")
    clip_stop = property(_get_clip_stop, _set_clip_stop, _del_clip_stop,
                        """Ending position of the Clip in the media file.""")
    offset = property(_get_offset, _set_offset, _del_offset,
                      """Offset amount for a Clip.  Needed if the first Episode video is not used in a multi-video Clip. """)
    audio = property(_get_audio, _set_audio, _del_audio,
                     """ Is Audio included in the first video in a multi-video clip? """)
    sort_order = property(_get_sort_order, _set_sort_order, _del_sort_order,
                        """Sort Order position within the parent Collection.""")
    keyword_list = property(_get_kwlist, _set_kwlist, _del_kwlist,
                        """The list of keywords that have been applied to
                        the Clip.""")
    series_id = property(_get_series_id, None, None,
                        "ID for the Series from which this Clip was created, if the (bridge) Episode still exists")
    episode_id = property(_get_episode_id, None, None,
                        "ID for the Episode from which this Clip was created, if it still exists")
