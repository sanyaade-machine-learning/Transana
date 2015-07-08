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

"""This module implements the Quote class as part of the Data Objects."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "Document DEBUG is ON!"

# import wxPython
import wx
# import Python's datetime module
import datetime
# import Python's os module
import os
# import Python's types module
import types
# import Transana's Clip Keyword Object
import ClipKeywordObject
# import Transana's Collection Object
import Collection
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

class Quote(DataObject.DataObject):
    """This class defines the structure for a quote object.  A quote
    object describes a segemnt from a text-only document for analysis in Transana."""

    def __init__(self, num=None, quoteID=None, collectionID=None, collectionParent=0, skipText=False):
        """Initialize a Quote object."""
        #   skipText indicates that the XMLText can be left off.  This leads to significantly faster loading
        #     particularly when we start having large documents with embedded images.
        DataObject.DataObject.__init__(self)
        # Remember if we're supposed to skip the RTF Text
        self.skipText = skipText
        if type(num) in (int, long):
            self.db_load_by_num(num)
        elif (collectionID != None) and (quoteID != None):
            self.db_load_by_name(collectionID, quoteID, collectionParent)
        # If we're not loading an existing Quote ...
        else:
            self.collection_id = ''
            # ... initialize the start and end character positions to -1
            self.start_char = -1
            self.end_char = -1
            self.sort_order = 0
        # For Partial Transcript Editing, create a data structure for storing the transcript information by LINE
        self.lines = []
        # Initialize a paragraph counter
        self.paragraphs = 0
        # Create a data structure for tracking very large transcripts by section
        self.paragraphPointers = {}
        # If we have text in the transcript ...
        if (num != None) or (quoteID != None):
            # ... set up data structures needed for editing large paragraphs
            self.UpdateParagraphs()

# Public methods

    def __repr__(self):
        str = 'Quote Object:\n'
        str += "number = %s\n" % self.number
        str += "id = %s\n" % self.id
        str += "collection_num = %s\n" % self.collection_num
        str = str + "collection_id = %s\n" % self.collection_id
        str += "source_document_num = %s\n" % self.source_document_num
        str += "sort_order = %s\n" % self.sort_order
        str += "start_char = %s\n" % self.start_char
        str += "end_char = %s\n" % self.end_char
        str += "comment = %s\n" % self.comment
        str += "isLocked = %s\n" % self._isLocked
        str += "recordlock = %s\n" % self.recordlock
        str += "locktime = %s\n" % self.locktime
        str += "Keywords:\n"
        for kw in self._kwlist:
            str += '  ' + kw.keywordPair + '\n'
        str = str + "LastSaveTime = %s\n" % self.lastsavetime
        if len(self.text) > 150:
            str += self.text[:150] + '\n\n'   # "text not displayed due to length.\n\n"
        else:
            str += "text = %s\n\n" % self.text
        return str.encode('utf8')

    def __eq__(self, other):
        """ Object Equality function """
        if other == None:
            return False
        else:

            if DEBUG:

                print "Quote.__eq__():", len(self.__dict__.keys()), len(other.__dict__.keys())

                print self.__dict__.keys()
                print other.__dict__.keys()
                print
                
                for key in self.__dict__.keys():
                    print key, self.__dict__[key] == other.__dict__[key],
                    if self.__dict__[key] != other.__dict__[key]:
                        print self.__dict__[key], other.__dict__[key]
                    else:
                        print
                print
            
            return self.__dict__ == other.__dict__

    def db_load_by_name(self, collectionID, quoteID, collectionParent):
        """Load a record by ID / Name."""
        # If we're in Unicode mode, we need to encode the parameters so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            collectionID = collectionID.encode(TransanaGlobal.encoding)
            quoteID = quoteID.encode(TransanaGlobal.encoding)
        # Get a database connection
        db = DBInterface.get_db()
        # Craft a query to get Quote data
        query = """SELECT * FROM Quotes2 a, Collections2 b, QuotePositions2 c
            WHERE   QuoteID = %s AND
                    a.CollectNum = b.CollectNum AND
                    b.CollectID = %s AND
                    b.ParentCollectNum = %s AND
                    c.QuoteNum = a.QuoteNum
        """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, (quoteID, collectionID, collectionParent))
        # Get the number of rows returned
        # rowcount doesn't work for sqlite!
        if TransanaConstants.DBInstalled == 'sqlite3':
            # ... so assume one record returned and check later
            n = 1
        # If not sqlite ...
        else:
            # ... we can use rowcount to know how much data was returned
            n = c.rowcount
        # If we don't get exactly one result ...
        if (n != 1):
            # Close the cursor
            c.close()
            # Clear the current Document object
            self.clear()
            # Raise an exception saying the data is not found
            raise RecordNotFoundError, (quoteID, n)
        # If we get exactly one result ...
        else:
            # Get the data from the cursor
            r = DBInterface.fetch_named(c)
            # if sqlite and no data is returned ...
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}):
                # ... close the cursor ...
                c.close()
                # ... clear the Document object ...
                self.clear()
                # ... and raise an exception
                raise RecordNotFoundError, (quoteID, 0)
            # Load the data into the Document object
            self._load_row(r)
            # Refresh the Keywords
            self.refresh_keywords()
        # Close the Database cursor
        c.close()

    def db_load_by_num(self, num):
        """Load a record by record number."""
        # Get the database connection
        db = DBInterface.get_db()
        # If we're skipping the XML Text ...
        if self.skipText:
            # Define the query to load a Quote without text
            query = """SELECT a.QuoteNum, QuoteID, a.CollectNum, CollectID, SourceDocumentNum, SortOrder, a.Comment,
                              StartChar, EndChar,
                              a.RecordLock, a.LockTime, LastSaveTime
                         FROM Quotes2 a, QuotePositions2 b, Collections2 c
                         WHERE a.QuoteNum = %s AND
                               a.QuoteNum = b.QuoteNum AND
                               a.CollectNum = c.CollectNum
                    """
        # If we're NOT skipping the RTF Text ...
        else:
            # Define the query to load a Document with everything
            # Define the query to load a Quote without text
            query = """SELECT a.QuoteNum, QuoteID, a.CollectNum, CollectID, SourceDocumentNum, SortOrder, a.Comment,
                              StartChar, EndChar,
                              XMLText, a.RecordLock, a.LockTime, LastSaveTime
                         FROM Quotes2 a, QuotePositions2 b, Collections2 c
                         WHERE a.QuoteNum = %s AND
                               a.QuoteNum = b.QuoteNum AND
                               a.CollectNum = c.CollectNum
                    """
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
            # Refresh the Keywords
            self.refresh_keywords()
        # Close the database cursor
        c.close()

    def UpdateParagraphs(self):
        """ This method divides XML text up into paragraphs, needed for editing LONG documents """
        # Initialize (or re-initialize) the paragraph pointers dictionary
        self.paragraphPointers = {}
        # If there's a defined (saved) transcript object ...
        if (self.id != 0) and (self.text != None):
            
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
        """Save the record to the database using Insert or Update as appropriate."""

        # Define and implement Demo Version limits
        if TransanaConstants.demoVersion and (self.number == 0):
            # Get a DB Cursor
            c = DBInterface.get_db().cursor()
            # Find out how many Quote records exist
            c.execute('SELECT COUNT(QuoteNum) FROM Quotes2')
            res = c.fetchone()
            c.close()
            # Define the maximum number of records allowed
            maxQuotes = TransanaConstants.maxQuotes
            # Compare
            if res[0] >= maxQuotes:
                # If the limit is exceeded, create and display the error using a SaveError exception
                prompt = _('The Transana Demonstration limits you to %d Quote records.\nPlease cancel the "Add Quote" dialog to continue.')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                raise SaveError, prompt % maxQuotes

        # Sanity checks
        if (self.id == ""):
            raise SaveError, _("Quote ID is required.")
        elif (self.collection_num == 0):
            raise SaveError, _("Parent Collection number is required.")

        if DEBUG:
            print "Quote.db_save():  %s\n%s" % (self.text, type(self.text))

        # If we're in Unicode mode, ...
        if 'unicode' in wx.PlatformInfo:
            # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
            # variables for the data.  We don't want to change the underlying object values.  Also, this way,
            # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
            id = self.id.encode(TransanaGlobal.encoding)
            # If the comment is None in the database, we can't encode that!  Otherwise, we should encode.
            if self.comment != None:
                comment = self.comment.encode(TransanaGlobal.encoding)
            else:
                comment = self.comment
        else:
            # If we don't need to encode the string values, we still need to copy them to our local variables.
            id = self.id
            comment = self.comment

        if (len(self.text) > TransanaGlobal.max_allowed_packet):   # 8388000
            raise SaveError, _("This quote is too large for the database.  Please shorten it, split it into two parts\nor if you are importing an RTF document, remove some unnecessary RTF encoding.")

        fields = ("QuoteID", "CollectNum", "SourceDocumentNum", "SortOrder", "Comment", "XMLText", "LastSaveTime")
        values = (id, self.collection_num, self.source_document_num, self.sort_order, comment, self.text)

        if (self._db_start_save() == 0):
            # Duplicate Quote IDs within a Collection are not allowed.
            if DBInterface.record_match_count("Quotes2", ("QuoteID", "CollectNum"), (id, self.collection_num)) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Quote named "%s" already exists in this Collection.\nPlease enter a different Quote ID.'), 'utf8')
                else:
                    prompt = _('A Quote named "%s" already exists in this Collection.\nPlease enter a different Quote ID.')
                raise SaveError, prompt % self.id
            # Duplicate Quote ID with a Clip ID within a Collection are not allowed.
            if DBInterface.record_match_count("Clips2", ("ClipID", "CollectNum"), (id, self.collection_num)) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Clip named "%s" already exists in this Collection.\nPlease enter a different Quote ID.'), 'utf8')
                else:
                    prompt = _('A Clip named "%s" already exists in this Collection.\nPlease enter a different Quote ID.')
                raise SaveError, prompt % self.id

            # insert the new record
            query = "INSERT INTO Quotes2\n("
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
            # Duplicate Quote IDs within a Collection are not allowed.
            if DBInterface.record_match_count("Quotes2", ("QuoteID", "!QuoteNum", "CollectNum"),
                    (id, self.number, self.collection_num)) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Quote named "%s" already exists in this Collection.\nPlease enter a different Quote ID.'), 'utf8')
                else:
                    prompt = _('A Quote named "%s" already exists in this Collection.\nPlease enter a different Quote ID.')
                raise SaveError, prompt % self.id
            # Duplicate Quote ID with Clip ID within a Collection are not allowed.
            if DBInterface.record_match_count("Clips2", ("ClipID", "CollectNum"),
                                             (id, self.collection_num)) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Clip named "%s" already exists in this Collection.\nPlease enter a different Quote ID.'), 'utf8')
                else:
                    prompt = _('An Clip named "%s" already exists in this Collection.\nPlease enter a different Quote ID.')
                raise SaveError, prompt % self.id
            
            # OK to update the episode record
            query = """UPDATE Quotes2
                SET QuoteID = %s,
                    CollectNum = %s,
                    SourceDocumentNum = %s,
                    SortOrder = %s,
                    Comment = %s,
                    XMLText = %s,
                    LastSaveTime = CURRENT_TIMESTAMP
                WHERE QuoteNum = %s
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
        
        # If we're using Transactions ...
        if use_transactions:
            # ... start the transaction
            c.execute('BEGIN')
        # Execure the Save query
        c.execute(query, values)
        # If the object number is 0, we have a new object
        if self.number == 0:
            # ... signal that the number has been changed
            numberChanged = True
            # Load the auto-assigned new number record if necessary and the saved time.
            query = """
                      SELECT QuoteNum, LastSaveTime FROM Quotes2
                      WHERE QuoteID = %s AND
                            CollectNum = %s
                    """
            # Assemble the arguments for the query
            args = (id, self.collection_num)
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
            # ... signal that the record number was not changed
            numberChanged = False
            # Load the NEW last saved time.
            query = """
                      SELECT LastSaveTime FROM Quotes2
                      WHERE QuoteNum = %s
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
            # If we are dealing with an existing Episode, delete all the Keywords
            # in anticipation of putting them all back later
            DBInterface.delete_all_keywords_for_a_group(0, 0, 0, self.number, 0)

        # Now that we know the Quote Number, we need to add the Start and End Character Position information 
        # to the QuotePositions Table

        # First, delete old data if it exists
        query = "DELETE FROM QuotePositions2 WHERE QuoteNum = %s"
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, (self.number, ))
        # Now we can insert the new data
        query = "INSERT INTO QuotePositions2 (QuoteNum, DocumentNum, StartChar, EndChar) "
        query += "VALUES (%s, %s, %s, %s)"
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the query
        c.execute(query, (self.number, self.source_document_num, self.start_char, self.end_char))

        # Initialize a blank error prompt
        prompt = ''
        # Add the Document keywords back.  Iterate through the Keyword List
        for kws in self._kwlist:
            # Try to add the Clip Keyword record.  If it is NOT added, the keyword has been changed by another user!
            if not DBInterface.insert_clip_keyword(0, 0, 0, self.number, 0, kws.keywordGroup, kws.keyword, kws.example):
                # if the prompt isn't blank ...
                if prompt != '':
                    # ... add a couple of line breaks to it
                    prompt += u'\n\n'
                # Add the current keyword to the error prompt
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt += unicode(_('Keyword "%s : %s" cannot be added to Quote "%s".\nAnother user must have edited the keyword while you were adding it.'), 'utf8') % (kws.keywordGroup, kws.keyword, self.id)

        # If there is an error prompt ...
        if prompt != '':
            # If the Quote's Number was changed ...
            if numberChanged:
                # ... change it back to zero!!
                self.number = 0
            # Undo the database save transaction
            if use_transactions:
                c.execute('ROLLBACK')
            # Close the Database Cursor
            c.close()
            # Complete the error prompt
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt += u'\n\n' + unicode(_('Please remove, and possibly replace, this keyword.'), 'utf8')
            # ... raise a SaveError exception using the error prompt
            raise SaveError, prompt
        # If there's no error prompt ...
        else:
            # ... Commit the database transaction
            if use_transactions:
                c.execute('COMMIT')
            # Close the Database Cursor
            c.close()

        # For Partial Transcript Editing, update the Paragraph Information for long transcripts
        self.UpdateParagraphs()

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database."""
        result = 1
        try:
            # Initialize delete operation, begin transaction if necessary
            (db, c) = self._db_start_delete(use_transactions)

            # Delete the QuotePosition data, if it exists
            query = "DELETE FROM QuotePositions2 WHERE QuoteNum = %s"
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Execute the query
            c.execute(query, (self.number, ))

            notes = []

            # Detect, Load, and Delete all Quote Notes.
            notes = self.get_note_nums()
            for note_num in notes:
                note = Note.Note(note_num)
                result = result and note.db_delete(0)
                del note
            del notes

            # Delete all related references in the ClipKeywords table
            if result:
                DBInterface.delete_all_keywords_for_a_group(0, 0, 0, self.number, 0)

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

    def clear_keywords(self):
        """Clear the keyword list."""
        self._kwlist = []
        
    def refresh_keywords(self):
        """Clear the keyword list and refresh it from the database."""
        self._kwlist = []
        kwpairs = DBInterface.list_of_keywords(Quote=self.number)
        for data in kwpairs:
            tempClipKeyword = ClipKeywordObject.ClipKeyword(data[0], data[1], quoteNum=self.number, example=data[2])
            self._kwlist.append(tempClipKeyword)
        
    def add_keyword(self, kwg, kw):
        """Add a keyword to the keyword list."""
        # We need to check to see if the keyword is already in the keyword list
        keywordFound = False
        # Iterate through the list
        for quoteKeyword in self._kwlist:
            # If we find a match, set the flag and quit looking.
            if (quoteKeyword.keywordGroup == kwg) and (quoteKeyword.keyword == kw):
                keywordFound = True
                break

        # If the keyword is not found, add it.  (If it's already there, we don't need to do anything!)
        if not keywordFound:
            # Create an appropriate ClipKeyword Object
            tempClipKeyword = ClipKeywordObject.ClipKeyword(kwg, kw, quoteNum=self.number)
            # Add it to the Keyword List
            self._kwlist.append(tempClipKeyword)

    def remove_keyword(self, kwg, kw):
        """Remove a keyword from the keyword list."""
        # Let's assume the Delete will fail (or be refused by the user) until it actually happens
        delResult = False

        # We need to find the keyword in the keyword list
        # Iterate through the keyword list 
        for index in range(len(self._kwlist)):
            # Look for the entry to be deleted
            if (self._kwlist[index].keywordGroup == kwg) and (self._kwlist[index].keyword == kw):
                # If the entry is found, delete it and stop looking
                del self._kwlist[index]
                delResult = True
                break
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

    def lock_record(self):
        """ Override the DataObject Lock Method """
        # If we're using the single-user version of Transana, we just need to ...
        if not TransanaConstants.singleUserVersion:
            # ... confirm that the quote has not been altered by another user since it was loaded.
            # To do this, first pull the LastSaveTime for this record from the database.
            query = """
                      SELECT LastSaveTime FROM Quotes2
                      WHERE QuoteNum = %s
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

    def GetNodeData(self, includeQuote=True):
        """ Returns the Node Data list (list of parent collections) needed for Database Tree Manipulation """
        # Load the Quote's collection
        tempCollection = Collection.Collection(self.collection_num)
        # If we're including the Quote in the Node Data ...
        if includeQuote:
            # ... get the Collection's node data and tack the Quote's ID on the end.
            return tempCollection.GetNodeData() + (self.id,)
        # If we're NOT including the Quote in the Node data ...
        else:
            # ... just return the Collection's Node Data.
            return tempCollection.GetNodeData()

    def GetNodeString(self, includeQuote=True):
        """ Returns a string that delineates the full nested collection structure for the present quote.
            if includeQuote=False, only the Collection path will be returned. """
        # Load the Quote's Collection
        tempCollection = Collection.Collection(self.collection_num)
        # If we're including the Quote in the Node String ...
        if includeQuote:
            # ... get the Collection's node string and tack the Quote's ID on the end.
            return tempCollection.GetNodeString() + ' > ' + self.id
        # If we're NOT including the Quote in the Node String ...
        else:
            # ... just return the Collection's Node String.
            return tempCollection.GetNodeString()
    

# Private methods
    # Implementation for Collection Number Property
    def _get_collection_num(self):
        return self._collection_num
    def _set_collection_num(self, num):
        self._collection_num = num
    def _del_collection_num(self):
        self._collection_num = 0

    # Implementation for Source Document Number Property
    def _get_source_document_num(self):
        return self._source_document_num
    def _set_source_document_num(self, num):
        self._source_document_num = num
    def _del_source_document_num(self):
        self._source_document_num = 0

    # Implementation for Text Property
    def _get_text(self):
        return self._text
    def _set_text(self, txt):
        self._text = txt
    def _del_text(self):
        self._text = ''

    # Implementation for Start Character Property
    def _get_start_char(self):
        return self._start_char
    def _set_start_char(self, num):
        self._start_char = num
    def _del_start_char(self):
        self._start_char = -1
                        
    # Implementation for End Character Property
    def _get_end_char(self):
        return self._end_char
    def _set_end_char(self, num):
        self._end_char = num
    def _del_end_char(self):
        self._end_char = -1

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

    def _get_kwlist(self):
        return self._kwlist
    def _set_kwlist(self, kwlist):
        self._kwlist = kwlist
    def _del_kwlist(self):
        self._kwlist = []


# Public properties
    collection_num = property(_get_collection_num, _set_collection_num, _del_collection_num,
                        """The Collection number, if associated with one.""")
    source_document_num = property(_get_source_document_num, _set_source_document_num, _del_source_document_num,
                        """The Source Document number, if associated with one.""")
    text = property(_get_text, _set_text, _del_text,
                        """Text of the transcript, stored in the database as a BLOB.""")
    start_char = property(_get_start_char, _set_start_char, _del_start_char,
                        """Value of the Quote's starting character in the Source Document.""")
    end_char = property(_get_end_char, _set_end_char, _del_end_char,
                        """Value of the Quote's ending character in the Source Document.""")
    locked_by_me = property(None, None, None,
                        """Determines if this instance owns the Quote lock.""")
    has_changed = property(_get_changed, _set_changed, _del_changed,
                        """Indicates whether the Document has been modified.""")
    lastsavetime = property(_get_lastsavetime, _set_lastsavetime, _del_lastsavetime,
                        """The timestamp of the last save (MU only).""")
    keyword_list = property(_get_kwlist, _set_kwlist, _del_kwlist,
                        """The list of keywords that have been applied to the Document.""")

    def _load_row(self, row):
    	self.number = row['QuoteNum']
        self.id = row['QuoteID']
        self.collection_num = row['CollectNum']
        self.collection_id = row['CollectID']
        self.source_document_num = row['SourceDocumentNum']
        self.start_char = row['StartChar']
        self.end_char = row['EndChar']
        self.sort_order = row['SortOrder']
        self.comment = row['Comment']
        self.recordlock = row['RecordLock']
        if self.recordlock != '':
            self._isLocked = True
        self.locktime = row['LockTime']
        # If we're NOT skipping the XML Text ...
        if not self.skipText:
            # Can I get away with assuming Unicode?
            # Here's the plan:
            #   test for rtf in here, if you find rtf, process normally
            #   if you don't find it, pass data off to some weirdo method in TranscriptEditor.py

            # 1 - Determine encoding, adjust if needed
            # 2 - enact the plan above

            # determine encoding, fix if needed
            if type(row['XMLText']).__name__ == 'array':

                if DEBUG:
                    print "Document._load_row(): 2", row['XMLText'].typecode
                
                if row['XMLText'].typecode == 'u':
                    self.text = row['XMLText'].tounicode()
                else:
                    self.text = row['XMLText'].tostring()
            else:
                self.text = row['XMLText']

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

        self.lastsavetime = row['LastSaveTime']
        self.changed = False

        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.collection_id = DBInterface.ProcessDBDataForUTF8Encoding(self.collection_id)
            self.comment = DBInterface.ProcessDBDataForUTF8Encoding(self.comment)
            self.recordlock = DBInterface.ProcessDBDataForUTF8Encoding(self.recordlock)
