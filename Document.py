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

"""This module implements the Document class as part of the Data Objects."""

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
# import Transana's Library Object
import Library
# import Transana's Constants
import TransanaConstants
# import Transana's Exceptions
from TransanaExceptions import *
# import Transana's Globals
import TransanaGlobal

class Document(DataObject.DataObject):
    """This class defines the structure for a document object.  A document
    object describes a text-only document for analysis in Transana."""

    def __init__(self, num=None, libraryID=None, documentID=None, skipText=False):
        """Initialize a Document object."""
        #   skipText indicates that the XMLText can be left off.  This leads to significantly faster loading
        #     particularly when we start having large documents with embedded images.
        DataObject.DataObject.__init__(self)
        # Remember if we're supposed to skip the RTF Text
        self.skipText = skipText
        if type(num) in (int, long):
            self.db_load_by_num(num)
        elif (libraryID != None) and (documentID != None):
            self.db_load_by_name(libraryID, documentID)
        else:
            self.library_id = ''
            self.quote_dict = {}
        # For Partial Transcript Editing, create a data structure for storing the transcript information by LINE
        self.lines = []
        # Initialize a paragraph counter
        self.paragraphs = 0
        # Create a data structure for tracking very large transcripts by section
        self.paragraphPointers = {}
        # If we have text in the transcript ...
        if ((num != None) or (documentID != None)) and not skipText:
            # ... set up data structures needed for editing large paragraphs
            self.UpdateParagraphs()

# Public methods

    def __repr__(self):
        str = 'Document Object:\n'
        str = str + "number = %s\n" % self.number
        str = str + "id = %s\n" % self.id
        str = str + "library_num = %s\n" % self.library_num
        str += "library_id = %s\n" % self.library_id
        str = str + "author = %s\n" % self.author
        str = str + "Comment = %s\n" % self.comment
        str += "Imported File = %s\n" % self.imported_file
        str += "Import Date = %s\n" % self.import_date
        str += "Document Length = %s  (%s)\n" % (self.document_length, len(self.text))
        str += "isLocked = %s\n" % self._isLocked
        str += "recordlock = %s\n" % self.recordlock
        str += "locktime = %s\n" % self.locktime
        str += "Keywords:\n"
        for kw in self._kwlist:
            str += '  ' + kw.keywordPair + '\n'
        str += "Quote Info:\n"
        keys = self.quote_dict.keys()
        keys.sort()
        for key in keys:
            str += '  %s : %s\n' % (key, self.quote_dict[key])
        str = str + "LastSaveTime = %s\n" % self.lastsavetime
        if len(self.text) > 250:
            str = str + self.text[:250] + '\n\n'   # "text not displayed due to length.\n\n"
        else:
            str = str + "text = %s\n\n" % self.text
        return str.encode('utf8')

    def __eq__(self, other):
        """ Object Equality function """
        if other == None:
            return False
        else:

            if DEBUG:

                print "Document.__eq__():", len(self.__dict__.keys()), len(other.__dict__.keys())

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

    def db_load_by_name(self, libraryID, documentID):
        """Load a record by ID / Name."""
        # If we're in Unicode mode, we need to encode the parameters so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            libraryID = libraryID.encode(TransanaGlobal.encoding)
            documentID = documentID.encode(TransanaGlobal.encoding)
        # Get a database connection
        db = DBInterface.get_db()
        # Craft a query to get Document data
        query = """SELECT * FROM Documents2 a, Series2 b
            WHERE   DocumentID = %s AND
                    a.LibraryNum = b.SeriesNum AND
                    b.SeriesID = %s
        """
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Get a database cursor
        c = db.cursor()
        # Execute the query
        c.execute(query, (documentID, libraryID))
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
            raise RecordNotFoundError, (documentID, n)
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
                raise RecordNotFoundError, (documentID, 0)
            # Load the data into the Document object
            self._load_row(r)
            # Refresh the Keywords
            self.refresh_keywords()
            # Refresh the Quote Position Dictionary
            self.refresh_quotes()
        # Close the Database cursor
        c.close()

    def db_load_by_num(self, num):
        """Load a record by record number."""
        # Get the database connection
        db = DBInterface.get_db()
        # If we're skipping the XML Text ...
        if self.skipText:
            # Define the query to load a Document without text
            query = """SELECT DocumentNum, DocumentID, LibraryNum, SeriesID, Author, Comment,
                              ImportedFile, ImportDate, DocumentLength, 
                              a.RecordLock, a.LockTime, LastSaveTime
                         FROM Documents2 a, Series2 b
                         WHERE   DocumentNum = %s AND
                                 a.LibraryNum = b.SeriesNum
                    """
        # If we're NOT skipping the RTF Text ...
        else:
            # Define the query to load a Document with everything
            query = """SELECT DocumentNum, DocumentID, LibraryNum, SeriesID, Author, Comment,
                              ImportedFile, ImportDate, DocumentLength, XMLText, 
                              a.RecordLock, a.LockTime, LastSaveTime
                         FROM Documents2 a, Series2 b
                         WHERE   DocumentNum = %s AND
                                 a.LibraryNum = b.SeriesNum
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
            # Load the data into the Transcript Object
            self._load_row(r)
            # If sqlite and not results are found ...
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}): 
                # ... close the database cursor ...
                c.close()
                # ... clear the current Transcript object ...
                self.clear()
                # ... and raise an exception
                raise RecordNotFoundError, (num, 0)
            # Refresh the Keywords
            self.refresh_keywords()
            # Refresh the Quote Position Dictionary
            self.refresh_quotes()
        # Close the database cursor
        c.close()

    def UpdateParagraphs(self):
        """ This method divides XML text up into paragraphs, needed for editing LONG documents """
        # Initialize (or re-initialize) the paragraph pointers dictionary
        self.paragraphPointers = {}
        # If there's a defined (saved) transcript object ...
        if (self.id != 0) and (self.text != None):
            
#            print "Document.UpdateParagraphs()"

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

    def db_save(self, use_transactions=True, ignore_filename=False):
        """Save the record to the database using Insert or Update as appropriate."""

        # Define and implement Demo Version limits
        if TransanaConstants.demoVersion and (self.number == 0):
            # Get a DB Cursor
            c = DBInterface.get_db().cursor()
            # Find out how many Document records exist
            c.execute('SELECT COUNT(DocumentNum) FROM Documents2')
            res = c.fetchone()
            c.close()
            # Define the maximum number of records allowed
            maxDocuments = TransanaConstants.maxDocuments
            # Compare
            if res[0] >= maxDocuments:
                # If the limit is exceeded, create and display the error using a SaveError exception
                prompt = _('The Transana Demonstration limits you to %d Document records.\nPlease cancel the "Add Document" dialog to continue.')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                raise SaveError, prompt % maxDocuments

        if DEBUG:
            print "Document.db_save():  %s\n%s" % (self.text, type(self.text))

        # Sanity checks
        if (self.id == ""):
            raise SaveError, _("Document ID is required.")
        elif self.library_num == 0:
            raise SaveError, _("This Document is not associated properly with a Library.")
        # If this is a NEW Document and a Filename has been provided, check to see if the file exists!
        elif (self.number == 0) and (not ignore_filename) and (self.imported_file != "") and (not os.path.exists(self.imported_file)):
            prompt = unicode(_('File "%s" cannot be found.  Try using the "Browse" button.'), 'utf8')
            raise SaveError, prompt % self.imported_file
        else:
            # Get a database cursor
            c = DBInterface.get_db().cursor()


            # If we're using transactions, start the transaction
            if use_transactions:
                c.execute('BEGIN')


            # If we're in Unicode mode, ...
            if 'unicode' in wx.PlatformInfo:
                # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
                # variables for the data.  We don't want to change the underlying object values.  Also, this way,
                # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
                id = self.id.encode(TransanaGlobal.encoding)
                # If the author is None in the database, we can't encode that!  Otherwise, we should encode.
                if self.author != None:
                    author = self.author.encode(TransanaGlobal.encoding)
                else:
                    author = self.author
                # If the comment is None in the database, we can't encode that!  Otherwise, we should encode.
                if self.comment != None:
                    comment = self.comment.encode(TransanaGlobal.encoding)
                else:
                    comment = self.comment
                # If the Imported File Name is None in the database, we can't encode that!  Otherwise, we should encode.
                if self.imported_file != None:
                    imported_file = self.imported_file.encode(TransanaGlobal.encoding)
                else:
                    imported_file = self.imported_file
            else:
                # If we don't need to encode the string values, we still need to copy them to our local variables.
                id = self.id
                author = self.author
                comment = self.comment
                imported_file = self.imported_file

            if (len(self.text) > TransanaGlobal.max_allowed_packet):   # 8388000
                raise SaveError, _("This document is too large for the database.  Please shorten it, split it into two parts\nor if you are importing an RTF document, remove some unnecessary RTF encoding.")

            fields = ("DocumentID", "LibraryNum", "Author", "Comment", "ImportedFile", "DocumentLength", "XMLText", "LastSaveTime")
            values = (id, self.library_num, author, comment, imported_file, self.document_length, self.text)

            if (self._db_start_save() == 0):
                # Duplicate Document IDs within a Library are not allowed.
                if DBInterface.record_match_count("Documents2", ("DocumentID", "LibraryNum"), (id, self.library_num)) > 0:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('A Document named "%s" already exists in this Library.\nPlease enter a different Document ID.'), 'utf8')
                    else:
                        prompt = _('A Document named "%s" already exists in this Library.\nPlease enter a different Document ID.')
                    raise SaveError, prompt % self.id
                # Duplicate Document ID with an Episode ID within a Library are not allowed.
                if DBInterface.record_match_count("Episodes2", ("EpisodeID", "SeriesNum"), (id, self.library_num)) > 0:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('An Episode named "%s" already exists in this Library.\nPlease enter a different Document ID.'), 'utf8')
                    else:
                        prompt = _('An Episode named "%s" already exists in this Library.\nPlease enter a different Document ID.')
                    raise SaveError, prompt % self.id

                # Add the ImportDate information
                fields += ("ImportDate", )
                values += ('%s' % datetime.datetime.now().strftime('%Y-%m-%d %H:%m:%S'),)

                # insert the new record
                query = "INSERT INTO Documents2\n("
                for field in fields:
                    query = "%s%s," % (query, field)
                query = query[:-1] + ')'
                query = "%s\nVALUES\n(" % query
                for value in values:
                    query = "%s%%s," % query
                # The last data value should be the SERVER's time stamp because we don't know if the clients are synchronized.
                # Even a couple minutes difference can cause problems, but with time zones, the different could be hours!
                query += 'CURRENT_TIMESTAMP)'

                if DEBUG:
                    import Dialogs
                    msg = query % values
                    dlg = Dialogs.InfoDialog(None, msg)
                    dlg.ShowModal()
                    dlg.Destroy()
                # Adjust the query for sqlite if needed
                query = DBInterface.FixQuery(query)
                # Execure the Save query
                c.execute(query, values)

            else:
                # Duplicate Document IDs within a Library are not allowed.
                if DBInterface.record_match_count("Documents2", ("DocumentID", "!DocumentNum", "LibraryNum"),
                        (id, self.number, self.library_num)) > 0:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('A Document named "%s" already exists in this Library.\nPlease enter a different Document ID.'), 'utf8')
                    else:
                        prompt = _('A Document named "%s" already exists in this Library.\nPlease enter a different Document ID.')
                    raise SaveError, prompt % self.id
                # Duplicate Document ID with Episode ID within a Library are not allowed.
                if DBInterface.record_match_count("Episodes2", ("EpisodeID", "SeriesNum"),
                        (id, self.library_num)) > 0:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('An Episode named "%s" already exists in this Library.\nPlease enter a different Document ID.'), 'utf8')
                    else:
                        prompt = _('An Episode named "%s" already exists in this Library.\nPlease enter a different Document ID.')
                    raise SaveError, prompt % self.id
                
                # OK to update the episode record
                query = """UPDATE Documents2
                    SET DocumentID = %s,
                        LibraryNum = %s,
                        Author = %s,
                        Comment = %s,
                        ImportedFile = %s,
                        DocumentLength = %s,
                        XMLText = %s,
                        LastSaveTime = CURRENT_TIMESTAMP
                    WHERE DocumentNum = %s
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

                # Execure the Save query
                c.execute(query, values)

                # We need to add the Start and End Character Position information 
                # to the QuotePositions Table.  (A new Document won't have any Quotes yet.)

                # First, delete old data if it exists
                query = "DELETE FROM QuotePositions2 WHERE DocumentNum = %s"
                # Adjust the query for sqlite if needed
                query = DBInterface.FixQuery(query)
                # Execute the query
                c.execute(query, (self.number, ))
                # If there are Quotes in the Quote Dictionary ...
                if len(self.quote_dict) > 0:
                    # Add MySQL-specific bulk insert SQL
                    if TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
                        query = "INSERT INTO QuotePositions2 (QuoteNum, DocumentNum, StartChar, EndChar) VALUES "
                        values = ()
                        for key in self.quote_dict.keys():
                            query += "(%s, %s, %s, %s), "
                            values += (key, self.number, self.quote_dict[key][0], self.quote_dict[key][1])
                        # Strip the final comma off the query!
                        query = query[:-2]
                        # Adjust the query for sqlite if needed
                        query = DBInterface.FixQuery(query)
                        # Execute the query
                        c.execute(query, values)
                    else:

#                        print "Document.db_save():  QuotePositions not BULK inserted for sqlite yet!!"

                        query = "INSERT INTO QuotePositions2 (QuoteNum, DocumentNum, StartChar, EndChar) VALUES "
                        query += "(%s, %s, %s, %s) "
                        for key in self.quote_dict.keys():
                            values = (key, self.number, self.quote_dict[key][0], self.quote_dict[key][1])

#                            print
#                            print
#                            print "Document.db_save():"
#                            print query
#                            print values
#                            print

                            # Adjust the query for sqlite if needed
                            query = DBInterface.FixQuery(query)
                            # Execute the query
                            c.execute(query, values)

            # If the object number is 0, we have a new object
            if self.number == 0:
                numberChanged = True
                # Load the auto-assigned new number record if necessary and the saved time.
                query = """
                          SELECT DocumentNum, ImportDate, LastSaveTime FROM Documents2
                          WHERE DocumentID = %s AND
                                LibraryNum = %s
                        """
                # Assemble the arguments for the query
                args = (id, self.library_num)
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
                    # ... load the Import Date
                    self.import_date = recs[0][1]
                    # ... and update the LastSaveTime
                    self.lastsavetime = recs[0][2]
                # If we don't have a single record ...
                else:
                    # ... raise an exception
                    raise RecordNotFoundError, (self.id, len(recs))

                # Close the temporary database cursor
                tempDBCursor.close()
            # If we're updating an EXISTING record ... 
            else:
                numberChanged = False
                # Load the NEW last saved time.
                query = """
                          SELECT LastSaveTime FROM Documents2
                          WHERE DocumentNum = %s
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
                DBInterface.delete_all_keywords_for_a_group(0, self.number, 0, 0, 0)

            # Initialize a blank error prompt
            prompt = ''
            # Add the Document keywords back.  Iterate through the Keyword List
            for kws in self._kwlist:
                # Try to add the Clip Keyword record.  If it is NOT added, the keyword has been changed by another user!
                if not DBInterface.insert_clip_keyword(0, self.number, 0, 0, 0, kws.keywordGroup, kws.keyword, kws.example):
                    # if the prompt isn't blank ...
                    if prompt != '':
                        # ... add a couple of line breaks to it
                        prompt += u'\n\n'
                    # Add the current keyword to the error prompt
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt += unicode(_('Keyword "%s : %s" cannot be added to Document "%s".\nAnother user must have edited the keyword while you were adding it.'), 'utf8') % (kws.keywordGroup, kws.keyword, self.id)


            # If there is an error prompt ...
            if prompt != '':
                # If the Episode Number was changed ...
                if numberChanged:
                    # ... change it back to zero!!
                    self.number = 0
                if use_transactions:
                    # Undo the database save transaction
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
                if use_transactions:
                    # ... Commit the database transaction
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

            # Delete all Document-based Filter Configurations
            #   Delete Document Keyword Map records
            DBInterface.delete_filter_records(17, self.number)
            #   Delete Document Keyword Visualization records
            DBInterface.delete_filter_records(18, self.number)
            #   Delete Document Report records
            DBInterface.delete_filter_records(19, self.number)
            #   Delete Document Data Export records
            DBInterface.delete_filter_records(20, self.number)
            
            # Look for Quotes before deleting a Document!
            quotes = self.quote_dict

            # If there are quotes in the list ...
            if len(quotes) > 0:
                # ... build a prompt for the warning dialog box
                prompt = _('Quotes have been created from Document "%s".\nThese quotes will become orphaned if you delete the document.\nDo you want to delete this document anyway?')
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
                    prompt = _('The delete has been cancelled.  Quotes have been created from Document "%s".')
                    # ... encode the prompt if needed ...
                    if 'unicode' in wx.PlatformInfo:
                        prompt = unicode(prompt, 'utf8')
                    # ... and use a DeleteError Exception to interupt the deletion.
                    raise DeleteError(prompt % self.id)

            # Quotes with this Document for the Source number need to be cleared from Quote
            # SourceDocumentNumber records.  This must be done LAST, after everything else.
            if result:
                DBInterface.ClearSourceDocumentRecords(self.number)

            # Fix the QuotePosition records
            query = """UPDATE QuotePositions2
                         SET DocumentNum = 0,
                             StartChar = -1,
                             EndChar = -1
                       WHERE DocumentNum = %s
                    """
            # Get a temporarly database cursor
            tempDBCursor = DBInterface.get_db().cursor()
            # Adjust the query for sqlite if needed
            query = DBInterface.FixQuery(query)
            # Execute the query
            tempDBCursor.execute(query, (self.number, ))

            notes = []

            # Detect, Load, and Delete all Document Notes.
            notes = self.get_note_nums()
            for note_num in notes:
                note = Note.Note(note_num)
                result = result and note.db_delete(0)
                del note
            del notes

            # Delete all related references in the ClipKeywords table
            if result:
                DBInterface.delete_all_keywords_for_a_group(0, self.number, 0, 0, 0)

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
        kwpairs = DBInterface.list_of_keywords(Document=self.number)
        for data in kwpairs:
            tempClipKeyword = ClipKeywordObject.ClipKeyword(data[0], data[1], documentNum=self.number, example=data[2])
            self._kwlist.append(tempClipKeyword)
        
    def add_keyword(self, kwg, kw):
        """Add a keyword to the keyword list."""
        # We need to check to see if the keyword is already in the keyword list
        keywordFound = False
        # Iterate through the list
        for documentKeyword in self._kwlist:
            # If we find a match, set the flag and quit looking.
            if (documentKeyword.keywordGroup == kwg) and (documentKeyword.keyword == kw):
                keywordFound = True
                break

        # If the keyword is not found, add it.  (If it's already there, we don't need to do anything!)
        if not keywordFound:
            # Create an appropriate ClipKeyword Object
            tempClipKeyword = ClipKeywordObject.ClipKeyword(kwg, kw, documentNum=self.number)
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

    def clear_quotes(self):
        """ Clear the Quote List """
        # Clear the Quote List
        self.quote_dict = {}

    def refresh_quotes(self):
        # Clear the Quote List
        self.clear_quotes()
        # Get the database connection
        db = DBInterface.get_db()
        # Get a database cursor
        c = db.cursor()
        # Define the Query.  No need for an ORDER BY clause!
        query = "SELECT QuoteNum, StartChar, EndChar FROM QuotePositions2 WHERE DocumentNum = %s"
        # Adjust the query for sqlite if needed
        query = DBInterface.FixQuery(query)
        # Execute the Query
        c.execute(query, (self.number, ))
        # Get the query results
        results = c.fetchall()
        # For each found Quote Position record ...
        for rec in results:
            # ... add the Quote to the Quote List
            self.add_quote(rec[0], rec[1], rec[2])

    def add_quote(self, quoteNum, startChar, endChar):
        """ Add a Quote to the Quote List """
        # Add or update the value in the Quote List
        self.quote_dict[quoteNum] = (startChar, endChar)

    def delete_quote(self, quoteNum):
        # If the current Quote is in the Document's Quote List ...
        if self.quote_dict.has_key(quoteNum):
            # ... delete it!
            del(self.quote_dict[quoteNum])

    def lock_record(self):
        """ Override the DataObject Lock Method """
        # If we're using the single-user version of Transana, we just need to ...
        if not TransanaConstants.singleUserVersion:
            # ... confirm that the document has not been altered by another user since it was loaded.
            # To do this, first pull the LastSaveTime for this record from the database.
            query = """
                      SELECT LastSaveTime FROM Documents2
                      WHERE DocumentNum = %s
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
    # Implementation for Library Number Property
    def _get_library_num(self):
        return self._library_num
    def _set_library_num(self, num):
        self._library_num = num
        if self.library_id == '':
            if num > 0:
                tmp = Library.Library(self.library_num)
                self.library_id = tmp.id
            else:
                self.library_id = ''
    def _del_library_num(self):
        self._library_num = 0
        self._library_id = ''

    # Implementation for Author Property
    def _get_author(self):
        return self._author
    def _set_author(self, name):
        self._author = name
    def _del_author(self):
        self._author = ''

    # Implementation for Text Property
    def _get_text(self):
        return self._text
    def _set_text(self, txt):
        self._text = txt
    def _del_text(self):
        self._text = ''

    # Implementation for Imported File Name Property
    def _get_imported_file(self):
        return self._imported_file
    def _set_imported_file(self, txt):
        if not txt is None:
            self._imported_file = txt
        else:
            self._imported_file = ''
    def _del_imported_file(self):
        self._imported_file = ''

    # Implementation for Import Date Property - NOTE that this does NOT try to change the value's type!
    def _get_import_date(self):
        return self._import_date
    def _set_import_date(self, txt):
        self._import_date = txt
    def _del_import_date(self):
        self._import_date = None

    # Implement the Document_Length Property
    def _get_document_length(self):
        return self._document_length
    def _set_document_length(self, length):
        self._document_length = length
    def _del_document_length(self):
        self._document_length = 0

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
    library_num = property(_get_library_num, _set_library_num, _del_library_num,
                        """The Library number, if associated with one.""")
    author = property(_get_author, _set_author, _del_author,
                        """The person who wrote the Document.""")
    text = property(_get_text, _set_text, _del_text,
                        """Text of the transcript, stored in the database as a BLOB.""")
    imported_file = property(_get_imported_file, _set_imported_file, _del_imported_file,
                             """Imported File Name""")
    import_date = property(_get_import_date, _set_import_date, _del_import_date,
                             """Import Date for the File""")
    document_length = property(_get_document_length, _set_document_length, _del_document_length,
                               """Document Length in Characters""")
    locked_by_me = property(None, None, None,
                        """Determines if this instance owns the Document lock.""")
    has_changed = property(_get_changed, _set_changed, _del_changed,
                        """Indicates whether the Document has been modified.""")
    lastsavetime = property(_get_lastsavetime, _set_lastsavetime, _del_lastsavetime,
                        """The timestamp of the last save (MU only).""")
    keyword_list = property(_get_kwlist, _set_kwlist, _del_kwlist,
                        """The list of keywords that have been applied to the Document.""")

    def _load_row(self, row):
    	self.number = row['DocumentNum']
        self.id = row['DocumentID']
        self.library_id = row['SeriesID']
        self.library_num = row['LibraryNum']
        self.author = row['Author']

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

        self.comment = row['Comment']
        self.imported_file = row['ImportedFile']
        self.import_date = row['ImportDate']
        self.document_length = row['DocumentLength']
        self.lastsavetime = row['LastSaveTime']
        self.changed = False

        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.library_id = DBInterface.ProcessDBDataForUTF8Encoding(self.library_id)
            self.author = DBInterface.ProcessDBDataForUTF8Encoding(self.author)
            self.comment = DBInterface.ProcessDBDataForUTF8Encoding(self.comment)
            self.imported_file = DBInterface.ProcessDBDataForUTF8Encoding(self.imported_file)
