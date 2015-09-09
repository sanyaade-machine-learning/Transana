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

"""This module implements the Keyword class as part of the Data Objects."""

# NOTE:  While the Keyword object is very SIMILAR to all the other Data Objects
#        in Transana, it will NOT be inherited from the DataObject class.  The
#        reason for this is that the DataObject class requires the database
#        record to have a RecordNumber field, and the Keywords Table does not
#        have (or at this time need) this field.  The following code is
#        obviously heavily based on DataObject.   DKW

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

import wx
from TransanaExceptions import *
import DBInterface
import Dialogs
import Misc
import TransanaConstants
import TransanaGlobal
import inspect
# import Python String module
import string
import types


class Keyword(object):
    """This class defines the structure for a Keyword object.  A Keyword object
    holds information about a Keyword within a Keyword Group."""

    def __init__(self, keywordGroup=None, keyword=None):
        """Initialize a Keyword object.  If a Keyword Group and Keyword
        are given, load it from the database."""
        self.clear()
        # Intialize values that must exist even if nothing is loaded.  Do this before loading!
        # Original Keyword Group and Original keyword come from the parameters passed in
        self.originalKeywordGroup = keywordGroup
        self.originalKeyword = keyword
        if isinstance(keywordGroup, types.StringTypes) and \
           isinstance(keyword, types.StringTypes):
            self.db_load_by_name(keywordGroup, keyword)

# Public methods

    def __repr__(self):
        str = "Keyword Object:\n"
        str = str + "Keyword Group = %s  (Original = %s)\n" % (self.keywordGroup, self.originalKeywordGroup)
        str = str + "Keyword = %s  (Original = %s)\n" % (self.keyword, self.originalKeyword)
        str = str + "Definition = %s\n" % self.definition
        str += "lineColorName = %s (%s)\n" % (self.lineColorName, self.lineColorDef)
        str += "drawMode = %s\n" % self.drawMode
        str += "lineWidth = %s\n" % self.lineWidth
        str += "lineStyle = %s\n" % self.lineStyle
#        str += "isLocked = %s\n" % self._isLocked
#        str += "recordlock = %s\n" % self.recordlock
#        str += "locktime = %s\n" % self.locktime
        return str

    def __eq__(self, other):
        if other == None:
            return False
        else:

            if self.__dict__ != other.__dict__:
                print "Keyword Objects NOT EQUAL:"
                print len(self.__dict__), len(other.__dict__)
                for key in self.__dict__.keys():
                    print key, type(self.__dict__[key]), type(other.__dict__[key]), self.__dict__[key] == other.__dict__[key]
                    
            return self.__dict__ == other.__dict__

    def clear(self):
        """Clear all properties, resetting them to default values."""
        attrdesc = inspect.classify_class_attrs(self.__class__)
        for attr in attrdesc:
            if attr[1] == "property":
                try:
                    delattr(self, attr[0])
                except AttributeError, e:
                    pass # probably a non-deletable attribute

    def checkEpisodesClipsSnapshotsForLocks(self):
        """ Checks Episodes, Clips, and Snapshots to see if a Keyword record is free of related locks """
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            originalKeywordGroup = self.originalKeywordGroup.encode(TransanaGlobal.encoding)
            originalKeyword = self.originalKeyword.encode(TransanaGlobal.encoding)
        else:
            originalKeywordGroup = self.originalKeywordGroup
            originalKeyword = self.originalKeyword

        # query the lock status for Documents that contain the Keyword 
        query = """SELECT c.RecordLock FROM Keywords2 a, ClipKeywords2 b, Documents2 c
                     WHERE a.KeywordGroup = %s AND
                           a.Keyword = %s AND
                           a.KeywordGroup = b.KeywordGroup AND
                           a.Keyword = b.Keyword AND
                           b.DocumentNum <> 0 AND
                           b.DocumentNum = c.DocumentNum AND
                           (c.RecordLock <> '' AND
                            c.RecordLock IS NOT NULL)"""
        values = (originalKeywordGroup, originalKeyword)
        c = DBInterface.get_db().cursor()
        query = DBInterface.FixQuery(query)
        c.execute(query, values)
        result = c.fetchall()
        RecordCount = len(result)
        c.close()


        # If no Document that contain the record are locked, check the lock status of Episodes that contain the Keyword
        if RecordCount == 0:
            # query the lock status for Episodes that contain the Keyword 
            query = """SELECT c.RecordLock FROM Keywords2 a, ClipKeywords2 b, Episodes2 c
                         WHERE a.KeywordGroup = %s AND
                               a.Keyword = %s AND
                               a.KeywordGroup = b.KeywordGroup AND
                               a.Keyword = b.Keyword AND
                               b.EpisodeNum <> 0 AND
                               b.EpisodeNum = c.EpisodeNum AND
                               (c.RecordLock <> '' AND
                                c.RecordLock IS NOT NULL)"""
            values = (originalKeywordGroup, originalKeyword)
            c = DBInterface.get_db().cursor()
            query = DBInterface.FixQuery(query)
            c.execute(query, values)
            result = c.fetchall()
            RecordCount = len(result)
            c.close()

        # If no Documents or Episodes that contain the record are locked, check the lock status of Quotes that
        # contain the Keyword
        if RecordCount == 0:
            # query the lock status for Quotes that contain the Keyword 
            query = """SELECT c.RecordLock FROM Keywords2 a, ClipKeywords2 b, Quotes2 c
                         WHERE a.KeywordGroup = %s AND
                               a.Keyword = %s AND
                               a.KeywordGroup = b.KeywordGroup AND
                               a.Keyword = b.Keyword AND
                               b.QuoteNum <> 0 AND
                               b.QuoteNum = c.QuoteNum AND
                               (c.RecordLock <> '' AND
                                c.RecordLock IS NOT NULL)"""
            values = (originalKeywordGroup, originalKeyword)
            c = DBInterface.get_db().cursor()
            query = DBInterface.FixQuery(query)
            c.execute(query, values)
            result = c.fetchall()
            RecordCount = len(result)
            c.close()

        # If no Documents, Episodes, or Quotes that contain the record are locked, check the lock status of
        # Clips that contain the Keyword
        if RecordCount == 0:
            # query the lock status for Clips that contain the Keyword 
            query = """SELECT c.RecordLock FROM Keywords2 a, ClipKeywords2 b, Clips2 c
                         WHERE a.KeywordGroup = %s AND
                               a.Keyword = %s AND
                               a.KeywordGroup = b.KeywordGroup AND
                               a.Keyword = b.Keyword AND
                               b.ClipNum <> 0 AND
                               b.ClipNum = c.ClipNum AND
                               (c.RecordLock <> '' AND
                                c.RecordLock IS NOT NULL)"""
            values = (originalKeywordGroup, originalKeyword)
            c = DBInterface.get_db().cursor()
            query = DBInterface.FixQuery(query)
            c.execute(query, values)
            result = c.fetchall()
            RecordCount = len(result)
            c.close()

        # If no Documents, Episodes, Quotes, or Clips that contain the record are locked, check the lock status
        # of Snapshots that contain the Keyword for Whole Snapshot coding
        if RecordCount == 0:
            # query the lock status for Snapshots that contain the Keyword 
            query = """SELECT c.RecordLock FROM Keywords2 a, ClipKeywords2 b, Snapshots2 c
                         WHERE a.KeywordGroup = %s AND
                               a.Keyword = %s AND
                               a.KeywordGroup = b.KeywordGroup AND
                               a.Keyword = b.Keyword AND
                               b.SnapshotNum <> 0 AND
                               b.SnapshotNum = c.SnapshotNum AND
                               (c.RecordLock <> '' AND
                                c.RecordLock IS NOT NULL)"""
            values = (originalKeywordGroup, originalKeyword)
            c = DBInterface.get_db().cursor()
            query = DBInterface.FixQuery(query)
            c.execute(query, values)
            result = c.fetchall()
            RecordCount = len(result)
            c.close()

        # If no Documents, Episodes, Quotes, Clips, or whole Snapshots that contain the record are locked,
        # check the lock status of Snapshots that contain the Keyword for Snapshot Coding
        if RecordCount == 0:
            # query the lock status for Snapshots that contain the Keyword 
            query = """SELECT c.RecordLock FROM Keywords2 a, SnapshotKeywords2 b, Snapshots2 c
                         WHERE a.KeywordGroup = %s AND
                               a.Keyword = %s AND
                               a.KeywordGroup = b.KeywordGroup AND
                               a.Keyword = b.Keyword AND
                               b.SnapshotNum <> 0 AND
                               b.SnapshotNum = c.SnapshotNum AND
                               (c.RecordLock <> '' AND
                                c.RecordLock IS NOT NULL)"""
            values = (originalKeywordGroup, originalKeyword)
            c = DBInterface.get_db().cursor()
            query = DBInterface.FixQuery(query)
            c.execute(query, values)
            result = c.fetchall()
            RecordCount = len(result)
            c.close()

        if RecordCount != 0:
            LockName = result[0][0]
            return (RecordCount, LockName)
        else:
            return (RecordCount, None)
        
    def removeDuplicatesForMerge(self):
        """  When merging keywords, we need to remove instances of the OLD keyword that already exist in
             Documents, Episodes, Quotes, Clips, or Snapshots that also contain the NEW keyword.  (Doing
             it this way reduces overhead.) """
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
            # variables for the data.  We don't want to change the underlying object values.  Also, this way,
            # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
            originalKeywordGroup = self.originalKeywordGroup.encode(TransanaGlobal.encoding)
            originalKeyword = self.originalKeyword.encode(TransanaGlobal.encoding)
            keywordGroup = self.keywordGroup.encode(TransanaGlobal.encoding)
            keyword = self.keyword.encode(TransanaGlobal.encoding)
        else:
            originalKeywordGroup = self.originalKeywordGroup
            originalKeyword = self.originalKeyword
            keywordGroup = self.keywordGroup
            keyword = self.keyword

        # Look for Episodes, Clips, and Whole-Snapshots that have BOTH the original and the merge keywords
        query = """SELECT a.EpisodeNum, a.DocumentNum, a.ClipNum, a.QuoteNum, a.SnapshotNum, a.KeywordGroup, a.Keyword
                     FROM ClipKeywords2 a, ClipKeywords2 b
                     WHERE a.KeywordGroup = %s AND
                           a.Keyword = %s AND
                           b.KeywordGroup = %s AND
                           b.Keyword = %s AND
                           a.EpisodeNum = b.EpisodeNum AND
                           a.DocumentNum = b.DocumentNum AND
                           a.ClipNum = b.ClipNum AND
                           a.QuoteNum = b.QuoteNum AND
                           a.SnapshotNum = b.SnapshotNum"""
        values = (originalKeywordGroup, originalKeyword, keywordGroup, keyword)
        c = DBInterface.get_db().cursor()
        query = DBInterface.FixQuery(query)
        c.execute(query, values)
        # Remember the list of what would become duplicate entries
        result = c.fetchall()

        # Prepare a query for deleting the duplicate
        query = """ DELETE FROM ClipKeywords2
                      WHERE EpisodeNum = %s AND
                            DocumentNum = %s AND
                            ClipNum = %s AND
                            QuoteNum = %s AND 
                            SnapshotNum = %s AND
                            KeywordGroup = %s AND
                            Keyword = %s """
        query = DBInterface.FixQuery(query)
        # Go through the list of duplicates ...
        for line in result:
            # ... and delete the original keyword listing, leaving the other (merge) record untouched.
            values = (line[0], line[1], line[2], line[3], line[4], line[5], line[6])
            c.execute(query, values)

        # For Snapshot Coding, we don't want to LOSE any of the drawn shapes, so we rename the OLD
        # Keyword Records to match the new Keyword Records
        query = """UPDATE SnapshotKeywords2
                     SET KeywordGroup = %s,
                         Keyword = %s
                     WHERE KeywordGroup = %s AND
                           Keyword = %s """
        values = (keywordGroup, keyword, originalKeywordGroup, originalKeyword)
        query = DBInterface.FixQuery(query)
        c.execute(query, values)
            
        # Look for Snapshots that have STYLES for BOTH the original and the merge keywords
        query = """SELECT * FROM SnapshotKeywordStyles2 a, SnapshotKeywordStyles2 b
                     WHERE a.KeywordGroup = %s AND
                           a.Keyword = %s AND
                           b.KeywordGroup = %s AND
                           b.Keyword = %s AND
                           a.SnapshotNum = b.SnapshotNum """
        values = (originalKeywordGroup, originalKeyword, keywordGroup, keyword)
        query = DBInterface.FixQuery(query)
        c.execute(query, values)
        # Remember the list of what would become duplicate entries
        result = c.fetchall()

        # Prepare a query for deleting the duplicate
        query = """ DELETE FROM SnapshotKeywordStyles2
                      WHERE SnapshotNum = %s AND
                            KeywordGroup = %s AND
                            Keyword = %s """
        query = DBInterface.FixQuery(query)
        # Go through the list of duplicates ...
        for line in result:
            # ... and delete the original keyword listing, leaving the other (merge) record untouched.
            values = (line[0], line[1], line[2])
            c.execute(query, values)
            
        c.close()
        

    def lock_record(self):
        """Lock a record.  If the lock is unable to be obtained, a
        RecordLockedError exception is raised with the username of the lock
        holder passed."""

        if (self.originalKeywordGroup == None) or \
           (self.originalKeyword == None):           # no record loaded?
            return

        tablename = self._table()
        
        db = DBInterface.get_db()
        c = db.cursor()

        lq = self._get_db_fields(('RecordLock', 'LockTime'), c)
        if (lq[1] == None) or (lq[0] == "") or ((DBInterface.ServerDateTime() - lq[1]).days > 1):
            (EpisodeClipLockCount, LockName) = self.checkEpisodesClipsSnapshotsForLocks()
            if EpisodeClipLockCount == 0:
                # Lock the record
                self._set_db_fields(    ('RecordLock', 'LockTime'),
                                        (DBInterface.get_username(),
                                        str(DBInterface.ServerDateTime())[:-3]), c)
                c.close()
            else:
                c.close()
                raise RecordLockedError, LockName
        else:
            # We just raise an exception here since GUI code isn't appropriate.
            c.close()
            raise RecordLockedError, lq[0]  # Pass name of person

        
    def unlock_record(self):
        """Unlock a record."""

        self._set_db_fields(    ('RecordLock', 'LockTime'),
                                ('', None), None)

    def db_load_by_name(self, keywordGroup, keyword):
        """Load a record.  Raise a RecordNotFound exception
        if record is not found in database."""
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            keywordGroup = keywordGroup.encode(TransanaGlobal.encoding)
            keyword = keyword.encode(TransanaGlobal.encoding)

        db = DBInterface.get_db()
        query = """
        SELECT * FROM Keywords2
            WHERE KeywordGroup = %s AND
                  Keyword = %s
        """
        query = DBInterface.FixQuery(query)
        c = db.cursor()
        c.execute(query, (keywordGroup, keyword))

        # rowcount doesn't work for sqlite!
        if TransanaConstants.DBInstalled == 'sqlite3':
            n = 1
        else:
            n = c.rowcount
        if (n != 1):
            c.close()
            self.clear()
            raise RecordNotFoundError, (keywordGroup + ':' + keyword, n)
        else:
            r = DBInterface.fetch_named(c)
            if (TransanaConstants.DBInstalled == 'sqlite3') and (r == {}): 
                c.close()
                self.clear()
                raise RecordNotFoundError, (keywordGroup + ':' + keyword, 0)
            self._load_row(r)
            

        c.close()

    # TODO:  We need to change the model for all Data Object Saves.  This needs to be
    #        a boolean function so that if the Save fails, the Properties Dialog can
    #        remain open, giving the user a chance to fix whatever caused the Save to
    #        fail.  Remember to do this here AND in DataObject.
    def db_save(self, use_transactions=True):
        """Save the record to the database using Insert or Update as
        appropriate."""
        # Define and implement Demo Version limits
        if TransanaConstants.demoVersion and (self._db_start_save() == 0):
            # Get a DB Cursor
            c = DBInterface.get_db().cursor()
            # Find out how many records exist
            c.execute('SELECT COUNT(Keyword) from Keywords2')
            res = c.fetchone()
            c.close()
            # Define the maximum number of records allowed
            maxKeywords = TransanaConstants.maxKeywords
            # Compare
            if res[0] >= maxKeywords:
                # If the limit is exceeded, create and display the error using a SaveError exception
                prompt = _('The Transana Demonstration limits you to %d Keyword records.\nPlease cancel the "Add Keyword" dialog to continue.')
                if 'unicode' in wx.PlatformInfo:
                    prompt = unicode(prompt, 'utf8')
                raise SaveError, prompt % maxKeywords

        # Validity Checks
        if (self.keywordGroup == ''):
            raise SaveError, _('Keyword Group is required.')
        elif (self.keyword == ''):
            raise SaveError, _('Keyword is required.')

        # If we're in Unicode mode, ...
        if 'unicode' in wx.PlatformInfo:
            # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
            # variables for the data.  We don't want to change the underlying object values.  Also, this way,
            # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
            keywordGroup = self.keywordGroup.encode(TransanaGlobal.encoding)
            keyword = self.keyword.encode(TransanaGlobal.encoding)
            if self.originalKeywordGroup != None:
                originalKeywordGroup = self.originalKeywordGroup.encode(TransanaGlobal.encoding)
            else:
                originalKeywordGroup = None
            if self.originalKeyword != None:
                originalKeyword = self.originalKeyword.encode(TransanaGlobal.encoding)
            else:
                originalKeyword = None
            definition = self.definition.encode(TransanaGlobal.encoding)
            lineColorName = self.lineColorName.encode(TransanaGlobal.encoding)
        else:
            # If we don't need to encode the string values, we still need to copy them to our local variables.
            keywordGroup = self.keywordGroup
            keyword = self.keyword
            if self.originalKeywordGroup != None:
                originalKeywordGroup = self.originalKeywordGroup
            else:
                originalKeywordGroup = None
            if self.originalKeyword != None:
                originalKeyword = self.originalKeyword
            else:
                originalKeyword = None
            definition = self.definition
            lineColorName = self.lineColorName

        values = (keywordGroup, keyword, definition, lineColorName, self.lineColorDef, self.drawMode, self.lineWidth, self.lineStyle)

        if (self._db_start_save() == 0):
            # duplicate Keywords are not allowed
            if DBInterface.record_match_count("Keywords2", \
                            ("KeywordGroup", "Keyword"), \
                            (keywordGroup, keyword) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = _('A Keyword named "%s : %s" already exists.')
                    if type(prompt) == str:
                        prompt = unicode(prompt, 'utf8')
                    prompt = prompt % (self.keywordGroup, self.keyword)
                else:
                    prompt = _('A Keyword named "%s : %s" already exists.') % (self.keywordGroup, self.keyword)
                raise SaveError, prompt

            # insert the new Keyword
            query = """
            INSERT INTO Keywords2
                (KeywordGroup, Keyword, Definition, LineColorName, LineColorDef, DrawMode, LineWidth, LineStyle)
                VALUES
                (%s, %s, %s, %s, %s, %s, %s, %s)
            """
            c = DBInterface.get_db().cursor()
            query = DBInterface.FixQuery(query)
            c.execute(query, values)
#            if TransanaConstants.DBInstalled in ['sqlite3']:
#                DBInterface.get_db().commit()
            c.close()
            # When inserting, we're not merging keywords!
            mergeKeywords = False
        else:
            # check for dupes, which are not allowed if either the Keyword Group or Keyword have been changed.
            if (DBInterface.record_match_count("Keywords2", \
                            ("KeywordGroup", "Keyword"), \
                            (keywordGroup, keyword) ) > 0) and \
               ((originalKeywordGroup != keywordGroup) or \
                (originalKeyword.lower() != keyword.lower())):
                # If duplication is found, ask the user if we should MERGE the keywords.
                if 'unicode' in wx.PlatformInfo:
                    oKG = unicode(originalKeywordGroup, 'utf8')
                    oKW = unicode(originalKeyword, 'utf8')
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Keyword named "%s : %s" already exists.  Do you want to merge\n"%s : %s" with "%s : %s"?'), 'utf8') % (self.keywordGroup, self.keyword, oKG, oKW, self.keywordGroup, self.keyword)
                else:
                    prompt = _('A Keyword named "%s : %s" already exists.  Do you want to merge\n"%s : %s" with "%s : %s"?') % (self.keywordGroup, self.keyword, originalKeywordGroup, originalKeyword, self.keywordGroup, self.keyword)
                dlg = Dialogs.QuestionDialog(None,  prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                # If the user wants to merge ...
                if result == wx.ID_YES:
                    # .. then signal the user's desire to merge.
                    mergeKeywords = True
                # If the user does NOT want to merge keywords ...
                else:
                    # ... then signal the user's desire NOT to merge (though this no longer matters!)
                    mergeKeywords = False
                    # ... and raise the duplicate keyword error exception
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('A Keyword named "%s : %s" already exists.'), 'utf8') % (self.keywordGroup, self.keyword)
                    else:
                        prompt = _('A Keyword named "%s : %s" already exists.') % (self.keywordGroup, self.keyword)
                    raise SaveError, prompt
            # If there are NO duplicate keywords ...
            else:
                # ... then signal that there is NO need to merge!
                mergeKeywords = False

            # NOTE:  This is a special instance.  Keywords behave differently than other DataObjects here!
            # Before we can save, we have to check to see if someone locked an Episode or a Clip that
            # uses our Keyword since we obtained the lock on the Keyword.  It doesn't make sense to me
            # to block editing Episode or Clips properties because some else is editing a Keyword it contains.
            # Mostly, editing Keywords will have to do with changing their definition field, not their keyword
            # group or keyword fields.  Because of that, we have to block the save of an altered keyword
            # because the locked Episode or Clip would retain the obsolete keyword and would lose the
            # new keyword record, so it would not appear in all Search results that it should.  I expect
            # this to be extraordinarily rare.
            (EpisodeClipLockCount, LockName) = self.checkEpisodesClipsSnapshotsForLocks()
            if EpisodeClipLockCount != 0 and \
               ((originalKeywordGroup != keywordGroup) or \
                (originalKeyword != keyword)):
                tempstr = _("""You cannot proceed because another user has recently started editing a Document, Episode, Quote, Clip, 
or Snapshot that uses Keyword "%s:%s".  If you change the Keyword now, 
that would corrupt the record that is currently locked by %s.  Please try again later.""")
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    tempstr = unicode(tempstr, 'utf8')
                raise SaveError(tempstr  % (self.originalKeywordGroup, self.originalKeyword, LockName))
            
            else:
                c = DBInterface.get_db().cursor()
                # If we're merging keywords ...
                if mergeKeywords:
                    # We'd better do this as a Transaction!
                    query = 'BEGIN'
                    c.execute(query)
                    # ... then we remove duplicate keywords and we DON'T rename the keyword
                    self.removeDuplicatesForMerge()
                # If we're NOT merging keywords ...
                else:
                    # update the record record with new values
                    query = """
                    UPDATE Keywords2
                        SET KeywordGroup = %s,
                            Keyword = %s,
                            Definition = %s,
                            LineColorName = %s,
                            LineColorDef = %s,
                            DrawMode = %s,
                            LineWidth = %s,
                            LineStyle = %s
                        WHERE KeywordGroup = %s AND
                              Keyword = %s
                    """
                    values = values + (originalKeywordGroup, originalKeyword)
                    query = DBInterface.FixQuery(query)
                    c.execute(query, values)

                # If the Keyword Group or Keyword has changed, we need to update all ClipKeyword records too.
                if ((originalKeywordGroup != keywordGroup) or \
                    (originalKeyword != keyword)):
                    query = """
                    UPDATE ClipKeywords2
                        SET KeywordGroup = %s,
                            Keyword = %s
                        WHERE KeywordGroup = %s AND
                              Keyword = %s
                    """
                    values = (keywordGroup, keyword, originalKeywordGroup, originalKeyword)
                    query = DBInterface.FixQuery(query)
                    c.execute(query, values)

                # If the Keyword Group or Keyword has changed, we need to update all Snapshot Keyword records too.
                if ((originalKeywordGroup != keywordGroup) or \
                    (originalKeyword != keyword)):
                    query = """
                    UPDATE SnapshotKeywords2
                        SET KeywordGroup = %s,
                            Keyword = %s
                        WHERE KeywordGroup = %s AND
                              Keyword = %s
                    """
                    values = (keywordGroup, keyword, originalKeywordGroup, originalKeyword)
                    query = DBInterface.FixQuery(query)
                    c.execute(query, values)

                # If the Keyword Group or Keyword has changed, we need to update all Snapshot Keyword Style records too.
                if ((originalKeywordGroup != keywordGroup) or \
                    (originalKeyword != keyword)):
                    query = """
                    UPDATE SnapshotKeywordStyles2
                        SET KeywordGroup = %s,
                            Keyword = %s
                        WHERE KeywordGroup = %s AND
                              Keyword = %s
                    """
                    values = (keywordGroup, keyword, originalKeywordGroup, originalKeyword)
                    query = DBInterface.FixQuery(query)
                    c.execute(query, values)

                # If we're merging Keywords, we need to DELETE the original keyword and end the transaction
                if mergeKeywords:
                    # Since we've already taken care of the Clip Keywords, we can just delete the keyword!
                    query = """ DELETE FROM Keywords2
                                  WHERE KeywordGroup = %s AND
                                        Keyword = %s"""
                    values = (originalKeywordGroup, originalKeyword)
                    query = DBInterface.FixQuery(query)
                    c.execute(query, values)
                    # If we make it this far, we can commit the transaction, 'cause we're done.
                    query = 'COMMIT'
                    c.execute(query)
##                if TransanaConstants.DBInstalled in ['sqlite3']:
##                    c.commit()
                c.close()
                # If the save is successful, we need to update the "original" values to reflect the new record key.
                # Otherwise, we can't unlock the proper record, among other things.
                self.originalKeywordGroup = self.keywordGroup
                self.originalKeyword = self.keyword
                
        # We need to signal if the we need to update (or delete) the keyword listing in the database tree.
        return not mergeKeywords

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database.  Raises
        RecordLockedError exception if record is locked and unable to be
        deleted."""
        tempstr = "Delete Keyword object has not been implemented."
        dlg = Dialogs.InfoDialog(TransanaGlobal.menuWindow, tempstr)
        dlg.ShowModal()
        dlg.Destroy()


# Private methods

    def _table(self):
        """Return the SQL table name."""
        t = "Keywords2"
        return t

    def _get_db_fields(self, fieldlist, c=None):
        """Get the values of fields from the database for the currently
        loaded record.  Use existing cursor if it exists, otherwise create
        a new one.  Return a tuple containing the values obtained."""
        
        if (self.originalKeywordGroup == None) or \
           (self.originalKeyword == None):           # no record loaded?
            return ()
        
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            originalKeywordGroup = self.originalKeywordGroup.encode(TransanaGlobal.encoding)
            originalKeyword = self.originalKeyword.encode(TransanaGlobal.encoding)
        else:
            originalKeywordGroup = self.originalKeywordGroup
            originalKeyword = self.originalKeyword

        tablename = self._table()
        
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
                "  WHERE KeywordGroup = %s AND\n" + \
                "        Keyword = %s\n"
        query = DBInterface.FixQuery(query)
        c.execute(query, (originalKeywordGroup, originalKeyword))

        qr = c.fetchone()       # get query row results
        if (close_c):
            c.close()
        return qr

    def _set_db_fields(self, fields, values, c=None):
        """Set the values of fields in the database for the currently loaded
        record.  Use existing cursor if it exists, otherwise create a new
        one."""

        if (self.originalKeywordGroup == None) or \
           (self.originalKeyword == None):           # no record loaded?
            return
 
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            originalKeywordGroup = self.originalKeywordGroup.encode(TransanaGlobal.encoding)
            originalKeyword = self.originalKeyword.encode(TransanaGlobal.encoding)
        else:
            originalKeywordGroup = self.originalKeywordGroup
            originalKeyword = self.originalKeyword

        tablename = self._table()

        close_c = 0
        if (c == None):
            close_c = 1
            db = DBInterface.get_db()
            c = db.cursor()

        fv = ""
        for f, v in map(None, fields, values):
            fv = fv + f + " = " + "%s,\n\t\t"
        fv = fv[:-4]
        
        query = "UPDATE " + tablename + "\n" + \
                "  SET " + fv + "\n" + \
                "  WHERE KeywordGroup = %s AND\n" + \
                "        Keyword = %s\n"
        values = values + (originalKeywordGroup, originalKeyword)
        query = DBInterface.FixQuery(query)
        c.execute(query, values)
        
        if (close_c):
            c.close()

    def _db_start_save(self):
        """Return 0 if creating new record, 1 if updating an existing one."""
        tname = type(self).__name__
        if (self.keywordGroup == ""):
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Cannot save a %s with a blank Keyword Group."), 'utf8') % tname
            else:
                prompt = _("Cannot save a %s with a blank Keyword Group.") % tname
            raise SaveError, prompt
        elif (self.keyword == ""):
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Cannot save a %s with a blank Keyword."), 'utf8') % tname
            else:
                prompt = _("Cannot save a %s with a blank Keyword.") % tname
            raise SaveError, prompt
        else:
            # Verify record lock is still good
            db = DBInterface.get_db()

            if ((self.originalKeywordGroup == None) and \
                (self.originalKeyword == None)) or \
               ((self.record_lock == DBInterface.get_username()) and
               ((DBInterface.ServerDateTime() - self.lock_time).days <= 1)):
                c = db.cursor()
                # If record num is 0, this is a NEW record and needs to be
                # INSERTed.  Otherwise, it is an existing record to be UPDATEd.
                if (self.originalKeywordGroup == None) or \
                   (self.originalKeyword == None):           # no record loaded?
                    return 0
                else:
                    return 1
            else:
                raise SaveError, _("Record lock no longer valid.")
                
                        
    def _db_start_delete(self, use_transactions):
        """Initialize delete operation and begin transaction if necessary.
        This is a helper method intended for sub-class db_delete() methods."""
        if (self.originalKeywordGroup == None) or \
           (self.originalKeyword == None):          
            self.clear()
            raise DeleteError, _("Invalid record number (0)")
        self.lock_record()

        db = DBInterface.get_db()
        c = db.cursor()
 
        if use_transactions:
            query = "BEGIN"   # Begin a transaction
            c.execute(query)

        return (db, c)


    def _db_do_delete(self, use_transactions, c, result):
        """Do the actual record delete and handle the transaction as needed.
        This is a helper method intended for sub-class db_delete() methods."""
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            originalKeywordGroup = self.originalKeywordGroup.encode(TransanaGlobal.encoding)
            originalKeyword = self.originalKeyword.encode(TransanaGlobal.encoding)
        else:
            originalKeywordGroup = self.originalKeywordGroup
            originalKeyword = self.originalKeyword

        tablename = self._table()

        query = "DELETE FROM " + tablename + \
                "  WHERE KeywordGroup = %s AND" + \
                "        Keyword = %s"
        values = values + (originalKeywordGroup, originalKeyword)
                
        c.execute(query, values)

        if (use_transactions):
            # Commit the transaction
            if (result):
                c.execute("COMMIT")
            else:
                # Rollback transaction because some part failed
                c.execute("ROLLBACK")
            if (self.originalKeywordGroup == None) or \
               (self.originalKeyword == None):           # no record loaded?
                    self.unlock_record()
        
        return

    def _load_row(self, r):
        self.keywordGroup = r['KeywordGroup']
        self.keyword = r['Keyword']
        self.definition = r['Definition']
        if self.definition == None:
            self.definition = ''
        self.lineColorName = r['LineColorName']
        self.lineColorDef = r['LineColorDef']
        self.drawMode = r['DrawMode']
        self.lineWidth = r['LineWidth']
        self.lineStyle = r['LineStyle']
        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.keywordGroup = DBInterface.ProcessDBDataForUTF8Encoding(self.keywordGroup)
            self.keyword = DBInterface.ProcessDBDataForUTF8Encoding(self.keyword)
            self.definition = DBInterface.ProcessDBDataForUTF8Encoding(self.definition)
            self.lineColorName = DBInterface.ProcessDBDataForUTF8Encoding(self.lineColorName)

    def _get_keywordGroup(self):
        return self._keywordGroup
    
    def _set_keywordGroup(self, keywordGroup):
        
        # ALSO SEE Dialogs.add_kw_group_ui().  The same errors are caught there.
    
        # Make sure parenthesis characters are not allowed in Keyword Group.  Remove them if necessary.
        if (string.find(keywordGroup, '(') > -1) or (string.find(keywordGroup, ')') > -1):
            keywordGroup = string.replace(keywordGroup, '(', '')
            keywordGroup = string.replace(keywordGroup, ')', '')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Keyword Groups cannot contain parenthesis characters.\nYour Keyword Group has been renamed to "%s".'), 'utf8') % keywordGroup
            else:
                prompt = _('Keyword Groups cannot contain parenthesis characters.\nYour Keyword Group has been renamed to "%s".') % keywordGroup
            dlg = Dialogs.ErrorDialog(None, prompt)
            dlg.ShowModal()
            dlg.Destroy()
        # Colons are not allowed in Keyword Groups.  Remove them if necessary.
        if keywordGroup.find(":") > -1:
            keywordGroup = keywordGroup.replace(':', '')
            if 'unicode' in wx.PlatformInfo:
                msg = unicode(_('You may not use a colon (":") in the Keyword Group name.  Your Keyword Group has been changed to\n"%s"'), 'utf8')
            else:
                msg = _('You may not use a colon (":") in the Keyword Group name.  Your Keyword Group has been changed to\n"%s"')
            dlg = Dialogs.ErrorDialog(None, msg % keywordGroup)
            dlg.ShowModal()
            dlg.Destroy()
        # Let's make sure we don't exceed the maximum allowed length for a Keyword Group.
        # First, let's see what the max length is.
        maxLen = TransanaGlobal.maxKWGLength
        # Check to see if we've exceeded the max length
        if len(keywordGroup) > maxLen:
            # If so, truncate the Keyword Group
            keywordGroup = keywordGroup[:maxLen]
            # Display a message to the user describing the trunctions
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                msg = unicode(_('Keyword Group is limited to %d characters.  Your Keyword Group has been changed to\n"%s"'), 'utf8')
            else:
                msg = _('Keyword Group is limited to %d characters.  Your Keyword Group has been changed to\n"%s"')
            dlg = Dialogs.ErrorDialog(None, msg % (maxLen, keywordGroup))
            dlg.ShowModal()
            dlg.Destroy()
        # Remove white space from the Keyword Group.
        self._keywordGroup = Misc.unistrip(keywordGroup)
        
    def _del_keywordGroup(self):
        self._keywordGroup = ""
    
    def _get_keyword(self):
        return self._keyword
    def _set_keyword(self, keyword):
        # Make sure parenthesis characters are not allowed in Keywords
        if (string.find(keyword, '(') > -1) or (string.find(keyword, ')') > -1):
            keyword = string.replace(keyword, '(', '')
            keyword = string.replace(keyword, ')', '')
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Keywords cannot contain parenthesis characters.\nYour Keyword has been renamed to "%s".'), 'utf8') % keyword
            else:
                prompt = _('Keywords cannot contain parenthesis characters.\nYour Keyword has been renamed to "%s".') % keyword
            dlg = Dialogs.ErrorDialog(None, prompt)
            dlg.ShowModal()
            dlg.Destroy()
        self._keyword = Misc.unistrip(keyword)
        
    def _del_keyword(self):
        self._keyword = ""

    def _get_definition(self):
        return self._definition
    def _set_definition(self, definition):
        self._definition = definition
    def _del_definition(self):
        self._definition = ""

    def _get_lineColorName(self):
        return self._lineColorName
    def _set_lineColorName(self, lineColorName):
        self._lineColorName = lineColorName
    def _del_lineColorName(self):
        self._lineColorName = ""

    def _get_lineColorDef(self):
        return self._lineColorDef
    def _set_lineColorDef(self, lineColorDef):
        self._lineColorDef = lineColorDef
    def _del_lineColorDef(self):
        self._lineColorDef = ""

    def _get_drawMode(self):
        return self._drawMode
    def _set_drawMode(self, drawMode):
        self._drawMode = drawMode
    def _del_drawMode(self):
        self._drawMode = ""

    def _get_lineWidth(self):
        return self._lineWidth
    def _set_lineWidth(self, lineWidth):
        if lineWidth != '':
            self._lineWidth = int(lineWidth)
        else:
            self._lineWidth = 0
    def _del_lineWidth(self):
        self._lineWidth = 0

    def _get_lineStyle(self):
        return self._lineStyle
    def _set_lineStyle(self, lineStyle):
        self._lineStyle = lineStyle
    def _del_lineStyle(self):
        self._lineStyle = ""

    def _get_rl(self):
        return self._get_db_fields(('RecordLock',))[0]

    def _get_lt(self):
        # Get the Record Lock Time from the Database
        lt = self._get_db_fields(('LockTime',))
        # If a Lock Time has been specified ...
        if len(lt) > 0:
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

# Public properties
    keywordGroup = property(_get_keywordGroup, _set_keywordGroup, _del_keywordGroup,
                        """The Keyword Group.""")
    keyword = property(_get_keyword, _set_keyword, _del_keyword,
                        """The keyword.""")
    definition = property(_get_definition, _set_definition, _del_definition,
                          """ The keyword Definition""")
    lineColorName = property(_get_lineColorName, _set_lineColorName, _del_lineColorName,
                        """The Line Color Name.""")
    lineColorDef = property(_get_lineColorDef, _set_lineColorDef, _del_lineColorDef,
                        """The Line Color Def.""")
    drawMode = property(_get_drawMode, _set_drawMode, _del_drawMode,
                        """The coding Draw Mode.""")
    lineWidth = property(_get_lineWidth, _set_lineWidth, _del_lineWidth,
                        """The coding Line Width.""")
    lineStyle = property(_get_lineStyle, _set_lineStyle, _del_lineStyle,
                        """The coding Line Style.""")
    record_lock = property(_get_rl, None, None,
                        """Username of person who has locked the record
                        (Read only).""")
    lock_time = property(_get_lt, None, None,
                        """Time of the last record lock (Read only).""")
