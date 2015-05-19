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

"""This module implements the Keyword class as part of the Data Objects."""

# NOTE:  While the Keyword object is very SIMILAR to all the other Data Objects
#        in Transana, it will NOT be inherited from the DataObject class.  The
#        reason for this is that the DataObject class requires the database
#        record to have a RecordNumber field, and the Keywords Table does not
#        have (or at this time need) this field.  The following code is
#        obviously heavily based on DataObject.   DKW

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'
# Based on code/ideas/logic from UKeywordObject Delphi unit by DKW

import wx
from TransanaExceptions import *
import DBInterface
import Dialogs
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
        return str

    def clear(self):
        """Clear all properties, resetting them to default values."""
        attrdesc = inspect.classify_class_attrs(self.__class__)
        for attr in attrdesc:
            if attr[1] == "property":
                try:
                    delattr(self, attr[0])
                except AttributeError, e:
                    pass # probably a non-deletable attribute

    def checkEpisodesAndClipsForLocks(self):
        """ Checks Episodes and Clips to see if a Keyword record is free of related locks """
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            originalKeywordGroup = self.originalKeywordGroup.encode(TransanaGlobal.encoding)
            originalKeyword = self.originalKeyword.encode(TransanaGlobal.encoding)
        else:
            originalKeywordGroup = self.originalKeywordGroup
            originalKeyword = self.originalKeyword

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
        c.execute(query, values)
        result = c.fetchall()
        RecordCount = len(result)
        c.close()

        # If no Episodes that contain the record are locked, check the lock status of Clips that contain the Keyword
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
            c.execute(query, values)
            result = c.fetchall()
            RecordCount = len(result)
            c.close()

        if RecordCount != 0:
            LockName = result[0][0]
            return (RecordCount, LockName)
        else:
            return (RecordCount, None)
        

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
            (EpisodeClipLockCount, LockName) = self.checkEpisodesAndClipsForLocks()
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
        c = db.cursor()
        c.execute(query, (keywordGroup, keyword))
        n = c.rowcount
        if (n != 1):
            c.close()
            self.clear()
            raise RecordNotFoundError, (keywordGroup + ':' + keyword, n)
        else:
            r = DBInterface.fetch_named(c)
            self._load_row(r)
            

        c.close()

    # TODO:  We need to change the model for all Data Object Saves.  This needs to be
    #        a boolean function so that if the Save fails, the Properties Dialog can
    #        remain open, giving the user a chance to fix whatever caused the Save to
    #        fail.  Remember to do this here AND in DataObject.
    def db_save(self):
        """Save the record to the database using Insert or Update as
        appropriate."""

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

        values = (keywordGroup, keyword, definition)

        if (self._db_start_save() == 0):
            # duplicate Keywords are not allowed
            if DBInterface.record_match_count("Keywords2", \
                            ("KeywordGroup", "Keyword"), \
                            (keywordGroup, keyword) ) > 0:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Keyword named "%s : %s" already exists.'), 'utf8') % (self.keywordGroup, self.keyword)
                else:
                    prompt = _('A Keyword named "%s : %s" already exists.') % (self.keywordGroup, self.keyword)
                raise SaveError, prompt

            # insert the new Keyword
            query = """
            INSERT INTO Keywords2
                (KeywordGroup, Keyword, Definition)
                VALUES
                (%s,%s,%s)
            """
            c = DBInterface.get_db().cursor()
            c.execute(query, values)
            c.close()
        else:
            # check for dupes, which are not allowed if either the Keyword Group or Keyword have been changed.
            if (DBInterface.record_match_count("Keywords2", \
                            ("KeywordGroup", "Keyword"), \
                            (keywordGroup, keyword) ) > 0) and \
               ((originalKeywordGroup != keywordGroup) or \
                (originalKeyword != keyword)):
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Keyword named "%s : %s" already exists.'), 'utf8') % (self.keywordGroup, self.keyword)
                else:
                    prompt = _('A Keyword named "%s : %s" already exists.') % (self.keywordGroup, self.keyword)
                raise SaveError, prompt

            # NOTE:  This is a special instance.  Keywords behave differently than other DataObjects here!
            # Before we can save, we have to check to see if someone locked an Episode or a Clip that
            # uses our Keyword since we obtained the lock on the Keyword.  It doesn't make sense to me
            # to block editing Episode or Clips properties because some else is editing a Keyword it contains.
            # Mostly, editing Keywords will have to do with changing their definition field, not their keyword
            # group or keyword fields.  Because of that, we have to block the save of an altered keyword
            # because the locked Episode or Clip would retain the obsolete keyword and would lose the
            # new keyword record, so it would not appear in all Search results that it should.  I expect
            # this to be extraordinarily rare.
            (EpisodeClipLockCount, LockName) = self.checkEpisodesAndClipsForLocks()
            if EpisodeClipLockCount != 0 and \
               ((originalKeywordGroup != keywordGroup) or \
                (originalKeyword != keyword)):
                tempstr = _("""You cannot proceed because another user has recently started editing an Episode or a Clip that uses\n
Keyword "%s:%s".  If you change the Keyword now, that would corrupt the record that is \n
currently locked by %s.  Please try again later.""")
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    tempstr = unicode(tempstr, 'utf8')
                dlg = Dialogs.ErrorDialog(TransanaGlobal.menuWindow, tempstr  % (self.originalKeywordGroup, self.originalKeyword, LockName))
                dlg.ShowModal()
                dlg.Destroy()
            
            else:
                # update the record
                query = """
                UPDATE Keywords2
                    SET KeywordGroup = %s,
                        Keyword = %s,
                        Definition = %s
                    WHERE KeywordGroup = %s AND
                          Keyword = %s
                """
                values = values + (originalKeywordGroup, originalKeyword)
                c = DBInterface.get_db().cursor()
                c.execute(query, values)
                c.close()

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
                    c = DBInterface.get_db().cursor()
                    c.execute(query, values)
                    c.close()
                # If the save is successful, we need to update the "original" values to reflect the new record key.
                # Otherwise, we can't unlock the proper record, among other things.
                self.originalKeywordGroup = self.keywordGroup
                self.originalKeyword = self.keyword

    def db_delete(self, use_transactions=1):
        """Delete this object record from the database.  Raises
        RecordLockedError exception if record is locked and unable to be
        deleted."""
        tempstr = "Delete Keyword object has not been implemented."
        dlg = wx.MessageDialog(TransanaGlobal.menuWindow, tempstr, "Keyword Object", wx.OK | wx.ICON_EXCLAMATION)
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
                
        c.execute(query, values + (originalKeywordGroup, originalKeyword))
        
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
            query = "Begin\n"   # Begin a transaction
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

        query = "DELETE FROM " + tablename + "\n" + \
                "  WHERE KeywordGroup = %s AND\n" + \
                "        Keyword = %s\n"
                
        c.execute(query, values + (originalKeywordGroup, originalKeyword))

        if (use_transactions):
            # Commit the transaction
            if (result):
                c.execute("COMMIT\n")
            else:
                # Rollback transaction because some part failed
                c.execute("ROLLBACK\n")
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
        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.keywordGroup = DBInterface.ProcessDBDataForUTF8Encoding(self.keywordGroup)
            self.keyword = DBInterface.ProcessDBDataForUTF8Encoding(self.keyword)
            self.definition = DBInterface.ProcessDBDataForUTF8Encoding(self.definition)

    def _get_keywordGroup(self):
        return self._keywordGroup
    def _set_keywordGroup(self, keywordGroup):
        # Make sure parenthesis characters are not allowed in Keyword Group
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
        self._keywordGroup = keywordGroup.strip()
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
        self._keyword = keyword.strip()
    def _del_keyword(self):
        self._keyword = ""

    def _get_definition(self):
        return self._definition
    def _set_definition(self, definition):
        self._definition = definition
    def _del_definition(self):
        self._definition = ""

    def _get_rl(self):
        return self._get_db_fields(('RecordLock',))[0]

    def _get_lt(self):
        return self._get_db_fields(('LockTime',))[0]

# Public properties
    keywordGroup = property(_get_keywordGroup, _set_keywordGroup, _del_keywordGroup,
                        """The Keyword Group.""")
    keyword = property(_get_keyword, _set_keyword, _del_keyword,
                        """The keyword.""")
    definition = property(_get_definition, _set_definition, _del_definition,
                          """ The keyword Definition""")
    record_lock = property(_get_rl, None, None,
                        """Username of person who has locked the record
                        (Read only).""")
    lock_time = property(_get_lt, None, None,
                        """Time of the last record lock (Read only).""")
