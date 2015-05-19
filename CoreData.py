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

"""This module implements the CoreData class as part of the Data Objects."""

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

# Import Python's Internationalization Module
import gettext

# import wxPython
import wx
# Import the DataObject class the CoreData descends from
import DataObject
# Import the Transana Database Interface
import DBInterface
# Import MySQLdb for its DateTime class, used to implement the CoreData Date field
import MySQLdb
# Import the Transana Exceptions
from TransanaExceptions import *
import TransanaGlobal
import types
# import Python's sys module
import sys

class CoreData(DataObject.DataObject):
    """This class defines the structure for the Core Data object.  This
    implements Dublin Core Metadata information about a media file.  See
    http://www.dublincore.org for more information on this specification."""

    def __init__(self, name=None):
        """Initialize a CoreData object."""
        # Initialize the parent object
        DataObject.DataObject.__init__(self)
        # If a name is provided, load the appropriate data from the database
        if isinstance(name, types.StringTypes):
            self.db_load(name)
        # If no Name is provided, start with a blank object
        else:
            self.clear()

# Public methods
    def __repr__(self):
        """ Presents a String Representation of the Core Data Object """
        str = 'Core Data Record:\n'
        str = str + 'Number = %s\n' % self.number
        str = str + 'id = %s\n' % self.id
        str = str + 'comment (full-path file name) = %s\n' % self.comment
        str = str + 'title = %s\n' % self.title
        str = str + 'creator = %s\n' % self.creator
        str = str + 'subject = %s\n' % self.subject
        str = str + 'description = %s\n' % self.description
        str = str + 'publisher = %s\n' % self.publisher
        str = str + 'contributor = %s\n' % self.contributor
        str = str + 'dc_date = %s\n' % self.dc_date
        str = str + 'dc_type = %s\n' % self.dc_type
        str = str + 'format = %s\n' % self.format
        str = str + 'source = %s\n' % self.source
        str = str + 'language = %s\n' % self.language
        str = str + 'relation = %s\n' % self.relation
        str = str + 'coverage = %s\n' % self.coverage
        str = str + 'rights = %s\n\n' % self.rights
        return str

    def clear(self):
        """ Remove all data from the Core Data object's properties """
        # from DataObject
        self.number = 0
        self.id = ''
        self.comment = ''
        # from CoreData
        self.title = ''
        self.creator = ''
        self.subject = ''
        self.description = ''
        self.publisher = ''
        self.contributor = ''
        # dc_date is a DateTime object, or None
        self.dc_date = None
        self.dc_type = ''
        self.format = ''
        self.source = ''
        # Default Language should be subject to Internationalization/Localization
        self.language = _('en')
        self.relation = ''
        self.coverage = ''
        self.rights = ''

    def db_load(self, name):
        """Load a record by Name."""
        # If we're in Unicode mode, we need to encode the parameter so that the query will work right.
        if 'unicode' in wx.PlatformInfo:
            name = name.encode(TransanaGlobal.encoding)
        # Define the SQL that loads a Core Data record
        query = """ SELECT * FROM CoreData2
                      WHERE Identifier = %s """
        # Get a Database Cursor
        dbCursor = DBInterface.get_db().cursor()
        # Execute the SQL Statement using the Database Cursor
        dbCursor.execute(query, name)

        # Check the Query Results and load the data
        rowCount = dbCursor.rowcount
        # If anything other that 1 record is returned, we have a problem
        if (rowCount != 1):
            # Close the Database Cursor
            dbCursor.close()
            # Raise an Exception
            raise RecordNotFoundError, (name, rowCount)
        # If exactly 1 record is returned, load the data into the Core Data Object
        else:
            # Get the Raw Data and prepare it for loading into the Object
            data = DBInterface.fetch_named(dbCursor)
            # Load the prepared Raw Data into the Object
            self._load_row(data)

        # Close the Database Cursor            
        dbCursor.close()
        

    def db_save(self):
        """ Save the record to the database using Insert or Update as appropriate. """

        # Sanity checks
        # Core Data Records must have an Identifier (pathless filename)
        if self.id == "":
            raise SaveError, _("Blank Identifier in CoreData Record")

        # If we're in Unicode mode, ...
        if 'unicode' in wx.PlatformInfo:
            # Encode strings to UTF8 before saving them.  The easiest way to handle this is to create local
            # variables for the data.  We don't want to change the underlying object values.  Also, this way,
            # we can continue to use the Unicode objects where we need the non-encoded version. (error messages.)
            id = self.id.encode(TransanaGlobal.encoding)
            title = self.title.encode(TransanaGlobal.encoding)
            creator = self.creator.encode(TransanaGlobal.encoding)
            subject = self.subject.encode(TransanaGlobal.encoding)
            description = self.description.encode(TransanaGlobal.encoding)
            publisher = self.publisher.encode(TransanaGlobal.encoding)
            contributor = self.contributor.encode(TransanaGlobal.encoding)
            dc_type = self.dc_type.encode(TransanaGlobal.encoding)
            format = self.format.encode(TransanaGlobal.encoding)
            source = self.source.encode(TransanaGlobal.encoding)
            language = self.language.encode(TransanaGlobal.encoding)
            relation = self.relation.encode(TransanaGlobal.encoding)
            coverage = self.coverage.encode(TransanaGlobal.encoding)
            rights = self.rights.encode(TransanaGlobal.encoding)
        else:
            # If we don't need to encode the string values, we still need to copy them to our local variables.
            id = self.id
            title = self.title
            creator = self.creator
            subject = self.subject
            description = self.description
            publisher = self.publisher
            contributor = self.contributor
            dc_type = self.dc_type
            format = self.format
            source = self.source
            language = self.language
            relation = self.relation
            coverage = self.coverage
            rights = self.rights
        
        # Identify the Fields for a CoreData Database Record
        fields = ('Identifier', 'Title', 'Creator', 'Subject', 'Description', \
                  'Publisher', 'Contributor', 'DCDate', 'DCType', \
                  'Format', 'Source', 'Language', 'Relation', \
                  'Coverage', 'Rights')
        # Arrange the Data Values for the Core Data Record to match the fields variable above
        values = (id, title, creator, subject, description, \
                  publisher, contributor, self.dc_date_db, dc_type, \
                  format, source, language, relation, \
                  coverage, rights)

        # Determine if we have a new record (_db_start_save() == 0) or an existing record (inherited from DataObject)
        if (self._db_start_save() == 0):
            # Duplicate Identifiers are not allowed.
            if DBInterface.record_match_count("CoreData2", ("Identifier",), (id,)) > 0:
                # If a duplicate is found, interrupt the Save by raising an exception
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A Core Data record named "%s" already exists.'), 'utf8')
                else:
                    prompt = _('A Core Data record named "%s" already exists.')
                raise SaveError, prompt % self.id

            # Insert the new record into the table.
            # First, build the appropriate SQL Statment, starting with Insert
            query = "INSERT INTO CoreData2\n("
            # Add all defined fields to the SQL Insert Statement
            for field in fields:
                query = "%s%s," % (query, field)
            # Remove the final (unwanted) comma, as we are done adding fields
            query = query[:-1] + ')'
            # Add the SQL Values keyword
            query = "%s\nVALUES\n(" % query
            # Add the defined values to match the defined fields
            for value in values:
                query = "%s%%s," % query
            # Remove the final (unwanted) comma, as we are done adding values
            query = query[:-1] + ')'
        else:
            # check for duplicate records.  (This should not be possible!)
            if DBInterface.record_match_count("CoreData2", ("Identifier", "!CoreDataNum"), (id, self.number) ) > 0:
                # If a duplicate is found, interrupt the Save by raising an exception
                raise SaveError, _("A Core Data Record with that ID already exists.")
            
            # Define the SQL Statement for updating an existing record
            query = """UPDATE CoreData2
                SET Identifier = %s,
                    Title = %s,
                    Creator = %s,
                    Subject = %s,
                    Description = %s, 
                    Publisher = %s,
                    Contributor = %s,
                    DCDate = %s,
                    DCType = %s, 
                    Format = %s,
                    Source = %s,
                    Language = %s,
                    Relation = %s, 
                    Coverage = %s,
                    Rights = %s
                WHERE CoreDataNum = %s
            """
            # Add the Record Number to the Values for the WHERE Clause
            values = values + (self.number,)

        # Get a Database Cursor
        dbCursor = DBInterface.get_db().cursor()
        # Execute the SQL Query with the data values assembled above
        dbCursor.execute(query, values)

        # If the number field is 0, we've just added a new record to the database.
        # In this case, we need to now load the database record to get the new Record Number.
        if (self.number == 0):
            # Load the auto-assigned new number record
            self.db_load(self.id)           

        # Close the Database Cursor
        dbCursor.close()
            
    
    def db_delete(self, use_transactions=1):
        """Delete this object record from the database."""
        # Initialize delete operation, begin transaction if necessary
        (database, dbCursor) = self._db_start_delete(use_transactions)
        # Assume success
        result = 1
        
        # Delete the actual record.
        self._db_do_delete(use_transactions, dbCursor, result)

        # Close the Database Cursor
        dbCursor.close()
        # Clear all values from the Core Data Object
        self.clear()

        # Return the Result to the calling procedure
        return result
    
# Private methods

    def _load_row(self, row):
        """ Load the row of data returned from the database into the Core Data Object's properties """
        self.number = row['CoreDataNum']
        self.id = row['Identifier']
        self.title = row['Title']
        self.creator = row['Creator']
        self.subject = row['Subject']
        self.description = row['Description']
        self.publisher = row['Publisher']
        self.contributor = row['Contributor']
        self.dc_date = row['DCDate']
        self.dc_type = row['DCType']
        self.format = row['Format']
        self.source = row['Source']
        self.language = row['Language']
        self.relation = row['Relation']
        self.coverage = row['Coverage']
        self.rights = row['Rights']
        # If we're in Unicode mode, we need to encode the data from the database appropriately.
        # (unicode(var, TransanaGlobal.encoding) doesn't work, as the strings are already unicode, yet aren't decoded.)
        if 'unicode' in wx.PlatformInfo:
            self.id = DBInterface.ProcessDBDataForUTF8Encoding(self.id)
            self.title = DBInterface.ProcessDBDataForUTF8Encoding(self.title)
            self.creator = DBInterface.ProcessDBDataForUTF8Encoding(self.creator)
            self.subject = DBInterface.ProcessDBDataForUTF8Encoding(self.subject)
            self.description = DBInterface.ProcessDBDataForUTF8Encoding(self.description)
            self.publisher = DBInterface.ProcessDBDataForUTF8Encoding(self.publisher)
            self.contributor = DBInterface.ProcessDBDataForUTF8Encoding(self.contributor)
            self.dc_type = DBInterface.ProcessDBDataForUTF8Encoding(self.dc_type)
            self.format = DBInterface.ProcessDBDataForUTF8Encoding(self.format)
            self.source = DBInterface.ProcessDBDataForUTF8Encoding(self.source)
            self.language = DBInterface.ProcessDBDataForUTF8Encoding(self.language)
            self.relation = DBInterface.ProcessDBDataForUTF8Encoding(self.relation)
            self.coverage = DBInterface.ProcessDBDataForUTF8Encoding(self.coverage)
            self.rights = DBInterface.ProcessDBDataForUTF8Encoding(self.rights)

    # Property Getters and Setters
    # (NOTE:  Number, Id, and Comment are provided by DataObject)

    # Title
    def _get_title(self):
        return self._title
    def _set_title(self, title):
        self._title = title
    def _del_title(self):
        self._title = ""

    # Creator
    def _get_creator(self):
        return self._creator
    def _set_creator(self, creator):
        self._creator = creator
    def _del_creator(self):
        self._creator = ""

    # Subject
    def _get_subject(self):
        return self._subject
    def _set_subject(self, subject):
        self._subject = subject
    def _del_subject(self):
        self._subject = ""

    # Description
    def _get_description(self):
        return self._description
    def _set_description(self, description):
        self._description = description
    def _del_description(self):
        self._description = ""

    # Publisher
    def _get_publisher(self):
        return self._publisher
    def _set_publisher(self, publisher):
        self._publisher = publisher
    def _del_publisher(self):
        self._publisher = ""

    # Contributor
    def _get_contributor(self):
        return self._contributor
    def _set_contributor(self, contributor):
        self._contributor = contributor
    def _del_contributor(self):
        self._contributor = ""

    # Date
    # TODO:  Localization of Date Format
    def _get_dc_date(self):
        # Although _dc_date is a MySQLdb.DateTime object, Date will be presented as a 'MM/DD/YYYY' string!
        if self._dc_date == None:
            return ''
        else:
            return "%02d/%02d/%s" % (self._dc_date.month, self._dc_date.day, self._dc_date.year)
    def _set_dc_date(self, dc_date):
        # Although _dc_date is a MySQLdb.DateTime object, Date will be presented for ingestion as a 'MM/DD/YYYY' string!
        # A None Object or a blank string should set _dc_date to None
        if (dc_date == '') or (dc_date == None) or (dc_date == u'  /  /    '):
            self._dc_date = None
        else:
            try:
                # A 'MM/DD/YYYY' string needs to be parsed.
                if isinstance(dc_date, types.StringTypes):
                    (month, day, year) = dc_date.split('/')
                    # The wxMaskedTextCtrl returns '  /  /    ' if left empty.
                    # We need to convert the Month into an Integer, setting it to 1 if empty
                    if month.strip() == '':
                        month = 1
                    else:
                        month = int(month)
                    # We need to convert the Day into an Integer, setting it to 1 if empty
                    if day.strip() == '':
                        day = 1
                    else:
                        day = int(day)
                    # We need to convert the Year into an Integer, setting it to 0 if empty
                    if year.strip() == '':
                        year = 0
                    else:
                        year = int(year)
                    # if all values were empty or bogus, set _dc_date to None, but if we have SOMETHING to save,
                    # save it.
                    if (year != 0) or (month != 0) or (day != 0):
                        self._dc_date = MySQLdb.Date(int(year), int(month), int(day))
                    else:
                        self._dc_date = None
                # If we get a non-string argument, just save it.  
                else:
                    self._dc_date = dc_date
            except:
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Date Format Error.  The Date must be in the format MM/DD/YYYY.\n%s'), 'utf8')
                else:
                    prompt = _('Date Format Error.  The Date must be in the format MM/DD/YYYY.\n%s')
                raise SaveError, prompt % sys.exc_info()[1]
    def _del_dc_date(self):
        self._dc_date = None

    def _get_dc_date_db(self):
        """ Read-only version of the Date in a MySQL Friendly format """
        # Although _dc_date is a MySQLdb.DateTime object, the database needs the Date to be
        # presented as a 'YYYY-MM-DD' string!
        if self._dc_date == None:
            return None
        else:
            return "%04d/%02d/%02d" % (self._dc_date.year, self._dc_date.month, self._dc_date.day)

    # Type
    def _get_dc_type(self):
        return self._dc_type
    def _set_dc_type(self, dc_type):
        self._dc_type = dc_type
    def _del_dc_type(self):
        self._dc_type = ""

    # Format
    def _get_format(self):
        return self._format
    def _set_format(self, format):
        self._format = format
    def _del_format(self):
        self._format = ""

    # Source
    def _get_source(self):
        return self._source
    def _set_source(self, source):
        self._source = source
    def _del_source(self):
        self._source = ""

    # Language
    def _get_language(self):
        return self._language
    def _set_language(self, language):
        self._language = language
    def _del_language(self):
        self._language = ""

    # Relation
    def _get_relation(self):
        return self._relation
    def _set_relation(self, relation):
        self._relation = relation
    def _del_relation(self):
        self._relation = ""

    # Coverage
    def _get_coverage(self):
        return self._coverage
    def _set_coverage(self, coverage):
        self._coverage = coverage
    def _del_coverage(self):
        self._coverage = ""

    # Rights
    def _get_rights(self):
        return self._rights
    def _set_rights(self, rights):
        self._rights = rights
    def _del_rights(self):
        self._rights = ""


    # Define Public Properties

    descriptionText = _("See http://www.dublincore.org for details.")
    title = property(_get_title, _set_title, _del_title, descriptionText)
    creator = property(_get_creator, _set_creator, _del_creator, descriptionText)
    subject = property(_get_subject, _set_subject, _del_subject, descriptionText)
    description = property(_get_description, _set_description, _del_description, descriptionText)
    publisher = property(_get_publisher, _set_publisher, _del_publisher, descriptionText)
    contributor = property(_get_contributor, _set_contributor, _del_contributor, descriptionText)
    dc_date = property(_get_dc_date, _set_dc_date, _del_dc_date, descriptionText)
    dc_date_db = property(_get_dc_date_db, None, None, 'dc_date in the format needed for the MySQL Database')
    dc_type = property(_get_dc_type, _set_dc_type, _del_dc_type, descriptionText)
    format = property(_get_format, _set_format, _del_format, descriptionText)
    source = property(_get_source, _set_source, _del_source, descriptionText)
    language = property(_get_language, _set_language, _del_language, descriptionText)
    relation = property(_get_relation, _set_relation, _del_relation, descriptionText)
    coverage = property(_get_coverage, _set_coverage, _del_coverage, descriptionText)
    rights = property(_get_rights, _set_rights, _del_rights, descriptionText)
