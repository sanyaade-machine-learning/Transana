# Copyright (C) 2003 - 2008 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Data Import for Transana based on the Transana XML schema."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'
# Patch sent by David Fraser to eliminate need for mx module

DEBUG = False
if DEBUG:
    print "XMLImport DEBUG is ON!"

# import wxPython
import wx
# Import the mx DateTime module
#import mx.DateTime
import datetime
import time

# import necessary Transana modules
import Clip
import ClipKeywordObject
import Collection
import CoreData
import DBInterface
import Dialogs
import Episode
import Keyword
import Misc
import Note
import Series
import TransanaGlobal
import Transcript

# import Python's os and sys modules
import os
import sys
# import Python's Regular Expression parser
import re
# import Python's cPickle module
import cPickle

MENU_FILE_EXIT = wx.NewId()

class XMLImport(Dialogs.GenForm):
   """ This window displays a variety of GUI Widgets. """
   def __init__(self,parent,id,title):
       Dialogs.GenForm.__init__(self, parent, id, title, (550,150), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Import Database')
       # Define the minimum size for this dialog as the initial size
       self.SetSizeHints(550, 150)

       # Import Message Layout
       lay = wx.LayoutConstraints()
       lay.top.SameAs(self.panel, wx.Top, 10)
       lay.left.SameAs(self.panel, wx.Left, 10)
       lay.right.SameAs(self.panel, wx.Right, 10)
       lay.height.AsIs()
       # If the XML filename path is not empty, we need to tell the user.
       if DBInterface.IsDatabaseEmpty():
           prompt = _('Please select a Transana XML File to import.')
       else:
           prompt = _('Your current database is not empty.  Please note that any duplicate object names will cause your import to fail.\nPlease select a Transana XML File to import.')
       importText = wx.StaticText(self.panel, -1, prompt)
       importText.SetConstraints(lay)

       # XML Filename Layout
       lay = wx.LayoutConstraints()
       lay.top.Below(importText, 10)
       lay.left.SameAs(self.panel, wx.Left, 10)
       lay.width.PercentOf(self.panel, wx.Width, 80)  # 80% width
       lay.height.AsIs()
       self.XMLFile = self.new_edit_box(_("Transana-XML Filename"), lay, '')
       self.XMLFile.SetDropTarget(EditBoxFileDropTarget(self.XMLFile))

       # Browse button layout
       lay = wx.LayoutConstraints()
       lay.top.SameAs(self.XMLFile, wx.Top)
       lay.left.RightOf(self.XMLFile, 10)
       lay.right.SameAs(self.panel, wx.Right, 10)
       lay.bottom.SameAs(self.XMLFile, wx.Bottom)
       browse = wx.Button(self.panel, wx.ID_FILE1, _("Browse"), wx.DefaultPosition)
       browse.SetConstraints(lay)
       wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)

       self.Layout()
       self.SetAutoLayout(True)
       self.CenterOnScreen()

       # We need to know the encoding of the import file, which differs depending on the
       # TransanaXML Version.  Let's assume UTF-8 unless we have to change it.
       self.importEncoding = 'utf8'

       self.XMLFile.SetFocus()


   def Import(self):
       # use the LONGEST title here to set the width of the dialog box!
       progress = wx.ProgressDialog(_('Transana XML Import'), _('Importing Transcript records (This may be slow because of the size of Transcript records.)'), style = wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)
       if progress.GetSize()[0] > 800:
            progress.SetSize((800, progress.GetSize()[1]))
            progress.Centre()
       # Initialize the Transana-XML Version.  This allows differential processing based on the version used to
       # create the Transana-XML data file
       self.XMLVersionNumber = 0.0
       # Initialize dictionary variables used to keep track of data values that may change during import.
       recNumbers = {}
       recNumbers['Series'] = {0:0}
       recNumbers['Episode'] = {0:0}
       recNumbers['Transcript'] = {0:0}
       recNumbers['Collection'] = {0:0}
       recNumbers['Clip'] = {0:0}
       recNumbers['OldClip'] = {0:0}
       recNumbers['Note'] = {0:0}
       clipTranscripts = {}
       clipStartStop = {}

       # Get the database connection
       db = DBInterface.get_db()
       if db != None:
           # Begin Database Transaction 
           dbCursor = db.cursor()
           SQLText = 'BEGIN'
           dbCursor.execute(SQLText)
           dbCursor.close()

       # We need to track the number of lines read and processed from the input file.
       lineCount = 0 

       # Start exception handling
       try:
           # Assume we're good to continue unless informed otherwise
           contin = True
           # Open the XML file
           f = file(self.XMLFile.GetValue(), 'r')

           # Initialize objectType and dataType, which are used to parse the file
           objectType = None 
           dataType = None
           # For each line in the file ...
           for line in f:
               # ... increment the line counter
               lineCount += 1
               # We can't just use strip() here.  Strings ending with the a with a grave (Alt-133) get corrupted!
               line = Misc.unistrip(line)

               if DEBUG:
                   print "Line %d: '%s' %s %s" % (lineCount, line, objectType, dataType)

               # Begin line processing.
               # Generally, we figure out what kind of object we're importing (objectType) and
               # then figure out what object property we're importing (dataType) and then we
               # import the data.  But of course, there's a bit more to it than that.

               # Code for updating the Progress Bar
               if line.upper() == '<SERIESFILE>':
                   progress.Update(0, _('Importing Series records'))
               elif line.upper() == '<EPISODEFILE>':
                   progress.Update(9, _('Importing Episode records'))
               elif line.upper() == '<COREDATAFILE>':
                   progress.Update(18, _('Importing Core Data records'))
               elif line.upper() == '<COLLECTIONFILE>':
                   progress.Update(27, _('Importing Collection records'))
               elif line.upper() == '<CLIPFILE>':
                   progress.Update(36, _('Importing Clip records'))
               elif line.upper() == '<TRANSCRIPTFILE>':
                   progress.Update(45, _('Importing Transcript records (This may be slow because of the size of Transcript records.)'))
               elif line.upper() == '<KEYWORDFILE>':
                   progress.Update(55, _('Importing Keyword records'))
               elif line.upper() == '<CLIPKEYWORDFILE>':
                   progress.Update(64, _('Importing Clip Keyword records'))
               elif line.upper() == '<NOTEFILE>':
                   progress.Update(73, _('Importing Note records'))
               elif line.upper() == '<FILTERFILE>':
                   progress.Update(82, _('Importing Filter records'))

               # Transana XML Version Checking
               elif line.upper() == '<TRANSANAXMLVERSION>':
                   objectType = None
                   dataType = 'XMLVersionNumber'
               elif line.upper() == '</TRANSANAXMLVERSION>':

                   # Version 1.0 -- Original Transana XML for Transana 2.0 release
                   # Version 1.1 -- Unicode encoding added to Transana XML for Transana 2.1 release
                   # Version 1.2 -- Filter Table added to Transana XML for Transana 2.11 release
                   # Version 1.3 -- FilterData handling changed to accomodate Unicode data for Transana 2.21 release
                   # Version 1.4 -- Database structure changed to accomodate Multiple Transcript Clips for Transana 2.30 release.

                   # Transana-XML version 1.0 ...
                   if self.XMLVersionNumber == '1.0':
                       # ... used Latin1 encoding
                       self.importEncoding = 'latin-1'
                   # Transana-XML versions 1.1 through 1.4 ...
                   elif self.XMLVersionNumber in ['1.1', '1.2', '1.3', '1.4']:
                       # ... use UTF8 encoding
                       self.importEncoding = 'utf8'
                   # All other Transana XML versions ...
                   else:
                       # ... haven't been defined yet.  We'd better prevent importing data from a LATER Transana-XML
                       # version, as we probably don't know how to process everything in it.
                       msg = _('The Database you are trying to import was created with a later version\nof Transana.  Please upgrade your copy of Transana and try again.')
                       dlg = Dialogs.ErrorDialog(None, msg)
                       dlg.ShowModal()
                       dlg.Destroy()
                       # This error means we can't continue processing the file.
                       contin = False
                       break

               # Code for Creating and Saving Objects
               elif line.upper() == '<SERIES>':
                   currentObj = Series.Series()
                   objectType = 'Series'
                   dataType = None

               elif line.upper() == '<EPISODE>':
                   currentObj = Episode.Episode()
                   objectType = 'Episode'
                   dataType = None

               elif line.upper() == '<COREDATA>':
                   currentObj = CoreData.CoreData()
                   objectType = 'CoreData'
                   dataType = None

               elif line.upper() == '<TRANSCRIPT>':
                   currentObj = Transcript.Transcript()
                   objectType = 'Transcript'
                   dataType = None

               elif line.upper() == '<COLLECTION>':
                   currentObj = Collection.Collection()
                   objectType = 'Collection'
                   dataType = None

               elif line.upper() == '<CLIP>':
                   currentObj = Clip.Clip()
                   objectType = 'Clip'
                   dataType = None

               elif line.upper() == '<KEYWORDREC>':
                   currentObj = Keyword.Keyword()
                   objectType = 'Keyword'
                   dataType = None

               elif line.upper() == '<CLIPKEYWORD>':
                   currentObj = ClipKeywordObject.ClipKeyword('', '')
                   objectType = 'ClipKeyword'
                   dataType = None

               elif line.upper() == '<NOTE>':
                   currentObj = Note.Note()
                   objectType = 'Note'
                   dataType = None

               elif line.upper() == '<FILTER>':
                    # There is not an Object for the Filter table.  We have to create the data record by hand.
                    currentObj = None
                    self.FilterReportType = None
                    self.FilterScope = None
                    self.FilterConfigName = None
                    self.FilterFilterDataType = None
                    self.FilterFilterData = ''
                    objectType = 'Filter'
                    dataType = None

               # If we're closing a data record in the XML, we need to SAVE the data object.
               elif line.upper() == '</SERIES>' or \
                    line.upper() == '</EPISODE>' or \
                    line.upper() == '</COREDATA>' or \
                    line.upper() == '</TRANSCRIPT>' or \
                    line.upper() == '</COLLECTION>' or \
                    line.upper() == '</CLIP>' or \
                    line.upper() == '</KEYWORDREC>' or \
                    line.upper() == '</CLIPKEYWORD>' or \
                    line.upper() == '</NOTE>' or \
                    line.upper() == '</FILTER>':
                   dataType = None

                   # Saves are one area where problems will arise if the data's not clean.
                   # We can trap some of these problems here.
                   try:
                       # We can't just keep the numbers that were assigned in the exporting database.
                       # Objects use the presence of a number to update rather than insert, and we need to
                       # insert here.  Therefore, we'll strip the record number here, but remember it for
                       # user later.
                       if objectType == 'Series' or \
                          objectType == 'Episode' or \
                          objectType == 'Transcript' or \
                          objectType == 'Collection' or \
                          objectType == 'Clip' or \
                          objectType == 'Note':
                           oldNumber = currentObj.number
                           currentObj.number = 0

                           # Clip Trancript, Start, and Stop times need to be transferred to Clip Transcripts
                           # if we're from a pre-version 1.4 transcript!  In earlier versions, Clip Transcript Number
                           # was a Clip property.  Starting in 1.4, a clip can have multiple transcript objects
                           # attached.  Multiple-Transcript Clips also need to copy Clip Start and Stop times
                           # as part of the Clip Transcript.

                           # NOTE:  Because we have to look up old and new Clip Numbers to manipulate Transcript data,
                           #        the Clips MUST be processed BEFORE the Transcripts.  That's why they come earlier in
                           #        the data files.
                           
                           # First, check for a Transcript and the XML Version
                           if (objectType == 'Transcript') and (self.XMLVersionNumber in ['1.1', '1.2', '1.3']):
                               # see if we have a Clip transcript
                               if currentObj.clip_num > 0:
                                   # The source transcript is saved at this point as the OLD transcript number, to be updated
                                   # later in the last step.  To determine that, look up the Clip's (current number) OLD Clip
                                   # number in the recNumbers['OldClip'] dictionary, and use that to look up the source
                                   # Transcript number in the clipTranscripts dictionary.
                                   if clipTranscripts.has_key(recNumbers['OldClip'][currentObj.clip_num]):
                                       currentObj.source_transcript = clipTranscripts[recNumbers['OldClip'][currentObj.clip_num]]
                                   # If there's no record in clipTranscripts, we've got an orphaned clips!
                                   else:
                                       currentObj.source_transcript = 0
                                   # Pre-1.4 clip transcripts were always singles, so didn't have Sort Order.  Default to 0.
                                   currentObj.sort_order = 0
                                   # Look up the Clip's start time in the clipStartStop dictionary, which is keyed to the
                                   # clip's OLD number, which we have to look up using recNumbers['OldClip']
                                   if clipStartStop.has_key((recNumbers['OldClip'][currentObj.clip_num], 'Start')):
                                       currentObj.clip_start = clipStartStop[(recNumbers['OldClip'][currentObj.clip_num], 'Start')]
                                   else:
                                       currentObj.clip_start = 0
                                   # Look up the Clip's stop time in the clipStartStop dictionary, which is keyed to the
                                   # clip's OLD number, which we have to look up using recNumbers['OldClip']
                                   if clipStartStop.has_key((recNumbers['OldClip'][currentObj.clip_num], 'Stop')):
                                       currentObj.clip_stop = clipStartStop[(recNumbers['OldClip'][currentObj.clip_num], 'Stop')]
                                   else:
                                       currentObj.clip_stop = 0
                               
                           # To prevent the ClipObject from trying to save the Clip Transcript,
                           # (which it doesn't yet have in the import), we must set its transcript(s) number(s) to -1.
                           if objectType == 'Clip':
                               for tr in currentObj.transcripts:
                                   tr.number = -1

                           if DEBUG and (objectType == 'Transcript') and False:
                               tmpdlg = wx.MessageDialog(self, currentObj.__repr__())
                               tmpdlg.ShowModal()
                               tmpdlg.Destroy()

                           # Save the data object
                           currentObj.db_save()
                           # Let's keep a record of the old and new object numbers for each object saved.
                           recNumbers[objectType][oldNumber] = currentObj.number

                           # If we've just saved a Clip ...
                           if objectType == 'Clip':
                               # ... we need to save the reverse lookup data, so we can find the clip's OLD
                               # number based on it's new number.  This must be post-save, as that's when
                               # the new object number gets assigned.
                               recNumbers['OldClip'][currentObj.number] = oldNumber

                       elif  objectType == 'CoreData':
                           currentObj.number = 0
                           currentObj.db_save()

                       elif objectType == 'Keyword':
                           currentObj.db_save()
                           
                       elif objectType == 'ClipKeyword':
                           if (currentObj.episodeNum != 0) or (currentObj.clipNum != 0):
                               currentObj.db_save()
                               
                       elif objectType == 'Filter':
                           # Starting with XML Version 1.3, we have to deal with encoding issues for the Filter data
                           if not self.XMLVersionNumber in ['1.0', '1.1', '1.2']:
                               if self.FilterFilterDataType in ['1', '2', '3', '5', '6', '7']:
                                   # Unpack the pickled data, which must be done differently depending on its current form
                                   if type(self.FilterFilterData).__name__ == 'array':
                                       data = cPickle.loads(self.FilterFilterData.tostring())
                                   elif type(self.FilterFilterData).__name__ == 'unicode':
                                       data = cPickle.loads(self.FilterFilterData.encode('utf-8'))
                                   else:
                                       data = cPickle.loads(self.FilterFilterData)
                                   # Re-pickle the data.  It's in a common form now that's somehow friendlier.
                                   data = cPickle.dumps(data)

                               elif self.FilterFilterDataType in ['4', '8']:
                                   pass

                               # All other data is encoded and needs to be decoded, but was not pickled.
                               else:
                                   self.FilterFilterData = DBInterface.ProcessDBDataForUTF8Encoding(self.FilterFilterData)

                           # Certain FilterDataTypes need to have their DATA adjusted for the new object numbers!
                           # This should be done before the save.
                           # So if we have Clips or Notes Filter Data ...
                           if self.FilterFilterDataType in ['2', '8']:
                               # ... initialize a List for accepting the altered Filter Data
                               filterData = []
                               # Unpack the pickled data, which must be done differently depending on its current form
                               if type(self.FilterFilterData).__name__ == 'array':
                                   data = cPickle.loads(self.FilterFilterData.tostring())
                               elif type(self.FilterFilterData).__name__ == 'unicode':
                                   data = cPickle.loads(self.FilterFilterData.encode('utf-8'))
                               else:
                                   data = cPickle.loads(self.FilterFilterData)
                               # Iterate through the data records
                               for dataRec in data:
                                   # If we have a Clip record ...
                                   if self.FilterFilterDataType == '2':
                                       # ... and if the Collection Number still exists in the new data set ...
                                       if recNumbers['Collection'].has_key(dataRec[1]):
                                           # ... get the new Collection Number ...
                                           collNum = recNumbers['Collection'][dataRec[1]]
                                           # ... and substitute it into the data record
                                           filterData.append((dataRec[0], collNum, dataRec[2]))
                                   # If we have a Notes record ...
                                   elif self.FilterFilterDataType == '8':
                                       # ... and if the Note Number still exists in the new data set ...
                                       if recNumbers['Note'].has_key(dataRec[0]):
                                           # ... get the new Note number ...
                                           noteNum = recNumbers['Note'][dataRec[0]]
                                           # ... and substitute it into the data record
                                           filterData.append((noteNum, ) + dataRec[1:])
                                    # NOTE that data records without new references are automatically dropped from
                                    #      the filter data!
                               # Now re-pickle the filter data
                               self.FilterFilterData = cPickle.dumps(filterData)
                           # Create the query to save the Filter record    
                           query = """ INSERT INTO Filters2
                                           (ReportType, ReportScope, ConfigName, FilterDataType, FilterData)
                                         VALUES
                                           (%s, %s, %s, %s, %s) """
                           # Build the values to match the query, including the pickled Clip data
                           values = (self.FilterReportType, self.FilterScope, self.FilterConfigName, self.FilterFilterDataType, self.FilterFilterData)
                           # Save the Filter data
                           if db != None:
                               dbCursor = db.cursor()
                               dbCursor.execute(query, values)
                               dbCursor.close()

                   except:

                       if DEBUG:
                           print
                           print sys.exc_info()[0], sys.exc_info()[1]
                           print
                           if (objectType == 'Transcript'):
                               tmpdlg = wx.MessageDialog(self, currentObj.__repr__())
                               tmpdlg.ShowModal()
                               tmpdlg.Destroy()
                           import traceback
                           traceback.print_exc(file=sys.stdout)
                           print
                           print
                           
                       # If an error arises, for now, let's interrupt the import process.  It may be possible
                       # to eliminate this line later, allowing the import to continue even if there is a problem.
                       contin = False
                       # let's build a detailed error message, if we can.
                       if 'unicode' in wx.PlatformInfo:
                           # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                           prompt = unicode(_('A problem has been detected importing a %s record'), 'utf8')
                       else:
                           prompt = _('A problem has been detected importing a %s record')
                       msg = prompt % objectType
                       if (not objectType in ['Keyword', 'ClipKeyword', 'Filter']) and (currentObj.id != ''):
                           if 'unicode' in wx.PlatformInfo:
                               # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                               prompt = unicode(_('named "%s".'), 'utf8')
                           else:
                               prompt = _('named "%s".')
                           msg = msg +  ' ' + prompt % currentObj.id
                       else:
                           msg = msg + '.'
                           # One specific error we need to trap is bogus Transcript records that have lost
                           # their parents.  This happened to at least one user.
                           if objectType == 'Transcript':
                               msg = msg + '\n' + _('The Transcript is for')
                               if currentObj.episode_num > 0:
                                    if 'unicode' in wx.PlatformInfo:
                                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                        prompt = unicode(_('Episode %d.'), 'utf8')
                                    else:
                                        prompt = _('Episode %d.')
                                    msg = msg + ' ' + prompt % currentObj.episode_num
                               elif currentObj.clip_num > 0: 
                                   if 'unicode' in wx.PlatformInfo:
                                       # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                       prompt = unicode(_('Clip %d.'), 'utf8')
                                   else:
                                       prompt = _('Clip %d.')
                                   msg = msg + ' ' + prompt % currentObj.clip_num
                               else: 
                                   if 'unicode' in wx.PlatformInfo:
                                       # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                       prompt = unicode(_('Episode 0, Clip 0, Transcript Record %d.'), 'utf8')
                                   else:
                                       prompt = _('Episode 0, Clip 0, Transcript Record %d.')
                                   msg = msg + ' ' + prompt % oldNumber
                           elif objectType == 'Keyword':
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('The record is for Keyword "%s:%s".'), 'utf8')
                                else:
                                    prompt = _('The record is for Keyword "%s:%s".')
                                msg = msg + '\n' + prompt % (currentObj.keywordGroup, currentObj.keyword)
                       if 'unicode' in wx.PlatformInfo:
                           # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                           prompt = unicode(_('You need to correct this record in XML file %s.'), 'utf8')
                           prompt2 = unicode(_('The %s record ends at line %d.'), 'utf8')
                       else:
                           prompt = _('You need to correct this record in XML file %s.')
                           prompt2 = _('The %s record ends at line %d.')
                       msg = msg + '\n' +  prompt % self.XMLFile.GetValue() + '\n' + \
                                           prompt2 % (objectType, lineCount)
                       # Display our carefully crafted error message to the user.
                       errordlg = Dialogs.ErrorDialog(None, msg)
                       errordlg.ShowModal()
                       errordlg.Destroy()
                   currentObj = None
                   objectType = None

               # Code for determining Property Type for populating Object Properties
               elif line.upper() == '<NUM>':
                   dataType = 'Num'

               elif line.upper() == '<ID>':
                   dataType = 'ID'

               elif line.upper() == '<COMMENT>':
                   dataType = 'Comment'

               elif line.upper() == '<OWNER>':
                   dataType = 'Owner'

               elif line.upper() == '<DEFAULTKEYWORDGROUP>':
                   dataType = 'DKG'

               elif line.upper() == '<SERIESNUM>':
                   dataType = 'SeriesNum'

               elif line.upper() == '<EPISODENUM>':
                   dataType = 'EpisodeNum'

               elif line.upper() == '<TRANSCRIPTNUM>':
                   dataType = 'TranscriptNum'

               elif line.upper() == '<COLLECTNUM>':
                   dataType = 'CollectNum'

               elif line.upper() == '<CLIPNUM>':
                   dataType = 'ClipNum'

               elif line.upper() == '<DATE>':
                   dataType = 'date'

               elif line.upper() == '<MEDIAFILE>':
                   dataType = 'MediaFile'

               elif line.upper() == '<LENGTH>':
                   dataType = 'Length'

               elif line.upper() == '<TITLE>':
                   dataType = 'Title'

               elif line.upper() == '<CREATOR>':
                   dataType = 'Creator'

               elif line.upper() == '<SUBJECT>':
                   dataType = 'Subject'

               elif line.upper() == '<DESCRIPTION>':
                   dataType = 'Description'

               # Because Description can be many lines long, we need to explicitly close this datatype when
               # the closing XML tag is found
               elif line.upper() == '</DESCRIPTION>':
                   dataType = None

               elif line.upper() == '<PUBLISHER>':
                   dataType = 'Publisher'

               elif line.upper() == '<CONTRIBUTOR>':
                   dataType = 'Contributor'

               elif line.upper() == '<TYPE>':
                   dataType = 'Type'

               elif line.upper() == '<FORMAT>':
                   dataType = 'Format'

               elif line.upper() == '<SOURCE>':
                   dataType = 'Source'

               elif line.upper() == '<LANGUAGE>':
                   dataType = 'Language'

               elif line.upper() == '<RELATION>':
                   dataType = 'Relation'

               elif line.upper() == '<COVERAGE>':
                   dataType = 'Coverage'

               elif line.upper() == '<RIGHTS>':
                   dataType = 'Rights'

               elif line.upper() == '<TRANSCRIBER>':
                   dataType = 'Transcriber'

               elif line.upper() == '<RTFTEXT>':
                   dataType = 'RTFText'

               # Because RTF Text can be many lines long, we need to explicitly close this datatype when
               # the closing XML tag is found
               elif line.upper() == '</RTFTEXT>':
                   dataType = None

               elif line.upper() == '<PARENTCOLLECTNUM>':
                   dataType = 'ParentCollectNum'

               elif line.upper() == '<CLIPSTART>':
                   dataType = 'ClipStart'

               elif line.upper() == '<CLIPSTOP>':
                   dataType = 'ClipStop'

               elif line.upper() == '<SORTORDER>':
                   dataType = 'SortOrder'

               elif line.upper() == '<KEYWORDGROUP>':
                   dataType = 'KWG'

               elif line.upper() == '<KEYWORD>':
                   dataType = 'KW'

               elif line.upper() == '<DEFINITION>':
                   dataType = 'Definition'

               # Because Definition Text can be many lines long, we need to explicitly close this datatype when
               # the closing XML tag is found
               elif line.upper() == '</DEFINITION>':
                   dataType = None

               elif line.upper() == '<EXAMPLE>':
                   dataType = 'Example'

               elif line.upper() == '<NOTETAKER>':
                   dataType = 'NoteTaker'

               elif line.upper() == '<NOTETEXT>':
                   dataType = 'NoteText'

               # Because Note Text can be many lines long, we need to explicitly close this datatype when
               # the closing XML tag is found
               elif line.upper() == '</NOTETEXT>':
                   dataType = None

               elif line.upper() == '<REPORTTYPE>':
                   dataType = 'ReportType'

               elif line.upper() == '<REPORTSCOPE>':
                   dataType = 'ReportScope'

               elif line.upper() == '<CONFIGNAME>':
                   dataType = 'ConfigName'

               elif line.upper() == '<FILTERDATATYPE>':
                   dataType = 'FilterDataType'

               elif line.upper() == '<FILTERDATA>':
                   dataType = 'FilterData'

               # Because Filter Data can be many lines long, we need to explicitly close this datatype when
               # the closing XML tag is found
               elif line.upper() == '</FILTERDATA>':
                   dataType = None

               # Code for populating Object Properties.
               # Unless data can stretch across mulitple lines, we should explicity undefine the dataType
               # once the data is captured.
               elif dataType == 'Num':
                   currentObj.number = int(line)
                   dataType = None

               elif dataType == 'ID':
                   currentObj.id = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Comment':
                   currentObj.comment = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Owner':
                   currentObj.owner = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'DKG':
                   currentObj.keyword_group = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'SeriesNum':
                   # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                   try:
                       currentObj.series_num = recNumbers['Series'][int(line)]
                   except:
                       pass
                   dataType = None

               elif dataType == 'EpisodeNum':
                   # We need to substitute the new Episode number for the old one.
                   # A user had a problem with a Transcript Record existing when the parent Episode
                   # had been deleted.  Therefore, let's check to see if the old Episode record existed
                   # by checking to see if the old episode number is a Key value in the Episode recNumbers table.
                   # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                   if (line.strip() != 'None'):
                       if recNumbers['Episode'].has_key(int(line)):
                           try:
                               currentObj.episode_num = recNumbers['Episode'][int(line)]
                           except:
                               pass
                       else:
                           # If the old record number doesn't exist, substitute 0 and show an error message.
                           currentObj.episode_num = 0
                           if isinstance(currentObj, ClipKeywordObject.ClipKeyword):
                               if 'unicode' in wx.PlatformInfo:
                                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                   prompt = unicode(_('Episode Number %s cannot be found for %s at line number %d.'), 'utf8')
                               else:
                                   prompt = _('Episode Number %s cannot be found for %s at line number %d.')
                               vals = (line, objectType, lineCount)
                           else:
                               if 'unicode' in wx.PlatformInfo:
                                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                   prompt = unicode(_('Episode Number %s cannot be found for %s record %d at line number %d.'), 'utf8')
                               else:
                                   prompt = _('Episode Number %s cannot be found for %s record %d at line number %d.')
                               vals = (line, objectType, currentObj.number, lineCount)
                           errordlg = Dialogs.ErrorDialog(None, prompt % vals)
                           errordlg.ShowModal()
                           errordlg.Destroy()
                   dataType = None

               elif dataType == 'TranscriptNum':
                   # if we're dealing with a CLIP's SOURCE TRANSCRIPT ...
                   if objectType == 'Clip':
                       if line != '0':
                           # ... we need to save the Clip's Source Transcript number for later, as it has been moved
                           # from the Clip object to the Clip Transcript object as of Transana-XML 1.4.
                           clipTranscripts[currentObj.number] = line
                   # To be clear, a Transcript's NUMBER goes to currentObj.number, while its TRANSCRIPTNUM
                   # is actually its SOURCE TRANSCRIPT, not its OBJECT NUMBER.
                   elif objectType == 'Transcript':
                       # Not all re-mapped Transcript Numbers are known while processing Transcripts.  Therefore, store the
                       # OLD TranscriptNum as the SourceTranscript and we'll convert it later!
                       try:
                           currentObj.source_transcript = line
                       except:
                           pass
                   # If we're dealing with anything but a Clip or Transcript ...
                   else:
                       # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                       try:
                           currentObj.transcript_num = recNumbers['Transcript'][int(line)]
                       except:
                           pass
                   dataType = None

               elif dataType == 'CollectNum':
                   # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                   try:
                       currentObj.collection_num = recNumbers['Collection'][int(line)]
                   except:
                       pass
                   dataType = None

               elif dataType == 'ClipNum':
                   # Handle object property naming inconsistency here!
                   if objectType in ['Transcript', 'Note']:
                       # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                       try:
                           currentObj.clip_num = recNumbers['Clip'][int(line)]
                       except:
                           pass
                   elif objectType == 'ClipKeyword':
                       try:
                           currentObj.clipNum = recNumbers['Clip'][int(line)]
                       except:
                           if 'unicode' in wx.PlatformInfo:
                               # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                               prompt = unicode(_('Clip Number %s cannot be found for a %s record at line number %d.\n(This is due to an incomplete Clip deletion and is not a problem.)'), 'utf8')
                           else:
                               prompt = _('Clip Number %s cannot be found for a %s record at line number %d.\n(This is due to an incomplete Clip deletion and is not a problem.)')
                           errordlg = Dialogs.ErrorDialog(None, prompt  % (line, objectType, lineCount))
                           errordlg.ShowModal()
                           errordlg.Destroy()
                   else:
                       currentObj.clipNum = recNumbers['Clip'][int(line)]
                            
                   dataType = None

               elif dataType == 'date':
                   # Importing dates can be a little tricky.  Let's trap conversion errors
                   try:
                       # If we're dealing with an Episode record ...
                       if objectType == 'Episode':
                           # Make sure it's in a form we recognize
                           # See if the date is in YYYY-MM-DD format (produced by XMLExport.py).
                           # If not, substitute slashes for dashes, as this file likely came from Delphi!
                           reStr = '\d{4}-\d+-\d+'
                           if re.compile(reStr).match(line) == None:
                               line = line.replace('-', '/')
                               timeformat = "%m/%d/%Y"
                           # The date should be stored in the Episode record as a date object
                           else:
                               timeformat = "%Y-%m-%d"
                           # Check to see if we've got extraneous time data appended.  If so, remove it!
                           # (This is reliably signalled by the presence of a space.)
                           if line.find(' ') > -1:
                               line = line.split(' ')[0]
                           # This works fine on Windows, and it works on the Mac under Python.  But on the
                           # Mac from an executable app, this line causes an ImportError exception.  If that
                           # arises, we'll have to parse the time format manually!
                           try:
                               timetuple = time.strptime(line, timeformat)
                           except ImportError:
                               # Break the string into it's components based on the divider from the timeformat.
                               tempTime = line.split(timeformat[2])
                               # create the timetuple value manually.  timeformat tells us if we're in YMD or MDY format.
                               if timeformat[1] == 'Y':
                                   timetuple = (int(tempTime[0]), int(tempTime[1]), int(tempTime[2]), 0, 0, 0, 1, 107, -1)
                               else:
                                   timetuple = (int(tempTime[2]), int(tempTime[0]), int(tempTime[1]), 0, 0, 0, 1, 107, -1)
                           currentObj.tape_date = datetime.datetime(*timetuple[:7])                           
                       # If we're dealing with a Core Data record ...
                       elif objectType == 'CoreData':
                           # Unfortunately, the Delphi exporter for 1.2x data and the Python exporter for 2.x data
                           # produce dates in different formats.  (Oops.  Sorry about that.)
                           # If the form is from Delphi, D-M-Y, we need to rearrange it into MM/DD/Y format
                           if line.find('-') > -1:
                               dateParts = line.split('-')
                               date = '%02d/%02d/%d' % (int(dateParts[1]), int(dateParts[0]), int(dateParts[2]))
                           # Otherwise, the format from Python is already MM/DD/Y form, so we should be okay
                           else:
                               date = line
                           # These dates are stored internally in the Core Data record as formatted strings.
                           currentObj.dc_date = date
                   except:
                       import traceback
                       traceback.print_exc(file=sys.stdout)
                       # Display the Exception Message, allow "continue" flag to remain true
                       if 'unicode' in wx.PlatformInfo:
                           # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                           prompt = unicode(_('Date Import Failure: "%s"'), 'utf8')
                           prompt2 = unicode(_("Exception %s: %s"), 'utf8')
                       else:
                           prompt = _('Date Import Failure: "%s"')
                           prompt2 = _("Exception %s: %s")
                       msg = prompt % line +'\n' + prompt2 % (sys.exc_info()[0], sys.exc_info()[1])
                       errordlg = Dialogs.ErrorDialog(None, msg)
                       errordlg.ShowModal()
                       errordlg.Destroy()
                       
                   dataType = None

               elif dataType == 'MediaFile':
                   currentObj.media_filename = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Length':
                   currentObj.tape_length = line
                   dataType = None

               elif dataType == 'Title':
                   currentObj.title = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Creator':
                   currentObj.creator = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Subject':
                   currentObj.subject = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Description':
                   # If this is not our first line, add a newline character before our new text.  Otherwise, all the
                   # text is added as a single line.
                   if currentObj.description != '':
                       currentObj.description = currentObj.description + '\n'
                   currentObj.description = currentObj.description + self.ProcessLine(line)
                   # We DO NOT reset DataType here, as Description may be many lines long!
                   # dataType = None

               elif dataType == 'Publisher':
                   currentObj.publisher = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Contributor':
                   currentObj.contributor = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Type':
                   currentObj.dc_type = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Format':
                   currentObj.format = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Source':
                   currentObj.source = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Language':
                   currentObj.language = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Relation':
                   currentObj.relation = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Coverage':
                   currentObj.coverage = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Rights':
                   currentObj.rights = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Transcriber':
                   currentObj.transcriber = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'RTFText':
                   # Add Line Breaks to the text to match the incoming lines.
                   # Otherwise, the transcript might be messed up, with the first word of the next line
                   # being truncated.
                   if currentObj.text <> '':
                       currentObj.text = currentObj.text + '\n'
                   currentObj.text = currentObj.text + line
                   # We DO NOT reset DataType here, as RTFText may be many lines long!
                   # dataType = None

               elif dataType == 'ParentCollectNum':
                   currentObj.parent = recNumbers['Collection'][int(line)]
                   dataType = None

               elif dataType == 'ClipStart':
                   currentObj.clip_start = int(line)
                   # If we have a Clip object ...
                   if objectType == 'Clip':
                       if line != '0':
                           # ... save the Clip Start time so it can be added to the Clip Transcript record too.
                           clipStartStop[(currentObj.number, 'Start')] = int(line)
                   dataType = None

               elif dataType == 'ClipStop':
                   currentObj.clip_stop = int(line)
                   # If we have a Clip object ...
                   if objectType == 'Clip':
                       if line != '0':
                           # ... save the Clip Stop time so it can be added to the Clip Transcript record too.
                           clipStartStop[(currentObj.number, 'Stop')] = int(line)
                   dataType = None

               elif dataType == 'SortOrder':
                   currentObj.sort_order = line
                   dataType = None

               elif dataType == 'KWG':
                   currentObj.keywordGroup = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'KW':
                   currentObj.keyword = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'Definition':
                   # If this is not our first line, add a newline character before our new text.  Otherwise, all the
                   # text is added as a single line.
                   if currentObj.definition != '':
                       currentObj.definition = currentObj.definition + '\n'
                   # We changed the way the definition was stored in the database for Transana 2.30, XML Version 1.4.
                   if (self.XMLVersionNumber in ['1.1', '1.2', '1.3']):
                       currentObj.definition = currentObj.definition + self.ProcessLine(line)
                   else:
                       currentObj.definition = currentObj.definition + line.decode(self.importEncoding)
                   # We DO NOT reset DataType here, as Definition may be many lines long!
                   # dataType = None

               elif dataType == 'Example':
                   currentObj.example = line
                   dataType = None

               elif dataType == 'NoteTaker':
                   currentObj.author = self.ProcessLine(line)
                   dataType = None

               elif dataType == 'NoteText':
                   # If this is not our first line, add a newline character before our new text.  Otherwise, all the
                   # text is added as a single line.
                   if currentObj.text != '':
                       currentObj.text = currentObj.text + '\n'
                   currentObj.text = currentObj.text + line.decode(self.importEncoding)
                   # We DO NOT reset DataType here, as NoteText may be many lines long!
                   # dataType = None

               elif dataType == 'ReportType':
                    self.FilterReportType = line
                    dataType = None

               elif dataType == 'ReportScope':
                    if self.FilterReportType in ['5', '6', '7', '10', '14']:
                        self.FilterScope = recNumbers['Series'][int(line)]
                    elif self.FilterReportType in ['1', '2', '3', '8', '11']:
                        self.FilterScope = recNumbers['Episode'][int(line)]
                    # Collection Clip Data Export (ReportType 4) only needs translation if ReportScope != 0
                    elif (self.FilterReportType in ['12']) or ((self.FilterReportType == '4') and (int(line) != 0)):
                        self.FilterScope = recNumbers['Collection'][int(line)]
                    elif self.FilterReportType in ['13']:
                        # FilterScopes for ReportType 13 (Notes Report) are constants, not object numbers!
                        self.FilterScope = int(line)
                    # Collection Clip Data Export (ReportType 4) for ReportScope 0, the Collection Root, needs
                    # a FilterScope of 0
                    # Saved Search (ReportType 15) needs no modifications, but setting FilterScope to 0 allows the SAVE!
                    elif ((self.FilterReportType == '4') and (int(line) == 0)) or (self.FilterReportType == '15'):
                        self.FilterScope = 0
                    else:
                       if 'unicode' in wx.PlatformInfo:
                           # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                           prompt = unicode(_('An error occurred during Database Import.\nThere is an unsupported Filter Report Type (%s) in the Filter table. \nYou may wish to upgrade Tranana and try again.'), 'utf8')
                       else:
                           prompt = _('An error occurred during Database Import.\nThere is an unsupported Filter Report Type (%s) in the Filter table. \nYou may wish to upgrade Tranana and try again.')
                       errordlg = Dialogs.ErrorDialog(self, prompt % (self.FilterReportType))
                       errordlg.ShowModal()
                       errordlg.Destroy()
                    dataType = None

               elif dataType == 'ConfigName': 
                    self.FilterConfigName = DBInterface.ProcessDBDataForUTF8Encoding(line)
                    dataType = None

               elif dataType == 'FilterDataType': 
                    self.FilterFilterDataType = line
                    dataType = None

               elif dataType == 'FilterData': 
                   # If this is not our first line, add a newline character before our new text.  Otherwise, all the
                   # text is added as a single line.
                   if self.FilterFilterData != '':
                       self.FilterFilterData += '\n'
                   # Starting with XML Version 1.3, we have to deal with encoding issues differently
                   if self.XMLVersionNumber in ['1.0', '1.1', '1.2']:
                       # We just add the encoded line, without the ProcessLine() call, because Filter Data is stored in the Database
                       # as a BLOB, and thus is encoded and handled differently.
                       self.FilterFilterData += unicode(line, self.importEncoding)
                   # Starting with XML Version 1.3 ...
                   else:
                       # we don't encode on a line by line basis for FilterData
                       self.FilterFilterData += line
                   # We DO NOT reset DataType here, as Filter Data may be many lines long!
                   # dataType = None

               elif dataType == 'XMLVersionNumber':
                   self.XMLVersionNumber = line

               # If we're not continuing, stop processing! 
               if not contin:
                   break

           if contin: 
               # Since Clips were imported before Transcripts, the Originating Transcript Numbers in the Clip Records
               # are incorrect.  We must update them now.
               progress.Update(91, _('Updating Source Transcript Numbers in Clip Transcript records'))
               if db != None:
                   dbCursor = db.cursor()
                   dbCursor2 = db.cursor()
                   SQLText = 'SELECT TranscriptNum, SourceTranscriptNum, ClipNum FROM Transcripts2 WHERE ClipNum > 0'
                   dbCursor.execute(SQLText)
                   for (TranscriptNum, SourceTranscriptNum, ClipNum) in dbCursor.fetchall():
                       SQLText = """ UPDATE Transcripts2
                                     SET SourceTranscriptNum = %s
                                     WHERE TranscriptNum = %s """
                       # It is possible that the originating Transcript has been deleted.  If so,
                       # we need to set the TranscriptNum to 0.  We accomplish this by adding the
                       # missing Transcript Number to our recNumbers list with a value of 0
                       if not(SourceTranscriptNum in recNumbers['Transcript'].keys()):
                           recNumbers['Transcript'][SourceTranscriptNum] = 0
                       dbCursor2.execute(SQLText, (recNumbers['Transcript'][SourceTranscriptNum], TranscriptNum))

                   dbCursor.close()
                   dbCursor2.close()

               # If we made it this far, we can commit the database transaction
               SQLText = 'COMMIT'
           else:
               # If contin is False, there's been an error and we should roll back the database transaction
               SQLText = 'ROLLBACK'
           # Execute the COMMIT or ROLLBACK
           dbCursor = db.cursor()
           dbCursor.execute(SQLText)
           dbCursor.close()
           
       except:
           if DEBUG:
               import traceback
               traceback.print_exc(file=sys.stdout)
           if 'unicode' in wx.PlatformInfo:
               # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
               prompt = unicode(_('An error occurred during Database Import.\n%s\n%s'), 'utf8')
           else:
               prompt = _('An error occurred during Database Import.\n%s\n%s')
           errordlg = Dialogs.ErrorDialog(self, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
           errordlg.ShowModal()
           errordlg.Destroy()
           dbCursor = db.cursor()
           SQLText = 'ROLLBACK'
           dbCursor.execute(SQLText)
           dbCursor.close()

       try:
           f.close()
       except:
           pass

       TransanaGlobal.menuWindow.ControlObject.DataWindow.DBTab.tree.refresh_tree()

       progress.Update(100)
       progress.Destroy()


   def ProcessLine(self, txt):
       """ Process most lines read from the XML file to apply the proper encoding, if needed. """
       if 'unicode' in wx.PlatformInfo:
           # If we've got a String instead of a Unicode object ...
           if type(txt) == str:
               # ... convert the string to Unicode using the import encoding
               txt = unicode(txt, self.importEncoding)
           # If we're not reading a file encoded with UTF-8 encoding, we need to ...
           if (self.importEncoding != 'utf8'):
               # ... and then convert it to UTF-8
               txt = txt.encode('utf8')
           # Now perform the UTF-8 encoding needed for the database.
           txt = DBInterface.ProcessDBDataForUTF8Encoding(txt)
       # Return the encoded text to the calling method
       return(txt)


   def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        fs = wx.FileSelector(_("Select an XML file to import"),
                        TransanaGlobal.configData.videoPath,
                        "",
                        "", 
                        _("Transana-XML Files (*.tra)|*.tra|XML Files (*.xml)|*.xml|All files (*.*)|*.*"), 
                        wx.OPEN | wx.FILE_MUST_EXIST)
        # If user didn't cancel ..
        if fs != "":
            self.XMLFile.SetValue(fs)

   def CloseWindow(self, event):
       self.Close()


# This simple derrived class let's the user drop files onto an edit box
class EditBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, editbox):
        wx.FileDropTarget.__init__(self)
        self.editbox = editbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        self.editbox.SetValue(files[0])


if __name__ == '__main__':
    class MyApp(wx.App):
       def OnInit(self):
          frame = XMLImport(None, -1, "Transana XML Import")
          self.SetTopWindow(frame)
          return True
          

    app = MyApp(0)
    app.MainLoop()
