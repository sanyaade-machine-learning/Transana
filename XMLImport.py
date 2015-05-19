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

import Dialogs
import Series
import Episode
import CoreData
import Transcript
import Collection
import Clip
import Keyword
import ClipKeywordObject
import Note
import DBInterface
import TransanaGlobal

# import Python's os and sys modules
import os
import sys
# import Python's Regular Expression parser
import re

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
       self.XMLFile = self.new_edit_box(_("XML Filename"), lay, '')
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
       XMLVersionNumber = 0.0
       recNumbers = {}
       recNumbers['Series'] = {0:0}
       recNumbers['Episode'] = {0:0}
       recNumbers['Transcript'] = {0:0}
       recNumbers['Collection'] = {0:0}
       recNumbers['Clip'] = {0:0}

       clipTranscripts = {}

       db = DBInterface.get_db()
       if db != None:
           # Begin Database Transaction 
           dbCursor = db.cursor()
           SQLText = 'BEGIN'
           dbCursor.execute(SQLText)
           dbCursor.close()

       lineCount = 0 
       
       try:
           contin = True
           f = file(self.XMLFile.GetValue(), 'r')

           objectType = None 
           dataType = None
           for line in f:
               lineCount += 1
               line = line.strip()

               if DEBUG:
                   print "Line %d: '%s' %s %s" % (lineCount, line, objectType, dataType)

               # Code for updating the Progress Bar

               if line.upper() == '<SERIESFILE>':
                   progress.Update(0, _('Importing Series records'))
               elif line.upper() == '<EPISODEFILE>':
                   progress.Update(10, _('Importing Episode records'))
               elif line.upper() == '<COREDATAFILE>':
                   progress.Update(20, _('Importing Core Data records'))
               elif line.upper() == '<COLLECTIONFILE>':
                   progress.Update(30, _('Importing Collection records'))
               elif line.upper() == '<CLIPFILE>':
                   progress.Update(40, _('Importing Clip records'))
               elif line.upper() == '<TRANSCRIPTFILE>':
                   progress.Update(50, _('Importing Transcript records (This may be slow because of the size of Transcript records.)'))
               elif line.upper() == '<KEYWORDFILE>':
                   progress.Update(60, _('Importing Keyword records'))
               elif line.upper() == '<CLIPKEYWORDFILE>':
                   progress.Update(70, _('Importing Clip Keyword records'))
               elif line.upper() == '<NOTEFILE>':
                   progress.Update(80, _('Importing Note records'))

               # Transana XML Version Checking
               elif line.upper() == '<TRANSANAXMLVERSION>':
                   objectType = None
                   dataType = 'XMLVersionNumber'

               elif line.upper() == '</TRANSANAXMLVERSION>':
                   # importEncoding reflects the encoding used to create the Transana XML file now being imported.
                   # Version 1.0 -- Original Transana XML for Transana 2.0 release
                   # Version 1.1 -- Unicode encoding added to Transana XML for Transana 2.1 release
                   if XMLVersionNumber == '1.0':
                       self.importEncoding = 'latin-1'
                   elif XMLVersionNumber in ['1.1']:
                       self.importEncoding = 'utf8'
                   else: 
                       msg = _('The Database you are trying to import was created with a later version\nof Transana.  Please upgrade your copy of Transana and try again.')
                       dlg = Dialogs.ErrorDialog(None, msg)
                       dlg.ShowModal()
                       dlg.Destroy()
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

               elif line.upper() == '</SERIES>' or \
                    line.upper() == '</EPISODE>' or \
                    line.upper() == '</COREDATA>' or \
                    line.upper() == '</TRANSCRIPT>' or \
                    line.upper() == '</COLLECTION>' or \
                    line.upper() == '</CLIP>' or \
                    line.upper() == '</KEYWORDREC>' or \
                    line.upper() == '</CLIPKEYWORD>' or \
                    line.upper() == '</NOTE>':
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
                          objectType == 'Clip':
                           oldNumber = currentObj.number
                           currentObj.number = 0

                           # TO prevent the ClipObject from trying to save the Clip Transcript,
                           # (which it doesn't yet have in the import),
                           # we must set its clip_transcript_num to -1.
                           if objectType == 'Clip':
                               currentObj.clip_transcript_num = -1

#                           if DEBUG and (objectType == 'Transcript'):
#                               tmpdlg = wx.MessageDialog(self, currentObj.__repr__())
#                               tmpdlg.ShowModal()
#                               tmpdlg.Destroy()
                               
                           currentObj.db_save()
                           # Let's keep a record of the old and new object numbers for each object saved.
                           recNumbers[objectType][oldNumber] = currentObj.number
                           
                       elif  objectType == 'CoreData' or \
                             objectType == 'Note':
                           currentObj.number = 0

#                           if DEBUG and (objectType == 'Note'):
#                               tmpdlg = wx.MessageDialog(self, currentObj.__repr__())
#                               tmpdlg.ShowModal()
#                               tmpdlg.Destroy()
                               
                           currentObj.db_save()

                       elif objectType == 'Keyword':
                           currentObj.db_save()
                       elif objectType == 'ClipKeyword':
                           if (currentObj.episodeNum != 0) or (currentObj.clipNum != 0):
                               currentObj.db_save()
                   except:
                       if DEBUG:
                           import traceback
                           traceback.print_exc(file=sys.stdout)
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
                       if (not objectType in ['Keyword', 'ClipKeyword']) and (currentObj.id != ''):
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

               # Code for determing Property Type for populating Object Properties

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

               # Code for populating Object Properties

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
                   currentObj.series_num = recNumbers['Series'][int(line)]
                   dataType = None

               elif dataType == 'EpisodeNum':
                   # We need to substitute the new Episode number for the old one.
                   # A user had a problem with a Transcript Record existing when the parent Episode
                   # had been deleted.  Therefore, let's check to see if the old Episode record existed
                   # by checking to see if the old episode number is a Key value in the Episode recNumbers table.
                   if recNumbers['Episode'].has_key(int(line)):
                       currentObj.episode_num = recNumbers['Episode'][int(line)]
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
                   if objectType == 'Clip':
                       currentObj.transcript_num = line
                       if line != '0':
                           clipTranscripts[currentObj.number] = line
                   else:
                       currentObj.transcript_num = recNumbers['Transcript'][int(line)]
                   dataType = None

               elif dataType == 'CollectNum':
                   currentObj.collection_num = recNumbers['Collection'][int(line)]
                   dataType = None

               elif dataType == 'ClipNum':
                   if objectType in ['Transcript', 'Note']:
                       currentObj.clip_num = recNumbers['Clip'][int(line)]
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
#                           # Make sure it's in a form compatible with mx.DateTime.
#                           # mxDateTime appears to accept YYYY-MM-DD and MM/DD/YYYY as legitimate forms.
                           # Make sure it's in a form we recognize
                           # See if the date is in YYYY-MM-DD format (produced by XMLExport.py).
                           # If not, substitute slashes for dashes, as this file likely came from Delphi!
                           reStr = '\d{4}-\d+-\d+'
                           if re.compile(reStr).match(line) == None:
                               line = line.replace('-', '/')
#                           # The date should be stored in the Episode record as a mx.DateTime object
#                           currentObj.tape_date = mx.DateTime.DateTimeFrom(line)
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
                   currentObj.clip_start = line
                   dataType = None

               elif dataType == 'ClipStop':
                   currentObj.clip_stop = line
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
                   currentObj.definition = currentObj.definition + self.ProcessLine(line)

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
                   # We just add the encoded line, without the ProcessLine() call, because NoteText is stored in the Database
                   # as a BLOB, and thus is encoded and handled differently.
                   currentObj.text = currentObj.text + unicode(line, self.importEncoding)
                   # We DO NOT reset DataType here, as NoteText may be many lines long!
                   # dataType = None

               elif dataType == 'XMLVersionNumber':
                   XMLVersionNumber = line

               # If we're not continuing, stop processing! 
               if not contin:
                   break

           if contin: 
               # Since Clips were imported before Transcripts, the Originating Transcript Numbers in the Clip Records
               # are incorrect.  We must update them now.
               progress.Update(90, _('Updating Transcript Numbers in Clip records'))
               if db != None:
                   dbCursor = db.cursor()
                   dbCursor2 = db.cursor()
                   SQLText = 'SELECT ClipNum, TranscriptNum FROM Clips2'
                   dbCursor.execute(SQLText)
                   for (ClipNum, TranscriptNum) in dbCursor.fetchall():
                       SQLText = """ UPDATE Clips2
                                     SET TranscriptNum = %s
                                     WHERE ClipNum = %s """
                       # It is possible that the originating Transcript has been deleted.  If so,
                       # we need to set the TranscriptNum to 0.  We accomplish this by adding the
                       # missing Transcript Number to our recNumbers list with a value of 0
                       if not(TranscriptNum in recNumbers['Transcript'].keys()):
                           recNumbers['Transcript'][TranscriptNum] = 0
                       dbCursor2.execute(SQLText, (recNumbers['Transcript'][TranscriptNum], ClipNum))

                   dbCursor.close()
                   dbCursor2.close()

               SQLText = 'COMMIT'
           else:
               SQLText = 'ROLLBACK'
               
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

       f.close()

       TransanaGlobal.menuWindow.ControlObject.DataWindow.DBTab.tree.refresh_tree()

       progress.Update(100)
       progress.Destroy()

#       print
#       print "recNumber Records:"
#       for record in recNumbers:
#           for val in recNumbers[record]:
#               print record, val, recNumbers[record][val]

   def ProcessLine(self, txt):
       """ Process most lines read from the XML file to apply the proper encoding, if needed. """
       if 'unicode' in wx.PlatformInfo:
           # If we're not reading a file encoded with UTF-8 encoding, we need to ...
           if self.importEncoding != 'utf8':
               # ... convert the string to Unicode using the import encoding
               unicodeTxt = unicode(txt, self.importEncoding)
               # ... and then convert it to UTF-8
               txt = unicodeTxt.encode('utf8')
           # Now perform the UTF-8 encoding needed for the database.
           txt = DBInterface.ProcessDBDataForUTF8Encoding(txt)
       # Return the encoded text to the calling method
       return(txt)


   def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        fs = wx.FileSelector(_("Select an XML file to import"),
                        os.path.dirname(sys.argv[0]),
                        "",
                        "", 
                        _("XML Files (*.xml)|*.xml|All files (*.*)|*.*"), 
                        wx.OPEN | wx.FILE_MUST_EXIST)
        # If user didn't cancel ..
        if fs != "":
            self.XMLFile.SetValue(fs)

   def CloseWindow(self, event):
       self.Close()

#   def get_input(self):
#        """Custom input routine."""
#        # Inherit base input routine
#        gen_input = Dialogs.GenForm.get_input(self)
#        if gen_input:
#            self.XMLFilename = gen_input[_('XML Filename')]
#        else:
#            self.XMLFilename = None
#        return self



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
