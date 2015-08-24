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

"""This module implements the Data Import for Transana based on the Transana XML schema."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'
# Patch sent by David Fraser to eliminate need for mx module

DEBUG = False
DEBUG_Exceptions = True
if DEBUG or DEBUG_Exceptions:
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
import Document
import Episode
import KeywordObject as Keyword
import Misc
import Note
import Quote
import Library
import Snapshot
import TransanaConstants
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
    def __init__(self, parent, id, title, importData=None):
        """ Initialize the XML Import dialog and framework.
            importData can contain the name of a Transana-XML file and the encoding used for that file, in which case
            the XML Import Dialog does not need to be displayed to the user. """
        # Create the Dialog Box
        Dialogs.GenForm.__init__(self, parent, id, title, (550,150), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Import Database')

        # Remember the import data, if any is passed
        self.importData = importData
        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create the main form prompt
        prompt = _('Please select a Transana XML File to import.')
        importText = wx.StaticText(self.panel, -1, prompt)

        # Add the import message to the dialog box
        r1Sizer.Add(importText, 0)
        
        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Add the Import File Name element
        self.XMLFile = self.new_edit_box(_("Transana-XML Filename"), v1, '')
        # If importData is provided ...
        if importData != None:
            # ... the first element is the Transana-XML file name to be imported
            self.XMLFile.SetValue(importData[0])
        # Make this text box a File Drop Target
        self.XMLFile.SetDropTarget(EditBoxFileDropTarget(self.XMLFile))

        # Add the element sizer to the row sizer
        r2Sizer.Add(v1, 1, wx.EXPAND)

        # Add a spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Browse button
        browse = wx.Button(self.panel, wx.ID_FILE1, _("Browse"), wx.DefaultPosition)
        # Add the Browse Method to the Browse Button
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)

        # Add the element to the row sizer
        r2Sizer.Add(browse, 0, wx.ALIGN_BOTTOM)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            r2Sizer.Add((2, 0))

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # We have encoding issues on import.  At the very least, single-user Transana 2.42 on Windows using
        # the Chinese prompts export data using GBK instead of UTF-8.  What can I say?  I screwed up.
        # To fix this, we need to add an option for specifying the import encoding to be used.

        # We need to know the encoding of the import file, which differs depending on the
        # TransanaXML Version, the Transana version, and the language used during export.
        # Let's assume UTF-8 unless we have to change it.
        self.importEncoding = 'utf8'

        # Add a Horizontal sizer
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a Vertical sizer to go in the Horizontal sizer
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Define the options for the Encoding choice box.  This must be done as two parallel lists, rather than a dictionary
        choices = [_('Most Transana database export files'),
                   _('Chinese data from single-user Transana 2.1 - 2.42 on Windows'),
                   _('Russian data from single-user Transana 2.1 - 2.42 on Windows'),
                   _('Eastern European data from single-user Transana 2.1 - 2.42 on Windows'),
                   _('Greek data from single-user Transana 2.1 - 2.42 on Windows'),
                   _('Japanese data from single-user Transana 2.1 - 2.42 on Windows'),
                   _('Transana 2.0 to Transana 2.05 database export files')]
        # Use a matching list to define the encodings that go with each of the Encoding options
        self.encodingOptions = ['utf8', 'gbk', 'koi8_r', 'iso8859_2', 'iso8859_7', 'cp932', 'latin1']
        # Create a Choice Box where the user can select an import encoding, based on information about how the
        # Transana-XML file in question was created.  This adds it to the Vertical Sizer created above.
        self.chImportEncoding = self.new_choice_box(_('Exported by:'), v2, choices, 0)
        # If importData is provided ...
        if importData != None:
            # ... set the Import Encoding based on the second element of importData
            if importData[1] == 'utf8':
                self.chImportEncoding.SetSelection(0)
            elif importData[1] == 'gbk':
                self.chImportEncoding.SetSelection(1)
            elif importData[1] == 'koi8_r':
                self.chImportEncoding.SetSelection(2)
            elif importData[1] == 'iso8859_2':
                self.chImportEncoding.SetSelection(3)
            elif importData[1] == 'iso8859_7':
                self.chImportEncoding.SetSelection(4)
            elif importData[1] == 'cp932':
                self.chImportEncoding.SetSelection(5)
            elif importData[1] == 'latin1':
                self.chImportEncoding.SetSelection(6)

        # Add the Vertical sizer to the Horizontal sizer
        r3Sizer.Add(v2, 1, wx.EXPAND)
        # Add the Horizontal Sizer to the form's main sizer.
        mainSizer.Add(r3Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the buttons
        self.create_buttons(sizer=btnSizer)
        # Add the button sizer to the main sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            mainSizer.Add((0, 2))

        # Set the PANEL's main sizer
        self.panel.SetSizer(mainSizer)
        # Tell the PANEL to auto-layout
        self.panel.SetAutoLayout(True)
        # Lay out the Panel
        self.panel.Layout()
        # Lay out the panel on the form
        self.Layout()
        # Resize the form to fit the contents
        self.Fit()

        # Get the new size of the form
        (width, height) = self.GetSizeTuple()
        # Reset the form's size to be at least the specified minimum width
        self.SetSize(wx.Size(max(550, width), height))
        # Define the minimum size for this dialog as the current size, and define height as unchangeable
        self.SetSizeHints(max(550, width), height, -1, height)
        # Center the form on screen
        TransanaGlobal.CenterOnPrimary(self)

        # Set focus to the XML file field
        self.XMLFile.SetFocus()

    def UnEscape(self, inpStr):
        """ Replaces "&amp;", "&gt;", and "&lt;" with "&", ">", and "<" 
            >, <, and & all need to be replaced, but &amp;, &gt;, and &lt; needs to survive!"""
        # Find the first Greater Than in the document
        chrPos = inpStr.find('&gt;')
        # While there are more Greater Thans ...
        while (chrPos > -1):
            # ... replace the Greater Than with the escape string ...
            inpStr = inpStr[:chrPos] + '>' + inpStr[chrPos + 4:]
            # ... and look for the NEXT Greater Than after the replacement
            chrPos = inpStr.find('&gt;', chrPos + 1)
        # Find the first Less Than in the document
        chrPos = inpStr.find('&lt;')
        # While there are more Less Thans ...
        while (chrPos > -1):
            # ... replace the Less Than with the escape string ...
            inpStr = inpStr[:chrPos] + '<' + inpStr[chrPos + 4:]
            # ... and look for the NEXT Less Than after the replacement
            chrPos = inpStr.find('&lt;', chrPos + 1)
        # Find the first Ampersand in the document
        chrPos = inpStr.find('&amp;')
        # While there are more Ampersands ...
        while (chrPos > -1):
            # ... replace the Ampersand with the escape string ...
            inpStr = inpStr[:chrPos] + '&' + inpStr[chrPos + 5:]
            # ... and look for the NEXT Ampersand after the replacement
            chrPos = inpStr.find('&amp;', chrPos + 1)
        # Return the modified string
        return inpStr

    def Import(self):
       """ Handle the Import request """
       # use the LONGEST title here to set the width of the dialog box!
       progress = wx.ProgressDialog(_('Transana XML Import'),
                                    _('Importing Transcript records (This may be slow because of the size of Transcript records.)') + '\nTest',
                                    parent=self,
                                    style = wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)

       # Initialize the Transana-XML Version.  This allows differential processing based on the version used to
       # create the Transana-XML data file
       self.XMLVersionNumber = 0.0
       # Initialize dictionary variables used to keep track of data values that may change during import.
       recNumbers = {}
       recNumbers['Libraries'] = {0:0}
       recNumbers['Document'] = {0:0}
       recNumbers['Episode'] = {0:0}
       recNumbers['Transcript'] = {0:0}
       recNumbers['Collection'] = {0:0}
       recNumbers['Quote'] = {0:0}
       recNumbers['Clip'] = {0:0}
       recNumbers['OldClip'] = {0:0}
       recNumbers['Snapshot'] = {0:0}
       recNumbers['Note'] = {0:0}
       quotePosition = {}
       clipTranscripts = {}
       clipStartStop = {}

       # Get the database connection
       db = DBInterface.get_db()
       if db != None:
           # Begin Database Transaction 
           dbCursor = db.cursor()
           SQLText = 'BEGIN'
           dbCursor.execute(SQLText)

#           print "XMLImport - Begin Transaction"

       # We need to track the number of lines read and processed from the input file.
       lineCount = 0
       objCountNumber = 1
       
       # We need to track how many Transcript records were in the database PRIOR to import!
       # (Needed for merging non-overlapping databases!  Otherwise, we lose SourceTranscriptNum information.)

       # First, let's find out the largest Transcript Number currently in use
       SQLText = 'SELECT MAX(TranscriptNum) FROM Transcripts2'
       dbCursor.execute(SQLText)
       tmpVal = dbCursor.fetchone()
       # Let's remember that largest transcript number
       transcriptCount = tmpVal[0]
       # If we have an EMPTY database, we get "None" rather than 0.
       if transcriptCount == None:
           # Fix that.
           transcriptCount = 0

       # Define some dictionaries for processing XML tags
       MainHeads = {}
       MainHeads['<SERIESFILE>'] =               { 'progPct' : 0,
                                                   'progPrompt' : _('Importing Library records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<DOCUMENTFILE>'] =             { 'progPct' : 5,
                                                   'progPrompt' : _('Importing Document records (This may be slow because of the size of Document records.)'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<EPISODEFILE>'] =              { 'progPct' : 10,
                                                   'progPrompt' : _('Importing Episode records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<COREDATAFILE>'] =             { 'progPct' : 14,
                                                   'progPrompt' : _('Importing Core Data records'),
                                                   'skipCheck' : True,
                                                   'skipValue' : False }
       MainHeads['<COLLECTIONFILE>'] =           { 'progPct' : 19,
                                                   'progPrompt' : _('Importing Collection records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<QUOTEFILE>'] =                { 'progPct' : 24,
                                                   'progPrompt' : _('Importing Quote records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<QUOTEPOSITIONFILE>'] =        { 'progPct' : 29,
                                                   'progPrompt' : _('Importing Quote Position records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<CLIPFILE>'] =                 { 'progPct' : 33,
                                                   'progPrompt' : _('Importing Clip records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<ADDITIONALVIDSFILE>'] =       { 'progPct' : 38,
                                                   'progPrompt' : _('Importing Additional Video records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<TRANSCRIPTFILE>'] =           { 'progPct' : 43,
                                                   'progPrompt' : _('Importing Transcript records (This may be slow because of the size of Transcript records.)'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<SNAPSHOTFILE>'] =             { 'progPct' : 48,
                                                   'progPrompt' : _('Importing Snapshot records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<KEYWORDFILE>'] =              { 'progPct' : 52,
                                                   'progPrompt' : _('Importing Keyword records'),
                                                   'skipCheck' : True,
                                                   'skipValue' : False }
       MainHeads['<CLIPKEYWORDFILE>'] =          { 'progPct' : 57,
                                                   'progPrompt' : _('Importing Clip Keyword records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<SNAPSHOTKEYWORDFILE>'] =      { 'progPct' : 62,
                                                   'progPrompt' : _('Importing Snapshot Keyword records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<SNAPSHOTKEYWORDSTYLEFILE>'] = { 'progPct' : 67,
                                                   'progPrompt' : _('Importing Snapshot Coding Style records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<NOTEFILE>'] =                 { 'progPct' : 71,
                                                   'progPrompt' : _('Importing Note records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }
       MainHeads['<FILTERFILE>'] =               { 'progPct' : 76,
                                                   'progPrompt' : _('Importing Filter records'),
                                                   'skipCheck' : False,
                                                   'skipValue' : False }

       # Because Definition, Description, FilterData, NoteText, RTFText, XMLText can be many lines long, we need to 
       # explicitly close these datatypes when the closing XML tag is found
       DataTypes = { '<AUDIO>'               : 'Audio',
                     '<CLIPNUM>'             : 'ClipNum',
                     '<CLIPSTART>'           : 'ClipStart',
                     '<CLIPSTOP>'            : 'ClipStop',
                     '<CREATOR>'             : 'Creator',
                     '<COLLECTNUM>'          : 'CollectNum',
                     '<COLORDEF>'            : 'ColorDef',
                     '<COLORNAME>'           : 'ColorName',
                     '<COMMENT>'             : 'Comment',
                     '<CONFIGNAME>'          : 'ConfigName',
                     '<CONTRIBUTOR>'         : 'Contributor',
                     '<COVERAGE>'            : 'Coverage',
                     '<DATE>'                : 'date',
                     '<DEFAULTKEYWORDGROUP>' : 'DKG',
                     '<DEFINITION>'          : 'Definition',
                     '</DEFINITION>'         : None,
                     '<DESCRIPTION>'         : 'Description',
                     '</DESCRIPTION>'        : None,
                     '<DOCUMENTNUM>'         : 'DocumentNum',
                     '<DRAWMODE>'            : 'DrawMode',
                     '<ENDCHAR>'             : 'EndChar',
                     '<EPISODENUM>'          : 'EpisodeNum',
                     '<EXAMPLE>'             : 'Example',
                     '<FILTERDATA>'          : 'FilterData',
                     '</FILTERDATA>'         : None,
                     '<FILTERDATATYPE>'      : 'FilterDataType',
                     '<FORMAT>'              : 'Format',
                     '<ID>'                  : 'ID',
                     '<IMAGECOORDSX>'        : 'ImageCoordsX',
                     '<IMAGECOORDSY>'        : 'ImageCoordsY',
                     '<IMAGESCALE>'          : 'ImageScale',
                     '<IMAGESIZEH>'          : 'ImageSizeH',
                     '<IMAGESIZEW>'          : 'ImageSizeW',
                     '<KEYWORD>'             : 'KW',
                     '<KEYWORDGROUP>'        : 'KWG',
                     '<LANGUAGE>'            : 'Language',
                     '<LENGTH>'              : 'Length',
                     '<LINESTYLE>'           : 'LineStyle',
                     '<LINEWIDTH>'           : 'LineWidth',
                     '<MEDIAFILE>'           : 'MediaFile',
                     '<MINTRANSCRIPTWIDTH>'  : 'MinTranscriptWidth',
                     '<NOTETAKER>'           : 'NoteTaker',
                     '<NOTETEXT>'            : 'NoteText',
                     '</NOTETEXT>'           : None,
                     '<NUM>'                 : 'Num',
                     '<OFFSET>'              : 'Offset',
                     '<OWNER>'               : 'Owner',
                     '<PARENTCOLLECTNUM>'    : 'ParentCollectNum',
                     '<PUBLISHER>'           : 'Publisher',
                     '<QUOTENUM>'            : 'QuoteNum',
                     '<RELATION>'            : 'Relation',
                     '<REPORTSCOPE>'         : 'ReportScope',
                     '<REPORTTYPE>'          : 'ReportType',
                     '<RIGHTS>'              : 'Rights',
                     '<RTFTEXT>'             : 'RTFText',
                     '</RTFTEXT>'            : None,
                     '<SERIESNUM>'           : 'SeriesNum',
                     '<SNAPSHOTDURATION>'    : 'SnapshotDuration',
                     '<SNAPSHOTNUM>'         : 'SnapshotNum',
                     '<SNAPSHOTTIMECODE>'    : 'SnapshotTimeCode',
                     '<SORTORDER>'           : 'SortOrder',
                     '<SOURCE>'              : 'Source',
                     '<STARTCHAR>'           : 'StartChar',
                     '<SUBJECT>'             : 'Subject',
                     '<TITLE>'               : 'Title',
                     '<TRANSCRIBER>'         : 'Transcriber',
                     '<TRANSCRIPTNUM>'       : 'TranscriptNum',
                     '<TYPE>'                : 'Type',
                     '<VISIBLE>'             : 'Visible',
                     '<X1>'                  : 'X1',
                     '<X2>'                  : 'X2',
                     '<XMLTEXT>'             : 'XMLText',
                     '</XMLTEXT>'            : None,
                     '<Y1>'                  : 'Y1',
                     '<Y2>'                  : 'Y2' }

       # Start exception handling
       try:
           # Assume we're good to continue unless informed otherwise
           contin = True
           # Open the XML file
           f = file(self.XMLFile.GetValue(), 'r')

           # Initialize objectType and dataType, which are used to parse the file
           objectType = None 
           dataType = None
           # Initialize constants for whether the "Skip additional messages" checkbox should be displayed as part of error messages
           skipCheck = False
           skipValue = False

           # Some collections may have become children of LATER collections not yet imported.
           # Let's keep a list of instances of when this occurs so we can fix it after all the collections are read.
           collectionsToUpdate = []

           # For each line in the file ...
           for line in f:
               # ... increment the line counter
               lineCount += 1
               # If we DON'T have RTF Text or Note Text ...
               if not (dataType in  ['XMLText', 'RTFText', 'NoteText']):
                   # ... we can't just use strip() here.  Strings ending with the a with a grave (Alt-133) get corrupted!
                   line = Misc.unistrip(line)
               # If we have RTF Text or Note Text, we don't want to strip leading white space ...
               else:
                   # ... so we do unicode stripping only of the RIGHT side here.
                   line = Misc.unistrip(line, left=False)

               if DEBUG and not dataType in ['XMLText', 'RTFText']:
                   print "Line %d: '%s' %s %s" % (lineCount, line[:40], objectType, dataType)

               # Create an upper case version of line once (as repeated upper() calls were taking a LOT of time on profiling!)
               lineUpper = line.upper()

               # Begin line processing.
               # Generally, we figure out what kind of object we're importing (objectType) and
               # then figure out what object property we're importing (dataType) and then we
               # import the data.  But of course, there's a bit more to it than that.

               # For the sake of simple optimization, let's see if we're dealing with an XML command to start with

               if ((line != '') and not (dataType in ['XMLText', 'RTFText']) and (line[0] == '<')) or \
                  ((dataType == 'NoteText') and (lineUpper.lstrip() == '</NOTETEXT>')):

                   # Code for updating the Progress Bar
                   if lineUpper in MainHeads.keys():
                       progress.Update(MainHeads[lineUpper]['progPct'], MainHeads[lineUpper]['progPrompt'])
                       # These records should NEVER skip error messages
                       skipCheck = MainHeads[lineUpper]['skipCheck']
                       skipValue = MainHeads[lineUpper]['skipValue']

                   elif lineUpper.lstrip() in DataTypes.keys():
                       dataType = DataTypes[lineUpper.lstrip()]

                   # When we finish the Collections import section ...
                   elif lineUpper == '</COLLECTIONFILE>':
                       # ... iterate through the list of Collections that need to be updated because they are parented by collections
                       #     with a larger collection number ...
                       for col in collectionsToUpdate:
                           # ... Load the appropriate collection (after translating the collection number) ...
                           tmpColl = Collection.Collection(recNumbers['Collection'][col])
                           # ... lock the collection record ...
                           tmpColl.lock_record()
                           # ... Update the Parent Collection Number by translating the parent collection number ...
                           tmpColl.parent = recNumbers['Collection'][tmpColl.parent]

                           if DEBUG_Exceptions:
                               print "XMLImport 1:  Saving ", type(currentObj), objCountNumber
                               objCountNumber += 1;

                           # ... save the collection ...
                           tmpColl.db_save(use_transactions=False)
                           # ... and unlock the collection record.
                           tmpColl.unlock_record()

                   # Transana XML Version Checking
                   elif lineUpper == '<TRANSANAXMLVERSION>':
                       objectType = None
                       dataType = 'XMLVersionNumber'
                   elif lineUpper == '</TRANSANAXMLVERSION>':

                       # Version 1.0 -- Original Transana XML for Transana 2.0 release
                       # Version 1.1 -- Unicode encoding added to Transana XML for Transana 2.1 release
                       # Version 1.2 -- Filter Table added to Transana XML for Transana 2.11 release
                       # Version 1.3 -- FilterData handling changed to accomodate Unicode data for Transana 2.21 release
                       # Version 1.4 -- Database structure changed to accomodate Multiple Transcript Clips for Transana 2.30 release.
                       # Version 1.5 -- Database structure changed to accomodate Multiple Media Files for Transana 2.40
                       # Version 1.6 -- Added MinTranscriptWidth, XML format for transcripts, and character escapes for Transana 2.50 release
                       # Version 1.7 -- Added Snapshots, Snapshot Keywords, and Snapshot Coding Styles for Transana 2.60
                       # Version 1.8 -- Character Encoding Rules changed completely!!
                       # Version 2.0 -- Transana 3.0.  Documents and Quotes added

                       # Transana-XML version 1.0 ...
                       if self.XMLVersionNumber == '1.0':
                           # ... used Latin1 encoding
                           self.importEncoding = 'latin1'
                       # Transana-XML versions 1.1 through 1.4 ...
                       elif self.XMLVersionNumber in ['1.1', '1.2', '1.3', '1.4', '1.5']:
                           # ... use the encoding selected by the user
                           self.importEncoding = self.encodingOptions[self.chImportEncoding.GetSelection()]
                       # Transana-XML version 1.6 ...
                       elif self.XMLVersionNumber in ['1.6', '1.7', '1.8', '2.0']:
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
                   elif lineUpper == '<SERIES>':
                       currentObj = Library.Library()
                       objectType = 'Libraries'
                       dataType = None

                   elif lineUpper == '<DOCUMENT>':
                       currentObj = Document.Document()
                       objectType = 'Document'
                       dataType = None

                   elif lineUpper == '<EPISODE>':
                       currentObj = Episode.Episode()
                       objectType = 'Episode'
                       dataType = None

                   elif lineUpper == '<COREDATA>':
                       currentObj = CoreData.CoreData()
                       objectType = 'CoreData'
                       dataType = None

                   elif lineUpper == '<TRANSCRIPT>':
                       currentObj = Transcript.Transcript()
                       objectType = 'Transcript'
                       dataType = None

                   elif lineUpper == '<COLLECTION>':
                       currentObj = Collection.Collection()
                       objectType = 'Collection'
                       dataType = None

                   elif lineUpper == '<QUOTE>':
                       currentObj = Quote.Quote()
                       objectType = 'Quote'
                       dataType = None

                   elif lineUpper == '<QUOTEPOSITION>':
                       # There is not an Object for the Snapshot Keywords table.  We have to create the data record by hand.
                       currentObj = None
                       self.snapshotKeyword = { 'QuoteNum'     :  0,
                                                'DocumentNum'  :  0,
                                                'StartChar'    :  -1,
                                                'EndChar'      :  -1 }
                       objectType = 'QuotePosition'
                       dataType = None

                   elif lineUpper == '<CLIP>':
                       currentObj = Clip.Clip()
                       objectType = 'Clip'
                       dataType = None

                   elif lineUpper == '<SNAPSHOT>':
                       currentObj = Snapshot.Snapshot()
                       objectType = 'Snapshot'
                       dataType = None

                   elif lineUpper == '<SNAPSHOTKEYWORD>':
                       # There is not an Object for the Snapshot Keywords table.  We have to create the data record by hand.
                       currentObj = None
                       self.snapshotKeyword = { 'SnapshotNum'  :  0,
                                                'KeywordGroup' :  '',
                                                'Keyword'      :  '',
                                                'X1'           :  0,
                                                'Y1'           :  0,
                                                'X2'           :  0,
                                                'Y2'           :  0,
                                                'Visible'      :  False }
                       objectType = 'SnapshotKeyword'
                       dataType = None

                   elif lineUpper == '<SNAPSHOTKEYWORDSTYLE>':
                       # There is not an Object for the Snapshot Keyword Styles table.  We have to create the data record by hand.
                       currentObj = None
                       self.snapshotKeywordStyle = { 'SnapshotNum'  :  0,
                                                     'KeywordGroup' :  '',
                                                     'Keyword'      :  '',
                                                     'DrawMode'     :  '',
                                                     'ColorName'    :  '',
                                                     'ColorDef'     :  '',
                                                     'LineWidth'    :  0,
                                                     'LineStyle'    :  '' }
                       objectType = 'SnapshotKeywordStyle'
                       dataType = None

                   elif lineUpper == '<ADDVID>':
                       # Additional Video Files don't exactly have their own object type.  They're part of Episode and Clip
                       # records.  So we'll just use a Dictionary object for their data import.
                       currentObj = {}
                       objectType = 'AddVid'
                       dataType = None

                   elif lineUpper == '<KEYWORDREC>':
                       currentObj = Keyword.Keyword()
                       objectType = 'Keyword'
                       dataType = None

                   elif lineUpper == '<CLIPKEYWORD>':
                       currentObj = ClipKeywordObject.ClipKeyword('', '')
                       objectType = 'ClipKeyword'
                       dataType = None

                   elif lineUpper == '<NOTE>':
                       currentObj = Note.Note()
                       objectType = 'Note'
                       dataType = None

                   elif lineUpper == '<FILTER>':
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
                   elif lineUpper == '</SERIES>' or \
                        lineUpper == '</DOCUMENT>' or \
                        lineUpper == '</EPISODE>' or \
                        lineUpper == '</COREDATA>' or \
                        lineUpper == '</TRANSCRIPT>' or \
                        lineUpper == '</COLLECTION>' or \
                        lineUpper == '</QUOTE>' or \
                        lineUpper == '</QUOTEPOSITION>' or \
                        lineUpper == '</CLIP>' or \
                        lineUpper == '</SNAPSHOT>' or \
                        lineUpper == '</SNAPSHOTKEYWORD>' or \
                        lineUpper == '</SNAPSHOTKEYWORDSTYLE>' or \
                        lineUpper == '</ADDVID>' or \
                        lineUpper == '</KEYWORDREC>' or \
                        lineUpper == '</CLIPKEYWORD>' or \
                        lineUpper == '</NOTE>' or \
                        lineUpper == '</FILTER>':
                       dataType = None

                       # Saves are one area where problems will arise if the data's not clean.
                       # We can trap some of these problems here.
                       try:
                           # We can't just keep the numbers that were assigned in the exporting database.
                           # Objects use the presence of a number to update rather than insert, and we need to
                           # insert here.  Therefore, we'll strip the record number here, but remember it for
                           # user later.
                           if objectType == 'Libraries' or \
                              objectType == 'Document' or \
                              objectType == 'Episode' or \
                              objectType == 'Transcript' or \
                              objectType == 'Collection' or \
                              objectType == 'Quote' or \
                              objectType == 'Clip' or \
                              objectType == 'Snapshot' or \
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
                                   tmpdlg = Dialogs.InfoDialog(self, currentObj.__repr__())
                                   tmpdlg.ShowModal()
                                   tmpdlg.Destroy()
                               elif DEBUG and (objectType == 'Transcript'):
                                   print currentObj
                                   print

                               if DEBUG_Exceptions:
                                   print "XMLImport 2:  Saving ", type(currentObj), objCountNumber
                                   objCountNumber += 1;

                               if objectType == 'Document':
                                   # Save the data object
                                   currentObj.db_save(use_transactions=False, ignore_filename=True)
                               else:
                                   # Save the data object
                                   currentObj.db_save(use_transactions=False)
                               # Let's keep a record of the old and new object numbers for each object saved.
                               recNumbers[objectType][oldNumber] = currentObj.number

                               # If we've just saved a Clip ...
                               if objectType == 'Clip':
                                   # ... we need to save the reverse lookup data, so we can find the clip's OLD
                                   # number based on it's new number.  This must be post-save, as that's when
                                   # the new object number gets assigned.
                                   recNumbers['OldClip'][currentObj.number] = oldNumber

                           elif objectType == 'QuotePosition':
                               # Define the query to update the Quote Position data in the database.
                               # (This should be faster than loading the Quote, adding the position values, and saving the Quote.
                               query = "UPDATE QuotePositions2 SET DocumentNum = %s, StartChar = %s, EndChar = %s WHERE QuoteNum = %s"
                               query = DBInterface.FixQuery(query)
                               # Get the data for each insert query
                               data = (quotePosition['DocumentNum'], quotePosition['StartChar'], quotePosition['EndChar'], quotePosition['QuoteNum'])
                               # Execute the query
                               dbCursor.execute(query, data)

                           elif objectType == 'AddVid':
                               # Additional Video records don't have a proper object type, so we have to do the saves the hard way.
                               # Make sure all necessary elements are present
                               if not currentObj.has_key('EpisodeNum'):
                                   currentObj['EpisodeNum'] = 0
                               if not currentObj.has_key('ClipNum'):
                                   currentObj['ClipNum'] = 0
                               if not currentObj.has_key('VidLength'):
                                   currentObj['VidLength'] = 0
                               if not currentObj.has_key('Offset'):
                                   currentObj['Offset'] = 0
                               if not currentObj.has_key('Audio'):
                                   currentObj['Audio'] = 0

                               # Define the query to insert the additional media files into the database
                               query = "INSERT INTO AdditionalVids2 (EpisodeNum, ClipNum, MediaFile, VidLength, Offset, Audio) VALUES (%s, %s, %s, %s, %s, %s)"
                               query = DBInterface.FixQuery(query)
                               # Substitute the generic OS seperator "/" for the Windows "\".
                               tmpFilename = currentObj['MediaFile'].replace('\\', '/')
                               # Encode the filename
                               tmpFilename = tmpFilename.encode(TransanaGlobal.encoding)
                               # Get the data for each insert query
                               data = (currentObj['EpisodeNum'], currentObj['ClipNum'], tmpFilename, currentObj['VidLength'], currentObj['Offset'], currentObj['Audio'])
                               # Execute the query
                               dbCursor.execute(query, data)

                           elif  objectType == 'CoreData':
                               currentObj.number = 0

                               if DEBUG_Exceptions:
                                   print "XMLImport 3:  Saving ", type(currentObj), objCountNumber
                                   objCountNumber += 1;

                               currentObj.db_save()

                           elif objectType == 'Keyword':

                               if DEBUG_Exceptions:
                                   print "XMLImport 4:  Saving ", type(currentObj), objCountNumber
                                   objCountNumber += 1;

                               currentObj.db_save(use_transactions=False)
                               
                           elif objectType == 'ClipKeyword':
                               if (currentObj.documentNum != 0) or \
                                  (currentObj.episodeNum != 0) or \
                                  (currentObj.quoteNum != 0) or \
                                  (currentObj.clipNum != 0) or \
                                  (currentObj.snapshotNum != 0):

                                   if DEBUG_Exceptions:
                                       print "XMLImport 5:  Saving ", type(currentObj), objCountNumber
                                       objCountNumber += 1;

                                   currentObj.db_save()

                           elif objectType == 'SnapshotKeyword':
                               if self.snapshotKeyword['SnapshotNum'] > 0:
                                   # Create the query to save the Snapshot Keyword record    
                                   query = """ INSERT INTO SnapshotKeywords2
                                                 (SnapshotNum, KeywordGroup, Keyword, x1, y1, x2, y2, visible)
                                               VALUES
                                                 (%s, %s, %s, %s, %s, %s, %s, %s) """
                                   query = DBInterface.FixQuery(query)
                                   # Build the values to match the query
                                   values = (self.snapshotKeyword['SnapshotNum'],
                                             self.snapshotKeyword['KeywordGroup'].encode('utf8'),
                                             self.snapshotKeyword['Keyword'].encode('utf8'),
                                             self.snapshotKeyword['X1'],
                                             self.snapshotKeyword['Y1'],
                                             self.snapshotKeyword['X2'],
                                             self.snapshotKeyword['Y2'],
                                             self.snapshotKeyword['Visible'])
                                   # Save the Snapshot Keyword data
                                   if db != None:
                                       dbCursor.execute(query, values)
                                   
                           elif objectType == 'SnapshotKeywordStyle':
                               if self.snapshotKeywordStyle['SnapshotNum'] > 0:
                                   # Create the query to save the Snapshot Keyword Style record    
                                   query = """ INSERT INTO SnapshotKeywordStyles2
                                                 (SnapshotNum, KeywordGroup, Keyword, DrawMode, LineColorName, LineColorDef, LineWidth, LineStyle)
                                               VALUES
                                                 (%s, %s, %s, %s, %s, %s, %s, %s) """
                                   query = DBInterface.FixQuery(query)
                                   # Build the values to match the query
                                   values = (self.snapshotKeywordStyle['SnapshotNum'],
                                             self.snapshotKeywordStyle['KeywordGroup'].encode('utf8'),
                                             self.snapshotKeywordStyle['Keyword'].encode('utf8'),
                                             self.snapshotKeywordStyle['DrawMode'],
                                             self.snapshotKeywordStyle['ColorName'].encode('utf8'),
                                             self.snapshotKeywordStyle['ColorDef'],
                                             self.snapshotKeywordStyle['LineWidth'],
                                             self.snapshotKeywordStyle['LineStyle'])
                                   # Save the Snapshot Keyword Style data
                                   if db != None:
                                       dbCursor.execute(query, values)
                                   
                           elif objectType == 'Filter':
                               # Starting with XML Version 1.3, we have to deal with encoding issues for the Filter data
                               if not self.XMLVersionNumber in ['1.0', '1.1', '1.2']:
                                   # With XML version 2.0, additional filter data types need to be unpickled
                                   if (self.FilterFilterDataType in ['1', '2', '3', '5', '6', '7']) or \
                                      ((self.XMLVersionNumber in ['2.0']) and (self.FilterFilterDataType in ['8', '18', '19', '20'])):
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

                                   # All other data EXCEPT SAVED SEARCHES is encoded and needs to be decoded, but was not pickled.
                                   elif self.FilterReportType != '15':
                                       self.FilterFilterData = DBInterface.ProcessDBDataForUTF8Encoding(self.FilterFilterData)

                                   # Saved Searches (Added for 2.50)
                                   elif self.FilterReportType == '15' and self.FilterScope == '0':
                                       # If the Filter Data is a string (it always should be!) ...
                                       if isinstance(self.FilterFilterData, str):
                                           # ... then decode it using the import encoding.
                                           self.FilterFilterData = self.FilterFilterData.decode(self.importEncoding)
                                       # Now encode the filter data using the file encoding
                                       self.FilterFilterData = self.FilterFilterData.encode(TransanaGlobal.encoding)
                               # Encode the Filter Configuration Name using the file encoding
                               self.FilterConfigName = self.FilterConfigName.encode(TransanaGlobal.encoding)

                               if DEBUG:
                                   print "Filter Rec:", self.FilterReportType, self.FilterScope, self.FilterFilterDataType

                               # Certain FilterDataTypes need to have their DATA adjusted for the new object numbers!
                               # This should be done before the save.
                               # So if we have Clips or Notes Filter Data ...
                               if (self.FilterFilterDataType in ['2', '8', '20']) or \
                                  ((self.FilterReportType == '15') and (self.FilterScope == 1)):
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
                                       # If we have a Save Collections record ...
                                       elif self.FilterReportType == '15':
                                           # ... and if the Collection Number still exists in the new data set ...
                                           if recNumbers['Collection'].has_key(dataRec[0]):
                                               # ... get the new Collection Number ...
                                               collNum = recNumbers['Collection'][dataRec[0]]
                                               # ... and substitute it into the data record
                                               filterData.append((collNum, dataRec[1]))
                                       # If we have a Save Quotes record ...
                                       elif self.FilterFilterDataType == '20':
                                           # ... and if the Colllection Number still exists in the new data set ...
                                           if recNumbers['Collection'].has_key(dataRec[1]):
                                               # ... get the new Collection Number ...
                                               collNum = recNumbers['Collection'][dataRec[1]]
                                               # ... and substitute it into the data record
                                               filterData.append((dataRec[0], collNum, dataRec[2]))
                                   # Now re-pickle the filter data
                                   self.FilterFilterData = cPickle.dumps(filterData)
                               # Create the query to save the Filter record    
                               query = """ INSERT INTO Filters2
                                               (ReportType, ReportScope, ConfigName, FilterDataType, FilterData)
                                             VALUES
                                               (%s, %s, %s, %s, %s) """
                               query = DBInterface.FixQuery(query)

                               if DEBUG:
                                   print 'XMLImport: Saving Filter Record:', self.FilterReportType, self.FilterScope, \
                                         self.FilterConfigName, self.FilterFilterDataType, type(self.FilterFilterData)
                               
                               # Build the values to match the query, including the pickled Clip data
                               values = (self.FilterReportType, self.FilterScope, self.FilterConfigName, self.FilterFilterDataType,
                                         self.FilterFilterData)
                               # Save the Filter data
                               if db != None:
                                   dbCursor.execute(query, values)
                       except:

                           if DEBUG or DEBUG_Exceptions:
                               print
                               print sys.exc_info()[0], sys.exc_info()[1]
                               print
                               if (objectType == 'Transcript'):
                                   tmpdlg = Dialogs.InfoDialog(self, currentObj.__repr__())
                                   tmpdlg.ShowModal()
                                   tmpdlg.Destroy()
                               import traceback
                               traceback.print_exc(file=sys.stdout)
                               print
                               print

                           # If we haven't been told to skip error messages of this type ...
                           if not skipValue:
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
                               if objectType == 'CoreData':
                                   msg = msg + '.'
                                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                   prompt = unicode(_('The Core Data record is for media file "%s".'), 'utf8')
                                   # Explain about Keyword Definitions
                                   prompt += '\n\n' + unicode(_('The existing record will not be updated, but the Database import will continue.'), 'utf8')
                                   msg = msg + '\n\n' + prompt % (currentObj.id)
                                   # So the keyword already exists.  Let's continue the import anyway!  This is a minor issue.
                                   contin = True

                               elif objectType == 'QuotePosition':
                                   prompt = unicode('for Quote %s.', 'utf8')
                                   msg = msg +  ' ' + prompt % quotePosition['QuoteNum']
                               elif (not objectType in ['AddVid', 'Keyword', 'ClipKeyword', 'Filter']) and (currentObj.id != ''):
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

                                       msg = msg + u'\n\n' + unicode(_('The Transcript is for'), 'utf8')
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
                                            prompt = unicode(_('The record is for Keyword "%s:%s".') + '  ', 'utf8')
                                            # Explain about Keyword Definitions
                                            prompt += '\n\n' + unicode(_('The Keyword Definition will not be updated, but the Database import will continue.'), 'utf8')
                                        else:
                                            prompt = _('The record is for Keyword "%s:%s".') + '  '
                                            # Explain about Keyword Definitions
                                            prompt += '\n\n' + unicode(_('The Keyword Definition will not be updated, but the Database import will continue.'), 'utf8')
                                        msg = msg + '\n\n' + prompt % (currentObj.keywordGroup, currentObj.keyword)
                                        # So the keyword already exists.  Let's continue the import anyway!  This is a minor issue.
                                        contin = True
                               # If we're interrupting and cancelling the import ...
                               if not contin:
                                   # ... we need to tell the user where to intervene.
                                   if 'unicode' in wx.PlatformInfo:
                                       # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                       prompt = unicode(_('You need to correct this record in XML file %s.'), 'utf8')
                                       prompt2 = unicode(_('The %s record ends at line %d.'), 'utf8')
                                   else:
                                       prompt = _('You need to correct this record in XML file %s.')
                                       prompt2 = _('The %s record ends at line %d.')
                                   # Add the intervention information to the error message
                                   msg = msg + '\n' +  prompt % self.XMLFile.GetValue() + '\n' + \
                                                       prompt2 % (objectType, lineCount)
                               # Display our carefully crafted error message to the user.
                               errordlg = Dialogs.ErrorDialog(None, msg, includeSkipCheck=skipCheck)
                               errordlg.ShowModal()
                               # if skipping error messages is an option ...
                               if skipCheck:
                                   # ... see if the Skip Error Messages checkbox has been checked
                                   skipValue = errordlg.GetSkipCheck()
                               errordlg.Destroy()

                       currentObj = None
                       objectType = None

               else:

                   # Also for the sake of minimalist optimization, let's deal with the RTFText and XMLTest 
                   # datatypes first, because we spend a LOT of time here in most imports.  Let's find 
                   # this with fewer preliminary "if" checks!

                   # Because XML Text can be many lines long, we need to explicitly close this datatype when
                   # the closing XML tag is found.  Since left stripping is skipped during XMLText reads, we need
                   # to add the lstrip() call here.
                   if lineUpper.lstrip() in ['</XMLTEXT>', '</RTFTEXT>']:
                       dataType = None

                   elif dataType in ['XMLText', 'RTFText']:
                       # Add Line Breaks to the text to match the incoming lines.
                       # Otherwise, the transcript might be messed up, with the first word of the next line
                       # being truncated.

                       # If this is the FIRST LINE ...
                       if currentObj.text == '':
                           # If we have an XML richtext specification ...
                           if line[:10] == '<richtext ':
                               # ... add the XML Header, which was stripped out during export because it breaks XML
                               currentObj.text = '<?xml version="1.0" encoding="UTF-8"?>\n'
                       else:
                           currentObj.text = currentObj.text + '\n'

                       currentObj.text = currentObj.text + line
                       # We DO NOT reset DataType here, as RTFText may be many lines long!
                       # dataType = None

                   # ignore blank lines, except in the Notes Text, where they represent blank lines!
                   elif (line == '') and (dataType != 'NoteText'):
                       pass

                   # Code for populating Object Properties.
                   # Unless data can stretch across mulitple lines, we should explicity undefine the dataType
                   # once the data is captured.
                   elif dataType == 'Num':
                       if not objectType in ['QuotePosition', 'AddVid']:
                           currentObj.number = int(line)
                       elif objectType == 'QuotePosition':
                           quotePosition['QuoteNum'] = recNumbers['Quote'][int(line)]
                       elif objectType == 'AddVid':
                           currentObj['AddVidNum'] = int(line)
                       dataType = None

                   elif dataType == 'ID':
                       currentObj.id = self.ProcessLine(line)
                       dataType = None

                       if objectType == 'Document':
                           st = _('Importing Document records (This may be slow because of the size of Document records.)')
                           st += '\n  '
                           st += _("Document")
                           st += ' '
                           st += currentObj.id.encode(TransanaGlobal.encoding)
                           st += '  (%d)' % currentObj.number
                           progress.Update(5, st)

                       if objectType == 'Transcript':
                           st = _('Importing Transcript records (This may be slow because of the size of Transcript records.)')
                           st += '\n  '
                           st += _("Transcript")
                           st += ' '
                           st += currentObj.id.encode(TransanaGlobal.encoding)
                           st += '  (%d)' % currentObj.number
                           progress.Update(43, st)

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
                       if objectType in ['Document']:
                           # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                           try:
                               currentObj.library_num = recNumbers['Libraries'][int(line)]
                           except:
                               pass
                       else:
                           # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                           try:
                               currentObj.series_num = recNumbers['Libraries'][int(line)]
                           except:
                               pass
                       dataType = None

                   elif dataType == 'DocumentNum':
                       if objectType in ['Quote']:
                           # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                           try:
                               currentObj.source_document_num = recNumbers['Document'][int(line)]
                           except:
                               pass
                       elif objectType == 'QuotePosition':
                           quotePosition['DocumentNum'] = recNumbers['Document'][int(line)]
                       elif objectType == 'ClipKeyword':
                           # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                           try:
                               currentObj.documentNum = recNumbers['Document'][int(line)]
                           except:
                               pass
                       else:
                           # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                           try:
                               currentObj.document_num = recNumbers['Document'][int(line)]
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
                                   if objectType == 'ClipKeyword':
                                       currentObj.episodeNum = recNumbers['Episode'][int(line)]
                                   elif objectType != 'AddVid':
                                       currentObj.episode_num = recNumbers['Episode'][int(line)]
                                   else:
                                       currentObj['EpisodeNum'] = recNumbers['Episode'][int(line)]
                               except:
                                   pass
                           else:
                               # If the old record number doesn't exist, substitute 0 and show an error message.
                               if objectType != 'AddVid':
                                   currentObj.episode_num = 0
                               if isinstance(currentObj, ClipKeywordObject.ClipKeyword):
                                   if 'unicode' in wx.PlatformInfo:
                                       # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                       prompt = unicode(_('Episode Number %s cannot be found for %s at line number %d.'), 'utf8')
                                   else:
                                       prompt = _('Episode Number %s cannot be found for %s at line number %d.')
                                   vals = (line, objectType, lineCount)
                               elif objectType == 'AddVid':
                                   if 'unicode' in wx.PlatformInfo:
                                       # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                       prompt = unicode(_('Episode Number %s cannot be found for %s record %d at line number %d.'), 'utf8')
                                   else:
                                       prompt = _('Episode Number %s cannot be found for %s record %d at line number %d.')
                                   vals = (line, _("Additional Video"), currentObj['AddVidNum'], lineCount)
                               else:
                                   if 'unicode' in wx.PlatformInfo:
                                       # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                       prompt = unicode(_('Episode Number %s cannot be found for %s record %d at line number %d.'), 'utf8')
                                   else:
                                       prompt = _('Episode Number %s cannot be found for %s record %d at line number %d.')
                                   vals = (line, objectType, currentObj.number, lineCount)
                               # This should be INFORMATION rather than ERROR!
                               errordlg = Dialogs.InfoDialog(None, prompt % vals)
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

                   elif dataType == 'QuoteNum':
                       if objectType == 'ClipKeyword':
                           # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                           try:
                               currentObj.quoteNum = recNumbers['Quote'][int(line)]
                           except:
                               pass
                       else:
                           # "line" may not be an integer, but could be "None".  Trap this and skip it if so.
                           try:
                               currentObj.quote_num = recNumbers['Quote'][int(line)]
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
                               # This should be INFORMATION rather than ERROR!
                               errordlg = Dialogs.InfoDialog(None, prompt  % (line, objectType, lineCount))
                               errordlg.ShowModal()
                               errordlg.Destroy()
                       elif objectType == 'AddVid':
                           currentObj['ClipNum'] = recNumbers['Clip'][int(line)]
                       else:
                           currentObj.clipNum = recNumbers['Clip'][int(line)]
                                
                       dataType = None

                       if objectType == 'Transcript':
                           progress.Update(43, _('Importing Transcript records (This may be slow because of the size of Transcript records.)') + \
                                               '\n  ' + _("Clip Transcript") + ' %d' % currentObj.clip_num)

                   elif dataType == 'SnapshotNum':
                       if objectType == 'SnapshotKeyword':
                           self.snapshotKeyword['SnapshotNum'] = recNumbers['Snapshot'][int(line)]
                       elif objectType == 'SnapshotKeywordStyle':
                           self.snapshotKeywordStyle['SnapshotNum'] = recNumbers['Snapshot'][int(line)]
                       elif objectType == 'ClipKeyword':
                           try:
                               currentObj.snapshotNum = recNumbers['Snapshot'][int(line)]
                           except:
                               if 'unicode' in wx.PlatformInfo:
                                   # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                   prompt = unicode(_('Snapshot Number %s cannot be found for a %s record at line number %d.\n(This is due to an incomplete Snapshot deletion and is not a problem.)'), 'utf8')
                               else:
                                   prompt = _('Snapshot Number %s cannot be found for a %s record at line number %d.\n(This is due to an incomplete Snapshot deletion and is not a problem.)')
                               # This should be INFORMATION rather than ERROR!
                               errordlg = Dialogs.InfoDialog(None, prompt  % (line, objectType, lineCount))
                               errordlg.ShowModal()
                               errordlg.Destroy()
                       else:
                           currentObj.snapshot_num = recNumbers['Snapshot'][int(line)]
                       dataType = None

                   elif dataType == 'date':
                       # Importing dates can be a little tricky.  Let's trap conversion errors
                       try:
                           if objectType in ['Document']:
                               currentObj.import_date = self.ProcessLine(line)
                           # If we're dealing with an Episode record ...
                           if objectType == 'Episode':
                               # Make sure it's in a form we recognize
                               # See if the date is in YYYY-MM-DD format (produced by XMLExport.py).
                               # If not, substitute slashes for dashes, as this file likely came from Delphi!
                               reStr = '\d{4}-\d+-\d+'
                               reStr2 = '\d{4}/\d+/\d+'
                               if re.compile(reStr).match(line) != None:
                                   timeformat = "%Y-%m-%d"
                               # The date should be stored in the Episode record as a date object
                               elif re.compile(reStr2).match(line) != None:
                                   timeformat = "%Y/%m/%d"
                               # The date should be stored in the Episode record as a date object
                               else:
                                   line = line.replace('-', '/')
                                   timeformat = "%m/%d/%Y"
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
                       if objectType in ['Document']:
                           currentObj.imported_file = self.ProcessLine(line)
                       elif objectType == 'AddVid':
                           currentObj['MediaFile'] = self.ProcessLine(line)
                       elif objectType == 'Snapshot':
                           currentObj.image_filename = self.ProcessLine(line)
                       else:
                           currentObj.media_filename = self.ProcessLine(line)
                       dataType = None

                   elif dataType == 'Length':
                       if objectType == 'Document':
                           currentObj.document_length = line
                       elif objectType != 'AddVid':
                           currentObj.tape_length = line
                       else:
                           currentObj['VidLength'] = line
                       dataType = None

                   elif dataType == 'Offset':
                       if objectType != 'AddVid':
                           currentObj.offset = line
                       else:
                           currentObj['Offset'] = line
                       dataType = None

                   elif dataType == 'Audio':
                       if objectType != 'AddVid':
                           currentObj.audio = line
                       else:
                           currentObj['Audio'] = line
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

                   elif dataType == 'ParentCollectNum':
                       # If the parent collection has already been defined, and thus has a known record number ...
                       if int(line) in recNumbers['Collection']:
                           # ... save the updated parent collection number
                           currentObj.parent = recNumbers['Collection'][int(line)]
                       # If the parent collection number is not yet knows, because the parent collection hasn't been processed yet ...
                       else:
                           # ... add this collection (by number) to the list of collections that need to be updated later ...
                           collectionsToUpdate.append(currentObj.number)
                           # ... and store the UNTRANSLATED parent collection number data to be translated later.
                           currentObj.parent = int(line)
                           
                       dataType = None

                   elif dataType == 'StartChar':
                       if objectType == 'QuotePosition':
                           quotePosition['StartChar'] = int(line)

                   elif dataType == 'EndChar':
                       if objectType == 'QuotePosition':
                           quotePosition['EndChar'] = int(line)

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

                   elif dataType == 'MinTranscriptWidth':
                       currentObj.minTranscriptWidth = int(line)
                       dataType = None

                   elif dataType == 'SortOrder':
                       currentObj.sort_order = line
                       dataType = None

                   elif dataType == 'ImageScale':
                       currentObj.image_scale = float(line)
                       dataType = None

                   elif dataType == 'ImageCoordsX':
                       if len(currentObj.image_coords) == 2:
                           currentObj.image_coords = (float(line), currentObj.image_coords[1])
                       else:
                           currentObj.image_coords = (float(line), 0.0)
                       dataType = None

                   elif dataType == 'ImageCoordsY':
                       if len(currentObj.image_coords) == 2:
                           currentObj.image_coords = (currentObj.image_coords[0], float(line))
                       else:
                           currentObj.image_coords = (0.0, float(line))
                       dataType = None

                   elif dataType == 'ImageSizeW':
                       if len(currentObj.image_size) == 2:
                           currentObj.image_size = (int(line), currentObj.image_size[1])
                       else:
                           currentObj.image_size = (int(line), 0)
                       dataType = None

                   elif dataType == 'ImageSizeH':
                       if len(currentObj.image_size) == 2:
                           currentObj.image_size = (currentObj.image_size[0], int(line))
                       else:
                           currentObj.image_size = (0, int(line))
                       dataType = None

                   elif dataType == 'SnapshotTimeCode':
                       currentObj.episode_start = int(line)
                       dataType = None

                   elif dataType == 'SnapshotDuration':
                       currentObj.episode_duration = int(line)
                       dataType = None

                   elif dataType == 'X1':
                       self.snapshotKeyword['X1'] = round(float(line))
                       dataType = None

                   elif dataType == 'Y1':
                       self.snapshotKeyword['Y1'] = round(float(line))
                       dataType = None

                   elif dataType == 'X2':
                       self.snapshotKeyword['X2'] = round(float(line))
                       dataType = None

                   elif dataType == 'Y2':
                       self.snapshotKeyword['Y2'] = round(float(line))
                       dataType = None

                   elif dataType == 'Visible':
                       self.snapshotKeyword['Visible'] = int(line)
                       dataType = None

                   elif dataType == 'DrawMode':
                       if objectType == 'Keyword':
                           currentObj.drawMode = self.ProcessLine(line)
                       elif objectType == 'SnapshotKeywordStyle':
                           self.snapshotKeywordStyle['DrawMode'] = self.ProcessLine(line)
                       dataType = None

                   elif dataType == 'ColorName':
                       if objectType == 'Keyword':
                           currentObj.lineColorName = self.ProcessLine(line)
                       elif objectType == 'SnapshotKeywordStyle':
                           self.snapshotKeywordStyle['ColorName'] = self.ProcessLine(line)
                       dataType = None

                   elif dataType == 'ColorDef':
                       if objectType == 'Keyword':
                           currentObj.lineColorDef = self.ProcessLine(line)
                       elif objectType == 'SnapshotKeywordStyle':
                           self.snapshotKeywordStyle['ColorDef'] = self.ProcessLine(line)
                       dataType = None

                   elif dataType == 'LineWidth':
                       if objectType == 'Keyword':
                           currentObj.lineWidth = int(line)
                       elif objectType == 'SnapshotKeywordStyle':
                           self.snapshotKeywordStyle['LineWidth'] = int(line)
                       dataType = None

                   elif dataType == 'LineStyle':
                       if objectType == 'Keyword':
                           currentObj.lineStyle = self.ProcessLine(line)
                       elif objectType == 'SnapshotKeywordStyle':
                           self.snapshotKeywordStyle['LineStyle'] = self.ProcessLine(line)
                       dataType = None

                   elif dataType == 'KWG':
                       if objectType == 'SnapshotKeyword':
                           self.snapshotKeyword['KeywordGroup'] = self.ProcessLine(line)
                       elif objectType == 'SnapshotKeywordStyle':
                           self.snapshotKeywordStyle['KeywordGroup'] = self.ProcessLine(line)
                       else:
                           currentObj.keywordGroup = self.ProcessLine(line)
                       dataType = None

                   elif dataType == 'KW':
                       if objectType == 'SnapshotKeyword':
                           self.snapshotKeyword['Keyword'] = self.ProcessLine(line)
                       elif objectType == 'SnapshotKeywordStyle':
                           self.snapshotKeywordStyle['Keyword'] = self.ProcessLine(line)
                       else:
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
                           # We changed the encoding here from importEncoding to UTF-8 no matter what for version 2.50.
                           currentObj.definition = currentObj.definition + line.decode('utf8')  # (self.importEncoding)
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
                       # NOTE:  we always use UTF8 here, not self.importEncoding!
                       currentObj.text = currentObj.text + line.decode('utf8')
                       # We DO NOT reset DataType here, as NoteText may be many lines long!
                       # dataType = None

                   elif dataType == 'ReportType':
                        self.FilterReportType = line
                        dataType = None

                   elif dataType == 'ReportScope':
                        if self.FilterReportType in ['5', '6', '7', '10', '14']:
                            self.FilterScope = recNumbers['Libraries'][int(line)]
                        elif self.FilterReportType in ['17', '18', '19', '20']:
                            self.FilterScope = recNumbers['Document'][int(line)]
                        elif self.FilterReportType in ['1', '2', '3', '8', '11']:
                            self.FilterScope = recNumbers['Episode'][int(line)]
                        # Collection Clip Data Export (ReportType 4) only needs translation if ReportScope != 0
                        elif (self.FilterReportType in ['12', '16']) or ((self.FilterReportType == '4') and (int(line) != 0)):
                            self.FilterScope = recNumbers['Collection'][int(line)]
                        elif self.FilterReportType in ['13', '15']:
                            # FilterScopes for ReportType 13 (Notes Report) are constants, not object numbers!
                            self.FilterScope = int(line)
                        # Collection Clip Data Export (ReportType 4) for ReportScope 0, the Collection Root, needs
                        # a FilterScope of 0
                        # Saved Search (ReportType 15) needs no modifications, but setting FilterScope to 0 allows the SAVE!
                        elif ((self.FilterReportType == '4') and (int(line) == 0)):
                            self.FilterScope = 0
                        else:
                           if 'unicode' in wx.PlatformInfo:
                               # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                               prompt = unicode(_('An error occurred during Database Import.\nThere is an unsupported Filter Report Type (%s) in the Filter table. \nYou may wish to upgrade Transana and try again.'), 'utf8')
                           else:
                               prompt = _('An error occurred during Database Import.\nThere is an unsupported Filter Report Type (%s) in the Filter table. \nYou may wish to upgrade Transana and try again.')
                           errordlg = Dialogs.ErrorDialog(self, prompt % (self.FilterReportType))
                           errordlg.ShowModal()
                           errordlg.Destroy()
                        dataType = None

                   elif dataType == 'ConfigName':
                        # Struggling to get the encoding correct.  ProcessLine(line) appears to work even in Chinese.
                        self.FilterConfigName = self.ProcessLine(line)  # line.decode('utf8')  # DBInterface.ProcessDBDataForUTF8Encoding(line)
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
               progress.Update(81, _('Updating Source Transcript Numbers in Clip Transcript records'))
               if db != None:
                   dbCursor2 = db.cursor()
                   # Get all NEW transcript records.  We DON'T want to process transcript records that were in the database prior
                   # to the import, as they won't be in recNumbers and thus we'd lose all SourceTranscript records!
                   SQLText = 'SELECT TranscriptNum, SourceTranscriptNum, ClipNum FROM Transcripts2 WHERE ClipNum > 0 AND TranscriptNum > %s'
                   SQLText = DBInterface.FixQuery(SQLText)
                   dbCursor.execute(SQLText, (transcriptCount, ))
                   # create the SQL for updating the SourceTranscriptNum of all new transcripts
                   SQLText = """ UPDATE Transcripts2
                                 SET SourceTranscriptNum = %s
                                 WHERE TranscriptNum = %s """
                   SQLText = DBInterface.FixQuery(SQLText)
                   # For each new Transcript record ...
                   for (TranscriptNum, SourceTranscriptNum, ClipNum) in dbCursor.fetchall():
                       # It is possible that the originating Transcript has been deleted.  If so,
                       # we need to set the TranscriptNum to 0.  We accomplish this by adding the
                       # missing Transcript Number to our recNumbers list with a value of 0
                       if not(SourceTranscriptNum in recNumbers['Transcript'].keys()):
                           recNumbers['Transcript'][SourceTranscriptNum] = 0
                       dbCursor2.execute(SQLText, (recNumbers['Transcript'][SourceTranscriptNum], TranscriptNum))

                   dbCursor2.close()

               if db != None:

                   # Hyperlinks in Documents

                   # We need a secong database cursor for updates
                   dbCursor2 = db.cursor()
                   
                   progress.Update(86, _('Updating HyperLinks in Documents'))
                   # Get all Document records
                   SQLText = 'SELECT DocumentNum, XMLText FROM Documents2'
                   SQLText = DBInterface.FixQuery(SQLText)
                   dbCursor.execute(SQLText)
                   # create the SQL for updating the XMLText of the Document
                   SQLText = """ UPDATE Documents2
                                 SET XMLText = %s
                                 WHERE DocumentNum = %s """
                   SQLText = DBInterface.FixQuery(SQLText)
                   # For each Document record ...
                   for (DocumentNum, XMLText) in dbCursor.fetchall():
                       dbCursor2.execute(SQLText, (self.UpdateHyperlinks(XMLText, recNumbers), DocumentNum))

                   progress.Update(90, _('Updating HyperLinks in Quotes'))
                   # Get all Quote records
                   SQLText = 'SELECT QuoteNum, XMLText FROM Quotes2'
                   SQLText = DBInterface.FixQuery(SQLText)
                   dbCursor.execute(SQLText)
                   # create the SQL for updating the XMLText of the Quote
                   SQLText = """ UPDATE Quotes2
                                 SET XMLText = %s
                                 WHERE QuoteNum = %s """
                   SQLText = DBInterface.FixQuery(SQLText)
                   # For each Quote record ...
                   for (QuoteNum, XMLText) in dbCursor.fetchall():
                       dbCursor2.execute(SQLText, (self.UpdateHyperlinks(XMLText, recNumbers), QuoteNum))

                   progress.Update(95, _('Updating HyperLinks in Transcripts'))
                   # Get all Transcript records
                   SQLText = 'SELECT TranscriptNum, RTFText FROM Transcripts2'
                   SQLText = DBInterface.FixQuery(SQLText)
                   dbCursor.execute(SQLText)
                   # create the SQL for updating the RTFText of the Transcript
                   SQLText = """ UPDATE Transcripts2
                                 SET RTFText = %s
                                 WHERE TranscriptNum = %s """
                   SQLText = DBInterface.FixQuery(SQLText)
                   # For each Transcript record ...
                   for (QuoteNum, XMLText) in dbCursor.fetchall():
                       dbCursor2.execute(SQLText, (self.UpdateHyperlinks(XMLText, recNumbers), QuoteNum))

                   # Close the secondary database cursor
                   dbCursor2.close()

               # If we made it this far, we can commit the database transaction
               SQLText = 'COMMIT'
           else:
               # If contin is False, there's been an error and we should roll back the database transaction
               SQLText = 'ROLLBACK'

#           print "TRANSACTION ENDED!  1  ", SQLText
           
           # Execute the COMMIT or ROLLBACK
           dbCursor.execute(SQLText)
           dbCursor.close()

       # Handle IO Errors
       except IOError, e:
           filename = self.XMLFile.GetValue()
           if 'unicode' in wx.PlatformInfo:
               # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
               prompt = unicode(_('File "%s" was not found.  Try using the "Browse" button to locate your file.'), 'utf8')
           else:
               prompt = _('File "%s" was not found.  Try using the "Browse" button to locate your file.')
           errordlg = Dialogs.ErrorDialog(self, prompt % filename)
           errordlg.ShowModal()
           errordlg.Destroy()
           SQLText = 'ROLLBACK'

#           print "TRANSACTION ENDED!  2  ", SQLText
           
           dbCursor.execute(SQLText)
           dbCursor.close()
       except:
           if DEBUG or DEBUG_Exceptions:
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
           SQLText = 'ROLLBACK'

#           print "TRANSACTION ENDED!  3  ", SQLText
          
           dbCursor.execute(SQLText)
           dbCursor.close()

       try:
           f.close()
       except:
           pass

       # If importData is NOT passed in ...
       if self.importData == None:
           # .. then we need to update Transana's Database Tree, which we don't need to do when importData IS passed in.
           TransanaGlobal.menuWindow.ControlObject.DataWindow.DBTab.tree.refresh_tree()

       progress.Update(100)

       # If we are importing a pre-2.40 Transana-XML file ...
       if self.XMLVersionNumber in ['1.0', '1.1', '1.2', '1.3', '1.4']:
           # ... then we need to update the transcript records so that clips made from other clips are parented
           # by the source EPISODE, not the source CLIP.
           DBInterface.UpdateTranscriptRecsfor240(self)
       progress.Destroy()

       # DO NOT CLOSE THE DATABASE!!!!
       # db.close()


    def ProcessLine(self, txt):
        """ Process most lines read from the XML file to apply the proper encoding, if needed. """
        if 'unicode' in wx.PlatformInfo:

#            print "XMLImport.ProcessLine(1):", type(txt) # , txt.encode(self.importEncoding), self.importEncoding, self.XMLVersionNumber

            # If we're not reading a file encoded with UTF-8 encoding, we need to ...
            if (self.importEncoding != 'utf8'):
                # ... initialize a STRING variable.  (txt is UNICODE, but the WRONG ENCODING.  Damn.)
                s = ''

                # NOTE:  We shouldn't have to do this.  I must've screwed up the encoding at some point in XMLExport.py.
                #        In essence, we need txt.decode('utf8').decode(self.importEncoding), but that 

                if self.importEncoding != 'latin1':
                    # For each character in the unicode TXT string ...
                    for x in txt.decode('utf8'):
                        # ... add the appropriate character to the string S variable
                        s += chr(ord(x))
                else:
                    s = txt

                # Start Exception Handling.  (Japanese Filter Names weren't encoded right in 2.42, which makes this necessary.)
                try:
                    # Now we DECODE this using the the importEncoding selected by the user.
                    txt = s.decode(self.importEncoding)
                # If a UnicodeDecodeError is thrown ...
                except UnicodeDecodeError:
                    # .. just do a straight decode.  It's not right, but I can't figure out the right decoding,
                    # and at least this way the filters aren't LOST, just renamed to an unreadable form.
                    txt = txt.decode('utf8')

            # If we've got a String instead of a Unicode object ...
            elif type(txt) == str:
                # ... convert the string to Unicode using the import encoding
                txt = unicode(txt, self.importEncoding)
            # Process Escaped characters (& >, <)
            txt = self.UnEscape(txt)
            # If we ARE using UTF-8 ....
            if (self.importEncoding == 'utf8') and (self.XMLVersionNumber in ['1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7']):
                # ... perform the UTF-8 encoding needed for the database.
                txt = DBInterface.ProcessDBDataForUTF8Encoding(txt)

#            print "XMLImport.ProcessLine(2):", type(txt), txt.encode(self.importEncoding), self.importEncoding, self.XMLVersionNumber
#            print

        # Return the encoded text to the calling method
        return(txt)

    def UpdateHyperlinks(self, XMLText, recNumbers):

        if not ('url="transana:' in XMLText):
            return XMLText

        lines = XMLText.split('\n')
        XMLText = ''
        for line in lines:
            

           # We need to catch and update Hyperlinks!!  (A single line may have more than one Hyperlink!)
           if ('url="transana:' in line):

##               print line
               
               # Let's build a List of the parts of this line, to rebuild it later!
               lineArray = []
               while 'url="transana:' in line:
                   # Get the section BEFORE the URL
                   lineArray.append(line[:line.find('url="transana:')])
                   line = line[line.find('url="transana:'):]
                   # Get the first part of the URL
                   lineArray.append(line[:line.find(':') + 1])
                   line = line[line.find(':') + 1:]
                   # Process object by Type
                   if line[:5] == 'Quote':
                       linkType = 'Quote'
                   elif line[:4] == 'Clip':
                       linkType = 'Clip'
                   elif line[:8] == 'Snapshot':
                       linkType = 'Snapshot'
                   elif line[:4] == 'Note':
                       linkType = 'Note'

                   try:
                       # Get the object type
                       lineArray.append(line[:len(linkType) + 1])
                       line = line[len(linkType) + 1:]
                       linkNum = int(line[:line.find('"')])
                       lineArray.append("%s" % recNumbers[linkType][linkNum])
                       line = line[line.find('"'):]
                   except KeyError:
                       prompt = unicode(_("Bad hyperlink.  Linked %s no longer exists."), 'utf8')
                       # Display our carefully crafted error message to the user.
                       errordlg = Dialogs.ErrorDialog(None, prompt % linkType)
                       errordlg.ShowModal()
                       errordlg.Destroy()

                       # Substitute a 0 for the object number and complete the link edit
                       lineArray.append("%s" % 0)
                       line = line[line.find('"'):]
                   
               lineArray.append(line)
               line = ''
               for x in range(len(lineArray)):
                   line += "%s" % lineArray[x]

##               print line
##               print

           XMLText += line
        return XMLText

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
