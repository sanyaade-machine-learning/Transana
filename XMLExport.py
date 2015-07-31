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

"""This module implements the Data Export for Transana based on the Transana XML schema."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Ronnie Steinmetz, Jonathan Beavers'

DEBUG = False
DEBUG2 = False
if DEBUG or DEBUG2:
    print "XMLExport DEBUG is ON!!"

import wx

import Dialogs
import DBInterface
import TransanaConstants
import TransanaGlobal
if TransanaConstants.USESRTC:
    import RichTextEditCtrl_RTC
    import TranscriptEditor_STC  # for conversion
import RichTextEditCtrl
import cPickle
import datetime
import pickle
import os
# import Python's Regular Expresions
import re
import sys

# Set the encoding for export.
# Use UTF-8 regardless of the current encoding for consistency in the Transana XML files
EXPORT_ENCODING = 'utf8'
ENCODE_PROPERLY = True

class XMLExport(Dialogs.GenForm):
    """ This window displays a variety of GUI Widgets. """
    def __init__(self,parent,id,title):
        Dialogs.GenForm.__init__(self, parent, id, title, (550,150), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Export Database')

        # If we're using the RichTextCtrl ...
        if TransanaConstants.USESRTC:
            # Create an invisible RichTextCtrl so we can convert RTF format transcripts
            # to the RTC's XML format.  It's much faster for loading, and is XML compliant.
            import TranscriptEditor_RTC
            self.invisibleRTC = TranscriptEditor_RTC.TranscriptEditor(self)  # RichTextEditCtrl_RTC.RichTextEditCtrl(self)
            self.invisibleRTC.Show(False)

        # Create an invisible instance of RichTextEditCtrl.
        # If we're USING RTC, this allows us to convert STC formattted transcripts.
        # If we're NOT using RTC, this allows us to
        # get away with converting fastsaved documents to RTF behind the scenes.
        # Then we can simply pull the RTF data out of this object and write it
        # to the desired file.
        self.invisibleSTC = RichTextEditCtrl.RichTextEditCtrl(self)
        self.invisibleSTC.Show(False)

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create the main form prompt
        prompt = _('Please create an Transana XML File for export.')
        exportText = wx.StaticText(self.panel, -1, prompt)

        # Add the export message to the dialog box
        r1Sizer.Add(exportText, 0)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Add the Export File Name element
        self.XMLFile = self.new_edit_box(_("Transana-XML Filename"), v1, '')
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

    def Escape(self, inpStr):
        """ Replaces "&", "<", and ">" with the XML friendly "&amp;", &gt;", and "&lt;"
            >, <, and & all need to be replaced, but &amp;, &gt;, and &lt; needs to survive! """
        # We can't just casually replace ampersands, as there are ampersands in the replacement!
        # Find the first Ampersand in the document
        chrPos = inpStr.find('&')
        # While there are more Ampersands ...
        while (chrPos > -1):
            # ... replace the Ampersand with the escape string ...
            inpStr = inpStr[:chrPos] + '&amp;' + inpStr[chrPos + 1:]
            # ... and look for the NEXT Ampersand after the replacement
            chrPos = inpStr.find('&', chrPos + 1)
        # Find the first Greater Than in the document
        chrPos = inpStr.find('>')
        # While there are more Greater Thans ...
        while (chrPos > -1):
            # ... replace the Greater Than with the escape string ...
            inpStr = inpStr[:chrPos] + '&gt;' + inpStr[chrPos + 1:]
            # ... and look for the NEXT Greater Than after the replacement
            chrPos = inpStr.find('>', chrPos + 1)
        # Find the first Less Than in the document
        chrPos = inpStr.find('<')
        # While there are more Less Thans ...
        while (chrPos > -1):
            # ... replace the Less Than with the escape string ...
            inpStr = inpStr[:chrPos] + '&lt;' + inpStr[chrPos + 1:]
            # ... and look for the NEXT Less Than after the replacement
            chrPos = inpStr.find('<', chrPos + 1)
        # Return the modified string
        return inpStr

    def Export(self):
        # use the LONGEST title here!  That determines the size of the Dialog Box.
        progress = wx.ProgressDialog(_('Transana XML Export'), _('Exporting Transcript records (This may be slow because of the size of Transcript records.)'), style = wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)
        if progress.GetSize()[0] > 800:
            progress.SetSize((800, progress.GetSize()[1]))
            progress.Centre()

        db = DBInterface.get_db()

        # If we're using sqlite ...
        if TransanaConstants.DBInstalled in ['sqlite3']:
            # ... switch the text_factory from string to unicode here so we don't have to mess with decoding.
            db.text_factory = unicode

        try:
            fs = self.XMLFile.GetValue()
            if (fs[-4:].lower() != '.xml') and (fs[-4:].lower() != '.tra'):
                fs = fs + '.tra'
            # On the Mac, if no path is specified, the data is exported to a file INSIDE the application bundle, 
            # where no one will be able to find it.  Let's put it in the user's HOME directory instead.
            # I'm okay with not handling this on Windows, where it will be placed in the Program's folder
            # but it CAN be found.  (There's no easy way on Windows to determine the location of "My Documents"
            # especially if the user has moved it.)
            if "__WXMAC__" in wx.PlatformInfo:
                # if the specified file has no path specification ...
                if fs.find(os.sep) == -1:
                    # ... then prepend the HOME folder
                    fs = os.getenv("HOME") + os.sep + fs
            f = file(fs, 'w')
            progress.Update(0, _('Writing Headers'))
            self.WriteXMLDTD(f)

            f.write('<Transana>\n');
            f.write('  <TransanaXMLVersion>\n');
            # Version 1.0 -- Original Transana XML for Transana 2.0 release
            # Version 1.1 -- Unicode encoding added to Transana XML for Transana 2.1 release
            # Version 1.2 -- Filter Table added to Transana XML for Transana 2.11 release
            # Version 1.3 -- FilterData handling changed to accomodate Unicode data for Transana 2.21 release
            # Version 1.4 -- Database structure changed to accomodate Multiple Transcript Clips, BLOB keyword definitions
            #                for Transana 2.30 release.
            # Version 1.5 -- Database structure changed to accomodate Multiple Simultaneous Media Files for the
            #                Transana 2.40 release.
            # Version 1.6 -- Added XML format for transcripts and character escapes for Transana 2.50 release
            # Version 1.7 -- Database Structure changed for Still Image Snapshots for Transana 2.60 release
            # Version 1.8 -- Character Encoding Rules changed completely!!
            # Version 2.0 -- Transana 3.0, Documents and Quotes added
            if ENCODE_PROPERLY:
                f.write('    2.0\n')
            else:
                f.write('    1.7\n')
            f.write('  </TransanaXMLVersion>\n')

            progress.Update(5, _('Writing Library Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT SeriesNum, SeriesID, SeriesComment, SeriesOwner, DefaultKeywordGroup FROM Series2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <SeriesFile>\n')
                    for seriesRec in data:
                        self.WriteSeriesRec(f, seriesRec)
                    f.write('  </SeriesFile>\n')
                dbCursor.close()

            progress.Update(11, _('Writing Document Records  (This will seem slow because of the size of the Document Records.)'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT DocumentNum, DocumentID, LibraryNum, Author, Comment, ImportedFile, ImportDate, DocumentLength, XMLText FROM Documents2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <DocumentFile>\n')
                    for documentRec in data:
                        self.WriteDocumentRec(f, progress, documentRec)
                    f.write('  </DocumentFile>\n')
                dbCursor.close()

            progress.Update(17, _('Writing Episode Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT EpisodeNum, EpisodeID, SeriesNum, TapingDate, MediaFile, EpLength, EpComment FROM Episodes2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <EpisodeFile>\n')
                    for episodeRec in data:
                        self.WriteEpisodeRec(f, episodeRec)
                    f.write('  </EpisodeFile>\n')
                dbCursor.close()

            progress.Update(23, _('Writing Core Data Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = """SELECT CoreDataNum, Identifier, Title, Creator, Subject, Description, Publisher,
                                    Contributor, DCDate, DCType, Format, Source, Language, Relation, Coverage, Rights
                                    FROM CoreData2"""
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <CoreDataFile>\n')
                    for coreDataRec in data:
                        self.WriteCoreDataRec(f, coreDataRec)
                    f.write('  </CoreDataFile>\n')
                dbCursor.close()

            progress.Update(28, _('Writing Collection Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT CollectNum, CollectID, ParentCollectNum, CollectComment, CollectOwner, DefaultKeywordGroup FROM Collections2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <CollectionFile>\n')
                    for collectionRec in data:
                        self.WriteCollectionRec(f, collectionRec)
                    f.write('  </CollectionFile>\n')
                dbCursor.close()

            progress.Update(34, _('Writing Quote Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT QuoteNum, QuoteID, CollectNum, SourceDocumentNum, SortOrder, Comment, XMLText FROM Quotes2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <QuoteFile>\n')
                    for quoteRec in data:
                        self.WriteQuoteRec(f, quoteRec)
                    f.write('  </QuoteFile>\n')
                dbCursor.close()

            progress.Update(39, _('Writing Quote Position Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT QuoteNum, DocumentNum, StartChar, EndChar FROM QuotePositions2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <QuotePositionFile>\n')
                    for quotePosRec in data:
                        self.WriteQuotePosRec(f, quotePosRec)
                    f.write('  </QuotePositionFile>\n')
                dbCursor.close()

            progress.Update(45, _('Writing Clip Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT ClipNum, ClipID, CollectNum, EpisodeNum, MediaFile, ClipStart, ClipStop, ClipOffset, Audio, '
                SQLText += 'ClipComment, SortOrder FROM Clips2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <ClipFile>\n')
                    for clipRec in data:
                        self.WriteClipRec(f, clipRec)
                    f.write('  </ClipFile>\n')
                dbCursor.close()

            progress.Update(50, _('Writing Additional Media File Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT AddVidNum, EpisodeNum, ClipNum, MediaFile, VidLength, Offset, Audio FROM AdditionalVids2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <AdditionalVidsFile>\n')
                    for additionalMediaFileRec in data:
                        self.WriteAdditionalMediaFileRec(f, additionalMediaFileRec)
                    f.write('  </AdditionalVidsFile>\n')
                dbCursor.close()

            progress.Update(56, _('Writing Transcript Records  (This will seem slow because of the size of the Transcript Records.)'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT TranscriptNum, TranscriptID, EpisodeNum, SourceTranscriptNum, ClipNum, SortOrder, Transcriber, '
                SQLText += 'ClipStart, ClipStop, Comment, MinTranscriptWidth, RTFText FROM Transcripts2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <TranscriptFile>\n')
                    for transcriptRec in data:
                        self.WriteTranscriptRec(f, progress, transcriptRec)
                    f.write('  </TranscriptFile>\n')
                dbCursor.close()

            progress.Update(62, _('Writing Snapshot Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT SnapshotNum, SnapshotID, CollectNum, ImageFile, ImageScale, ImageCoordsX, ImageCoordsY, '
                SQLText += 'ImageSizeW, ImageSizeH, EpisodeNum, TranscriptNum, SnapshotTimeCode, SnapshotDuration, '
                SQLText += 'SnapshotComment, SortOrder FROM Snapshots2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <SnapshotFile>\n')
                    for snapshotRec in data:
                        self.WriteSnapshotRec(f, snapshotRec)
                    f.write('  </SnapshotFile>\n')
                dbCursor.close()

            progress.Update(68, _('Writing Keyword Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT KeywordGroup, Keyword, Definition, LineColorName, LineColorDef, DrawMode, LineWidth, LineStyle FROM Keywords2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <KeywordFile>\n')
                    for keywordRec in data:
                        self.WriteKeywordRec(f, keywordRec)
                    f.write('  </KeywordFile>\n')
                dbCursor.close()

            progress.Update(73, _('Writing Clip Keyword Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT EpisodeNum, DocumentNum, ClipNum, QuoteNum, SnapshotNum, KeywordGroup, Keyword, Example FROM ClipKeywords2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <ClipKeywordFile>\n')
                    for clipKeywordRec in data:
                        self.WriteClipKeywordRec(f, clipKeywordRec)
                    f.write('  </ClipKeywordFile>\n')
                dbCursor.close()

            progress.Update(79, _('Writing Snapshot Keywords Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT SnapshotNum, KeywordGroup, Keyword, x1, y1, x2, y2, visible FROM SnapshotKeywords2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <SnapshotKeywordFile>\n')
                    for snapshotKeywordRec in data:
                        self.WriteSnapshotKeywordRec(f, snapshotKeywordRec)
                    f.write('  </SnapshotKeywordFile>\n')
                dbCursor.close()

            progress.Update(84, _('Writing Snapshot Coding Style Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT SnapshotNum, KeywordGroup, Keyword, DrawMode, LineColorName, LineColorDef, LineWidth, LineStyle '
                SQLText += 'FROM SnapshotKeywordStyles2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <SnapshotKeywordStyleFile>\n')
                    for snapshotKeywordStyleRec in data:
                        self.WriteSnapshotKeywordStyleRec(f, snapshotKeywordStyleRec)
                    f.write('  </SnapshotKeywordStyleFile>\n')
                dbCursor.close()

            progress.Update(90, _('Writing Note Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, CollectNum, ClipNum, SnapshotNum, DocumentNum, '
                SQLText += 'QuoteNum, TranscriptNum, NoteTaker, NoteText FROM Notes2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <NoteFile>\n')
                    for noteRec in data:
                        self.WriteNoteRec(f, noteRec)
                    f.write('  </NoteFile>\n')
                dbCursor.close()

            progress.Update(95, _('Writing Filter Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT ReportType, ReportScope, ConfigName, FilterDataType, FilterData FROM Filters2'
                dbCursor.execute(SQLText)
                data = dbCursor.fetchall()
                if len(data) > 0:
                    f.write('  <FilterFile>\n')
                    for filterRec in data:
                        self.WriteFilterRec(f, filterRec)
                    f.write('  </FilterFile>\n')
                dbCursor.close()

            f.write('</Transana>\n');

            f.flush()

            dbCursor.close()
            
        except:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('An error occurred during Database Export.\n%s\n%s'), 'utf8')
            else:
                prompt = _('An error occurred during Database Export.\n%s\n%s')
            errordlg = Dialogs.ErrorDialog(self, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
            errordlg.ShowModal()
            errordlg.Destroy()

            if DEBUG or DEBUG2:
                import traceback
                traceback.print_exc(file=sys.stdout)
            dbCursor.close()
        finally:
            # If we're using sqlite ...
            if TransanaConstants.DBInstalled in ['sqlite3']:
                # ... we need to go back to using strings rather than unicode objects from the database
                db.text_factory = str

        f.close()
        progress.Update(100)
        progress.Destroy()

    def WriteXMLDTD(self, f):
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write('<!DOCTYPE TransanaData [\n')
        f.write('  <!ELEMENT TransanaXMLVersion (#PCDATA)>\n')
        f.write('  <!ELEMENT SeriesFile (Series)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT Num (#PCDATA)>\n')
        f.write('  <!ELEMENT ID (#PCDATA)>\n')
        f.write('  <!ELEMENT Comment (#PCDATA)>\n')
        f.write('  <!ELEMENT Owner (#PCDATA)>\n')
        f.write('  <!ELEMENT DefaultKeywordGroup (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT Series (#PCDATA|Num|ID|Comment|Owner|DefaultKeywordGroup)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT DocumentFile (Document)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT SeriesNum (#PCDATA)>\n')
        f.write('  <!ELEMENT NoteTaker (#PCDATA)>\n')
        f.write('  <!ELEMENT MediaFile (#PCDATA)>\n')
        f.write('  <!ELEMENT Date (#PCDATA)>\n')
        f.write('  <!ELEMENT Length (#PCDATA)>\n')
        f.write('  <!ELEMENT XMLText (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT Document (#PCDATA|Num|ID|SeriesNum|Author|Comment|MediaFile|Date|Length|XMLText)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT EpisodeFile (Episode)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT Episode (#PCDATA|Num|ID|SeriesNum|Date|MediaFile|Length|Comment)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT CoreDataFile (CoreData)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT Title (#PCDATA)>\n')
        f.write('  <!ELEMENT Creator (#PCDATA)>\n')
        f.write('  <!ELEMENT Subject (#PCDATA)>\n')
        f.write('  <!ELEMENT Description (#PCDATA)>\n')
        f.write('  <!ELEMENT Publisher (#PCDATA)>\n')
        f.write('  <!ELEMENT Contributor (#PCDATA)>\n')
        f.write('  <!ELEMENT Type (#PCDATA)>\n')
        f.write('  <!ELEMENT Format (#PCDATA)>\n')
        f.write('  <!ELEMENT Source (#PCDATA)>\n')
        f.write('  <!ELEMENT Language (#PCDATA)>\n')
        f.write('  <!ELEMENT Relation (#PCDATA)>\n')
        f.write('  <!ELEMENT Coverage (#PCDATA)>\n')
        f.write('  <!ELEMENT Rights (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT CoreData (#PCDATA|Num|ID|Title|Creator|Subject|Description|Publisher|Contributor|Date|Type|Format|Source|Language|Relation|Coverage|Rights)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT TranscriptFile (Transcript)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT EpisodeNum (#PCDATA)>\n')
        f.write('  <!ELEMENT TranscriptNum (#PCDATA)>\n')
        f.write('  <!ELEMENT ClipNum (#PCDATA)>\n')
        f.write('  <!ELEMENT SortOrder (#PCDATA)>\n')
        f.write('  <!ELEMENT Transcriber (#PCDATA)>\n')
        f.write('  <!ELEMENT ClipStart (#PCDATA)>\n')
        f.write('  <!ELEMENT ClipStop (#PCDATA)>\n')
        f.write('  <!ELEMENT MinTranscriptWidth (#PCDATA)>\n')
        f.write('  <!ELEMENT RTFText (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT Transcript (#PCDATA|Num|ID|EpisodeNum|TranscriptNum|ClipNum|SortOrder|Transcriber|ClipStart|ClipStop|Comment|RTFText)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT CollectionFile (Collection)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT ParentCollectNum (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT Collection (#PCDATA|Num|ID|ParentCollectNum|Comment|Owner|DefaultKeywordGroup)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT QuoteFile (Quote)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT DocumentNum (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT Quote (#PCDATA|Num|ID|CollectNum|DocumentNum|SortOrder|Comment|XMLText)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT QuotePositionFile (QuotePosition)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT StartChar (#PCDATA)>\n')
        f.write('  <!ELEMENT EndChar (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT QuotePosition (#PCDATA|Num|DocumentNum|StartChar|EndChar)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT ClipFile (Clip)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT CollectNum (#PCDATA)>\n')
        f.write('  <!ELEMENT Offset (#PCDATA)>\n')
        f.write('  <!ELEMENT Audio (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT Clip (#PCDATA|Num|ID|CollectNum|EpisodeNum|MediaFile|ClipStart|ClipStop|Offset|Audio|Comment|SortOrder)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT SnapshotFile (Snapshot)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT ImageScale (#PCDATA)>\n')
        f.write('  <!ELEMENT ImageCoordsX (#PCDATA)>\n')
        f.write('  <!ELEMENT ImageCoordsY (#PCDATA)>\n')
        f.write('  <!ELEMENT ImageSizeW (#PCDATA)>\n')
        f.write('  <!ELEMENT ImageSizeH (#PCDATA)>\n')
        f.write('  <!ELEMENT SnapshotTimeCode (#PCDATA)>\n')
        f.write('  <!ELEMENT SnapshotDuration (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT Snapshot (#PCDATA|Num|ID|CollectNum|MediaFile|ImageScale|ImageCoordsX|ImageCoordsY|ImageSizeW|ImageSizeH|EpisodeNum|TranscriptNum|SnapshotTimeCode|SnapshotDuration|Comment|SortOrder)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT AdditionalVids (AdditionalVidRec)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT AddVid (#PCDATA|Num|EpisodeNum|ClipNum|MediaFile|Length|Offset|Audio)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT KeywordFile (KeywordRec)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT KeywordGroup (#PCDATA)>\n')
        f.write('  <!ELEMENT Keyword (#PCDATA)>\n')
        f.write('  <!ELEMENT Definition (#PCDATA)>\n')
        f.write('  <!ELEMENT ColorName (#PCDATA)>\n')
        f.write('  <!ELEMENT ColorDef (#PCDATA)>\n')
        f.write('  <!ELEMENT DrawMode (#PCDATA)>\n')
        f.write('  <!ELEMENT LineWidth (#PCDATA)>\n')
        f.write('  <!ELEMENT LineStyle (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT KeywordRec (#PCDATA|KeywordGroup|Keyword|Definition|ColorName|ColorDef|DrawMode|LineWidth|LineStyle)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT ClipKeywordFile (ClipKeyword)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT DocumentNum (#PCDATA)>\n')
        f.write('  <!ELEMENT QuoteNum (#PCDATA)>\n')
        f.write('  <!ELEMENT SnapshotNum (#PCDATA)>\n')
        f.write('  <!ELEMENT Example (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT ClipKeyword (#PCDATA|EpisodeNum|DocumentNum|ClipNum|QuoteNum|SnapshotNum|KeywordGroup|Keyword|Example)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT SnapshotKeywordFile (SnapshotKeyword)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT X1 (#PCDATA)>\n')
        f.write('  <!ELEMENT Y1 (#PCDATA)>\n')
        f.write('  <!ELEMENT X2 (#PCDATA)>\n')
        f.write('  <!ELEMENT Y2 (#PCDATA)>\n')
        f.write('  <!ELEMENT Visible (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT SnapshotKeyword (#PCDATA|SnapshotNum|KeywordGroup|Keyword|X1|Y1|X2|Y2|Visible)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT SnapshotKeywordStyleFile (SnapshotKeywordStyle)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT SnapshotKeywordStyle (#PCDATA|SnapshotNum|KeywordGroup|Keyword|DrawMode|ColorName|ColorDef|LineWidth|LineStyle)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT NoteFile (Note)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT NoteText (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT Note (#PCDATA|Num|ID|SeriesNum|EpisodeNum|CollectNum|ClipNum|TranscriptNum|NoteTaker|NoteText)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT FilterFile (Filter)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT ReportType (#PCDATA)>\n')
        f.write('  <!ELEMENT ReportScope (#PCDATA)>\n')
        f.write('  <!ELEMENT ConfigName (#PCDATA)>\n')
        f.write('  <!ELEMENT FilterDataType (#PCDATA)>\n')
        f.write('  <!ELEMENT FilterData (#PCDATA)>\n')
        f.write('\n')
        f.write('  <!ELEMENT Filter (#PCDATA|ReportType|ReportScope|ConfigName|FilterDataType|FilterData)*>\n')
        f.write('\n')
        f.write('  <!ELEMENT Transana (#PCDATA|SeriesFile|EpisodeFile|CoreDataFile|TranscriptFile|CollectionFile|ClipFile|KeywordFile|ClipKeywordFile|NoteFile|FilterFile)*>\n')
        f.write(']>\n')
        f.write('\n')

    def WriteSeriesRec(self, f, seriesRec):
        (SeriesNum, SeriesID, SeriesComment, SeriesOwner, DefaultKeywordGroup) = seriesRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            SeriesID = DBInterface.ProcessDBDataForUTF8Encoding(SeriesID)
            SeriesComment = DBInterface.ProcessDBDataForUTF8Encoding(SeriesComment)
            SeriesOwner = DBInterface.ProcessDBDataForUTF8Encoding(SeriesOwner)
            DefaultKeywordGroup = DBInterface.ProcessDBDataForUTF8Encoding(DefaultKeywordGroup)

        f.write('    <Series>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % SeriesNum)
        f.write('      </Num>\n')
        f.write('      <ID>\n')
        f.write('        %s\n' % self.Escape(SeriesID.encode(EXPORT_ENCODING)))
        f.write('      </ID>\n')
        if SeriesComment != '':
            f.write('      <Comment>\n')
            f.write('        %s\n' % self.Escape(SeriesComment.encode(EXPORT_ENCODING)))
            f.write('      </Comment>\n')
        if SeriesOwner != '':
            f.write('      <Owner>\n')
            f.write('        %s\n' % self.Escape(SeriesOwner.encode(EXPORT_ENCODING)))
            f.write('      </Owner>\n')
        if DefaultKeywordGroup != '':
            f.write('      <DefaultKeywordGroup>\n')
            f.write('        %s\n' % self.Escape(DefaultKeywordGroup.encode(EXPORT_ENCODING)))
            f.write('      </DefaultKeywordGroup>\n')
        f.write('    </Series>\n')

    def WriteDocumentRec(self, f, progress, documentRec):
        (DocumentNum, DocumentID, LibraryNum, Author, Comment, ImportedFile, ImportDate, DocumentLength, XMLText) = documentRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            DocumentID = DBInterface.ProcessDBDataForUTF8Encoding(DocumentID)
            Author = DBInterface.ProcessDBDataForUTF8Encoding(Author)
            Comment = DBInterface.ProcessDBDataForUTF8Encoding(Comment)
            ImportedFile = DBInterface.ProcessDBDataForUTF8Encoding(ImportedFile)

        f.write('    <Document>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % DocumentNum)
        f.write('      </Num>\n')
        f.write('      <ID>\n')
        f.write('        %s\n' % self.Escape(DocumentID.encode(EXPORT_ENCODING)))

        if DEBUG:
            try:
                print "Document ID ", DocumentID
            except:
                print "Document Number ", DocumentNum
                
        f.write('      </ID>\n')
        f.write('      <SeriesNum>\n')
        f.write('        %s\n' % LibraryNum)
        f.write('      </SeriesNum>\n')
        if (Author != None) and (Author != ''):
            f.write('      <NoteTaker>\n')
            f.write('        %s\n' % self.Escape(Author.encode(EXPORT_ENCODING)))
            f.write('      </NoteTaker>\n')
        if (Comment != None) and (Comment != ''):
            f.write('      <Comment>\n')
            f.write('        %s\n' % self.Escape(Comment.encode(EXPORT_ENCODING)))
            f.write('      </Comment>\n')
        if (ImportedFile != None) and (ImportedFile != ''):
            f.write('      <MediaFile>\n')
            f.write('        %s\n' % self.Escape(ImportedFile.encode(EXPORT_ENCODING)))
            f.write('      </MediaFile>\n')
        if ImportDate != None:
            f.write('      <Date>\n')
            f.write('        %s\n' % ImportDate)
            f.write('      </Date>\n')
        if (DocumentLength != '') and (DocumentLength > 0):
            f.write('      <Length>\n')
            f.write('        %s\n' % DocumentLength)
            f.write('      </Length>\n')
        if XMLText != '':
            # Extract the XML Text from the DB's array structure, if needed
            # Okay, this isn't so straight-forward any more.
            # With MySQL for Python 0.9.x, XMLText is of type str.
            # With MySQL for Python 1.2.0, XMLText is of type array.  It could then either be a
            # character string (typecode == 'c') or a unicode string (typecode == 'u'), which then
            # need to be interpreted differently.
            if type(XMLText).__name__ == 'array':
                if XMLText.typecode == 'u':
                    XMLText = XMLText.tounicode()
                else:
                    XMLText = XMLText.tostring()
            elif isinstance(XMLText, unicode):

                XMLText = XMLText.encode('utf8')
                
            f.write('      <XMLText>\n')

            if DEBUG2:
                print "type(XMLText) =", type(XMLText),

            # Unlike with Transcripts, we should *ALWAYS* have XML Text here
            if (type(XMLText).__name__ != 'NoneType') and (len(XMLText) > 5) and (XMLText[:5].upper() == '<?XML'):

                if DEBUG2:
                    print "XML Type"
                    
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt1 = unicode(_('Writing Document Records  (This will seem slow because of the size of the Document Records.)'), 'utf8')
                prompt2 = '\n' + unicode(_('Exporting %s'), 'utf8')

                progress.Update(11, prompt1 + prompt2 % DocumentID)
                progress.Refresh()

                # If XML, we can just use it as is after we strip off the XML Header Line, which breaks XML.
                # Due to changes in the header across wxPython versions, we need to use a regular expression.
                # xmlData = XMLText[39:]

                # Match "<?xml" plus any number of characters that are not ">" plus ">" plus whitespace
                REGEX = "<\?xml[^>]*>[\s]*"
                regexCompiled = re.compile(REGEX)
                regexSearch = regexCompiled.search(XMLText)
                xmlData = XMLText[regexSearch.end():]

            # now simply write the XML data to the file.  (This does NOT need to be encoded, as the XML already is!)
            # (but check to make sure there's actually XML data there!)
            if XMLText != None:
                f.write('%s' % XMLText.rstrip())
            # ... add an extra line break here!
            f.write('\n')
            f.write('      </XMLText>\n')

            # On exporting old, huge databases, you can end up with hundreds of "Importing" popups.
            # This hopefully will allow them to close properly rather than building up.
            wx.YieldIfNeeded()

        f.write('    </Document>\n')

    def WriteEpisodeRec(self, f, episodeRec):
        (EpisodeNum, EpisodeID, SeriesNum, TapingDate, MediaFile, EpLength, EpComment) = episodeRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            EpisodeID = DBInterface.ProcessDBDataForUTF8Encoding(EpisodeID)
            MediaFile = DBInterface.ProcessDBDataForUTF8Encoding(MediaFile)
            EpComment = DBInterface.ProcessDBDataForUTF8Encoding(EpComment)

        f.write('    <Episode>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % EpisodeNum)
        f.write('      </Num>\n')
        f.write('      <ID>\n')
        f.write('        %s\n' % self.Escape(EpisodeID.encode(EXPORT_ENCODING)))
        f.write('      </ID>\n')
        f.write('      <SeriesNum>\n')
        f.write('        %s\n' % SeriesNum)
        f.write('      </SeriesNum>\n')
        if TapingDate != None:
            f.write('      <Date>\n')
            f.write('        %s\n' % TapingDate)
            f.write('      </Date>\n')
        f.write('      <MediaFile>\n')
        f.write('        %s\n' % self.Escape(MediaFile.encode(EXPORT_ENCODING)))
        f.write('      </MediaFile>\n')
        if (EpLength != '') and (EpLength != 0):
            f.write('      <Length>\n')
            f.write('        %s\n' % EpLength)
            f.write('      </Length>\n')
        if EpComment != '':
            f.write('      <Comment>\n')
            f.write('        %s\n' % self.Escape(EpComment.encode(EXPORT_ENCODING)))
            f.write('      </Comment>\n')
        f.write('    </Episode>\n')

    def WriteCoreDataRec(self, f, coreDataRec):
        (CoreDataNum, Identifier, Title, Creator, Subject, Description, Publisher,
         Contributor, DCDate, DCType, Format, Source, Language, Relation, Coverage, Rights) = coreDataRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            Identifier = DBInterface.ProcessDBDataForUTF8Encoding(Identifier)
            Title = DBInterface.ProcessDBDataForUTF8Encoding(Title)
            Creator = DBInterface.ProcessDBDataForUTF8Encoding(Creator)
            Subject = DBInterface.ProcessDBDataForUTF8Encoding(Subject)
            Description = DBInterface.ProcessDBDataForUTF8Encoding(Description)
            Publisher = DBInterface.ProcessDBDataForUTF8Encoding(Publisher)
            Contributor = DBInterface.ProcessDBDataForUTF8Encoding(Contributor)
            DCType = DBInterface.ProcessDBDataForUTF8Encoding(DCType)
            Format = DBInterface.ProcessDBDataForUTF8Encoding(Format)
            Source = DBInterface.ProcessDBDataForUTF8Encoding(Source)
            Language = DBInterface.ProcessDBDataForUTF8Encoding(Language)
            Relation = DBInterface.ProcessDBDataForUTF8Encoding(Relation)
            Coverage = DBInterface.ProcessDBDataForUTF8Encoding(Coverage)
            Rights = DBInterface.ProcessDBDataForUTF8Encoding(Rights)

        f.write('    <CoreData>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % CoreDataNum)
        f.write('      </Num>\n')
        f.write('      <ID>\n')
        f.write('        %s\n' % self.Escape(Identifier.encode(EXPORT_ENCODING)))
        f.write('      </ID>\n')
        if Title != '':
            f.write('      <Title>\n')
            f.write('        %s\n' % self.Escape(Title.encode(EXPORT_ENCODING)))
            f.write('      </Title>\n')
        if Creator != '':
            f.write('      <Creator>\n')
            f.write('        %s\n' % self.Escape(Creator.encode(EXPORT_ENCODING)))
            f.write('      </Creator>\n')
        if Subject != '':
            f.write('      <Subject>\n')
            f.write('        %s\n' % self.Escape(Subject.encode(EXPORT_ENCODING)))
            f.write('      </Subject>\n')
        if Description != '':
            f.write('      <Description>\n')
            f.write('        %s\n' % self.Escape(Description.encode(EXPORT_ENCODING)))
            f.write('      </Description>\n')
        if Publisher != '':
            f.write('      <Publisher>\n')
            f.write('        %s\n' % self.Escape(Publisher.encode(EXPORT_ENCODING)))
            f.write('      </Publisher>\n')
        if Contributor != '':
            f.write('      <Contributor>\n')
            f.write('        %s\n' % self.Escape(Contributor.encode(EXPORT_ENCODING)))
            f.write('      </Contributor>\n')
        if DCDate != None:
            # If we have a string or unicode object ...
            if isinstance(DCDate, str) or isinstance(DCDate, unicode):
                # ... convert it to a datetime object!
                DCDate = datetime.datetime.strptime(DCDate, '%Y/%m/%d')
            f.write('      <Date>\n')
            f.write('        %s/%s/%s\n' % (DCDate.month, DCDate.day, DCDate.year))
            f.write('      </Date>\n')
        if DCType != '':
            f.write('      <Type>\n')
            f.write('        %s\n' % self.Escape(DCType.encode(EXPORT_ENCODING)))
            f.write('      </Type>\n')
        if Format != '':
            f.write('      <Format>\n')
            f.write('        %s\n' % self.Escape(Format.encode(EXPORT_ENCODING)))
            f.write('      </Format>\n')
        if Source != '':
            f.write('      <Source>\n')
            f.write('        %s\n' % self.Escape(Source.encode(EXPORT_ENCODING)))
            f.write('      </Source>\n')
        if Language != '':
            f.write('      <Language>\n')
            f.write('        %s\n' % self.Escape(Language.encode(EXPORT_ENCODING)))
            f.write('      </Language>\n')
        if Relation != '':
            f.write('      <Relation>\n')
            f.write('        %s\n' % self.Escape(Relation.encode(EXPORT_ENCODING)))
            f.write('      </Relation>\n')
        if Coverage != '':
            f.write('      <Coverage>\n')
            f.write('        %s\n' % self.Escape(Coverage.encode(EXPORT_ENCODING)))
            f.write('      </Coverage>\n')
        if Rights != '':
            f.write('      <Rights>\n')
            f.write('        %s\n' % self.Escape(Rights.encode(EXPORT_ENCODING)))
            f.write('      </Rights>\n')
        f.write('    </CoreData>\n')

    def WriteCollectionRec(self, f, collectionRec):
        (CollectNum, CollectID, ParentCollectNum, CollectComment, CollectOwner, DefaultKeywordGroup) = collectionRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            CollectID = DBInterface.ProcessDBDataForUTF8Encoding(CollectID)
            CollectComment = DBInterface.ProcessDBDataForUTF8Encoding(CollectComment)
            CollectOwner = DBInterface.ProcessDBDataForUTF8Encoding(CollectOwner)
            DefaultKeywordGroup = DBInterface.ProcessDBDataForUTF8Encoding(DefaultKeywordGroup)

        f.write('    <Collection>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % CollectNum)
        f.write('      </Num>\n')
        f.write('      <ID>\n')
        f.write('        %s\n' % self.Escape(CollectID.encode(EXPORT_ENCODING)))
        f.write('      </ID>\n')
        if (ParentCollectNum != '') and (ParentCollectNum != 0):
            f.write('      <ParentCollectNum>\n')
            f.write('        %s\n' % ParentCollectNum)
            f.write('      </ParentCollectNum>\n')
        if CollectComment != '':
            f.write('      <Comment>\n')
            f.write('        %s\n' % self.Escape(CollectComment.encode(EXPORT_ENCODING)))
            f.write('      </Comment>\n')
        if CollectOwner != '':
            f.write('      <Owner>\n')
            f.write('        %s\n' % self.Escape(CollectOwner.encode(EXPORT_ENCODING)))
            f.write('      </Owner>\n')
        if DefaultKeywordGroup != '':
            f.write('      <DefaultKeywordGroup>\n')
            f.write('        %s\n' % self.Escape(DefaultKeywordGroup.encode(EXPORT_ENCODING)))
            f.write('      </DefaultKeywordGroup>\n')
        f.write('    </Collection>\n')

    def WriteQuoteRec(self, f, quoteRec):
        (QuoteNum, QuoteID, CollectNum, SourceDocumentNum, SortOrder, Comment, XMLText) = quoteRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            QuoteID = DBInterface.ProcessDBDataForUTF8Encoding(QuoteID)
            Comment = DBInterface.ProcessDBDataForUTF8Encoding(Comment)

        f.write('    <Quote>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % QuoteNum)
        f.write('      </Num>\n')
        f.write('      <ID>\n')
        f.write('        %s\n' % self.Escape(QuoteID.encode(EXPORT_ENCODING)))
        f.write('      </ID>\n')
        if CollectNum != None:
            f.write('      <CollectNum>\n')
            f.write('        %s\n' % CollectNum)
            f.write('      </CollectNum>\n')
        if SourceDocumentNum != None:
            f.write('      <DocumentNum>\n')
            f.write('        %s\n' % SourceDocumentNum)
            f.write('      </DocumentNum>\n')
        if (SortOrder != '') and (SortOrder != 0):
            f.write('      <SortOrder>\n')
            f.write('        %s\n' % SortOrder)
            f.write('      </SortOrder>\n')
        if Comment != '':
            f.write('      <Comment>\n')
            f.write('        %s\n' % self.Escape(Comment.encode(EXPORT_ENCODING)))
            f.write('      </Comment>\n')
        if XMLText != '':
            # Extract the XML Text from the DB's array structure, if needed
            # Okay, this isn't so straight-forward any more.
            # With MySQL for Python 0.9.x, XMLText is of type str.
            # With MySQL for Python 1.2.0, XMLText is of type array.  It could then either be a
            # character string (typecode == 'c') or a unicode string (typecode == 'u'), which then
            # need to be interpreted differently.
            if type(XMLText).__name__ == 'array':
                if XMLText.typecode == 'u':
                    XMLText = XMLText.tounicode()
                else:
                    XMLText = XMLText.tostring()
            elif isinstance(XMLText, unicode):

                XMLText = XMLText.encode('utf8')
                
            f.write('      <XMLText>\n')

            if DEBUG2:
                print "type(XMLText) =", type(XMLText),

            # Unlike with Transcripts, we should *ALWAYS* have XML Text here
            if (type(XMLText).__name__ != 'NoneType') and (len(XMLText) > 5) and (XMLText[:5].upper() == '<?XML'):

                if DEBUG2:
                    print "XML Type"
                    
##                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
##                prompt1 = unicode(_('Writing Document Records  (This will seem slow because of the size of the Document Records.)'), 'utf8')
##                prompt2 = '\n' + unicode(_('Exporting %s'), 'utf8')

##                progress.Update(11, prompt1 + prompt2 % DocumentID)
##                progress.Refresh()

                # If XML, we can just use it as is after we strip off the XML Header Line, which breaks XML.
                # Due to changes in the header across wxPython versions, we need to use a regular expression.
                # xmlData = XMLText[39:]

                # Match "<?xml" plus any number of characters that are not ">" plus ">" plus whitespace
                REGEX = "<\?xml[^>]*>[\s]*"
                regexCompiled = re.compile(REGEX)
                regexSearch = regexCompiled.search(XMLText)
                xmlData = XMLText[regexSearch.end():]

            # now simply write the XML data to the file.  (This does NOT need to be encoded, as the XML already is!)
            # (but check to make sure there's actually XML data there!)
            if XMLText != None:
                f.write('%s' % XMLText.rstrip())
            # ... add an extra line break here!
            f.write('\n')
            f.write('      </XMLText>\n')

            # On exporting old, huge databases, you can end up with hundreds of "Importing" popups.
            # This hopefully will allow them to close properly rather than building up.
            wx.YieldIfNeeded()

        f.write('    </Quote>\n')

    def WriteQuotePosRec(self, f, quotePosRec):
        (QuoteNum, DocumentNum, StartChar, EndChar) = quotePosRec

        f.write('    <QuotePosition>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % QuoteNum)
        f.write('      </Num>\n')
        if (DocumentNum != None) and (DocumentNum > 0):
            f.write('      <DocumentNum>\n')
            f.write('        %s\n' % DocumentNum)
            f.write('      </DocumentNum>\n')
        f.write('      <StartChar>\n')
        f.write('        %s\n' % StartChar)
        f.write('      </StartChar>\n')
        f.write('      <EndChar>\n')
        f.write('        %s\n' % EndChar)
        f.write('      </EndChar>\n')
        f.write('    </QuotePosition>\n')

    def WriteClipRec(self, f, clipRec):
        (ClipNum, ClipID, CollectNum, EpisodeNum, MediaFile, ClipStart, ClipStop, ClipOffset, ClipAudio, ClipComment, SortOrder) = clipRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            ClipID = DBInterface.ProcessDBDataForUTF8Encoding(ClipID)
            MediaFile = DBInterface.ProcessDBDataForUTF8Encoding(MediaFile)
            ClipComment = DBInterface.ProcessDBDataForUTF8Encoding(ClipComment)

        f.write('    <Clip>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % ClipNum)
        f.write('      </Num>\n')
        f.write('      <ID>\n')
        f.write('        %s\n' % self.Escape(ClipID.encode(EXPORT_ENCODING)))
        f.write('      </ID>\n')
        if CollectNum != None:
            f.write('      <CollectNum>\n')
            f.write('        %s\n' % CollectNum)
            f.write('      </CollectNum>\n')
        if EpisodeNum != None:
            f.write('      <EpisodeNum>\n')
            f.write('        %s\n' % EpisodeNum)
            f.write('      </EpisodeNum>\n')
        f.write('      <MediaFile>\n')
        f.write('        %s\n' % self.Escape(MediaFile.encode(EXPORT_ENCODING)))
        f.write('      </MediaFile>\n')
        f.write('      <ClipStart>\n')
        f.write('        %s\n' % ClipStart)
        f.write('      </ClipStart>\n')
        f.write('      <ClipStop>\n')
        f.write('        %s\n' % ClipStop)
        f.write('      </ClipStop>\n')
        if ClipOffset != 0:
            f.write('      <Offset>\n')
            f.write('        %s\n' % ClipOffset)
            f.write('      </Offset>\n')
        f.write('      <Audio>\n')
        f.write('        %s\n' % ClipAudio)
        f.write('      </Audio>\n')
        if ClipComment != '':
            f.write('      <Comment>\n')
            f.write('        %s\n' % self.Escape(ClipComment.encode(EXPORT_ENCODING)))
            f.write('      </Comment>\n')
        if (SortOrder != '') and (SortOrder != 0):
            f.write('      <SortOrder>\n')
            f.write('        %s\n' % SortOrder)
            f.write('      </SortOrder>\n')
        f.write('    </Clip>\n')

    def WriteAdditionalMediaFileRec(self, f, additionalVidRec):
        (AddVidNum, EpisodeNum, ClipNum, MediaFile, VidLength, Offset, Audio) = additionalVidRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            MediaFile = DBInterface.ProcessDBDataForUTF8Encoding(MediaFile)

        f.write('    <AddVid>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % AddVidNum)
        f.write('      </Num>\n')
        if EpisodeNum > 0:
            f.write('      <EpisodeNum>\n')
            f.write('        %s\n' % EpisodeNum)
            f.write('      </EpisodeNum>\n')
        if ClipNum > 0:
            f.write('      <ClipNum>\n')
            f.write('        %s\n' % ClipNum)
            f.write('      </ClipNum>\n')
        f.write('      <MediaFile>\n')
        f.write('        %s\n' % self.Escape(MediaFile.encode(EXPORT_ENCODING)))
        f.write('      </MediaFile>\n')
        if (VidLength != '') and (VidLength != 0):
            f.write('      <Length>\n')
            f.write('        %s\n' % VidLength)
            f.write('      </Length>\n')
        if Offset != 0:
            f.write('      <Offset>\n')
            f.write('        %s\n' % Offset)
            f.write('      </Offset>\n')
        f.write('      <Audio>\n')
        f.write('        %s\n' % Audio)
        f.write('      </Audio>\n')
        f.write('    </AddVid>\n')

    def WriteSnapshotRec(self, f, snapshotRec):
        (SnapshotNum, SnapshotID, CollectNum, ImageFile, ImageScale, ImageCoordsX, ImageCoordsY, 
         ImageSizeW, ImageSizeH, EpisodeNum, TranscriptNum, SnapshotTimeCode, SnapshotDuration, 
         SnapshotComment, SortOrder) = snapshotRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            SnapshotID = DBInterface.ProcessDBDataForUTF8Encoding(SnapshotID)
            ImageFile = DBInterface.ProcessDBDataForUTF8Encoding(ImageFile)
            SnapshotComment = DBInterface.ProcessDBDataForUTF8Encoding(SnapshotComment)

        f.write('    <Snapshot>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % SnapshotNum)
        f.write('      </Num>\n')
        f.write('      <ID>\n')
        f.write('        %s\n' % self.Escape(SnapshotID.encode(EXPORT_ENCODING)))
        f.write('      </ID>\n')
        if (CollectNum != None) and (CollectNum > 0):
            f.write('      <CollectNum>\n')
            f.write('        %s\n' % CollectNum)
            f.write('      </CollectNum>\n')
        f.write('      <MediaFile>\n')
        f.write('        %s\n' % self.Escape(ImageFile.encode(EXPORT_ENCODING)))
        f.write('      </MediaFile>\n')
        f.write('      <ImageScale>\n')
        f.write('        %s\n' % ImageScale)
        f.write('      </ImageScale>\n')
        f.write('      <ImageCoordsX>\n')
        f.write('        %s\n' % ImageCoordsX)
        f.write('      </ImageCoordsX>\n')
        f.write('      <ImageCoordsY>\n')
        f.write('        %s\n' % ImageCoordsY)
        f.write('      </ImageCoordsY>\n')
        f.write('      <ImageSizeW>\n')
        f.write('        %s\n' % ImageSizeW)
        f.write('      </ImageSizeW>\n')
        f.write('      <ImageSizeH>\n')
        f.write('        %s\n' % ImageSizeH)
        f.write('      </ImageSizeH>\n')
        if (EpisodeNum != None) and (EpisodeNum > 0):
            f.write('      <EpisodeNum>\n')
            f.write('        %s\n' % EpisodeNum)
            f.write('      </EpisodeNum>\n')
        if (TranscriptNum != None) and (TranscriptNum > 0):
            f.write('      <TranscriptNum>\n')
            f.write('        %s\n' % TranscriptNum)
            f.write('      </TranscriptNum>\n')
        f.write('      <SnapshotTimeCode>\n')
        f.write('        %s\n' % SnapshotTimeCode)
        f.write('      </SnapshotTimeCode>\n')
        f.write('      <SnapshotDuration>\n')
        f.write('        %s\n' % SnapshotDuration)
        f.write('      </SnapshotDuration>\n')
        if SnapshotComment != '':
            f.write('      <Comment>\n')
            f.write('        %s\n' % self.Escape(SnapshotComment.encode(EXPORT_ENCODING)))
            f.write('      </Comment>\n')
        if (SortOrder != '') and (SortOrder != 0):
            f.write('      <SortOrder>\n')
            f.write('        %s\n' % SortOrder)
            f.write('      </SortOrder>\n')
        f.write('    </Snapshot>\n')

    def WriteTranscriptRec(self, f, progress, transcriptRec):
        (TranscriptNum, TranscriptID, EpisodeNum, SourceTranscriptNum, ClipNum, SortOrder, Transcriber,
         ClipStart, ClipStop, Comment, MinTranscriptWidth, RTFText) = transcriptRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            TranscriptID = DBInterface.ProcessDBDataForUTF8Encoding(TranscriptID)
            Transcriber = DBInterface.ProcessDBDataForUTF8Encoding(Transcriber)
            Comment = DBInterface.ProcessDBDataForUTF8Encoding(Comment)

        f.write('    <Transcript>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % TranscriptNum)
        f.write('      </Num>\n')
        if TranscriptID != '':
            f.write('      <ID>\n')
            f.write('        %s\n' % self.Escape(TranscriptID.encode(EXPORT_ENCODING)))

            if DEBUG:
                try:
                    print "Transcript ID ", TranscriptID
                except:
                    print "Transcript Number ", TranscriptNum
                
            f.write('      </ID>\n')
        if (EpisodeNum != '') and (EpisodeNum != 0):
            f.write('      <EpisodeNum>\n')
            f.write('        %s\n' % EpisodeNum)
            f.write('      </EpisodeNum>\n')
        if (SourceTranscriptNum != '') and (SourceTranscriptNum != 0):
            f.write('      <TranscriptNum>\n')
            f.write('        %s\n' % SourceTranscriptNum)
            f.write('      </TranscriptNum>\n')
        if (ClipNum != '') and (ClipNum != 0):
            f.write('      <ClipNum>\n')
            f.write('        %s\n' % ClipNum)
            f.write('      </ClipNum>\n')
        if (SortOrder != None) and (SortOrder != 0):
            f.write('      <SortOrder>\n')
            f.write('        %s\n' % SortOrder)
            f.write('      </SortOrder>\n')
        if (Transcriber != None) and (Transcriber != ''):
            f.write('      <Transcriber>\n')
            f.write('        %s\n' % self.Escape(Transcriber.encode(EXPORT_ENCODING)))
            f.write('      </Transcriber>\n')
        if (ClipStart != None):
            f.write('      <ClipStart>\n')
            f.write('        %s\n' % ClipStart)
            f.write('      </ClipStart>\n')
        if (ClipStop != None):
            f.write('      <ClipStop>\n')
            f.write('        %s\n' % ClipStop)
            f.write('      </ClipStop>\n')
        if (Comment != None) and (Comment != ''):
            f.write('      <Comment>\n')
            f.write('        %s\n' % self.Escape(Comment.encode(EXPORT_ENCODING)))
            f.write('      </Comment>\n')
        if (MinTranscriptWidth != None) and (MinTranscriptWidth != '') and (MinTranscriptWidth > 0):
            f.write('      <MinTranscriptWidth>\n')
            f.write('        %s\n' % MinTranscriptWidth)
            f.write('      </MinTranscriptWidth>\n')
        if RTFText != '':
            # Extract the RTF Text from the DB's array structure, if needed
            # Okay, this isn't so straight-forward any more.
            # With MySQL for Python 0.9.x, RTFText is of type str.
            # With MySQL for Python 1.2.0, RTFText is of type array.  It could then either be a
            # character string (typecode == 'c') or a unicode string (typecode == 'u'), which then
            # need to be interpreted differently.
            if type(RTFText).__name__ == 'array':
                if RTFText.typecode == 'u':
                    RTFText = RTFText.tounicode()
                else:
                    RTFText = RTFText.tostring()
            elif isinstance(RTFText, unicode):

                RTFText = RTFText.encode('utf8')
                
            f.write('      <RTFText>\n')

            if DEBUG2:
                print "type(RTFText) =", type(RTFText),

            # Determine if we have RTF Text, a pickled wxSTC Object, or XML Text
            if (type(RTFText).__name__ != 'NoneType') and (len(RTFText) > 5) and (RTFText[:5].upper() == '<?XML'):

                if DEBUG2:
                    print "XML Type"
                    
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt1 = unicode(_('Writing Transcript Records  (This will seem slow because of the size of the Transcript Records.)'), 'utf8')
                prompt2 = '\n' + unicode(_('Exporting %s'), 'utf8')

                progress.Update(56, prompt1 + prompt2 % TranscriptID)
                progress.Refresh()

                # If XML, we can just use it as is after we strip off the XML Header Line, which breaks XML.
                # Due to changes in the header across wxPython versions, we need to use a regular expression.
                # rtfData = RTFText[39:]

                # Match "<?xml" plus any number of characters that are not ">" plus ">" plus whitespace
                REGEX = "<\?xml[^>]*>[\s]*"
                regexCompiled = re.compile(REGEX)
                regexSearch = regexCompiled.search(RTFText)
                rtfData = RTFText[regexSearch.end():]

            elif (type(RTFText).__name__ != 'NoneType') and (len(RTFText) > 6) and (RTFText[:6].upper() == '{\\RTF1'):

                if DEBUG2:
                    print "RTF Type"
                    
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt1 = unicode(_('Writing Transcript Records  (This will seem slow because of the size of the Transcript Records.)'), 'utf8')
                prompt2 = unicode(_('\nConverting %s'), 'utf8')

                progress.Update(56, prompt1 + prompt2 % TranscriptID)

                # If RTF  ...
                # If we're using the RichTextCtrl ...
                if TransanaConstants.USESRTC:
                    # Load the RTF into the invisible RichTextCtrl
                    self.invisibleRTC.LoadRTFData(RTFText)
                    # Hide Time Code data that might not be hidden properly
                    self.invisibleRTC.HideTimeCodeData()
                    # ... and extract the XML for it.  This makes export slower, import faster, and means
                    # we ALWAYS have XML in the export file rather than a mix of XML and RTF.
                    rtfData = self.invisibleRTC.GetFormattedSelection('XML')

                    # If XML, we can just use it as is after we strip off the XML Header Line, which breaks XML.
                    # Due to changes in the header across wxPython versions, we need to use a regular expression.
                    # rtfData = rtfData[39:]

                    # Match "<?xml" plus any number of characters that are not ">" plus ">" plus whitespace
                    REGEX = "<\?xml[^>]*>[\s]*"
                    regexCompiled = re.compile(REGEX)
                    regexSearch = regexCompiled.search(rtfData)
                    rtfData = rtfData[regexSearch.end():]

                # If we're using the StyledTextCtrl ...
                else:
                    # ... then we just use the RTF as is.
                    rtfData = RTFText

            # If we have STC formatted data ...
            else:

                if DEBUG2:
                    print "STC Type"

                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt1 = unicode(_('Writing Transcript Records  (This will seem slow because of the size of the Transcript Records.)'), 'utf8')
                prompt2 = unicode(_('\nConverting %s'), 'utf8')

                progress.Update(56, prompt1 + prompt2 % TranscriptID)

                # unpickle the text and style info
                (bufferContents, specs, attrs) = pickle.loads(RTFText)
                # Clear the invisible STC
                self.invisibleSTC.ClearDoc()

                # DKW  Although these all are called in ClearDoc(), it appears necessary to repeat them
                #      here.  That's because ClearDoc() actually populates a few styles, and they interfere
                #      with the ones being brought in from the pickled RTFText.  Otherwise, Transcripts are 
                #      subtly changed on XML Export.  In particular, the Jeffersonian Symbols don't survive.
                self.invisibleSTC.StyleClearAll()
                self.invisibleSTC.style_specs = []
                self.invisibleSTC.style_attrs = []
                self.invisibleSTC.num_styles = 0

                # you have to apply the styles of the document in order
                # for the document to load properly.
                for x in specs:
                    self.invisibleSTC.GetStyleAccessor(x)

                # feed the data info invisibleSTC.
                self.invisibleSTC.AddStyledText(bufferContents)
                # extract the data as RTF.
                rtfData = self.invisibleSTC.GetRTFBuffer()

                # If we're using the RichTextCtrl ...
                if TransanaConstants.USESRTC:
                    # Load the STC-converted-to-RTF data into the invisible RichTextCtrl
                    self.invisibleRTC.LoadRTFData(rtfData)
                    # ... and extract the XML for it.  This makes export slower, import faster, and means
                    # we ALWAYS have XML in the export file rather than a mix of XML and RTF.
                    rtfData = self.invisibleRTC.GetFormattedSelection('XML')

                    # If XML, we can just use it as is after we strip off the XML Header Line, which breaks XML.
                    # Due to changes in the header across wxPython versions, we need to use a regular expression.
                    # rtfData = rtfData[39:]

                    # Match "<?xml" plus any number of characters that are not ">" plus ">" plus whitespace
                    REGEX = "<\?xml[^>]*>[\s]*"
                    regexCompiled = re.compile(REGEX)
                    regexSearch = regexCompiled.search(rtfData)
                    rtfData = rtfData[regexSearch.end():]

            # now simply write the RTF data to the file.  (This does NOT need to be encoded, as the RTF already is!)
            # (but check to make sure there's actually RTF data there!)
            if rtfData != None:
                f.write('%s' % rtfData.rstrip())
            # ... add an extra line break here!
            f.write('\n')
            f.write('      </RTFText>\n')

            # On exporting old, huge databases, you can end up with hundreds of "Converting RTF" popups.
            # This hopefully will allow them to close properly rather than building up.
            wx.YieldIfNeeded()
            
        f.write('    </Transcript>\n')

    def WriteKeywordRec(self, f, keywordRec):
        (KeywordGroup, Keyword, Definition, LineColorName, LineColorDef, DrawMode, LineWidth, LineStyle) = keywordRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            KeywordGroup = DBInterface.ProcessDBDataForUTF8Encoding(KeywordGroup)
            Keyword = DBInterface.ProcessDBDataForUTF8Encoding(Keyword)
            Definition = DBInterface.ProcessDBDataForUTF8Encoding(Definition)

        f.write('    <KeywordRec>\n')
        f.write('      <KeywordGroup>\n')
        f.write('        %s\n' % self.Escape(KeywordGroup.encode(EXPORT_ENCODING)))
        f.write('      </KeywordGroup>\n')
        f.write('      <Keyword>\n')
        f.write('        %s\n' % self.Escape(Keyword.encode(EXPORT_ENCODING)))
        f.write('      </Keyword>\n')
        if Definition != '':
            f.write('      <Definition>\n')
            # Okay, this isn't so straight-forward any more.
            # With Transana 2.30, Definition becomes a BLOB of type array.  It could then either be a
            # character string (typecode == 'c') or a unicode string (typecode == 'u'), which then
            # need to be interpreted differently.
            if type(Definition).__name__ == 'array':
                if (Definition.typecode == 'u'):
                    Definition = Definition.tounicode()
                else:
                    Definition = Definition.tostring()
                    if ('unicode' in wx.PlatformInfo):
                        try:
                            Definition = unicode(Definition, EXPORT_ENCODING)
                        except UnicodeDecodeError, e:
                            Definition = unicode(Definition.decode(TransanaGlobal.encoding))
            elif isinstance(Definition, str):
                Definition = DBInterface.ProcessDBDataForUTF8Encoding(Definition)
            f.write('        %s\n' % self.Escape(Definition.encode(EXPORT_ENCODING)))
            f.write('      </Definition>\n')
        if LineColorName != '':
            f.write('      <ColorName>\n')
            f.write('        %s\n' % self.Escape(LineColorName.encode(EXPORT_ENCODING)))
            f.write('      </ColorName>\n')
        if LineColorDef != '':
            f.write('      <ColorDef>\n')
            f.write('        %s\n' % LineColorDef)
            f.write('      </ColorDef>\n')
        if DrawMode != '':
            f.write('      <DrawMode>\n')
            f.write('        %s\n' % DrawMode)
            f.write('      </DrawMode>\n')
        if LineWidth > 0:
            f.write('      <LineWidth>\n')
            f.write('        %s\n' % LineWidth)
            f.write('      </LineWidth>\n')
        if LineStyle != '':
            f.write('      <LineStyle>\n')
            f.write('        %s\n' % LineStyle)
            f.write('      </LineStyle>\n')
        f.write('    </KeywordRec>\n')

    def WriteClipKeywordRec(self, f, clipKeywordRec):
        (EpisodeNum, DocumentNum, ClipNum, QuoteNum, SnapshotNum, KeywordGroup, Keyword, Example) = clipKeywordRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            KeywordGroup = DBInterface.ProcessDBDataForUTF8Encoding(KeywordGroup)
            Keyword = DBInterface.ProcessDBDataForUTF8Encoding(Keyword)

        f.write('    <ClipKeyword>\n')
        if (EpisodeNum != '') and (EpisodeNum != 0):
            f.write('      <EpisodeNum>\n')
            f.write('        %s\n' % EpisodeNum)
            f.write('      </EpisodeNum>\n')
        if (DocumentNum != '') and (DocumentNum != 0):
            f.write('      <DocumentNum>\n')
            f.write('        %s\n' % DocumentNum)
            f.write('      </DocumentNum>\n')
        if (ClipNum != '') and (ClipNum != 0):
            f.write('      <ClipNum>\n')
            f.write('        %s\n' % ClipNum)
            f.write('      </ClipNum>\n')
        if (QuoteNum != '') and (QuoteNum != 0):
            f.write('      <QuoteNum>\n')
            f.write('        %s\n' % QuoteNum)
            f.write('      </QuoteNum>\n')
        if (SnapshotNum != '') and (SnapshotNum != 0):
            f.write('      <SnapshotNum>\n')
            f.write('        %s\n' % SnapshotNum)
            f.write('      </SnapshotNum>\n')
        f.write('      <KeywordGroup>\n')
        f.write('        %s\n' % self.Escape(KeywordGroup.encode(EXPORT_ENCODING)))
        f.write('      </KeywordGroup>\n')
        f.write('      <Keyword>\n')
        f.write('        %s\n' % self.Escape(Keyword.encode(EXPORT_ENCODING)))
        f.write('      </Keyword>\n')
        if (Example != '') and (Example != 0) and (Example != u'0'):
            f.write('      <Example>\n')
            f.write('        %s\n' % Example)
            f.write('      </Example>\n')
        f.write('    </ClipKeyword>\n')

    def WriteSnapshotKeywordRec(self, f, snapshotKeywordRec):
        (SnapshotNum, KeywordGroup, Keyword, x1, y1, x2, y2, visible) = snapshotKeywordRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            KeywordGroup = DBInterface.ProcessDBDataForUTF8Encoding(KeywordGroup)
            Keyword = DBInterface.ProcessDBDataForUTF8Encoding(Keyword)

        f.write('    <SnapshotKeyword>\n')
        f.write('      <SnapshotNum>\n')
        f.write('        %s\n' % SnapshotNum)
        f.write('      </SnapshotNum>\n')
        f.write('      <KeywordGroup>\n')
        f.write('        %s\n' % self.Escape(KeywordGroup.encode(EXPORT_ENCODING)))
        f.write('      </KeywordGroup>\n')
        f.write('      <Keyword>\n')
        f.write('        %s\n' % self.Escape(Keyword.encode(EXPORT_ENCODING)))
        f.write('      </Keyword>\n')
        f.write('      <X1>\n')
        f.write('        %s\n' % x1)
        f.write('      </X1>\n')
        f.write('      <Y1>\n')
        f.write('        %s\n' % y1)
        f.write('      </Y1>\n')
        f.write('      <X2>\n')
        f.write('        %s\n' % x2)
        f.write('      </X2>\n')
        f.write('      <Y2>\n')
        f.write('        %s\n' % y2)
        f.write('      </Y2>\n')
        f.write('      <Visible>\n')
        f.write('        %s\n' % visible)
        f.write('      </Visible>\n')
        f.write('    </SnapshotKeyword>\n')

    def WriteSnapshotKeywordStyleRec(self, f, snapshotKeywordStyleRec):
        (SnapshotNum, KeywordGroup, Keyword, DrawMode, LineColorName, LineColorDef, LineWidth, LineStyle) = snapshotKeywordStyleRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            KeywordGroup = DBInterface.ProcessDBDataForUTF8Encoding(KeywordGroup)
            Keyword = DBInterface.ProcessDBDataForUTF8Encoding(Keyword)
            LineColorName = DBInterface.ProcessDBDataForUTF8Encoding(LineColorName)

        f.write('    <SnapshotKeywordStyle>\n')
        f.write('      <SnapshotNum>\n')
        f.write('        %s\n' % SnapshotNum)
        f.write('      </SnapshotNum>\n')
        f.write('      <KeywordGroup>\n')
        f.write('        %s\n' % self.Escape(KeywordGroup.encode(EXPORT_ENCODING)))
        f.write('      </KeywordGroup>\n')
        f.write('      <Keyword>\n')
        f.write('        %s\n' % self.Escape(Keyword.encode(EXPORT_ENCODING)))
        f.write('      </Keyword>\n')
        f.write('      <DrawMode>\n')
        f.write('        %s\n' % DrawMode)
        f.write('      </DrawMode>\n')
        f.write('      <ColorName>\n')
        f.write('        %s\n' % self.Escape(LineColorName.encode(EXPORT_ENCODING)))
        f.write('      </ColorName>\n')
        f.write('      <ColorDef>\n')
        f.write('        %s\n' % LineColorDef)
        f.write('      </ColorDef>\n')
        f.write('      <LineWidth>\n')
        f.write('        %s\n' % LineWidth)
        f.write('      </LineWidth>\n')
        f.write('      <LineStyle>\n')
        f.write('        %s\n' % LineStyle)
        f.write('      </LineStyle>\n')
        f.write('    </SnapshotKeywordStyle>\n')

    def WriteNoteRec(self, f, noteRec):
        (NoteNum, NoteID, SeriesNum, EpisodeNum, CollectNum, ClipNum, SnapshotNum, DocumentNum, QuoteNum, TranscriptNum, \
         NoteTaker, NoteText) = noteRec

        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            NoteID = DBInterface.ProcessDBDataForUTF8Encoding(NoteID)
            NoteTaker = DBInterface.ProcessDBDataForUTF8Encoding(NoteTaker)
            NoteText = DBInterface.ProcessDBDataForUTF8Encoding(NoteText)

        f.write('    <Note>\n')
        f.write('      <Num>\n')
        f.write('        %s\n' % NoteNum)
        f.write('      </Num>\n')
        f.write('      <ID>\n')
        f.write('        %s\n' % self.Escape(NoteID.encode(EXPORT_ENCODING)))
        f.write('      </ID>\n')
        # Note Library Numbers could be None instead of 0.
        if (SeriesNum != 0) and (SeriesNum != None):
            f.write('      <SeriesNum>\n')
            f.write('        %s\n' % SeriesNum)
            f.write('      </SeriesNum>\n')
        # Note Episode Numbers could be None instead of 0.
        if (EpisodeNum != 0) and (EpisodeNum != None):
            f.write('      <EpisodeNum>\n')
            f.write('        %s\n' % EpisodeNum)
            f.write('      </EpisodeNum>\n')
        # Note Collection Numbers could be None instead of 0.
        if (CollectNum != 0) and (CollectNum != None):
            f.write('      <CollectNum>\n')
            f.write('        %s\n' % CollectNum)
            f.write('      </CollectNum>\n')
        # Note Clip Numbers could be None instead of 0.
        if (ClipNum != 0) and (ClipNum != None):
            f.write('      <ClipNum>\n')
            f.write('        %s\n' % ClipNum)
            f.write('      </ClipNum>\n')
        # Note Document Numbers could be None instead of 0.
        if (DocumentNum != 0) and (DocumentNum != None):
            f.write('      <DocumentNum>\n')
            f.write('        %s\n' % DocumentNum)
            f.write('      </DocumentNum>\n')
        # Note Quote Numbers could be None instead of 0.
        if (QuoteNum != 0) and (QuoteNum != None):
            f.write('      <QuoteNum>\n')
            f.write('        %s\n' % QuoteNum)
            f.write('      </QuoteNum>\n')
        # Note Snapshot Numbers could be None instead of 0.
        if (SnapshotNum != 0) and (SnapshotNum != None):
            f.write('      <SnapshotNum>\n')
            f.write('        %s\n' % SnapshotNum)
            f.write('      </SnapshotNum>\n')
        # Note Transcript Numbers could be None instead of 0.
        if (TranscriptNum != 0) and (TranscriptNum != None):
            f.write('      <TranscriptNum>\n')
            f.write('        %s\n' % TranscriptNum)
            f.write('      </TranscriptNum>\n')
        if NoteTaker != '':
            f.write('      <NoteTaker>\n')
            f.write('        %s\n' % self.Escape(NoteTaker.encode(EXPORT_ENCODING)))
            f.write('      </NoteTaker>\n')
        if NoteText != '':
            # Okay, this isn't so straight-forward any more.
            # With MySQL for Python 0.9.x, NoteText is of type str.
            # With MySQL for Python 1.2.0, NoteText is of type array.  It could then either be a
            # character string (typecode == 'c') or a unicode string (typecode == 'u'), which then
            # need to be interpreted differently.
            # This is because NoteText is a BLOB field in the database.
            if type(NoteText).__name__ == 'array':
                if (NoteText.typecode == 'u'):
                    NoteText = NoteText.tounicode()
                else:
                    NoteText = NoteText.tostring()
                    if ('unicode' in wx.PlatformInfo):
                        try:
                            NoteText = unicode(NoteText, EXPORT_ENCODING)
                        except UnicodeDecodeError, e:
                            NoteText = unicode(NoteText.decode(TransanaGlobal.encoding))
            # If the NoteText is a string ...
            elif type(NoteText).__name__ == 'str':
                # ... we need to convert it to a Unicode object so the conversion below will work!
                NoteText = unicode(NoteText, TransanaGlobal.encoding)
            f.write('      <NoteText>\n')
            # If note text is found ...
            if NoteText != None:
                # ... add it to the output file ...
                f.write('%s\n' % self.Escape(NoteText.encode(EXPORT_ENCODING)))
            # ... but if NO note text is found ...
            else:
                # ... explicitly note that!  (This was crashing the export/import process.)
                f.write('%s\n' % _('(No Note Text found.)').encode(EXPORT_ENCODING))
            f.write('      </NoteText>\n')
        f.write('    </Note>\n')

    def WriteFilterRec(self, f, filterRec):
        (ReportType, ReportScope, ConfigName, FilterDataType, FilterData) = filterRec

        if DEBUG:
            print "FilterData rec:", ReportType, ReportScope, ConfigName.encode('utf8'), FilterDataType,
            if FilterData == None:
                print "FilterData == None"
            else:
                print
                
        # If we're encoding things (for 1.8 format) and we're using MySQL ...
        if ENCODE_PROPERLY and TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
            # ... encode the value
            ConfigName = DBInterface.ProcessDBDataForUTF8Encoding(ConfigName)

        if FilterData != None:
            f.write('    <Filter>\n')
            f.write('      <ReportType>\n')
            f.write('        %s\n' % ReportType)
            f.write('      </ReportType>\n')
            f.write('      <ReportScope>\n')
            f.write('        %s\n' % ReportScope)
            f.write('      </ReportScope>\n')
            f.write('      <ConfigName>\n')
            f.write('        %s\n' % self.Escape(ConfigName.encode(EXPORT_ENCODING)))
            f.write('      </ConfigName>\n')
            f.write('      <FilterDataType>\n')
            f.write('        %s\n' % FilterDataType)
            f.write('      </FilterDataType>\n')

            # The following code was used for TransanaXML version 1.2.  If the FilterData included unicode
            # characters, particularly Russian or Chinese characters, the export would fail
            
            # FilterData is a BLOB field in the database.  Therefore, it's probably of type array, and needs to be converted.
            # if type(FilterData).__name__ == 'array':
            #     if (FilterData.typecode == 'u'):
            #         FilterData = FilterData.tounicode()
            #     else:
            #         FilterData = FilterData.tostring()
            #         if ('unicode' in wx.PlatformInfo):
            #             FilterData = unicode(FilterData, EXPORT_ENCODING)
            # f.write('      <FilterData>\n')
            # f.write('        %s\n' % FilterData.encode(EXPORT_ENCODING))
            # f.write('      </FilterData>\n')

            # For TransanaXML version 1.3, we fix problems with encoding in the Filter Data.  We're not just
            # unpickling and repickling the data.  We're converting it to a form that's more friendly for output.

            # For FilterDataTypes for Episodes (1), Clips (2), Keywords(3), Keyword Groups (5), Transcripts (6),
            # Collection (7), Notes(8), Snapshots (18), Documents (19), or Quotes (20), we have LIST data to process.
            if FilterDataType in [1, 2, 3, 5, 6, 7, 8, 18, 19, 20]:
                # Initialize the fileterDataList to None, in case there's an unpickling error
                filterDataList = None
                # Because of the BLOB size problem, some filter data may not unpickle correctly.  Better trap it.
                try:
                    # MySQL for Python often saves the data as an array ...
                    if type(FilterData).__name__ == 'array':
                        # ... so convert it to a string and un-pickle it
                        filterDataList = cPickle.loads(FilterData.tostring())
                    # If we have a unicode object ...
                    elif isinstance(FilterData, unicode):
                        # ... unpickle it while encoding it
                        filterDataList = cPickle.loads(FilterData.encode(EXPORT_ENCODING))
                    # If it's not an array ...
                    else:
                        # ... just un-pickle the data.
                        filterDataList = cPickle.loads(FilterData)
                # If we get an UnpicklingError exception ...
                except cPickle.UnpicklingError, e:
                    # Construct an error message that tells the user exactly what filter is causing
                    # the problem so it can be fixed.
                    errorMsg = _("There is a problem with a Filter Data record.") + "\n\n"
                    # Tell the user what type of report is having a problem.
                    errorMsg += "  " + _("ReportType:      %s")
                    if ReportType == 1:
                        errorMsg += "   (" + _("Keyword Map") + ")"
                    elif ReportType == 2:
                        errorMsg += "   (" + _("Keyword Visualization") + ")"
                    elif ReportType == 3:
                        errorMsg += "   (" + _("Episode Analytic Data Export") + ")"
                    elif ReportType == 4:
                        errorMsg += "   (" + _("Collection Analytic Data Export") + ")"
                    elif ReportType == 5:
                        errorMsg += "   (" + _("Library Keyword Sequence Map") + ")"
                    elif ReportType == 6:
                        errorMsg += "   (" + _("Library Keyword Bar Graph") + ")"
                    elif ReportType == 7:
                        errorMsg += "   (" + _("Library Keyword Percentage Map") + ")"
                    elif ReportType == 8:
                        errorMsg += "   (" + _("Episode  Data Coder Reliability Export") + ")"
                    elif ReportType == 9:
                        errorMsg += "   (" + _("Keyword Summary Report") + ")"
                    elif ReportType == 10:
                        errorMsg += "   (" + _("Library Report") + ")"
                    elif ReportType == 11:
                        errorMsg += "   (" + _("Episode Report") + ")"
                    elif ReportType == 12:
                        errorMsg += "   (" + _("Collection Report") + ")"
                    elif ReportType == 13:
                        errorMsg += "   (" + _("Notes Report") + ")"
                    elif ReportType == 14:
                        errorMsg += "   (" + _("Library Analytic Data Export") + ")"
                    elif ReportType == 15:
                        errorMsg += "   (" + _("Saved Search") + ")"
                    elif ReportType == 16:
                        errorMsg += "   (" + _("Collection Keyword Map") + ")"
                    elif ReportType == 17:
                        errorMsg += "   (" + _("Document Keyword Map") + ")"
                    elif ReportType == 18:
                        errorMsg += "   (" + _("Document Keyword Visualization") + ")"
                    elif ReportType == 19:
                        errorMsg += "   (" + _("Document Report") + ")"
                    elif ReportType == 20:
                        errorMsg += "   (" + _("Document Analytic Data Export") + ")"
                    errorMsg += "\n"
                    # Tell the user which data object is having a problem.
                    errorMsg += "  " + _("ReportScope:     %s")
                    if ReportType in [5, 6, 7, 10, 14]:
                        import Library
                        tempLibrary = Library.Library(ReportScope)
                        errorMsg += "  (" + _('Library') + ' "%s")' % tempLibrary.id
                    elif ReportType in [1, 2, 3, 8, 11]:
                        import Episode
                        tempEpisode = Episode.Episode(ReportScope)
                        errorMsg += "  (" + _("Episode") + ' "%s")' % tempEpisode.id
                    elif ReportType in [4, 12]:
                        if ReportScope > 0:
                            import Collection
                            tempCollection = Collection.Collection(ReportScope)
                            errorMsg += "  (" + _("Collection") + ' "%s")' % tempCollection.id
                    elif ReportType in [13]:
                        if ReportScope == 1:
                            errorMsg += "  " + _("(All Notes)")
                        elif ReportScope == 2:
                            errorMsg += "  " + _("(Library Notes)")
                        elif ReportScope == 3:
                            errorMsg += "  " + _("(Episode Notes)")
                        elif ReportScope == 4:
                            errorMsg += "  " + _("(Transcript Notes)")
                        elif ReportScope == 5:
                            errorMsg += "  " + _("(Collection Notes)")
                        elif ReportScope == 6:
                            errorMsg += "  " + _("(Clip Notes)")
                        elif ReportScope == 7:
                            errorMsg += "  " + _("(Snapshot Notes)")
                        elif ReportScope == 8:
                            errorMsg += "  " + _("(Document Notes)")
                        elif ReportScope == 9:
                            errorMsg += "  " + _("(Quote Notes)")
                    elif ReportType in [17, 18, 19, 20]:
                        if ReportScope > 0:
                            import Document
                            tempDocument = Document.Document(ReportScope)
                            errorMsg += "  (" + _("Document") + ' "%s")' % tempDocument.id
                    errorMsg += "\n"
                    # Tell the user which config file is having the problem.
                    errorMsg += "  " + _("ConfigName:      %s") + "\n"
                    # Tell the user which part of the filter is having the problem.
                    errorMsg += "  " + _("FilterDataType:  %s")
                    if FilterDataType == 1:
                        errorMsg += "   (" + _("Episodes") + ")"
                    elif FilterDataType == 2:
                        errorMsg += "   (" + _("Clips") + ")"
                    elif FilterDataType == 3:
                        errorMsg += "   (" + _("Keywords") + ")"
                    elif FilterDataType == 4:
                        errorMsg += "   (" + _("Keyword Colors") + ")"
                    elif FilterDataType == 8:
                        errorMsg += "   (" + _("Notes") + ")"
                    elif FilterDataType == 18:
                        errorMsg += "   (" + _("Snapshots") + ")"
                    elif FilterDataType == 19:
                        errorMsg += "   (" + _("Documents") + ")"
                    elif FilterDataType == 20:
                        errorMsg += "   (" + _("Quotes") + ")"
                    errorMsg += "\n\n"
                    errorMsg += _("This Filter Configuration needs to be corrected or the data record needs to be removed from the export file.")
                    # Convert the error message to Unicode
                    if ('unicode' in wx.PlatformInfo) and (type(errorMsg) == type('')):
                        errorMsg = unicode(errorMsg, 'utf8')
                    # Gather the data to be put into the error message
                    errorMsgData = (ReportType, ReportScope, ConfigName, FilterDataType)
                    # Finally display the error message
                    errorDlg = Dialogs.ErrorDialog(self, errorMsg % errorMsgData)
                    errorDlg.ShowModal()
                    errorDlg.Destroy()

            # For FilterDataType for Keyword Colors (4), we have DICTIONARY data that's already been encoded.
            # For FilterDataType for Notes (8), we have a LIST that's already been encoded.
            elif FilterDataType in [4, 8]:
                # MySQL for Python often saves the data as an array ...
                if type(FilterData).__name__ == 'array':
                    # ... so convert it to a string
                    filterDataList = FilterData.tostring()
                else:
                    filterDataList = FilterData

            # Other filter data types (0, and 9 though 17 so far as of 2.21) are of UNPICKLED, UNENCODED data
            else:
                # MySQL for Python often saves the data as an array ...
                if type(FilterData).__name__ == 'array':
                    # ... so convert it to a string
                    filterDataList = FilterData.tostring()
                # If we have a unicode object ...
                elif isinstance(FilterData, unicode):
                    # ... encode it
                    filterDataList = FilterData.encode(EXPORT_ENCODING)
                else:
                    filterDataList = FilterData

            # If we have data from FilterDataTypes 1, 2, 3, 5, 6, 7, 8, 18, 19, 20 ...
            if FilterDataType in [1, 2, 3, 5, 6, 7, 8, 18, 19, 20]:
                # ... we need to re-pickle the data
                FilterData = cPickle.dumps(filterDataList)
            # Otherwise ...
            else:
                # ... the data's ready for output.
                FilterData = filterDataList

            f.write('      <FilterData>\n')
            f.write('        %s\n' % self.Escape(FilterData))
            f.write('      </FilterData>\n')
            f.write('    </Filter>\n')

    def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        fs = wx.FileSelector(_("Select an XML file for export"),
                        TransanaGlobal.configData.videoPath,
                        "",
                        "", 
                        _("Transana-XML Files (*.tra)|*.tra|XML Files (*.xml)|*.xml|All files (*.*)|*.*"), 
                        wx.SAVE)
        # If user didn't cancel ..
        if fs != "":
            self.XMLFile.SetValue(fs)


# This simple derrived class let's the user drop files onto an edit box
class EditBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, editbox):
        wx.FileDropTarget.__init__(self)
        self.editbox = editbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        self.editbox.SetValue(files[0])
