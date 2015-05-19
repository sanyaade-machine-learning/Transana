# Copyright (C) 2003 - 2007 The Board of Regents of the University of Wisconsin System 
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

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Jonathan Beavers <jonathan.beavers@gmail.com>'

DEBUG = False
if DEBUG:
    print "XMLExport DEBUG is ON!!"

import wx

import Dialogs
import DBInterface
import TransanaGlobal
from RichTextEditCtrl import RichTextEditCtrl
import cPickle
import pickle
import os
import sys

class XMLExport(Dialogs.GenForm):
    """ This window displays a variety of GUI Widgets. """
    def __init__(self,parent,id,title):
        Dialogs.GenForm.__init__(self, parent, id, title, (550,150), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Export Database')
        # Define the minimum size for this dialog as the initial size
        self.SetSizeHints(550, 150)

	# Create an invisible instance of RichTextEditCtrl. This allows us to
	# get away with converting fastsaved documents to RTF behind the scenes.
	# Then we can simply pull the RTF data out of this object and write it
	# to the desired file.
	self.invisibleSTC = RichTextEditCtrl(self)
	self.invisibleSTC.Show(False)

        # Emport Message Layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.right.SameAs(self.panel, wx.Right, 10)
        lay.height.AsIs()
        # If the XML filename path is not empty, we need to tell the user.
        prompt = _('Please create an Transana XML File for export.')
        exportText = wx.StaticText(self.panel, -1, prompt)
        exportText.SetConstraints(lay)

        # XML Filename Layout
        lay = wx.LayoutConstraints()
        lay.top.Below(exportText, 10)
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

        self.XMLFile.SetFocus()

    def Export(self):
        # Set the encoding for export.
        # Use UTF-8 regardless of the current encoding for consistency in the Transana XML files
        EXPORT_ENCODING = 'utf8'
        # use the LONGEST title here!  That determines the size of the Dialog Box.
        progress = wx.ProgressDialog(_('Transana XML Export'), _('Exporting Transcript records (This may be slow because of the size of Transcript records.)'), style = wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)
        if progress.GetSize()[0] > 800:
            progress.SetSize((800, progress.GetSize()[1]))
            progress.Centre()

        db = DBInterface.get_db()
       
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
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n');
            f.write('<!DOCTYPE TransanaData [\n');
            f.write('  <!ELEMENT TransanaXMLVersion (#PCDATA)>\n');
            f.write('  <!ELEMENT SeriesFile (Series)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT Num (#PCDATA)>\n');
            f.write('  <!ELEMENT ID (#PCDATA)>\n');
            f.write('  <!ELEMENT Comment (#PCDATA)>\n');
            f.write('  <!ELEMENT Owner (#PCDATA)>\n');
            f.write('  <!ELEMENT DefaultKeywordGroup (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT Series (#PCDATA|Num|ID|Comment|Owner|DefaultKeywordGroup)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT EpisodeFile (Episode)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT SeriesNum (#PCDATA)>\n');
            f.write('  <!ELEMENT Date (#PCDATA)>\n');
            f.write('  <!ELEMENT MediaFile (#PCDATA)>\n');
            f.write('  <!ELEMENT Length (#PCDATA)>\n');
            f.write('  <!ELEMENT Comment (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT Episode (#PCDATA|Num|ID|SeriesNum|Date|MediaFile|Length|Comment)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT CoreDataFile (CoreData)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT Title (#PCDATA)>\n');
            f.write('  <!ELEMENT Creator (#PCDATA)>\n');
            f.write('  <!ELEMENT Subject (#PCDATA)>\n');
            f.write('  <!ELEMENT Description (#PCDATA)>\n');
            f.write('  <!ELEMENT Publisher (#PCDATA)>\n');
            f.write('  <!ELEMENT Contributor (#PCDATA)>\n');
            f.write('  <!ELEMENT Type (#PCDATA)>\n');
            f.write('  <!ELEMENT Format (#PCDATA)>\n');
            f.write('  <!ELEMENT Source (#PCDATA)>\n');
            f.write('  <!ELEMENT Language (#PCDATA)>\n');
            f.write('  <!ELEMENT Relation (#PCDATA)>\n');
            f.write('  <!ELEMENT Coverage (#PCDATA)>\n');
            f.write('  <!ELEMENT Rights (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT CoreData (#PCDATA|Num|ID|Title|Creator|Subject|Description|Publisher|Contributor|Date|Type|Format|Source|Language|Relation|Coverage|Rights)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT TranscriptFile (Transcript)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT EpisodeNum (#PCDATA)>\n');
            f.write('  <!ELEMENT ClipNum (#PCDATA)>\n');
            f.write('  <!ELEMENT Transcriber (#PCDATA)>\n');
            f.write('  <!ELEMENT RTFText (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT Transcript (#PCDATA|Num|ID|EpisodeNum|ClipNum|Transcriber|Comment|RTFText)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT CollectionFile (Collection)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT ParentCollectNum (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT Collection (#PCDATA|Num|ID|ParentCollectNum|Comment|Owner|DefaultKeywordGroup)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT ClipFile (Clip)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT CollectNum (#PCDATA)>\n');
            f.write('  <!ELEMENT TranscriptNum (#PCDATA)>\n');
            f.write('  <!ELEMENT ClipStart (#PCDATA)>\n');
            f.write('  <!ELEMENT ClipStop (#PCDATA)>\n');
            f.write('  <!ELEMENT SortOrder (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT Clip (#PCDATA|Num|ID|CollectNum|TranscriptNum|ClipStart|ClipStop|Comment|SortOrder)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT KeywordFile (KeywordRec)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT KeywordGroup (#PCDATA)>\n');
            f.write('  <!ELEMENT Keyword (#PCDATA)>\n');
            f.write('  <!ELEMENT Definition (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT KeywordRec (#PCDATA|KeywordGroup|Keyword|Definition)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT ClipKeywordFile (ClipKeyword)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT Example (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT ClipKeyword (#PCDATA|EpisodeNum|ClipNum|KeywordGroup|Keyword|Example)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT NoteFile (Note)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT NoteTaker (#PCDATA)>\n');
            f.write('  <!ELEMENT NoteText (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT Note (#PCDATA|Num|ID|SeriesNum|EpisodeNum|CollectNum|ClipNum|TranscriptNum|NoteTaker|NoteText)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT FilterFile (Filter)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT ReportType (#PCDATA)>\n');
            f.write('  <!ELEMENT ReportScope (#PCDATA)>\n');
            f.write('  <!ELEMENT ConfigName (#PCDATA)>\n');
            f.write('  <!ELEMENT FilterDataType (#PCDATA)>\n');
            f.write('  <!ELEMENT FilterData (#PCDATA)>\n');
            f.write('\n');
            f.write('  <!ELEMENT Filter (#PCDATA|ReportType|ReportScope|ConfigName|FilterDataType|FilterData)*>\n');
            f.write('\n');
            f.write('  <!ELEMENT Transana (#PCDATA|SeriesFile|EpisodeFile|CoreDataFile|TranscriptFile|CollectionFile|ClipFile|KeywordFile|ClipKeywordFile|NoteFile|FilterFile)*>\n');
            f.write(']>\n');
            f.write('\n');
            f.write('<Transana>\n');
            f.write('  <TransanaXMLVersion>\n');
            # Version 1.0 -- Original Transana XML for Transana 2.0 release
            # Version 1.1 -- Unicode encoding added to Transana XML for Transana 2.1 release
            # Version 1.2 -- Filter Table added to Transana XML for Transana 2.11 release
            # Version 1.3 -- FilterData handling changed to accomodate Unicode data for Transana 2.21 release
            f.write('    1.3\n');
            f.write('  </TransanaXMLVersion>\n');

            progress.Update(9, _('Writing Series Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT SeriesNum, SeriesID, SeriesComment, SeriesOwner, DefaultKeywordGroup FROM Series2'
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('  <SeriesFile>\n')
                for (SeriesNum, SeriesID, SeriesComment, SeriesOwner, DefaultKeywordGroup) in dbCursor.fetchall():
                    f.write('    <Series>\n')
                    f.write('      <Num>\n')
                    f.write('        %s\n' % SeriesNum)
                    f.write('      </Num>\n')
                    f.write('      <ID>\n')
                    f.write('        %s\n' % SeriesID.encode(EXPORT_ENCODING))
                    f.write('      </ID>\n')
                    if SeriesComment != '':
                        f.write('      <Comment>\n')
                        f.write('        %s\n' % SeriesComment.encode(EXPORT_ENCODING))
                        f.write('      </Comment>\n')
                    if SeriesOwner != '':
                        f.write('      <Owner>\n')
                        f.write('        %s\n' % SeriesOwner.encode(EXPORT_ENCODING))
                        f.write('      </Owner>\n')
                    if DefaultKeywordGroup != '':
                        f.write('      <DefaultKeywordGroup>\n')
                        f.write('        %s\n' % DefaultKeywordGroup.encode(EXPORT_ENCODING))
                        f.write('      </DefaultKeywordGroup>\n')
                    f.write('    </Series>\n')
                if dbCursor.rowcount > 0:
                    f.write('  </SeriesFile>\n')
                dbCursor.close()

            progress.Update(18, _('Writing Episode Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT EpisodeNum, EpisodeID, SeriesNum, TapingDate, MediaFile, EpLength, EpComment FROM Episodes2'
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('  <EpisodeFile>\n')
                for (EpisodeNum, EpisodeID, SeriesNum, TapingDate, MediaFile, EpLength, EpComment) in dbCursor.fetchall():
                    f.write('    <Episode>\n')
                    f.write('      <Num>\n')
                    f.write('        %s\n' % EpisodeNum)
                    f.write('      </Num>\n')
                    f.write('      <ID>\n')
                    f.write('        %s\n' % EpisodeID.encode(EXPORT_ENCODING))
                    f.write('      </ID>\n')
                    f.write('      <SeriesNum>\n')
                    f.write('        %s\n' % SeriesNum)
                    f.write('      </SeriesNum>\n')
                    if TapingDate != None:
                        f.write('      <Date>\n')
                        f.write('        %s\n' % TapingDate)
                        f.write('      </Date>\n')
                    f.write('      <MediaFile>\n')
                    f.write('        %s\n' % MediaFile.encode(EXPORT_ENCODING))
                    f.write('      </MediaFile>\n')
                    if EpLength != '':
                        f.write('      <Length>\n')
                        f.write('        %s\n' % EpLength)
                        f.write('      </Length>\n')
                    if EpComment != '':
                        f.write('      <Comment>\n')
                        f.write('        %s\n' % EpComment.encode(EXPORT_ENCODING))
                        f.write('      </Comment>\n')
                    f.write('    </Episode>\n')
                if dbCursor.rowcount > 0:
                    f.write('  </EpisodeFile>\n')
                dbCursor.close()

            progress.Update(27, _('Writing Core Data Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = """SELECT CoreDataNum, Identifier, Title, Creator, Subject, Description, Publisher,
                                    Contributor, DCDate, DCType, Format, Source, Language, Relation, Coverage, Rights
                                    FROM CoreData2"""
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('  <CoreDataFile>\n')
                for (CoreDataNum, Identifier, Title, Creator, Subject, Description, Publisher,
                     Contributor, DCDate, DCType, Format, Source, Language, Relation, Coverage, Rights) in dbCursor.fetchall():
                    f.write('    <CoreData>\n')
                    f.write('      <Num>\n')
                    f.write('        %s\n' % CoreDataNum)
                    f.write('      </Num>\n')
                    f.write('      <ID>\n')
                    f.write('        %s\n' % Identifier.encode(EXPORT_ENCODING))
                    f.write('      </ID>\n')
                    if Title != '':
                        f.write('      <Title>\n')
                        f.write('        %s\n' % Title.encode(EXPORT_ENCODING))
                        f.write('      </Title>\n')
                    if Creator != '':
                        f.write('      <Creator>\n')
                        f.write('        %s\n' % Creator.encode(EXPORT_ENCODING))
                        f.write('      </Creator>\n')
                    if Subject != '':
                        f.write('      <Subject>\n')
                        f.write('        %s\n' % Subject.encode(EXPORT_ENCODING))
                        f.write('      </Subject>\n')
                    if Description != '':
                        f.write('      <Description>\n')
                        f.write('        %s\n' % Description.encode(EXPORT_ENCODING))
                        f.write('      </Description>\n')
                    if Publisher != '':
                        f.write('      <Publisher>\n')
                        f.write('        %s\n' % Publisher.encode(EXPORT_ENCODING))
                        f.write('      </Publisher>\n')
                    if Contributor != '':
                        f.write('      <Contributor>\n')
                        f.write('        %s\n' % Contributor.encode(EXPORT_ENCODING))
                        f.write('      </Contributor>\n')
                    if DCDate != None:
                        f.write('      <Date>\n')
                        f.write('        %s/%s/%s\n' % (DCDate.month, DCDate.day, DCDate.year))
                        f.write('      </Date>\n')
                    if DCType != '':
                        f.write('      <Type>\n')
                        f.write('        %s\n' % DCType.encode(EXPORT_ENCODING))
                        f.write('      </Type>\n')
                    if Format != '':
                        f.write('      <Format>\n')
                        f.write('        %s\n' % Format.encode(EXPORT_ENCODING))
                        f.write('      </Format>\n')
                    if Source != '':
                        f.write('      <Source>\n')
                        f.write('        %s\n' % Source.encode(EXPORT_ENCODING))
                        f.write('      </Source>\n')
                    if Language != '':
                        f.write('      <Language>\n')
                        f.write('        %s\n' % Language.encode(EXPORT_ENCODING))
                        f.write('      </Language>\n')
                    if Relation != '':
                        f.write('      <Relation>\n')
                        f.write('        %s\n' % Relation.encode(EXPORT_ENCODING))
                        f.write('      </Relation>\n')
                    if Coverage != '':
                        f.write('      <Coverage>\n')
                        f.write('        %s\n' % Coverage.encode(EXPORT_ENCODING))
                        f.write('      </Coverage>\n')
                    if Rights != '':
                        f.write('      <Rights>\n')
                        f.write('        %s\n' % Rights.encode(EXPORT_ENCODING))
                        f.write('      </Rights>\n')
                    f.write('    </CoreData>\n')
                if dbCursor.rowcount > 0:
                    f.write('  </CoreDataFile>\n')
                dbCursor.close()

            progress.Update(36, _('Writing Collection Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT CollectNum, CollectID, ParentCollectNum, CollectComment, CollectOwner, DefaultKeywordGroup FROM Collections2'
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('  <CollectionFile>\n')
                for (CollectNum, CollectID, ParentCollectNum, CollectComment, CollectOwner, DefaultKeywordGroup) in dbCursor.fetchall():
                    f.write('    <Collection>\n')
                    f.write('      <Num>\n')
                    f.write('        %s\n' % CollectNum)
                    f.write('      </Num>\n')
                    f.write('      <ID>\n')
                    f.write('        %s\n' % CollectID.encode(EXPORT_ENCODING))
                    f.write('      </ID>\n')
                    if ParentCollectNum != '':
                        f.write('      <ParentCollectNum>\n')
                        f.write('        %s\n' % ParentCollectNum)
                        f.write('      </ParentCollectNum>\n')
                    if CollectComment != '':
                        f.write('      <Comment>\n')
                        f.write('        %s\n' % CollectComment.encode(EXPORT_ENCODING))
                        f.write('      </Comment>\n')
                    if CollectOwner != '':
                        f.write('      <Owner>\n')
                        f.write('        %s\n' % CollectOwner.encode(EXPORT_ENCODING))
                        f.write('      </Owner>\n')
                    if DefaultKeywordGroup != '':
                        f.write('      <DefaultKeywordGroup>\n')
                        f.write('        %s\n' % DefaultKeywordGroup.encode(EXPORT_ENCODING))
                        f.write('      </DefaultKeywordGroup>\n')
                    f.write('    </Collection>\n')
                if dbCursor.rowcount > 0:
                    f.write('  </CollectionFile>\n')
                dbCursor.close()

            progress.Update(45, _('Writing Clip Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT ClipNum, ClipID, CollectNum, EpisodeNum, TranscriptNum, MediaFile, ClipStart, ClipStop, ClipComment, SortOrder FROM Clips2'
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('  <ClipFile>\n')
                for (ClipNum, ClipID, CollectNum, EpisodeNum, TranscriptNum, MediaFile, ClipStart, ClipStop, ClipComment, SortOrder) in dbCursor.fetchall():
                    f.write('    <Clip>\n')
                    f.write('      <Num>\n')
                    f.write('        %s\n' % ClipNum)
                    f.write('      </Num>\n')
                    f.write('      <ID>\n')
                    f.write('        %s\n' % ClipID.encode(EXPORT_ENCODING))
                    f.write('      </ID>\n')
                    if CollectNum != None:
                        f.write('      <CollectNum>\n')
                        f.write('        %s\n' % CollectNum)
                        f.write('      </CollectNum>\n')
                    if EpisodeNum != None:
                        f.write('      <EpisodeNum>\n')
                        f.write('        %s\n' % EpisodeNum)
                        f.write('      </EpisodeNum>\n')
                    if TranscriptNum != None:
                        f.write('      <TranscriptNum>\n')
                        f.write('        %s\n' % TranscriptNum)
                        f.write('      </TranscriptNum>\n')
                    f.write('      <MediaFile>\n')
                    f.write('        %s\n' % MediaFile.encode(EXPORT_ENCODING))
                    f.write('      </MediaFile>\n')
                    f.write('      <ClipStart>\n')
                    f.write('        %s\n' % ClipStart)
                    f.write('      </ClipStart>\n')
                    f.write('      <ClipStop>\n')
                    f.write('        %s\n' % ClipStop)
                    f.write('      </ClipStop>\n')
                    if ClipComment != '':
                        f.write('      <Comment>\n')
                        f.write('        %s\n' % ClipComment.encode(EXPORT_ENCODING))
                        f.write('      </Comment>\n')
                    if SortOrder != '':
                        f.write('      <SortOrder>\n')
                        f.write('        %s\n' % SortOrder)
                        f.write('      </SortOrder>\n')
                    f.write('    </Clip>\n')
                if dbCursor.rowcount > 0:
                    f.write('  </ClipFile>\n')
                dbCursor.close()

            progress.Update(54, _('Writing Transcript Records  (This will seem slow because of the size of the Transcript Records.)'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT TranscriptNum, TranscriptID, EpisodeNum, ClipNum, Transcriber, Comment, RTFText FROM Transcripts2'

                if DEBUG:
                    print "Selecting Transcripts"
                    
                dbCursor.execute(SQLText)

                if DEBUG:
                    print "%d Transcripts selected." % dbCursor.rowcount
                    
                if dbCursor.rowcount > 0:
                    f.write('  <TranscriptFile>\n')
                for (TranscriptNum, TranscriptID, EpisodeNum, ClipNum, Transcriber, Comment, RTFText) in dbCursor.fetchall():

                    if DEBUG:
                        print "TranscriptNum =", TranscriptNum
                        
                    f.write('    <Transcript>\n')
                    f.write('      <Num>\n')
                    f.write('        %s\n' % TranscriptNum)
                    f.write('      </Num>\n')
                    if TranscriptID != '':
                        f.write('      <ID>\n')
                        f.write('        %s\n' % TranscriptID.encode(EXPORT_ENCODING))

                        if DEBUG:
                            try:
                                print "Transcript ID ", TranscriptID
                            except:
                                print "Transcript Number ", TranscriptNum
                            
                        f.write('      </ID>\n')
                    if EpisodeNum != '':
                        f.write('      <EpisodeNum>\n')
                        f.write('        %s\n' % EpisodeNum)
                        f.write('      </EpisodeNum>\n')
                    if ClipNum != '':
                        f.write('      <ClipNum>\n')
                        f.write('        %s\n' % ClipNum)
                        f.write('      </ClipNum>\n')
                    if (Transcriber != None) and (Transcriber != ''):
                        f.write('      <Transcriber>\n')
                        f.write('        %s\n' % Transcriber.encode(EXPORT_ENCODING))
                        f.write('      </Transcriber>\n')
                    if (Comment != None) and (Comment != ''):
                        f.write('      <Comment>\n')
                        f.write('        %s\n' % Comment.encode(EXPORT_ENCODING))
                        f.write('      </Comment>\n')
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
                        f.write('      <RTFText>\n')

                        if DEBUG:
                            print "type(RTFText) =", type(RTFText)
                            
                        # Determine if we have RTF Text or a pickled wxSTC Object
                        if (type(RTFText).__name__ != 'NoneType') and (len(RTFText) > 6) and (RTFText[:6].upper() != '{\\RTF1'):

                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt1 = unicode(_('Writing Transcript Records  (This will seem slow because of the size of the Transcript Records.)'), 'utf8')
                                prompt2 = unicode(_('\nConverting %s'), 'utf8')
                            else:
                                prompt1 = _('Writing Transcript Records  (This will seem slow because of the size of the Transcript Records.)')
                                prompt2 = _('\nConverting %s')

                            progress.Update(60, prompt1 + prompt2 % DBInterface.ProcessDBDataForUTF8Encoding(TranscriptID))

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

                            progress.Update(60, _('Writing Transcript Records  (This will seem slow because of the size of the Transcript Records.)'))
                            
                        else:
                            rtfData = RTFText
			# now simply write the RTF data to the file.  (This does NOT need to be encoded, as the RTF already is!)
			# (but check to make sure there's actually RTF data there!)
			if rtfData != None:
                            f.write('%s' % rtfData)
                        f.write('      </RTFText>\n')
                    f.write('    </Transcript>\n')
                if dbCursor.rowcount > 0:
                    f.write('  </TranscriptFile>\n')
                dbCursor.close()

            progress.Update(63, _('Writing Keyword Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT KeywordGroup, Keyword, Definition FROM Keywords2'
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('  <KeywordFile>\n')
                for (KeywordGroup, Keyword, Definition) in dbCursor.fetchall():
                    f.write('    <KeywordRec>\n')
                    f.write('      <KeywordGroup>\n')
                    f.write('        %s\n' % KeywordGroup.encode(EXPORT_ENCODING))
                    f.write('      </KeywordGroup>\n')
                    f.write('      <Keyword>\n')
                    f.write('        %s\n' % Keyword.encode(EXPORT_ENCODING))
                    f.write('      </Keyword>\n')
                    if Definition != '':
                        f.write('      <Definition>\n')
                        f.write('        %s\n' % Definition.encode(EXPORT_ENCODING))
                        f.write('      </Definition>\n')
                    f.write('    </KeywordRec>\n')
                if dbCursor.rowcount > 0:
                    f.write('  </KeywordFile>\n')
                dbCursor.close()

            progress.Update(72, _('Writing Clip Keyword Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT EpisodeNum, ClipNum, KeywordGroup, Keyword, Example FROM ClipKeywords2'
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('  <ClipKeywordFile>\n')
                for (EpisodeNum, ClipNum, KeywordGroup, Keyword, Example) in dbCursor.fetchall():
                    f.write('    <ClipKeyword>\n')
                    if EpisodeNum != '':
                        f.write('      <EpisodeNum>\n')
                        f.write('        %s\n' % EpisodeNum)
                        f.write('      </EpisodeNum>\n')
                    if ClipNum != '':
                        f.write('      <ClipNum>\n')
                        f.write('        %s\n' % ClipNum)
                        f.write('      </ClipNum>\n')
                    f.write('      <KeywordGroup>\n')
                    f.write('        %s\n' % KeywordGroup.encode(EXPORT_ENCODING))
                    f.write('      </KeywordGroup>\n')
                    f.write('      <Keyword>\n')
                    f.write('        %s\n' % Keyword.encode(EXPORT_ENCODING))
                    f.write('      </Keyword>\n')
                    if Example != '':
                        f.write('      <Example>\n')
                        f.write('        %s\n' % Example)
                        f.write('      </Example>\n')
                    f.write('    </ClipKeyword>\n')
                if dbCursor.rowcount > 0:
                    f.write('  </ClipKeywordFile>\n')
                dbCursor.close()

            progress.Update(81, _('Writing Note Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, CollectNum, ClipNum, TranscriptNum, NoteTaker, NoteText FROM Notes2'
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('  <NoteFile>\n')
                for (NoteNum, NoteID, SeriesNum, EpisodeNum, CollectNum, ClipNum, TranscriptNum, NoteTaker, NoteText) in dbCursor.fetchall():
                    f.write('    <Note>\n')
                    f.write('      <Num>\n')
                    f.write('        %s\n' % NoteNum)
                    f.write('      </Num>\n')
                    f.write('      <ID>\n')
                    f.write('        %s\n' % NoteID.encode(EXPORT_ENCODING))
                    f.write('      </ID>\n')
                    # Note Series Numbers could be None instead of 0.
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
                    # Note Transcript Numbers could be None instead of 0.
                    if (TranscriptNum != 0) and (TranscriptNum != None):
                        f.write('      <TranscriptNum>\n')
                        f.write('        %s\n' % TranscriptNum)
                        f.write('      </TranscriptNum>\n')
                    if NoteTaker != '':
                        f.write('      <NoteTaker>\n')
                        f.write('        %s\n' % NoteTaker.encode(EXPORT_ENCODING))
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
                        f.write('      <NoteText>\n')
                        # If note text is found ...
                        if NoteText != None:
                            # ... add it to the output file ...
                            f.write('        %s\n' % NoteText.encode(EXPORT_ENCODING))
                        # ... but if NO note text is found ...
                        else:
                            # ... explicitly note that!  (This was crashing the export/import process.)
                            f.write('        %s\n' % _('(No Note Text found.)').encode(EXPORT_ENCODING))
                        f.write('      </NoteText>\n')
                    f.write('    </Note>\n')
                if dbCursor.rowcount > 0:
                    f.write('  </NoteFile>\n')
                dbCursor.close()

            progress.Update(90, _('Writing Filter Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = 'SELECT ReportType, ReportScope, ConfigName, FilterDataType, FilterData FROM Filters2'
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('  <FilterFile>\n')
                for (ReportType, ReportScope, ConfigName, FilterDataType, FilterData) in dbCursor.fetchall():

                    if DEBUG:
                        print "FilterData rec:", ReportType, ReportScope, ConfigName.encode('utf8'), FilterDataType,
                        if FilterData == None:
                            print "FilterData == None"
                        else:
                            print
                            
                    if FilterData != None:
                        f.write('    <Filter>\n')
                        f.write('      <ReportType>\n')
                        f.write('        %s\n' % ReportType)
                        f.write('      </ReportType>\n')
                        f.write('      <ReportScope>\n')
                        f.write('        %s\n' % ReportScope)
                        f.write('      </ReportScope>\n')
                        f.write('      <ConfigName>\n')
                        f.write('        %s\n' % ConfigName.encode(EXPORT_ENCODING))
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
                        # and Collection (7), we have LIST data to process.
                        if FilterDataType in [1, 2, 3, 5, 6, 7]:
                            # Initialize the fileterDataList to None, in case there's an unpickling error
                            filterDataList = None
                            # Because of the BLOB size problem, some filter data may not unpickle correctly.  Better trap it.
                            try:
                                # MySQL for Python often saves the data as an array ...
                                if type(FilterData).__name__ == 'array':
                                    # ... so convert it to a string and un-pickle it
                                    filterDataList = cPickle.loads(FilterData.tostring())
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
                                    errorMsg += "   (" + _("Episode Clip Data Export") + ")"
                                elif ReportType == 4:
                                    errorMsg += "   (" + _("Collection Clip Data Export") + ")"
                                elif ReportType == 5:
                                    errorMsg += "   (" + _("Series Keyword Sequence Map") + ")"
                                elif ReportType == 6:
                                    errorMsg += "   (" + _("Series Keyword Bar Graph") + ")"
                                elif ReportType == 7:
                                    errorMsg += "   (" + _("Series Keyword Percentage Map") + ")"
                                elif ReportType == 8:
                                    errorMsg += "   (" + _("Episode Clip Data Coder Reliability Export") + ")"
                                elif ReportType == 9:
                                    errorMsg += "   (" + _("Keyword Summary Report") + ")"
                                elif ReportType == 10:
                                    errorMsg += "   (" + _("Series Report") + ")"
                                elif ReportType == 11:
                                    errorMsg += "   (" + _("Episode Report") + ")"
                                elif ReportType == 12:
                                    errorMsg += "   (" + _("Collection Report") + ")"
                                elif ReportType == 13:
                                    errorMsg += "   (" + _("Notes Report") + ")"
                                elif ReportType == 14:
                                    errorMsg += "   (" + _("Series Clip Data Export") + ")"
                                errorMsg += "\n"
                                # Tell the user which data object is having a problem.
                                errorMsg += "  " + _("ReportScope:     %s")
                                if ReportType in [5, 6, 7, 10, 14]:
                                    import Series
                                    tempSeries = Series.Series(ReportScope)
                                    errorMsg += "  (" + _("Series") + ' "%s")' % tempSeries.id
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
                                        errorMsg += "  " + _("(Series Notes)")
                                    elif ReportScope == 3:
                                        errorMsg += "  " + _("(Episode Notes)")
                                    elif ReportScope == 4:
                                        errorMsg += "  " + _("(Transcript Notes)")
                                    elif ReportScope == 5:
                                        errorMsg += "  " + _("(Collection Notes)")
                                    elif ReportScope == 6:
                                        errorMsg += "  " + _("(Clip Notes)")
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

                        # Other filter data types (9 though 17 so far as of 2.21) are of UNPICKLED, UNENCODED data
                        else:
                            # MySQL for Python often saves the data as an array ...
                            if type(FilterData).__name__ == 'array':
                                # ... so convert it to a string
                                filterDataList = FilterData.tostring()
                            else:
                                filterDataList = FilterData
                            # Encode the string using the export encoding
                            filterDataList = filterDataList.encode(EXPORT_ENCODING)

                        # If we have data from FilterDataTypes 1, 2, 3, 5, 6, or 7 ...
                        if FilterDataType in [1, 2, 3, 5, 6, 7]:
                            # ... we need to re-pickle the data
                            FilterData = cPickle.dumps(filterDataList)
                        # Otherwise ...
                        else:
                            # ... the data's ready for output.
                            FilterData = filterDataList

                        f.write('      <FilterData>\n')
                        f.write('        %s\n' % FilterData)
                        f.write('      </FilterData>\n')
                        f.write('    </Filter>\n')
                if dbCursor.rowcount > 0:
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

            if DEBUG:
                import traceback
                traceback.print_exc(file=sys.stdout)
            dbCursor.close()

        f.close()
        progress.Update(100)
        progress.Destroy()


    def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        fs = wx.FileSelector(_("Select an XML file for export"),
                        TransanaGlobal.programDir,
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
