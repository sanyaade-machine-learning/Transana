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

"""This module implements the Selective Data Export function. """

__author__ = 'Ronny Steinmetz'

DEBUG = False
if DEBUG:
    print "SelectiveDataExport DEBUG is ON!!"

import wx
import os
import Dialogs
import DBInterface
import TransanaConstants
import TransanaGlobal
if TransanaConstants.USESRTC:
    from RichTextEditCtrl_RTC import RichTextEditCtrl
else:
    from RichTextEditCtrl import RichTextEditCtrl
import cPickle
import pickle
import sys
import Library
import Episode
import Collection
import Clip
import FilterDialog
import Misc
import XMLExport

class SelectiveDataExport(Dialogs.GenForm):
    """ This windows displays a variety of GUI Widgets. """
    def __init__(self,parent,id,seriesNum=0, episodeNum=0, collectionNum=0):
        # Remember the Library, episode or collection that triggered creation of this export
        self.seriesNum = seriesNum
        self.episodeNum = episodeNum
        self.collectionNum = collectionNum

        # Set the encoding for export.
        # Use UTF-8 regardless of the current encoding for consistency in the Transana XML files
        EXPORT_ENCODING = 'utf8'

        # Create a form to get the name of the file to receive the data
        # Define the form title
        title = _("Transana Selective Data Export")
        Dialogs.GenForm.__init__(self, parent, id, title, (550,150), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Selective Data Export')
	# Create an invisible instance of RichTextEditCtrl. This allows us to
	# get away with converting fastsaved documents to RTF behind the scenes.
	# Then we can simply pull the RTF data out of this object and write it
	# to the desired file.
	self.invisibleSTC = RichTextEditCtrl(self)
	self.invisibleSTC.Show(False)

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Export Message
        # If the XML filename path is not empty, we need to tell the user.
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
        # Export Filename Layout
        self.XMLFile = self.new_edit_box(_("Transana Export XML-Filename"), lay, '')
        self.XMLFile.SetDropTarget(EditBoxFileDropTarget(self.XMLFile))
        # Add the element sizer to the row sizer
        r2Sizer.Add(v1, 1, wx.EXPAND)

        # Add a spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Browse button
        browse = wx.Button(self.panel, wx.ID_FILE1, _("Browse"), wx.DefaultPosition)
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)
        # Add the element to the row sizer
        r2Sizer.Add(browse, 0, wx.ALIGN_BOTTOM)

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

    def Export(self):
        # use the LONGEST title here!  That determines the size of the Dialog Box.
        progress = wx.ProgressDialog(_('Transana Selective Data Export'), _('Exporting Transcript records (This may be slow because of the size of Transcript records.)'), style = wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)
        if progress.GetSize()[0] > 800:
            progress.SetSize((800, progress.GetSize()[1]))
            progress.Centre()

        db = DBInterface.get_db()

        try:
            # Get the user specified file name
            fs = self.XMLFile.GetValue()
            # Ensure that the file name has the proper extension
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
            # Open the output file for writing
            f = file(fs, 'w')
            progress.Update(0, _('Writing Headers'))

            # Importing Definitions from XMLExport
            XMLExportObject = XMLExport.XMLExport(self, -1, '')

            # Writing XML's DTD Headers
            XMLExportObject.WritingXMLDTD(f)

            # Writing Library Records
            progress.Update(9, _('Writing Library Records'))
            seriesList = []

            if db != None:
                dbCursor = db.cursor()
            
                if self.seriesNum <> 0:
                    SQLText = """SELECT SeriesNum, SeriesID, SeriesComment, SeriesOwner, DefaultKeywordGroup
                                 FROM Series2
                                 WHERE SeriesNum = %s"""
                    dbCursor.execute(SQLText, self.seriesNum)
                    if dbCursor.rowcount > 0:
                        f.write('  <SeriesFile>\n')
                        XMLExportObject.WritingSeriesRecords(f, dbCursor)
                    if dbCursor.rowcount > 0:
                        f.write('  </SeriesFile>\n')

                if self.episodeNum <> 0:
                    SQLText = """SELECT a.SeriesNum, SeriesID, SeriesComment, SeriesOwner, DefaultKeywordGroup
                                 FROM Series2 a, Episodes b
                                 WHERE a.SeriesNum = b.SeriesNum
                                 AND b.EpisodeNum = %s"""
                    dbCursor.execute(SQLText, self.episodeNum)
                    if dbCursor.rowcount > 0:
                        f.write('  <SeriesFile>\n')
                        XMLExportObject.WritingSeriesRecords(f, dbCursor)
                    if dbCursor.rowcount > 0:
                        f.write('  </SeriesFile>\n')

                if self.collectionNum <> 0:
                    SQLText = """SELECT a.SeriesNum, SeriesID, SeriesComment, SeriesOwner, DefaultKeywordGroup
                                 FROM Series2 a, Episodes b, Clip c, Collect d
                                 WHERE a.SeriesNum = b.SeriesNum
                                 AND b.EpisodeNum = c.EpisodeNum
                                 AND c.CollectNum = d.CollectNum
                                 AND d.CollectNum = %s"""
                    dbCursor.execute(SQLText, self.collectNum)
                    if dbCursor.rowcount > 0:
                        f.write('  <SeriesFile>\n')
                        XMLExportObject.WritingSeriesRecords(f, dbCursor)
                    if dbCursor.rowcount > 0:
                        f.write('  </SeriesFile>\n')

                # Populating list with episodes
                for seriesRec in dbCursor.fetchall():
                    if seriesRec not in seriesList:
                        seriesList.append(seriesRec[0])

                dbCursor.close()


            if DEBUG:
                print "Library Records:", seriesList
                    
            # Writing Episode Records
            progress.Update(18, _('Writing Episode Records'))
            episodesList = []
            
            if db != None:
                dbCursor = db.cursor()
                SQLText = """SELECT EpisodeNum, EpisodeID, SeriesNum, TapingDate, MediaFile, EpLength, EpComment
                             FROM Episodes2
                             WHERE SeriesNum = %s"""
#                episodes = FakeCursorObject()
                dbCursor.execute(SQLText, seriesRec[0])
                    
                if dbCursor.rowcount > 0:
                    f.write('  <EpisodeFile>\n')
                    XMLExportObject.WritingEpisodeRecords(f, dbCursor)
                if dbCursor.rowcount > 0:
                    f.write('  </EpisodeFile>\n')

                # Populating list with episodes
                for episodeRec in dbCursor.fetchall():
                    if episodeRec not in episodesList:
                        episodesList.append(episodeRec[0])

                if DEBUG:
                    print "Episode Records:", episodesList
                    
                dbCursor.close()

            # Writing Core Data Records
            progress.Update(27, _('Writing Core Data Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = """SELECT CoreDataNum, Identifier, Title, Creator, Subject, Description, Publisher, Contributor, DCDate, DCType, Format, Source, Language, Relation, Coverage, Rights
                             FROM CoreData2
                             WHERE Identifier = 'Volume.mpg'"""
                dbCursor.execute(SQLText)
                if dbCursor.rowcount > 0:
                    f.write('   <CoreDataFile>\n')
                    XMLExportObject.WritingCoreDataRecords(f, dbCursor)
                if dbCursor.rowcount > 0:
                     f.write('   </CoreDataFile>\n')
                dbCursor.close()

            # Writing Collection Records
            progress.Update(36, _('Writing Collection Records'))
            collectionsList = []
            
            if db != None:
                dbCursor = db.cursor()
                SQLText = """SELECT a.CollectNum, CollectID, ParentCollectNum, CollectComment, CollectOwner, a.DefaultKeywordGroup
                             FROM Collections2 a, Clips2 b
                             WHERE a.CollectNum = b.CollectNum AND
                                   b.EpisodeNum = %s"""
                collections = FakeCursorObject()
                dbCursor.execute(SQLText, episodeRec[0])

                # Populating list with episodes
                for collectionRec in dbCursor.fetchall():

                    print collectionRec
                    print collectionsList
                    print (collectionRec[0], collectionRec[1]) not in collectionsList
                    print
                    
                    if (collectionRec[0], collectionRec[1]) not in collectionsList:
                        collectionsList.append((collectionRec[0], collectionRec[1]))
                        collections.append(collectionRec)

                if dbCursor.rowcount > 0:
                    f.write('  <CollectionFile>\n')
                    XMLExportObject.WritingCollectionRecords(f, collections)
                if dbCursor.rowcount > 0:
                    f.write('  </CollectionFile>\n')

                if DEBUG:
                    print "Collection Recs:", collectionsList
                    
                dbCursor.close()

            #Writing Clip Records
            progress.Update(45, _('Writing Clip Records'))
            clipsList = []
            
            if db != None:
                dbCursor = db.cursor()
                SQLText = """SELECT ClipNum, ClipID, CollectNum, a.EpisodeNum, a.MediaFile, ClipStart, ClipStop, ClipComment, SortOrder
                             FROM Clips2 a, Episodes2 b
                             WHERE a.EpisodeNum = %s"""
#                clips = FakeCursorObject()
                dbCursor.execute(SQLText, episodeRec[0])
                if dbCursor.rowcount > 0:
                    f.write('  <ClipFile>\n')
                    XMLExportObject.WritingClipRecords(f, dbCursor)
                if dbCursor.rowcount > 0:
                    f.write('  </ClipFile>\n')

                # Populating list with clips
                for clipRec in dbCursor.fetchall():
                    if clipRec not in clipsList:
                        clipsList.append(clipRec[0])

                if DEBUG:
                    print "Clip Recs:", clipsList

                dbCursor.close()

            #Writing Transcript Records
            transcriptsList = []
            progress.Update(54, _('Writing Transcript Records  (This will seem slow because of the size of the Transcript Records.)'))
            if db != None:
                dbCursor = db.cursor()
                #Querying all transcripts based on Episodes
                SQLText1 = """SELECT TranscriptNum, TranscriptID, EpisodeNum, SourceTranscriptNum, ClipNum, SortOrder, Transcriber, ClipStart, ClipStop, Comment, RTFText
                             FROM Transcripts2
                             WHERE EpisodeNum = %s
                             AND ClipNum = 0"""
                #Querying all transcripts based on Clips
                SQLText2 = """SELECT TranscriptNum, TranscriptID, EpisodeNum, SourceTranscriptNum, ClipNum, SortOrder, Transcriber, ClipStart, ClipStop, Comment, RTFText
                             FROM Transcripts2
                             WHERE ClipNum = %s"""

                if DEBUG:
                    print "Selecting Transcripts"
                    
#                transcripts = FakeCursorObject()

                # Adding all transcripts based on episodes to the list of all transcripts
                for episodeRec in episodesList:
                    dbCursor.execute(SQLText1, episodeRec)
                    for record in dbCursor.fetchall():
                        transcriptsList.append(record)

                # Adding all transcripts based on clips to the list of all transcripts
                for clipRec in clipsList:
                    dbCursor.execute(SQLText2, clipRec)
                    for record in dbCursor.fetchall():
                        transcriptsList.append(record)

                if DEBUG:
                    print "%d Transcripts selected." % dbCursor.rowcount
                    
                if dbCursor.rowcount > 0:
                    f.write('  <TranscriptFile>\n')
                    #Writing all transcripts into the xml file
                    XMLExportObject.WritingTranscriptRecords(f, dbCursor, progress)
                if dbCursor.rowcount > 0:
                    f.write('  </TranscriptFile>\n')

                if DEBUG:
                    print
                    print transcriptsList

                dbCursor.close()

            # Collecting ClipKeyword data
            clipKeywordsList = []
            keywordsList = []
            
            # Collecting ClipKeywords from Episodes
            # Iterate through the Episode list...
            for episodeRec in episodesList:

                if DEBUG:
                    print
                    print "episodeRec =", episodeRec
                
                # ... add each episode's keyword to the episodeClipKeywords list...
                episodeClipKeywords = DBInterface.list_of_keywords(Episode = episodeRec)

                if DEBUG:
                    print
                    print "Episode ClipKeywords =", episodeClipKeywords
                
                # ... and add this list to the actual ClipKeyword list.
                for episodeClip in episodeClipKeywords:
                    if episodeClip not in clipKeywordsList:
                        clipKeywordsList.append(episodeClip)

                for clipKeyword in clipKeywordsList:

                    if DEBUG:
                        print
                        print clipKeyword[0], clipKeyword[1]
                    
                    # If the keyword isn't already in in the Keyword List...
                    if (clipKeyword[0], clipKeyword[1]) not in keywordsList:
                        # ... add the keyword to the Keyword List.
                        keywordsList.append((clipKeyword[0], clipKeyword[1]))

            if DEBUG:
                print
                print "Keword List =", keywordsList
                
            # Collecting ClipKeywords from Clips            
            # Iterate through the Clip List...
            for clipRec in clipsList:
            
                if DEBUG:
                    print
                    print "clipRec =", clipRec

                # ... add each clip's keyword to the clipClipKeywords list
                clipClipKeywords = DBInterface.list_of_keywords(Clip = clipRec)

                if DEBUG:
                    print
                    print "Clip ClipKeywords =", clipClipKeywords

                # ... and add this to the actual ClipKeyword list.
                for clipClipKeyword in clipClipKeywords:
                    if clipClipKeyword not in clipKeywordsList:
                        clipKeywordsList.append(clipClipKeyword)

                for clipKeyword in clipKeywordsList:

                    if DEBUG:
                        print
                        print clipKeyword[0], clipKeyword[1]
                    
                    # If the keyword isn't already in the Keyword List...
                    if (clipKeyword[0], clipKeyword[1]) not in keywordsList:
                        # ... add the keyword to the Keyword List.
                        keywordsList.append((clipKeyword[0], clipKeyword[1]))
                        
            if DEBUG:
                print
                print "Keyword List =", keywordsList

            # Writing Keyword Records
            progress.Update(63, _('Writing Keyword Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = """SELECT KeywordGroup, Keyword, Definition
                             FROM Keywords2
                             WHERE KeywordGroup = %s
                             AND Keyword = %s"""

                allKeywords = FakeCursorObject()
                
                for clipKeyword in keywordsList:
                    dbCursor.execute(SQLText, (clipKeyword[0], clipKeyword[1]))
                    for record in dbCursor.fetchall():
                        allKeywords.append(record)
                    
                if dbCursor.rowcount > 0:
                    f.write('  <KeywordFile>\n')
                    XMLExportObject.WritingKeywordRecords(f, allKeywords)
                if dbCursor.rowcount > 0:
                    f.write('  </KeywordFile>\n')
                dbCursor.close()

            # Writing ClipKeyword Records
            progress.Update(72, _('Writing Clip Keyword Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = """SELECT EpisodeNum, ClipNum, KeywordGroup, Keyword, Example
                             FROM ClipKeywords2
                             WHERE KeywordGroup = %s
                             AND Keyword = %s"""

                allClipKeywords = FakeCursorObject()

                for clipKeyword in clipKeywordsList:
                    dbCursor.execute(SQLText, (clipKeyword[0], clipKeyword[1]))
                    for record in dbCursor.fetchall():
                        allClipKeywords.append(record)
                
                if dbCursor.rowcount > 0:
                    f.write('  <ClipKeywordFile>\n')
                    XMLExportObject.WritingClipKeywordRecords(f, allClipKeywords)
                if dbCursor.rowcount > 0:
                    f.write('  </ClipKeywordFile>\n')
                dbCursor.close()

            # Collecting Note data
            notesList = []

            # Collecting NoteData from Library...
            for seriesRec in seriesList:
                seriesNotesList = DBInterface.list_of_notes(Library = seriesRec)
                
            if DEBUG:
                print
                print "Library Notes List =", seriesNotesList
                               
            for seriesNote in seriesNotesList:
                # If the note isn't already in the notesList...
                if seriesNote not in notesList:
                    # ... add the note to the notesList.
                    notesList.append(seriesNote)
            
            # Collecting NoteData from Episodes
            # Iterate through the Episode list...
            for episodeRec in episodesList:
                # ... add each episode's note to the episodeNoteList...
                episodeNotesList = DBInterface.list_of_notes(Episode = episodeRec)

            if DEBUG:
                print
                print "Episode Note List =", episodeNotesList

                for episodeNote in episodeNotesList:
                    # If the note isn't already in the Note List...
                    if episodeNote not in notesList:
                        # ... add the note to the Note List.
                        notesList.append(episodeNote)

            # Collecting NoteData from Collections
            # Iterate through the Collection list...
            
            for collectionRec in collectionsList:
                # ... add each collection's note to the collectionNoteList...
                collectionNotesList = DBInterface.list_of_notes(Collection = collectionRec[0])

                for collectionNote in collectionNotesList:
                    # If the note isn't alreadz in the Note List...
                    if collectionNote not in notesList:
                        # ... add the note to the Note List
                        notesList.append(collectionNote)

            if DEBUG:
                print
                print "Collection Note List =", collectionNotesList

            if DEBUG:
                print
                print "Note List =", notesList

            progress.Update(81, _('Writing Note Records'))
            if db != None:
                dbCursor = db.cursor()
                SQLText = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, CollectNum, ClipNum, TranscriptNum, NoteTaker, NoteText
                           FROM Notes2
                           WHERE NoteID = %s"""
                
                allNotes = FakeCursorObject()
                
                for note in notesList:
                    dbCursor.execute(SQLText, note)
                    for record in dbCursor.fetchall():
                        allNotes.append(record)
                    
                if dbCursor.rowcount > 0:
                    f.write('  <NoteFile>\n')
                    XMLExportObject.WritingNoteRecords(f, allNotes)
                if dbCursor.rowcount > 0:
                    f.write('  </NoteFile>\n')
                dbCursor.close()

            XMLExportObject.Destroy()
                
        except:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('An error occurred during Selective Data Export.\n%s\n%s'), 'utf8')
            else:
                prompt = _('An error occurred during Selective Data Export.\n%s\n%s')
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


class FakeCursorObject(object):
    def __init__(self):
        self.data = []

    def append(self, datum):
        self.data.append(datum)

    def fetchall(self):
        return self.data
