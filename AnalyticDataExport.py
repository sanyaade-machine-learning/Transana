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

"""This module implements the Analytic Data Export function. """

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

import wx
# import Python's os module
import os
import Dialogs
import Library
import Document
import Episode
import Collection
import Quote
import Clip
import DBInterface
import FilterDialog
import TransanaConstants
import TransanaExceptions
import TransanaGlobal
import Misc
# import Python's codecs module to make reading and writing UTF-8 text files 
import codecs


class AnalyticDataExport(Dialogs.GenForm):
    """ This class creates the tab-delimited text file that is the Analytic Data Export. """
    def __init__(self, parent, id, libraryNum=0, documentNum=0, episodeNum=0, collectionNum=0):
        # Remember the Library, Document, Episode or Collection that triggered creation of this report
        self.libraryNum = libraryNum
        self.documentNum = documentNum
        self.episodeNum = episodeNum
        self.collectionNum = collectionNum

        # Create a form to get the name of the file to receive the data
        # Define the form title
        title = _("Transana Analytic Data Export")
        # Create the form itself
        Dialogs.GenForm.__init__(self, parent, id, title, (550,150), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Analytic Data Export')

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Header Message
        prompt = _('Please create a Transana Analytic Data File for export.')
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

        # Export Filename
        self.exportFile = self.new_edit_box(_("Export Filename"), v1, '')
        self.exportFile.SetDropTarget(EditBoxFileDropTarget(self.exportFile))
        # Add the element sizer to the row sizer
        r2Sizer.Add(v1, 1, wx.EXPAND)

        # Add a spacer to the row sizer        
        r2Sizer.Add((10, 0))

        # Browse button
        browse = wx.Button(self.panel, wx.ID_FILE1, _("Browse"), wx.DefaultPosition)
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
        self.SetSize(wx.Size(max(550, width), max(100, height)))
        # Define the minimum size for this dialog as the current size, and define height as unchangeable
        self.SetSizeHints(max(550, width), max(100, height), -1, max(100, height))
        # Center the form on screen
        TransanaGlobal.CenterOnPrimary(self)

        # Set focus to the Export File Name field
        self.exportFile.SetFocus()


    def Export(self):
        """ Export the Analytic Data to a Tab-delimited file """
        # Initialize values for data structures for this report
        # The Episode List is the list of Episodes to be sent to the Filter Dialog for the Library report
        episodeList = []
        # The Document List is the list of Documents to be sent to the Filter Dialog for the Library report
        documentList = []
        # The Quote List is the list of Quotes to be sent to the Filter Dialog
        quoteList = []
        # The Quote Lookup allows us to find the Quote Number based on the data from the Quote List
        quoteLookup = {}
        # The Clip List is the list of Clips to be sent to the Filter Dialog
        clipList = []
        # The Clip Lookup allows us to find the Clip Number based on the data from the Clip List
        clipLookup = {}
        # The Keyword List is the list of Keywords to be sent to the Filter Dialog
        keywordList = []
        # Show a WAIT cursor.  Assembling the data can take noticable time in some cases.
        TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))

        # If we have an Library Number, we set up the Library Analytic Data Export
        if self.libraryNum <> 0:
            # Get the Library record
            tempLibrary = Library.Library(self.libraryNum)

            # obtain a list of all Documents in the Library
            tempDocumentList = DBInterface.list_of_documents(tempLibrary.number)
            # initialize the temporary Quote List
            tempQuoteList = []
            # iterate through the Document List ...
            for documentRecord in tempDocumentList:
                # ... and add each Document's Quotes to the Temporary Quote list
                tempQuoteList += DBInterface.list_of_quotes_by_document(documentRecord[0])
                # Add the Document data to the Filter Dialog's Document List
                documentList.append((documentRecord[1], tempLibrary.id, True))
            # For all the Quotes ...
            for quoteRecord in tempQuoteList:
                # ... add the Quote to the Quote List for filtering ...
                quoteList.append((quoteRecord['QuoteID'], quoteRecord['CollectNum'], True))
                # ... retain a pointer to the Quote Number keyed to the Quote ID and Collection Number ...
                quoteLookup[(quoteRecord['QuoteID'], quoteRecord['CollectNum'])] = quoteRecord['QuoteNum']
                # ... now get all the keywords for this Quote ...
                quoteKeywordList = DBInterface.list_of_keywords(Quote = quoteRecord['QuoteNum'])
                # ... and iterate through the list of Quote keywords.
                for quoteKeyword in quoteKeywordList:
                    # If the keyword isn't already in the Keyword List ...
                    if (quoteKeyword[0], quoteKeyword[1], True) not in keywordList:
                        # ... add the keyword to the keyword list for filtering.
                        keywordList.append((quoteKeyword[0], quoteKeyword[1], True))

            # obtain a list of all Episodes in that Library
            tempEpisodeList = DBInterface.list_of_episodes_for_series(tempLibrary.id)
            # initialize the temporary clip List
            tempClipList = []
            # iterate through the Episode List ...
            for episodeRecord in tempEpisodeList:
                # ... and add each Episode's Clips to the Temporary clip list
                tempClipList += DBInterface.list_of_clips_by_episode(episodeRecord[0])
                # Add the Episode data to the Filter Dialog's Episode List
                episodeList.append((episodeRecord[1], tempLibrary.id, True))
            # For all the Clips ...
            for clipRecord in tempClipList:
                # ... add the Clip to the Clip List for filtering ...
                clipList.append((clipRecord['ClipID'], clipRecord['CollectNum'], True))
                # ... retain a pointer to the Clip Number keyed to the Clip ID and Collection Number ...
                clipLookup[(clipRecord['ClipID'], clipRecord['CollectNum'])] = clipRecord['ClipNum']
                # ... now get all the keywords for this Clip ...
                clipKeywordList = DBInterface.list_of_keywords(Clip = clipRecord['ClipNum'])
                # ... and iterate through the list of clip keywords.
                for clipKeyword in clipKeywordList:
                    # If the keyword isn't already in the Keyword List ...
                    if (clipKeyword[0], clipKeyword[1], True) not in keywordList:
                        # ... add the keyword to the keyword list for filtering.
                        keywordList.append((clipKeyword[0], clipKeyword[1], True))
            
        # If we have a Document Number, we set up the Document Analytic Data Export
        elif self.documentNum <> 0:
            # First, we get a list of all the Quotes for the Document specified
            tempQuoteList = DBInterface.list_of_quotes_by_document(self.documentNum)
            # For all the Quotes ...
            for quoteRecord in tempQuoteList:
                # ... add the Quote to the Quote List for filtering ...
                quoteList.append((quoteRecord['QuoteID'], quoteRecord['CollectNum'], True))
                # ... retain a pointer to the Quote Number keyed to the Quote ID and Collection Number ...
                quoteLookup[(quoteRecord['QuoteID'], quoteRecord['CollectNum'])] = quoteRecord['QuoteNum']
                # ... now get all the keywords for this Quote ...
                quoteKeywordList = DBInterface.list_of_keywords(Quote = quoteRecord['QuoteNum'])
                # ... and iterate through the list of Quote keywords.
                for quoteKeyword in quoteKeywordList:
                    # If the keyword isn't already in the Keyword List ...
                    if (quoteKeyword[0], quoteKeyword[1], True) not in keywordList:
                        # ... add the keyword to the keyword list for filtering.
                        keywordList.append((quoteKeyword[0], quoteKeyword[1], True))

        # If we have an Episode Number, we set up the Episode Analytic Data Export
        elif self.episodeNum <> 0:
            # First, we get a list of all the Clips for the Episode specified
            tempClipList = DBInterface.list_of_clips_by_episode(self.episodeNum)
            # For all the Clips ...
            for clipRecord in tempClipList:
                # ... add the Clip to the Clip List for filtering ...
                clipList.append((clipRecord['ClipID'], clipRecord['CollectNum'], True))
                # ... retain a pointer to the Clip Number keyed to the Clip ID and Collection Number ...
                clipLookup[(clipRecord['ClipID'], clipRecord['CollectNum'])] = clipRecord['ClipNum']
                # ... now get all the keywords for this Clip ...
                clipKeywordList = DBInterface.list_of_keywords(Clip = clipRecord['ClipNum'])
                # ... and iterate through the list of clip keywords.
                for clipKeyword in clipKeywordList:
                    # If the keyword isn't already in the Keyword List ...
                    if (clipKeyword[0], clipKeyword[1], True) not in keywordList:
                        # ... add the keyword to the keyword list for filtering.
                        keywordList.append((clipKeyword[0], clipKeyword[1], True))

        # If we don't have Library Number, Document Number, or Episode number, but DO have a Collection Number, we set
        # up the Clips for the Collection specified.  If we have neither, it's the GLOBAL Analytic Data Export,
        # requesting ALL the Quotes and Clips in the database!  We can handle both of these cases together.
        else:
            # If we have a specific collection specified ...
            if self.collectionNum <> 0:
                # ... load the specified collection.  We need its data.
                tempCollection = Collection.Collection(self.collectionNum)
                # Put the selected Collection's data into the Collection List as a starting place.
                tempCollectionList = [(tempCollection.number, tempCollection.id, tempCollection.parent)]
            # If we don't have any selected collection ...
            else:
                # ... then we should initialise the Collection List with data for all top-level collections, with parent = 0
                tempCollectionList = DBInterface.list_of_collections()
            # Iterate through the Collection List as long as it has entries
            while len(tempCollectionList) > 0:
                # Get the list of Quotes for the current Collection
                tempQuoteList = DBInterface.list_of_quotes_by_collectionnum(tempCollectionList[0][0])
                # For all the Quotes ...
                for (quoteNo, quoteName, collNo, sourceDocNo) in tempQuoteList:
                    # ... add the Quote to the Quote List for filtering ...
                    quoteList.append((quoteName, collNo, True))
                    # ... retain a pointer to the Quote Number keyed to the Quote ID and Collection Number ...
                    quoteLookup[(quoteName, collNo)] = quoteNo
                    # ... now get all the keywords for this Quote ...
                    quoteKeywordList = DBInterface.list_of_keywords(Quote = quoteNo)
                    # ... and iterate through the list of Quote keywords.
                    for quoteKeyword in quoteKeywordList:
                        # If the keyword isn't already in the Keyword List ...
                        if (quoteKeyword[0], quoteKeyword[1], True) not in keywordList:
                            # ... add the keyword to the keyword list for filtering.
                            keywordList.append((quoteKeyword[0], quoteKeyword[1], True))

                # Get the list of Clips for the current Collection
                tempClipList = DBInterface.list_of_clips_by_collection(tempCollectionList[0][1], tempCollectionList[0][2])
                # For all the Clips ...
                for (clipNo, clipName, collNo) in tempClipList:
                    # ... add the Clip to the Clip List for filtering ...
                    clipList.append((clipName, collNo, True))
                    # ... retain a pointer to the Clip Number keyed to the Clip ID and Collection Number ...
                    clipLookup[(clipName, collNo)] = clipNo
                    # ... now get all the keywords for this Clip ...
                    clipKeywordList = DBInterface.list_of_keywords(Clip = clipNo)
                    # ... and iterate through the list of clip keywords.
                    for clipKeyword in clipKeywordList:
                        # If the keyword isn't already in the Keyword List ...
                        if (clipKeyword[0], clipKeyword[1], True) not in keywordList:
                            # ... add the keyword to the keyword list for filtering.
                            keywordList.append((clipKeyword[0], clipKeyword[1], True))

                # Get the nested collections for the current collection and add them to the Collection List
                tempCollectionList += DBInterface.list_of_collections(tempCollectionList[0][0])
                # Remove the current Collection from the list.  We're done with it.
                del(tempCollectionList[0])

        # Put the Quote List in alphabetical order in preparation for Filtering..
        quoteList.sort()
        # Put the Clip List in alphabetical order in preparation for Filtering..
        clipList.sort()
        # Put the Keyword List in alphabetical order in preparation for Filtering.
        keywordList.sort()

        # Prepare the Filter Dialog.
        # Set the title for the Filter Dialog
        title = unicode(_("Analytic Data Export Filter Dialog"), 'utf8')
        # If we have a Library-based report ...
        if self.libraryNum != 0:
            # ... reportType 14 indicates Library Analytic Data Export to the Filter Dialog
            reportType = 14
            # ... the reportScope is the Library Number.
            reportScope = self.libraryNum
        # If we have a Document-based report ...
        elif self.documentNum != 0:
            # ... reportType 20 indicates Document Analytic Data Export to the Filter Dialog
            reportType = 20
            # ... the reportScope is the Document Number.
            reportScope = self.documentNum
        # If we have an Episode-based report ...
        elif self.episodeNum != 0:
            # ... reportType 3 indicates Episode Analytic Data Export to the Filter Dialog
            reportType = 3
            # ... the reportScope is the Episode Number.
            reportScope = self.episodeNum
        # If we have a Collection-based report ...
        else:
            # ... reportType 4 indicates Collection Analytic Data Export to the Filter Dialog
            reportType = 4
            # ... the reportScope is the Collection Number.
            reportScope = self.collectionNum

        showDocuments = (len(documentList) > 0)
        showEpisodes = (len(episodeList) > 0)
        showQuotes = (len(quoteList) > 0)
        showClips = (len(clipList) > 0)

        # If we are basing the report on a Library ...
        if self.libraryNum != 0:
            # ... create a Filter Dialog, passing all the necessary parameters.  We want to include the Episode List
            dlgFilter = FilterDialog.FilterDialog(None, -1, title, reportType=reportType, reportScope=reportScope,
                                                  documentFilter=showDocuments,
                                                  episodeFilter=showEpisodes,
                                                  quoteFilter=showQuotes,
                                                  clipFilter=showClips,
                                                  keywordFilter=True)
        # If we are basing the report on a Collection (but not the Collection Root) ...
        elif self.collectionNum != 0:
            # ... create a Filter Dialog, passing all the necessary parameters.  We want to be able to include Nested Collections
            dlgFilter = FilterDialog.FilterDialog(None, -1, title, reportType=reportType, reportScope=reportScope,
                                                  quoteFilter=showQuotes, clipFilter=showClips, keywordFilter=True,
                                                  reportContents=True, showNestedData=True)
        # If we are basing the report on a Document ...
        elif self.documentNum != 0:
            # ... create a Filter Dialog, passing all the necessary parameters.  We DON'T need Nested Collections
            dlgFilter = FilterDialog.FilterDialog(None, -1, title, reportType=reportType, reportScope=reportScope,
                                                  quoteFilter=True, keywordFilter=True)
        # If we are basing the report on an Episode ...
        elif self.episodeNum != 0:
            # ... create a Filter Dialog, passing all the necessary parameters.  We DON'T need Nested Collections
            dlgFilter = FilterDialog.FilterDialog(None, -1, title, reportType=reportType, reportScope=reportScope,
                                                  clipFilter=True, keywordFilter=True)
        # If we are doing the report on the Collection Root report, which MUST have nested data) ...
        else:
            # ... create a Filter Dialog, passing all the necessary parameters.  We DON'T need Nested Collections
            dlgFilter = FilterDialog.FilterDialog(None, -1, title, reportType=reportType, reportScope=reportScope,
                                                  quoteFilter=showQuotes, clipFilter=showClips, keywordFilter=True)
        # If we have a Library-based report ...
        if self.libraryNum != 0:
            if showDocuments:
                # ... populate the Document Data Structure
                dlgFilter.SetDocuments(documentList)
            if showEpisodes:
                # ... populate the Episode Data Structure
                dlgFilter.SetEpisodes(episodeList)
        # Populate the Quote, Clip and Keyword Data Structures
        if showQuotes:
            dlgFilter.SetQuotes(quoteList)
        if showClips:
            dlgFilter.SetClips(clipList)
        dlgFilter.SetKeywords(keywordList)

        # ... get the list of existing configuration names.
        profileList = dlgFilter.GetConfigNames()
        # If (translated) "Default" is in the list ...
        # (NOTE that the default config name is stored in English, but gets translated by GetConfigNames!)
        if unicode(_('Default'), 'utf8') in profileList:
            # ... set the Filter Dialog to use this filter
            dlgFilter.configName = unicode(_('Default'), TransanaGlobal.encoding)
            # Temporarily set loadDefault to True for the Filter Dialog.  This disables the Filter Load dialog.
            # (We don't use the FilterDialog parameter, as that disables loading other Filters!)
            dlgFilter.loadDefault = True
            # We need to load the config
            dlgFilter.OnFileOpen(None)
            # Now we turn loadDefault back off so we can load other filters if we want.
            dlgFilter.loadDefault = False
            
        # restore the cursor, now that the data is set up for the filter dialog
        TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

        # If the user clicks OK ...
        if dlgFilter.ShowModal() == wx.ID_OK:
            # Set the WAIT cursor.  It can take a while to build the data file.
            TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
            # If we have a Library-based report ...
            if self.libraryNum != 0:
                if showDocuments:
                    documentList = dlgFilter.GetDocuments()
                if showEpisodes:
                    # ... get the revised Episode data from the Filter Dialog
                    episodeList = dlgFilter.GetEpisodes()
            # Get the revised Quote, Clip, and Keyword data from the Filter Dialog
            if showQuotes:
                quoteList = dlgFilter.GetQuotes()
            if showClips:
                clipList = dlgFilter.GetClips()
            keywordList = dlgFilter.GetKeywords()
            # If we have a Collection-based report ...
            if self.collectionNum != 0:
                # ... get the setting for including nested collections (not relevant for other reports)
                showNested = dlgFilter.GetShowNestedData()
            # If we have a report other than based on a Collection ...
            else:
                # ... nesting is meaningless, so we can just initialize this variable to False.
                showNested = False

            # Get the user-specified File Name
            fs = self.exportFile.GetValue()
            # Ensure that the file name has the proper extension.
            if fs[-4:].lower() != '.txt':
                fs = fs + '.txt'
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
            # Open the output file for writing.
            f = codecs.open(fs, 'w', 'utf8')    # file(fs, 'w')

            prompt = unicode(_('Collection Name\tItem Type\tItem Name\tSource File\tStart\tStop\tLength'), 'utf8')
            # Write the Header line.  We're creating a tab-delimited file, so we'll use tabs to separate the items.
            f.write(prompt)
            # Add keywords to the Header.  Iterate through the Keyword List.
            for keyword in keywordList:
                # See if the user has left the keyword "checked" in the filter dialog.
                if keyword[2]:
                    # Encode and write all "checked" keywords to the Header.
                    kwg = keyword[0]
                    kw = keyword[1]
                    f.write('\t%s : %s' % (kwg, kw))
            # Add a line break to signal the end of the Header line. 
            f.write('\n')

            # Now iterate through the Quote List
            for quoteRec in quoteList:
                # See if the user has left the Quote "checked" in the filter dialog.
                # Also, if we are using a collection report, either Nested Data should be requested OR the current
                # Quote should be from the main collection if it is to be included in the report.
                if quoteRec[2] and ((self.collectionNum == 0) or (showNested) or (quoteRec[1] == self.collectionNum)):
                    # Load the Quote data.  The QuoteLookup dictionary allows this easily.
                    # No need to load the Quote Text, which can be slow to load.
                    quote = Quote.Quote(quoteLookup[quoteRec[0], quoteRec[1]], skipText=True)
                    # Get the collection the Quote is from.
                    collection = Collection.Collection(quote.collection_num)
                    # Encode string values using the Export Encoding
                    collectionID = collection.GetNodeString()
                    quoteID = quote.id
                    try:
                        document = Document.Document(quote.source_document_num)
                        documentID = document.id
                        quoteSourceFilename = document.imported_file
                        # If we're doing a Library report, we need the Quote's source document and Library for Document Filter comparison.
                        if self.libraryNum != 0:
                            library = Library.Library(document.library_num)
                            libraryID = library.id
                    # If we have an orphaned Quote ...
                    except TransanaExceptions.RecordNotFoundError, e:
                        # ... then we don't know these values!
                        documentID = ''
                        quoteSourceFilename = _('Source Document unknown')
                        libraryID = 0
                        
                    # Implement Document filtering if needed.  If we have a Library Report, we need to confirm that the Source Document
                    # is "checked" in the filter list.  (If we don't have a Library Report, this check isn't needed.)
                    if (self.libraryNum == 0) or ((documentID == '') and (libraryID == '')) or ((documentID, libraryID, True) in documentList):
                        # Write the Quote's data values to the output file.  We're creating a tab-delimited file,
                        # so we'll use tabs to separate the items.
                        f.write('%s\t%s\t%s\t%s\t%s\t%s\t%d' % (collectionID, '1', quoteID, quoteSourceFilename,
                                                            quote.start_char, quote.end_char,
                                                            (quote.end_char - quote.start_char)))

                        # Now we iterate through the keyword list ...
                        for keyword in keywordList:
                            # ... looking only at those keywords the user left "checked" in the filter dialog ...
                            if keyword[2]:
                                # ... and check to see if the Quote HAS the keyword.
                                if quote.has_keyword(keyword[0], keyword[1]):
                                    # If so, we write a "1", indicating True.
                                    f.write('\t1')
                                else:
                                    # If not, we write a "0", indicating False.
                                    f.write('\t0')
                        # Add a line break to signal the end of the Quote record
                        f.write('\n')

            # Now iterate through the Clip List
            for clipRec in clipList:
                # See if the user has left the clip "checked" in the filter dialog.
                # Also, if we are using a collection report, either Nested Data should be requested OR the current
                # clip should be from the main collection if it is to be included in the report.
                if clipRec[2] and ((self.collectionNum == 0) or (showNested) or (clipRec[1] == self.collectionNum)):
                    # Load the Clip data.  The ClipLookup dictionary allows this easily.
                    # No need to load the Clip Transcripts, which can be slow to load.
                    clip = Clip.Clip(clipLookup[clipRec[0], clipRec[1]], skipText=True)
                    # Get the collection the clip is from.
                    collection = Collection.Collection(clip.collection_num)
                    # Encode string values using the Export Encoding
                    collectionID = collection.GetNodeString()
                    clipID = clip.id
                    clipMediaFilename = clip.media_filename
                    # If we're doing a Library report, we need the clip's source episode and Library for Episode Filter comparison.
                    if self.libraryNum != 0:
                        episode = Episode.Episode(clip.episode_num)
                        library = Library.Library(episode.series_num)
                    # Implement Episode filtering if needed.  If we have a Library Report, we need to confirm that the Source Episode
                    # is "checked" in the filter list.  (If we don't have a Library Report, this check isn't needed.)
                    if (self.libraryNum == 0) or ((episode.id, library.id, True) in episodeList):
                        # Write the Clip's data values to the output file.  We're creating a tab-delimited file,
                        # so we'll use tabs to separate the items.
                        f.write('%s\t%s\t%s\t%s\t%s\t%s\t%10.4f' % (collectionID, '2', clipID, clipMediaFilename,
                                                            Misc.time_in_ms_to_str(clip.clip_start), Misc.time_in_ms_to_str(clip.clip_stop),
                                                            (clip.clip_stop - clip.clip_start) / 1000.0))

                        # Now we iterate through the keyword list ...
                        for keyword in keywordList:
                            # ... looking only at those keywords the user left "checked" in the filter dialog ...
                            if keyword[2]:
                                # ... and check to see if the Clip HAS the keyword.
                                if clip.has_keyword(keyword[0], keyword[1]):
                                    # If so, we write a "1", indicating True.
                                    f.write('\t1')
                                else:
                                    # If not, we write a "0", indicating False.
                                    f.write('\t0')
                        # Add a line break to signal the end of the Clip record
                        f.write('\n')

            # Flush the output file's buffer (probably unnecessary)
            f.flush()
            # Close the output file.
            f.close()
            # Restore the cursor when we're done.
            TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
            # If so, create a prompt to inform the user and ask to overwrite the file.
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Clip Data has been exported to file "%s".'), 'utf8')
            else:
                prompt = _('Clip Data has been exported to file "%s".')
            # Create the dialog to inform the user
            dlg2 = Dialogs.InfoDialog(self, prompt % fs)
            # Show the Info dialog.
            dlg2.ShowModal()
            # Destroy the Information dialog
            dlg2.Destroy()                    

        # Destroy the Filter Dialog.  We're done with it.
        dlgFilter.Destroy()
        
    def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        fs = wx.FileSelector(_("Select a text file for export"),
                        TransanaGlobal.configData.videoPath,
                        "",
                        "", 
                        _("Text Files (*.txt)|*.txt|All files (*.*)|*.*"), 
                        wx.SAVE)
        # If user didn't cancel ..
        if fs != "":
            self.exportFile.SetValue(fs)


# This simple derrived class let's the user drop files onto an edit box
class EditBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, editbox):
        wx.FileDropTarget.__init__(self)
        self.editbox = editbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        self.editbox.SetValue(files[0])
