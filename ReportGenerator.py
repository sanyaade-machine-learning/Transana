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

"""This module implements the Report Generator. """
# This module combines functionality that used to be divided into Collection Summary Report and
# Keyword Usage Report modules.

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

DEBUG = False

# import Python's datetime module
import datetime
# import Python's locale module
import locale
# import the Python String module
import string
# import Python's sys module
import sys
# import the python xml-sax module
import xml.sax
# import wxPython
import wx
# import wxPython's Rich Text Ctrl module
import wx.richtext as richtext
# Import Transana's Clip object
import Clip
# import Transana's Collection Object
import Collection
# Import Transana's Database Interface
import DBInterface
# Import Transana's Dialog Boxes
import Dialogs
# import Transana's Document Object
import Document
# Import Transana's Episode object
import Episode
# import Transana's Filter Dialog
import FilterDialog
# Import Transana's Keyword object
import KeywordObject as Keyword
# Import Transana's Miscellaneous functions
import Misc
# Import Transana's Note object
import Note
# import the Transana XML-to-RTC Import Parser
import PyXML_RTCImportParser
# import Transana's Quote Object
import Quote
# Import Transana's Library Object
import Library
# import Transana's Snapshot Object
import Snapshot
# Import Transana's Snapshot Window for loading images
import SnapshotWindow
# Import Transana's Text Report infrastructure
import TextReport
# Import Transana's Constants
import TransanaConstants
# import Transana's Exceptions
import TransanaExceptions
# import Transana's Global variables
import TransanaGlobal
# import Transana's Transcript Object
import Transcript
# import Transana's Transcript Editor - Rich Text Ctrl version
import TranscriptEditor_RTC


class ReportGenerator(wx.Object):
    """ This class creates and displays the Object Reports, formerly the Keyword Usage Report and the Collection Summary Report """
    def __init__(self, **kwargs):
        """ Create the Object Report
              If a seriesName is passed, all Episodes and Document in that *Library* and their keywords should be listed.
              If a documentName is passed, all Quotes from that Document, regardless of Collection, and their Document Keywords should be listed.
              If an episodeName is passed, all Clips from that Episode, regardless of Collection, and their Clip keywords should be listed.
              If a collection is passed, all Clips in that Collection, regardless of source Episode, and their Clip keywords should be listed.
              If a searchSeries is passed, use the treeCtrl to determine the Episodes and Documents that should be included.
              if a searchCollection is passed, use the treeCtrl to determine the Clips that should be included. """
        # Parameters can include:
        # controlObject=None
        # title=''
        # seriesName=None,
        # documentName=None,
        # episodeName=None,
        # collection=None,
        # searchSeries=None,
        # searchColl=None,
        # treeCtrl=None,
        # showNested=False,
        # showHypertext=False,
        # showFile=True,
        # showTime=True,
        # showDocImportDate=True,
        # showSourceInfo=True,
        # showQuoteText=True,
        # showTranscripts=False,
        # showSnapshotImage=True,
        # showSnapshotCoding=True,
        # showKeywords=False,
        # showComments=False,
        # showCollectionNotes=False
        # showDocumentNotes=False
        # showClipNotes=False
        # showSnapshotNotes=False

        # Remember the parameters passed in and set values for all variables, even those NOT passed in.
        if kwargs.has_key('controlObject'):
            self.ControlObject = kwargs['controlObject']
        else:
            self.ControlObject = None
        # Specify the Report Title
        if kwargs.has_key('title'):
            self.title = kwargs['title']
        else:
            self.title = ''
        if kwargs.has_key('seriesName'):
            self.seriesName = kwargs['seriesName']
        else:
            self.seriesName = None
        if kwargs.has_key('documentName'):
            self.documentName = kwargs['documentName']
        else:
            self.documentName = None
        if kwargs.has_key('episodeName'):
            self.episodeName = kwargs['episodeName']
        else:
            self.episodeName = None
        if kwargs.has_key('collection'):
            self.collection = kwargs['collection']
        else:
            self.collection = None
        if kwargs.has_key('searchSeries'):
            self.searchSeries = kwargs['searchSeries']
        else:
            self.searchSeries = None
        if kwargs.has_key('searchColl'):
            self.searchColl = kwargs['searchColl']
        else:
            self.searchColl = None
        if kwargs.has_key('treeCtrl'):
            self.treeCtrl = kwargs['treeCtrl']
        else:
            self.treeCtrl = None
        if kwargs.has_key('showNested') and kwargs['showNested']:
            self.showNested = True
        else:
            self.showNested = False
        if kwargs.has_key('showHyperlink') and kwargs['showHyperlink']:
            self.showHyperlink = True
        else:
            self.showHyperlink = False
        if kwargs.has_key('showFile') and kwargs['showFile']:
            self.showFile = True
        else:
            self.showFile = False
        if kwargs.has_key('showTime') and kwargs['showTime']:
            self.showTime = True
        else:
            self.showTime = False
        if kwargs.has_key('showDocImportDate') and kwargs['showDocImportDate']:
            self.showDocImportDate = True
        else:
            self.showDocImportDate = False
        if kwargs.has_key('showSourceInfo') and kwargs['showSourceInfo']:
            self.showSourceInfo = True
        else:
            self.showSourceInfo = False
        if kwargs.has_key('showQuoteText') and kwargs['showQuoteText']:
            self.showQuoteText = True
        else:
            self.showQuoteText = False
        if kwargs.has_key('showTranscripts') and kwargs['showTranscripts']:
            self.showTranscripts = True
        else:
            self.showTranscripts = False
        if kwargs.has_key('showSnapshotImage'):
            self.showSnapshotImage = kwargs['showSnapshotImage']
        else:
            self.showSnapshotImage = 0
        if kwargs.has_key('showSnapshotCoding'):
            self.showSnapshotCoding = kwargs['showSnapshotCoding']
        else:
            self.showSnapshotCoding = 0
        if kwargs.has_key('showKeywords') and kwargs['showKeywords']:
            self.showKeywords = True
        else:
            self.showKeywords = False
        if kwargs.has_key('showComments') and kwargs['showComments']:
            self.showComments = True
        else:
            self.showComments = False
        if kwargs.has_key('showCollectionNotes') and kwargs['showCollectionNotes']:
            self.showCollectionNotes = True
        else:
            self.showCollectionNotes = False
        if kwargs.has_key('showQuoteNotes') and kwargs['showQuoteNotes']:
            self.showQuoteNotes = True
        else:
            self.showQuoteNotes = False
        if kwargs.has_key('showClipNotes') and kwargs['showClipNotes']:
            self.showClipNotes = True
        else:
            self.showClipNotes = False
        if kwargs.has_key('showSnapshotNotes') and kwargs['showSnapshotNotes']:
            self.showSnapshotNotes = True
        else:
            self.showSnapshotNotes = False

        # Filter Configuration Name -- initialize to nothing
        self.configName = ''

        # Get the local locale, which will set the appropriate date formatting for the %x parameter below.
        locale.setlocale(locale.LC_ALL, '')

        # Create the TextReport object, which forms the basis for text-based reports.
        self.report = TextReport.TextReport(None, title=self.title, displayMethod=self.OnDisplay,
                                            filterMethod=self.OnFilter, helpContext="Transana's Text Reports")
        # If a Control Object has been passed in ...
        if self.ControlObject != None:
            # ... register this report with the Control Object (which adds it to the Windows Menu)
            self.ControlObject.AddReportWindow(self.report)
            # Register the Control Object with the Report
            self.report.ControlObject = self.ControlObject
        # Define the main (Episode or Clip) Filter List (which will differ depending on the report type)
        self.filterList = []
        # Define the Document Filter List, which will only be used for some reports
        self.documentFilterList = []
        # Define the Quote Filter List.  This probably is redundant with the filterList, I'm not sure yet.
        self.quoteFilterList = []
        # Define the Snapshot Filter List, which will only be used for some reports
        self.snapshotFilterList = []
        # Define the Keyword Filter List as well, which does NOT differ based on report type
        self.keywordFilterList = []
        # To speed report creation, freeze GUI updates based on changes to the report text
        self.report.reportText.Freeze()
        # Trigger the ReportText method that causes the report to be displayed.
        self.report.CallDisplay()
        # Apply the Default Filter, if one exists
        self.report.OnFilter(None)
        # Now that we're done, remove the freeze
        self.report.reportText.Thaw()

    def OnDisplay(self, reportText):
        """ This method, required by TextReport, populates the TextReport.  The reportText parameter is
            the wxSTC control from the TextReport object.  It needs to be in the report parent because
            the TextReport doesn't know anything about the actual data.  """
        # Create minorList as a blank Dictionary Object
        minorList = {}
        # We need variables to count the number of quotes displayed and to accumulate their total length.
        self.quoteCount = 0
        self.quoteTotalLength = 0
        # We need variables to count the number of clips displayed and to accumulate their total time.
        self.clipCount = 0
        self.clipTotalTime = 0.0
        # We need variables to count the number of snapshots displayed.
        self.snapshotCount = 0
        self.snapshotTotalTime = 0.0
        # Determine if we need to populate the Filter Lists.  If it hasn't already been done, we should do it.
        # If it has already been done, no need to do it again.
        if (self.filterList == []) and (self.documentFilterList == []) and \
           (self.quoteFilterList == []) and (self.snapshotFilterList == []):
            populateFilterList = True
        else:
            populateFilterList = False
        # Initialize the variable that tracks the request to skip "Image Not Loaded" errors
        skippingImageError = False

        # Make the control writable
        reportText.SetReadOnly(False)

        # ... Set the Style for the Heading
        reportText.SetTxtStyle(fontFace='Courier New', fontSize=16, fontBold=True, fontUnderline=True)
        # Set report margins, the left and right margins to 0.  The RichTextPrinting infrastructure handles that!
        # Center the title, and add spacing after.
        reportText.SetTxtStyle(parLeftIndent = 0, parRightIndent = 0, parAlign=wx.TEXT_ALIGNMENT_CENTER,
                               parSpacingBefore = 0, parSpacingAfter = 12)
        # Add the Title to the page
        reportText.WriteText(self.title + '\n')

        # If a Collection is passed in ...
        if self.collection != None:
            # The major label for objects for the Collection report is Clip
            majorLabel = _('Clip')
            # An empty Collection (number == 0) signals the request for the Global report
            if self.collection.number == 0:
                # The global report has no subtitle.
                self.subtitle = ''
                # We initialize the major list to an empty list for the global report
                majorList = []
            # A non-empty collection signals a scoped report.
            else:
                # Add a subtitle and ...
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("Collection: %s"), 'utf8')
                else:
                    prompt = _("Collection: %s")
                self.subtitle = prompt % self.collection.GetNodeString()
                
                # initialize the Major List of report elements.  Clips and Snapshots will be sorted and included.
                majorList = []
                # initialize a Dictionary for all Report Artifacts (for sorting!)
                tmpDict = {}
                
                # Get a list of all Clips in the Collection.
                tmpClipList = DBInterface.list_of_clips_by_collectionnum(self.collection.number, includeSortOrder=True)
                # For each Clip ...
                for x in tmpClipList:
                    # ... add the Clip to the Dictionary with the Sort Order as the key and with the Object Type added
                    tmpDict[x[3]] = (('Clip',) + x)

                if TransanaConstants.proVersion:
                    # Get a list of all Quotes in the Collection.
                    tmpQuoteList = DBInterface.list_of_quotes_by_collectionnum(self.collection.number, includeSortOrder=True)
                    # For each Quote ...
                    for x in tmpQuoteList:
                        # ... add the Quote to the Dictionary with the Sort Order as the key and with the Object Type added
                        tmpDict[x[3]] = (('Quote',) + x)

                    # Get a list of all Snapshots in the Collection.
                    tmpSnapshotList = DBInterface.list_of_snapshots_by_collectionnum(self.collection.number, includeSortOrder=True)
                    # For each Snapshot ...
                    for x in tmpSnapshotList:
                        # ... add the Snapshot to the Dictionary with the Sort Order as the key and with the Object Type added
                        tmpDict[x[3]] = (('Snapshot',) + x)
                        
                # Get the Dictionary's Keys
                order = tmpDict.keys()
                # Sort the Dictionary's Keys
                order.sort()
                # For each element in the sorted list of Keys ...
                for x in order:
                    # Add the elemnt to the Major List.
                    majorList.append(tmpDict[x][:-1])

            # If we're supposed to show Nested Collection data ...
            if self.showNested:
                # ... we first need to get the nested collections for the top level
                nestedCollections = DBInterface.list_of_collections(self.collection.number)
                # As long as there are entries in the list of nested collections that haven't been processed ...
                while len(nestedCollections) > 0:
                    # ... extract the data from the top of the nested collection list ...
                    (collNum, collName, parentCollNum) = nestedCollections[0]
                    # ... and remove that entry from the list.
                    del(nestedCollections[0])
                    # initialize a Dictionary for all Report Artifacts (for sorting!)
                    tmpDict = {}
                    # Get a list of all Clips in the Collection.
                    tmpClipList = DBInterface.list_of_clips_by_collectionnum(collNum, includeSortOrder=True)
                    # For each Clip ...
                    for x in tmpClipList:
                        # ... add the Clip to the Dictionary with the Sort Order as the key and with the Object Type added
                        tmpDict[x[3]] = (('Clip',) + x)
                    if TransanaConstants.proVersion:
                        # Get a list of all Quotes in the Collection.
                        tmpQuoteList = DBInterface.list_of_quotes_by_collectionnum(collNum, includeSortOrder=True)
                        # For each Quote ...
                        for x in tmpQuoteList:
                            # ... add the Quote to the Dictionary with the Sort Order as the key and with the Object Type added
                            tmpDict[x[3]] = (('Quote',) + x)
                        # Get a list of all Snapshots in the Collection.
                        tmpSnapshotList = DBInterface.list_of_snapshots_by_collectionnum(collNum, includeSortOrder=True)
                        # For each Snapshot ...
                        for x in tmpSnapshotList:
                            # ... add the Snapshot to the Dictionary with the Sort Order as the key and with the Object Type added
                            tmpDict[x[3]] = (('Snapshot',) + x)
                    # Get the Dictionary's Keys
                    order = tmpDict.keys()
                    # Sort the Dictionary's Keys
                    order.sort()
                    # For each element in the sorted list of Keys ...
                    for x in order:
                        # Add the elemnt to the Major List.
                        majorList.append(tmpDict[x][:-1])

                    # Then get the nested collections under the new collection and add them to the Nested Collection list
                    # They get added at the FRONT of the list so that the report will mirror the organization of the
                    # database Tree.
                    nestedCollections = DBInterface.list_of_collections(collNum) + nestedCollections

            # Put all the Keywords for the Clips and Snapshots in the majorList in the minorList.
            # Start by iterating through the Major List
            for (objType, objNo, objName, collNo) in majorList:
                # Create a Minor List dictionary entry, indexed to clip or snapshot number, for the keywords.
                if objType == 'Quote':
                    minorList[(objType, objNo)] = DBInterface.list_of_keywords(Quote = objNo)
                elif objType == 'Clip':
                    minorList[(objType, objNo)] = DBInterface.list_of_keywords(Clip = objNo)
                elif objType == 'Snapshot':
                    minorList[(objType, objNo)] = DBInterface.list_of_keywords(Snapshot = objNo)
                # If we're populating Filter Lists ...
                if populateFilterList:
                    # If we have a Quote ...
                    if objType == 'Quote':
                        # ... add it to the Quote Filter List
                        listToPopulate = self.quoteFilterList
                    # If we have a Clip ...
                    elif objType == 'Clip':
                        # ... add it to the regular Filter List
                        listToPopulate = self.filterList
                    # If we have a Snapshot ...
                    elif objType == 'Snapshot':
                        # ... add it to the Snapshot Filter List
                        listToPopulate = self.snapshotFilterList
                    # ... then add the Artifact data to the appropiate Filter List, initially checked ...
                    listToPopulate.append((objName, collNo, True))
                    # ... and iterate through that clip's keywords or the snapshot's whole snapshot keywords ...
                    for (kwg, kw, ex) in minorList[(objType, objNo)]:
                        # ... check to see if the entry is NOT already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... and add the keyword entry to the Keyword Filter List if it's not already there.
                            self.keywordFilterList.append((kwg, kw, True))

                    # If we have a Snapshot ...
                    if objType == 'Snapshot':
                        # ... get a list of the Snapshot's Detail Coding
                        tmpList = DBInterface.list_of_snapshot_detail_keywords(Snapshot = objNo)
                        # For each Keyword Group : Keyword pair ...
                        for (kwg, kw) in tmpList:
                            # ... check to see if the entry is NOT already in the list ...
                            if (kwg, kw, True) not in self.keywordFilterList:
                                # ... and add the keyword entry to the Keyword Filter List if it's not already there.
                                self.keywordFilterList.append((kwg, kw, True))

        # If a Document Name is passed in ...
        elif self.documentName != None:
            # ...  add a subtitle
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Document: %s"), 'utf8')
            else:
                prompt = _("Episode: %s")
            self.subtitle = prompt % self.documentName
            # First, get the Document Object ...
            docObj = Document.Document(libraryID=self.seriesName, documentID=self.documentName)

            # initialize the Major List of report elements.  Quotes and Snapshots will be sorted and included.
            majorList = []
            # initialize a Dictionary for all Report Artifacts (for sorting!)
            tmpDict = {}
                
            # Get a list of all Quotes created from the Document
            tmpQuoteList = DBInterface.list_of_quotes_by_document(docObj.number)
            # For each Quote ...
            for x in tmpQuoteList:
                # ... specify that this is a Quote
                x['Type'] = 'Quote'
                # ... add the Quote to the Dictionary with the Sort Order as the key and with the Object Type added
                tmpDict[(x['StartChar'], x['EndChar'], x['CollectID'], x['CollectNum'], x['QuoteID'], 'Quote')] = x

##            if TransanaConstants.proVersion:
##                # Get a list of all Snapshots in the Collection.
##                tmpSnapshotList = DBInterface.list_of_snapshots_by_episode(epObj.number)
##                # For each Snapshot ...
##                for x in tmpSnapshotList:
##                    # ... specify that this is a Snapshot
##                    x['Type'] = 'Snapshot'
##                    # ... add the Snapshot to the Dictionary with the Sort Order as the key and with the Object Type added
##                    tmpDict[(x['SnapshotStart'], x['SnapshotStop'], x['CollectID'], x['CollectNum'], x['SnapshotID'], 'Snapshot')] = x
            # Get the Dictionary's Keys
            order = tmpDict.keys()
            # Sort the Dictionary's Keys
            order.sort()
            # For each element in the sorted list of Keys ...
            for x in order:
                # Add the elemnt to the Major List.
                majorList.append(tmpDict[x])

            # Put all the Keywords for the Quotes and Snapshots in the majorList in the minorList.
            # Start by iterating through the Major List
            for item in majorList:
                # Create a Minor List dictionary entry, indexed to quote or snapshot number, for the keywords.
                if item['Type'] == 'Quote':
                    minorList[(item['Type'], item['QuoteNum'])] = DBInterface.list_of_keywords(Quote = item['QuoteNum'])
##                elif item['Type'] == 'Snapshot':
##                    minorList[(item['Type'], item['SnapshotNum'])] = DBInterface.list_of_keywords(Snapshot = item['SnapshotNum'])
                # If we're populating Filter Lists ...
                if populateFilterList:
                    # If we have a Snapshot ...
                    if item['Type'] == 'Snapshot':
                        # ... add it to the Snapshot Filter List
                        listToPopulate = self.snapshotFilterList
                        objName = item['SnapshotID']
                        objNo = item['SnapshotNum']
                    # If we DON'T have a Snapshot ...
                    else:
                        # ... add it to the Quote Filter List
                        listToPopulate = self.quoteFilterList
                        objName = item['QuoteID']
                        objNo = item['QuoteNum']
                    # ... then add the Artifact data to the appropiate Filter List, initially checked ...
                    listToPopulate.append((objName, item['CollectNum'], True))
                    # ... and iterate through that quote's keywords or the snapshot's whole snapshot keywords ...
                    for (kwg, kw, ex) in minorList[(item['Type'], objNo)]:
                        # ... check to see if the entry is NOT already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... and add the keyword entry to the Keyword Filter List if it's not already there.
                            self.keywordFilterList.append((kwg, kw, True))

##                    # If we have a Snapshot ...
##                    if item['Type'] == 'Snapshot':
##                        # ... get a list of the Snapshot's Detail Coding
##                        tmpList = DBInterface.list_of_snapshot_detail_keywords(Snapshot = item['SnapshotNum'])
##                        # For each Keyword Group : Keyword pair ...
##                        for (kwg, kw) in tmpList:
##                            # ... check to see if the entry is NOT already in the list ...
##                            if (kwg, kw, True) not in self.keywordFilterList:
##                                # ... and add the keyword entry to the Keyword Filter List if it's not already there.
##                                self.keywordFilterList.append((kwg, kw, True))

        # If an Episode Name is passed in ...
        elif self.episodeName != None:
            # ...  add a subtitle
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Episode: %s"), 'utf8')
            else:
                prompt = _("Episode: %s")
            self.subtitle = prompt % self.episodeName
            # First, get the Episode Object ...
            epObj = Episode.Episode(series = self.seriesName, episode = self.episodeName)

            # initialize the Major List of report elements.  Clips and Snapshots will be sorted and included.
            majorList = []
            # initialize a Dictionary for all Report Artifacts (for sorting!)
            tmpDict = {}
                
            # Get a list of all Clips created from the Episode
            tmpClipList = DBInterface.list_of_clips_by_episode(epObj.number)
            # For each Clip ...
            for x in tmpClipList:
                # ... specify that this is a Clip
                x['Type'] = 'Clip'
                # ... add the Clip to the Dictionary with the Sort Order as the key and with the Object Type added
                tmpDict[(x['ClipStart'], x['ClipStop'], x['CollectID'], x['CollectNum'], x['ClipID'], 'Clip')] = x

            if TransanaConstants.proVersion:
                # Get a list of all Snapshots in the Collection.
                tmpSnapshotList = DBInterface.list_of_snapshots_by_episode(epObj.number)
                # For each Snapshot ...
                for x in tmpSnapshotList:
                    # ... specify that this is a Snapshot
                    x['Type'] = 'Snapshot'
                    # ... add the Snapshot to the Dictionary with the Sort Order as the key and with the Object Type added
                    tmpDict[(x['SnapshotStart'], x['SnapshotStop'], x['CollectID'], x['CollectNum'], x['SnapshotID'], 'Snapshot')] = x
            # Get the Dictionary's Keys
            order = tmpDict.keys()
            # Sort the Dictionary's Keys
            order.sort()
            # For each element in the sorted list of Keys ...
            for x in order:
                # Add the elemnt to the Major List.
                majorList.append(tmpDict[x])

            # Put all the Keywords for the Clips and Snapshots in the majorList in the minorList.
            # Start by iterating through the Major List
            for item in majorList:
                # Create a Minor List dictionary entry, indexed to clip or snapshot number, for the keywords.
                if item['Type'] == 'Clip':
                    minorList[(item['Type'], item['ClipNum'])] = DBInterface.list_of_keywords(Clip = item['ClipNum'])
                elif item['Type'] == 'Snapshot':
                    minorList[(item['Type'], item['SnapshotNum'])] = DBInterface.list_of_keywords(Snapshot = item['SnapshotNum'])
                # If we're populating Filter Lists ...
                if populateFilterList:
                    # If we have a Snapshot ...
                    if item['Type'] == 'Snapshot':
                        # ... add it to the Snapshot Filter List
                        listToPopulate = self.snapshotFilterList
                        objName = item['SnapshotID']
                        objNo = item['SnapshotNum']
                    # If we DON'T have a Snapshot ...
                    else:
                        # ... add it to the regular Filter List
                        listToPopulate = self.filterList
                        objName = item['ClipID']
                        objNo = item['ClipNum']
                    # ... then add the Artifact data to the appropiate Filter List, initially checked ...
                    listToPopulate.append((objName, item['CollectNum'], True))
                    # ... and iterate through that clip's keywords or the snapshot's whole snapshot keywords ...
                    for (kwg, kw, ex) in minorList[(item['Type'], objNo)]:
                        # ... check to see if the entry is NOT already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... and add the keyword entry to the Keyword Filter List if it's not already there.
                            self.keywordFilterList.append((kwg, kw, True))

                    # If we have a Snapshot ...
                    if item['Type'] == 'Snapshot':
                        # ... get a list of the Snapshot's Detail Coding
                        tmpList = DBInterface.list_of_snapshot_detail_keywords(Snapshot = item['SnapshotNum'])
                        # For each Keyword Group : Keyword pair ...
                        for (kwg, kw) in tmpList:
                            # ... check to see if the entry is NOT already in the list ...
                            if (kwg, kw, True) not in self.keywordFilterList:
                                # ... and add the keyword entry to the Keyword Filter List if it's not already there.
                                self.keywordFilterList.append((kwg, kw, True))

        # If a Library Name is passed in ...            
        elif self.seriesName != None:
            # ...  add a subtitle
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Library: %s"), 'utf8')
            else:
                prompt = _("Library: %s")
            self.subtitle = prompt % self.seriesName
            # The label for our Major unit should reflect that these are Episodes
            majorLabel = _('Episode')
            # Initialize the Major List
            majorList = []

            # Get the Library Object
            tmpLibraryObj = Library.Library(self.seriesName)
            # Get a Dictionary of all items in this Library
            tempDict = DBInterface.dictionary_of_documents_and_episodes(tmpLibraryObj)
            # Get the keys to the dictionary
            keys = tempDict.keys()
            # Sort the keys so the report will be displayed in the correct order
            keys.sort()
            # For each Key in the data list ...
            for key in keys:
                # ... get the data object's Name from the dictionary Key ...
                objName = key[0]
                # ... and get the Object's Type, Number, and Parent Number from the dictionary Value
                (objType, objNum, objParentNum) = tempDict[key]
                # Put the Item in the Major List
                majorList.append((objType, objNum, objName, objParentNum))
                # If we have a Document ...
                if objType == 'Document':
                    # Put all the Keywords for the Document in the majorList in the minorList
                    minorList[(objType, objNum)] = DBInterface.list_of_keywords(Document = objNum)
                # If we have an Episode ...
                elif objType == 'Episode':
                    # Put all the Keywords for the Episodes in the majorList in the minorList
                    minorList[(objType, objNum)] = DBInterface.list_of_keywords(Episode = objNum)
                # If we're populating the Filter Lists ...
                if populateFilterList:
                    if objType == 'Document':
                        # ... Add the Document data to the document Filter List ...
                        self.documentFilterList.append((objName, self.seriesName, True))
                    else:
                        # ... Add the Episode data to the main Filter List ...
                        self.filterList.append((objName, self.seriesName, True))
                    # ... Iterate through the keywords that were just added to the Minor List (only for this Key) ...
                    for (kwg, kw, ex) in minorList[(objType, objNum)]:
                        # .. and IF they're not already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... add them to the Keyword Filter List
                            self.keywordFilterList.append((kwg, kw, True))

        # If this report is called for a SearchLibraryResult, we build the majorList based on the contents of the Tree Control.
        elif (self.searchSeries != None) and (self.treeCtrl != None):
            # Get the Search Result Name for the subtitle
            searchResultNode = self.searchSeries
            # Start a loop to move up the tree.  Keep going until interrupted.
            while True:
                # Move up to the Parent of the current node
                searchResultNode = self.treeCtrl.GetItemParent(searchResultNode)
                # Get the Data for the new node
                tempData = self.treeCtrl.GetPyData(searchResultNode)
                # If we are at the SearchResultsNode ...
                if tempData.nodetype == 'SearchResultsNode':
                    # ... we can stop moving up.  Break the loop.
                    break
            # Now build the subtitle
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Search Result: %s  Library: %s"), 'utf8')
            else:
                prompt = _("Search Result: %s  Library: %s")
            self.subtitle = prompt % (self.treeCtrl.GetItemText(searchResultNode), self.treeCtrl.GetItemText(self.searchSeries))
            # The majorLabel is for Episodes in this case
            majorLabel = _('Episode')
            # Initialize the majorList to an empty list
            majorList = []
            # Get the first Child node from the searchColl collection
            (item, cookie) = self.treeCtrl.GetFirstChild(self.searchSeries)
            # Process all children in the searchLibrary Library.  (IsOk() fails when all children are processed.)
            while item.IsOk():
                # Get the child item's Name
                itemText = self.treeCtrl.GetItemText(item)
                # Get the child item's Node Data
                itemData = self.treeCtrl.GetPyData(item)
                # See if the item is a Document
                if itemData.nodetype == 'SearchDocumentNode':
                    # If it's a Document, add the Document's Node Data to the majorList
                    majorList.append(('Document', itemData.recNum, itemText, itemData.parent))
                    # If we're populating the Filter Lists ...
                    if populateFilterList:
                        # ... Add the Episode data to the main Filter List ...
                        self.documentFilterList.append((itemText, self.treeCtrl.GetItemText(self.treeCtrl.GetItemParent(item)), True))
                # See if the item is an Episode
                elif itemData.nodetype == 'SearchEpisodeNode':
                    # If it's an Episode, add the Episode's Node Data to the majorList
                    majorList.append(('Episode', itemData.recNum, itemText, itemData.parent))
                    # If we're populating the Filter Lists ...
                    if populateFilterList:
                        # ... Add the Episode data to the main Filter List ...
                        self.filterList.append((itemText, self.treeCtrl.GetItemText(self.treeCtrl.GetItemParent(item)), True))
                # Get the next Child Item and continue the loop
                (item, cookie) = self.treeCtrl.GetNextChild(self.searchSeries, cookie)

##            print "ReportGenerator.OnDisplay():  Search Library Report"
##            print "majorList:"
##            for x in range(len(majorList)):
##                print x, majorList[x]
##            print

            # Once we have the Episodes in the majorList, we can gather their keywords into the minorList.
            # Start by iterating through the Major List
            for (objType, EpNo, epName, epParentNo) in majorList:
                # If we have a Document ...
                if objType == 'Document':
                    # Get all the keywords for the indicated Document and add them to the Minor List, keyed to the Document Name.
                    minorList[('Document', EpNo)] = DBInterface.list_of_keywords(Document = EpNo)
                # If we have an Episode ...
                elif objType == 'Episode':
                    # Get all the keywords for the indicated Episode and add them to the Minor List, keyed to the Episode Name.
                    minorList[('Episode', EpNo)] = DBInterface.list_of_keywords(Episode = EpNo)
                # If we're populating the Filter Lists ...
                if populateFilterList:
                    # ... Iterate through the keywords that were just added to the Minor List (only for this Key) ...
                    for (kwg, kw, ex) in minorList[(objType, EpNo)]:
                        # .. and IF they're not already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... add them to the Keyword Filter List
                            self.keywordFilterList.append((kwg, kw, True))

##            print "minorList:"
##            for x in range(len(minorList)):
##                print x, minorList[x]
##            print

        # If this report is called for a SearchCollectionResult, we build the majorList based on the contents of the Tree Control.
        elif (self.searchColl != None) and (self.treeCtrl != None):
            # Get the Search Result Name for the subtitle
            searchResultNode = self.searchColl
            # Start a loop to move up the tree.  Keep going until interrupted.
            while True:
                # Move up to the Parent of the current node
                searchResultNode = self.treeCtrl.GetItemParent(searchResultNode)
                # Get the Data for the new node
                tempData = self.treeCtrl.GetPyData(searchResultNode)
                # If we are at the SearchResultsNode ...
                if tempData.nodetype in ['SearchRootNode', 'SearchResultsNode']:
                    # ... we can stop moving up.  Break the loop.
                    break
            # Now build the subtitle
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Search Result: %s  Collection: %s"), 'utf8')
            else:
                prompt = _("Search Result: %s  Collection: %s")
            self.subtitle = prompt % (self.treeCtrl.GetItemText(searchResultNode), self.treeCtrl.GetItemText(self.searchColl))
            # The majorLabel is for Clips in this case
            majorLabel = _('Clip')
            # Initialize the majorList to an empty list
            majorList = []
            # Extracting data from the treeCtrl requires a "cookie" value, which is initialized to 0
            cookie = 0
            # Get the first Child node from the searchColl collection
            (item, cookie) = self.treeCtrl.GetFirstChild(self.searchColl)
            # Create an empty list for Nested Collections so we can recurse through them
            nestedCollections = []
            # While looking at children, we need a pointer to the parent node
            currentNode = self.searchColl
            # Process all children in the searchColl collection
            while item.IsOk():
                # Get the item's Name
                itemText = self.treeCtrl.GetItemText(item)
                # Get the item's Node Data
                itemData = self.treeCtrl.GetPyData(item)
                # See if the item is a Clip
                if itemData.nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']:
                    if itemData.nodetype == 'SearchQuoteNode':
                        objType = 'Quote'
                    elif itemData.nodetype == 'SearchClipNode':
                        objType = 'Clip'
                    elif itemData.nodetype == 'SearchSnapshotNode':
                        objType = 'Snapshot'
                    # If it's a Clip, add the Clip's Node Data to the majorList
                    majorList.append((objType, itemData.recNum, itemText, itemData.parent))
                    # If we're populating the Filter List ...
                    if populateFilterList:
                        # If we have a Quote ...
                        if objType == 'Quote':
                            # ... add it to the regular Filter List
                            listToPopulate = self.quoteFilterList
                        # If we have a Clip ...
                        elif objType == 'Clip':
                            # ... add it to the regular Filter List
                            listToPopulate = self.filterList
                        # If we have a Snapshot ...
                        elif objType == 'Snapshot':
                            # ... add it to the Snapshot Filter List
                            listToPopulate = self.snapshotFilterList
                        # ... then add the Artifact data to the appropiate Filter List, initially checked ...
                        listToPopulate.append((itemText, itemData.parent, True))
                # If we have a Collection Node ...
                elif self.showNested and (itemData.nodetype == 'SearchCollectionNode'):
                    # ... add it to the list of nested Collections to be processed
                    nestedCollections.append(item)

                # When we get to the last Child Item for the current node, ...
                if item == self.treeCtrl.GetLastChild(currentNode):
                    # ... check to see if there are nested collections that need to be processed.  If so ...
                    if len(nestedCollections) > 0:
                        # ... set the current node pointer to the first nested collection ...
                        currentNode = nestedCollections[0]
                        # ... get the first child node of the nested collection ...
                        (item, cookie) = self.treeCtrl.GetFirstChild(nestedCollections[0])
                        # ... and remove the nested collection from the list waiting to be processed
                        del(nestedCollections[0])
                    # If there are no nested collections to be processed ...
                    else:
                        # ... stop looping.  We're done.
                        break
                # If we're not at the Last Child Item ...
                else:
                    # ... get the next Child Item and continue the loop
                    (item, cookie) = self.treeCtrl.GetNextChild(currentNode, cookie)

            # Put all the Keywords for the Clips and Snapshots in the majorList in the minorList.
            # Start by iterating through the Major List
            for (objType, objNo, objName, collNo) in majorList:
                # Create a Minor List dictionary entry, indexed to clip or snapshot number, for the keywords.
                if objType == 'Quote':
                    minorList[(objType, objNo)] = DBInterface.list_of_keywords(Quote = objNo)
                elif objType == 'Clip':
                    minorList[(objType, objNo)] = DBInterface.list_of_keywords(Clip = objNo)
                elif objType == 'Snapshot':
                    minorList[(objType, objNo)] = DBInterface.list_of_keywords(Snapshot = objNo)
                # If we're populating Filter Lists ...
                if populateFilterList:
                    # ... and iterate through that clip's keywords or the snapshot's whole snapshot keywords ...
                    for (kwg, kw, ex) in minorList[(objType, objNo)]:
                        # ... check to see if the entry is NOT already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... and add the keyword entry to the Keyword Filter List if it's not already there.
                            self.keywordFilterList.append((kwg, kw, True))

                    # If we have a Snapshot ...
                    if objType == 'Snapshot':
                        # ... get a list of the Snapshot's Detail Coding
                        tmpList = DBInterface.list_of_snapshot_detail_keywords(Snapshot = objNo)
                        # For each Keyword Group : Keyword pair ...
                        for (kwg, kw) in tmpList:
                            # ... check to see if the entry is NOT already in the list ...
                            if (kwg, kw, True) not in self.keywordFilterList:
                                # ... and add the keyword entry to the Keyword Filter List if it's not already there.
                                self.keywordFilterList.append((kwg, kw, True))

        # ...  add a subtitle
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("Filter Configuration: %s"), 'utf8')
        else:
            prompt = _("Filter Configuration: %s")

        # There's a problem in wxPython 2.8.12.0 on Windows.  If you use too many format changes in too
        # long a document, you run out of Windows GDI resources.  This problem can be ameliorated (but not
        # totally eliminated) by reducing the number of times BOLD text is used.
        #
        # So if we're on Windows and have more than 350 elements in the report, DON'T use BOLD as much!
        # This flag is the signal.
        useBold = True   # not ((len(majorList) > 350) and  ('wxMSW' in wx.PlatformInfo))

        # If a subtitle is defined ...
        if self.subtitle != '':
            # ... set the subtitle font
            reportText.SetTxtStyle(fontSize=12, fontBold=False, fontUnderline=False, parSpacingBefore = 0, parSpacingAfter = 0)
            # Add the subtitle to the page
            reportText.WriteText(self.subtitle + '\n')
            # Finish the paragraph
#            reportText.Newline()
        if self.configName != '':
            self.configLine = prompt % self.configName
            # ... set the subtitle font
            reportText.SetTxtStyle(fontSize=10, fontBold=False, fontUnderline=False, parSpacingBefore = 0, parSpacingAfter = 0)
            # Add the subtitle to the page
            reportText.WriteText(self.configLine + '\n')

        # Initialize the initial data structure that will be turned into the report
        self.data = []
        # Create a Dictionary Data Structure to accumulate Keyword Counts
        keywordCounts = {}
        # Create a Dictionary Data Structure to accumulate Keyword Times
        keywordTimes = {}
        keywordLengths = {}
        # Because Snapshot records are coded two different ways, we need to be able to keep track of what
        # we've already counted in clipCount and ClipTotalTime so we don't count it twice.
        self.itemsCounted = []

        # The majorList and minorList are constructed differently for the Episode and Document versions of the report,
        # and so the report must be built differently here too!
        if (self.episodeName == None) and (self.documentName == None):
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                majorLabel = unicode(majorLabel, 'utf8')

            # Let's keep track of the current collection name so we can tell when to print
            # the Collection Header and Collection Notes if that's how the report is configured
            # If we have a Collection-based Report ...
            if (self.collection != None):
                # ... put garbage into the workingCollection variable so it will trigger printing for the
                # first Collection
                workingCollection = ('NoCollectionString', 0)
            # If we're not doing a Collection Report ...
            else:
                # Initialize workingCollection to an empty string, as we won't be using it.
                workingCollection = ''

            # If there are 20 or more items in the list, or at least 3 images ...
            if (self.collection != None) and ((len(majorList) >= 20) or (len(self.snapshotFilterList) > 3)):
                # ... create a Progress Dialog.  (The PARENT is needed to prevent the report being hidden
                #     behind Transana!)
                progress = wx.ProgressDialog(self.title, _('Assembling report contents'), parent=self.report)

            # Iterate through the major list
            for (objType, groupNo, group, parentCollNo) in majorList:

                # If our majorLabel is Clip/Snapshot ...
                if majorLabel.encode('utf8') in [_('Document'), _('Episode'), _('Clip'), _('Snapshot'), _('Quote')]:
                    # ... set the majorLabel to match the object type (but translated)
                    if objType == 'Document':
                        majorLabel = _('Document')
                    elif objType == 'Episode':
                        majorLabel = _('Episode')
                    elif objType == 'Clip':
                        majorLabel = _('Clip')
                    elif objType == 'Snapshot':
                        majorLabel = _('Snapshot')
                    elif objType == _('Quote'):
                        majorLabel = _('Quote')
                        
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    majorLabel = unicode(majorLabel, 'utf8')

                # If a Collection Name is passed in ...
                if self.collection != None:
                    # ... then our Filter comparison is based on Clip data
                    filterVal = (group, parentCollNo, True)
                # If a Library Name is passed in ...            
                elif self.seriesName != None:
                    # ... then our Filter comparison is based on Episode data
                    filterVal = (group, self.seriesName, True)
                # If this report is called for a SearchLibraryResult ...
                elif (self.searchSeries != None) and (self.treeCtrl != None):
                    # ... then our Filter comparison is based on the search Library from the TreeCtrl
                    filterVal = (group, self.treeCtrl.GetItemText(self.searchSeries), True)
                # If this report is called for a SearchCollectionResult ...
                elif (self.searchColl != None) and (self.treeCtrl != None):
                    # ... then our Filter comparison is based on Search Collection data
                    filterVal = (group, parentCollNo, True)

                # now that we have the filter comparison data, we see if it's actually in the Filter List.
                if ((objType == 'Document') and (filterVal in self.documentFilterList)) or \
                   ((objType == 'Snapshot') and (filterVal in self.snapshotFilterList)) or \
                   ((objType == 'Quote')    and (filterVal in self.quoteFilterList)) or \
                   (filterVal in self.filterList):
                    # If we have Collection-based data ...
                    if (self.collection != None) or ((self.searchColl != None) and (self.treeCtrl != None)):
                        # ... load the collection the current clip is in
                        tempColl = Collection.Collection(parentCollNo)

                        # Check to see if we're showing Collection headers, if we're showing nested collections (since
                        # there's no point showing collection headers if there aren't different collections!), and
                        # see if the new collection is different from the collection of the last clip displayed.
                        if (workingCollection != '') and \
                           (self.showNested or self.showComments or self.showCollectionNotes) and \
                           (workingCollection[0] != tempColl.GetNodeString()):
                            # Format text for the next section of the report
                            reportText.SetTxtStyle(fontSize=12, fontBold=useBold, fontUnderline=False,
                                                   parAlign = wx.TEXT_ALIGNMENT_LEFT,
                                                   parLeftIndent = 0,
                                                   parSpacingBefore = 24, parSpacingAfter = 0)
                            # Add the Collections header and data to the report
                            reportText.WriteText(_('Collection: '))
#                            reportText.SetTxtStyle(fontBold=False)
                            reportText.WriteText('%s\n' % tempColl.GetNodeString())
                            
                            # If we are supposed to show Comments ...
                            if self.showComments:
                                # ... if the collection has a comment ...
                                if tempColl.comment != u'':
                                    # Set the font for the comments
                                    reportText.SetTxtStyle(fontSize=10, fontBold=useBold, parLeftIndent=63, parRightIndent=63,
                                                           parSpacingBefore = 0, parSpacingAfter = 0)
                                    # Add the header to the report
                                    reportText.WriteText(_('Collection Comment:\n'))
                                    reportText.SetTxtStyle(fontBold=False, parLeftIndent=127,
                                                           parSpacingBefore = 0, parSpacingAfter = 0)
                                    # Add the content of the Collection Comment to the report
                                    reportText.WriteText('%s\n' % tempColl.comment)

                            # If we're supposed to show Collection Notes ...
                            if self.showCollectionNotes:
                                # ... get a list of notes, including their object numbers
                                notesList = DBInterface.list_of_notes(Collection=tempColl.number, includeNumber=True)
                                # If there are notes for this Clip ...
                                if (len(notesList) > 0):
                                    # Set the font for the Notes
                                    reportText.SetTxtStyle(fontSize=10, fontBold=useBold,
                                                           parLeftIndent=63, parRightIndent=63,
                                                           parSpacingBefore = 0, parSpacingAfter=0)
                                    # Add the header to the report
                                    reportText.WriteText(_('Collection Notes:\n'))
#                                    reportText.Newline()
                                    # Iterate throught the list of notes ...
                                    for note in notesList:
                                        # ... load each note ...
                                        tempNote = Note.Note(note[0])
                                        reportText.SetTxtStyle(fontBold=useBold, parLeftIndent=127, parSpacingBefore = 0, parSpacingAfter = 0)
                                        # Add the note ID to the report
                                        reportText.WriteText('%s\n' % tempNote.id)
#                                        reportText.Newline()
                                        # Turn bold off.
                                        reportText.SetTxtStyle(fontBold=False, parLeftIndent=190)
                                        # Add the note text to the report (rstrip() prevents formatting problems when a note ends with blank lines)
                                        reportText.WriteText('%s\n' % tempNote.text.rstrip())
#                                        reportText.Newline()
                            # Update the workingCollection variable with the data for the current collection so we'll
                            # be able to tell when the collection changes
                            workingCollection = (tempColl.GetNodeString(), tempColl.number)
                            
                            # We need to indent EVERYTHING else to adjust for these headers
                            baseIndent = 63
                        else:
                            # We DON'T need to indent later paragraphs
                            baseIndent = 63
                    else:
                        # We DON'T need to indent later paragraphs
                        baseIndent = 0
                            
                    # ... Set the formatting for the report, including turning off previous formatting
                    reportText.SetTxtStyle(fontSize = 12, fontBold = True, fontUnderline = False,
                                           parAlign = wx.TEXT_ALIGNMENT_LEFT,
                                           parLeftIndent = baseIndent, parRightIndent = 0,
                                           parSpacingBefore = 36, parSpacingAfter = 2)

                    if DEBUG:
                        print "%s, %s (majorLabel, group) l=%d (baseindent), r=0, b=36, a=0" % (majorLabel, group, baseIndent)

                    # Add the group name to the report
                    reportText.WriteText('%s: ' % majorLabel)

                    # If we're showing Hyperlinks to Clips/Snapshots ...
                    if self.showHyperlink:
                        # Define a Hyperlink Style (Blue, underlined)
                        urlStyle = richtext.RichTextAttr()
                        urlStyle.SetFontFaceName('Courier New')
                        urlStyle.SetFontSize(12)
                        urlStyle.SetTextColour(wx.BLUE)
                        urlStyle.SetFontUnderlined(True)
                        # Apply the Hyperlink Style
                        reportText.BeginStyle(urlStyle)
                        # Insert the Hyperlink information, object type and object number
                        reportText.BeginURL("transana:%s=%d" % (objType, groupNo))

                    # Add the group name to the report
                    reportText.WriteText('%s\n' % group)
                    # End the paragraph
#                    reportText.Newline()

                    # If we're showing Hyperlinks to Clips/Snapshots ...
                    if self.showHyperlink:
                        # End the Hyperlink
                        reportText.EndURL()
                        # Stop using the Hyperlink Style
                        reportText.EndStyle()

                    # ... Set the formatting for the report, including turning off previous formatting
                    reportText.SetTxtStyle(fontSize = 10, fontBold = useBold,
                                           parLeftIndent = baseIndent + 63, parRightIndent = 63,
                                           parSpacingBefore = 0, parSpacingAfter = 0)

                    # If we have Collection-based data, we add some Clip-specific data
                    if (self.collection != None) or ((self.searchColl != None) and (self.treeCtrl != None)):
                        # Add the header to the report
                        reportText.WriteText(_('Collection:'))
                        reportText.SetTxtStyle(fontBold = False)
                        # Add the data to the report, the full Collection path in this case
                        reportText.WriteText('  %s\n' % (tempColl.GetNodeString(),))
                        # If we're looking at a Quote ...
                        if objType == 'Quote':
                            # Get the full Quote data
                            quoteObj = Quote.Quote(groupNo)
                            tmpObj = quoteObj
                            try:
                                # If we have a Quote, load the Source Document!
                                tmpDoc = Document.Document(num=tmpObj.source_document_num)
                            except TransanaExceptions.RecordNotFoundError:
                                tmpDoc = None
                        # If we're looking at a Clip ...
                        elif objType == 'Clip':
                            # Get the full Clip data
                            clipObj = Clip.Clip(groupNo)
                            tmpObj = clipObj
                        # If we're looking at a Snapshot ...
                        elif objType == 'Snapshot':
                            # Get the full Snapshot data
                            snapshotObj = Snapshot.Snapshot(groupNo, suppressEpisodeError = True)
                            tmpObj = snapshotObj
                        # If we're supposed to show the Media File Name ...
                        if self.showFile:
                            # Turn bold on.
                            reportText.SetTxtStyle(fontBold = useBold)
                            if objType == 'Quote':
                                # Add the header to the report
                                reportText.WriteText(_('Source File:'))
                            else:
                                # Add the header to the report
                                reportText.WriteText(_('File:'))
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False)
                            # If we have a Quote ...
                            if objType == 'Quote':
                                if tmpDoc != None:
                                    # Add the data to the report, the file name in this case
                                    reportText.WriteText(_('  %s\n') % tmpDoc.imported_file)
                                    prompt = unicode(_("imported on %s\n"), "utf8")
                                    # sqlite gives a string rather than a datetime object
                                    if isinstance(tmpDoc.import_date, str):
                                        # start exception handling in case of formatting problems
                                        try:
                                            # Convert the date string to a datetime object
                                            tmpDate = datetime.datetime.strptime(tmpDoc.import_date, '%Y-%m-%d %H:%M:%S')
                                            # Display the correct date
                                            reportText.WriteText(prompt % tmpDate.strftime('%x'))
                                        # If the conversion fails ...
                                        except ValueError:
                                            # ... display the un-converted string.
                                            reportText.WriteText(prompt % tmpDoc.import_date)
                                    # MySQL returns a datetime object.
                                    else:
                                        reportText.WriteText(prompt % tmpDoc.import_date.strftime('%x'))
#                                else:
#                                    reportText.WriteText(_("The source Document for this Quote has been deleted.") + u'\n')
                            # If we have a Clip ...
                            elif objType == 'Clip':
                                # Add the data to the report, the file name in this case
                                reportText.WriteText(_('  %s\n') % tmpObj.media_filename)
                                # Add Additional Media File info
                                for mediaFile in tmpObj.additional_media_files:
                                    reportText.WriteText(_('       %s\n') % mediaFile['filename'])
                            # If we have a Snapshot ...
                            elif objType == 'Snapshot':
                                # Add the data to the report, the file name in this case
                                reportText.WriteText(_('  %s\n') % tmpObj.image_filename)

                        # If we're supposed to show the Clip/Snapshot Time data ...
                        if self.showTime:
                            # We DON'T show this if we have a Snapshot with no defined duration
                            if (objType == 'Clip') or (objType == 'Quote') or \
                               ((objType == 'Snapshot') and (tmpObj.episode_num > 0) and (tmpObj.episode_duration > 0)):
                                # Turn bold on.
                                reportText.SetTxtStyle(fontBold = useBold, parSpacingAfter = 0)
                                if objType == 'Quote':
                                    # Add the header to the report
                                    reportText.WriteText(_('Position:'))
                                else:
                                    # Add the header to the report
                                    reportText.WriteText(_('Time:'))
                                # Turn bold off.
                                reportText.SetTxtStyle(fontBold = False)
                                if objType == 'Quote':
                                    # Add the data to the report, the Quote start and end characters in this case
                                    reportText.WriteText('  %s - %s   (' % (quoteObj.start_char, quoteObj.end_char))
                                # If we have a Clip ...
                                elif objType == 'Clip':
                                    # Add the data to the report, the Clip start and stop times in this case
                                    reportText.WriteText('  %s - %s   (' % (Misc.time_in_ms_to_str(clipObj.clip_start), Misc.time_in_ms_to_str(clipObj.clip_stop)))
                                # If we have a Snapshot ...
                                elif objType == 'Snapshot':
                                    # Add the data to the report, the Snapshot start and stop times in this case
                                    reportText.WriteText('  %s - %s   (' % (Misc.time_in_ms_to_str(tmpObj.episode_start), Misc.time_in_ms_to_str(tmpObj.episode_start + tmpObj.episode_duration)))
                                # Turn bold on.
                                reportText.SetTxtStyle(fontBold = useBold)
                                # Add the header to the report
                                reportText.WriteText(_('Length:'))
                                # Turn bold off.
                                reportText.SetTxtStyle(fontBold = False)
                                # If we have a Quote ...
                                if objType == 'Quote':
                                    # Add the data to the report, the Quote length in this case
                                    reportText.WriteText('  %s)\n' % (quoteObj.end_char - quoteObj.start_char))
                                # If we have a Clip ...
                                elif objType == 'Clip':
                                    # Add the data to the report, the Clip Length in this case
                                    reportText.WriteText('  %s)\n' % (Misc.time_in_ms_to_str(clipObj.clip_stop - clipObj.clip_start)))
                                # If we have a Snapshot ...
                                elif objType == 'Snapshot':
                                    # Add the data to the report, the Snapshot Duration in this case
                                    reportText.WriteText('  %s)\n' % (Misc.time_in_ms_to_str(tmpObj.episode_duration)))

                        # If we have a Quote ...
                        if objType == 'Quote':
                            # Increment the Item Counter
                            self.quoteCount += 1
                            # Add the Quote's length to the Quote Total Length accumulator
                            self.quoteTotalLength += quoteObj.end_char - quoteObj.start_char
                        # If we have a Clip ...
                        if objType == 'Clip':
                            # Increment the Item Counter
                            self.clipCount += 1
                            # Add the Clip's length to the Clip Total Time accumulator
                            self.clipTotalTime += clipObj.clip_stop - clipObj.clip_start
                        # If we have a Snapshot ...
                        elif (objType == 'Snapshot'):
                            # Increment the Item Counter
                            self.snapshotCount += 1
                            if (tmpObj.episode_num > 0):
                                # Add the Snapshot's length to the Total Time accumulator
                                self.snapshotTotalTime += snapshotObj.episode_duration

                        # If we're supposed to show Source Information, and the item HAS source information ...
                        if self.showSourceInfo:
                            reportText.SetTxtStyle(fontSize = 10, parSpacingAfter = 0)
                            if (objType == 'Quote'):
                                if(tmpDoc != None):
                                    # Turn bold on.
                                    reportText.SetTxtStyle(fontBold = useBold)
                                    # Add the header to the report
                                    reportText.WriteText(_('Library:'))
                                    reportText.SetTxtStyle(fontBold = False)
                                    # Add the data to the report
                                    reportText.WriteText('  %s\n' % tmpDoc.library_id)
                            else:
                                if tmpObj.series_id != '':
                                    # Turn bold on.
                                    reportText.SetTxtStyle(fontBold = useBold)
                                    # Add the header to the report
                                    reportText.WriteText(_('Library:'))
                                    reportText.SetTxtStyle(fontBold = False)
                                    # Add the data to the report
                                    reportText.WriteText('  %s\n' % tmpObj.series_id)
                                if tmpObj.episode_num > 0:
                                    # Turn bold on.
                                    reportText.SetTxtStyle(fontBold = useBold)
                                    # Add the header to the report
                                    reportText.WriteText(_('Episode:'))
                                    reportText.SetTxtStyle(fontBold = False)
                                    # Add the data to the report
                                    reportText.WriteText('  %s\n' % tmpObj.episode_id)

                        # If we're supposed to show Source Information or Quote Text or Clip Transcripts ...
                        if self.showSourceInfo or self.showQuoteText or self.showTranscripts:
                            reportText.SetTxtStyle(fontSize = 10, parSpacingAfter = 0)

                            if (objType == 'Quote'):
                                # Turn bold on.
                                reportText.SetTxtStyle(fontBold =useBold)
                                # Add the header to the report
                                reportText.WriteText(_('Document:'))
                                # Turn bold off.
                                reportText.SetTxtStyle(fontBold = False)
                                # If a Source Document was found ...
                                if tmpDoc != None:
                                    # Add the data to the report, the Source Document ID in this case
                                    reportText.WriteText('  %s\n' % (tmpDoc.id,))
                                # if no Source Document is found, we have an orphan.
                                else:
                                    # Add the data to the report, the Source Document ID in this case
                                    reportText.WriteText('  %s\n' % _('The original Document has been deleted.'))

                                if self.showQuoteText:
                                    # Turn bold on.
                                    reportText.SetTxtStyle(fontSize = 10, fontFace = 'Courier New', fontBold = useBold,
                                                           parLeftIndent = baseIndent + 63, parRightIndent = 0,
                                                           parSpacingBefore = 0, parSpacingAfter = 0)
                                    # Add the header to the report
                                    reportText.WriteText(_('Quote Text:\n'))
                                
                                    # Turn bold off.
                                    reportText.SetTxtStyle(fontBold = False)
                                    # Add the Quote Text to the report.  Quote text *must* be in XML format.

                                    # Strip the time codes for the report
                                    text = reportText.StripTimeCodes(tmpObj.text)

                                    # Create the Transana XML to RTC Import Parser.  This is needed so that we can
                                    # pull XML transcripts into the existing RTC without resetting the contents of
                                    # the reportText RTC, which wipes out all accumulated Report data.
                                    # Pass the reportText RTC and the desired additional margins in.
                                    handler = PyXML_RTCImportParser.XMLToRTCHandler(reportText, (127 + baseIndent, 127))
                                    # Parse the transcript text, adding it to the reportText RTC
                                    xml.sax.parseString(text, handler)

                            elif (objType == 'Clip'):
                                # Iterate through the clips transcripts
                                for tr in clipObj.transcripts:
                                    if self.showSourceInfo:
                                        # Default the Episode Transcript to None in case the load fails
                                        episodeTranscriptObj = None
                                        # Begin exception handling
                                        try:
                                            # If the Clip Object has a defined Source Transcript ...
                                            if tr.source_transcript > 0:
                                                # ... try to load that source transcript
                                                # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                                                episodeTranscriptObj = Transcript.Transcript(tr.source_transcript, skipText=True)
                                        # if the record is not found (orphaned Clip)
                                        except TransanaExceptions.RecordNotFoundError:
                                            # We don't need to do anything.
                                            pass

                                        # If an Episode Transcript was found ...
                                        if episodeTranscriptObj != None:
                                            # Turn bold on.
                                            reportText.SetTxtStyle(fontSize = 10, fontFace = 'Courier New', fontBold = useBold,
                                                                   parLeftIndent = baseIndent + 63, parRightIndent = 0,
                                                                   parSpacingBefore = 0, parSpacingAfter = 0)
                                            # Add the header to the report
                                            reportText.WriteText(_('Episode Transcript:'))
                                            # Turn bold off.
                                            reportText.SetTxtStyle(fontBold = False)
                                            # Add the data to the report, the Episode Transcript ID in this case
                                            reportText.WriteText('  %s\n' % (episodeTranscriptObj.id,))
    #                                        reportText.Newline()
                                        # if no Episode Transcript is found, we have an orphan.
                                        else:
                                            # Turn bold on.
                                            reportText.SetTxtStyle(fontBold =useBold)
                                            # Add the header to the report
                                            reportText.WriteText(_('Episode Transcript:'))
                                            # Turn bold off.
                                            reportText.SetTxtStyle(fontBold = False)
                                            # Add the data to the report, the Episode Transcript ID in this case
                                            reportText.WriteText('  %s\n' % _('The Episode Transcript has been deleted.'))
    #                                        reportText.Newline()

                                    if self.showTranscripts:
                                        # Turn bold on.
                                        reportText.SetTxtStyle(fontSize = 10, fontFace = 'Courier New', fontBold = useBold,
                                                               parLeftIndent = baseIndent + 63, parRightIndent = 0,
                                                               parSpacingBefore = 0, parSpacingAfter = 0)
                                        # Add the header to the report
                                        reportText.WriteText(_('Clip Transcript:\n'))
    #                                    reportText.Newline()

                                        # Turn bold off.
                                        reportText.SetTxtStyle(fontBold = False)
                                        # Add the Transcript to the report
                                        # Clip Transcripts could be in the old RTF format, or they could have been
                                        # updated to the new XML format.  These require different processing.

                                        # If we have a Rich Text Format document ...
                                        if tr.text[:5].lower() == u'{\\rtf':
                                            # Create a temporary RTC control
                                            tmpTxtCtrl = TranscriptEditor_RTC.TranscriptEditor(reportText.parent, pos=(-20, -20), suppressGDIWarning = True)
                                            # ... import the RTF data into the report text
                                            tmpTxtCtrl.LoadRTFData(tr.text, clearDoc=True)
                                            # Pull the data back out of the control as XML
                                            tmpText = tmpTxtCtrl.GetFormattedSelection('XML')
                                            # Strip the time codes for the report
                                            tmpText = reportText.StripTimeCodes(tmpText)

                                            # Create the Transana XML to RTC Import Parser.  This is needed so that we can
                                            # pull XML transcripts into the existing RTC without resetting the contents of
                                            # the reportText RTC, which wipes out all accumulated Report data.
                                            # Pass the reportText RTC and the desired additional margins in.
                                            handler = PyXML_RTCImportParser.XMLToRTCHandler(reportText, (127 + baseIndent, 127))
                                            # Parse the transcript text, adding it to the reportText RTC
                                            xml.sax.parseString(tmpText, handler)

                                            del(handler)

                                            # Destroy the temporary RTC control
                                            tmpTxtCtrl.Destroy()

                                        # If we have an XML document ...
                                        elif tr.text[:5].lower() == u'<?xml':
                                            # Strip the time codes for the report
                                            tr.text = reportText.StripTimeCodes(tr.text)

                                            # Create the Transana XML to RTC Import Parser.  This is needed so that we can
                                            # pull XML transcripts into the existing RTC without resetting the contents of
                                            # the reportText RTC, which wipes out all accumulated Report data.
                                            # Pass the reportText RTC and the desired additional margins in.
                                            handler = PyXML_RTCImportParser.XMLToRTCHandler(reportText, (127 + baseIndent, 127))
                                            # Parse the transcript text, adding it to the reportText RTC
                                            xml.sax.parseString(tr.text, handler)

                                        # If we have a transcript that is neither RTF nor XML (shouldn't happen!)
                                        else:
                                            # ... then just import it directly.  Treat it as plain text.
                                            # (rstrip() prevents formatting problems when a transcript ends with blank lines)
                                            reportText.WriteText(tr.text.rstrip())

                            # If we have a Snapshot ...
                            elif objType == 'Snapshot':
                                if self.showSourceInfo:
                                    # ... if the Snapshot has a defined Transcript ...
                                    if (tmpObj.transcript_num > 0) and (tmpObj.episode_num > 0):
                                        # Turn bold on.
                                        reportText.SetTxtStyle(fontSize = 10, fontFace = 'Courier New', fontBold = useBold,
                                                               parLeftIndent = baseIndent + 63, parRightIndent = 0,
                                                               parSpacingBefore = 0, parSpacingAfter = 0)
                                        # Add the header to the report
                                        reportText.WriteText(_('Episode Transcript:'))
                                        reportText.SetTxtStyle(fontBold = False)
                                        # Add the data to the report
                                        reportText.WriteText('  %s\n' % tmpObj.transcript_id)
                            
                        # If we have a Snapshot, and we're displaying Full, Medium, or Small images, show the actual IMAGE
                        if (objType == 'Snapshot') and (self.showSnapshotImage in [0, 1, 2]):
                            # Start Exception Handling
                            try:
                                # Open a HIDDEN Snapshot Window
                                tmpSnapshotWindow = SnapshotWindow.SnapshotWindow(TransanaGlobal.menuWindow, -1, tmpObj.id, tmpObj, showWindow=False)
                                # Get the cropped, coded image from the Snapshot Window
                                tmpBMP = tmpSnapshotWindow.CopyBitmap()
                                # Close the hidden Snapshot Window
                                tmpSnapshotWindow.Close()
                                # Explicitly Delete the Temporary Snapshot Window
                                tmpSnapshotWindow.Destroy()
                                
                                # Convert the Bitmap to an Image so it can be rescaled
                                tmpImage = tmpBMP.ConvertToImage()
                                # Get the Image Size
                                (imgWidth, imgHeight) = tmpImage.GetSize()
                                # We need the SMALLER of the current image size and the current Transcript Window size
                                # (Adjust width for scrollbar size!)
                                maxWidth = min(float(imgWidth), (reportText.GetSize()[0] - 20.0) * 0.72)
                                maxHeight = min(float(imgHeight), reportText.GetSize()[1] * 0.80)
                                # if we're using "Medium" size ...
                                if self.showSnapshotImage == 1:
                                    # ... set the image max size to 500 pixels
                                    maxWidth = min(maxWidth, 500)
                                    maxHeight = min(maxHeight, 500)
                                # If we're using "Small" size ...
                                elif self.showSnapshotImage == 2:
                                    # ... set the image max size to 250 pixels
                                    maxWidth = min(maxWidth, 250)
                                    maxHeight = min(maxHeight, 250)
                                # Determine the scaling factor for adjusting the image size
                                scaleFactor = min(maxWidth / float(imgWidth), maxHeight / float(imgHeight))
                                # If the image is too BIG, it needs to be re-scaled ...
                                if scaleFactor < 1.0:
                                    # ... so rescale the image to fit in the current Transcript window.  Use slower high quality rescale.
                                    tmpImage.Rescale(int(imgWidth * scaleFactor), int(imgHeight * scaleFactor), quality=wx.IMAGE_QUALITY_HIGH)
                                # If we have anything but a large image ...
                                if self.showSnapshotImage > 0:
                                    # ... alter the Paragraph Spacing here.
                                    reportText.SetTxtStyle(parSpacingBefore = 24, parSpacingAfter = 24)
                                # If we have a large image ...
                                else:
                                    # ... alter the Paragraph Spacing here, and DON'T indent the image itself, so it can be as large as possible!
                                    reportText.SetTxtStyle(parLeftIndent = 0, parSpacingBefore = 24, parSpacingAfter = 24)
                                # Add the image to the transcript
                                reportText.WriteImage(tmpImage)
                                # Delete the temporary image, bitmap, and device contexts
                                tmpImage.Destroy()
                                tmpBMP.Destroy()
                                # Add some more blank space.
                                reportText.WriteText('\n')
                                # Reset the Paragraph Spacing here
                                reportText.SetTxtStyle( parLeftIndent = baseIndent + 63, parSpacingBefore = 0, parSpacingAfter = 0)

                            # Detect Image Loading problems
                            except TransanaExceptions.ImageLoadError, e:

                                if DEBUG:
                                    print "ReportGenerator.OnDisplay():"
                                    print tmpObj.GetNodeString()
                                    print
                                    print sys.exc_info()[0]
                                    print sys.exc_info()[1]
                                    import traceback
                                    traceback.print_exc(file=sys.stdout)
                                    print

                                # If we're displaying Image Load Errors ...
                                if not skippingImageError:
                                    # ... build the error message and display it.
                                    tmpDlg = Dialogs.ErrorDialog(self.report, e.explanation, includeSkipCheck=True)
                                    tmpDlg.ShowModal()
                                    # See if the user is requesting that further messages be skipped
                                    skippingImageError = tmpDlg.GetSkipCheck()
                                    tmpDlg.Destroy()
                                
                                # Alter the Paragraph Spacing here
                                reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 0)
                                # Add some more blank space.
                                reportText.WriteText('\n')
                                # Add the image to the transcript
                                reportText.WriteText(e.explanation)
                                # Alter the Paragraph Spacing here
                                reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 24)
                                # Add some more blank space.
                                reportText.WriteText('\n')
                                # Reset the Paragraph Spacing here
                                reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 0)

                            # Detect PyAssertionError
                            except wx._core.PyAssertionError, e:

                                if DEBUG:
                                    print "ReportGenerator.OnDisplay():"
                                    print tmpObj.GetNodeString()
                                    print
                                    print sys.exc_info()[0]
                                    print sys.exc_info()[1]
                                    import traceback
                                    traceback.print_exc(file=sys.stdout)
                                    print

                                # ... build the error message and display it.
                                tmpDlg = Dialogs.ErrorDialog(self.report, e.message)
                                tmpDlg.ShowModal()
                                tmpDlg.Destroy()
                                
                                # Alter the Paragraph Spacing here
                                reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 0)
                                # Add some more blank space.
                                reportText.WriteText('\n')
                                # Add the image to the transcript
                                reportText.WriteText(e.message)
                                # Alter the Paragraph Spacing here
                                reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 24)
                                # Add some more blank space.
                                reportText.WriteText('\n')
                                # Reset the Paragraph Spacing here
                                reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 0)
                                

                        # We want to display the Coding Key for the Snapshot.  But we only want to display the information
                        # for those codes that are VISIBLE in the image (even if they're not included on the Snapshot!)

                        # If we're supposed to show Snapshot Coding ...
                        if (objType == 'Snapshot') and self.showSnapshotCoding:
                            # Create an empty list to hold the keys of visible codes
                            tmpKeys = []
                            # For each Coding Object in the Snapshot ...
                            for key in tmpObj.codingObjects.keys():
                                # ... if the Coding Object is visible and the Keyword it represents isn't already in the tmpKeys list ...
                                if tmpObj.codingObjects[key]['visible'] and \
                                   (not (tmpObj.codingObjects[key]['keywordGroup'], tmpObj.codingObjects[key]['keyword']) in tmpKeys):
                                    # ... add the Keyword to the tmpKeys List
                                    tmpKeys.append((tmpObj.codingObjects[key]['keywordGroup'], tmpObj.codingObjects[key]['keyword']))

                            # If there are keywords in the list
                            if len(tmpKeys) > 0:
                                # Turn bold on.
                                reportText.SetTxtStyle(fontBold = useBold, parLeftIndent=baseIndent + 63, parRightIndent = 0,
                                                       parSpacingAfter = 0)
                                # Add the header to the report
                                reportText.WriteText(_('Snapshot Coding Key:'))
                                # Turn bold off.
                                reportText.SetTxtStyle(fontBold = False)
                                # Finish the Line
                                reportText.WriteText('\n')
#                                reportText.Newline()
                                # Set formatting for Keyword lines
                                reportText.SetTxtStyle(parLeftIndent = baseIndent + 127, parRightIndent = 0,
                                                       parSpacingBefore = 0, parSpacingAfter = 0)
                                # Sort the Keywords
                                tmpKeys.sort()
                                # For each keyword ...
                                for key in tmpKeys:

                                    # See if the keyword should be included, based on the Keyword Filter List
                                    if (key[0], key[1], True) in self.keywordFilterList:

                                        # Add the keyword
                                        reportText.WriteText('%s : %s  (' % (key[0], key[1]))

                                        # if THIS Snapshot with THIS Keyword has NOT already been counted ...
                                        if not (('Snapshot', tmpObj.number, key[0], key[1]) in self.itemsCounted):
                                            # Add this Keyword to the Keyword Counts
                                            if keywordCounts.has_key('%s : %s' % (key[0], key[1])):
                                                keywordCounts['%s : %s' % (key[0], key[1])] += 1
                                                if tmpObj.episode_num > 0:
                                                    keywordTimes['%s : %s' % (key[0], key[1])] += tmpObj.episode_duration
                                            else:
                                                keywordCounts['%s : %s' % (key[0], key[1])] = 1
                                                keywordTimes['%s : %s' % (key[0], key[1])] = 0
                                                keywordLengths['%s : %s' % (key[0], key[1])] = 0
                                                if tmpObj.episode_num > 0:
                                                    keywordTimes['%s : %s' % (key[0], key[1])] = tmpObj.episode_duration
                                            # Remember that THIS Snapshot with THIS Keyword HAS been counted now
                                            self.itemsCounted.append(('Snapshot', tmpObj.number, key[0], key[1]))
                                        # Get the Graphic for the Coding Key
                                        tmpImage = SnapshotWindow.CodingKeyGraphic(tmpObj.keywordStyles[key])
                                        # Add the Image to the Report
                                        reportText.WriteImage(tmpImage)
                                        # Delete the temporary image
                                        tmpImage.Destroy()
                                        # Add the Shape Description and a Line Feed
                                        reportText.WriteText(')\n')

                        # If there are 20 or more items in the list, or at least 3 images ...
                        if (self.collection != None) and ((len(majorList) >= 20) or (len(self.snapshotFilterList) > 3)):
                            # ... update the progress bar
                            progress.Update(int(float(self.quoteCount + self.clipCount + self.snapshotCount) / float(len(majorList)) * 100))

                    # If we have a Library Report ...
                    else:
                        if objType == 'Episode':
                            # Get the full Episode data
                            tmpObj = Episode.Episode(groupNo)
                            fileName = tmpObj.media_filename
                            addFiles = tmpObj.additional_media_files
                        elif objType == 'Document':
                            tmpObj = Document.Document(groupNo)
                            fileName = tmpObj.imported_file
                            addFiles = []
                        # If we're supposed to show the Media File Name ...
                        if self.showFile:
                            # Turn bold on.
                            reportText.SetTxtStyle(fontBold = True, parSpacingAfter = 0)
                            # Add the header to the report
                            reportText.WriteText(_('File:'))
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False)
                            # Add the data to the report, the file name in this case
                            reportText.WriteText(_('  %s\n') % fileName)
#                            reportText.Newline()
                            # Add Additional Media File info
                            for mediaFile in addFiles:
                                reportText.WriteText(_('       %s\n') % mediaFile['filename'])
#                                reportText.Newline()

                        # If we're supposed to show the Episode Time data ...
                        if self.showTime and (objType == 'Episode') and (tmpObj.tape_length > 0):
                            # Turn bold on.
                            reportText.SetTxtStyle(fontBold = True, parSpacingAfter = 0)
                            # Add the header to the report
                            reportText.WriteText(_('Length:'))
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False)
                            # Add the data to the report, the Episode Length in this case
                            reportText.WriteText('  %s\n' % (Misc.time_in_ms_to_str(tmpObj.tape_length)))
#                            reportText.Newline()
#                            reportText.SetTxtStyle(parSpacingAfter = 0)

                        # If we're supposed to show the Document File Import Date data ...
                        if self.showDocImportDate and (objType == 'Document') and (tmpObj.import_date != None):
                            # Turn bold on.
                            reportText.SetTxtStyle(fontBold = True, parSpacingAfter = 0)
                            # Add the header to the report
                            reportText.WriteText(_('Import Date:'))
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False)
                            # Add the data to the report, the Document Import Date in this case
                            # sqlite gives a string rather than a datetime object
                            if isinstance(tmpObj.import_date, str):
                                # start exception handling in case of formatting problems
                                try:
                                    # Convert the date string to a datetime object
                                    tmpDate = datetime.datetime.strptime(tmpObj.import_date, '%Y-%m-%d %H:%M:%S')
                                    # Display the correct date
                                    reportText.WriteText('  %s\n' % tmpDate.strftime('%x'))
                                # If the conversion fails ...
                                except ValueError:
                                    # ... display the un-converted string.
                                    reportText.WriteText('  %s\n' % tmpObj.import_date)
                            # MySQL returns a datetime object.
                            else:
                                reportText.WriteText('  %s\n' % tmpObj.import_date.strftime('%x'))
#                            reportText.Newline()
#                            reportText.SetTxtStyle(parSpacingAfter = 0)

                        if objType == 'Document':
                            # Increment the Document Counter (Yeah, we're using quoteCount)
                            self.quoteCount += 1
                            self.quoteTotalLength += tmpObj.document_length
                        elif objType == 'Episode':
                            # Increment the Episode Counter (Yeah, we're using clipCount)
                            self.clipCount += 1
                            # Add the Episode's length to the Episode Total Time accumulator (Yeah, we're using clipTotalTime)
                            self.clipTotalTime += tmpObj.tape_length

                    # Reset the font.  It could have been contaminated by the Clip Transcript
                    reportText.SetTxtStyle(fontFace='Courier New', fontSize = 10, fontColor = wx.Colour(0, 0, 0),
                                           fontBgColor = wx.Colour(255, 255, 255), fontBold = False, fontItalic = False,
                                           fontUnderline = False)
                    reportText.SetTxtStyle(parAlign = wx.TEXT_ALIGNMENT_LEFT,
                                           parLineSpacing = wx.TEXT_ATTR_LINE_SPACING_NORMAL,
                                           parLeftIndent = 0,
                                           parRightIndent = 0,
                                           parSpacingBefore = 0,
                                           parSpacingAfter = 0,
                                           parTabs = [])

                    if DEBUG:
                        print "RESET all FONT and PARAGRAPH settings"

                    # if we are supposed to show Keywords ...
                    if self.showKeywords:

                        # If there are keywords in the list ... (even if they might all get filtered out ...)
                        if len(minorList[(objType, groupNo)]) > 0:
                            # Turn bold on.
                            reportText.SetTxtStyle(fontBold = useBold, parLeftIndent=baseIndent + 63, parRightIndent = 0,
                                                   parSpacingAfter = 0)
                            # Add the header to the report
                            if self.seriesName != None:
                                if objType == 'Episode':
                                    reportText.WriteText(_('Episode Keywords:'))
                                elif objType == 'Document':
                                    reportText.WriteText(_('Document Keywords:'))
                            else:
                                if majorLabel.encode('utf8') == _('Snapshot'):
                                    reportText.WriteText(_('Whole') + ' ')
                                reportText.WriteText(majorLabel + ' ')
                                reportText.WriteText(_('Keywords:'))
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False)
                            # Finish the Line
                            reportText.WriteText('\n')
#                                reportText.Newline()
                            # Set formatting for Keyword lines
                            reportText.SetTxtStyle(parLeftIndent = baseIndent + 127, parRightIndent = 0, parSpacingAfter = 0)
                        # Iterate through the list of Keywords for the group
                        for (keywordGroup, keyword, example) in minorList[(objType, groupNo)]:
                            # See if the keyword should be included, based on the Keyword Filter List
                            if (keywordGroup, keyword, True) in self.keywordFilterList:
                                # Add the Keyword to the report
                                reportText.WriteText('%s : %s\n' % (keywordGroup, keyword))
#                                reportText.Newline()
                                if objType in ['Document', 'Episode']:
                                    # if THIS Object with THIS Keyword has NOT already been counted ...
                                    if not ((objType, tmpObj.number, keywordGroup, keyword) in self.itemsCounted):
                                        # Add this Keyword to the Keyword Counts
                                        if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] += 1
                                        else:
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                                            keywordTimes['%s : %s' % (keywordGroup, keyword)] = 0
                                            keywordLengths['%s : %s' % (keywordGroup, keyword)] = 0
                                        # Remember that THIS Episode with THIS Keyword HAS been counted now
                                        self.itemsCounted.append((objType, tmpObj.number, keywordGroup, keyword))
                                elif objType == 'Quote':
                                    # if THIS Quote with THIS Keyword has NOT already been counted ...
                                    if not (('Quote', tmpObj.number, keywordGroup, keyword) in self.itemsCounted):
                                        # Add this Keyword to the Keyword Counts
                                        if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] += 1
                                            if (self.episodeName != None) or (self.collection != None) or (self.searchColl != None):
                                                if keywordLengths.has_key('%s : %s' % (keywordGroup, keyword)):
                                                    keywordLengths['%s : %s' % (keywordGroup, keyword)] += tmpObj.end_char - tmpObj.start_char
                                                else:
                                                    keywordLengths['%s : %s' % (keywordGroup, keyword)] = tmpObj.end_char - tmpObj.start_char
                                        else:
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                                            keywordTimes['%s : %s' % (keywordGroup, keyword)] = 0
                                            keywordLengths['%s : %s' % (keywordGroup, keyword)] = 0
                                            if (self.episodeName != None) or (self.collection != None) or (self.searchColl != None):
                                                keywordLengths['%s : %s' % (keywordGroup, keyword)] += tmpObj.end_char - tmpObj.start_char
                                        # Remember that THIS Clip with THIS Keyword HAS been counted now
                                        self.itemsCounted.append(('Quote', tmpObj.number, keywordGroup, keyword))
                                elif objType == 'Clip':
                                    # if THIS Clip with THIS Keyword has NOT already been counted ...
                                    if not (('Clip', clipObj.number, keywordGroup, keyword) in self.itemsCounted):
                                        # Add this Keyword to the Keyword Counts
                                        if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] += 1
                                            if (self.episodeName != None) or (self.collection != None) or (self.searchColl != None):
                                                if keywordTimes.has_key('%s : %s' % (keywordGroup, keyword)):
                                                    keywordTimes['%s : %s' % (keywordGroup, keyword)] += clipObj.clip_stop - clipObj.clip_start
                                                else:
                                                    keywordTimes['%s : %s' % (keywordGroup, keyword)] = clipObj.clip_stop - clipObj.clip_start
                                        else:
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                                            keywordTimes['%s : %s' % (keywordGroup, keyword)] = 0
                                            keywordLengths['%s : %s' % (keywordGroup, keyword)] = 0
                                            if (self.episodeName != None) or (self.collection != None) or (self.searchColl != None):
                                                keywordTimes['%s : %s' % (keywordGroup, keyword)] += clipObj.clip_stop - clipObj.clip_start
                                        # Remember that THIS Clip with THIS Keyword HAS been counted now
                                        self.itemsCounted.append(('Clip', clipObj.number, keywordGroup, keyword))
                                # If we have a Snapshot
                                elif objType == 'Snapshot':
                                    # if THIS Snapshot with THIS Keyword has NOT already been counted ...
                                    if not (('Snapshot', tmpObj.number, keywordGroup, keyword) in self.itemsCounted):
                                        # Add this Keyword to the Keyword Counts
                                        if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] += 1
                                            if (tmpObj.episode_num > 0) and \
                                               ((self.episodeName != None) or (self.collection != None) or (self.searchColl != None)):
                                                keywordTimes['%s : %s' % (keywordGroup, keyword)] += tmpObj.episode_duration
                                        else:
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                                            keywordTimes['%s : %s' % (keywordGroup, keyword)] = 0
                                            keywordLengths['%s : %s' % (keywordGroup, keyword)] = 0
                                            if (tmpObj.episode_num > 0) and \
                                               ((self.episodeName != None) or (self.collection != None) or (self.searchColl != None)):
                                                keywordTimes['%s : %s' % (keywordGroup, keyword)] += tmpObj.episode_duration
                                        # Remember that THIS Snapshot with THIS Keyword HAS been counted now
                                        self.itemsCounted.append(('Snapshot', tmpObj.number, keywordGroup, keyword))
                                else:
                                    print "Line 1787", objType
                                    
                    # if we are supposed to show Comments ...
                    if self.showComments:
                        if objType == 'Quote':
                            tmpObj = tmpObj
                        elif objType == 'Clip':
                            tmpObj = clipObj
                        elif objType == 'Snapshot':
                            tmpObj = snapshotObj
                        # If there are keywords in the list ... (even if they might all get filtered out ...)
                        if tmpObj.comment != u'':
                            # Set the font for the Comment
                            reportText.SetTxtStyle(fontSize = 10, fontBold = useBold,
                                                   parLeftIndent = baseIndent + 63, parRightIndent = 0,
                                                   parSpacingBefore = 0, parSpacingAfter = 0)
                            # Add the header to the report
                            reportText.WriteText(majorLabel + ' ')
                            reportText.WriteText(_('Comment:'))
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False)
                            reportText.WriteText('\n')
#                            reportText.Newline()
                            # Format the Clip Comment
                            reportText.SetTxtStyle(parLeftIndent = baseIndent + 127, parRightIndent = 127,
                                                   parSpacingBefore = 0, parSpacingAfter = 0)
                            # Add the content of the Clip Comment to the report
                            reportText.WriteText('%s\n' % tmpObj.comment)
#                            reportText.Newline()

                    # If we are supposed to show Quote, Clip or Snapshot Notes ...
                    if self.showQuoteNotes or self.showClipNotes or self.showSnapshotNotes:
                        # initialize the Notes List to empty so that notes of a type that should NOT be displayed
                        # will be properly skipped
                        notesList = []
                        # ... get a list of notes, including their object numbers
                        if self.showQuoteNotes and (objType == 'Quote'):
                            notesList = DBInterface.list_of_notes(Quote=tmpObj.number, includeNumber=True)
                            prompt = _('Quote Notes:\n')
                        elif self.showClipNotes and (objType == 'Clip'):
                            notesList = DBInterface.list_of_notes(Clip=tmpObj.number, includeNumber=True)
                            prompt = _('Clip Notes:\n')
                        elif self.showSnapshotNotes and (objType == 'Snapshot'):
                            notesList = DBInterface.list_of_notes(Snapshot=tmpObj.number, includeNumber=True)
                            prompt = _('Snapshot Notes:\n')
                        # If there are notes for this Clip ...
                        if len(notesList) > 0:
                            # Set the font for the notes
                            reportText.SetTxtStyle(fontSize = 10, fontBold = useBold,
                                                   parLeftIndent = baseIndent + 63, parRightIndent = 0,
                                                   parSpacingBefore = 0, parSpacingAfter = 0)

                            # Add the header to the report
                            reportText.WriteText(prompt)
#                            reportText.Newline()
                            # Iterate throught the list of notes ...
                            for note in notesList:
                                # ... load each note ...
                                tempNote = Note.Note(note[0])
                                # Turn bold on.
                                reportText.SetTxtStyle(fontBold = useBold,
                                                       parLeftIndent = baseIndent + 127, parRightIndent = 0,
                                                       parSpacingBefore = 0, parSpacingAfter = 0)

                                # Add the note ID to the report
                                reportText.WriteText('%s\n' % tempNote.id)
#                                reportText.Newline()
                                # Turn bold off, format Note
                                reportText.SetTxtStyle(fontBold = False,
                                                       parLeftIndent = baseIndent + 190, parRightIndent = 127,
                                                       parSpacingBefore = 0, parSpacingAfter = 0)

                                # Add the note text to the report
                                # (rstrip() prevents formatting problems following a note with a blank line at the end!)
                                reportText.WriteText('%s\n' % tempNote.text.rstrip())

#                                reportText.Newline()

            # If there are 20 or more items in the list, or at least 3 images ...
            if (self.collection != None) and ((len(majorList) >= 20) or (len(self.snapshotFilterList) > 3)):
                # Destroy the progress bar
                progress.Destroy()

        # If this IS an Episode-based or Document-based report ...
        else:

            # If there are 20 or more items in the list, or at least 3 images ...
            if ((len(majorList) >= 20) or (len(self.snapshotFilterList) > 3)):
                # ... create a Progress Dialog.  (The PARENT is needed to prevent the report being hidden
                #     behind Transana!)
                progress = wx.ProgressDialog(self.title, _('Assembling report contents'), parent=self.report)

            # If this is a Document Report ...
            if self.documentName != '':
                try:
                    tmpDoc = Document.Document(libraryID = self.seriesName, documentID = self.documentName)
                except:
                    tmpDoc = None

            # Iterate through the major list
            for itemRecord in majorList:
                if itemRecord['Type'] == 'Quote':
                    # our Filter comparison is based on Quote data
                    filterVal = (itemRecord['QuoteID'], itemRecord['CollectNum'], True)
                    filterList = self.quoteFilterList
                    prompt = _('Quote')
                    # Load the Quote Object
                    tmpObj = Quote.Quote(num=itemRecord['QuoteNum'])
                elif itemRecord['Type'] == 'Clip':
                    # our Filter comparison is based on Clip data
                    filterVal = (itemRecord['ClipID'], itemRecord['CollectNum'], True)
                    filterList = self.filterList
                    prompt = _('Clip')
                    # Load the Clip Object
                    tmpObj = Clip.Clip(itemRecord['ClipNum'])
                elif itemRecord['Type'] == 'Snapshot':
                    # our Filter comparison is based on Clip data
                    filterVal = (itemRecord['SnapshotID'], itemRecord['CollectNum'], True)                    
                    filterList = self.snapshotFilterList
                    prompt = _('Snapshot')
                    # Load the Snapshot Object
                    tmpObj = Snapshot.Snapshot(itemRecord['SnapshotNum'], suppressEpisodeError = True)
                # now that we have the filter comparison data, we see if it's actually in the Filter List.
                if filterVal in filterList:
                    # First, load the collection the current clip is in
                    collectionObj = Collection.Collection(itemRecord['CollectNum'])
                    # Set the font for the heading.
                    reportText.SetTxtStyle(fontSize = 12, fontBold = True)
                    reportText.SetTxtStyle(parAlign = wx.TEXT_ALIGNMENT_LEFT,
                                           parLeftIndent = 0,
                                           parSpacingBefore = 24, parSpacingAfter = 2)
                    # Add the header to the report
                    reportText.WriteText(prompt)
                    # Add the data to the report, the Clip/Snapshot ID in this case
                    reportText.WriteText(": ")
                    
                    # If we're showing Hyperlinks to Clips/Snapshots ...
                    if self.showHyperlink:
                        # Define a Hyperlink Style (Blue, underlined)
                        urlStyle = richtext.RichTextAttr()
                        urlStyle.SetFontFaceName('Courier New')
                        urlStyle.SetFontSize(12)
                        urlStyle.SetTextColour(wx.BLUE)
                        urlStyle.SetFontUnderlined(True)
                        # Apply the Hyperlink Style
                        reportText.BeginStyle(urlStyle)
                        # Insert the Hyperlink information, object type and object number
                        reportText.BeginURL("transana:%s=%d" % (itemRecord['Type'], tmpObj.number))

                    reportText.WriteText("%s\n" % (tmpObj.id,))
#                    reportText.Newline()

                    # If we're showing Hyperlinks to Clips/Snapshots ...
                    if self.showHyperlink:
                        # End the Hyperlink
                        reportText.EndURL();
                        # Stop using the Hyperlink Style
                        reportText.EndStyle();

                    # Set the font for the rest of the record
                    reportText.SetTxtStyle(fontSize = 10, fontBold = useBold, parLeftIndent = 63,
                                           parSpacingBefore = 0, parSpacingAfter = 0)
                    # Add the header to the report
                    reportText.WriteText(_('Collection:'))
                    # Turn bold off.
                    reportText.SetTxtStyle(fontBold = False)
                    # Add the data to the report, the full Collection path in this case
                    reportText.WriteText('  %s\n' % (collectionObj.GetNodeString(),))
#                    reportText.Newline()

                    # If we're supposed to show Media File name information ...
                    if self.showFile:
                        # Turn bold on.
                        reportText.SetTxtStyle(fontBold = useBold)
                        if itemRecord['Type'] == 'Quote':
                            # Add the header to the report
                            reportText.WriteText(_('Source File:'))
                        else:
                            # Add the header to the report
                            reportText.WriteText(_('File:'))
                        # Turn bold off.
                        reportText.SetTxtStyle(fontBold = False)
                        if itemRecord['Type'] == 'Quote':
                            if tmpDoc != None:
                                # Add the data to the report, the Quote Source file name in this case
                                reportText.WriteText('  %s ' % tmpDoc.imported_file)
                                prompt = unicode(_("imported on %s\n"), "utf8")
                                # sqlite gives a string rather than a datetime object
                                if isinstance(tmpDoc.import_date, str):
                                    # start exception handling in case of formatting problems
                                    try:
                                        # Convert the date string to a datetime object
                                        tmpDate = datetime.datetime.strptime(tmpDoc.import_date, '%Y-%m-%d %H:%M:%S')
                                        # Display the correct date
                                        reportText.WriteText(prompt % tmpDate.strftime('%x'))
                                    # If the conversion fails ...
                                    except ValueError:
                                        # ... display the un-converted string.
                                        reportText.WriteText(prompt % tmpDoc.import_date)
                                # MySQL returns a datetime object.
                                else:
                                    reportText.WriteText(prompt % tmpDoc.import_date.strftime('%x'))
                            else:
                                reportText.WriteText(_('Unknown') + '\n')
                        elif itemRecord['Type'] == 'Clip':
                            # Add the data to the report, the Clip media file name in this case
                            reportText.WriteText('  %s\n' % tmpObj.media_filename)
                            # Add Additional Media File info
                            for mediaFile in tmpObj.additional_media_files:
                                reportText.WriteText(_('       %s\n') % mediaFile['filename'])
                        elif itemRecord['Type'] == 'Snapshot':
                            # Add the data to the report, the Snapshot media file name in this case
                            reportText.WriteText('  %s\n' % tmpObj.image_filename)

                    # If we're supposed to show Clip Time  or Quote Position data ...
                    if self.showTime:
                        # Turn bold on.
                        reportText.SetTxtStyle(fontBold = useBold, parSpacingAfter = 0)
                        if itemRecord['Type'] == 'Quote':
                            # Add the header to the report
                            reportText.WriteText(_('Position:'))
                        else:
                            # Add the header to the report
                            reportText.WriteText(_('Time:'))
                        # Turn bold off.
                        reportText.SetTxtStyle(fontBold = False)
                        if itemRecord['Type'] == 'Quote':
                            # Add the data to the report, the Quote start and end characters in this case
                            reportText.WriteText('  %s - %s   (' % (itemRecord['StartChar'], itemRecord['EndChar']))
                        elif itemRecord['Type'] == 'Clip':
                            # Add the data to the report, the Clip start and stop times in this case
                            reportText.WriteText('  %s - %s   (' % (Misc.time_in_ms_to_str(itemRecord['ClipStart']), Misc.time_in_ms_to_str(itemRecord['ClipStop'])))
                        elif itemRecord['Type'] == 'Snapshot':
                            # Add the data to the report, the Snapshot start and stop times in this case
                            reportText.WriteText('  %s - %s   (' % (Misc.time_in_ms_to_str(tmpObj.episode_start), Misc.time_in_ms_to_str(tmpObj.episode_start + tmpObj.episode_duration)))
                        # Turn bold on.
                        reportText.SetTxtStyle(fontBold = useBold)
                        # Add the header to the report
                        reportText.WriteText(_('Length:'))
                        # Turn bold off.
                        reportText.SetTxtStyle(fontBold = False)
                        if itemRecord['Type'] == 'Quote':
                            # Add the data to the report, the Clip length in this case
                            reportText.WriteText('  %s)\n' % (itemRecord['EndChar'] - itemRecord['StartChar'],))
                        elif itemRecord['Type'] == 'Clip':
                            # Add the data to the report, the Clip length in this case
                            reportText.WriteText('  %s)\n' % (Misc.time_in_ms_to_str(itemRecord['ClipStop'] - itemRecord['ClipStart']),))
                        elif itemRecord['Type'] == 'Snapshot':
                            # Add the data to the report, the Snapshot duration in this case
                            reportText.WriteText('  %s)\n' % (Misc.time_in_ms_to_str(tmpObj.episode_duration),))
                        reportText.SetTxtStyle(parSpacingBefore = 0) #, parSpacingAfter = 0)
                            
                    # If we have a Quote ...
                    if itemRecord['Type'] == 'Quote':
                        # Increment the Quote Counter
                        self.quoteCount += 1
                        self.quoteTotalLength += tmpObj.end_char - tmpObj.start_char
                    # If we have a Clip ...
                    if itemRecord['Type'] == 'Clip':
                        # Increment the Clip Counter
                        self.clipCount += 1
                        # Add the Clip's length to the Clip Total Time accumulator
                        self.clipTotalTime += tmpObj.clip_stop - tmpObj.clip_start
                    # If we have a Snapshot ...
                    elif itemRecord['Type'] == 'Snapshot':
                        # Increment the Snapshot Counter
                        self.snapshotCount += 1
                        if tmpObj.episode_num > 0:
                            # Add the Snapshot's length to the Total Time accumulator
                            self.snapshotTotalTime += tmpObj.episode_duration

	    	    # If we're supposed to show Source Information, and the item HAS source information ...
		    if self.showSourceInfo:
		        reportText.SetTxtStyle(fontSize = 10, parSpacingAfter = 0)
                        if itemRecord['Type'] == 'Quote':
                            if tmpDoc != None:
                                # Turn bold on.
                                reportText.SetTxtStyle(fontBold = useBold)
                                # Add the header to the report
                                reportText.WriteText(_('Library:'))
                                reportText.SetTxtStyle(fontBold = False)
                                # Add the data to the report
                                reportText.WriteText('  %s\n' % tmpDoc.library_id)
                        else:
                            if tmpObj.series_id != '':
                                # Turn bold on.
                                reportText.SetTxtStyle(fontBold = useBold)
                                # Add the header to the report
                                reportText.WriteText(_('Library:'))
                                reportText.SetTxtStyle(fontBold = False)
                                # Add the data to the report
                                reportText.WriteText('  %s\n' % tmpObj.series_id)
                            if tmpObj.episode_num > 0:
                                # Turn bold on.
                                reportText.SetTxtStyle(fontBold = useBold)
                                # Add the header to the report
                                reportText.WriteText(_('Episode:'))
                                reportText.SetTxtStyle(fontBold = False)
                                # Add the data to the report
                                reportText.WriteText('  %s\n' % tmpObj.episode_id)

                    # If we are supposed to show Quote Text and the quote HAS a source document
                    #   or Clip Transcripts and the clip HAS a transcript ...
                    if self.showSourceInfo or self.showQuoteText or self.showTranscripts:
                        reportText.SetTxtStyle(fontSize = 10, parSpacingAfter = 0)

                        if (itemRecord['Type'] == 'Quote'):
                            # Turn bold on.
                            reportText.SetTxtStyle(fontBold =useBold)
                            # Add the header to the report
                            reportText.WriteText(_('Document:'))
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False)
                            # If a Source Document was found ...
                            if tmpDoc != None:
                                # Add the data to the report, the Source Document ID in this case
                                reportText.WriteText('  %s\n' % (tmpDoc.id,))
                            # if no Source Document is found, we have an orphan.
                            else:
                                # Add the data to the report, the Source Document ID in this case
                                reportText.WriteText('  %s\n' % _('The original Document has been deleted.'))

                            if self.showQuoteText:
                                # Turn bold on.
                                reportText.SetTxtStyle(fontSize = 10, fontFace = 'Courier New', fontBold = useBold,
                                                       parLeftIndent = 63, parRightIndent = 0,
                                                       parSpacingBefore = 0, parSpacingAfter = 0)
                                # Add the header to the report
                                reportText.WriteText(_('Quote Text:\n'))
                            
                                # Turn bold off.
                                reportText.SetTxtStyle(fontBold = False)
                                # Add the Quote Text to the report.  Quote text *must* be in XML format.

                                # Strip the time codes for the report
                                text = reportText.StripTimeCodes(tmpObj.text)

                                # Create the Transana XML to RTC Import Parser.  This is needed so that we can
                                # pull XML transcripts into the existing RTC without resetting the contents of
                                # the reportText RTC, which wipes out all accumulated Report data.
                                # Pass the reportText RTC and the desired additional margins in.
                                handler = PyXML_RTCImportParser.XMLToRTCHandler(reportText, (127, 127))
                                # Parse the transcript text, adding it to the reportText RTC
                                xml.sax.parseString(text, handler)

                        elif (itemRecord['Type'] == 'Clip'):
                            # for each Clip transcript:
                            for tr in tmpObj.transcripts:
                                if self.showSourceInfo:
                                    # Default the Episode Transcript to None in case the load fails
                                    episodeTranscriptObj = None
                                    # Begin exception handling
                                    try:
                                        # If the Clip Object has a defined Source Transcript ...
                                        if tr.source_transcript > 0:
                                            # ... try to load that source transcript
                                            # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                                            episodeTranscriptObj = Transcript.Transcript(tr.source_transcript, skipText=True)
                                    # if the record is not found (orphaned Clip)
                                    except TransanaExceptions.RecordNotFoundError:
                                        # We don't need to do anything.
                                        pass

                                    # If an Episode Transcript was found ...
                                    if episodeTranscriptObj != None:
                                        # Set the font for the Clip Transcript Header
                                        reportText.SetTxtStyle(fontSize = 10, fontFace = 'Courier New', fontBold = useBold,
                                                               parLeftIndent = 63, parRightIndent = 0,
                                                               parSpacingBefore = 0, parSpacingAfter = 0)
                            
                                        # Add the header to the report
                                        reportText.WriteText(_('Episode Transcript:'))
                                        # Turn bold off.
                                        reportText.SetTxtStyle(fontBold = False)
                                        # Add the data to the report, the Episode Transcript ID in this case
                                        reportText.WriteText('  %s\n' % (episodeTranscriptObj.id,))
                                    # if no Episode Transcript is found, we have an orphan.
                                    else:

                                        print "************************************************************************"
                                        print "*                      MISSING EPISODE TRANSCRIPT                      *"
                                        print "************************************************************************"

                                        # Turn bold on.
                                        reportText.SetTxtStyle(fontBold =useBold)
                                        # Add the header to the report
                                        reportText.WriteText(_('Episode Transcript:'))
                                        # Turn bold off.
                                        reportText.SetTxtStyle(fontBold = False)
                                        # Add the data to the report, the Episode Transcript ID in this case
                                        reportText.WriteText('  %s\n' % _('The Episode Transcript has been deleted.'))
#                                        reportText.Newline()

                                if self.showTranscripts:
                                    # Turn bold on.
                                    reportText.SetTxtStyle(fontSize = 10, fontFace = 'Courier New', fontBold = useBold,
                                                           parLeftIndent = 63, parRightIndent = 0,
                                                           parSpacingBefore = 0, parSpacingAfter = 0)
                                    # Add the header to the report
                                    reportText.WriteText(_('Clip Transcript:\n'))
                                
                                    # Turn bold off.
                                    reportText.SetTxtStyle(fontBold = False)
                                    # Add the Transcript to the report
                                    # Clip Transcripts could be in the old RTF format, or they could have been
                                    # updated to the new XML format.  These require different processing.

                                    # If we have a Rich Text Format document ...
                                    if tr.text[:5].lower() == u'{\\rtf':

                                        # Create a temporary RTC control
                                        tmpTxtCtrl = TranscriptEditor_RTC.TranscriptEditor(reportText.parent, pos=(-20, -20), suppressGDIWarning = True)
                                        # ... import the RTF data into the report text
                                        tmpTxtCtrl.LoadRTFData(tr.text, clearDoc=True)
                                        # Pull the data back out of the control as XML
                                        tmpText = tmpTxtCtrl.GetFormattedSelection('XML')
                                        # Strip the time codes for the report
                                        tmpText = reportText.StripTimeCodes(tmpText)

                                        # Create the Transana XML to RTC Import Parser.  This is needed so that we can
                                        # pull XML transcripts into the existing RTC without resetting the contents of
                                        # the reportText RTC, which wipes out all accumulated Report data.
                                        # Pass the reportText RTC and the desired additional margins in.
                                        handler = PyXML_RTCImportParser.XMLToRTCHandler(reportText, (127, 127))
                                        # Parse the transcript text, adding it to the reportText RTC
                                        xml.sax.parseString(tmpText, handler)

                                        # Destroy the temporary RTC control
                                        tmpTxtCtrl.Destroy()

                                    # If we have an XML document ...
                                    elif tr.text[:5].lower() == u'<?xml':
                                        # Strip the time codes for the report
                                        tr.text = reportText.StripTimeCodes(tr.text)

                                        # Create the Transana XML to RTC Import Parser.  This is needed so that we can
                                        # pull XML transcripts into the existing RTC without resetting the contents of
                                        # the reportText RTC, which wipes out all accumulated Report data.
                                        # Pass the reportText RTC and the desired additional margins in.
                                        handler = PyXML_RTCImportParser.XMLToRTCHandler(reportText, (127, 127))
                                        # Parse the transcript text, adding it to the reportText RTC
                                        xml.sax.parseString(tr.text, handler)

                                    # If we have a transcript that is neither RTF nor XML (shouldn't happen!)
                                    else:
                                        # ... then just import it directly.  Treat it as plain text.
                                        # (rstrip() prevents formatting problems when a transcript ends with blank lines)
                                        reportText.WriteText(tr.text.rstrip())

                        # If we have a Snapshot ...
                        elif itemRecord['Type'] == 'Snapshot':
                            if self.showSourceInfo:
                                # ... if the Snapshot has a defined Transcript ...
                                if tmpObj.transcript_num > 0:
                                    # Turn bold on.
                                    reportText.SetTxtStyle(fontSize = 10, fontFace = 'Courier New', fontBold = useBold,
                                                           parLeftIndent = 63, parRightIndent = 0,
                                                           parSpacingBefore = 0, parSpacingAfter = 0)
                                    # Add the header to the report
                                    reportText.WriteText(_('Episode Transcript:'))
                                    reportText.SetTxtStyle(fontBold = False)
                                    # Add the data to the report
                                    reportText.WriteText('  %s\n' % tmpObj.transcript_id)

                    # If we have a Snapshot, and we're displaying Full, Medium, or Small images, show the actual IMAGE
                    if (itemRecord['Type'] == 'Snapshot') and (self.showSnapshotImage in [0, 1, 2]):
                        try:
                            # Open a HIDDEN Snapshot Window
                            tmpSnapshotWindow = SnapshotWindow.SnapshotWindow(TransanaGlobal.menuWindow, -1, tmpObj.id, tmpObj, showWindow=False)
                            # Get the cropped, coded image from the Snapshot Window
                            tmpBMP = tmpSnapshotWindow.CopyBitmap()
                            # Close the hidden Snapshot Window
                            tmpSnapshotWindow.Close()
                            # Explicitly Delete the Temporary Snapshot Window
                            tmpSnapshotWindow.Destroy()

                            # Convert the Bitmap to an Image so it can be rescaled
                            tmpImage = tmpBMP.ConvertToImage()
                            # Get the Image Size
                            (imgWidth, imgHeight) = tmpImage.GetSize()
                            # We need the SMALLER of the current image size and the current Transcript Window size
                            # (Adjust width for scrollbar size!)
                            maxWidth = min(float(imgWidth), (reportText.GetSize()[0] - 20.0) * 0.75)
                            maxHeight = min(float(imgHeight), reportText.GetSize()[1] * 0.80)
                            # if we're using "Medium" size ...
                            if self.showSnapshotImage == 1:
                                # ... set the image max size to 500 pixels
                                maxWidth = min(maxWidth, 500)
                                maxHeight = min(maxHeight, 500)
                            # If we're using "Small" size ...
                            elif self.showSnapshotImage == 2:
                                # ... set the image max size to 250 pixels
                                maxWidth = min(maxWidth, 250)
                                maxHeight = min(maxHeight, 250)
                            # Determine the scaling factor for adjusting the image size
                            scaleFactor = min(maxWidth / float(imgWidth), maxHeight / float(imgHeight))
                            # If the image is too BIG, it needs to be re-scaled ...
                            if scaleFactor < 1.0:
                                # ... so rescale the image to fit in the current Transcript window.  Use slower high quality rescale.
                                tmpImage.Rescale(int(imgWidth * scaleFactor), int(imgHeight * scaleFactor), quality=wx.IMAGE_QUALITY_HIGH)

                            # If we have anything but a large image ...
                            if self.showSnapshotImage > 0:
                                # ... alter the Paragraph Spacing here.
                                reportText.SetTxtStyle(parSpacingBefore = 24, parSpacingAfter = 24)
                            # If we have a large image ...
                            else:
                                # ... alter the Paragraph Spacing here, and DON'T indent the image itself, so it can be as large as possible!
                                reportText.SetTxtStyle(parLeftIndent = 0, parSpacingBefore = 24, parSpacingAfter = 24)
                            # Add the image to the transcript
                            reportText.WriteImage(tmpImage)
                            # Delete the temporary image, bitmap, and device contexts
                            tmpImage.Destroy()
                            tmpBMP.Destroy()
                            # Add some more blank space.
                            reportText.WriteText('\n')
                            # Reset the Paragraph Spacing here
                            reportText.SetTxtStyle(parLeftIndent = 63, parSpacingBefore = 0, parSpacingAfter = 0)

                        # Detect Image Loading problems
                        except TransanaExceptions.ImageLoadError, e:

                            if DEBUG:
                                print "ReportGenerator.OnDisplay():"
                                print tmpObj.GetNodeString()
                                print
                                print sys.exc_info()[0]
                                print sys.exc_info()[1]
                                import traceback
                                traceback.print_exc(file=sys.stdout)
                                print

                            # If we're displaying Image Load Errors ...
                            if not skippingImageError:
                                # ... build the error message and display it.
                                tmpDlg = Dialogs.ErrorDialog(self.report, e.explanation, includeSkipCheck=True)
                                tmpDlg.ShowModal()
                                # See if the user is requesting that further messages be skipped
                                skippingImageError = tmpDlg.GetSkipCheck()
                                tmpDlg.Destroy()
                            
                            # Alter the Paragraph Spacing here
                            reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 0)
                            # Add some more blank space.
                            reportText.WriteText('\n')
                            # Add the image to the transcript
                            reportText.WriteText(e.explanation)
                            # Alter the Paragraph Spacing here
                            reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 24)
                            # Add some more blank space.
                            reportText.WriteText('\n')
                            # Reset the Paragraph Spacing here
                            reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 0)

                        # Detect PyAssertionError
                        except wx._core.PyAssertionError, e:

                            if DEBUG:
                                print "ReportGenerator.OnDisplay():"
                                print tmpObj.GetNodeString()
                                print
                                print sys.exc_info()[0]
                                print sys.exc_info()[1]
                                import traceback
                                traceback.print_exc(file=sys.stdout)
                                print

                            # ... build the error message and display it.
                            tmpDlg = Dialogs.ErrorDialog(self.report, tmpObj.GetNodeString() + '\n' + e.message)
                            tmpDlg.ShowModal()
                            tmpDlg.Destroy()
                            
                            # Alter the Paragraph Spacing here
                            reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 0)
                            # Add some more blank space.
                            reportText.WriteText('\n')
                            # Add the image to the transcript
                            reportText.WriteText(e.message)
                            # Alter the Paragraph Spacing here
                            reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 24)
                            # Add some more blank space.
                            reportText.WriteText('\n')
                            # Reset the Paragraph Spacing here
                            reportText.SetTxtStyle(parSpacingBefore = 0, parSpacingAfter = 0)


                        # We want to display the Coding Key for the Snapshot.  But we only want to display the information
                        # for those codes that are VISIBLE in the image (even if they're not included on the Snapshot!)

                    # If we're supposed to show Snapshot Coding ...
                    if (itemRecord['Type'] == 'Snapshot') and self.showSnapshotCoding:
                        # Create an empty list to hold the keys of visible codes
                        tmpKeys = []
                        # For each Coding Object in the Snapshot ...
                        for key in tmpObj.codingObjects.keys():
                            # ... if the Coding Object is visible and the Keyword it represents isn't already in the tmpKeys list ...
                            if tmpObj.codingObjects[key]['visible'] and \
                               (not (tmpObj.codingObjects[key]['keywordGroup'], tmpObj.codingObjects[key]['keyword']) in tmpKeys):
                                # ... add the Keyword to the tmpKeys List
                                tmpKeys.append((tmpObj.codingObjects[key]['keywordGroup'], tmpObj.codingObjects[key]['keyword']))

                        # If there are keywords in the list
                        if len(tmpKeys) > 0:
                            # Turn bold on.
                            reportText.SetTxtStyle(fontBold = useBold, parLeftIndent=63, parRightIndent = 0,
                                                   parSpacingAfter = 0)
                            # Add the header to the report
                            reportText.WriteText(_('Snapshot Coding Key:'))
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False)
                            # Finish the Line
                            reportText.WriteText('\n')
#                                reportText.Newline()
                            # Set formatting for Keyword lines
                            reportText.SetTxtStyle(parLeftIndent = 127, parRightIndent = 0,
                                                   parSpacingBefore = 0, parSpacingAfter = 0)
                            # Sort the Keywords
                            tmpKeys.sort()
                            # For each keyword ...
                            for key in tmpKeys:
                                
                                # See if the keyword should be included, based on the Keyword Filter List
                                if (key[0], key[1], True) in self.keywordFilterList:

                                    # Add the keyword
                                    reportText.WriteText('%s : %s  (' % (key[0], key[1]))
                                    # if THIS Snapshot with THIS Keyword has NOT already been counted ...
                                    if not (('Snapshot', tmpObj.number, key[0], key[1]) in self.itemsCounted):
                                        # Add this Keyword to the Keyword Counts
                                        if keywordCounts.has_key('%s : %s' % (key[0], key[1])):
                                            keywordCounts['%s : %s' % (key[0], key[1])] += 1
                                            if tmpObj.episode_num > 0:
                                                keywordTimes['%s : %s' % (key[0], key[1])] += tmpObj.episode_duration
                                        else:
                                            keywordCounts['%s : %s' % (key[0], key[1])] = 1
                                            keywordTimes['%s : %s' % (key[0], key[1])] = 0
                                            keywordLengths['%s : %s' % (key[0], key[1])] = 0
                                            if tmpObj.episode_num > 0:
                                                keywordTimes['%s : %s' % (key[0], key[1])] += tmpObj.episode_duration
                                        # Remember that THIS Snapshot with THIS Keyword HAS been counted now
                                        self.itemsCounted.append(('Snapshot', tmpObj.number, key[0], key[1]))

                                    # Get the Graphic for the Coding Key
                                    tmpImage = SnapshotWindow.CodingKeyGraphic(tmpObj.keywordStyles[key])
                                    # Add the Image to the Report
                                    reportText.WriteImage(tmpImage)
                                    # Delete the temporary image
                                    tmpImage.Destroy()
                                    # Add the Shape Description and a Line Feed
                                    reportText.WriteText(')\n')

                    # Reset the font.  It could have been contaminated by the Clip Transcript
                    reportText.SetTxtStyle(fontFace='Courier New', fontSize = 10, fontColor = wx.Colour(0, 0, 0),
                                           fontBgColor = wx.Colour(255, 255, 255), fontBold = False, fontItalic = False,
                                           fontUnderline = False)
                    reportText.SetTxtStyle(parAlign = wx.TEXT_ALIGNMENT_LEFT,
                                           parLineSpacing = wx.TEXT_ATTR_LINE_SPACING_NORMAL,
                                           parLeftIndent = 0,
                                           parRightIndent = 0,
                                           parSpacingBefore = 0,
                                           parSpacingAfter = 0,
                                           parTabs = [])

                    if DEBUG:
                        print "RESET all FONT and PARAGRAPH settings"

                    # if we are supposed to show Keywords ...
                    if self.showKeywords:
                        # If we have a Quote ...
                        if itemRecord['Type'] == 'Quote':
                            # ... use the Quote Number
                            recNum = itemRecord['QuoteNum']
                            prompt = _('Quote Keywords:')
                        # If we have a Clip ...
                        elif itemRecord['Type'] == 'Clip':
                            # ... use the Clip Number
                            recNum = itemRecord['ClipNum']
                            prompt = _('Clip Keywords:')
                        # if we have a Snapshot ...
                        elif itemRecord['Type'] == 'Snapshot':
                            # use the Snapshot Number
                            recNum = itemRecord['SnapshotNum']
                            prompt = _('Whole Snapshot Keywords:')
                        # If there are keywords in the list ... (even if they might all get filtered out ...)
                        if len(minorList[(itemRecord['Type'], recNum)]) > 0:
                            # Set Formatting
                            reportText.SetTxtStyle(fontSize = 10, fontBold = useBold, parLeftIndent = 63, parRightIndent = 0,
                                                   parLineSpacing = wx.TEXT_ATTR_LINE_SPACING_NORMAL,
                                                   parSpacingBefore = 0, parSpacingAfter = 0)
                            # Add the header to the report
                            reportText.WriteText(prompt)
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False)
                            reportText.WriteText('\n')
#                            reportText.Newline()
                            reportText.SetTxtStyle(parLeftIndent = 127, parSpacingBefore = 0, parSpacingAfter = 0)

                        # Iterate through the list of Keywords for the group
                        for (keywordGroup, keyword, example) in minorList[(itemRecord['Type'], recNum)]:
                            # See if the keyword should be included, based on the Keyword Filter List
                            if (keywordGroup, keyword, True) in self.keywordFilterList:
                                # Add the Keyword to the report
                                reportText.WriteText('%s : %s\n' % (keywordGroup, keyword))
#                                reportText.Newline()
                                # If we have a Clip ...
                                if itemRecord['Type'] == 'Quote':
                                    # if THIS Quote with THIS Keyword has NOT already been counted ...
                                    if not (('Quote', tmpObj.number, keywordGroup, keyword) in self.itemsCounted):
                                        # Add this Keyword to the Keyword Counts
                                        if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] += 1
                                            keywordLengths['%s : %s' % (keywordGroup, keyword)] += tmpObj.end_char - tmpObj.start_char
                                        else:
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                                            keywordTimes['%s : %s' % (keywordGroup, keyword)] = 0
                                            keywordLengths['%s : %s' % (keywordGroup, keyword)] = tmpObj.end_char - tmpObj.start_char
                                        # Remember that THIS Quote with THIS Keyword HAS been counted now
                                        self.itemsCounted.append(('Quote', tmpObj.number, keywordGroup, keyword))
                                # If we have a Clip ...
                                elif itemRecord['Type'] == 'Clip':
                                    # if THIS Clip with THIS Keyword has NOT already been counted ...
                                    if not (('Clip', tmpObj.number, keywordGroup, keyword) in self.itemsCounted):
                                        # Add this Keyword to the Keyword Counts
                                        if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] += 1
                                            keywordTimes['%s : %s' % (keywordGroup, keyword)] += tmpObj.clip_stop - tmpObj.clip_start
                                        else:
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                                            keywordTimes['%s : %s' % (keywordGroup, keyword)] = tmpObj.clip_stop - tmpObj.clip_start
                                            keywordLengths['%s : %s' % (keywordGroup, keyword)] = 0
                                        # Remember that THIS Clip with THIS Keyword HAS been counted now
                                        self.itemsCounted.append(('Clip', tmpObj.number, keywordGroup, keyword))
                                # If we have a Snapshot ...
                                elif itemRecord['Type'] == 'Snapshot':
                                    # if THIS Snapshot with THIS Keyword has NOT already been counted ...
                                    if not (('Snapshot', tmpObj.number, keywordGroup, keyword) in self.itemsCounted):
                                        # Add this Keyword to the Keyword Counts
                                        if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] += 1
                                            if (tmpObj.episode_num > 0) and \
                                               ((self.episodeName != None) or (self.collection != None) or (self.searchColl != None)):
                                                keywordTimes['%s : %s' % (keywordGroup, keyword)] += tmpObj.episode_duration
                                        else:
                                            keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                                            keywordTimes['%s : %s' % (keywordGroup, keyword)] = 0
                                            keywordLengths['%s : %s' % (keywordGroup, keyword)] = 0
                                            if (tmpObj.episode_num > 0) and \
                                               ((self.episodeName != None) or (self.collection != None) or (self.searchColl != None)):
                                                keywordTimes['%s : %s' % (keywordGroup, keyword)] += tmpObj.episode_duration
                                        # Remember that THIS Snapshot with THIS Keyword HAS been counted now
                                        self.itemsCounted.append(('Snapshot', tmpObj.number, keywordGroup, keyword))
                                    
                    # if we are supposed to show Comments ...
                    if self.showComments:
                        # If there are keywords in the list ... (even if they might all get filtered out ...)
                        if itemRecord['Comment'] != u'':
                            # Set the font for the Comments
                            reportText.SetTxtStyle(fontSize = 10, fontBold = useBold, parLeftIndent = 63,
                                                   parSpacingBefore = 0, parSpacingAfter = 0)
                            # If we have a Quote ...
                            if itemRecord['Type'] == 'Quote':
                                # Add the header to the report
                                reportText.WriteText(_('Quote Comment:\n'))
                            # If we have a Clip ...
                            elif itemRecord['Type'] == 'Clip':
                                # Add the header to the report
                                reportText.WriteText(_('Clip Comment:\n'))
                            elif itemRecord['Type'] == 'Snapshot':
                                # Add the header to the report
                                reportText.WriteText(_('Snapshot Comment:\n'))
                            # Turn bold off.
                            reportText.SetTxtStyle(fontBold = False, parLeftIndent = 127,
                                                   parSpacingBefore = 0, parSpacingAfter = 0)
                            # Add the content of the Clip Comment to the report
                            reportText.WriteText('%s\n' % tmpObj.comment)
#                            reportText.Newline()

                    # If we are supposed to show Clip or Snapshot Notes ...
                    if self.showQuoteNotes or self.showClipNotes or self.showSnapshotNotes:
                        # initialize the Notes List to empty so that notes of a type that should NOT be displayed
                        # will be properly skipped
                        notesList = []
                        # ... get a list of notes, including their object numbers
                        if self.showQuoteNotes and (itemRecord['Type'] == 'Quote'):
                            notesList = DBInterface.list_of_notes(Quote=tmpObj.number, includeNumber=True)
                            prompt = _('Quote Notes:\n')
                        elif self.showClipNotes and (itemRecord['Type'] == 'Clip'):
                            notesList = DBInterface.list_of_notes(Clip=tmpObj.number, includeNumber=True)
                            prompt = _('Clip Notes:\n')
                        elif self.showSnapshotNotes and (itemRecord['Type'] == 'Snapshot'):
                            notesList = DBInterface.list_of_notes(Snapshot=tmpObj.number, includeNumber=True)
                            prompt = _('Snapshot Notes:\n')
                        # If there are notes for this object ...
                        if len(notesList) > 0:
                            # Set the font for the Notes
                            reportText.SetTxtStyle(fontSize = 10, fontBold = useBold, parLeftIndent = 63,
                                                   parSpacingBefore = 0, parSpacingAfter = 0)
                            # Add the header to the report
                            reportText.WriteText(prompt)
#                            reportText.Newline()
                            # Iterate throught the list of notes ...
                            for note in notesList:
                                # ... load each note ...
                                tempNote = Note.Note(note[0])
                                # Turn bold on.
                                reportText.SetTxtStyle(fontBold = useBold, parLeftIndent = 127, parRightIndent = 127,
                                                       parSpacingBefore = 0, parSpacingAfter = 0)
                                # Add the note ID to the report
                                reportText.WriteText('%s\n' % tempNote.id)
#                                reportText.Newline()
                                # Turn bold off.
                                reportText.SetTxtStyle(fontBold = False, parLeftIndent = 190, parRightIndent = 127,
                                                       parSpacingBefore = 0, parSpacingAfter = 0)
                                # Add the note text to the report (rstrip() prevents formatting problems when a note ends with blank lines)
                                reportText.WriteText('%s\n' % tempNote.text.rstrip())
#                                reportText.Newline()

                    # If there are 20 or more items in the list, or at least 3 images ...
                    if ((len(majorList) >= 20) or (len(self.snapshotFilterList) > 3)):
                        # Determine the correct percentage figure ...
                        val = min(100, int(float(self.quoteCount + self.clipCount + self.snapshotCount) / float(len(majorList)) * 100))
                        # ... update the progress bar
                        progress.Update(val)

                    reportText.WriteText('\n')
#                    reportText.Newline()

            # If there are 20 or more items in the list, or at least 3 images ...
            if ((len(majorList) >= 20) or (len(self.snapshotFilterList) > 3)):
                # Destroy the progress bar
                progress.Destroy()

        if 'wxMac' in wx.PlatformInfo:
            keyWidth = 50
            macAdjust = 40
        else:
            keyWidth = 52
            macAdjust = 0

        # Now add the Report Summary
        # If there's data in the majorList OR if we're showing TIMES or if we're showing Snapshot Coding and there's coding to show ...
        if (self.showKeywords and (len(majorList) != 0)) or \
           self.showTime or \
           (self.showSnapshotCoding and (len(self.snapshotFilterList) > 0) and (len(keywordCounts) > 0)):
            # First, set the font for the summary
            reportText.SetTxtStyle(fontSize = 12, fontBold = True, parLeftIndent = 0, parRightIndent = 0,
                                   parSpacingBefore = 36, parSpacingAfter = 0, overrideGDIWarning = True) 
            # Add the section heading to the report
            reportText.WriteText(_("Summary") + '\n')
#            reportText.SetTxtStyle(fontSize = 10, fontBold = False, overrideGDIWarning = True)
            # Set the font for the Keywords
#            reportText.Newline()
            reportText.SetTxtStyle(fontSize = 10, fontBold = False,
                                   parLeftIndent = 32, parRightIndent = 0, parSpacingBefore = 0, parSpacingAfter = 0,
                                   parTabs = [1143 - macAdjust], overrideGDIWarning = True)

            # Get a list of the Keyword Group : Keyword pairs that have been used
            countKeys = keywordCounts.keys()
            # Sort the list
            countKeys.sort()
            # Add the sorted keywords to the summary with their counts
            for key in countKeys:
                prompt = '%s\t%5d'
                if (self.documentName != None) or (self.episodeName != None) or (self.collection != None) or (self.searchColl != None):
                    if len(keywordLengths) > 0:
                        # Add the text to the report.
                        prompt += '  %10d'
                    if len(keywordTimes) > 0:
                        # Add the text to the report.
                        prompt += '  %10s'
                data = (key[:keyWidth], keywordCounts[key])

                if (self.documentName != None) or (self.episodeName != None) or (self.collection != None) or (self.searchColl != None):
                    if len(keywordLengths) > 0:
                        if keywordLengths.has_key(key):
                            data += (keywordLengths[key],)
                        else:
                            data += (0,)
                    if len(keywordTimes) > 0:
                        if keywordTimes.has_key(key):
                            data += (Misc.time_in_ms_to_str(keywordTimes[key]),)
                        else:
                            data += (Misc.time_in_ms_to_str(0), )

                prompt += '\n'
                # Add the text to the report.
                reportText.WriteText(prompt % data)
#                reportText.Newline()

            # Show total number of items reported
            if self.quoteCount + self.clipCount + self.snapshotCount > 0:
                # Set the font for the data
                reportText.SetTxtStyle(fontSize = 10, parSpacingBefore = 0, parSpacingAfter = 0,
                                       parTabs = [1121 - macAdjust, 1250 - macAdjust], overrideGDIWarning = True)
                # Add the total Item Count
                reportText.WriteText(u'\n' + (unicode(_('Items:'), 'utf8') + u'\t%6d\n') % (self.quoteCount + self.clipCount + self.snapshotCount))
#                reportText.Newline()

            # If we have Documents / Quotes ...
            if self.quoteCount > 0:
                # If we have a Library Report ...
                if self.collection == None and self.documentName == None and self.episodeName == None:
                    prompt = u'  ' + unicode(_('Documents:'), 'utf8') + u'\t%6d'
                else:
                    prompt = u'  ' + unicode(_('Quotes:'), 'utf8') + u'\t%6d'
                data = (self.quoteCount,)
                if self.showKeywords or self.showTime:
                    prompt += u'  %10s'
                    data += (self.quoteTotalLength,)
                # Add the total Item Count
                reportText.WriteText(prompt % data)
                reportText.WriteText('\n')
                
            # If we have Episodes / Clips ...
            if self.clipCount > 0:
                # If we have a Library Report ...
                if self.collection == None and self.documentName == None and self.episodeName == None:
                    prompt = u'  ' + unicode(_('Episodes:'), 'utf8') + u'\t%6d'
                else:
                    prompt = u'  ' + unicode(_('Clips:'), 'utf8') + u'\t%6d'
                data = (self.clipCount,)
                if self.showKeywords or self.showTime:
                    if ((self.documentName != None) or (self.episodeName != None) or \
                        (self.collection != None) or (self.searchColl != None)) and \
                        (len(keywordLengths) > 0):
                        prompt += u'            '
                    prompt += u'   %s'
                    data += (Misc.time_in_ms_to_str(self.clipTotalTime),)
                # Add the total Item Count
                reportText.WriteText(prompt % data)
                reportText.WriteText('\n')

            # If we have Snapshots ...
            if self.snapshotCount > 0:
                prompt = u'  ' + unicode(_('Snapshots:'), 'utf8') + u'\t%6d'
                data = (self.snapshotCount,)
                if self.showKeywords or self.showTime:
                    if ((self.documentName != None) or (self.episodeName != None) or \
                        (self.collection != None) or (self.searchColl != None)) and \
                        (len(keywordLengths) > 0):
                        prompt += u'            '
                    prompt += u'   %s'
                    data += (Misc.time_in_ms_to_str(self.snapshotTotalTime),)
                # Add the total Item Count
                reportText.WriteText(prompt % data)
                reportText.WriteText('\n')

        # Make the control read only, now that it's done
        reportText.SetReadOnly(True)

    def OnFilter(self, event):
        """ This method, required by TextReport, implements the call to the Filter Dialog.  It needs to be
            in the report parent because the TextReport doesn't know the appropriate filter parameters. """
        # See if we're loading the Default profile.  This is signalled by an event of None!
        if event == None:
            loadDefault = True
        else:
            loadDefault = False
        # Sort the Keyword Filter List
        self.keywordFilterList.sort()
        # If a Collection Name is passed in ...
        if (self.collection != None) or ((self.searchColl != None) and (self.treeCtrl != None)):
            # First, let's determine the correct Report Scope.  If we have a regular collection ...
            if self.collection != None:
                # ... we can use the collection number
                reportScope = self.collection.number
            # If we don't have a regular collection, we are using a Search Collection.  In that case ...
            else:
                # ... we get the collection number from the tree control's PyData for that tree entry.
                reportScope = self.treeCtrl.GetPyData(self.searchColl).recNum
            # See if there are Quotes in the Quote Filter List
            quoteFilter = (len(self.quoteFilterList) > 0)
            # See if there are Clips in the Filter List
            clipFilter = (len(self.filterList) > 0)
            # See if there are Snapshots in the Snapshot Filter List
            snapshotFilter = (len(self.snapshotFilterList) > 0)
            # See if there are Keywords in the Filter List
            keywordFilter = (len(self.keywordFilterList) > 0)
            # Define the Filter Dialog.  We need reportType 12 to identify the Collection Report, and we
            # need only the Clip Filter and Keyword Filter for this report.  We want to show the file name,
            # clip time data, Clip Transcripts, Clip Keywords, Comments, Collection Note, Clip Note, and
            # Nested Data options.
            dlgFilter = FilterDialog.FilterDialog(self.report,
                                                  -1,
                                                  self.title,
                                                  reportType=12,
                                                  reportScope=reportScope,
                                                  loadDefault=loadDefault,
                                                  configName=self.configName,
                                                  quoteFilter=quoteFilter,
                                                  clipFilter=clipFilter,
                                                  snapshotFilter=snapshotFilter,
                                                  keywordFilter=keywordFilter,
                                                  reportContents=True,
                                                  showFile=self.showFile,
                                                  showTime=self.showTime,
                                                  showSourceInfo=self.showSourceInfo,
                                                  showQuoteText=self.showQuoteText,
                                                  showClipTranscripts=self.showTranscripts,
                                                  showSnapshotImage=self.showSnapshotImage,
                                                  showSnapshotCoding=self.showSnapshotCoding,
                                                  showKeywords=self.showKeywords,
                                                  showComments=self.showComments,
                                                  showNestedData=self.showNested,
                                                  showHyperlink=self.showHyperlink,
                                                  showQuoteNotes=self.showQuoteNotes,
                                                  showCollectionNotes=self.showCollectionNotes,
                                                  showClipNotes=self.showClipNotes,
                                                  showSnapshotNotes=self.showSnapshotNotes )
            # If there are Quotes ...
            if quoteFilter:
                # Populate the Filter Dialog with Quotes
                dlgFilter.SetQuotes(self.quoteFilterList)
            # If there are Clips ...
            if clipFilter:
                # Populate the Filter Dialog with Clips
                dlgFilter.SetClips(self.filterList)
            # if there are Snapshots ...
            if snapshotFilter:
                # ... populate the Filter Dialog with Snapshots
                dlgFilter.SetSnapshots(self.snapshotFilterList)
            # If there are Keywords ...
            if keywordFilter:
                # Populate the Filter Dialog with Keywords
                dlgFilter.SetKeywords(self.keywordFilterList)
            # If we're loading the Default configuration ...
            if loadDefault:
                # ... get the list of existing configuration names.
                profileList = dlgFilter.GetConfigNames()
                # If (translated) "Default" is in the list ...
                # (NOTE that the default config name is stored in English, but gets translated by GetConfigNames!)
                if unicode(_('Default'), 'utf8') in profileList:
                    # ... then signal that we need to load the config.
                    dlgFilter.OnFileOpen(None)
                    # Fake that we asked the user for a filter name and got an OK
                    result = wx.ID_OK
                # If we're loading a Default profile, but there's none in the list, we can skip
                # the rest of the Filter method by pretending we got a Cancel from the user.
                else:
                    result = wx.ID_CANCEL
            # If we're not loading a Default profile ...
            else:
                # ... we need to show the Filter Dialog here.
                result = dlgFilter.ShowModal()
                
            # If the user clicks OK (or we have a Default config)
            if result == wx.ID_OK:
                # ... get the filter data ...
                if quoteFilter:
                    self.quoteFilterList = dlgFilter.GetQuotes()
                    self.showQuoteText = dlgFilter.GetShowQuoteText()
                    self.showQuoteNotes = dlgFilter.GetShowQuoteNotes()
                # ... get the filter data ...
                if clipFilter:
                    self.filterList = dlgFilter.GetClips()
                    self.showTranscripts = dlgFilter.GetShowClipTranscripts()
                    self.showClipNotes = dlgFilter.GetShowClipNotes()
                if snapshotFilter:
                    self.snapshotFilterList = dlgFilter.GetSnapshots()
                    self.showSnapshotImage = dlgFilter.GetShowSnapshotImage()
                    self.showSnapshotCoding = dlgFilter.GetShowSnapshotCoding()
                    self.showSnapshotNotes = dlgFilter.GetShowSnapshotNotes()
                if keywordFilter:
                    self.keywordFilterList = dlgFilter.GetKeywords()
                    self.showKeywords = dlgFilter.GetShowKeywords()
                if clipFilter or snapshotFilter:
                    self.showFile = dlgFilter.GetShowFile()
                    self.showTime = dlgFilter.GetShowTime()
                    self.showSourceInfo = dlgFilter.GetShowSourceInfo()
                self.showComments = dlgFilter.GetShowComments()
                self.showCollectionNotes = dlgFilter.GetShowCollectionNotes()
                self.showNested = dlgFilter.GetShowNestedData()
                self.showHyperlink = dlgFilter.GetShowHyperlink()
                # Remember the configuration name for later reuse
                self.configName = dlgFilter.configName

                dlgFilter.Destroy()
                
                # ... and signal the TextReport that the filter is to be applied.
                return True
            # If the filter is cancelled by the user ...
            else:
                # ... signal the TextReport that the filter is NOT to be applied.
                return False

        # If a Document Name is passed in ...
        elif (self.documentName != None):
            # Load the Document object
            tempDocument = Document.Document(libraryID=self.seriesName, documentID=self.documentName, skipText=True)
            # See if there are Quotes in the QuoteFilter List
            quoteFilter = (len(self.quoteFilterList) > 0)
##            # See if there are Snapshots in the Snapshot Filter List
##            snapshotFilter = (len(self.snapshotFilterList) > 0)
            # See if there are Keywords in the Filter List
            keywordFilter = (len(self.keywordFilterList) > 0)
            # Define the Filter Dialog.  We need reportType 19 to identify the Document Report, and we
            # need only the Quote Filter and Keyword Filter for this report.  We want to show the file name,
            # quote position data, Clip Transcripts, Clip Keywords, Comments and clip notes options.
            dlgFilter = FilterDialog.FilterDialog(self.report,
                                                  -1,
                                                  self.title,
                                                  reportType=19,
                                                  reportScope=tempDocument.number,
                                                  loadDefault=loadDefault,
                                                  configName=self.configName,
                                                  quoteFilter=quoteFilter,
##                                                  snapshotFilter=snapshotFilter,
                                                  keywordFilter=keywordFilter,
                                                  reportContents=True,
                                                  showHyperlink=self.showHyperlink,
                                                  showFile=self.showFile,
                                                  showTime=self.showTime,
                                                  showSourceInfo=self.showSourceInfo,
                                                  showQuoteText=self.showQuoteText,
##                                                  showSnapshotImage=self.showSnapshotImage,
##                                                  showSnapshotCoding=self.showSnapshotCoding,
                                                  showKeywords=self.showKeywords,
                                                  showComments=self.showComments,
                                                  showQuoteNotes=self.showQuoteNotes  ) # ,
##                                                  showSnapshotNotes=self.showSnapshotNotes )
            # If there are Quotes ...
            if quoteFilter:
                # Populate the Filter Dialog with Quotes
                dlgFilter.SetQuotes(self.quoteFilterList)
##            # if there are Snapshots ...
##            if snapshotFilter:
##                # ... populate the Filter Dialog with Snapshots
##                dlgFilter.SetSnapshots(self.snapshotFilterList)
            if keywordFilter:
                # Populate the Filter Dialog with Keywords
                dlgFilter.SetKeywords(self.keywordFilterList)
            # If we're loading the Default configuration ...
            if loadDefault:
                # ... get the list of existing configuration names.
                profileList = dlgFilter.GetConfigNames()
                # If (translated) "Default" is in the list ...
                # (NOTE that the default config name is stored in English, but gets translated by GetConfigNames!)
                if unicode(_('Default'), 'utf8') in profileList:
                    # ... then signal that we need to load the config.
                    dlgFilter.OnFileOpen(None)
                    # Fake that we asked the user for a filter name and got an OK
                    result = wx.ID_OK
                # If we're loading a Default profile, but there's none in the list, we can skip
                # the rest of the Filter method by pretending we got a Cancel from the user.
                else:
                    result = wx.ID_CANCEL
            # If we're not loading a Default profile ...
            else:
                # ... we need to show the Filter Dialog here.
                result = dlgFilter.ShowModal()
                
            # If the user clicks OK (or we have a Default config)
            if result == wx.ID_OK:
                # ... get the filter data ...
                if quoteFilter:
                    self.quoteFilterList = dlgFilter.GetQuotes()
                    self.showQuoteText = dlgFilter.GetShowQuoteText()
                    self.showQuoteNotes = dlgFilter.GetShowQuoteNotes()
##                if snapshotFilter:
##                    self.snapshotFilterList = dlgFilter.GetSnapshots()
##                    self.showSnapshotImage = dlgFilter.GetShowSnapshotImage()
##                    self.showSnapshotCoding = dlgFilter.GetShowSnapshotCoding()
##                    self.showSnapshotNotes = dlgFilter.GetShowSnapshotNotes()
                if keywordFilter:
                    self.keywordFilterList = dlgFilter.GetKeywords()
                    self.showKeywords = dlgFilter.GetShowKeywords()
                if quoteFilter or snapshotFilter:
                    self.showFile = dlgFilter.GetShowFile()
                    self.showTime = dlgFilter.GetShowTime()
                    self.showSourceInfo = dlgFilter.GetShowSourceInfo()
                self.showHyperlink = dlgFilter.GetShowHyperlink()
                self.showComments = dlgFilter.GetShowComments()
                # Remember the configuration name for later reuse
                self.configName = dlgFilter.configName
                # ... and signal the TextReport that the filter is to be applied.
                return True
            # If the filter is cancelled by the user ...
            else:
                # ... signal the TextReport that the filter is NOT to be applied.
                return False
            
        # If an Episode Name is passed in ...
        elif (self.episodeName != None):
            # Load the Episode to get the Episode Number
            tempEpisode = Episode.Episode(episode=self.episodeName, series=self.seriesName)
            # See if there are Clips in the Filter List
            clipFilter = (len(self.filterList) > 0)
            # See if there are Snapshots in the Snapshot Filter List
            snapshotFilter = (len(self.snapshotFilterList) > 0)
            # See if there are Keywords in the Filter List
            keywordFilter = (len(self.keywordFilterList) > 0)
            # Define the Filter Dialog.  We need reportType 11 to identify the Episode Report, and we
            # need only the Clip Filter and Keyword Filter for this report.  We want to show the file name,
            # clip time data, Clip Transcripts, Clip Keywords, Comments and clip notes options.
            dlgFilter = FilterDialog.FilterDialog(self.report,
                                                  -1,
                                                  self.title,
                                                  reportType=11,
                                                  reportScope=tempEpisode.number,
                                                  loadDefault=loadDefault,
                                                  configName=self.configName,
                                                  clipFilter=clipFilter,
                                                  snapshotFilter=snapshotFilter,
                                                  keywordFilter=keywordFilter,
                                                  reportContents=True,
                                                  showHyperlink=self.showHyperlink,
                                                  showFile=self.showFile,
                                                  showTime=self.showTime,
                                                  showSourceInfo=self.showSourceInfo,
                                                  showClipTranscripts=self.showTranscripts,
                                                  showSnapshotImage=self.showSnapshotImage,
                                                  showSnapshotCoding=self.showSnapshotCoding,
                                                  showKeywords=self.showKeywords,
                                                  showComments=self.showComments,
                                                  showClipNotes=self.showClipNotes,
                                                  showSnapshotNotes=self.showSnapshotNotes )
            # If there are Clips ...
            if clipFilter:
                # Populate the Filter Dialog with Clips
                dlgFilter.SetClips(self.filterList)
            # if there are Snapshots ...
            if snapshotFilter:
                # ... populate the Filter Dialog with Snapshots
                dlgFilter.SetSnapshots(self.snapshotFilterList)
            if keywordFilter:
                # Populate the Filter Dialog with Keywords
                dlgFilter.SetKeywords(self.keywordFilterList)
            # If we're loading the Default configuration ...
            if loadDefault:
                # ... get the list of existing configuration names.
                profileList = dlgFilter.GetConfigNames()
                # If (translated) "Default" is in the list ...
                # (NOTE that the default config name is stored in English, but gets translated by GetConfigNames!)
                if unicode(_('Default'), 'utf8') in profileList:
                    # ... then signal that we need to load the config.
                    dlgFilter.OnFileOpen(None)
                    # Fake that we asked the user for a filter name and got an OK
                    result = wx.ID_OK
                # If we're loading a Default profile, but there's none in the list, we can skip
                # the rest of the Filter method by pretending we got a Cancel from the user.
                else:
                    result = wx.ID_CANCEL
            # If we're not loading a Default profile ...
            else:
                # ... we need to show the Filter Dialog here.
                result = dlgFilter.ShowModal()
                
            # If the user clicks OK (or we have a Default config)
            if result == wx.ID_OK:
                # ... get the filter data ...
                if clipFilter:
                    self.filterList = dlgFilter.GetClips()
                    self.showTranscripts = dlgFilter.GetShowClipTranscripts()
                    self.showClipNotes = dlgFilter.GetShowClipNotes()
                if snapshotFilter:
                    self.snapshotFilterList = dlgFilter.GetSnapshots()
                    self.showSnapshotImage = dlgFilter.GetShowSnapshotImage()
                    self.showSnapshotCoding = dlgFilter.GetShowSnapshotCoding()
                    self.showSnapshotNotes = dlgFilter.GetShowSnapshotNotes()
                if keywordFilter:
                    self.keywordFilterList = dlgFilter.GetKeywords()
                    self.showKeywords = dlgFilter.GetShowKeywords()
                if clipFilter or snapshotFilter:
                    self.showFile = dlgFilter.GetShowFile()
                    self.showTime = dlgFilter.GetShowTime()
                    self.showSourceInfo = dlgFilter.GetShowSourceInfo()
                self.showHyperlink = dlgFilter.GetShowHyperlink()
                self.showComments = dlgFilter.GetShowComments()
                # Remember the configuration name for later reuse
                self.configName = dlgFilter.configName
                # ... and signal the TextReport that the filter is to be applied.
                return True
            # If the filter is cancelled by the user ...
            else:
                # ... signal the TextReport that the filter is NOT to be applied.
                return False
            
        # If a Library Name is passed in ...            
        elif (self.seriesName != None) or ((self.searchSeries != None) and (self.treeCtrl != None)):
            # Load the Library to get the Library Number
            tempLibrary = Library.Library(self.seriesName)
            # See if there are Episodes in the Filter List
            episodeFilter = (len(self.filterList) > 0)
            # See if there are Documents in the Filter List
            documentFilter = (len(self.documentFilterList) > 0)
            # See if there are Keywords in the Filter List
            keywordFilter = (len(self.keywordFilterList) > 0)
            # Define the Filter Dialog.  We need reportType 10 to identify the Library Report and we
            # need only the Episode Filter and the Keyord Filter for this report.
            dlgFilter = FilterDialog.FilterDialog(self.report,
                                                  -1,
                                                  self.title,
                                                  reportType=10,
                                                  reportScope=tempLibrary.number,
                                                  loadDefault=loadDefault,
                                                  configName=self.configName,
                                                  episodeFilter=episodeFilter,
                                                  documentFilter=documentFilter,
                                                  keywordFilter=keywordFilter,
                                                  reportContents=True,
                                                  showFile=self.showFile,
                                                  showTime=self.showTime,
                                                  showDocImportDate=self.showDocImportDate,
                                                  showKeywords=self.showKeywords)
            # Populate the Filter Dialog with the Episode, Document, and Keyword Filter lists
            if episodeFilter:
                dlgFilter.SetEpisodes(self.filterList)
            if documentFilter:
                dlgFilter.SetDocuments(self.documentFilterList)
            if keywordFilter:
                dlgFilter.SetKeywords(self.keywordFilterList)
            # If we're loading the Default configuration ...
            if loadDefault:
                # ... get the list of existing configuration names.
                profileList = dlgFilter.GetConfigNames()
                # If (translated) "Default" is in the list ...
                # (NOTE that the default config name is stored in English, but gets translated by GetConfigNames!)
                if unicode(_('Default'), 'utf8') in profileList:
                    # ... then signal that we need to load the config.
                    dlgFilter.OnFileOpen(None)
                    # Fake that we asked the user for a filter name and got an OK
                    result = wx.ID_OK
                # If we're loading a Default profile, but there's none in the list, we can skip
                # the rest of the Filter method by pretending we got a Cancel from the user.
                else:
                    result = wx.ID_CANCEL
            # If we're not loading a Default profile ...
            else:
                # ... we need to show the Filter Dialog here.
                result = dlgFilter.ShowModal()
                
            # If the user clicks OK (or we have a Default config)
            if result == wx.ID_OK:
                # ... get the filter data ...
                if episodeFilter:
                    self.filterList = dlgFilter.GetEpisodes()
                    self.showTime = dlgFilter.GetShowTime()
                if documentFilter:
                    self.documentFilterList = dlgFilter.GetDocuments()
                    self.showDocImportDate = dlgFilter.GetShowDocImportDate()
                if episodeFilter or documentFilter:
                    self.showFile = dlgFilter.GetShowFile()
                if keywordFilter:
                    self.keywordFilterList = dlgFilter.GetKeywords()
                    self.showKeywords = dlgFilter.GetShowKeywords()
                # Remember the configuration name for later reuse
                self.configName = dlgFilter.configName
                # ... and signal the TextReport that the filter is to be applied.
                return True
            # If the filter is cancelled by the user ...
            else:
                # ... signal the TextReport that the filter is NOT to be applied.
                return False
