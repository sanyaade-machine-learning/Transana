# Copyright (C) 2003 - 2010 The Board of Regents of the University of Wisconsin System 
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

# import the Python String module
import string
# import wxPython
import wx
# Import Transana's Clip object
import Clip
# import Transana's Collection Object
import Collection
# Import Transana's Database Interface
import DBInterface
# Import Transana's Dialog Boxes
import Dialogs
# Import Transana's Episode object
import Episode
# import Transana's Filter Dialog
import FilterDialog
# Import Transana's Keyword object
import Keyword
# Import Transana's Miscellaneous functions
import Misc
# Import Transana's Note object
import Note
# Import Transana's Series Object
import Series
# Import Transana's Text Report infrastructure
import TextReport
# import Transana's Exceptions
import TransanaExceptions
# import Transana's Global variables
import TransanaGlobal
# import Transana's Transcript Object
import Transcript


class ReportGenerator(wx.Object):
    """ This class creates and displays the Object Reports, formerly the Keyword Usage Report and the Collection Summary Report """
    def __init__(self, **kwargs):
        """ Create the Object Report
              If a seriesName is passed, all Episodes in that Series and their Episode keywords should be listed.
              If an episodeName is passed, all Clips from that Episode, regardless of Collection, and their Clip keywords should be listed.
              If a collection is passed, all Clips in that Collection, regardless of source Episode, and their Clip keywords should be listed.
              If a searchSeries is passed, use the treeCtrl to determine the Episodes that should be included.
              if a searchCollection is passed, use the treeCtrl to determine the Clips that should be included. """
        # Parameters can include:
        # title=''
        # seriesName=None,
        # episodeName=None,
        # collection=None,
        # searchSeries=None,
        # searchColl=None,
        # treeCtrl=None,
        # showNested=False,
        # showFile=True,
        # showTime=True,
        # showTranscripts=False,
        # showKeywords=False,
        # showComments=False,
        # showCollectionNotes=False
        # showClipNotes=False

        # Remember the parameters passed in and set values for all variables, even those NOT passed in.
        # Specify the Report Title
        if kwargs.has_key('title'):
            self.title = kwargs['title']
        else:
            self.title = ''
        if kwargs.has_key('seriesName'):
            self.seriesName = kwargs['seriesName']
        else:
            self.seriesName = None
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
        if kwargs.has_key('showFile') and kwargs['showFile']:
            self.showFile = True
        else:
            self.showFile = False
        if kwargs.has_key('showTime') and kwargs['showTime']:
            self.showTime = True
        else:
            self.showTime = False
        if kwargs.has_key('showTranscripts') and kwargs['showTranscripts']:
            self.showTranscripts = True
        else:
            self.showTranscripts = False
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
        if kwargs.has_key('showClipNotes') and kwargs['showClipNotes']:
            self.showClipNotes = True
        else:
            self.showClipNotes = False

        # Filter Configuration Name -- initialize to nothing
        self.configName = ''

        # Create the TextReport object, which forms the basis for text-based reports.
        self.report = TextReport.TextReport(None, title=self.title, displayMethod=self.OnDisplay,
                                            filterMethod=self.OnFilter, helpContext="Keyword Usage Report")

        # Define the Filter List (which will differ depending on the report type)
        self.filterList = []
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
        # We need variables to count the number of clips displayed and to accumulate their total time.
        self.clipCount = 0
        self.clipTotalTime = 0.0
        # Determine if we need to populate the Filter Lists.  If it hasn't already been done, we should do it.
        # If it has already been done, no need to do it again.
        if self.filterList == []:
            populateFilterList = True
        else:
            populateFilterList = False

        # Make the control writable
        reportText.SetReadOnly(False)
        # Set the font for the Report Title
        reportText.SetFont('Courier New', 13, 0x000000, 0xFFFFFF)
        # Make the font Bold
        reportText.SetBold(True)
        # Get the style specified associated with this font
        style = reportText.GetStyleAccessor("size:13,face:Courier New,fore:#000000,back:#ffffff,bold")
        # Get spaces appropriate to centering the title
        centerSpacer = self.report.GetCenterSpacer(style, self.title)
        # Insert the spaces to center the title
        reportText.InsertStyledText(centerSpacer)
        # Turn on underlining now (because we don't want the spaces to be underlined)
        reportText.SetUnderline(True)
        # Add the Report Title
        reportText.InsertStyledText(self.title)
        # Turn off underlining and bold
        reportText.SetUnderline(False)
        reportText.SetBold(False)

        # If a Collection is passed in ...
        if self.collection != None:
            # The major label for objects for the Collection report is Clip
            majorLabel = _('Clip:')
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
                # ... use the Clips in the Collection to start the majorList.
                majorList = DBInterface.list_of_clips_by_collection(self.collection.id, self.collection.parent)
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
                    # now get the clips for the new collection and add them to the Major (clip) list
                    majorList += DBInterface.list_of_clips_by_collection(collName, parentCollNum)
                    # Then get the nested collections under the new collection and add them to the Nested Collection list
                    # They get added at the FRONT of the list so that the report will mirror the organization of the
                    # database Tree.
                    nestedCollections = DBInterface.list_of_collections(collNum) + nestedCollections
            
            # Put all the Keywords for the Clips in the majorList in the minorList.
            # Start by iterating through the Major List
            for (clipNo, clipName, collNo) in majorList:
                # Create a Minor List dictionary entry, indexed to clip number, for the keywords.
                minorList[clipNo] = DBInterface.list_of_keywords(Clip = clipNo)
                # If we're populating Filter Lists ...
                if populateFilterList:
                    # ... then add the CLIP data to the (Clip) Filter List, initially checked ...
                    self.filterList.append((clipName, collNo, True))
                    # ... and iterate through that clip's keywords ...
                    for (kwg, kw, ex) in minorList[clipNo]:
                        # ... check to see if the entry is NOT already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... and add the keyword entry to the Keyword Filter List if it's not already there.
                            self.keywordFilterList.append((kwg, kw, True))

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
            # ... and use the Clips from the Episode for the majorList. 
            majorList = DBInterface.list_of_clips_by_episode(epObj.number)
            # Put all the Keywords for the Clips in the majorList in the minorList.  Iterate through the Major List ...
            for clipRecord in majorList:
                # ... load the Clip ...
                clipObj = Clip.Clip(clipRecord['ClipID'], clipRecord['CollectID'], clipRecord['ParentCollectNum'])
                # ... Get the Clip's Keywords and put them in the Minor List keyed to the Clip data
                minorList[(clipRecord['ClipID'], clipRecord['CollectID'], clipRecord['ParentCollectNum'])] = DBInterface.list_of_keywords(Clip = clipObj.number)
                # If we're populating the Filter Lists ...
                if populateFilterList:
                    # ... Add the Clip data to the main Filter List ...
                    self.filterList.append((clipRecord['ClipID'], clipRecord['CollectNum'], True))
                    # ... Iterate through the keywords that were just added to the Minor List (only for this Key) ...
                    for (kwg, kw, ex) in minorList[(clipRecord['ClipID'], clipRecord['CollectID'], clipRecord['ParentCollectNum'])]:
                        # .. and IF they're not already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... add them to the Keyword Filter List
                            self.keywordFilterList.append((kwg, kw, True))

        # If a Series Name is passed in ...            
        elif self.seriesName != None:
            # ...  add a subtitle
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Series: %s"), 'utf8')
            else:
                prompt = _("Series: %s")
            self.subtitle = prompt % self.seriesName
            # The label for our Major unit should reflect that these are Episodes
            majorLabel = _('Episode:')
            # Get the Episodes from the Series for the majorList
            majorList = DBInterface.list_of_episodes_for_series(self.seriesName)
            # Iterate through the Episodes in the Major List ...
            for (EpNo, epName, epParentNo) in majorList:
                # Put all the Keywords for the Episodes in the majorList in the minorList
                minorList[EpNo] = DBInterface.list_of_keywords(Episode = EpNo)
                # If we're populating the Filter Lists ...
                if populateFilterList:
                    # ... Add the Episode data to the main Filter List ...
                    self.filterList.append((epName, self.seriesName, True))
                    # ... Iterate through the keywords that were just added to the Minor List (only for this Key) ...
                    for (kwg, kw, ex) in minorList[EpNo]:
                        # .. and IF they're not already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... add them to the Keyword Filter List
                            self.keywordFilterList.append((kwg, kw, True))

        # If this report is called for a SearchSeriesResult, we build the majorList based on the contents of the Tree Control.
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
                prompt = unicode(_("Search Result: %s  Series: %s"), 'utf8')
            else:
                prompt = _("Search Result: %s  Series: %s")
            self.subtitle = prompt % (self.treeCtrl.GetItemText(searchResultNode), self.treeCtrl.GetItemText(self.searchSeries))
            # The majorLabel is for Episodes in this case
            majorLabel = _('Episode:')
            # Initialize the majorList to an empty list
            majorList = []
            # Get the first Child node from the searchColl collection
            (item, cookie) = self.treeCtrl.GetFirstChild(self.searchSeries)
            # Process all children in the searchSeries Series.  (IsOk() fails when all children are processed.)
            while item.IsOk():
                # Get the child item's Name
                itemText = self.treeCtrl.GetItemText(item)
                # Get the child item's Node Data
                itemData = self.treeCtrl.GetPyData(item)
                # See if the item is an Episode
                if itemData.nodetype == 'SearchEpisodeNode':
                    # If it's an Episode, add the Episode's Node Data to the majorList
                    majorList.append((itemData.recNum, itemText, itemData.parent))
                    # If we're populating the Filter Lists ...
                    if populateFilterList:
                        # ... Add the Episode data to the main Filter List ...
                        self.filterList.append((itemText, self.treeCtrl.GetItemText(self.treeCtrl.GetItemParent(item)), True))
                # Get the next Child Item and continue the loop
                (item, cookie) = self.treeCtrl.GetNextChild(self.searchSeries, cookie)
            # Once we have the Episodes in the majorList, we can gather their keywords into the minorList.
            # Start by iterating through the Major List
            for (EpNo, epName, epParentNo) in majorList:
                # Get all the keywords for the indicated Episode and add them to the Minor List, keyed to the Episode Name.
                minorList[EpNo] = DBInterface.list_of_keywords(Episode = EpNo)
                # If we're populating the Filter Lists ...
                if populateFilterList:
                    # ... Iterate through the keywords that were just added to the Minor List (only for this Key) ...
                    for (kwg, kw, ex) in minorList[EpNo]:
                        # .. and IF they're not already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... add them to the Keyword Filter List
                            self.keywordFilterList.append((kwg, kw, True))

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
            majorLabel = _('Clip:')
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
                if itemData.nodetype == 'SearchClipNode':
                    # If it's a Clip, add the Clip's Node Data to the majorList
                    majorList.append((itemData.recNum, itemText, itemData.parent))
                    # If we're populating the Filter List ...
                    if populateFilterList:
                        # ... add it to the Filter List as well.
                        self.filterList.append((itemText, itemData.parent, True))
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

            # Once we have the Clips in the majorList, we can gather their keywords into the minorList.
            # Iterate through the Major List ...
            for (clipNo, clipName, collNo) in majorList:
                # ... and get the keywords for each clip, adding them to the Minor List.
                minorList[clipNo] = DBInterface.list_of_keywords(Clip = clipNo)
                # If we're populating Filter Lists ...
                if populateFilterList:
                    # ... interate through the Minor List for this Key ...
                    for (kwg, kw, ex) in minorList[clipNo]:
                        # .. and IF they're not already in the list ...
                        if (kwg, kw, True) not in self.keywordFilterList:
                            # ... add them to the Keyword Filter List
                            self.keywordFilterList.append((kwg, kw, True))

        # If a subtitle is defined ...
        if self.subtitle != '':
            # ... set the font for the subtitle ...
            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
            # ... get the style specifier for that font ...
            style = reportText.GetStyleAccessor("size:10,face:Courier New,fore:#000000,back:#ffffff")
            # ... get the spaces needed to center the subtitle ...
            centerSpacer = self.report.GetCenterSpacer(style, self.subtitle)
            # ... and insert the spacer and the subtitle.
            reportText.InsertStyledText('\n' + centerSpacer + self.subtitle)
        if self.configName != '':
            # ...  add a subtitle
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Filter Configuration: %s"), 'utf8')
            else:
                prompt = _("Filter Configuration: %s")
            self.configLine = prompt % self.configName
            # ... set the font for the subtitle ...
            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
            # ... get the style specifier for that font ...
            style = reportText.GetStyleAccessor("size:10,face:Courier New,fore:#000000,back:#ffffff")
            # ... get the spaces needed to center the subtitle ...
            centerSpacer = self.report.GetCenterSpacer(style, self.configLine)
            # ... and insert the spacer and the subtitle.
            reportText.InsertStyledText('\n' + centerSpacer + self.configLine)
        # Skip a couple of lines.
        reportText.InsertStyledText('\n\n')

        # Initialize the initial data structure that will be turned into the report
        self.data = []
        # Create a Dictionary Data Structure to accumulate Keyword Counts
        keywordCounts = {}
        # Create a Dictionary Data Structure to accumulate Keyword Times
        keywordTimes = {}

        # The majorList and minorList are constructed differently for the Episode version of the report,
        # and so the report must be built differently here too!
        if self.episodeName == None:
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
            # Iterate through the major list
            for (groupNo, group, parentCollNo) in majorList:
                # If a Collection Name is passed in ...
                if self.collection != None:
                    # ... then our Filter comparison is based on Clip data
                    filterVal = (group, parentCollNo, True)
                # If a Series Name is passed in ...            
                elif self.seriesName != None:
                    # ... then our Filter comparison is based on Episode data
                    filterVal = (group, self.seriesName, True)
                # If this report is called for a SearchSeriesResult ...
                elif (self.searchSeries != None) and (self.treeCtrl != None):
                    # ... then our Filter comparison is based on the search series from the TreeCtrl
                    filterVal = (group, self.treeCtrl.GetItemText(self.searchSeries), True)
                # If this report is called for a SearchCollectionResult ...
                elif (self.searchColl != None) and (self.treeCtrl != None):
                    # ... then our Filter comparison is based on Search Collection data
                    filterVal = (group, parentCollNo, True)

                # now that we have the filter comparison data, we see if it's actually in the Filter List.
                if filterVal in self.filterList:
                    # If we have Collection-based data ...
                    if (self.collection != None) or ((self.searchColl != None) and (self.treeCtrl != None)):
                        # ... load the collection the current clip is in
                        tempColl = Collection.Collection(parentCollNo)
                        # Check to see if we're showing Collection headers, if we're showing nested collections (since
                        # there's no point showing collection headers if there aren't different collections!), and
                        # see if the new collection is different from the collection of the last clip displayed.
                        if (workingCollection != '') and self.showNested and (workingCollection[0] != tempColl.GetNodeString()):
                            # If so, print Collection-specific information.
                            # Set the font for the title
                            reportText.SetFont('Courier New', 12, 0x000000, 0xFFFFFF)
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Print the header
                            reportText.InsertStyledText(_('Collection: '))
                            # Turn bold off.
                            reportText.SetBold(False)
                            # print the collection data
                            reportText.InsertStyledText('%s\n' % tempColl.GetNodeString())

                            # If we are supposed to show Comments ...
                            if self.showComments:
                                # ... if the collection has a comment ...
                                if tempColl.comment != u'':
                                    # Set the font for the comments
                                    reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                                    # Turn bold on.
                                    reportText.SetBold(True)
                                    # Add the header to the report
                                    reportText.InsertStyledText(_('Collection Comment:') + '\n')
                                    # Turn bold off.
                                    reportText.SetBold(False)
                                    # Add the content of the Collection Comment to the report
                                    reportText.InsertStyledText('%s\n' % tempColl.comment)

                            # If we're supposed to show Collection Notes ...
                            if self.showCollectionNotes:
                                # ... get a list of notes, including their object numbers
                                notesList = DBInterface.list_of_notes(Collection=tempColl.number, includeNumber=True)
                                # If there are notes for this Clip ...
                                if (len(notesList) > 0):
                                    # Set the font for the Notes
                                    reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                                    # Turn bold on.
                                    reportText.SetBold(True)
                                    # Add the header to the report
                                    reportText.InsertStyledText(_('Collection Notes:') + '\n')
                                    # Turn bold off.
                                    reportText.SetBold(False)
                                    # Iterate throught the list of notes ...
                                    for note in notesList:
                                        # ... load each note ...
                                        tempNote = Note.Note(note[0])
                                        # Turn bold on.
                                        reportText.SetBold(True)
                                        # Add the note ID to the report
                                        reportText.InsertStyledText('%s\n' % tempNote.id)
                                        # Turn bold off.
                                        reportText.SetBold(False)
                                        # Add the note text to the report
                                        reportText.InsertStyledText('%s\n' % tempNote.text)
                            # Update the workingCollection variable with the data for the current collection so we'll
                            # be able to tell when the collection changes
                            workingCollection = (tempColl.GetNodeString(), tempColl.number)
                            # Add a blank line to the report
                            reportText.InsertStyledText('\n')
                            
                    # First, set the font for the heading.
                    reportText.SetFont('Courier New', 12, 0x000000, 0xFFFFFF)
                    # ... and make it bold.
                    reportText.SetBold(True)
                    # Add the group name to the report
                    reportText.InsertStyledText('%s %s\n' % (majorLabel, group))
                    # Turn bold off
                    reportText.SetBold(False)
                    # Set the font for the Clip Data
                    reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                    # If we have Collection-based data, we add some Clip-specific data
                    if (self.collection != None) or ((self.searchColl != None) and (self.treeCtrl != None)):
                        # Turn bold on.
                        reportText.SetBold(True)
                        # Add the header to the report
                        reportText.InsertStyledText(_('Collection:'))
                        # Turn bold off.
                        reportText.SetBold(False)
                        # Add the data to the report, the full Collection path in this case
                        reportText.InsertStyledText('  %s\n' % (tempColl.GetNodeString(),))
                        # Get the full Clip data
                        clipObj = Clip.Clip(groupNo)
                        
                        # If we're supposed to show the Media File Name ...
                        if self.showFile:
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('File:'))
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Add the data to the report, the file name in this case
                            reportText.InsertStyledText(_('  %s\n') % clipObj.media_filename)
                            # Add Additional Media File info
                            for mediaFile in clipObj.additional_media_files:
                                reportText.InsertStyledText(_('       %s\n') % mediaFile['filename'])
                            
                        # If we're supposed to show the Clip Time data ...
                        if self.showTime:
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('Time:'))
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Add the data to the report, the Clip start and stop times in this case
                            reportText.InsertStyledText('  %s - %s   (' % (Misc.time_in_ms_to_str(clipObj.clip_start), Misc.time_in_ms_to_str(clipObj.clip_stop)))
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('Length:'))
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Add the data to the report, the Clip Length in this case
                            reportText.InsertStyledText('  %s)\n' % (Misc.time_in_ms_to_str(clipObj.clip_stop - clipObj.clip_start)))
                            
                        # Increment the Clip Counter
                        self.clipCount += 1
                        # Add the Clip's length to the Clip Total Time accumulator
                        self.clipTotalTime += clipObj.clip_stop - clipObj.clip_start
                        
                        # If we are supposed to show Clip Transcripts, and the clip HAS a transcript ...
                        if self.showTranscripts:
                            # Skip a line
                            reportText.InsertStyledText('\n')
                            # Iterate through the clips transcripts
                            for tr in clipObj.transcripts:
                                # Default the Episode Transcript to None in case the load fails
                                episodeTranscriptObj = None
                                # Begin exception handling
                                try:
                                    # If the Clip Object has a defined Source Transcript ...
                                    if tr.source_transcript > 0:
                                        # ... try to load that source transcript
                                        episodeTranscriptObj = Transcript.Transcript(tr.source_transcript)
                                # if the record is not found (orphaned Clip)
                                except TransanaExceptions.RecordNotFoundError:
                                    # We don't need to do anything.
                                    pass

                                # Set the font for the clip transcript headers
                                reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                                
                                # If an Episode Transcript was found ...
                                if episodeTranscriptObj != None:
                                    # Turn bold on.
                                    reportText.SetBold(True)
                                    # Add the header to the report
                                    reportText.InsertStyledText(_('Episode Transcript:'))
                                    # Turn bold off.
                                    reportText.SetBold(False)
                                    # Add the data to the report, the Episode Transcript ID in this case
                                    reportText.InsertStyledText('  %s\n' % (episodeTranscriptObj.id,))
                                # if no Episode Transcript is found, we have an orphan.
                                else:
                                    # Turn bold on.
                                    reportText.SetBold(True)
                                    # Add the header to the report
                                    reportText.InsertStyledText(_('Episode Transcript:'))
                                    # Turn bold off.
                                    reportText.SetBold(False)
                                    # Add the data to the report, the lack of an Episode Transcript in this case
                                    reportText.InsertStyledText('  %s\n' % _('The Episode Transcript has been deleted.'))

                                # Turn bold on.
                                reportText.SetBold(True)
                                # Add the header to the report
                                reportText.InsertStyledText(_('Clip Transcript:') + '\n')

                                # Turn bold off.
                                reportText.SetBold(False)
                                # Add the Transcript to the report
                                reportText.InsertRTFText(tr.text)
                                reportText.InsertStyledText('\n')
                                
                    # If we have a Series Report ...
                    else:
                        # Get the full Episode data
                        episodeObj = Episode.Episode(groupNo)
                        # If we're supposed to show the Media File Name ...
                        if self.showFile:
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('File:'))
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Add the data to the report, the file name in this case
                            reportText.InsertStyledText(_('  %s\n') % episodeObj.media_filename)
                            # Add Additional Media File info
                            for mediaFile in episodeObj.additional_media_files:
                                reportText.InsertStyledText(_('       %s\n') % mediaFile['filename'])
                        # If we're supposed to show the Episode Time data ...
                        if self.showTime and (episodeObj.tape_length > 0):
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('Length:'))
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Add the data to the report, the Clip Length in this case
                            reportText.InsertStyledText('  %s\n' % (Misc.time_in_ms_to_str(episodeObj.tape_length)))
                            # Add the Episode's length to the Episode Total Time accumulator (Yeah, we're using clipTotalTime)
                            self.clipTotalTime += episodeObj.tape_length
                        # Increment the Episode Counter (Yeah, we're using clipCount)
                        self.clipCount += 1
                            
                                
                    # if we are supposed to show Keywords ...
                    if self.showKeywords:
                        # Set the font for the Keywords
                        reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                        # Skip a line
                        reportText.InsertStyledText('\n')
                        # If there are keywords in the list ... (even if they might all get filtered out ...)
                        if len(minorList[groupNo]) > 0:
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            if self.seriesName != None:
                                reportText.InsertStyledText(_('Episode Keywords:') + '\n')
                            else:
                                reportText.InsertStyledText(_('Clip Keywords:') + '\n')
                            # Turn bold off.
                            reportText.SetBold(False)
                        # Iterate through the list of Keywords for the group
                        for (keywordGroup, keyword, example) in minorList[groupNo]:
                            # See if the keyword should be included, based on the Keyword Filter List
                            if (keywordGroup, keyword, True) in self.keywordFilterList:
                                # Add the Keyword to the report
                                reportText.InsertStyledText('  %s : %s\n' % (keywordGroup, keyword))
                                # Add this Keyword to the Keyword Counts
                                if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                                    keywordCounts['%s : %s' % (keywordGroup, keyword)] += 1
                                    if (self.episodeName != None) or (self.collection != None) or (self.searchColl != None):
                                        keywordTimes['%s : %s' % (keywordGroup, keyword)] += clipObj.clip_stop - clipObj.clip_start
                                else:
                                    keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                                    if (self.episodeName != None) or (self.collection != None) or (self.searchColl != None):
                                        keywordTimes['%s : %s' % (keywordGroup, keyword)] = clipObj.clip_stop - clipObj.clip_start
                                    
                    # if we are supposed to show Comments ...
                    if self.showComments:
                        # If there are keywords in the list ... (even if they might all get filtered out ...)
                        if clipObj.comment != u'':
                            # Set the font for the Comment
                            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('Clip Comment:') + '\n')
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Add the content of the Clip Comment to the report
                            reportText.InsertStyledText('%s\n' % clipObj.comment)

                    # If we are supposed to show Clip Notes ...
                    if self.showClipNotes:
                        # ... get a list of notes, including their object numbers
                        notesList = DBInterface.list_of_notes(Clip=clipObj.number, includeNumber=True)
                        # If there are notes for this Clip ...
                        if len(notesList) > 0:
                            # Set the font for the notes
                            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('Clip Notes:') + '\n')
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Iterate throught the list of notes ...
                            for note in notesList:
                                # ... load each note ...
                                tempNote = Note.Note(note[0])
                                # Turn bold on.
                                reportText.SetBold(True)
                                # Add the note ID to the report
                                reportText.InsertStyledText('%s\n' % tempNote.id)
                                # Turn bold off.
                                reportText.SetBold(False)
                                # Add the note text to the report
                                reportText.InsertStyledText('%s\n' % tempNote.text)

                    # Add a blank line after each group
                    reportText.InsertStyledText('\n')
                            
        # If this IS an Episode-based report ...
        else:
            # Iterate through the major list
            for clipRecord in majorList:
                # our Filter comparison is based on Clip data
                filterVal = (clipRecord['ClipID'], clipRecord['CollectNum'], True)

                # now that we have the filter comparison data, we see if it's actually in the Filter List.
                if filterVal in self.filterList:
                    # First, load the collection the current clip is in
                    collectionObj = Collection.Collection(clipRecord['CollectNum'])
                    # Set the font for the heading.
                    reportText.SetFont('Courier New', 12, 0x000000, 0xFFFFFF)
                    # ... and make it bold.
                    reportText.SetBold(True)
                    # Add the header to the report
                    reportText.InsertStyledText(_("Clip:"))
                    # Turn bold off
                    reportText.SetBold(False)
                    # Add the data to the report, the Clip ID in this case
                    reportText.InsertStyledText("  %s\n" % (clipRecord['ClipID'],))
                    # Set the font for the rest of the record
                    reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                    # Turn bold on.
                    reportText.SetBold(True)
                    # Add the header to the report
                    reportText.InsertStyledText(_('Collection:'))
                    # Turn bold off.
                    reportText.SetBold(False)
                    # Add the data to the report, the full Collection path in this case
                    reportText.InsertStyledText('  %s\n' % (collectionObj.GetNodeString(),))
                    # Load the Clip Object
                    clipObj = Clip.Clip(clipRecord['ClipNum'])

                    # If we're supposed to show Media File name information ...
                    if self.showFile:
                        # Turn bold on.
                        reportText.SetBold(True)
                        # Add the header to the report
                        reportText.InsertStyledText(_('File:'))
                        # Turn bold off.
                        reportText.SetBold(False)
                        # Add the data to the report, the Clip media file name in this case
                        reportText.InsertStyledText('  %s\n' % clipObj.media_filename)
                        # Add Additional Media File info
                        for mediaFile in clipObj.additional_media_files:
                            reportText.InsertStyledText(_('       %s\n') % mediaFile['filename'])

                    # If we're supposed to show Clip Time data ...
                    if self.showTime:
                        # Turn bold on.
                        reportText.SetBold(True)
                        # Add the header to the report
                        reportText.InsertStyledText(_('Time:'))
                        # Turn bold off.
                        reportText.SetBold(False)
                        # Add the data to the report, the Clip start and stop times in this case
                        reportText.InsertStyledText('  %s - %s   (' % (Misc.time_in_ms_to_str(clipRecord['ClipStart']), Misc.time_in_ms_to_str(clipRecord['ClipStop'])))
                        # Turn bold on.
                        reportText.SetBold(True)
                        # Add the header to the report
                        reportText.InsertStyledText(_('Length:'))
                        # Turn bold off.
                        reportText.SetBold(False)
                        # Add the data to the report, the Clip length in this case
                        reportText.InsertStyledText('  %s)\n' % (Misc.time_in_ms_to_str(clipRecord['ClipStop'] - clipRecord['ClipStart']),))
                        
                    # Increment the Clip Counter
                    self.clipCount += 1
                    # Add the Clip's length to the Clip Total Time accumulator
                    self.clipTotalTime += clipRecord['ClipStop'] - clipRecord['ClipStart']
                    
                    # If we are supposed to show Clip Transcripts, and the clip HAS a transcript ...
                    if self.showTranscripts:
                        # print a blank line
                        reportText.InsertStyledText('\n')
                        # for each Clip transcript:
                        for tr in clipObj.transcripts:
                            # Default the Episode Transcript to None in case the load fails
                            episodeTranscriptObj = None
                            # Begin exception handling
                            try:
                                # If the Clip Object has a defined Source Transcript ...
                                if tr.source_transcript > 0:
                                    # ... try to load that source transcript
                                    episodeTranscriptObj = Transcript.Transcript(tr.source_transcript)
                            # if the record is not found (orphaned Clip)
                            except TransanaExceptions.RecordNotFoundError:
                                # We don't need to do anything.
                                pass

                            # Set the font for the Clip Transcript Header
                            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                            
                            # If an Episode Transcript was found ...
                            if episodeTranscriptObj != None:
                                # Turn bold on.
                                reportText.SetBold(True)
                                # Add the header to the report
                                reportText.InsertStyledText(_('Episode Transcript:'))
                                # Turn bold off.
                                reportText.SetBold(False)
                                # Add the data to the report, the Episode Transcript ID in this case
                                reportText.InsertStyledText('  %s\n' % (episodeTranscriptObj.id,))
                            # if no Episode Transcript is found, we have an orphan.
                            else:
                                # Turn bold on.
                                reportText.SetBold(True)
                                # Add the header to the report
                                reportText.InsertStyledText(_('Episode Transcript:'))
                                # Turn bold off.
                                reportText.SetBold(False)
                                # Add the data to the report, the lack of an Episode Transcript in this case
                                reportText.InsertStyledText('  %s\n' % _('The Episode Transcript has been deleted.'))

                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('Clip Transcript:') + '\n')
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Add the Transcript to the report
                            reportText.InsertRTFText(tr.text)                        
                            reportText.InsertStyledText('\n')
                        
                    # if we are supposed to show Keywords ...
                    if self.showKeywords:
                        # Set the font for the Keywords
                        reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                        # print a blank line
                        reportText.InsertStyledText('\n')
                        # If there are keywords in the list ... (even if they might all get filtered out ...)
                        if len(minorList[(clipRecord['ClipID'], clipRecord['CollectID'], clipRecord['ParentCollectNum'])]) > 0:
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('Clip Keywords:') + '\n')
                            # Turn bold off.
                            reportText.SetBold(False)
                        # Iterate through the list of Keywords for the group
                        for (keywordGroup, keyword, example) in minorList[(clipRecord['ClipID'], clipRecord['CollectID'], clipRecord['ParentCollectNum'])]:
                            # See if the keyword should be included, based on the Keyword Filter List
                            if (keywordGroup, keyword, True) in self.keywordFilterList:
                                # Add the Keyword to the report
                                reportText.InsertStyledText('  %s : %s\n' % (keywordGroup, keyword))
                                # Add this Keyword to the Keyword Counts
                                if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                                    keywordCounts['%s : %s' % (keywordGroup, keyword)] += 1
                                    keywordTimes['%s : %s' % (keywordGroup, keyword)] += clipObj.clip_stop - clipObj.clip_start
                                else:
                                    keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                                    keywordTimes['%s : %s' % (keywordGroup, keyword)] = clipObj.clip_stop - clipObj.clip_start
                                    
                    # if we are supposed to show Comments ...
                    if self.showComments:
                        # If there are keywords in the list ... (even if they might all get filtered out ...)
                        if clipRecord['Comment'] != u'':
                            # Set the font for the Comments
                            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('Clip Comment:') + '\n')
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Add the content of the Clip Comment to the report
                            reportText.InsertStyledText('%s\n' % clipRecord['Comment'])
                            
                    # If we are supposed to show Clip Notes ...
                    if self.showClipNotes:
                        # ... get a list of notes, including their object numbers
                        notesList = DBInterface.list_of_notes(Clip=clipObj.number, includeNumber=True)
                        # If there are notes for this Clip ...
                        if len(notesList) > 0:
                            # Set the font for the Notes
                            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                            # Turn bold on.
                            reportText.SetBold(True)
                            # Add the header to the report
                            reportText.InsertStyledText(_('Clip Notes:') + '\n')
                            # Turn bold off.
                            reportText.SetBold(False)
                            # Iterate throught the list of notes ...
                            for note in notesList:
                                # ... load each note ...
                                tempNote = Note.Note(note[0])
                                # Turn bold on.
                                reportText.SetBold(True)
                                # Add the note ID to the report
                                reportText.InsertStyledText('%s\n' % tempNote.id)
                                # Turn bold off.
                                reportText.SetBold(False)
                                # Add the note text to the report
                                reportText.InsertStyledText('%s\n' % tempNote.text)

                    # Add a blank line after each group
                    reportText.InsertStyledText('\n')

        # Now add the Report Summary
        # If there's data in the majorList ...
        if self.showKeywords and (len(majorList) != 0):
            # First, set the font for the summary
            reportText.SetFont('Courier New', 12, 0x000000, 0xFFFFFF)
            # Make the current font bold.
            reportText.SetBold(True)
            # Add the section heading to the report
            reportText.InsertStyledText(_("Summary\n"))
            # Turn bold off
            reportText.SetBold(False)
            # Set the font for the Keywords
            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
            # Get a list of the Keyword Group : Keyword pairs that have been used
            countKeys = keywordCounts.keys()
            # Sort the list
            countKeys.sort()
            # Add the sorted keywords to the summary with their counts
            for key in countKeys:
                # Right-pad the keyword with spaces so columns will line up right.
                st = key + '                                                            '[len(key):]
                if (self.episodeName != None) or (self.collection != None) or (self.searchColl != None):
                    # Add the text to the report.
                    reportText.InsertStyledText('  %60s  %5d  %10s\n' % (st[:60], keywordCounts[key], Misc.time_in_ms_to_str(keywordTimes[key])))
                else:
                    # Add the text to the report.
                    reportText.InsertStyledText('  %60s  %5d\n' % (st[:60], keywordCounts[key]))
        # If our Clip Counter shows the presence of Clips ...
        if self.clipCount > 0:
            # Set the font for the data
            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
            reportText.InsertStyledText('\n')
            # If we're showing Clip Time data ...
            if self.showTime:
                # Set the appropriate prompt for Episodes or Clips
                if self.seriesName != None:
                    prompt = _('  Episodes:  %8d                                      Total Time:  %s\n')
                else:
                    prompt = _('  Clips:  %8d                                         Total Time:  %s\n')
                # Add the Clip Count and Total Time data to the report
                reportText.InsertStyledText(prompt % (self.clipCount, Misc.time_in_ms_to_str(self.clipTotalTime)))
            # if we're not showing Clip Time data ...
            else:
                if self.seriesName != None:
                    prompt = _('  Episodes:  %4d\n')
                else:
                    prompt = _('  Clips:  %4d\n')
                # Add the Clip Count but NOT the Total Time data to the report
                reportText.InsertStyledText(prompt % self.clipCount)
            
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
            # Define the Filter Dialog.  We need reportType 12 to identify the Collection Report, and we
            # need only the Clip Filter and Keyword Filter for this report.  We want to show the file name,
            # clip time data, Clip Transcripts, Clip Keywords, Comments, Collection Note, Clip Note, and
            # Nested Data options.
            dlgFilter = FilterDialog.FilterDialog(self.report, -1, self.title, reportType=12,
                                                  reportScope=reportScope, loadDefault=loadDefault, configName=self.configName,
                                                  clipFilter=True, keywordFilter=True, reportContents=True,
                                                  showFile=self.showFile, showTime=self.showTime,
                                                  showClipTranscripts=self.showTranscripts, showClipKeywords=self.showKeywords,
                                                  showComments=self.showComments, showNestedData=self.showNested,
                                                  showCollectionNotes=self.showCollectionNotes, showClipNotes=self.showClipNotes)
            # Populate the Filter Dialog with the Clips and Keyword Filter lists
            dlgFilter.SetClips(self.filterList)
            dlgFilter.SetKeywords(self.keywordFilterList)
            # If we're loading the Default configuration ...
            if loadDefault:
                # ... get the list of existing configuration names.
                profileList = dlgFilter.GetConfigNames()
                # If (translated) "Default" is in the list ...
                # (NOTE that the default config name is stored in English, but gets translated by GetConfigNames!)
                if unicode(_('Default'), TransanaGlobal.encoding) in profileList:
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
                self.filterList = dlgFilter.GetClips()
                self.keywordFilterList = dlgFilter.GetKeywords()
                self.showFile = dlgFilter.GetShowFile()
                self.showTime = dlgFilter.GetShowTime()
                self.showTranscripts = dlgFilter.GetShowClipTranscripts()
                self.showKeywords = dlgFilter.GetShowClipKeywords()
                self.showComments = dlgFilter.GetShowComments()
                self.showCollectionNotes = dlgFilter.GetShowCollectionNotes()
                self.showClipNotes = dlgFilter.GetShowClipNotes()
                self.showNested = dlgFilter.GetShowNestedData()
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
            # Define the Filter Dialog.  We need reportType 11 to identify the Episode Report, and we
            # need only the Clip Filter and Keyword Filter for this report.  We want to show the file name,
            # clip time data, Clip Transcripts, Clip Keywords, Comments and clip notes options.
            dlgFilter = FilterDialog.FilterDialog(self.report, -1, self.title, reportType=11,
                                                  reportScope=tempEpisode.number, loadDefault=loadDefault, configName=self.configName,
                                                  clipFilter=True, keywordFilter=True, reportContents=True,
                                                  showFile=self.showFile, showTime=self.showTime,
                                                  showClipTranscripts=self.showTranscripts, showClipKeywords=self.showKeywords,
                                                  showComments=self.showComments, showClipNotes=self.showClipNotes)
            # Populate the Filter Dialog with the Clip and Keyword Filter lists
            dlgFilter.SetClips(self.filterList)
            dlgFilter.SetKeywords(self.keywordFilterList)
            # If we're loading the Default configuration ...
            if loadDefault:
                # ... get the list of existing configuration names.
                profileList = dlgFilter.GetConfigNames()
                # If (translated) "Default" is in the list ...
                # (NOTE that the default config name is stored in English, but gets translated by GetConfigNames!)
                if unicode(_('Default'), TransanaGlobal.encoding) in profileList:
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
                self.filterList = dlgFilter.GetClips()
                self.keywordFilterList = dlgFilter.GetKeywords()
                self.showFile = dlgFilter.GetShowFile()
                self.showTime = dlgFilter.GetShowTime()
                self.showTranscripts = dlgFilter.GetShowClipTranscripts()
                self.showKeywords = dlgFilter.GetShowClipKeywords()
                self.showComments = dlgFilter.GetShowComments()
                self.showClipNotes = dlgFilter.GetShowClipNotes()
                # Remember the configuration name for later reuse
                self.configName = dlgFilter.configName
                # ... and signal the TextReport that the filter is to be applied.
                return True
            # If the filter is cancelled by the user ...
            else:
                # ... signal the TextReport that the filter is NOT to be applied.
                return False
            
        # If a Series Name is passed in ...            
        elif (self.seriesName != None) or ((self.searchSeries != None) and (self.treeCtrl != None)):
            # Load the Series to get the Series Number
            tempSeries = Series.Series(self.seriesName)
            # Define the Filter Dialog.  We need reportType 10 to identify the Series Report and we
            # need only the Episode Filter and the Keyord Filter for this report.
            dlgFilter = FilterDialog.FilterDialog(self.report, -1, self.title, reportType=10,
                                                  reportScope=tempSeries.number, loadDefault=loadDefault, configName=self.configName,
                                                  episodeFilter=True, keywordFilter=True, reportContents=True, showFile=self.showFile,
                                                  showTime=self.showTime, showClipKeywords=self.showKeywords)
            # Populate the Filter Dialog with the Episode and Keyword Filter lists
            dlgFilter.SetEpisodes(self.filterList)
            dlgFilter.SetKeywords(self.keywordFilterList)
            # If we're loading the Default configuration ...
            if loadDefault:
                # ... get the list of existing configuration names.
                profileList = dlgFilter.GetConfigNames()
                # If (translated) "Default" is in the list ...
                # (NOTE that the default config name is stored in English, but gets translated by GetConfigNames!)
                if unicode(_('Default'), TransanaGlobal.encoding) in profileList:
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
                self.filterList = dlgFilter.GetEpisodes()
                self.keywordFilterList = dlgFilter.GetKeywords()
                self.showFile = dlgFilter.GetShowFile()
                self.showTime = dlgFilter.GetShowTime()
                self.showKeywords = dlgFilter.GetShowClipKeywords()
                # Remember the configuration name for later reuse
                self.configName = dlgFilter.configName
                # ... and signal the TextReport that the filter is to be applied.
                return True
            # If the filter is cancelled by the user ...
            else:
                # ... signal the TextReport that the filter is NOT to be applied.
                return False
