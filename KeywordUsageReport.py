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

"""This module implements the Keyword Usage Report. """

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

import string
import wx
import TransanaGlobal
import DBInterface
import Misc
import ReportPrintoutClass
import Episode
import Clip
import Keyword
import Dialogs

class KeywordUsageReport(wx.Object):
    """ This class creates and displays the Keyword Usage Report """
    def __init__(self, seriesName=None, episodeName=None, collection=None, searchSeries=None, searchColl=None, treeCtrl=None):
        # Specify the Report Title
        self.title = _("Keyword Usage Report")
        # Create minorList as a blank Dictionary Object
        minorList = {}
        # If a Collection Name is passed in ...
        if collection != None:
            # ...  add a subtitle and ...
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Collection: %s"), 'utf8')
            else:
                prompt = _("Collection: %s")
            self.subtitle = prompt % collection.id
            majorLabel = _('Clip:')
            # ... use the Clips in the Collection for the majorList.
            majorList = DBInterface.list_of_clips_by_collection(collection.id, collection.parent)
            # Put all the Keywords for the Clips in the majorList in the minorList
            for (clipNo, clipName, collNo) in majorList:
                minorList[clipName] = DBInterface.list_of_keywords(Clip = clipNo)

        # If an Episode Name is passed in ...
        elif episodeName != None:
            # ...  add a subtitle and ...
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Episode: %s"), 'utf8')
            else:
                prompt = _("Episode: %s")
            self.subtitle = prompt % episodeName
            # ... use the Clips from the Episode for the majorList
            epObj = Episode.Episode(series = seriesName, episode = episodeName)
            majorList = DBInterface.list_of_clips_by_episode(epObj.number)
            # Put all the Keywords for the Clips in the majorList in the minorList
            for clipRecord in majorList:
                clipObj = Clip.Clip(clipRecord['ClipID'], clipRecord['CollectID'], clipRecord['ParentCollectNum'])
                minorList[(clipRecord['ClipID'], clipRecord['CollectID'], clipRecord['ParentCollectNum'])] = DBInterface.list_of_keywords(Clip = clipObj.number)

        # If a Series Name is passed in ...            
        elif seriesName != None:
            # ...  add a subtitle and ...
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Series: %s"), 'utf8')
            else:
                prompt = _("Series: %s")
            self.subtitle = prompt % seriesName
            majorLabel = _('Episode:')
            # ... use the Episodes from the Series for the majorList
            majorList = DBInterface.list_of_episodes_for_series(seriesName)
            # Put all the Keywords for the Episodes in the majorList in the minorList
            for (EpNo, epName, epParentNo) in majorList:
                epObj = Episode.Episode(series = seriesName, episode = epName)
                minorList[epName] = DBInterface.list_of_keywords(Episode = epObj.number)

        # If this report is called for a SearchSeriesResult, we build the majorList based on the contents of the Tree Control.
        elif (searchSeries != None) and (treeCtrl != None):
            # Get the Search Result Name for the subtitle
            searchResultNode = searchSeries
            while 1:
                searchResultNode = treeCtrl.GetItemParent(searchResultNode)
                tempData = treeCtrl.GetPyData(searchResultNode)
                if tempData.nodetype == 'SearchResultsNode':
                    break
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Search Result: %s  Series: %s"), 'utf8')
            else:
                prompt = _("Search Result: %s  Series: %s")
            self.subtitle = prompt % (treeCtrl.GetItemText(searchResultNode), treeCtrl.GetItemText(searchSeries))
            # The majorLabel is for Episodes in this case
            majorLabel = _('Episode:')
            # Initialize the majorList to an empty list
            majorList = []
            # Get the first Child node from the searchColl collection
            (item, cookie) = treeCtrl.GetFirstChild(searchSeries)
            # Process all children in the searchSeries Series.  (IsOk() fails when all children are processed.)
            while item.IsOk():
                # Get the item's Name
                itemText = treeCtrl.GetItemText(item)
                # Get the item's Node Data
                itemData = treeCtrl.GetPyData(item)
                # See if the item is an Episode
                if itemData.nodetype == 'SearchEpisodeNode':
                    # If it's an Episode, add the Episode's Node Data to the majorList
                    majorList.append((itemData.recNum, itemText, itemData.parent))
                # Get the next Child Item and continue the loop
                (item, cookie) = treeCtrl.GetNextChild(searchSeries, cookie)
            # Once we have the Episodes in the majorList, we can gather their keywords into the minorList
            for (EpNo, epName, epParentNo) in majorList:
                epObj = Episode.Episode(series = treeCtrl.GetItemText(searchSeries), episode = epName)
                minorList[epName] = DBInterface.list_of_keywords(Episode = epObj.number)
        
        # If this report is called for a SearchCollectionResult, we build the majorList based on the contents of the Tree Control.
        elif (searchColl != None) and (treeCtrl != None):
            # Get the Search Result Name for the subtitle
            searchResultNode = searchColl
            while 1:
                searchResultNode = treeCtrl.GetItemParent(searchResultNode)
                tempData = treeCtrl.GetPyData(searchResultNode)
                if tempData.nodetype == 'SearchResultsNode':
                    break
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Search Result: %s  Collection: %s"), 'utf8')
            else:
                prompt = _("Search Result: %s  Collection: %s")
            self.subtitle = prompt % (treeCtrl.GetItemText(searchResultNode), treeCtrl.GetItemText(searchColl))
            # The majorLabel is for Clips in this case
            majorLabel = _('Clip:')
            # Initialize the majorList to an empty list
            majorList = []
            # Extracting data from the treeCtrl requires a "cookie" value, which is initialized to 0
            cookie = 0
            # Get the first Child node from the searchColl collection
            (item, cookie) = treeCtrl.GetFirstChild(searchColl)
            # Process all children in the searchColl collection
            while item.IsOk():
                # Get the item's Name
                itemText = treeCtrl.GetItemText(item)
                # Get the item's Node Data
                itemData = treeCtrl.GetPyData(item)
                # See if the item is a Clip
                if itemData.nodetype == 'SearchClipNode':
                    # If it's a Clip, add the Clip's Node Data to the majorList
                    majorList.append((itemData.recNum, itemText, itemData.parent))
                # When we get to the last Child Item, stop looping
                if item == treeCtrl.GetLastChild(searchColl):
                    break
                # If we're not at the Last Child Item, get the next Child Item and continue the loop
                else:
                    (item, cookie) = treeCtrl.GetNextChild(searchColl, cookie)
            # Once we have the Clips in the majorList, we can gather their keywords into the minorList
            for (clipNo, clipName, collNo) in majorList:
                minorList[clipName] = DBInterface.list_of_keywords(Clip = clipNo)
        
        
        # Initialize and fill the initial data structure that will be turned into the report
        self.data = []
        # Create a Dictionary Data Structure to accumulate Keyword Counts
        keywordCounts = {}

        # The majorList and minorList are constructed differently for the Episode version of the report,
        # and so the report must be built differently here too!
        if episodeName == None:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                majorLabel = unicode(majorLabel, 'utf8')
            # Iterate through the major list
            for (groupNo, group, parentCollNo) in majorList:
                # Use the group name as a Heading
                self.data.append((('Subheading', '%s %s' % (majorLabel, group)),))
                # Iterate through the list of Keywords for the group
                for (keywordGroup, keyword, example) in minorList[group]:
                    # Use the Keyword name as a Subheading
                    self.data.append((('Subtext', '%s : %s' % (keywordGroup, keyword)),))
                    # Add this Keyword to the Keyword Counts
                    if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                        keywordCounts['%s : %s' % (keywordGroup, keyword)] = keywordCounts['%s : %s' % (keywordGroup, keyword)] + 1
                    else:
                        keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                # Add a blank line after each group
                self.data.append((('Normal', ''),))
        else:
            # Iterate through the major list
            for clipRecord in majorList:
                # Use the group name as a Heading
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Collection: %s, Clip: %s'), 'utf8')
                else:
                    prompt = _('Collection: %s, Clip: %s')
                self.data.append((('Subheading', prompt % (clipRecord['CollectID'], clipRecord['ClipID'])), ('NormalRight', '(%s - %s)' % (Misc.time_in_ms_to_str(clipRecord['ClipStart']), Misc.time_in_ms_to_str(clipRecord['ClipStop'])))))  
                # Iterate through the list of Keywords for the group
                for (keywordGroup, keyword, example) in minorList[(clipRecord['ClipID'], clipRecord['CollectID'], clipRecord['ParentCollectNum'])]:
                    # Use the Keyword name as a Subheading
                    self.data.append((('Subtext', '%s : %s' % (keywordGroup, keyword)),))
                    # Add this Keyword to the Keyword Counts
                    if keywordCounts.has_key('%s : %s' % (keywordGroup, keyword)):
                        keywordCounts['%s : %s' % (keywordGroup, keyword)] = keywordCounts['%s : %s' % (keywordGroup, keyword)] + 1
                    else:
                        keywordCounts['%s : %s' % (keywordGroup, keyword)] = 1
                # Add a blank line after each group
                self.data.append((('Normal', ''),))

        # If there's data in the majorList ...
        if len(majorList) != 0:
            # Now add the Report Summary
            self.data.append((('Subheading', _('Summary')),))
            # Get a list of the Keyword Group : Keyword pairs that have been used
            countKeys = keywordCounts.keys()
            # Sort the list
            countKeys.sort()
            # Add the sorted keywords to the summary with their counts
            for key in countKeys:
                self.data.append((('Subtext', key), ('NormalRight', '%s' % keywordCounts[key])))
            # The initial data structure needs to be prepared.  What PrepareData() does is to create a graphic
            # object that is the correct size and dimensions for the type of paper selected, and to create
            # a datastructure that breaks the data sent in into separate pages, again based on the dimensions
            # of the paper currently selected.
            (self.graphic, self.pageData) = ReportPrintoutClass.PrepareData(TransanaGlobal.printData, self.title, self.data, self.subtitle)

            # Send the results of the PrepareData() call to the MyPrintout object, once for the print preview
            # version and once for the printer version.  
            printout = ReportPrintoutClass.MyPrintout(self.title, self.graphic, self.pageData, self.subtitle)
            printout2 = ReportPrintoutClass.MyPrintout(self.title, self.graphic, self.pageData, self.subtitle)

            # Create the Print Preview Object
            self.preview = wx.PrintPreview(printout, printout2, TransanaGlobal.printData)
            # Check for errors during Print preview construction
            if not self.preview.Ok():
                dlg = Dialogs.ErrorDialog(None, _("Print Preview Problem"))
                dlg.ShowModal()
                dlg.Destroy()
                return
            # Create the Frame for the Print Preview
            theWidth = max(wx.ClientDisplayRect()[2] - 180, 760)
            theHeight = max(wx.ClientDisplayRect()[3] - 200, 560)
            frame = wx.PreviewFrame(self.preview, None, _("Print Preview"), size=(theWidth, theHeight))
            frame.Centre()
            
            # Initialize the Frame for the Print Preview
            frame.Initialize()
            # Display the Print Preview Frame
            frame.Show(True)
        # If there's NO data in the majorList ...
        else:
            # If there are no clips to report, display an error message.
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('%s has no data for the Keyword Usage Report.'), 'utf8')
            else:
                prompt = _('%s has no data for the Keyword Usage Report.')
            dlg = wx.MessageDialog(None, prompt % self.subtitle, style = wx.OK | wx.ICON_EXCLAMATION)
            dlg.ShowModal()
            dlg.Destroy()
            
