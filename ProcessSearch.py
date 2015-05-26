# Copyright (C) 2003-2014 The Board of Regents of the University of Wisconsin System 
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

""" This module handles Search Requests and all related processing. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Import wxPython
import wx

# Import the Transana Series Object
import Series
# Import the Transana Episode Object
import Episode
# Import the Transana Collection Object
import Collection
# Import the Transana Database Interface
import DBInterface
# Import the Transana Search Dialog Box
import SearchDialog
# import Transana's Constants
import TransanaConstants
# Import Transana's Globals
import TransanaGlobal

# Import the Python String module
import string


class ProcessSearch(object):
    """ This class handles all processing related to Searching. """
    # searchName and searchTerms are used by unit_test_search
    def __init__(self, dbTree, searchCount, kwg=None, kw=None, searchName=None, searchTerms=None):
        """ Initialize the ProcessSearch class.  The dbTree parameter accepts a wxTreeCtrl as the Database Tree where
            Search Results should be displayed.  The searchCount parameter accepts the number that should be included
            in the Default Search Title. Optional kwg (Keyword Group) and kw (Keyword) parameters implement Quick Search
            for the keyword specified. """

        # Note the Database Tree that accepts Search Results
        self.dbTree = dbTree
        self.collectionList = []
        # If kwg and kw are None, we are doing a regular (full) search.
        if ((kwg == None) or (kw == None)) and (searchTerms == None):
            # Create the Search Dialog Box
            dlg = SearchDialog.SearchDialog(_("Search") + " %s" % searchCount)
            # Display the Search Dialog Box and record the Result
            result = dlg.ShowModal()
            # If the user selects OK ...
            if result == wx.ID_OK:
                # ... get the search name from the dialog
                searchName = dlg.searchName.GetValue().strip()
                # Search Name is required.  If it was eliminated, put it back!
                if searchName == '':
                    searchName = _("Search") + " %s" % searchCount

                # Get the Collections Tree from the Search Form
                collTree = dlg.ctcCollections
                # Get the Collections Tree's Root Node
                collNode = collTree.GetRootItem()
                # Get a list of all the Checked Collections in the Collections Tree
                self.collectionList = dlg.GetCollectionList(collTree, collNode, True)
                # ... and get the search terms from the dialog
                searchTerms = dlg.searchQuery.GetValue().split('\n')
                # Get the includeEpisodes info
                includeEpisodes = dlg.includeEpisodes.IsChecked()
                # Get the includeClips info
                includeClips = dlg.includeClips.IsChecked()
                # Get the includeSnapshots info
                includeSnapshots = dlg.includeSnapshots.IsChecked()
            # Destroy the Search Dialog Box
            dlg.Destroy()
        elif (searchTerms != None):
            # There's no dialog.  Just say the user said OK.
            result = wx.ID_OK
            # Include Episodes, Clips, and Snapshots
            includeEpisodes = True
            includeClips = True
            if TransanaConstants.proVersion:
                includeSnapshots = True
            else:
                includeSnapshots = False
        # if kwg and kw are passed in, we're doing a Quick Search
        else:
            # There's no dialog.  Just say the user said OK.
            result = wx.ID_OK
            # The Search Name is built from the kwg : kw combination
            searchName = "%s : %s" % (kwg, kw)
            # The Search Terms are just the keyword group and keyword passed in
            searchTerms = ["%s:%s" % (kwg, kw)]
            # Include Episodes, Clips and Snapshots
            includeEpisodes = True
            includeClips = True
            if TransanaConstants.proVersion:
                includeSnapshots = True
            else:
                includeSnapshots = False

        # If OK is pressed (or Quick Search), process the requested Search
        if result == wx.ID_OK:
            # Increment the Search Counter
            self.searchCount = searchCount + 1
            # The "Search" node itself is always item 0 in the node list
            searchNode = self.dbTree.select_Node((_("Search"),), 'SearchRootNode')
            # We need to collect a list of the named searches already done.
            namedSearches = []
            # Get the first child node from the Search root node
            (childNode, cookieVal) = self.dbTree.GetFirstChild(searchNode)
            # As long as there are child nodes ...
            while childNode.IsOk():
                # Add the node name to the named searches list ...
                namedSearches.append(self.dbTree.GetItemText(childNode))
                # ... and get the next child node
                (childNode, cookieVal) = self.dbTree.GetNextChild(childNode, cookieVal)
            # We need to give each search result a unique name.  So note the search count number
            nameIncrementValue = searchCount
            # As long as there's already a named search with the name we want to use ...
            while (searchName in namedSearches):
                # ... if this is our FIRST attempt ...
                if nameIncrementValue == searchCount:
                    # ... append the appropriate number on the end of the search name
                    searchName += unicode(_(' - Search %d'), 'utf8') % nameIncrementValue
                # ... if this is NOT our first attempt ...
                else:
                    # ... remove the previous number and add the appropriate next number to try
                    searchName = searchName[:searchName.rfind(' ')] + ' %d' % nameIncrementValue
                # Increment our counter by one.  We'll keep trying new numbers until we find one that works.
                nameIncrementValue += 1
            # As long as there's a search name (and there's no longer a way to eliminate it!
            if searchName != '':
                # Add a Search Results Node to the Database Tree
                nodeListBase = [_("Search"), searchName]
                self.dbTree.add_Node('SearchResultsNode', nodeListBase, 0, 0, expandNode=True)

                # Build the appropriate Queries based on the Search Query specified in the Search Dialog.
                # (This method parses the Natural Language Search Terms into queries for Episode Search
                #  Terms, for Clip Search Terms, and for Snapshot Search Terms, and includes the appropriate 
                #  Parameters to be used with the queries.  Parameters are not integrated into the queries 
                #  in order to allow for automatic processing of apostrophes and other text that could 
                #  otherwise interfere with the SQL execution.)
                (episodeQuery, clipQuery, wholeSnapshotQuery, snapshotCodingQuery, params) = self.BuildQueries(searchTerms)

                # Get a Database Cursor
                dbCursor = DBInterface.get_db().cursor()

                if includeEpisodes:
                    # Adjust query for sqlite, if needed
                    episodeQuery = DBInterface.FixQuery(episodeQuery)
                    # Execute the Series/Episode query
                    dbCursor.execute(episodeQuery, tuple(params))

                    # Process the results of the Series/Episode query
                    for line in DBInterface.fetchall_named(dbCursor):
                        # Add the new Transcript(s) to the Database Tree Tab.
                        # To add a Transcript, we need to build the node list for the tree's add_Node method to climb.
                        # We need to add the Series, Episode, and Transcripts to our Node List, so we'll start by loading
                        # the current Series and Episode
                        tempSeries = Series.Series(line['SeriesNum'])
                        tempEpisode = Episode.Episode(line['EpisodeNum'])
                        # Add the Search Root Node, the Search Name, and the current Series and Episode Names.
                        nodeList = (_('Search'), searchName, tempSeries.id, tempEpisode.id)
                        # Find out what Transcripts exist for each Episode
                        transcriptList = DBInterface.list_transcripts(tempSeries.id, tempEpisode.id)
                        # If the Episode HAS defined transcripts ...
                        if len(transcriptList) > 0:
                            # Add each Transcript to the Database Tree
                            for (transcriptNum, transcriptID, episodeNum) in transcriptList:
                                # Add the Transcript Node to the Tree.  
                                self.dbTree.add_Node('SearchTranscriptNode', nodeList + (transcriptID,), transcriptNum, episodeNum)
                        # If the Episode has no transcripts, it still has the keywords and SHOULD be displayed!
                        else:
                            # Add the Transcript-less Episode Node to the Tree.  
                            self.dbTree.add_Node('SearchEpisodeNode', nodeList, tempEpisode.number, tempSeries.number)

                if includeClips:
                    # Adjust query for sqlite, if needed
                    clipQuery = DBInterface.FixQuery(clipQuery)
                    # Execute the Collection/Clip query
                    dbCursor.execute(clipQuery, params)

                    # Process all results of the Collection/Clip query 
                    for line in DBInterface.fetchall_named(dbCursor):
                        # Add the new Clip to the Database Tree Tab.
                        # To add a Clip, we need to build the node list for the tree's add_Node method to climb.
                        # We need to add all of the Collection Parents to our Node List, so we'll start by loading
                        # the current Collection
                        tempCollection = Collection.Collection(line['CollectNum'])

                        # Add the current Collection Name, and work backwards from here.
                        nodeList = (tempCollection.id,)
                        # Repeat this process as long as the Collection we're looking at has a defined Parent...
                        while tempCollection.parent > 0:
                           # Load the Parent Collection
                           tempCollection = Collection.Collection(tempCollection.parent)
                           # Add this Collection's name to the FRONT of the Node List
                           nodeList = (tempCollection.id,) + nodeList
                        # Get the DB Values
                        tempID = line['ClipID']
                        # If we're in Unicode mode, format the strings appropriately
                        if 'unicode' in wx.PlatformInfo:
                            tempID = DBInterface.ProcessDBDataForUTF8Encoding(tempID)
                        # Now add the Search Root Node and the Search Name to the front of the Node List and the
                        # Clip Name to the back of the Node List
                        nodeList = (_('Search'), searchName) + nodeList + (tempID, )

                        # Add the Node to the Tree
                        self.dbTree.add_Node('SearchClipNode', nodeList, line['ClipNum'], line['CollectNum'], sortOrder=line['SortOrder'])

                if includeSnapshots:
                    # Adjust query for sqlite, if needed
                    wholeSnapshotQuery = DBInterface.FixQuery(wholeSnapshotQuery)
                    # Execute the Whole Snapshot query
                    dbCursor.execute(wholeSnapshotQuery, params)

                    # Since we have two sources of Snapshots that get included, we need to track what we've already
                    # added so we don't add the same Snapshot twice
                    addedSnapshots = []

                    # Process all results of the Whole Snapshot query 
                    for line in DBInterface.fetchall_named(dbCursor):
                        # Add the new Snapshot to the Database Tree Tab.
                        # To add a Snapshot, we need to build the node list for the tree's add_Node method to climb.
                        # We need to add all of the Collection Parents to our Node List, so we'll start by loading
                        # the current Collection
                        tempCollection = Collection.Collection(line['CollectNum'])

                        # Add the current Collection Name, and work backwards from here.
                        nodeList = (tempCollection.id,)
                        # Repeat this process as long as the Collection we're looking at has a defined Parent...
                        while tempCollection.parent > 0:
                           # Load the Parent Collection
                           tempCollection = Collection.Collection(tempCollection.parent)
                           # Add this Collection's name to the FRONT of the Node List
                           nodeList = (tempCollection.id,) + nodeList
                        # Get the DB Values
                        tempID = line['SnapshotID']
                        # If we're in Unicode mode, format the strings appropriately
                        if 'unicode' in wx.PlatformInfo:
                            tempID = DBInterface.ProcessDBDataForUTF8Encoding(tempID)
                        # Now add the Search Root Node and the Search Name to the front of the Node List and the
                        # Clip Name to the back of the Node List
                        nodeList = (_('Search'), searchName) + nodeList + (tempID, )

                        # Add the Node to the Tree
                        self.dbTree.add_Node('SearchSnapshotNode', nodeList, line['SnapshotNum'], line['CollectNum'], sortOrder=line['SortOrder'])
                        # Add the Snapshot to the list of Snapshots added to the Search Result
                        addedSnapshots.append(line['SnapshotNum'])
                        
                        tmpNode = self.dbTree.select_Node(nodeList[:-1], 'SearchCollectionNode', ensureVisible=False)
                        self.dbTree.SortChildren(tmpNode)
                    # Adjust query for sqlite if needed
                    snapshotCodingQuery = DBInterface.FixQuery(snapshotCodingQuery)
                    # Execute the Snapshot Coding query
                    dbCursor.execute(snapshotCodingQuery, params)

                    # Process all results of the Snapshot Coding query 
                    for line in DBInterface.fetchall_named(dbCursor):
                        # If the Snapshot is NOT already in the Search Results ...
                        if not (line['SnapshotNum'] in addedSnapshots):
                            # Add the new Snapshot to the Database Tree Tab.
                            # To add a Snapshot, we need to build the node list for the tree's add_Node method to climb.
                            # We need to add all of the Collection Parents to our Node List, so we'll start by loading
                            # the current Collection
                            tempCollection = Collection.Collection(line['CollectNum'])

                            # Add the current Collection Name, and work backwards from here.
                            nodeList = (tempCollection.id,)
                            # Repeat this process as long as the Collection we're looking at has a defined Parent...
                            while tempCollection.parent > 0:
                               # Load the Parent Collection
                               tempCollection = Collection.Collection(tempCollection.parent)
                               # Add this Collection's name to the FRONT of the Node List
                               nodeList = (tempCollection.id,) + nodeList
                            # Get the DB Values
                            tempID = line['SnapshotID']
                            # If we're in Unicode mode, format the strings appropriately
                            if 'unicode' in wx.PlatformInfo:
                                tempID = DBInterface.ProcessDBDataForUTF8Encoding(tempID)
                            # Now add the Search Root Node and the Search Name to the front of the Node List and the
                            # Clip Name to the back of the Node List
                            nodeList = (_('Search'), searchName) + nodeList + (tempID, )

                            # Add the Node to the Tree
                            self.dbTree.add_Node('SearchSnapshotNode', nodeList, line['SnapshotNum'], line['CollectNum'], sortOrder=line['SortOrder'])
                            # Add the Snapshot to the list of Snapshots added to the Search Result
                            addedSnapshots.append(line['SnapshotNum'])
                            
                            tmpNode = self.dbTree.select_Node(nodeList[:-1], 'SearchCollectionNode', ensureVisible=False)
                            self.dbTree.SortChildren(tmpNode)

            else:
                self.searchCount = searchCount

        # If the Search Dialog is cancelled, do NOT increment the Search Number                
        else:
            self.searchCount = searchCount


    def GetSearchCount(self):
        """ This method is called to determine whether the Search Counter was incremented, that is, whether the
            search was performed or cancelled. """
        return self.searchCount


    def BuildQueries(self, queryText):
        """ Convert natural language search terms (as structured by the Transana Search Dialog) into
            executable SQL that runs on MySQL. """

        # Here are a couple of sample SQL Statements generated by this code:
        #
        # Query:  "Demo:Geometry AND NOT Demo:Teacher Commentary"
        #
        # SELECT Ep.SeriesNum, SeriesID, Ep.EpisodeNum, EpisodeID,
        #        COUNT(CASE WHEN ((CK1.KeywordGroup = 'Demo') AND (CK1.Keyword = 'Geometry')) THEN 1 ELSE NULL END) V1,
        #        COUNT(CASE WHEN ((CK1.KeywordGroup = 'Demo') AND (CK1.Keyword = 'Teacher Commentary')) THEN 1 ELSE NULL END) V2
        #   FROM ClipKeywords2 CK1, Series2 Se, Episodes2 Ep
        #   WHERE (Ep.EpisodeNum = CK1.EpisodeNum) AND (Ep.SeriesNum = Se.SeriesNum) AND (CK1.EpisodeNum > 0)
        #   GROUP BY SeriesNum, SeriesID, EpisodeNum, EpisodeID
        #   HAVING (V1 > 0) AND (V2 = 0)
        #
        # SELECT Cl.CollectNum, ParentCollectNum, Cl.ClipNum, CollectID, ClipID,
        #        COUNT(CASE WHEN ((CK1.KeywordGroup = 'Demo') AND (CK1.Keyword = 'Geometry')) THEN 1 ELSE NULL END) V1,
        #        COUNT(CASE WHEN ((CK1.KeywordGroup = 'Demo') AND (CK1.Keyword = 'Teacher Commentary')) THEN 1 ELSE NULL END) V2
        #   FROM ClipKeywords2 CK1, Collections2 Co, Clips2 Cl
        #   WHERE (Cl.ClipNum = CK1.ClipNum) AND (Cl.CollectNum = Co.CollectNum) AND (CK1.ClipNum > 0)
        #   GROUP BY Cl.CollectNum, CollectID, ClipID
        #   HAVING (V1 > 0) AND (V2 = 0)
        
        # Initialize a Temporary Variable Counter
        tempVarNum = 0
        # Initialize a list for strings to store SQL "COUNT" lines
        countStrings = []
        # Initialize a list to hold the Search Parameters.
        # NOTE:  Parameters are passed separately rather than being integrated into the SQL so that
        #        MySQLdb can handle all parsing related to apostrophes and other non-SQL-friendly characters.
        params = []
        # Initialize a String to store the SQL "HAVING" clause
        havingStr = ''

        # We now will go through the Search Terms line by line and prepare to convert the Search Request to SQL
        for lineNum in range(len(queryText)):
            # Capture the Line being processed, and remove whitespace from either end
            tempStr = string.strip(queryText[lineNum])

            # Initialize the "Continuation" string, which holds a BOOLEAN Operator ("AND" or "OR")
            continStr = ''
            # Initialize the flag that signals the BOOLEAN "NOT" Operator
            notFlag = False
            # Initialize the counter that tracks the number of parentheses that are open and need to be closed.
            closeParen = 0

            # If a line ends with " AND"...
            if tempStr[-4:] == ' AND':
                # ... put the Boolean Operator into the Continuation String ...
                continStr = ' AND '
                # ... and remove it from the line being processed.
                tempStr = tempStr[:-4]

            # If a line ends with " OR"...
            if tempStr[-3:] == ' OR':
                # ... put the Boolean Operator into the Continuation String ...
                continStr = ' OR '
                # ... and remove it from the line being processed.
                tempStr = tempStr[:-3]

            # Process characters at the beginning of the Line, including open parens and the "NOT" operator.
            # NOTE:  The Search Dialog allows "(NOT", but not "NOT(".
            while (tempStr[0] == '(') or (tempStr[:4] == 'NOT '):
                # If the line starts with an open paren ...
                if tempStr[0] == '(':
                    # ... add it to the "HAVING" clause string ...
                    havingStr += '('
                    # ... and remove it from the line.
                    tempStr = tempStr[1:]
                # If the line starts with a "NOT" operator ...
                if tempStr[:4] == 'NOT ':
                    # ... set the NOT Flag ...
                    notFlag = True
                    # ... and remove it from the line.
                    tempStr = tempStr[4:]

            # Check for close parens in the line ...
            while tempStr.find(')') > -1:
                # ... keep track of how many are found in this line ...
                closeParen += 1
                # ... and remove them from the line.
                tempStr = tempStr[:tempStr.find(')')] + tempStr[tempStr.find(')') + 1:]

            # All that should be left in the line being processed now should be Keywords.
            if len(tempStr) > 0:
                # increment the Temporary Variable Counter.  (Every Keyword Group : Keyword combination gets a unique
                # Temporary Variable Number.)
                tempVarNum += 1

                # The presence of a variable (or it's absence if NOT has been specified) is signalled in SQL by a combination of
                # this "COUNT" statement, which creates a numbered variable in the SELECT Clause, and a "HAVING" line.
                # I can't adequately explain it, but it DOES work.
                # Please, don't mess with it.

                # Add a line to the SQL "COUNT" statements to indicate the presence or absence of a Keyword Group : Keyword pair
                tempStr2 = "COUNT(CASE WHEN ((CK1.KeywordGroup = %s) AND (CK1.Keyword = %s)) THEN 1 ELSE NULL END) " + "V%s" % tempVarNum
                countStrings.append(tempStr2)
                # Add the Keyword Group to the Parameters
                kwg = tempStr[:tempStr.find(':')]
                if 'unicode' in wx.PlatformInfo:
                    kwg = kwg.encode(TransanaGlobal.encoding)
                params.append(kwg)
                # Add the Keyword to the Parameters
                kw = tempStr[tempStr.find(':') + 1:]
                if 'unicode' in wx.PlatformInfo:
                    kw = kw.encode(TransanaGlobal.encoding)
                params.append(kw)
                # Add the Temporary Variable Number that corresponds to this Keyword Group : Keyword pair to the Parameters
#                params.append(tempVarNum)

                # If the "NOT" operator has been specified, we want the Temporary Variable to equal Zero in the "HAVING" clause
                if notFlag:
                    havingStr += '(V%s = 0)' % tempVarNum
                # If the "NOT" operator has not been specified, we want the Temporary Variable to be greater than Zero in the "HAVING" clause
                else:
                    havingStr += '(V%s > 0)' % tempVarNum

                # Add any closing parentheses that were specified to the end of the "HAVING" clause
                for x in range(closeParen):
                    havingStr += ')'
                # Add the appropriate Boolean Operator to the end of the "HAVING" clause, if one was specified
                havingStr += continStr

        # Before we continue, let's build the part of the query that implements the Collections selections
        # made on the Collections tab of the Search Form

        if len(self.collectionList) > 0:
            paramsCl = ()
            paramsSn = ()
            collectionSQL = ' AND ('
            for coll in self.collectionList:
                collectionSQL += "(%%s.CollectNum = %d) " % coll[0]
                if coll != self.collectionList[-1]:
                    collectionSQL += "or "
                paramsCl+= ('Cl',)
                paramsSn += ('Sn',)
            collectionSQL += ") "

        # Now that all the pieces (countStrings, params, and the havingStr) are assembled, we can build the
        # SQL Statements for the searches.

        # Define the start of the Series/Episode Query
        episodeSQL = 'SELECT Ep.SeriesNum, SeriesID, Ep.EpisodeNum, EpisodeID, '
        # Define the start of the Collection/Clip Query
        clipSQL = 'SELECT Cl.CollectNum, ParentCollectNum, Cl.ClipNum, CollectID, ClipID, SortOrder, '
        # Define the start of the Whole Snapshot Query
        wholeSnapshotSQL = 'SELECT Sn.CollectNum, ParentCollectNum, Sn.SnapshotNum, CollectID, SnapshotID, SortOrder, '
        # Define the start of the Snapshot Coding Query
        snapshotCodingSQL = 'SELECT Sn.CollectNum, ParentCollectNum, Sn.SnapshotNum, CollectID, SnapshotID, SortOrder, '

        # Add in the SQL "COUNT" variables that signal the presence or absence of Keyword Group : Keyword pairs
        for lineNum in range(len(countStrings)):
            # All SQL "COUNT" lines but he last one need to end with a comma
            if lineNum < len(countStrings)-1:
                tempStr = ', '
            # The last SQL "COUNT" line does not need to end with a comma
            else:
                tempStr = ' '

            # Add the SQL "COUNT" Line and seperator to the Series/Episode Query
            episodeSQL += countStrings[lineNum] + tempStr
            # Add the SQL "COUNT" Line and seperator to the Collection/Clip Query
            clipSQL += countStrings[lineNum] + tempStr
            # Add the SQL "COUNT" Line and seperator to the Whole Snapshot Query
            wholeSnapshotSQL += countStrings[lineNum] + tempStr
            # Add the SQL "COUNT" Line and seperator to the Snapshot Coding Query
            snapshotCodingSQL += countStrings[lineNum] + tempStr

        # Now add the rest of the SQL for the Series/Episode Query
        episodeSQL += 'FROM ClipKeywords2 CK1, Series2 Se, Episodes2 Ep '
        episodeSQL += 'WHERE (Ep.EpisodeNum = CK1.EpisodeNum) AND '
        episodeSQL += '(Ep.SeriesNum = Se.SeriesNum) AND '
        episodeSQL += '(CK1.EpisodeNum > 0) '
        episodeSQL += 'GROUP BY Ep.SeriesNum, SeriesID, Ep.EpisodeNum, EpisodeID '
        # Add in the SQL "HAVING" Clause that was constructed above
        episodeSQL += 'HAVING %s ' % havingStr

        # Now add the rest of the SQL for the Collection/Clip Query
        clipSQL += 'FROM ClipKeywords2 CK1, Collections2 Co, Clips2 Cl '
        clipSQL += 'WHERE (Cl.ClipNum = CK1.ClipNum) AND '
        clipSQL += '(Cl.CollectNum = Co.CollectNum) AND '
        clipSQL += '(CK1.ClipNum > 0) '
        if len(self.collectionList) > 0:
            clipSQL += collectionSQL % paramsCl
        clipSQL += 'GROUP BY Cl.CollectNum, CollectID, ClipID '
        # Add in the SQL "HAVING" Clause that was constructed above
        clipSQL += 'HAVING %s ' % havingStr
        # Add an "ORDER BY" Clause to preserve Clip Sort Order
        clipSQL += 'ORDER BY CollectID, SortOrder'

        # Now add the rest of the SQL for the Whole Snapshot Query
        wholeSnapshotSQL += 'FROM ClipKeywords2 CK1, Collections2 Co, Snapshots2 Sn '
        wholeSnapshotSQL += 'WHERE (Sn.SnapshotNum = CK1.SnapshotNum) AND '
        wholeSnapshotSQL += '(Sn.CollectNum = Co.CollectNum) AND '
        wholeSnapshotSQL += '(CK1.SnapshotNum > 0) '
        if len(self.collectionList) > 0:
            wholeSnapshotSQL += collectionSQL % paramsSn
        wholeSnapshotSQL += 'GROUP BY Sn.CollectNum, CollectID, SnapshotID '
        # Add in the SQL "HAVING" Clause that was constructed above
        wholeSnapshotSQL += 'HAVING %s ' % havingStr
        # Add an "ORDER BY" Clause to preserve Snapshot Sort Order
        wholeSnapshotSQL += 'ORDER BY CollectID, SortOrder'

        # Now add the rest of the SQL for the Snapshot Coding Query
        snapshotCodingSQL += 'FROM SnapshotKeywords2 CK1, Collections2 Co, Snapshots2 Sn '
        snapshotCodingSQL += 'WHERE (Sn.SnapshotNum = CK1.SnapshotNum) AND '
        snapshotCodingSQL += '(Sn.CollectNum = Co.CollectNum) AND '
        snapshotCodingSQL += '(CK1.SnapshotNum > 0) '
        # For Snapshot Coding, we ONLY want VISIBLE Keywords
        snapshotCodingSQL += 'AND (CK1.Visible = 1) '
        if len(self.collectionList) > 0:
            snapshotCodingSQL += collectionSQL % paramsSn
        snapshotCodingSQL += 'GROUP BY Sn.CollectNum, CollectID, SnapshotID '
        # Add in the SQL "HAVING" Clause that was constructed above
        snapshotCodingSQL += 'HAVING %s ' % havingStr
        # Add an "ORDER BY" Clause to preserve Snapshot Sort Order
        snapshotCodingSQL += 'ORDER BY CollectID, SortOrder'

#        tempParams = ()
#        for p in params:
#            tempParams = tempParams + (p,)
            
        # dlg = wx.TextEntryDialog(None, "Transana Series/Episode SQL Statement:", "Transana", episodeSQL % tempParams, style=wx.OK)
        # dlg.ShowModal()
        # dlg.Destroy()

#        dlg = wx.TextEntryDialog(None, "Transana Collection/Clip SQL Statement:", "Transana", clipSQL % tempParams, style=wx.OK)
#        dlg.ShowModal()
#        dlg.Destroy()

#        dlg = wx.TextEntryDialog(None, "Transana Whole Snapshot SQL Statement:", "Transana", wholeSnapshotSQL % tempParams, style=wx.OK)
#        dlg.ShowModal()
#        dlg.Destroy()

#        dlg = wx.TextEntryDialog(None, "Transana Snapshot Coding SQL Statement:", "Transana", snapshotCodingSQL % tempParams, style=wx.OK)
#        dlg.ShowModal()
#        dlg.Destroy()

        # Return the Series/Episode Query, the Collection/Clip Query, the Whole Snapshot Query, the Snapshot Coding Query, 
        # and the list of parameters to use with these queries to the calling routine.
        return (episodeSQL, clipSQL, wholeSnapshotSQL, snapshotCodingSQL, params)
