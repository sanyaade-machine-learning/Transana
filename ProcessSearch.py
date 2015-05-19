# Copyright (C) 2003-2007 The Board of Regents of the University of Wisconsin System 
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
# Import Transana's Globals
import TransanaGlobal

# Import the Python String module
import string


class ProcessSearch(object):
    """ This class handles all processing related to Searching. """
    def __init__(self, dbTree, searchCount):
        """ Initialize the ProcessSearch class.  The dbTree parameter accepts a wxTreeCtrl as the Database Tree where
            Search Results should be displayed.  The searchCount parameter accepts the number that should be included
            in the Default Search Title. """

        # Note the Database Tree that accepts Search Results
        self.dbTree = dbTree

        # Create the Search Dialog Box
        dlg = SearchDialog.SearchDialog(_("Search") + " %s" % searchCount)
        # Display the Search Dialog Box and record the Result
        result = dlg.ShowModal()
        # If OK is pressed, process the requested Search
        if result == wx.ID_OK:
            # Increment the Search Counter
            self.searchCount = searchCount + 1

            # The "Keywords" node itself is always item 0 in the node list
            searchNode = self.dbTree.select_Node((_("Search"),), 'SearchRootNode')
            namedSearches = []
            (childNode, cookieVal) = self.dbTree.GetFirstChild(searchNode)
            while childNode.IsOk():
                namedSearches.append(self.dbTree.GetItemText(childNode))
                (childNode, cookieVal) = self.dbTree.GetNextChild(childNode, cookieVal)

            searchName = dlg.searchName.GetValue().strip()

            while (searchName in namedSearches):
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('You already have a Search Result named "%s".\nPlease enter a new Search Name.'), 'utf8')
                else:
                    prompt = _('You already have a Search Result named "%s".\nPlease enter a new Search Name.')
                dlg2 = wx.TextEntryDialog(dlg, prompt % searchName, _('Transana Error'), searchName)
                if dlg2.ShowModal() == wx.ID_OK:
                    searchName = dlg2.GetValue().strip()
                else:
                    searchName = ''
                dlg2.Destroy()
                
            if searchName != '':
                # Add a Search Results Node to the Database Tree
                nodeListBase = [_("Search"), searchName]
                self.dbTree.add_Node('SearchResultsNode', nodeListBase, 0, 0, True)

                # Build the appropriate Queries based on the Search Query specified in the Search Dialog.
                # (This method parses the Natural Language Search Terms into queries for Episode Search
                #  Terms and for Clip Search Terms, and includes the appropriate Parameters to be used
                #  with the queries.  Parameters are not integrated into the queries in order to allow
                #  for automatic processing of apostrophes and other text that could otherwise interfere
                #  with the SQL execution.)
                (seriesQuery, collectionQuery, params) = self.BuildQueries(dlg.searchQuery)

                # Get a Database Cursor
                dbCursor = DBInterface.get_db().cursor()
                # Execute the Series/Episode query
                dbCursor.execute(seriesQuery, params)

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
                    # Add each Transcript to the Database Tree
                    for (transcriptNum, transcriptID, episodeNum) in transcriptList:
                        # Add the Transcript Node to the Tree.  
                        self.dbTree.add_Node('SearchTranscriptNode', nodeList + (transcriptID,), transcriptNum, episodeNum)

                # Execute the Collection/Clip query
                dbCursor.execute(collectionQuery, params)

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
                    self.dbTree.add_Node('SearchClipNode', nodeList, line['ClipNum'], line['CollectNum'])
            else:
                self.searchCount = searchCount

        # If the Search Dialog is cancelled, do NOT increment the Search Number                
        else:
            self.searchCount = searchCount

        # Destroy the Search Dialog Box
        dlg.Destroy()


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
        for lineNum in range(queryText.GetNumberOfLines()):
            # Capture the Line being processed, and remove whitespace from either end
            tempStr = string.strip(queryText.GetLineText(lineNum))

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
                countStrings.append("COUNT(CASE WHEN ((CK1.KeywordGroup = %s) AND (CK1.Keyword = %s)) THEN 1 ELSE NULL END) V%s")
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
                params.append(tempVarNum)

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

        # Now that all the pieces (countStrings, params, and the havingStr) are assembled, we can build the
        # SQL Statements for the searches.

        # Define the start of the Series/Episode Query
        seriesSQL = 'SELECT Ep.SeriesNum, SeriesID, Ep.EpisodeNum, EpisodeID, '
        # Define the start of the Collection/Clip Query
        collectionSQL = 'SELECT Cl.CollectNum, ParentCollectNum, Cl.ClipNum, CollectID, ClipID, '

        # Add in the SQL "COUNT" variables that signal the presence or absence of Keyword Group : Keyword pairs
        for lineNum in range(len(countStrings)):
            # All SQL "COUNT" lines but he last one need to end with a comma
            if lineNum < len(countStrings)-1:
                tempStr = ', '
            # The last SQL "COUNT" line does not need to end with a comma
            else:
                tempStr = ' '

            # Add the SQL "COUNT" Line and seperator to the Series/Episode Query
            seriesSQL += countStrings[lineNum] + tempStr
            # Add the SQL "COUNT" Line and seperator to the Collection/Clip Query
            collectionSQL += countStrings[lineNum] + tempStr

        # Now add the rest of the SQL for the Series/Episode Query
        seriesSQL += 'FROM ClipKeywords2 CK1, Series2 Se, Episodes2 Ep '
        seriesSQL += 'WHERE (Ep.EpisodeNum = CK1.EpisodeNum) AND '
        seriesSQL += '(Ep.SeriesNum = Se.SeriesNum) AND '
        seriesSQL += '(CK1.EpisodeNum > 0) '
        seriesSQL += 'GROUP BY SeriesNum, SeriesID, EpisodeNum, EpisodeID '
        # Add in the SQL "HAVING" Clause that was constructed above
        seriesSQL += 'HAVING %s' % havingStr

        # Now add the rest of the SQL for the Collection/Clip Query
        collectionSQL += 'FROM ClipKeywords2 CK1, Collections2 Co, Clips2 Cl '
        collectionSQL += 'WHERE (Cl.ClipNum = CK1.ClipNum) AND '
        collectionSQL += '(Cl.CollectNum = Co.CollectNum) AND '
        collectionSQL += '(CK1.ClipNum > 0) '
        collectionSQL += 'GROUP BY Cl.CollectNum, CollectID, ClipID '
        # Add in the SQL "HAVING" Clause that was constructed above
        collectionSQL += 'HAVING %s' % havingStr
        # Add an "ORDER BY" Clause to preserve Clip Sort Order
        collectionSQL += 'ORDER BY CollectID, SortOrder'

        # tempParams = ()
        # for p in params:
        #     tempParams = tempParams + (p,)
            
        # dlg = wx.TextEntryDialog(None, "Transana Series/Episode SQL Statement:", "Transana", seriesSQL % tempParams, style=wx.OK)
        # dlg.ShowModal()
        # dlg.Destroy()

        # dlg = wx.TextEntryDialog(None, "Transana Collection/Clip SQL Statement:", "Transana", collectionSQL % tempParams, style=wx.OK)
        # dlg.ShowModal()
        # dlg.Destroy()

        # Return the Series/Episode Query, the Collection/Clip Query, and the list of parameters to use with
        # both queries to the calling routine.
        return (seriesSQL, collectionSQL, params)
