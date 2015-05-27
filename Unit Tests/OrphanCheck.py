""" This program looks for orphaned data records in Transana """

import MySQLdb

import os, sys
import wx

class OrphanCheck(object):
    """ This is the main window for the Orphan Data Record Checker """
    def __init__(self, parent, ID, title):

        pw = raw_input('Enter Password:')
        print

        self.testsTried = 0
        self.testsPassed = 0

        config = wx.Config('Transana', 'Verception')

        hostName = config.Read('/2.0/hostMU', '')  # '128.104.150.81'
        dbName = config.Read('/2.0/databaseMU', '')  # 'Transana_250a2'
        
        print "Unit Test: OrphanCheck.py"
        print
        print "Database Server:", hostName
        print "Database:", dbName
        print
        print "Checking for Orphaned Database Records:"
        print
        
        # Get all Library Records
        #  Create a connection to the database

        DBConn = MySQLdb.Connection(host=hostName, user='DavidW', passwd=pw, port=3306)  #  db=dbName,
        #  Create a cursor and execute the appropriate query
        self.DBCursor = DBConn.cursor()

        self.DBCursor.execute('SET CHARACTER SET utf8')
        self.DBCursor.execute('SET character_set_connection = utf8')
        self.DBCursor.execute('SET character_set_client = utf8')
        self.DBCursor.execute('SET character_set_server = utf8')
        self.DBCursor.execute('SET character_set_database = utf8')
        self.DBCursor.execute('SET character_set_results = utf8')


        self.DBCursor.execute('USE %s' % dbName.encode('utf8'))

        print "Get Database Version"
        query = "SELECT Value FROM ConfigInfo WHERE KeyVal = 'DBVersion'"
        self.DBCursor.execute(query)
        Records = self.DBCursor.fetchone()
        DBVersion = int(Records[0])
        print "Database Version:", DBVersion
        print
        print

        print "Check for Library Notes without Library:"
        query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, DocumentNum, QuoteNum
                     FROM Notes2
                     WHERE (SeriesNum <> 0) AND
                           (SeriesNum not in (SELECT SeriesNum FROM Series2))"""
        self.ExecQuery(query)

        if DBVersion >= 260:
            print "Check for Documents without Library:"
            query = """SELECT DocumentNum, DocumentID, LibraryNum
                         FROM Documents2
                         WHERE (LibraryNum not in (SELECT SeriesNum FROM Series2))"""
            self.ExecQuery(query)

            print "Check for Document Keywords without Document:"
            if DBVersion > 260:
                query = """SELECT EpisodeNum, DocumentNum, ClipNum, QuoteNum, SnapshotNum, KeywordGroup, Keyword
                             FROM ClipKeywords2
                             WHERE (DocumentNum > 0) AND
                                   (DocumentNum not in (SELECT DocumentNum FROM Documents2))"""
            elif DBVersion > 250:
                query = 'SELECT EpisodeNum, ClipNum, SnapshotNum, KeywordGroup, Keyword FROM ClipKeywords2  WHERE (ClipNum = 0) AND (SnapshotNum = 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
            else:
                query = 'SELECT EpisodeNum, ClipNum, KeywordGroup, Keyword FROM ClipKeywords2  WHERE (ClipNum = 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
            self.ExecQuery(query)

            print "Check for QuotePositions without Document:"
            query = """SELECT QuoteNum, DocumentNum, StartChar, EndChar
                         FROM QuotePositions2
                         WHERE (DocumentNum > 0) AND
                               (DocumentNum not in (SELECT DocumentNum FROM Documents2))"""
            self.ExecQuery(query)

            print "Check for Document Notes without Document:"
            query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, DocumentNum, QuoteNum
                         FROM Notes2
                         WHERE (DocumentNum <> 0) AND
                               (DocumentNum not in (SELECT DocumentNum FROM Documents2))"""
            self.ExecQuery(query)

        print "Check for Episodes without Library:"
        query = 'SELECT EpisodeNum, EpisodeID, SeriesNum FROM Episodes2  WHERE (SeriesNum not in (SELECT SeriesNum FROM Series2))'
        self.ExecQuery(query)

        print "Check for Episode Transcripts without Episodes:"
        query = 'SELECT TranscriptNum, EpisodeNum, SourceTranscriptNum, ClipNum FROM Transcripts2  WHERE (ClipNum = 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
        self.ExecQuery(query)

        print "Check for Episode Additional Video Files without Episodes:"
        query = 'SELECT AddVidNum, EpisodeNum, ClipNum FROM AdditionalVids2  WHERE (ClipNum = 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
        self.ExecQuery(query)

        print "Check for Episode Keywords without Episodes:"
        if DBVersion > 260:
            query = """SELECT EpisodeNum, DocumentNum, ClipNum, QuoteNum, SnapshotNum, KeywordGroup, Keyword
                         FROM ClipKeywords2
                         WHERE (EpisodeNum > 0) AND
                               (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))"""
        elif DBVersion > 250:
            query = 'SELECT EpisodeNum, ClipNum, SnapshotNum, KeywordGroup, Keyword FROM ClipKeywords2  WHERE (ClipNum = 0) AND (SnapshotNum = 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
        else:
            query = 'SELECT EpisodeNum, ClipNum, KeywordGroup, Keyword FROM ClipKeywords2  WHERE (ClipNum = 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
        self.ExecQuery(query)

        print "Check for Episode Notes without Episodes:"
        query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, DocumentNum, QuoteNum
                     FROM Notes2
                     WHERE (EpisodeNum <> 0) AND
                           (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))"""
        self.ExecQuery(query)

        print "Check for Transcript Notes without Transcripts:"
        query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, DocumentNum, QuoteNum
                     FROM Notes2
                     WHERE (TranscriptNum <> 0) AND
                           (TranscriptNum not in (SELECT TranscriptNum FROM Transcripts2))"""
        self.ExecQuery(query)

        print "Check for Nested Collections without Parents:"
        query = 'SELECT CollectNum, CollectID, ParentCollectNum FROM Collections2  WHERE (ParentCollectNum <> 0) AND (ParentCollectNum not in (SELECT CollectNum FROM Collections2))'
        self.ExecQuery(query)

        if DBVersion >= 260:
            print "Check for Quote Notes without Quote:"
            query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, DocumentNum, QuoteNum
                         FROM Notes2
                         WHERE (QuoteNum <> 0) AND
                               (QuoteNum not in (SELECT QuoteNum FROM Quotes2))"""
            self.ExecQuery(query)

        print "Check for Collection Notes without Collections:"
        query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, DocumentNum, QuoteNum
                     FROM Notes2
                     WHERE (CollectNum <> 0) AND
                           (CollectNum not in (SELECT CollectNum FROM Collections2))"""
        self.ExecQuery(query)

        if DBVersion >= 270:
            print "Check for Quotes without Collections:"
            query = """SELECT QuoteNum, QuoteID, CollectNum
                         FROM Quotes2
                         WHERE (CollectNum not in (SELECT CollectNum FROM Collections2))"""
            self.ExecQuery(query)

            print "Check for QuotePositions without Quote:"
            query = """SELECT QuoteNum, DocumentNum, StartChar, EndChar
                         FROM QuotePositions2
                         WHERE (QuoteNum not in (SELECT QuoteNum FROM Quotes2))"""
            self.ExecQuery(query)

            print "Check for QuotePositions with negative StartChar:"
            query = """SELECT QuoteNum, DocumentNum, StartChar, EndChar
                         FROM QuotePositions2
                         WHERE StartChar < 0"""
            self.ExecQuery(query)

            print "Check for QuotePositions with negative EndChar:"
            query = """SELECT QuoteNum, DocumentNum, StartChar, EndChar
                         FROM QuotePositions2
                         WHERE EndChar < 0"""
            self.ExecQuery(query)

            print "Check for Quote Notes without Quote:"
            query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, DocumentNum, QuoteNum
                         FROM Notes2
                         WHERE (QuoteNum <> 0) AND
                               (QuoteNum not in (SELECT QuoteNum FROM Quotes2))"""
            self.ExecQuery(query)

            print "Check for Quotes without Source Documents:"
            query = """SELECT QuoteNum, QuoteID, CollectNum, SourceDocumentNum
                         FROM Quotes2
                         WHERE (SourceDocumentNum > 0) AND
                               (SourceDocumentNum not in (SELECT DocumentNum FROM Documents2))"""
            self.ExecQuery(query)

            # Known Orphans
            print "Known Quote orphans:  (SourceDocumentNum = 0)", 
            query = """SELECT QuoteNum, QuoteID, CollectNum, SourceDocumentNum
                         FROM Quotes2
                         WHERE SourceDocumentNum = 0"""
            self.DBCursor.execute(query)
            print self.DBCursor.rowcount
            print

            if self.DBCursor.rowcount > 0:
                for Record in self.DBCursor.fetchall():
                    print Record
                print

        print "Check for Clips without Collections:"
        query = 'SELECT ClipNum, ClipID, CollectNum FROM Clips2  WHERE (CollectNum not in (SELECT CollectNum FROM Collections2))'
        self.ExecQuery(query)

        print "Check for Clip Transcripts without Clips:"
        query = 'SELECT TranscriptNum, EpisodeNum, SourceTranscriptNum, ClipNum FROM Transcripts2  WHERE (ClipNum <> 0) AND (ClipNum not in (SELECT ClipNum FROM Clips2))'
        self.ExecQuery(query)

        print "Check for Clip Transcripts without Source Transcripts:"
        query = 'SELECT TranscriptNum, EpisodeNum, SourceTranscriptNum, ClipNum FROM Transcripts2  WHERE (SourceTranscriptNum <> 0) AND (SourceTranscriptNum not in (SELECT TranscriptNum FROM Transcripts2))'
        self.ExecQuery(query)

        # Known Orphans
        print "Known Clip orphans:  (SourceTranscriptNum = 0)", 
        query = 'SELECT TranscriptNum, EpisodeNum, SourceTranscriptNum, ClipNum FROM Transcripts2 WHERE (ClipNum > 0) AND (SourceTranscriptNum = 0)'
        self.DBCursor.execute(query)
        print self.DBCursor.rowcount
        print

        if self.DBCursor.rowcount > 0:
            for Record in self.DBCursor.fetchall():
                print Record
            print

        print "Check for Clips with a length of 0.0"
        query = 'SELECT ClipNum, ClipID, CollectNum FROM Clips2  WHERE ClipStart = ClipStop'
        self.ExecQuery(query)
 
        print "Check for Clip Additional Video Files without Clips:"
        query = 'SELECT AddVidNum, EpisodeNum, ClipNum FROM AdditionalVids2  WHERE (ClipNum <> 0) AND (ClipNum not in (SELECT ClipNum FROM Clips2))'
        self.ExecQuery(query)

        print "Check for Clip Keywords without Clips:"
        query = 'SELECT EpisodeNum, ClipNum, KeywordGroup, Keyword FROM ClipKeywords2  WHERE (ClipNum <> 0) AND (ClipNum not in (SELECT ClipNum FROM Clips2))'
        self.ExecQuery(query)

        print "Check for Clip Notes without Clips:"
        query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, DocumentNum, QuoteNum
                     FROM Notes2
                     WHERE (ClipNum <> 0) AND
                           (ClipNum not in (SELECT ClipNum FROM Clips2))"""
        self.ExecQuery(query)

        if DBVersion >= 260:
            print "Check for Snapshots without Episode:"
            query = 'SELECT SnapshotNum, SnapshotID, EpisodeNum FROM Snapshots2  WHERE (EpisodeNum > 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
            self.ExecQuery(query)

            print "Check for Snapshots without Transcripts:"
            query = 'SELECT SnapshotNum, SnapshotID, TranscriptNum FROM Snapshots2  WHERE (TranscriptNum > 0) AND (TranscriptNum not in (SELECT TranscriptNum FROM Transcripts2))'
            self.ExecQuery(query)

            print "Check for Snapshots without Collections:"
            query = 'SELECT SnapshotNum, SnapshotID, CollectNum FROM Snapshots2  WHERE (CollectNum not in (SELECT CollectNum FROM Collections2))'
            self.ExecQuery(query)

            print "Check for Whole Snapshot Keywords without Snapshots:"
            query = 'SELECT EpisodeNum, ClipNum, SnapshotNum, KeywordGroup, Keyword FROM ClipKeywords2  WHERE (SnapshotNum <> 0) AND (SnapshotNum not in (SELECT SnapshotNum FROM Snapshots2))'
            self.ExecQuery(query)

            print "Check for Snapshot Coding Keywords without Snapshots:"
            query = 'SELECT SnapshotNum, KeywordGroup, Keyword FROM SnapshotKeywords2  WHERE (SnapshotNum not in (SELECT SnapshotNum FROM Snapshots2))'
            self.ExecQuery(query)

            print "Check for Snapshot Keyword Styles without Snapshots:"
            query = 'SELECT SnapshotNum, KeywordGroup, Keyword FROM SnapshotKeywordStyles2  WHERE (SnapshotNum not in (SELECT SnapshotNum FROM Snapshots2))'
            self.ExecQuery(query)

            print "Check for Snapshot Notes without Snapshots:"
            query = """SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum, SnapshotNum, DocumentNum, QuoteNum
                         FROM Notes2
                         WHERE (SnapshotNum <> 0) AND
                         (SnapshotNum not in (SELECT SnapshotNum FROM Snapshots2))"""
            self.ExecQuery(query)

        keywordList = []
        print "Loading Keywords..."
        query = 'SELECT KeywordGroup, Keyword from Keywords2 ORDER BY KeywordGroup, Keyword'
        try:
            self.DBCursor.execute(query)
        except:
            print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
#        print "Count =", self.DBCursor.rowcount
        if self.DBCursor.rowcount > 0:
            Records = self.DBCursor.fetchall()
            for Record in Records:
                keywordList.append((Record[0], Record[1]))

#                print "Adding", Record

        print "Keywords loaded"
        print

        testPassed = True
        self.testsTried += 1
        print "Check for ClipKeywords without Keywords..."
        query = 'SELECT KeywordGroup, Keyword from ClipKeywords2'
        try:
            self.DBCursor.execute(query)
        except:
            print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
#        print "Count =", self.DBCursor.rowcount
        if self.DBCursor.rowcount > 0:
            Records = self.DBCursor.fetchall()
            for Record in Records:
                # Stripping required for Arabic "Transana Users" group!!
                if not((Record[0].strip(), Record[1].strip()) in keywordList):
                    testPassed = False
                    break
        if testPassed:
            self.testsPassed += 1
            print "PASSED"
            print
        else:
            for Record in Records:
                if not((Record[0].strip(), Record[1].strip()) in keywordList):
                    print Record, "not found."
            print


        if DBVersion >= 260:
            snapshotKeywordsList = []
            
            testPassed = True
            self.testsTried += 1
            print "Check for SnapshotKeywords without Keywords..."
            query = 'SELECT KeywordGroup, Keyword, SnapshotNum from SnapshotKeywords2'
            try:
                self.DBCursor.execute(query)
            except:
                print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
    #        print "Count =", self.DBCursor.rowcount
            if self.DBCursor.rowcount > 0:

                Records = self.DBCursor.fetchall()
                for Record in Records:
                    if not((Record[0], Record[1]) in keywordList):
                        testPassed = False

                    snapshotKeywordsList.append((Record[2], Record[0], Record[1]))

            if testPassed:
                self.testsPassed += 1
                print "PASSED"
                print
            else:
                for Record in Records:
                    if not((Record[0], Record[1]) in keywordList):
                        print Record, "not found."
                print

            testPassed = True
            self.testsTried += 1
            print "Check for SnapshotKeywordStyles without Keywords..."
            query = 'SELECT KeywordGroup, Keyword from SnapshotKeywordStyles2'
            try:
                self.DBCursor.execute(query)
            except:
                print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
    #        print "Count =", self.DBCursor.rowcount
            if self.DBCursor.rowcount > 0:
                Records = self.DBCursor.fetchall()
                for Record in Records:
                    if not((Record[0], Record[1]) in keywordList):
                        testPassed = False
                        break
            if testPassed:
                self.testsPassed += 1
                print "PASSED"
                print
            else:
                for Record in Records:
                    if not((Record[0], Record[1]) in keywordList):
                        print Record, "not found."
                print

            testPassed = True
            self.testsTried += 1
            print "Check for SnapshotKeywordStyles without Coding Objects..."
            query = 'SELECT SnapshotNum, KeywordGroup, Keyword from SnapshotKeywordStyles2'
            try:
                self.DBCursor.execute(query)
            except:
                print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
    #        print "Count =", self.DBCursor.rowcount
    #        print
    #        for rec in self.DBCursor.description:
    #            print rec[0], '\t\t',
    #        print
                
            if self.DBCursor.rowcount > 0:
                Records = self.DBCursor.fetchall()

                for Record in Records:

                    if not ((Record[0], Record[1], Record[2]) in snapshotKeywordsList):
                    
                        testPassed = False

                        break

                
            if testPassed:
                self.testsPassed += 1
                print "PASSED"
                print
            else:
                for Record in Records:
                    if not ((Record[0], Record[1], Record[2]) in snapshotKeywordsList):
                        print Record, "not found."
                print

        testPassed = True
        self.testsTried += 1
        print "Check for Sort Order issues in Collections ..."
        query = 'SELECT CollectNum, CollectID FROM Collections2'
        try:
            self.DBCursor.execute(query)
        except:
            print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
#        print "Count =", self.DBCursor.rowcount
#        print
#        for rec in self.DBCursor.description:
#            print rec[0], '\t\t',
#        print
            
        if self.DBCursor.rowcount > 0:
            Records = self.DBCursor.fetchall()

            DBCursor2 = DBConn.cursor()

            for Record in Records:

                SortOrder = {}

                query2 = 'SELECT ClipNum, ClipID, SortOrder FROM Clips2 WHERE CollectNum = %s' % Record[0]

                try:
                    DBCursor2.execute(query2)
                except:
                    print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
#                print "  Count =", DBCursor2.rowcount
#                print

                Records2 = DBCursor2.fetchall()

                for Record2 in Records2:
                    
#                    print "  Clip:", Record2[0], Record2[1], Record2[2]

                    if SortOrder.has_key(Record2[2]):
                        testPassed = False

#                        print '  Collection ', Records[0], Records[1]
                        print '  Clip %d - "%s" has the same Sort Order as %s "%s" in Collection "%s"' % \
                              (Record2[0], Record2[1], SortOrder[Record2[2]][0], SortOrder[Record2[2]][1], Record[1])
                        print

                    else:

                        SortOrder[Record2[2]] = ('Clip', Record2[1])

                if DBVersion >= 270:
                    query2 = 'SELECT QuoteNum, QuoteID, SortOrder FROM Quotes2 WHERE CollectNum = %s' % Record[0]

                    try:
                        DBCursor2.execute(query2)
                    except:
                        print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
    #                print "  Count =", DBCursor2.rowcount
    #                print

                    Records2 = DBCursor2.fetchall()

                    for Record2 in Records2:
                        
    #                    print "  Quote:", Record2[0], Record2[1], Record2[2]
                    
                        if SortOrder.has_key(Record2[2]):
                            testPassed = False

                            print '  Quote "%s" has the same Sort Order as %s "%s" in Collection "%s"' % \
                                  (Record2[1], SortOrder[Record2[2]][0], SortOrder[Record2[2]][1], Record[1])
                            print

                        else:

                            SortOrder[Record2[2]] = ('Quote', Record2[1])

                if DBVersion >= 260:
                    query2 = 'SELECT SnapshotNum, SnapshotID, SortOrder FROM Snapshots2 WHERE CollectNum = %s' % Record[0]

                    try:
                        DBCursor2.execute(query2)
                    except:
                        print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
    #                print "  Count =", DBCursor2.rowcount
    #                print

                    Records2 = DBCursor2.fetchall()

                    for Record2 in Records2:
                        
    #                    print "  Snapshot:", Record2[0], Record2[1], Record2[2]
                    
                        if SortOrder.has_key(Record2[2]):
                            testPassed = False

                            print '  Snapshot "%s" has the same Sort Order as %s "%s" in Collection "%s"' % \
                                  (Record2[1], SortOrder[Record2[2]][0], SortOrder[Record2[2]][1], Record[1])
                            print

                        else:

                            SortOrder[Record2[2]] = ('Snapshot', Record2[1])

        if testPassed:
            self.testsPassed += 1
            print "PASSED"
            print


        self.DBCursor.close()
        DBConn.close()
        MySQLdb.server_end()

        print "SUMMARY:  Tests conducted:  %d   Tests passed:  %d   Tests failed:  %d" % (self.testsTried, self.testsPassed, self.testsTried - self.testsPassed)

    def ExecQuery(self, query):
        self.testsTried += 1
        try:
            self.DBCursor.execute(query)

            if self.DBCursor.rowcount == 0:
                print "PASSED"
                self.testsPassed += 1
                print
            else:
                
                print "FAILED!!!  ******************************************"
                print
                
    #        print "Count =", self.DBCursor.rowcount
            if self.DBCursor.rowcount > 0:
                for rec in self.DBCursor.description:
                    print rec[0],
                print
                Records = self.DBCursor.fetchall()
                for Record in Records:
                    print Record
                print
            print
        except:
            print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
            print


# run the app
OrphanCheck(None, -1, "Orphan Check")
