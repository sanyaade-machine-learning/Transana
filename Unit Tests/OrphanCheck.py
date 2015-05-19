""" This program explores database records for Transana """

import MySQLdb

import os, sys
import wx

class DataExplorer(object):
    """ This is the main window for the Data Explorer """
    def __init__(self, parent, ID, title):
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
        
        # Get all Series Records
        #  Create a connection to the database

        pw = 'msp'
        
        DBConn = MySQLdb.Connection(host=hostName, db=dbName, user='DavidW', passwd=pw, port=3306)
        #  Create a cursor and execute the appropriate query
        self.DBCursor = DBConn.cursor()

##        print "Databases THIS USER can see:"
##        query = 'SHOW databases'
##        self.DBCursor.execute(query)
##        print "Count =", self.DBCursor.rowcount
##        if self.DBCursor.rowcount > 0:
##            for rec in self.DBCursor.description:
##                print rec[0],
##            print
##            Records = self.DBCursor.fetchall()
##            for Record in Records:
##                print Record
##            print
##        print
##
        print "Check for Series Notes without Series:"
        query = 'SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum FROM Notes2  WHERE (SeriesNum <> 0) AND (SeriesNum not in (SELECT SeriesNum FROM Series2))'
        self.ExecQuery(query)

        print "Check for Episodes without Series:"
        query = 'SELECT EpisodeNum, EpisodeID, SeriesNum FROM Episodes2  WHERE (SeriesNum not in (SELECT SeriesNum FROM Series2))'
        self.ExecQuery(query)

        print "Check for Episode Transcripts without Episodes:"
        query = 'SELECT TranscriptNum, EpisodeNum, SourceTranscriptNum, ClipNum FROM Transcripts2  WHERE (ClipNum = 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
        self.ExecQuery(query)

        print "Check for Episode Additional Video Files without Episodes:"
        query = 'SELECT AddVidNum, EpisodeNum, ClipNum FROM AdditionalVids2  WHERE (ClipNum = 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
        self.ExecQuery(query)

        print "Check for Episode Keywords without Episodes:"
        query = 'SELECT EpisodeNum, ClipNum, KeywordGroup, Keyword FROM ClipKeywords2  WHERE (ClipNum = 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
        self.ExecQuery(query)

        print "Check for Episode Notes without Episodes:"
        query = 'SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum FROM Notes2  WHERE (EpisodeNum <> 0) AND (EpisodeNum not in (SELECT EpisodeNum FROM Episodes2))'
        self.ExecQuery(query)

        print "Check for Transcript Notes without Transcripts:"
        query = 'SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum FROM Notes2  WHERE (TranscriptNum <> 0) AND (TranscriptNum not in (SELECT TranscriptNum FROM Transcripts2))'
        self.ExecQuery(query)

        print "Check for Nested Collections without Parents:"
        query = 'SELECT CollectNum, CollectID, ParentCollectNum FROM Collections2  WHERE (ParentCollectNum <> 0) AND (ParentCollectNum not in (SELECT CollectNum FROM Collections2))'
        self.ExecQuery(query)

        print "Check for Collection Notes without Collections:"
        query = 'SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum FROM Notes2  WHERE (CollectNum <> 0) AND (CollectNum not in (SELECT CollectNum FROM Collections2))'
        self.ExecQuery(query)

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
        print "Known orphans:  ", 
        query = 'SELECT TranscriptNum, EpisodeNum, SourceTranscriptNum, ClipNum FROM Transcripts2 WHERE (ClipNum > 0) AND (SourceTranscriptNum = 0)'
        self.DBCursor.execute(query)
        print self.DBCursor.rowcount
        print

        print "Check for Clip Additional Video Files without Clips:"
        query = 'SELECT AddVidNum, EpisodeNum, ClipNum FROM AdditionalVids2  WHERE (ClipNum <> 0) AND (ClipNum not in (SELECT ClipNum FROM Clips2))'
        self.ExecQuery(query)

        print "Check for Clip Keywords without Clips:"
        query = 'SELECT EpisodeNum, ClipNum, KeywordGroup, Keyword FROM ClipKeywords2  WHERE (ClipNum <> 0) AND (ClipNum not in (SELECT ClipNum FROM Clips2))'
        self.ExecQuery(query)

        print "Check for Clip Notes without Clips:"
        query = 'SELECT NoteNum, NoteID, SeriesNum, EpisodeNum, TranscriptNum, CollectNum, ClipNum FROM Notes2  WHERE (ClipNum <> 0) AND (ClipNum not in (SELECT ClipNum FROM Clips2))'
        self.ExecQuery(query)

        self.DBCursor.close()
        DBConn.close()
        MySQLdb.server_end()

        print "SUMMARY:  Tests conducted:  %d   Tests passed:  %d   Tests failed:  %d" % (self.testsTried, self.testsPassed, self.testsTried - self.testsPassed)

    def ExecQuery(self, query):
        self.testsTried += 1
        try:
            self.DBCursor.execute(query)

        except:
            print "Exception!", sys.exc_info()[0], sys.exc_info()[1]
        if self.DBCursor.rowcount == 0:
            print "PASSED, ",
            self.testsPassed += 1
        else:
            
            print "FAILED!!!  ******************************************"
        print "Count =", self.DBCursor.rowcount
        if self.DBCursor.rowcount > 0:
            for rec in self.DBCursor.description:
                print rec[0],
            print
            Records = self.DBCursor.fetchall()
            for Record in Records:
                print Record
            print
        print


# run the app
DataExplorer(None, -1, "Data Explorer")
