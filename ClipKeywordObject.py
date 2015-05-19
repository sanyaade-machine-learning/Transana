# Copyright (C) 2003-2012 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the ClipKeyword class.

   The ClipKeyword class is designed to retain all important information for a KeywordGroup:Keyword
   pair as used within Episode and Clip objects.  The primary differences between this and the
   Keyword class is that this object retains information about whether this KeywordGroup:Keyword
   pair instance serves as a Keyword Example, and the ClipKeyword does not concern itself with
   the Keyword Definition.

   Usage:
     ClipKeyword(keywordGroup, keyword, episodeNum = 0, clipNum = 0, example = 0)
       keywordGroup and keyword are mandatory.  episodeNum and clipNum will only be useful
       outside of the context of an Episode or Clip, although this context is exactly where
       I anticipate this object will be mostly used.  example defaults to 0, as few
       ClipKeywords are examples.

   Properties:
     keywordGroup   Keyword Group
     keyword        Keyword
     keywordPair    read-only KeywordGroup:Keyword string
     episodeNum     Episode Number (0 for Clip)
     clipNum        Clip Number (0 for Episode)
     example        0 if not a Keyword Example instance, 1 if it is

   Methods:
     None
"""

import types
import DBInterface
# import Transana's Globals
import TransanaGlobal

# import wxPython only for unicode testing
import wx

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'


class ClipKeyword(object):
    """ The ClipKeyword Object holds all data for a Keyword.  This can be held in lists in the Episode and Clip objects. """

    def __init__(self, keywordGroup, keyword, episodeNum = 0, clipNum = 0, example=0):
        """ Initialize the ClipKeyword Object.  keywordGroup and keyword are required.  Either an episodeNum
            or a clipNum must be specified. """
        # Store all data values passed in on initialization
        self.keywordGroup = keywordGroup
        self.keyword = keyword
        self.episodeNum = episodeNum
        self.clipNum = clipNum
        self.example = example

    def __repr__(self):
        """ Provides a String Representation of the ClipKeyword Object. """
        str = 'Clip Keyword:\n'
        str = str + '  episodeNum = %s\n' % self.episodeNum
        str = str + '  clipNum = %s\n' % self.clipNum
        str = str + '  keywordGroup = %s\n' % self.keywordGroup
        str = str + '  keyword = %s\n' % self.keyword
        str = str + '  keywordPair = %s\n' % self.keywordPair
        str = str + '  example = %s (%s)\n\n' % (self.example, type(self.example))
        return str.encode('utf8')

    def db_save(self):
        """ Saves ClipKeyword record to Database """
        # NOTE:  This routine, at present, is ONLY used by the Database Import routine.
        #        Therefore, it does no checking for duplicate records.  If you want to
        #        use it for other purposes, you probably have to make it smarter!

        if 'unicode' in wx.PlatformInfo:
            keywordGroup = self.keywordGroup.encode(TransanaGlobal.encoding)
            keyword = self.keyword.encode(TransanaGlobal.encoding)
        else:
            keywordGroup = self.keywordGroup
            keyword = self.keyword
        dbCursor = DBInterface.get_db().cursor()
        SQLText = """ INSERT INTO ClipKeywords2
                        (EpisodeNum, ClipNum, KeywordGroup, Keyword, Example)
                      VALUES
                        (%s, %s, %s, %s, %s) """
        values = (self.episodeNum, self.clipNum, keywordGroup, keyword, self.example)
        dbCursor.execute(SQLText, values)
        dbCursor.close()
    
    # Define Property getters and setters
    # Keyword Group Property
    def _getKeywordGroup(self):
        return self._keywordGroup
    def _setKeywordGroup(self, keywordGroup):
        self._keywordGroup = keywordGroup
    def _delKeywordGroup(self):
        self._keywordGroup = ''

    # Keyword Property
    def _getKeyword(self):
        return self._keyword
    def _setKeyword(self, keyword):
        self._keyword = keyword
    def _delKeyword(self):
        self._keyword = ''

    # read-only Keyword Pair Property
    def _getKeywordPair(self):
        return self._keywordGroup + ' : ' + self._keyword

    # Episode Number Property
    def _getEpisodeNum(self):
        return self._episodeNum
    def _setEpisodeNum(self, episodeNum):
        self._episodeNum = episodeNum
    def _delEpisodeNum(self):
        self._episodeNum = 0

    # Clip Number Property
    def _getClipNum(self):
        return self._clipNum
    def _setClipNum(self, clipNum):
        self._clipNum = clipNum
    def _delClipNum(self):
        self._clipNum = 0

    # Example Property
    def _getExample(self):
        return self._example
    def _setExample(self, example):
        if isinstance(example, types.StringTypes):
            try:
                example = int(example)
            except:
                example = 0
        self._example = example
    def _delExample(self):
        self._example = 0


    # Define Object Properties
    keywordGroup = property(_getKeywordGroup, _setKeywordGroup, _delKeywordGroup, "Keyword Group.")
    keyword = property(_getKeyword, _setKeyword, _delKeyword, "Keyword.")
    # keywordPair is a read-only property that returns the "KeywordGroup:Keyword" string
    keywordPair = property(_getKeywordPair, doc="Read-only KWG:KW Pair")
    episodeNum = property(_getEpisodeNum, _setEpisodeNum, _delEpisodeNum, "Episode Number.")
    clipNum = property(_getClipNum, _setClipNum, _delClipNum, "Clip Number.")
    example = property(_getExample, _setExample, _delExample, "Keyword Example.")
