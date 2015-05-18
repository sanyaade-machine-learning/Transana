# Copyright (C) 2004 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Keyword Summary Report. """

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

import string
import wx
import TransanaGlobal
import DBInterface
import ReportPrintoutClass
import Keyword

class KeywordSummaryReport(wx.Object):
    """ This class creates and displays the Keyword Summary Report """
    def __init__(self, keywordGroupName=None):
        # Specify the Report Title
        self.title = _("Keyword Summary Report")
        # If a Keyword Group Name is passed in, add a subtitle and use it as the Report's keywordGroupList.
        if keywordGroupName != None:
            self.subtitle = _("Keyword Group: %s") % keywordGroupName
            keywordGroupList = [keywordGroupName]
        # If no Keyword Group Name is passed in, get a list of all Keyword Groups as the Report's keywordGroupList.
        else:
            self.subtitle = ''
            keywordGroupList = DBInterface.list_of_keyword_groups()
        
        # Initialize and fill the initial data structure that will be turned into the report
        self.data = []

        # Iterate through the list of Keyword Groups
        for keywordGroup in keywordGroupList:
            # Use the Keyword Group name as a Heading
            self.data.append((('Heading', keywordGroup),))
            # Get the list of Keywords defined for that group
            keywordList = DBInterface.list_of_keywords_by_group(keywordGroup)
            # Iterate through the list of Keywords
            for keyword in keywordList:
                # Use the Keyword name as a Subheading
                self.data.append((('Subheading', keyword),))
                # Load the Keyword object
                keywordObject = Keyword.Keyword(keywordGroup, keyword)
                # Add the Keyword Definition to the Report
                # Keyword Definitions can have line breaks embedded in them.  If so, break them up into separate lines here!
                definitionLines = string.split(keywordObject.definition, '\n')
                # Add each line of the keyword definition to the report using the "Subtext" style
                for ln in definitionLines:
                    self.data.append((('Subtext', ln),))
                # Add a blank line after each definition
                self.data.append((('Normal', ''),))

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
            self.SetStatusText(_("Print Preview Problem"))
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
