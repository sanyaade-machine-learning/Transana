# Copyright (C) 2004 - 2012  The Board of Regents of the University of Wisconsin System 
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

# import the Python os module
import os
# import the Python string module
import string
# import wxPython
import wx
# Import Transana's Constants
import TransanaConstants
if TransanaConstants.USESRTC:
    import wx.richtext as richtext
else:
    # import the wxPython StyledTextCtrl (for some constant definitions)
    from wx import stc
# import Transana's Database interface
import DBInterface
# import Transana's Dialog Boxes
import Dialogs
# import Transana's Filter Dialog Box
import FilterDialog
# Import Transana's keyword object
import KeywordObject as Keyword
# Import Transana's Text Report infrastructure
import TextReport
# Import Transana's Global variables
import TransanaGlobal
# Import Transana's Printout Class (for DPI definition)
import TranscriptPrintoutClass


class KeywordSummaryReport(wx.Object):
    """ This class creates and displays the Keyword Summary Report """
    def __init__(self, keywordGroupName=None):
        """ Create the Keyword Summary Report.  If the keywordGroupName parameter is specified, the report scope
            is a single (specified) keyword group.  Otherwise, it should encompass all Keyword Groups. """
        # Remember the keyword group name passed if, if any
        self.keywordGroupName = keywordGroupName
        # Specify the Report Title
        self.title = unicode(_("Keyword Summary Report"), 'utf8')
        # If no Keyword Group Name is specified ...
        if self.keywordGroupName == None:
            # ... then we want to include the filterMethod specification.  This is the Keyword Summary Report that
            # spans Keyword Groups, and we may want to filter out some keyword groups.
            self.report = TextReport.TextReport(None, title=self.title, displayMethod=self.OnDisplay,
                                                filterMethod=self.OnFilter, helpContext="Keyword Summary Report")
        # If a Keyword Group Name is specified ...
        else:
            # ... then we do NOT want to include the filterMethod specification.  This is the Keyword Summary Report
            # that specifies a single Keyword Group, so no filtering is needed.
            self.report = TextReport.TextReport(None, title=self.title, displayMethod=self.OnDisplay,
                                                helpContext="Keyword Summary Report")
        # Initialize the Keyword Group list for the Filter Dialog.  Start with an empty list.
        self.keywordGroupFilterList = []
        # Get a list of all Keyword Groups
        kwgList = DBInterface.list_of_keyword_groups()
        # Iterate through the list of Keyword Groups ...
        for kwg in kwgList:
            # ... and add them to the Keyword Group List, along with a boolean suggesting they should be displayed initially
            self.keywordGroupFilterList.append((kwg, True))
        # Trigger the ReportText method that causes the report to be displayed.
        self.report.CallDisplay()

    def OnDisplay(self, reportText):
        """ This method, required by TextReport, populates the TextReport.  The reportText parameter is
            the wxSTC control from the TextReport object.  It needs to be in the report parent because
            the TextReport doesn't know anything about the actual data.  """
        # Make the control writable
        reportText.SetReadOnly(False)

        # If we're using the RichTextCtrl ...
        if TransanaConstants.USESRTC:
            # ... Set the Style for the Heading
            reportText.SetTxtStyle(fontFace='Courier New', fontSize=16, fontBold=True, fontUnderline=True)
            # Add the Title to the page
            reportText.WriteText(self.title)
            # Set report margins, the left and right margins to 0.  The RichTextPrinting infrastructure handles that!
            # Center the title, and add spacing after.
            reportText.SetTxtStyle(parLeftIndent = 0, parRightIndent = 0, 
                                   parAlign=wx.TEXT_ALIGNMENT_CENTER, parSpacingAfter = 20)
            # End the paragraph
            reportText.Newline()
        # If we're using the Styled Text Ctrl ...
        else:
            # ... Set the font for the Report Title
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

        # If a Keyword Group Name is passed in ...
        if self.keywordGroupName != None:
            # ... build the prompt for the subtitle ...
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Keyword Group: %s"), 'utf8')
            else:
                prompt = _("Keyword Group: %s")
            # ... add the keyword group name to the subtitle ...
            self.subtitle = prompt % self.keywordGroupName
            # ... and add the passed in keyword group as the only one to be included in the report.
            keywordGroupList = [self.keywordGroupName]
        # If no Keyword Group Name is passed in ...
        else:
            # ... we don't need a subtitle ...
            self.subtitle = ''
            # and we need a list of all Keyword Groups as the list of groups to be included in the Report.
            keywordGroupList = DBInterface.list_of_keyword_groups()

        # If a subtitle is defined ...
        if self.subtitle != '':
            # If we're using the RichTextCtrl ...
            if TransanaConstants.USESRTC:
                # ... set the subtitle font
                reportText.SetTxtStyle(fontSize=10, fontBold=False, fontUnderline=False)
                # Add the subtitle to the page
                reportText.WriteText(self.subtitle)
                # Finish the paragraph
                reportText.Newline()
            # If we're using the Styled Text Ctrl ...
            else:
                # ... set the font for the subtitle ...
                reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                # ... get the style specifier for that font ...
                style = reportText.GetStyleAccessor("size:10,face:Courier New,fore:#000000,back:#ffffff")
                # ... get the spaces needed to center the subtitle ...
                centerSpacer = self.report.GetCenterSpacer(style, self.subtitle)
                # ... and insert the spacer and the subtitle.
                reportText.InsertStyledText('\n' + centerSpacer + self.subtitle)

        # If we're using the RichTextCtrl ...
        if TransanaConstants.USESRTC:
            # Add a blank line
            reportText.Newline()
        # If we're using the Styled Text Ctrl ...
        else:
            # Skip a couple of lines.
            reportText.InsertStyledText('\n\n')

        # Initialize an empty list for Keyword Examples
        keywordExamplesList = []
        # Iterate through the list of all keyword examples ...
        for KWE in DBInterface.list_of_keyword_examples():
            # ... adding the keywords to the KeywordExamples List
            keywordExamplesList.append((KWE[2], KWE[3]))

        # Get the graphic for the Keyword Example indicator
        kweGraphic = wx.Image(os.path.join(TransanaGlobal.programDir, "images", "Clip16.xpm"), wx.BITMAP_TYPE_XPM)

        # Iterate through the list of Keyword Groups
        for keywordGroup in keywordGroupList:
            # Check to see if the keyword group is supposed to be displayed by looking for it in the Filter list
            if (keywordGroup, True) in self.keywordGroupFilterList:
                # If we're using the RichTextCtrl ...
                if TransanaConstants.USESRTC:
                    # ... Set the formatting for the report, including turning off previous formatting
                    reportText.SetTxtStyle(fontSize = 12, fontBold = True, fontUnderline = False)
                    
                    # Add the Keyword Group to the Report
                    reportText.WriteText(keywordGroup)
                    # ... Set the formatting for the report, including turning off previous formatting
                    reportText.SetTxtStyle(parAlign = wx.TEXT_ALIGNMENT_LEFT,
                                           parLeftIndent = 0,
                                           parSpacingBefore = 30, parSpacingAfter = 0)
                    # End the paragraph
                    reportText.Newline()
                # If we're using the Styled Text Ctrl ...
                else:
                    # Use the Keyword Group name as a Heading.  First, set the font.
                    reportText.SetFont('Courier New', 12, 0x000000, 0xFFFFFF)
                    # ... and make it bold.
                    reportText.SetBold(True)
                    # Add the Keyword Group to the report
                    reportText.InsertStyledText("%s\n" % keywordGroup)
                    # Turn bold off
                    reportText.SetBold(False)

                # Get the list of Keywords defined for that group
                keywordList = DBInterface.list_of_keywords_by_group(keywordGroup)
                # Iterate through the list of Keywords
                for keyword in keywordList:
                    # Load the Keyword object
                    keywordObject = Keyword.Keyword(keywordGroup, keyword)

                    # If we're using the RichTextCtrl ...
                    if TransanaConstants.USESRTC:
                        # if there's NO keyword definition ...
                        if keywordObject.definition.strip() == '':
                            # ... then we want paragraph spacing after THIS paragraph
                            parSpacingAfter = 20
                        # If there IS a keyword definition ...
                        else:
                            # ... we DON'T want paragraph spacing after yet.
                            parSpacingAfter = 0
                        # .25 inch left indent
                        reportText.SetTxtStyle(fontSize = 12, fontBold = False)
                        
                        # Write the Keyword
                        reportText.WriteText(keyword)

                        # .25 inch left indent
                        reportText.SetTxtStyle(parLeftIndent = 63,
                                               parSpacingBefore = 0,
                                               parSpacingAfter = parSpacingAfter)

                        # Check to see if this keyword has a Keyword Example
                        if (keywordGroup, keyword) in keywordExamplesList:
                            # If so, add the Keyword Example Indicator Graphic
                            reportText.WriteText(' ')
                            reportText.WriteImage(kweGraphic)
                            
                        # Finish the paragraph
                        reportText.Newline()

                        # If there IS a paragraph definition ...
                        if keywordObject.definition.strip() != '':
                            # Reduce font size, .5 inch left indent, have paragraph spacing after
                            reportText.SetTxtStyle(fontSize = 10,
                                                   parLeftIndent = 126, parRightIndent = 0,
                                                   parSpacingAfter = 20)

                            # Strip white space out of the definition.
                            reportText.WriteText(keywordObject.definition.strip())

                            # Finish the paragraph
                            reportText.Newline()
                    # If we're using the Styled Text Ctrl ...
                    else:
                        # Set the font for the Keyword itself
                        reportText.SetFont('Courier New', 12, 0x000000, 0xFFFFFF)
                        # Add the Keyword to the report
                        reportText.InsertStyledText('  %s\n' % keyword)

                        # Add the Keyword Definition to the Report
                        # Keyword Definitions can have line breaks embedded in them.  If so, break them up into separate lines here!
                        definitionLines = string.split(keywordObject.definition, '\n')
                        # Set the font for the Keyword Definitions
                        reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
                        # ... get the style specifier for that font ...
                        style = reportText.GetStyleAccessor("size:10,face:Courier New,fore:#000000,back:#ffffff")
                        # Add each line of the keyword definition to the report.
                        for ln in definitionLines:
                            # Check for blank lines.
                            if ln.strip() <> '':
                                # Add some blank space to indent the definition
                                reportText.InsertStyledText('    ')
                                # Divide the line up into individual words for word wrapping
                                words = ln.split()
                                # Initialize the line length to the spaces
                                lineLength = reportText.TextWidth(style, '    ')
                                # Iterate through the words in the definition.
                                for word in words:
                                    # Add the width of the word to the line length
                                    lineLength += reportText.TextWidth(style, '%s ' % word)
                                    # If the line length exceeds the width of the control (including margins and 1/4 inch for good measure) ...
                                    if lineLength > reportText.GetSizeTuple()[0] - int(TranscriptPrintoutClass.DPI * 2.25):
                                        
    #                                    print 'KeywordSummaryReport.OnDisplay(): ', word, lineLength, reportText.GetSizeTuple()[0] - (TranscriptPrintoutClass.DPI * 2), reportText.GetSizeTuple()[0], TranscriptPrintoutClass.DPI
                                        
                                        # ... reset the line length to the indent plus the width of the word
                                        lineLength = reportText.TextWidth(style, '    %s ' % word)
                                        # add a line break to the report
                                        reportText.InsertStyledText('\n    ')
                                    # Add the word to the report, followed by a space.
                                    reportText.InsertStyledText('%s ' % word)
                            # If we have a non-blank definition ...
                            if definitionLines != [u'']:
                                # ... add a line break to the report to follow that definition
                                reportText.InsertStyledText('\n')
                        # Add a blank line after each definition
                        reportText.InsertStyledText('\n')

        # Once we're done constructing the report, make it Read Only
        reportText.SetReadOnly(True)

    def OnFilter(self, event):
        """ This method, required by TextReport, implements the call to the Filter Dialog.  It needs to be
            in the report parent because the TextReport doesn't know the appropriate filter parameters. """
        # Define the Filter Dialog.  We need reportType 9 to identify the Keyword Summary Report and we
        # need only the Keyword Group Filter for this report.
        dlgFilter = FilterDialog.FilterDialog(self.report, -1, self.title, reportType=9,
                                              keywordGroupFilter=True)
        # Populate the Filter Dialog with the Keyword Group Filter list
        dlgFilter.SetKeywordGroups(self.keywordGroupFilterList)
        # If the filter is defined and accepted by the user ...
        if dlgFilter.ShowModal() == wx.ID_OK:
            # ... get the filter data ...
            self.keywordGroupFilterList = dlgFilter.GetKeywordGroups()
            # ... and signal the TextReport that the filter is to be applied.
            return True
        # If the filter is cancelled by the user ...
        else:
            # ... signal the TextReport that the filter is NOT to be applied.
            return False
