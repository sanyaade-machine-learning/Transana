#Copyright (C) 2003 - 2007  The Board of Regents of the University of Wisconsin System
#
#This program is free software; you can redistribute it and/or
#modify it under the terms of the GNU General Public License
#as published by the Free Software Foundation; either version 2
#of the License, or (at your option) any later version.
#
#This program is distributed in the hope that it will be useful,
#but WITHOUT ANY WARRANTY; without even the implied warranty of
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#GNU General Public License for more details.
#
#You should have received a copy of the GNU General Public License
#along with this program; if not, write to the Free Software
#Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

""" This class implements a Print Preview and Printing mechanism for RTF-based Transcripts within
    Transana and TEXT-based Notes, based on the wxPython Print Framework. """

__author__ = "David K. Woods <dwoods@wcer.wisc.edu>, Nate Case"

# This unit is designed to implement Print Preview and Print mechanisms for Transana's Transcript Printing and Note
# printing.  It is designed to receive a Transcript object or Note text, depending on
# whether we are printing a Transcript or a Note.  (Yeah, it's not consistent, but I thought being able to handle
# ANY plain text might be helpful elsewhere.)  This unit should handle all pagination automatically.
#
# To use it, do the following:
#
#  - Call the PrepareData function in this unit.  Parameters are:
#
#        a wxPrintData Object
#        the Title of your report, a String
#        a Transana Transcript Object OR a Transana Note's plain text
#        optionally, a subtitle for the report
#
#    This function returns a tuple of data elements required in the next step.  They are:
#
#        a wxBitmap object that is blank, but that is the correct size for the paper type send in the wxPrintData Object
#        a data structure where the data sent in has been divided up into pages of lines.
#
#    The call looks like this:
#
#        (self.graphic, self.pageData) = PrintoutClass.PrepareData(self.printData, self.title, self.data, self.subtitle)
#
#    It is necessary to do this BEFORE invoking the wxPython Print Framework objects (wxPreviewFrame) because
#    there is, I think, a bug in wxPython (or perhaps wxWindows) where the GetPageInfo method, which tells the
#    print framework how many pages are in the output, gets called before the OnPreparePrinting method, where
#    you are supposed to be able to manipulate your data and determine the number of pages in your report.
#    (See the wxWindows documentation for wxPrintout's OnPreparePrinting method.)  The wxBitmap object, although
#    empty, is necessary so that the proper Scaling Factor can be calculated for the wxPrintout object.
#
#  - Finally, you create a MyPrintout Object, passing your report title, your graphic object and your rearranged data
#    object returned by the PrepareData call above, and your optional subtitle.  This MyPrintout Object can then be
#    sent to the wxPrintPreview object or to a wxPrinter object's Print method.  The whole thing, for the data object
#    described above, looks something like this:
#
#    def OnPrintPreview(self, event):
#        (self.graphic, self.pageData) = PrintoutClass.PrepareData(self.printData, self.title, self.data, self.subtitle)
#
#        printout = PrintoutClass.MyPrintout(self.title, self.graphic, self.pageData, subtitle)
#        printout2 = PrintoutClass.MyPrintout(self.title, self.graphic, self.pageData, subtitle)
#
#        self.preview = wx.PrintPreview(printout, printout2, self.printData)
#        if not self.preview.Ok():
#            self.SetStatusText("Print Preview Problem")
#            return
#        frame2 = wx.PreviewFrame(self.preview, self, "Print Preview")
#        frame2.Initialize()
#        frame2.Show(True)

# import wxPython
import wx
# import the Python String class
import string
# import the Python sys module
import sys
# import the Transana Database Interface
import DBInterface
# import Transana Dialogs for Error Dialog
import Dialogs
# import the Transana Miscellaneous Functions
import Misc
# import the Transana Transcript Object
import Transcript
# import the Transana Collection Object
#import Collection
# import the Transana Clip Object
#import Clip
#import the Transana RTF Parser class
RTFModulePath = "rtf"
if sys.path.count(RTFModulePath) == 0:
    sys.path.append(RTFModulePath)  # Add to path if not already there
import RTFParser
# Import Transana's Globals
import TransanaGlobal

# For pages to be sized and proportioned correctly, we need different DPI scaling factors on different platforms.
if "__WXMAC__" in wx.PlatformInfo:
    DPI = 72
else:
    DPI = 96

#----------------------------------------------------------------------

# NOTE:  At some point, these values should become configurable via Page Formatting in the Transcript Editor
xMargin = DPI    # 1 inch horizontal margin by default
yMargin = DPI    # 1 inch vertical margin by default

DEFAULT_FONT_SIZE = 12
STYLE_NONE     = 0
STYLE_CENTER   = 1
STYLE_DRAW_BOX = 2

def GetDefaultFont():
    """Return a default font."""
    font = wx.TheFontList.FindOrCreateFont(DEFAULT_FONT_SIZE, wx.ROMAN, wx.NORMAL, wx.NORMAL, 0)
    color = 0x000000  # wx.BLACK
    style = [STYLE_NONE]
    return (font, color, style)

def GetTitleFont():
    """Return a default font."""
    font = wx.TheFontList.FindOrCreateFont(DEFAULT_FONT_SIZE * 2, wx.ROMAN, wx.NORMAL, wx.BOLD, 1)
    color = 0x000000  # wx.BLACK
    style = [STYLE_CENTER]
    return (font, color, style)


def RTFDocAttrTowxFont(docattr):
    """Convert a DocAttribute object from RTFParser into a wxFont."""
    
    fonttokens = docattr.font.split()

    if len(fonttokens) > 0 and fonttokens[-1].isdigit():
        # Font string follows format of: "Face name SIZE"
        ptsize = int(fonttokens[-1])
        face = string.join(fonttokens[:-1])
    else:
        # Assume that the .font attribute is just a face name
        face = docattr.font
        ptsize = docattr.fontsize

    if docattr.italic:
        fontstyle = wx.ITALIC
    else:
        fontstyle = wx.NORMAL

    if docattr.bold:
        weight = wx.BOLD
    else:
        weight = wx.NORMAL
        
    font = wx.TheFontList.FindOrCreateFont(ptsize, wx.DEFAULT, fontstyle, weight, docattr.underline, facename=face)
    color = docattr.fg
    style = [STYLE_NONE]
    return (font, color, style)
    
#----------------------------------------------------------------------

def ProcessRTF(dc, sizeX, sizeY, inText, pageData, thisPageData, datLines, xPos, yPos, yInc, titleHeight):
    """ Process Rich Text Format data for PrepareData() """
    # Parse the RTF transcript data into a usable data structure
    parser = RTFParser.RTFParser()
    parser.buf = inText
    parser.read_stream()
    # Now construct the data structure used in the for loop below, which expects a sequence of
    # tuples of the following format: ((wxFont, wxColor, style), text)
    #
    # where the first element is a wxFont which defines the style of the text
    # the second element is the text that uses this style
    data = []
    # Initialize lineWidth.  (Issue 230)
    lineWidth = 0
    (cur_font, cur_color, cur_style) = GetDefaultFont()
    for obj in parser.stream:
        if obj.attr:
            (cur_font, cur_color, cur_style) = RTFDocAttrTowxFont(obj.attr)
        else:
            newitem = ((cur_font, cur_color, cur_style), obj.text)
            data.append(newitem)
    # Initialize the string variable used in constructing text elements
    tempLine = ''
    # Iterate through all the data sent in, one "line" at a time.
    # NOTE:  "line" here refers to a line in the incoming data structure, which
    #        probably does not correspond with a printed line on the printout.
    for line in data:
        # each line is made up of a font spec and the associated text.
        (fontSpec, text) = line
        # Find a Line Break character, if there is one.
        breakPos = text.find('\n')

        # Set the Device Context font to the identified fontSpec
        dc.SetFont(fontSpec[0])

        # There is a bug in wxPython.  wx.ColourRGB() transposes Red and Blue.  This hack fixes it!
        color = wx.ColourRGB(fontSpec[1])
        
        rgbValue = (color.Red() << 16) | (color.Green() << 8) | color.Blue()
        
        dc.SetTextForeground(wx.ColourRGB(rgbValue))

        # First, let's see if this starts with a line break
        while breakPos == 0:

            # If so, add it to the datLines structure
            datLines = datLines + ((fontSpec, '\n'),)
            # Add the current datLines structure to the page
            thisPageData.append(datLines)

            (dummy, yInc) = dc.GetTextExtent('Xy')

            # Reset datLines and tempLine
            datLines = ()
            tempLine = ''
            # Reset horizontal position to the horizontal margin
            xPos = xMargin
            lineWidth = 0
            # Increment the vertical position by the height of the line plus a blank line
            yPos = yPos + yInc                 # + fontSpec.GetPointSize() + 6
            # Reset yInc
            yInc = 0

            # It is possible there are additional line breaks, so let's see if there are more to process.
            text = text[1:]
            breakPos = text.find('\n')


        # Now let's check for line breaks elsewhere in the line.  Actually, this shouldn't happen!
#        if breakPos > 0:
#            print "Line Break inside the line, not at the beginning"

        # Break the line into words at whitespace breaks
        if ('unicode' in wx.PlatformInfo) and (type(text).__name__ == 'str'):
            words = []
            words.append(unicode(text, TransanaGlobal.encoding))
        else:

            # This "text.split()" call, necessary for line breaks, causes loss of whitespace.
            # We have to do some head stands to avoid it.
            words = text.split()
            
        # We need to retain leading whitespace
        if (len(text) > 0) and (text[0] == ' ') and (len(words) > 0):
            # This syntax captures all leading whitespace.
            words[0] = text[:text.find(words[0])] + words[0]
        # We also need to retain trailing whitespace
        if (len(text) > 0) and (text[len(text)-1] == ' ') and (len(words) > 0):
            words[len(words)-1] = words[len(words)-1] + ' '
        # Check for text that is ONLY whitespace, as this was getting lost if formatting was applied on both sides
        if (len(text) > 0) and (len(words) == 0):
            words = []
            words.append(text)

        # Iterate through the words
        for word in words:
            text = text[len(word):]

            while (len(text) > 0) and (string.whitespace.find(text[0]) > -1):
                word = word + text[0]
                text = text[1:]

            # Determine the line width if we add the current word to it
            (lineWidth, lineHeight) = dc.GetTextExtent(tempLine + word)
            # If this text element has the largest height, use that for the vertical increment (yInc)
            if lineHeight > yInc:
                yInc = lineHeight

            # If the line is still within our margins, add the word and a space to the temporary line
            if xPos + lineWidth < sizeX - xMargin + (2 * dc.GetTextExtent(" ")[1]):
                tempLine = tempLine + word
            # If the line would be too wide ...
            else:
                datLines = datLines + ((fontSpec, tempLine),)

                thisPageData.append(datLines)

                # Initialize a new line
                datLines = ()
                
                # Increment the vertical position marker
                yPos = yPos + yInc
                # Reset yInc
                yInc = 0

                # Check to see if we've reached the bottom of the page, with a one inch margin and one line's height
                if yPos >= sizeY - titleHeight - int(yMargin):

                    # Add the page to the final document data structure
                    pageData.append(thisPageData)
                    # Initialize a new page
                    thisPageData = []
                    # Reset the vertical position indicator
                    yPos = yMargin + titleHeight
                # Start a new temporary line with the word that did not fit on the last line
                tempLine = word

                (tempLineWidth, tempLineHeight) = dc.GetTextExtent(word)
                xPos = xMargin + tempLineWidth

        # When done looking at words, add the final part to the line we're building for this page
        datLines = datLines + ((fontSpec, tempLine),)

        # Add the line we're building to the Page, but don't add a blank line to the top of a page
        if (thisPageData != []) or ((len(datLines) != 1) or (datLines[0][1] != '')):
            thisPageData.append(datLines)

            xPos = xPos + lineWidth

            datLines = ()
            tempLine = ''

        # Check to see if we've reached the bottom of the page, with a one inch margin and one line's height.
        # DKW 5/24/2004 -- I've thrown in titleHeight here because the title height approximates the
        #                  height of the boxed entry on the Collection Summary Report, while it is 0 for
        #                  printing a Transcript.  TODO:  Make this more precise.
        if yPos >= sizeY - titleHeight - int(yMargin):

            # Add the page to the final document data structure
            pageData.append(thisPageData)
            # Initialize a new page
            thisPageData = []
            # Reset the vertical position indicator
            yPos = yMargin + titleHeight
    return (pageData, thisPageData, datLines, xPos, yPos, yInc)

def ProcessTXT(dc, sizeX, sizeY, inText, pageData, thisPageData, datLines, xPos, yPos, yInc, titleHeight):
    """ Process Plain Text data for PrepareData() """
    # Construct the data structure used in the for loop below, which expects a sequence of
    # tuples of the following format: ((wxFont, wxColor, style), text)
    #
    # where the first element is a wxFont which defines the style of the text
    # the second element is the text that uses this style
    data = []
    # Initialize lineWidth.  (Issue 230)
    lineWidth = 0
    fontSpec = GetDefaultFont()
    
    # Initialize the string variable used in constructing text elements
    tempLine = ''

    # each line is made up of a font spec and the associated text.
    lines = inText.split('\n')

    # Set the Device Context font to the identified fontSpec
    dc.SetFont(fontSpec[0])

    # There is a bug in wxPython.  wx.ColourRGB() transposes Red and Blue.  This hack fixes it!
    color = wx.ColourRGB(fontSpec[1])
    
    rgbValue = (color.Red() << 16) | (color.Green() << 8) | color.Blue()
    
    dc.SetTextForeground(wx.ColourRGB(rgbValue))

    # First, let's see if this starts with a line break
    for text in lines:
        # If so, add it to the datLines structure
        datLines = datLines + ((fontSpec, '\n'),)
        # Add the current datLines structure to the page
        thisPageData.append(datLines)

        (dummy, yInc) = dc.GetTextExtent('Xy')

        # Reset datLines and tempLine
        datLines = ()
        tempLine = ''
        # Reset horizontal position to the horizontal margin
        xPos = xMargin
        lineWidth = 0
        # Increment the vertical position by the height of the line plus a blank line
        yPos = yPos + yInc                 # + fontSpec.GetPointSize() + 6
        # Reset yInc
        yInc = 0

        # Break the line into words at whitespace breaks
        if ('unicode' in wx.PlatformInfo) and (type(text).__name__ == 'str'):
            words = []
            words.append(unicode(text, TransanaGlobal.encoding))
        else:

            # This "text.split()" call, necessary for line breaks, causes loss of whitespace.
            # We have to do some head stands to avoid it.
            words = text.split()
            
        # We need to retain leading whitespace
        if (len(text) > 0) and (text[0] == ' ') and (len(words) > 0):
            # This syntax captures all leading whitespace.
            words[0] = text[:text.find(words[0])] + words[0]
        # We also need to retain trailing whitespace
        if (len(text) > 0) and (text[len(text)-1] == ' ') and (len(words) > 0):
            words[len(words)-1] = words[len(words)-1] + ' '
        # Check for text that is ONLY whitespace, as this was getting lost if formatting was applied on both sides
        if (len(text) > 0) and (len(words) == 0):
            words = []
            words.append(text)

        # Iterate through the words
        for word in words:
            text = text[len(word):]

            while (len(text) > 0) and (string.whitespace.find(text[0]) > -1):
                word = word + text[0]
                text = text[1:]

            # Determine the line width if we add the current word to it
            (lineWidth, lineHeight) = dc.GetTextExtent(tempLine + word)
            # If this text element has the largest height, use that for the vertical increment (yInc)
            if lineHeight > yInc:
                yInc = lineHeight

            # If the line is still within our margins, add the word and a space to the temporary line 
            if xPos + lineWidth < sizeX - xMargin:
                tempLine = tempLine + word

            # If the line would be too wide ...
            else:
                datLines = datLines + ((fontSpec, tempLine),)

                thisPageData.append(datLines)

                # Initialize a new line
                datLines = ()
                
                # Increment the vertical position marker
                yPos = yPos + yInc
                # Reset yInc
                yInc = 0

                # Check to see if we've reached the bottom of the page, with a one inch margin and one line's height
                if yPos >= sizeY - titleHeight - int(yMargin):

                    # Add the page to the final document data structure
                    pageData.append(thisPageData)
                    # Initialize a new page
                    thisPageData = []
                    # Reset the vertical position indicator
                    yPos = yMargin + titleHeight
                # Start a new temporary line with the word that did not fit on the last line
                tempLine = word

                (tempLineWidth, tempLineHeight) = dc.GetTextExtent(word)
                xPos = xMargin + tempLineWidth

        # When done looking at words, add the final part to the line we're building for this page
        datLines = datLines + ((fontSpec, tempLine),)

        # Add the line we're building to the Page, but don't add a blank line to the top of a page
        if (thisPageData != []) or ((len(datLines) != 1) or (datLines[0][1] != '')):
            thisPageData.append(datLines)

            xPos = xPos + lineWidth

            datLines = ()
            tempLine = ''

        # Check to see if we've reached the bottom of the page, with a one inch margin and one line's height.
        # DKW 5/24/2004 -- I've thrown in titleHeight here because the title height approximates the
        #                  height of the boxed entry on the Collection Summary Report, while it is 0 for
        #                  printing a Transcript.  TODO:  Make this more precise.
        if yPos >= sizeY - titleHeight - int(yMargin):

            # Add the page to the final document data structure
            pageData.append(thisPageData)
            # Initialize a new page
            thisPageData = []
            # Reset the vertical position indicator
            yPos = yMargin + titleHeight
    return (pageData, thisPageData, datLines, xPos, yPos, yInc)

def GetPaperSize():
    """ Determine the size of the paper in pixels """
    # Get the current print definitions
    printData = TransanaGlobal.printData
    # Determine the type of paper being used so that the graphic can be set to the correct size
    papersize = printData.GetPaperId()
    if papersize in [wx.PAPER_LETTER, wx.PAPER_LETTERSMALL, wx.PAPER_NOTE, wx.PAPER_LETTER_ROTATED]:
        sizeX = int(8.5 * DPI)    # 8.5 inches x DPI dots per inch
        sizeY = 11 * DPI          # 11 inches x DPI dots per inch
    elif papersize == wx.PAPER_LEGAL:
        sizeX = int(8.5 * DPI)    # 8.5 inches x DPI dots per inch
        sizeY = 14 * DPI          # 14 inches x DPI dots per inch
    elif papersize in [wx.PAPER_A4, wx.PAPER_A4SMALL, wx.PAPER_A4_ROTATED]:
        sizeX = int(210 / 25.4 * DPI)    # 210 mm converted to inches x DPI dots per inch
        sizeY = int(297 / 25.4 * DPI)    # 297 mm converted to inches x DPI dots per inch
    elif papersize == wx.PAPER_CSHEET:
        sizeX = 17 * DPI          # 17 inches x DPI dots per inch
        sizeY = 22 * DPI          # 22 inches x DPI dots per inch
    elif papersize == wx.PAPER_DSHEET:
        sizeX = 22 * DPI          # 22 inches x DPI dots per inch
        sizeY = 34 * DPI          # 34 inches x DPI dots per inch
    elif papersize == wx.PAPER_ESHEET:
        sizeX = 34 * DPI          # 34 inches x DPI dots per inch
        sizeY = 44 * DPI          # 44 inches x DPI dots per inch
    elif (papersize == wx.PAPER_TABLOID) or (papersize == wx.PAPER_11X17):
        sizeX = 11 * DPI          # 11 inches x DPI dots per inch
        sizeY = 17 * DPI          # 17 inches x DPI dots per inch
    elif papersize == wx.PAPER_LEDGER:
        sizeX = 17 * DPI          # 17 inches x DPI dots per inch
        sizeY = 11 * DPI          # 11 inches x DPI dots per inch
    elif papersize == wx.PAPER_STATEMENT:
        sizeX = int(5.5 * DPI)    # 5.5 inches x DPI dots per inch
        sizeY = int(8.5 * DPI)    # 8.5 inches x DPI dots per inch
    elif papersize == wx.PAPER_EXECUTIVE:
        sizeX = int(7.25 * DPI)   # 7.25 inches x DPI dots per inch
        sizeY = int(10.5 * DPI)   # 10.5 inches x DPI dots per inch
    elif papersize in [wx.PAPER_A3, wx.PAPER_A3_ROTATED]:
        sizeX = int(297 / 25.4 * DPI)    # 210 mm converted to inches x DPI dots per inch
        sizeY = int(420 / 25.4 * DPI)    # 297 mm converted to inches x DPI dots per inch
    elif papersize in [wx.PAPER_A5, wx.PAPER_A5_ROTATED]:
        sizeX = int(148 / 25.4 * DPI)    # 148 mm converted to inches x DPI dots per inch
        sizeY = int(210 / 25.4 * DPI)    # 210 mm converted to inches x DPI dots per inch
    elif papersize == wx.PAPER_B4:
        sizeX = int(250 / 25.4 * DPI)    # 250 mm converted to inches x DPI dots per inch
        sizeY = int(354 / 25.4 * DPI)    # 354 mm converted to inches x DPI dots per inch
    elif papersize == wx.PAPER_B5:
        sizeX = int(182 / 25.4 * DPI)    # 182 mm converted to inches x DPI dots per inch
        sizeY = int(257 / 25.4 * DPI)    # 257 mm converted to inches x DPI dots per inch
    elif papersize == wx.PAPER_FOLIO:
        sizeX = int(8.5 * DPI)    # 8.5 inches x DPI dots per inch
        sizeY = 13 * DPI          # 13 inches x DPI dots per inch
    elif papersize == wx.PAPER_QUARTO:
        sizeX = int(215 / 25.4 * DPI)    # 215 mm converted to inches x DPI dots per inch
        sizeY = int(275 / 25.4 * DPI)    # 275 mm converted to inches x DPI dots per inch
    elif papersize == wx.PAPER_9X11:
        sizeX =  9 * DPI          #  9 inches x DPI dots per inch
        sizeY = 11 * DPI          # 11 inches x DPI dots per inch
    elif papersize == wx.PAPER_10X11:
        sizeX = 10 * DPI          # 10 inches x DPI dots per inch
        sizeY = 11 * DPI          # 11 inches x DPI dots per inch
    elif papersize == wx.PAPER_10X14:
        sizeX = 10 * DPI          # 10 inches x DPI dots per inch
        sizeY = 14 * DPI          # 14 inches x DPI dots per inch
    elif papersize == wx.PAPER_12X11:
        sizeX = 12 * DPI          # 12 inches x DPI dots per inch
        sizeY = 11 * DPI          # 11 inches x DPI dots per inch
    elif papersize == wx.PAPER_15X11:
        sizeX = 15 * DPI          # 15 inches x DPI dots per inch
        sizeY = 11 * DPI          # 11 inches x DPI dots per inch
    elif papersize == wx.PAPER_FANFOLD_US:
        sizeX = int(14.875 * DPI) # 14-7/8 inches x DPI dots per inch
        sizeY = 11 * DPI          # 11 inches x DPI dots per inch
    elif papersize == wx.PAPER_FANFOLD_STD_GERMAN:
        sizeX = int(8.5 * DPI)    # 8.5 inches x DPI dots per inch
        sizeY = 12 * DPI          # 12 inches x DPI dots per inch
    elif papersize == wx.PAPER_FANFOLD_LGL_GERMAN:
        sizeX = int(8.5 * DPI)    # 8.5 inches x DPI dots per inch
        sizeY = 13 * DPI          # 13 inches x DPI dots per inch
    else:
        # TODO:  Display error to user here!
        dlg = Dialogs.ErrorDialog(None, _('Transana does not recognize the Paper Size selected for you printer.\nIt is unlikely that your report will print correctly.\nPlease select a different Paper Size.'))
        dlg.ShowModal()
        dlg.Destroy()
        sizeX = int(8.5 * DPI)    # 8.5 inches x DPI dots per inch
        sizeY = 11 * DPI          # 11 inches x DPI dots per inch

    # If we are in landscape mode, reverse the paper dimensions
    if printData.GetOrientation() == wx.LANDSCAPE:
        temp = sizeX
        sizeX = sizeY
        sizeY = temp
    return (sizeX, sizeY)

def PrepareData(printData, transcriptObj=None, noteTxt=None, title='', subtitle=''):
    """ This method takes a data structure of an unknown number of data lines and prepares them for use by MyPrintout,
        Transana's custom wxPrintout Object.  The main task accomplished here is to divide the data into seperate Pages.
        It returns a wxBitmap Object and a list object needed by the MyPrintout class. """

    # Determine the size of the currently selected printer paper
    (sizeX, sizeY) = GetPaperSize()

    # Create an empty Bitmap image the size needed for the printout.  This serves two purposes:
    # 1.  It allows us to calculate line sizes so that the data can be divided up into pages correctly.
    # 2.  It sets the size correctly so that the wxPrintout Scaling Factor can be determined correctly.
    graphic = wx.EmptyBitmap(sizeX, sizeY)

    # Create a Device Context to go along with the Image
    dc = wx.BufferedDC(None, graphic)
    dc.Clear()

    # No title used for now
    titleHeight = 0
    # Determine the height of the Title, which influences where the normal text should be printed
    if title != '':
        dc.SetFont(GetTitleFont()[0])
        (lineWidth, lineHeight) = dc.GetTextExtent(title)
        titleHeight = lineHeight + int(DPI / 12)
    # If there is a subtitle, add its height to the title height
    if subtitle != '':
        dc.SetFont(GetDefaultFont()[0])
        (lineWidth, lineHeight) = dc.GetTextExtent(subtitle)
        # Subtitles are places 6 points down on the screen
        titleHeight = titleHeight + lineHeight + int(DPI / 12)

    dc.SetFont(GetDefaultFont()[0])

    # Go through all the data and parse it to pages here

    # Initialize the list structure
    pageData = []
    # Initialize a temporary structure for one page's data
    thisPageData = []

    # Start the horizontal position at the horizontal margin
    xPos = xMargin
    # Start the vertical position at the vertical margin
    yPos = yMargin + titleHeight
    # We need to know the height of the largest text element for each line, the yInc(rement).
    # Initialize this to 0
    yInc = 0

    # Each line is made up of tuples of text elements.  Initialize a tuple for this line
    datLines = ()

    # This report module is used for printing Transcripts and for printing the Collection Summary Report.
    # If transcriptObj is not None, we're printing a Transcript
    if transcriptObj != None:

        (pageData, thisPageData, datLines, xPos, yPos, yInc) = \
                   ProcessRTF(dc, sizeX, sizeY, transcriptObj.GetTranscriptWithoutTimeCodes(), pageData, thisPageData, datLines, xPos, yPos, yInc, titleHeight)

    # If transcriptObj is None, we're printing a Note.
    else:

        (pageData, thisPageData, datLines, xPos, yPos, yInc) = \
                   ProcessTXT(dc, sizeX, sizeY, noteTxt, pageData, thisPageData, datLines, xPos, yPos, yInc, titleHeight)

    # If there is a final page we were working on when we examined the last of the data,
    # add it to the final document data structure
    if thisPageData != []:
        pageData.append(thisPageData)

    if False:
        print
        print "PrepareData pageData:"
        for page in pageData:
            for line in page:
                for segment in line:
                    (fontSpec, text) = segment
                    if text != '\n':
                        print text,
                print
            print ' --------------- Page Break --------------- '
        print


    # Return the blank bitmap and the document data structure to the calling routine.
    return (graphic, pageData)
        

#----------------------------------------------------------------------

class MyPrintout(wx.Printout):
    """ This class creates a custom wxPrintout Object for Print Preview and Printing for Transana.  Parameters are
        Report Title, a blank graphic used for page sizing, a data structure that has all the lines for all the pages,
        and an optional subtitle.  The PrepareData() function in this unit builds the graphic and the data structure
        correctly, and should be called prior to creating a MyPrintout Object. """
    def __init__(self, title, graphic, pageData, subtitle=''):
        # Create a wxPrintout Object
        wx.Printout.__init__(self, title)
        # Store the data
        self.title = title
        self.subtitle = subtitle
        self.graphic = graphic
        self.pageData = pageData

    def OnBeginDocument(self, start, end):
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            return self.base_OnBeginDocument(start, end)
        else:
            return super(MyPrintout, self).OnBeginDocument(start, end)

    def OnEndDocument(self):
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            self.base_OnEndDocument()
        else:
            super(MyPrintout, self).OnEndDocument()

    def OnBeginPrinting(self):
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            self.base_OnBeginPrinting()
        else:
            super(MyPrintout, self).OnBeginPrinting()

    def OnEndPrinting(self):
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            self.base_OnEndPrinting()
        else:
            super(MyPrintout, self).OnEndPrinting()

    def OnPreparePrinting(self):
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            self.base_OnPreparePrinting()
        else:
            super(MyPrintout, self).OnPreparePrinting()

    def HasPage(self, page):
        # The HasPage function tells the framework if a given page number exists.
        # len(self.pageData) is the number of pages the report has!
        if page <= len(self.pageData):
            return True
        else:
            return False

    def GetPageInfo(self):
        # GetPageInfo function tells the framework how many pages there are in the full document.
        # You are supposed to be able to determine the number of pages in the
        # OnPreparePrinting method, but unfortunately GetPageInfo gets called
        # first and I don't see a way to alter the number of pages later.
        # len(self.pageData) is the number of pages the report has!
        return(1, len(self.pageData), 1, len(self.pageData))

    def OnPrintPage(self, page):
        """ This method actually builds the requested page for presentation to the screen or printer """
        # Get the Device Context
        dc = self.GetDC()

        # Determine the Print Scaling Factors by comparing the dimensions of the Device Context
        # (which has the printer's resolution times the paper size) with the Graphic's (from the
        # PrepareData() function) resolution (which is screen resolution times paper size).

        # Determine the size of the DC and the Graphic
        (dcX, dcY) = dc.GetSizeTuple()
        graphicX = self.graphic.GetWidth()
        graphicY = self.graphic.GetHeight()

        # Scaling Factors are the DC dimensions divided by the Graphics dimensions
        scaleX = float(dcX)/graphicX
        scaleY = float(dcY)/graphicY

        # Apply the scaling factors to the Device Context.  If you don't do this, the screen will look
        # fine but the printer version will be very, very tiny.
        dc.SetUserScale(scaleX, scaleY)

        titleHeight = 0
        # Add Title to the Printout
        if self.title != '':
            # Set the appropriate style and font
            dc.SetFont(GetTitleFont()[0])
            # Get the appropriate text
            line = self.title
            # Determine the width and height of the text
            (lineWidth, lineHeight) = dc.GetTextExtent(line)
            # Align and position the text, then add it to the Printout
            dc.DrawText(line, int(graphicX/2.0 - lineWidth/2.0), DPI)
            # Keep track of the title's height
            titleHeight = lineHeight + int(DPI / 12)

        # If there is a defined Subtitle, add it to the printout
        if self.subtitle != '':
            dc.SetFont(GetDefaultFont()[0])
            # Get the text
            line = self.subtitle
            # Get the width and height of the text
            (lineWidth, lineHeight) = dc.GetTextExtent(line)
            # Align and position the text, then add it to the Printout
            dc.DrawText(line, int(graphicX/2.0 - lineWidth/2.0), int((DPI * 17) / 16) + titleHeight)
            # Add the Subtitle's height to the titleHeight
            titleHeight = titleHeight + lineHeight + int(DPI / 8)

        # Place the Page Number in the lower right corner of the page
        defaultFont = GetDefaultFont()
        dc.SetFont(defaultFont[0])
        dc.SetTextForeground(wx.ColourRGB(defaultFont[1]))
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            txt = unicode(_("Page: %d"), 'utf8') % page
        else:
            txt = _("Page: %d") % page
        (lineWidth, lineHeight) = dc.GetTextExtent(txt)
        # Position and draw the text
        dc.DrawText(txt, graphicX - lineWidth - 40, graphicY - lineHeight - 40)
        # Print the lines for this page
        xPos = xMargin
        # Lines start 1/3 inch below the title.  (Title has 1-inch margin)
        yPos = yMargin + titleHeight
        # self.pageData[page - 1] is the data for THIS page.  Process each line.

        for lines in self.pageData[page - 1]:
            # We need to know the height of the largest text element for each line, the yInc(rement)
            yInc = 0
            # Each line is made up of one of more text elements.  Process each element, which has a style and text.
            for (fontSpec, line) in lines:
                (font, color, style) = fontSpec

                if line == '\n':
                    xPos = xMargin
                    (dummy, yInc) = dc.GetTextExtent('Xy')
                    # Increment the vertical position indicator
                    yPos = yPos + yInc
                    # Reset yInc
                    yInc = 0
                else:
                    # Set the font
                    dc.SetFont(font)
                    # Set the Device Context font to the identified fontSpec
                    dc.SetFont(fontSpec[0])

                    # There is a bug in wxPython.  wx.ColourRGB() transposes Red and Blue.  This hack fixes it!
                    color = wx.ColourRGB(color)
                    
                    rgbValue = (color.Red() << 16) | (color.Green() << 8) | color.Blue()
                    
                    dc.SetTextForeground(wx.ColourRGB(rgbValue))
                    
                    # Determine the left indent
                    xindent = 0 # self.styles[style]['indent']
                    # Determine the width and height of the line
                    (lineWidth, lineHeight) = dc.GetTextExtent(line)
                    # If this text element has the largest height, use that for the vertical increment (yInc)
                    if lineHeight > yInc:
                        yInc = lineHeight

                    if STYLE_DRAW_BOX in style:
                        brush = wx.TheBrushList.FindOrCreateBrush(wx.LIGHT_GREY, wx.SOLID)
                        dc.SetBrush(brush)
                        # The box isn't drawing right on the Mac.  Let's make an adjustment
                        if '__WXMAC__' in wx.PlatformInfo:
                            yPosAdjusted = yPos - 2
                        else:
                            yPosAdjusted = yPos
                        dc.DrawRectangle(xMargin - int(DPI / 8), yPosAdjusted, graphicX - (2 * xMargin) + int(DPI / 4), yInc * 3 + int(DPI / 24))
                    else:
                        brush = wx.TheBrushList.FindOrCreateBrush(wx.WHITE, wx.SOLID)
                        dc.SetBrush(brush)

                    if xPos + lineWidth > graphicX - int(xMargin / 2):
                        xPos = xMargin
                        # Increment the vertical position indicator
                        yPos = yPos + yInc
                        yInc = 0

                    # dc.DrawText('%d' % yPos, 0, yPos)
                    
                    if STYLE_CENTER in style:
                        dc.DrawText(line, int(graphicX/2.0 - lineWidth/2.0), yPos)
                    else:
                        # For now we left-align everything (it's all the transcripts support, anyway)
                        dc.DrawText(line, xPos, yPos)

                    # Align and position the text, then add it to the Printout
                    #if self.styles[style]['align'] == mstyLEFT:
                    #    dc.DrawText(line, DPI + xindent, yPos)
                    #elif self.styles[style]['align'] == mstyCENTER:
                    #    dc.DrawText(line, int(graphicX/2.0 - lineWidth/2.0), yPos)
                    #elif self.styles[style]['align'] == mstyRIGHT:
                    #    dc.DrawText(line, graphicX - DPI - lineWidth, yPos)
                    xPos = xPos + lineWidth

        # When done with the page, return True to indicate success
        return True

#----------------------------------------------------------------------
# If this unit is run as a stand-alone, the code below creates an application that is useful for testing new styles and features

if __name__ == '__main__':
    import gettext
    
    # Declare Control IDs
    # Menu Item File > Printer Setup
    M_FILE_PRINTSETUP    =  102
    # Menu Item File > Print Preview
    M_FILE_PRINTPREVIEW  =  103
    # Menu Item File > Print
    M_FILE_PRINT         =  104
    # Menu Item File > Exit
    M_FILE_EXIT          =  105

    class Main(wx.Frame):
        """ This is the Main Window for the Print Framework Test Program """
        def __init__(self, parent, ID, title):
            # Specify the Internationalization Base
            gettext.install("ReportPrintoutClass")
            # Create the basic Frame structure with a white background
            self.frame = wx.Frame.__init__(self, parent, ID, 'Transana - %s' % title, pos=(10, 10), size=(300, 150), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL | wx.NO_FULL_REPAINT_ON_RESIZE)
            self.title = title
            self.SetBackgroundColour(wx.WHITE)

            # Add a Menu Bar
            menuBar = wx.MenuBar()                                                        # Create the Menu Bar
            self.menuFile = wx.Menu()                                                     # Create the File Menu
            self.menuFile.Append(M_FILE_PRINTSETUP, "Printer Setup", "Set up Printer")      # Add "Printer Setup" to the File Menu
            self.menuFile.Append(M_FILE_PRINTPREVIEW, "Print Preview", "Preview your printed output") # Add "Print Preview" to the File Menu
            self.menuFile.Append(M_FILE_PRINT, "&Print", "Send your output to the Printer") # Add "Print" to the File Menu
            self.menuFile.Append(M_FILE_EXIT, "E&xit", "Exit the %s program" % self.title)  # Add "Exit" to the File Menu
            menuBar.Append(self.menuFile, '&File')                                          # Add the File Menu to the Menu Bar
            wx.EVT_MENU(self, M_FILE_PRINTSETUP, self.OnPrintSetup)                         # Attach File > Print Setup to a method
            wx.EVT_MENU(self, M_FILE_PRINTPREVIEW, self.OnPrintPreview)                     # Attach File > Print Preview to a method
            wx.EVT_MENU(self, M_FILE_PRINT, self.OnPrint)                                   # Attach File > Print to a method
            wx.EVT_MENU(self, M_FILE_EXIT, self.CloseWindow)                                # Attach CloseWindow to File > Exit
            self.SetMenuBar(menuBar)                                                        # Connect the Menu Bar to the Frame

            # Add a Status Bar
            self.CreateStatusBar()

            # Prepare the wxPrintData object for use in Printing
            self.printData = wx.PrintData()
            self.printData.SetPaperId(wx.PAPER_LETTER)

            # Show the Frame
            self.Show(True)

            # Get a filename of an RTF file to use as the transcript
            dlg = wx.FileDialog(None, wildcard="*.rtf", style=wx.OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                fname = dlg.GetPath()
                dlg.Destroy()
            else:
                dlg.Destroy()
                self.Close()
            
            # Create a dummy Transcript object
            self.transcriptObj = Transcript.Transcript()
            f = open(fname, "r")
            self.transcriptObj.text = f.read()
            f.close()

            # Initialize the pageData structure to an empty list
            self.pageData = []
            


        # Define the Method that implements Printer Setup
        def OnPrintSetup(self, event):
            """ Printer Setup method """
            # Create a Print Setup Dialog
            printerDialog = wx.PrintDialog(self)
            # Supply the existing PrintData to the Print Setup Dialog
            printerDialog.GetPrintDialogData().SetPrintData(self.printData)
            # Indicate that we want the Print Setup Dialog to be displayed
            printerDialog.GetPrintDialogData().SetSetupDialog(True)
            # Show the Print Setup Dialog
            if printerDialog.ShowModal() == wx.ID_OK:
                # Update the PrintData object's information
                self.printData = printerDialog.GetPrintDialogData().GetPrintData()
            # Destroy the Print Setup Dialog
            # printerDialog.Destroy()


        # Define the Method that implements Print Preview
        def OnPrintPreview(self, event):
            """ Print Preview Method """
            # Define a Subtitle
            subtitle = "Subtitles are optional"

            # We already have our report (our self.data structure) defined.  The data should be a list of
            # lines, and each line is a tuple of text elements.   Each text element is a tuple of Style and Text.
            # There can be multiple text elements on a line, but the report creator is responsible for making sure
            # they have different alignments and don't overlap.  A line can also be too long, in which case it is
            # automatically wrapped.
            #
            # The initial data structure needs to be prepared.  What PrepareData() does is to create a graphic
            # object that is the correct size and dimensions for the type of paper selected, and to create
            # a datastructure that breaks the data sent in into separate pages, again based on the dimensions
            # of the paper currently selected.
            (self.graphic, self.pageData) = PrepareData(self.printData, self.transcriptObj)

            # Send the results of the PrepareData() call to the MyPrintout object, once for the print preview
            # version and once for the printer version.  
            printout = MyPrintout(self.title, self.graphic, self.pageData, subtitle)
            printout2 = MyPrintout(self.title, self.graphic, self.pageData, subtitle)

            # Create the Print Preview Object
            self.preview = wx.PrintPreview(printout, printout2, self.printData)
            # Check for errors during Print preview construction
            if not self.preview.Ok():
                self.SetStatusText("Print Preview Problem")
                return
            # Create the Frame for the Print Preview
            theWidth = max(wx.ClientDisplayRect()[2] - 180, 760)
            theHeight = max(wx.ClientDisplayRect()[3] - 200, 560)
            frame2 = wx.PreviewFrame(self.preview, self, _("Print Preview"), size=(theWidth, theHeight))
            frame2.Centre()
            # Initialize the Frame for the Print Preview
            frame2.Initialize()
            # Display the Print Preview Frame
            frame2.Show(True)

        # Define the Method that implements Print
        def OnPrint(self, event):
            """ Print Method """
            # Create a PrintDialogData Object
            pdd = wx.PrintDialogData()
            # Send it the PrintData information
            pdd.SetPrintData(self.printData)
            # Create a Printer object with the PrintDialogData
            printer = wx.Printer(pdd)

            # Define a subtitle
            subtitle = "Subtitles are optional"
            
            # We already have our report (our self.data structure) defined.  The data should be a list of
            # lines, and each line is a tuple of text elements.   Each text element is a tuple of Style and Text.
            # There can be multiple text elements on a line, but the report creator is responsible for making sure
            # they have different alignments and don't overlap.  A line can also be too long, in which case it is
            # automatically wrapped.
            #
            # The initial data structure needs to be prepared.  What PrepareData() does is to create a graphic
            # object that is the correct size and dimensions for the type of paper selected, and to create
            # a datastructure that breaks the data sent in into separate pages, again based on the dimensions
            # of the paper currently selected.
            (self.graphic, self.pageData) = PrepareData(self.printData, self.transcriptObj)
            # Send the results of the PrepareData() call to the MyPrintout object
            printout = MyPrintout(self.title, self.graphic, self.pageData, subtitle)
            # send the MyPrintout object to the Printer
            if not printer.Print(self, printout):
                dlg = Dialogs.ErrorDialog(None, _("There was a problem printing this report."))
                dlg.ShowModal()
                dlg.Destoy()
            # NO!  REMOVED to prevent crash on 2nd print attempt following Filter Config.
            # else:
                # Save any changes that may have been made to the Printer Setup
            #     self.printData = printer.GetPrintDialogData().GetPrintData()
            # Destroy the MyPrintout Object
            printout.Destroy()


        # Define the Method that closes the Window on File > Exit
        def CloseWindow(self, event):
            """ Close Window method """
            self.Close()


        
    class MyApp(wx.App):
        def OnInit(self):
            # Define the Main Window
            frame = Main(None, -1, 'Report Generator Test')
            # Set the Main Window as the Top Window
            self.SetTopWindow(frame)
            # Report success
            return True

    # run the application
    app = MyApp(0)
    app.MainLoop()
