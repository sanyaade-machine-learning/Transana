#Copyright (C) 2003 - 2005  The Board of Regents of the University of Wisconsin System
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

""" This class implements a Print Preview and Printing mechanism, based on the wxPython Print Framework. """

__author__ = "David K. Woods <dwoods@wcer.wisc.edu>"

# This unit is designed to implement Print Preview and Print mechanisms for Transana's Reports.
# It is designed to allow the user to define a set of font styles for the report and send a data structure
# that indentifies styles with text to be printed.  This unit should handle all pagination automatically,
# and even does CRUDE line wrapping.  All styles are applied at the paragraph level.
#
# It is possible to include multiple text elements on the same line.  However, they will only be displayed
# in a legible way if the user programs the report to give them different alignments.  This allows, for
# example, the user to produce a report with left-aligned words on the same line as right-aligned counts.
#
# To use it, do the following:
#
#  - Edit the MyPrintStyles function below to reflect the styles you want to use in your report.  Each style
#    consists of a font definition (wxFont object), an alignment specification (left, center, right), and an
#    indent value (left margin) for left-aligned text elements.
#
#  - Prepare a List data structure with one entry for each line.  Each line is made up of a tuple of elements,
#    with a logical max of 3 elements per line (one left-aligned, one center-aligned, and one right-aligned).
#    Each of these elements is a tuple made up of a Style indicator (as defined in MyPrintStyles) and the
#    text for that element.
#
#    For example, look at the following data structure definition:
#        self.data = []
#        self.data.append((('Heading', 'This is Line 1, using "Heading" style'),))
#        self.data.append((('Subheading', 'This is Line 2, using "Subheading" style'),))
#        self.data.append((('Subtext', 'This is Line 3, using "Subtext" style'),))
#        self.data.append((('Subtext', "This is Line 4, using 'Subtext' style.  It's a really, really long line that will need to be parsed because it's too long to fit on a single line in the final printed version of the stupid thing.  I mean, this line is really, really, really long.  I mean it.  Really."),))
#        self.data.append((('Normal', 'This is Line 5, using "Normal" style'),))
#        self.data.append((('NormalCenter', 'This is Line 6, using "NormalCenter" style'),))
#        self.data.append((('NormalRight', 'This is Line 7, using "NormalRight" style'),))
#        self.data.append((('Normal', 'Line 8, Part 1, Left'), ('NormalCenter', 'Line 8, Part 2, Center'), ('NormalRight', 'Line 8, Part 3, "NormalRight" style')))
#
#    Obviously, you don't want to mix multiple-element lines with lines you can anticipate will be wrapped.
#
#  - Call the PrepareData function in this unit.  Parameters are:
#
#        a wxPrintData Object
#        the Title of your report, a String
#        the List data structure you prepared above
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
# import Transana Error Dialog
import Dialogs


# For pages to be sized and proportioned correctly, we need different DPI scaling factors on different platforms.
if "__WXMAC__" in wx.PlatformInfo:
    DPI = 80
else:
    DPI = 96

#----------------------------------------------------------------------

# Define Alignment Constants for MyStyles
mstyLEFT   = wx.NewId()
mstyCENTER = wx.NewId()
mstyRIGHT  = wx.NewId()

def MyPrintStyles():
    """ This method defines the Font Styles used by Transana's PrintoutClass framework.  It is used by both the
        PrepareData method and by the MyPrintout Class for setting up the appropriate data structure, displaying
        it for Print Preview, and printing it.  This method returns a dictionary data object. """
    # I want all fonts to have the same character set.  It is defined here to make it easy to change.  (This is not required.)
    fontFace = wx.SWISS

    # Styles are set up as dictionary items containing font, alignment, and indent information
    style = {}
    
    # 'Title' is 18 pt Bold Underlined, centered
    font = wx.Font(18, fontFace, wx.NORMAL, wx.BOLD, True)
    style['Title'] = {'font' : font, 'align': mstyCENTER, 'indent': 0}
    
    # 'Subtitle' is 14 pt unadorned, centered
    font = wx.Font(14, fontFace, wx.NORMAL, wx.NORMAL, False)
    style['Subtitle'] = {'font' : font, 'align': mstyCENTER, 'indent': 0}
    
    # 'Heading' is 14 pt Bold, left, no indent
    font = wx.Font(14, fontFace, wx.NORMAL, wx.BOLD, False)
    style['Heading'] = {'font' : font, 'align': mstyLEFT, 'indent': 0}
    
    # 'Subheading' is 12 pt unadorned, left, 12 indent
    font = wx.Font(12, fontFace, wx.NORMAL, wx.NORMAL, False)
    style['Subheading'] = {'font' : font, 'align': mstyLEFT, 'indent': 12}
    
    # 'Subtext' is 10 pt unadorned, left, 24 indent
    font = wx.Font(10, fontFace, wx.NORMAL, wx.NORMAL, False)
    style['Subtext'] = {'font' : font, 'align': mstyLEFT, 'indent': 24}
    
    # 'Normal' is 10 pt unadorned, left, no indent (The same font can be used for multiple definitions!)
    font = wx.Font(10, fontFace, wx.NORMAL, wx.NORMAL, False)
    style['Normal'] = {'font' : font, 'align': mstyLEFT, 'indent': 0}
    
    # 'NormalCenter' is 10 pt unadorned, Center, no indent  
    style['NormalCenter'] = {'font' : font, 'align': mstyCENTER, 'indent': 0}
    
    # NormalRight is 10 pt unadorned, right, no indent
    style['NormalRight'] = {'font' : font, 'align': mstyRIGHT, 'indent': 0}

    return style

#----------------------------------------------------------------------

def PrepareData(printData, title, data, subtitle=''):
    """ This method takes a data structure of an unknown number of data lines and prepares them for use by MyPrintout,
        Transana's custom wxPrintout Object.  The main task accomplished here is to divide the data into seperate Pages.
        It returns a wxBitmap Object and a list object needed by the MyPrintout class. """
    # First, determine the type of paper being used so that the graphic can be set to the correct size
    papersize = printData.GetPaperId()
    if (papersize == wx.PAPER_LETTER) or (papersize == wx.PAPER_LETTERSMALL) or (papersize == wx.PAPER_NOTE):
        sizeX = int(8.5 * DPI)    # 8.5 inches x DPI dots per inch
        sizeY = 11 * DPI          # 11 inches x DPI dots per inch
    elif papersize == wx.PAPER_LEGAL:
        sizeX = int(8.5 * DPI)    # 8.5 inches x DPI dots per inch
        sizeY = 14 * DPI          # 14 inches x DPI dots per inch
    elif (papersize == wx.PAPER_A4) or (papersize == wx.PAPER_A4SMALL):
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
    elif papersize == wx.PAPER_A3:
        sizeX = int(297 / 25.4 * DPI)    # 210 mm converted to inches x DPI dots per inch
        sizeY = int(420 / 25.4 * DPI)    # 297 mm converted to inches x DPI dots per inch
    elif papersize == wx.PAPER_A5:
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
    elif papersize == wx.PAPER_10X14:
        sizeX = 10 * DPI          # 10 inches x DPI dots per inch
        sizeY = 14 * DPI          # 14 inches x DPI dots per inch
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
        # Display error to user here!
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

    # Create an empty Bitmap image the size needed for the printout.  This serves two purposes:
    # 1.  It allows us to calculate line sizes so that the data can be divided up into pages correctly.
    # 2.  It sets the size correctly so that the wxPrintout Scaling Factor can be determined correctly.
    graphic = wx.EmptyBitmap(sizeX, sizeY)

    # Create a Device Context to go along with the Image
    dc = wx.BufferedDC(None, graphic)
    dc.Clear()

    # Get the Style Definitions used in building the page
    styles = MyPrintStyles()

    # Determine the height of the Title, which influences where the normal text should be printed
    dc.SetFont(styles['Title']['font'])
    (lineWidth, lineHeight) = dc.GetTextExtent(title)
    titleHeight = lineHeight

    # If there is a subtitle, add its height to the title height
    if subtitle != '':
        dc.SetFont(styles['Subtitle']['font'])
        (lineWidth, lineHeight) = dc.GetTextExtent(subtitle)
        # Subtitles are places 6 points down on the screen
        titleHeight = titleHeight + lineHeight + int(DPI / 16)

    # Go through all the data and parse it to pages here
    # Initialize the list structure
    pageData = []
    # To determine where Page Breaks fall, we need to track the vertical printing position, the "yPos".
    # Start with 1 inch margin, title height, and 1/3 inch between title and text
    yPos = DPI + titleHeight + int(DPI / 3)
    # Initialize a temporary structure for one page's data
    thisPageData = []
    # Iterate through all the data sent in, one line at a time
    for lines in data:
        # We need to know the height of the largest text element for each line, the yInc(rement)
        yInc = 0
        # Each line is made up of a tuple of text elements.  Initialize a tuple for this line
        datLines = ()
        # Iterate through the text elements for each line
        for (style, line) in lines:
            # Set the font for the identified Style
            dc.SetFont(styles[style]['font'])
            # Determine the height of the text
            (lineWidth, lineHeight) = dc.GetTextExtent(line)
            # A blank line has no font height here, so needs an alternate lineHeight value or it won't show up at all
            if line == '':
                # Use the font's height plus 6 to allow for ascenders and descenders
                lineHeight = styles[style]['font'].GetPointSize() + int(DPI / 16)
            # If this text element has the largest height, use that for the vertical increment (yInc)
            if lineHeight > yInc:
                yInc = lineHeight
            # If a line is too wide, break it into multiple lines.
            # Take the paper width, subtract 1-inch margins for each side of the paper (DPI x 2) and
            # subtract the indent value once.  (This implements left-indentation only)
            if lineWidth > sizeX - (DPI * 2) - styles[style]['indent']:
                # Break the line into words at whitespace breaks
                words = string.split(line)
                # Initialize a string so we can build new lines
                tempLine = ''
                # Iterate through the words
                for word in words:
                    # Determine the line width if we add the current word to it
                    (lineWidth, lineHeight) = dc.GetTextExtent(tempLine + word)
                    # If the line is still within our margins, add the word and a space to the temporary line 
                    if lineWidth < sizeX - (DPI * 2) - styles[style]['indent']:
                        tempLine = tempLine + word + ' '
                    # If the line would be too wide ...
                    else:
                        # ... add the line we're building to the Page, but don't add a blank line to the top of a page
                        if (thisPageData != []) or (tempLine != ''):
                            thisPageData.append(((style, tempLine),))
                        # Initialize a new line
                        datLines = ()
                        # Increment the vertical position marker
                        yPos = yPos + yInc
                        # Check to see if we've reached the bottom of the page, with a one inch margin and one line's height
                        if yPos > sizeY - DPI - int(DPI / 6):
                            # Add the page to the final document data structure
                            pageData.append(thisPageData)
                            # Initialize a new page
                            thisPageData = []
                            # Reset the vertical position indicator
                            yPos = DPI + titleHeight + int(DPI / 3)
                        # Start a new temporary line with the word that did not fit on the last line
                        tempLine = word + ' '
                # When done looking at words, add the final part to the line we're building for this page
                datLines = datLines + ((style, tempLine),)
            else:
                # If the line fits, add it to the line we're building for this page
                datLines = datLines + ((style, line),)
        # Add the line we're building to the Page, but don't add a blank line to the top of a page
        if (thisPageData != []) or ((len(datLines) != 1) or (datLines[0][1] != '')):
            thisPageData.append(datLines)
        # Increment the vertical position indicator
        yPos = yPos + yInc
        # Check to see if we've reached the bottom of the page, with a one inch margin and one line's height
        if yPos > sizeY - DPI - int(DPI / 6):
            # Add the page to the final document data structure
            pageData.append(thisPageData)
            # Initialize a new page
            thisPageData = []
            # Reset the vertical position indicator
            yPos = DPI + titleHeight + int(DPI / 3)

    # If there is a final page we were working on when we examined the last of the data,
    # add it to the final document data structure
    if thisPageData != []:
        pageData.append(thisPageData)

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
        # Access the Style definitions
        self.styles = MyPrintStyles()
        # Store the data
        self.title = title
        self.subtitle = subtitle
        self.graphic = graphic
        self.pageData = pageData

    def OnBeginDocument(self, start, end):
        return self.base_OnBeginDocument(start, end)

    def OnEndDocument(self):
        self.base_OnEndDocument()

    def OnBeginPrinting(self):
        self.base_OnBeginPrinting()

    def OnEndPrinting(self):
        self.base_OnEndPrinting()

    def OnPreparePrinting(self):
        self.base_OnPreparePrinting()

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

        # Add Title to the Printout
        # Set the appropriate style and font
        style = 'Title'
        dc.SetFont(self.styles[style]['font'])
        # Get the appropriate text
        line = self.title
        # Determine the width and height of the text
        (lineWidth, lineHeight) = dc.GetTextExtent(line)
        # Align and position the text, then add it to the Printout
        if self.styles[style]['align'] == mstyLEFT:
            dc.DrawText(line, DPI + xindent, DPI)
        elif self.styles[style]['align'] == mstyCENTER:
            dc.DrawText(line, int(graphicX/2.0 - lineWidth/2.0), DPI)
        elif self.styles[style]['align'] == mstyRIGHT:
            dc.DrawText(line, graphicX - DPI - lineWidth, DPI)
        # Keep track of the title's height
        titleHeight = lineHeight

        # If there is a defined Subtitle, add it to the printout
        if self.subtitle != '':
            # Set the style and font
            style = 'Subtitle'
            dc.SetFont(self.styles[style]['font'])
            # Get the text
            line = self.subtitle
            # Get the width and height of the text
            (lineWidth, lineHeight) = dc.GetTextExtent(line)
            # Align and position the text, then add it to the Printout
            if self.styles[style]['align'] == mstyLEFT:
                dc.DrawText(line, DPI + xindent, int((DPI * 17)/ 16) + titleHeight)
            elif self.styles[style]['align'] == mstyCENTER:
                dc.DrawText(line, int(graphicX/2.0 - lineWidth/2.0), int((DPI * 17)/ 16) + titleHeight)
            elif self.styles[style]['align'] == mstyRIGHT:
                dc.DrawText(line, graphicX - DPI - lineWidth, int((DPI * 17)/ 16) + titleHeight)
            # Add the Subtitle's height to the titleHeight
            titleHeight = titleHeight + lineHeight + int(DPI / 16)

        # Place the Page Number in the lower right corner of the page
        dc.SetFont(self.styles['Normal']['font'])
        txt = _("Page: %d") % page
        (lineWidth, lineHeight) = dc.GetTextExtent(txt)
        # Position and draw the text
        dc.DrawText(txt, graphicX - lineWidth - 40, graphicY - lineHeight - 40)

        # Print the lines for this page
        # Lines start 1/3 inch below the title.  (Title has 1-inch margin)
        yPos = DPI + titleHeight + int(DPI / 3)

        if len(self.pageData) == 0:
            errordlg = Dialogs.ErrorDialog(None, _('The requested Keyword Summary Report contains no data to display.'))
            errordlg.ShowModal()
            errordlg.Destroy()
            return False
        else:
        
            # self.pageData[page - 1] is the data for THIS page.  Process each line.
            for lines in self.pageData[page - 1]:
                # We need to know the height of the largest text element for each line, the yInc(rement)
                yInc = 0
                # Each line is made up of one of more text elements.  Process each element, which has a style and text.
                for (style, line) in lines:
                    # Set the font
                    dc.SetFont(self.styles[style]['font'])
                    # Determine the left indent
                    xindent = self.styles[style]['indent']
                    # Determine the width and height of the line
                    (lineWidth, lineHeight) = dc.GetTextExtent(line)
                    # A blank line has no font height here, so needs an alternate "lineHeight" value
                    if line == '':
                        lineHeight = self.styles[style]['font'].GetPointSize() + int(DPI / 16)
                    # If this text element has the largest height, use that for the vertical increment (yInc)
                    if lineHeight > yInc:
                        yInc = lineHeight
                    # Align and position the text, then add it to the Printout
                    if self.styles[style]['align'] == mstyLEFT:
                        dc.DrawText(line, DPI + xindent, yPos)
                    elif self.styles[style]['align'] == mstyCENTER:
                        dc.DrawText(line, int(graphicX/2.0 - lineWidth/2.0), yPos)
                    elif self.styles[style]['align'] == mstyRIGHT:
                        dc.DrawText(line, graphicX - DPI - lineWidth, yPos)
                # Increment the vertical position indicator
                yPos = yPos + yInc

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

            # Initialize and fill a sample data structure that will be turned into a report
            self.data = []
            self.data.append((('Heading', 'This is Line 1, using "Heading" style'),))
            self.data.append((('Subheading', 'This is Line 2, using "Subheading" style'),))
            self.data.append((('Subtext', 'This is Line 3, using "Subtext" style'),))
            self.data.append((('Subtext', "This is Line 4, using 'Subtext' style.  It's a really, really long line that will need to be parsed because it's too long to fit on a single line in the final printed version of the stupid thing.  I mean, this line is really, really, really long.  I mean it.  Really."),))
            self.data.append((('Normal', 'This is Line 5, using "Normal" style'),))
            self.data.append((('NormalCenter', 'This is Line 6, using "NormalCenter" style'),))
            self.data.append((('NormalRight', 'This is Line 7, using "NormalRight" style'),))
            self.data.append((('Normal', 'Line 8, Part 1, Left'), ('NormalCenter', 'Line 8, Part 2, Center'), ('NormalRight', 'Line 8, Part 3, "NormalRight" style')))
            self.data.append((('Heading', 'This is Line 9, using "Heading" style'),))
            self.data.append((('Normal', 'This is Line 10, using "Normal" style'),))
            self.data.append((('Heading', 'This is Line 11, using "Heading" style'),))
            self.data.append((('Normal', 'This is Line 12, using "Normal" style'),))
            self.data.append((('Normal', ''),))
            for x in range(14, 121):
                self.data.append((('Normal', 'This is Line %d, using "Normal" style' % x),))

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
            printerDialog.Destroy()


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
            (self.graphic, self.pageData) = PrepareData(self.printData, self.title, self.data, subtitle)

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
            (self.graphic, self.pageData) = PrepareData(self.printData, self.title, self.data, subtitle)
            # Send the results of the PrepareData() call to the MyPrintout object
            printout = MyPrintout(self.title, self.graphic, self.pageData, subtitle)
            # send the MyPrintout object to the Printer
            if not printer.Print(self, printout):
                dlg = Dialogs.ErrorDialog(None, _("There was a problem printing this report."))
                dlg.ShowModal()
                dlg.Destoy()
            else:
                # Save any changes that may have been made to the Printer Setup
                self.printData = printer.GetPrintDialogData().GetPrintData()
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
 
