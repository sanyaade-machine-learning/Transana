#Copyright (C) 2002 - 2007 The Board of Regents of the University of Wisconsin System

#This program is free software; you can redistribute it and/or
#modify it under the terms of the GNU General Public License
#as published by the Free Software Foundation; either version 2
#of the License, or (at your option) any later version.

#This program is distributed in the hope that it will be useful,
#but WITHOUT ANY WARRANTY; without even the implied warranty of
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#GNU General Public License for more details.

#You should have received a copy of the GNU General Public License
#along with this program; if not, write to the Free Software
#Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

# This code is taken from the wxPython Demo file "PrintFramework.py" and
# has been modified by David Woods

""" This unit creates a Print / Print Preview Framework for the Keyword Map and Series Map. """

# import wxPython
import wx

class MyPrintout(wx.Printout):
    """ This class creates and displays the Keyword Map """
    def __init__(self, title, canvas, multiPage=False, lineStart=0, lineHeight=0):
        """ Initialization Function parameters:
              title            Title for the Report
              canvas           The GraphicsControlClass object that contains the graphic object to be printed.
                               (Must have a DrawLines() function that causes the image to be completed.)
              multiPage=False  Indicates that the graphic should be printed across multiple pages rather than compressed.
              lineStart=0      Position of the first line in the graphic.
              lineHeight=0     The size of each line in the graphic.  """
        # Create the Print Preview dialog
        wx.Printout.__init__(self, title)
        # Remember the graphic canvas for later use
        self.canvas = canvas
        # Save parameter data for later use
        self.multiPage = multiPage
        self.lineStart = lineStart
        self.lineHeight = lineHeight

    def OnPreparePrinting(self):
        """ Prepare data for printing / previewing """
        # If we have a multi-page report, we need to create all of the pages here
        if self.multiPage:
            # Determine the dimensions of the Print Device Context
            (dcWidth, dcHeight) = self.GetDC().GetSizeTuple()
            # Determine the dimensions for the graphics canvas passed in
            canvasWidth = self.canvas.getWidth()
            canvasHeight = self.canvas.getHeight()

            # The header area to be repeated on each page for the Series Map reports is 48 pixels high.  It just is.
            headerHeight = 48
            # Get a graphic of the whole report image.
            # Create a Bitmap to draw on
            bigBMP = wx.EmptyBitmap(canvasWidth, canvasHeight)
            # Create a Device Context for the Bitmap
            bigDC = wx.BufferedDC(None, bigBMP)
            # Indicate that we want a solid background, rather than transparent
            bigDC.SetBackgroundMode(wx.SOLID)
            # Set the Background to White
            bigDC.SetBackground(wx.Brush(wx.WHITE))
            # Clear the DC
            bigDC.Clear()
            # Paint the background
            bigDC.FloodFill(1, 1, wx.WHITE)
            # Draw the report object on the DC
            self.canvas.DrawLines(bigDC)
            # Saving the bitmap as a graphic can be helpful in debugging!
            # bigBMP.SaveFile('c:\Documents and Settings\davidwoods\Desktop\TestImage.jpg', wx.BITMAP_TYPE_JPEG)

            # Create a data structure to hold the page images
            self.pageBMPs = []
            # Initialize a variable for counting the number of pages in the report
            self.numPages = 0
            # We need to count lines in the report to find page break positions.  Initialize the line counter.
            lineCount = 0
            # We need to track what pixel in the whole graphic was used for the last page break.
            # We initialize this to the starting position of the first line in the report.
            startingPixel = self.lineStart
            # Now let's create seperate images for each printed page.  We keep going until we have
            # processed all the pixels in the graphic canvas.
            while startingPixel < canvasHeight:
                # Increment the page counter.
                self.numPages += 1
                # Calculate the default height for this page, not including the space used for the page header.
                pageHeight = int((float(canvasWidth) / dcWidth) * dcHeight) - headerHeight
                # Create a full-size Bitmap image to hold the graphic for the individual page, including the page header.
                pageBMP = wx.EmptyBitmap(canvasWidth, pageHeight + headerHeight)
                
                # The problem we have to solve here is to figure out EXACTLY where in the graphic to break the page.
                # We don't want to break in the middle of a line, so we have to figure out where the space between lines
                # falls.
                # To start, initialize a page height variable to 0.
                newPageHeight = 0
                # We can move through the graphic one whole line at a time by incrementing the new page height by
                # the line height.  We do this as long as there is room on the page.
                while self.lineStart + lineCount * self.lineHeight < startingPixel + pageHeight + headerHeight - self.lineStart - self.lineHeight:
                    # While there's room, increment the line counter ...
                    lineCount += 1
                    # ... and add the height of the line to the new page height.
                    newPageHeight += self.lineHeight
                # Create a Device Context that will allow us to manipulate the Bitmap for this individual page
                pageDC = wx.BufferedDC(None, pageBMP)
                # Set the Background to White
                pageDC.SetBackground(wx.Brush(wx.WHITE))
                # Clear the DC
                pageDC.Clear()
                # Paint the Bitmap's background White, which is probably only needed on the last page!
                pageDC.FloodFill(1, 1, wx.WHITE)
                # Put the Page Header from the top of the whole graphic on the top of the page graphic
                pageDC.Blit(0, 0, canvasWidth, headerHeight, bigDC, 0, 0)
                # Put the next page-sized segment of the whole graphic onto the page graphic
                pageDC.Blit(0, headerHeight, canvasWidth, newPageHeight, bigDC, 0, startingPixel)
                # We want to add the page number to the page.  Start by creating the text label.
                pageNumText = _("Page %d") % self.numPages
                # Determine the size of the text label on the graphic to aid in positioning
                (textWidth, textHeight) = pageDC.GetTextExtent(pageNumText)
                # Begin drawing on the Page device context.
                pageDC.BeginDrawing()
                # Add the page number text in the lower right-hand corner of the page.
                pageDC.DrawText(pageNumText, canvasWidth - textWidth, pageHeight + headerHeight - textHeight)
                # End drawing on the Page device context.
                pageDC.EndDrawing()

                # Move the pointer to the whole graphic's position down by the length of this page 
                startingPixel += newPageHeight
                # Store the page image in a memory structure where we can get to it later
                self.pageBMPs.append(pageBMP)

                # Saving the bitmap as a graphic can be helpful in debugging!
                # pageBMP.SaveFile('c:\Documents and Settings\davidwoods\Desktop\TestImage%s.jpg' % page, wx.BITMAP_TYPE_JPEG)
        # If we're dealing with a single-page document, such as a Keyword Map ...
        else:
            # ... then we know we only have one page!
            self.numPages = 1

        # Call the base class' OnPreparePrinting() method, which we must do differently in different wxPython versions
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            self.base_OnPreparePrinting()
        else:
            super(MyPrintout, self).OnPreparePrinting()

    def OnBeginPrinting(self):
        # Call the base class' OnBeginPrinting() method, which we must do differently in different wxPython versions
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            self.base_OnBeginPrinting()
        else:
            super(MyPrintout, self).OnBeginPrinting()

    def OnEndPrinting(self):
        # Call the base class' OnEndPrinting() method, which we must do differently in different wxPython versions
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            self.base_OnEndPrinting()
        else:
            super(MyPrintout, self).OnEndPrinting()

    def OnBeginDocument(self, start, end):
        # Call the base class' OnBeginDocument() method, which we must do differently in different wxPython versions
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            return self.base_OnBeginDocument(start, end)
        else:
            return super(MyPrintout, self).OnBeginDocument(start, end)

    def OnEndDocument(self):
        # Call the base class' OnEndDocument() method, which we must do differently in different wxPython versions
        if wx.VERSION[0] == 2 and wx.VERSION[1] <= 6:
            self.base_OnEndDocument()
        else:
            super(MyPrintout, self).OnEndDocument()

    def HasPage(self, page):
        """ This function tells the wxPython Print Framework if the specified page has been defined. """
        # If the desired page is less than or equal to the defined number of pages, return True.  Otherwise, False.
        if page <= self.numPages:
            return True
        else:
            return False

    def GetPageInfo(self):
        """ This function tells the wxPython Print Framework how many pages there are. """
        return (1, self.numPages, 1, self.numPages)

    def OnPrintPage(self, page):
        """ Processing necessary for presenting a single page to the Preview or Printer """
        # Get the Printer Device Context
        dc = self.GetDC()

        # One possible method of setting scaling factors... (Modeled on the wxPython PrintFramework Demo.)

        # If we are set up to display multiple pages ...
        if self.multiPage:
            # ... start with the dimensions of a single page.  (All pages are the same size.)
            pageX = self.pageBMPs[0].GetWidth()
            pageY = self.pageBMPs[0].GetHeight()
        # If we only have one page ...
        else:
            # ... start with the dimensions of the graphic canvas.
            pageX = self.canvas.getWidth()
            pageY = self.canvas.getHeight()

        # Let's have at least 50 device units margin
        marginX = 50
        marginY = 50

        # Add the margin to the graphic size
        maxX = pageX + (2 * marginX)
        maxY = pageY + (2 * marginY)

        # Get the size of the DC in pixels
        (w, h) = dc.GetSizeTuple()

        # Calculate a suitable scaling factor
        scaleX = float(w) / maxX
        scaleY = float(h) / maxY

        # Use x or y scaling factor, whichever fits on the DC
        actualScale = min(scaleX, scaleY)

        # Calculate the position on the DC for centring the graphic
        posX = (w - (pageX * actualScale)) / 2.0
        posY = (h - (pageY * actualScale)) / 2.0

        # Set the scale and origin
        dc.SetUserScale(actualScale, actualScale)
        dc.SetDeviceOrigin(int(posX), int(posY))

        # If we have multiple pages ...
        if self.multiPage:
            # ... create a Buffered Device Context for THIS page from the prepared page graphics ...
            pageDC = wx.BufferedDC(None, self.pageBMPs[page - 1])
            # ... and copy the graphic information onto the Printer Device Context using the page graphic's dimensions.
            dc.Blit(0, 0, pageX, pageY, pageDC, 0, 0)
        # If we have a single page ...
        else:
            # ... we actually populate the graphic device context (ie. draw the graph) here.
            self.canvas.DrawLines(dc)

        return True
