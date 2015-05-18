#Copyright (C) 2002  The Board of Regents of the University of Wisconsin System

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

# This code is derived from the wxPython Demo file "PrintFramework.py" and
# has been modified by David Woods

""" This unit creates a Print / Print Preview Framework for the Keyword Map. """

from wxPython.wx         import *

#----------------------------------------------------------------------

class MyPrintout(wxPrintout):
    """ This class creates and displays the Keyword Map """
    def __init__(self, canvas):
        wxPrintout.__init__(self, _('Transana Keyword Map'))
        self.canvas = canvas

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
        # HasPage function has 1 page limit hardcoded!!
        if page <= 1:
            return true
        else:
            return false

    def GetPageInfo(self):
        # GetPageInfo function ???
        return (1, 1, 1, 1)

    def OnPrintPage(self, page):
        dc = self.GetDC()

        #-------------------------------------------
        # One possible method of setting scaling factors...

        maxX = self.canvas.getWidth()
        maxY = self.canvas.getHeight()

        # Let's have at least 50 device units margin
        marginX = 50
        marginY = 50

        # Add the margin to the graphic size
        maxX = maxX + (2 * marginX)
        maxY = maxY + (2 * marginY)

        # Get the size of the DC in pixels
        (w, h) = dc.GetSizeTuple()

        # Calculate a suitable scaling factor
        scaleX = float(w) / maxX
        scaleY = float(h) / maxY

        # Use x or y scaling factor, whichever fits on the DC
        actualScale = min(scaleX, scaleY)

        # Calculate the position on the DC for centring the graphic
        posX = (w - (self.canvas.getWidth() * actualScale)) / 2.0
        posY = (h - (self.canvas.getHeight() * actualScale)) / 2.0

        # Set the scale and origin
        dc.SetUserScale(actualScale, actualScale)
        dc.SetDeviceOrigin(int(posX), int(posY))

        self.canvas.DrawLines(dc)

        return true

