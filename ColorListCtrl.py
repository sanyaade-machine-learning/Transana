# Copyright (C) 2003 - 2008 The Board of Regents of the University of Wisconsin System 
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

""" This module implements a ColorListBox custom control for Transana. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx
# import the GenStaticBitmap widget, at Robin Dunn's suggestion
import wx.lib.statbmp

import TransanaGlobal

class ColorListCtrl(wx.Panel):
    """ This class is like a CheckListCtrl, but displays a COLOR box for the check box. """
    # While the ColorListCtrl is actually a panel, it needs to look and act like a ListCtrl to the parent.
    def __init__(self, parent):
        # Initialize a data structure to keep track of whether items are selected or deselected
        self.itemChecks = []
        # Create a panel to hold all the controls for the ColorListCtrl.  
        wx.Panel.__init__(self, parent, -1)
        # Create the main Sizer
        vSizer = wx.BoxSizer(wx.VERTICAL)
        # Add a spacer to allow space above the instructions
        vSizer.Add((1,4))
        # Create instructions and add to the sizer
        instructions = wx.StaticText(self, -1, _('Click the color box to select/unselect a keyword.  A color is "selected," white is "unselected."'))
        vSizer.Add(instructions, 0, wx.ALL, 2)
        # Create a ListCtrl to hold the items
        self.lc = wx.ListCtrl(self, -1, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        vSizer.Add(self.lc, 1, wx.EXPAND)
        # Bind the ListItemSelected method to the ListCtrl
        self.lc.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnListItemSelected)
        # Add a spacer to allow space above the instructions
        vSizer.Add((1,4))
        # Add text describing the Color Selection squares
        colorText = wx.StaticText(self, -1, _("Click keyword text above, then choose a color below to change the keyword's display color."))
        vSizer.Add(colorText, 0, wx.ALL, 2)

        # Create the keyword Color images
        # Start by initializing the ImageHandlers
        wx.InitAllImageHandlers()
        # Create an image list for 16 x 16 pixel images
        self.imageList = wx.ImageList(16,16)
        # Create a horizontal sizer to hold the color images
        colorHSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Create a dictionary that will allow the lookup of what actual color is associated with what color graphic
        self.colorIDs = {}
        # We need to add White to the list of Keyword Map colors for the blank item!
        tempColorList = TransanaGlobal.keywordMapColourSet + ['White']
        # Iterate through the defined colors for use with the Keyword Map.
        for colName in tempColorList:
            # Determine the color value that goes along with the color name
            colRGB = TransanaGlobal.transana_colorLookup[colName]
            # Create an empty bitmap
            bmp = wx.EmptyBitmap(16, 16)
            # Create a Device Context for manipulating the bitmap
            dc = wx.BufferedDC(None, bmp)
            # Begin the drawing process
            dc.BeginDrawing()
            # Paint the bitmap white
            dc.SetBackground(wx.Brush(wx.Colour(255, 255, 255)))
            # Clear the device context
            dc.Clear()
            # Define the pen to draw with
            pen = wx.Pen(wx.Colour(0, 0, 0), 1, wx.SOLID)
            # Set the Pen for the Device Context
            dc.SetPen(pen)
            # Define the brush to paint with in the defined color
            brush = wx.Brush(wx.Colour(colRGB[0], colRGB[1], colRGB[2]))
            # Set the Brush for the Device Context
            dc.SetBrush(brush)
            # Draw a black border around the color graphic, leaving a little white space
            dc.DrawRectangle(0, 2, 14, 14)
            # End the drawing process
            dc.EndDrawing()
            # Select a different object into the Device Context, which allows the bmp to be used.
            dc.SelectObject(wx.EmptyBitmap(5,5))
            # If we have a valid bitmap ...
            if bmp:
                # ... add that bitmap to the image list
                self.imageList.Add(bmp)

        # If there are fewer than 375 colors, we can do 25 colors per row.  Otherwise, there needs to be a maximum of
        # 15 rows so the form looks okay.  
        colorsPerRow = max((len(TransanaGlobal.transana_graphicsColorList) + 14) / 15, 25)
        # Initial a counter to -1, since our first element is  0
        counter = -1
        # Now iterate through the color list in a different order.  We want the colors to appear on
        # the form in RGB COLOR order, which groups similar colors, not the Keyword Map Color List
        # order, which separates similar colors.
        for (colName, colRGB) in TransanaGlobal.transana_graphicsColorList:
            # Not quite all the defined colors are in the Keyword List Colors.  We need the subset.
            if colName in tempColorList:
                # Get the correct bitmap from the imageList.
                bmp = self.imageList.GetBitmap(tempColorList.index(colName))
                # Create a Widget for the bitmap.  Use GenStaticBitmap rather than StaticBitmap so that
                # the mouse functions work on Mac.
                bmpCtrl = wx.lib.statbmp.GenStaticBitmap(self, -1, bmp)
                # Get the widget's ID and create a lookup entry that will give the color name for this widget
                self.colorIDs[bmpCtrl.GetId()] = colName
                # Bind the Mouse Left Up event.  This will allow us to catch mouse clicks for color selection.
                bmpCtrl.Bind(wx.EVT_LEFT_UP, self.OnLeftUp)
                # Increment the counter
                counter += 1
                # Let's split the color blocks into multiple rows by completing and re-starting the sizer
                # every time the counter hits the specified value
                if counter % colorsPerRow == 0:
                    # Add the color selection bar to the Panel's main sizer
                    vSizer.Add(colorHSizer, 0, wx.ALL, 2)
                    # Create a horizontal sizer to hold the color images
                    colorHSizer = wx.BoxSizer(wx.HORIZONTAL)
                # Add the widget to the sizer
                colorHSizer.Add(bmpCtrl, 0, wx.RIGHT, 2)
        # Add the color selection bar to the Panel's main sizer
        vSizer.Add(colorHSizer, 0, wx.ALL, 2)
        # Add the ImageList to the ListCtrl to make the images accessible
        self.lc.SetImageList(self.imageList, wx.IMAGE_LIST_SMALL)
        # Set the panel's main sizer and handle layout.
        self.SetSizer(vSizer)
        self.SetAutoLayout(True)
        self.Layout()

    def CheckItem(self, itemNum):
        """ "Check" an item.  Gives ColorListCtrl functionality similar to the CheckListCtrl mixin. """
        # Our itemChecks list is only maintained to the highest "True" value, and may be smaller than the ListCtrl.
        # If the item to be checked is beyond the end of the current lists...
        while len(self.itemChecks) <= itemNum:
            # ... then pad the list with "False" value to the appropriate size
            self.itemChecks.append(False)
        # Set the indicated item to True
        self.itemChecks[itemNum] = True
        # Set the item's image.  The appropriate image index is stored in the ListCtrl's internal ItemData value.
        self.lc.SetItemImage(itemNum, self.lc.GetItemData(itemNum))

    def DeleteAllItems(self):
        """ Delete all items in the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        # Delete all item check data
        self.itemChecks = []
        # Delete all items from the ListCtrl
        return self.lc.DeleteAllItems()

    def DeleteItem(self, itemNum):
        """ Delete an item from the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        # Remove the entry from the item check data
        del(self.itemChecks[itemNum])
        # Remove the item from the ListCtrl
        return self.lc.DeleteItem(itemNum)

    def EnsureVisible(self, itemNum):
        """ Ensure Visibility of a ListCtrl item.  (Gives our Panel ListCtrl functionality.) """
        return self.lc.EnsureVisible(itemNum)

    def GetItem(self, itemNum, colNum):
        """ Get a ListCtrl Item.  (Gives our Panel ListCtrl functionality.) """
        return self.lc.GetItem(itemNum, colNum)

    def GetItemCount(self):
        """ Count the items in the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        return self.lc.GetItemCount()

    def GetItemData(self, itemNum):
        """ Get Item Data for the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        return self.lc.GetItemData(itemNum)

    def GetNextItem(self, selItem, geometry, state):
        """ Get Next Item in the ListCtrl that matches the parameters.  (Gives our Panel ListCtrl functionality.) """
        return self.lc.GetNextItem(selItem, geometry, state)

    def InsertColumn(self, colNum, colTitle):
        """ Insert a Column in the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        return self.lc.InsertColumn(colNum, colTitle)

    def InsertStringItem(self, itemNum, itemStr):
        """ Insert a new String Item in the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        # Insert "False" as the default Item Check value
        self.itemChecks.insert(itemNum, False)
        # Insert the item in the ListCtrl
        return self.lc.InsertStringItem(itemNum, itemStr)
        
    def IsChecked(self, itemNum):
        """ Determine if an item is checked.  Gives ColorListCtrl functionality similar to the CheckListCtrl mixin. """
        # If the item is in the itemChecks list ...
        if len(self.itemChecks) > itemNum:
            # ... return the item value.
            return self.itemChecks[itemNum]
        # if it's NOT in the itemChecks list ...
        else:
            # ... then we know it's False.
            return False

    def OnLeftUp(self, event):
        """ Triggered upon release of the left mouse button in one of the Color Selection boxes """
        # The options include the defined Keyword Map colors plus white (for unselected)
        colorList = TransanaGlobal.keywordMapColourSet + ['White']
        # Determine what item (if any) is currently selected
        itemNum = self.lc.GetNextItem(-1, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
        # If an item is selected and a color is defined ...
        if (itemNum > -1):
            # White means to deselect the item ...
            if self.colorIDs[event.GetId()] == 'White':
                # ... so if the item is checked ...
                if self.IsChecked(itemNum):
                    # ... then uncheck it
                    self.ToggleItem(itemNum)
            # Any other color means to select that color for that item.
            else:
                # Make sure the item is checked ...
                self.CheckItem(itemNum)
                # ... and set the selected color as the Item Data.
                self.SetItemData(itemNum, colorList.index(self.colorIDs[event.GetId()]))

    def OnListItemSelected(self, event):
        """ Triggered by the selection of an item in the ListCtrl """
        # Call the underlying ListCtrl's OnListItemSelected method
        event.Skip()
        # Determine the item number of the selected item
        itemNum = self.lc.GetNextItem(-1, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
        # Get the mouse position on the screen
        (windowx, windowy) = wx.GetMousePosition()
        # Translate the mouse position to be relative to the ColorListCtrl
        pt = self.ScreenToClientXY(windowx, windowy)
        # If the mouse is on the color image associated with the list item ...
        if (pt[0] >= 6) and (pt[0] < 20):
            # ... toggle the time to change its "Checked" status ...
            self.ToggleItem(itemNum)
            # ... and de-select the item.  (You can't "check/uncheck" the selected item, so changing the "check"
            # back and forth without this was too hard.)
            self.SetItemState(itemNum, 0, wx.LIST_STATE_SELECTED)

    def SetColumnWidth(self, colNum, width):
        """ Set a column's width in the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        return self.lc.SetColumnWidth(colNum, width)

    def SetItemData(self, itemNum, dataVal):
        """ Set an item's Item Data in the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        # Set the Item Data in the ListCtrl
        self.lc.SetItemData(itemNum, dataVal)
        # If the item is "checked" ...
        if self.IsChecked(itemNum):
            # ... set the item's image.  The appropriate image index is stored in the ListCtrl's internal ItemData value.
            self.lc.SetItemImage(itemNum, self.lc.GetItemData(itemNum))
        # If the item is "unchecked" ...
        else:
            # ... set the item's image to blank, the last image in the imageList.
            self.lc.SetItemImage(itemNum, self.lc.GetImageList(wx.IMAGE_LIST_SMALL).GetImageCount() - 1)

    def SetItemState(self, itemNum, state, stateMask):
        """ Set an item's State in the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        return self.lc.SetItemState(itemNum, state, stateMask)

    def SetStringItem(self, itemNum, colNum, itemStr):
        """ Set a column's string value for the appropriate item in the ListCtrl.  (Gives our Panel ListCtrl functionality.) """
        return self.lc.SetStringItem(itemNum, colNum, itemStr)

    def ToggleItem(self, itemNum):
        """ Toggle an item's "checked" status.  Gives ColorListCtrl functionality similar to the CheckListCtrl mixin. """
        # If the list of itemChecks is too short ...
        if len(self.itemChecks) < itemNum:
            # ... we know it is currently false and can make it "True" (adjusting the length of the list automatically)
            # by calling CheckItem()
            self.CheckItem(itemNum)
        # If the item is already in the list ...
        else:
            # ... swap the value of itemChecks ...
            self.itemChecks[itemNum] = not self.itemChecks[itemNum]
            # If the item is now True ...
            if self.itemChecks[itemNum]:
                # ... set the item's image.  The appropriate image index is stored in the ListCtrl's internal ItemData value.
                self.lc.SetItemImage(itemNum, self.GetItemData(itemNum))
            # If the item is now False ...
            else:
                # ... set the item's image to blank, the last image in the imageList.
                self.lc.SetItemImage(itemNum, self.lc.GetImageList(wx.IMAGE_LIST_SMALL).GetImageCount() - 1)
