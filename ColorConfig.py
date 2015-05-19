# Copyright (C) 2003 - 2009 The Board of Regents of the University of Wisconsin System 
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

""" This dialog implements the Transana Color Configuration class.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx
# import wxPython's Color Selector
import wx.lib.colourselect
# import Transana's common dialogs
import Dialogs
# import the TransanaGlobal variables
import TransanaGlobal
# Import Python's os and sys modules
import os, sys

class ColorConfig(wx.Dialog):
    """ Transana's custom Color Configuration Dialog Box.  """
    def __init__(self, parent):
        """ Initialize the Color Configuration Dialog Box. """
        # initialize the current item to the top of the list
        self.currentItem = 0
        # Get a local copy of the graphics color list from TransanaGlobal
        self.colors = TransanaGlobal.transana_graphicsColorList
        # Get the configuration filename
        self.configFile = TransanaGlobal.configData.colorConfigFilename

        # Create the Font Dialog
        wx.Dialog.__init__(self, parent, -1, _('Graphics Color Configuration'), style=wx.CAPTION | wx.SYSTEM_MENU | wx.THICK_FRAME)
        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the main Sizer, which will hold the boxTop, boxMiddle, and boxButton sizers
        box = wx.BoxSizer(wx.VERTICAL)

        # Add the name of the Color Configuration File being used (Row 0)
        self.lblFileName = wx.StaticText(self, -1, self.configFile)
        box.Add(self.lblFileName, 0, wx.LEFT | wx.TOP, 10)

        # Create the sizer for Row 1
        boxRow1 = wx.BoxSizer(wx.HORIZONTAL)

        # Create a List Control for the Color List
        self.colorList = wx.ListCtrl(self, -1, size=(200, 350), style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        # Define the Item Selected event for the List Control
        self.colorList.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnItemSelected)
        # Add the List Control to the Sizer
        boxRow1.Add(self.colorList, 9, wx.EXPAND | wx.ALL, 10)

        # Create the sizer for the second Column in Row 1
        boxR1C2 = wx.BoxSizer(wx.VERTICAL)
        
        # Add an expandable top spacer
        boxR1C2.Add((1, 1), 1, wx.EXPAND)
        
        # Get the graphic for the Move Up button
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_UP, wx.ART_TOOLBAR, (16,16))
        # Create a bitmap button for the Move Up button
        self.btnUp = wx.BitmapButton(self, -1, bmp)
        # Set the Tool Tip for the Move Up button
        self.btnUp.SetToolTipString(_("Move up"))
        # Bind the button event for the Move Up button
        self.btnUp.Bind(wx.EVT_BUTTON, self.OnUpDown)
        # Add Move Up to the Sizer
        boxR1C2.Add(self.btnUp, 0, wx.ALIGN_CENTER | wx.RIGHT | wx.BOTTOM, 10)
        
        # Get the graphic for the Move Down button
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_DOWN, wx.ART_TOOLBAR, (16,16))
        # Create a bitmap button for the Move Down button
        self.btnDown = wx.BitmapButton(self, -1, bmp)
        # Set the Tool Tip for the Move Down button
        self.btnDown.SetToolTipString(_("Move down"))
        # Bind the button event for the Move Down button
        self.btnDown.Bind(wx.EVT_BUTTON, self.OnUpDown)
        # Add Move Down to the Sizer
        boxR1C2.Add(self.btnDown, 0, wx.ALIGN_CENTER | wx.RIGHT, 10)
        
        # Add an expandable Spacer to the sizer.  This allows the Up and Down buttons to center vertically.
        boxR1C2.Add((1, 1), 1, wx.EXPAND)
        # Add the Up and Down buttons to the Row 1 sizer
        boxRow1.Add(boxR1C2, 0, wx.EXPAND)

        # Add Row 1 to the Form's Main sizer
        box.Add(boxRow1, 9, wx.EXPAND | wx.GROW)

        # Create a sizer for the second row
        boxRow2 = wx.BoxSizer(wx.HORIZONTAL)
        
        # Add the color selector as Column 1
        boxR2C1 = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(self, -1, _("Color:"))
        boxR2C1.Add(lbl, 0, wx.LEFT, 10)
        # Create a Color Selector
        self.csColorSelect = wx.lib.colourselect.ColourSelect(self, -1, "", (128, 128, 128), size=wx.DefaultSize)
        # Bind the Color Selection event handler
        self.csColorSelect.Bind(wx.lib.colourselect.EVT_COLOURSELECT, self.OnColorSelect)
        boxR2C1.Add(self.csColorSelect, 0, wx.LEFT, 10)
        boxRow2.Add(boxR2C1, 0)

        # Add the Color Name as Column 2
        boxR2C2 = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(self, -1, _("Color Name:"))
        boxR2C2.Add(lbl, 0, wx.LEFT, 10)
        self.txtColorName = wx.TextCtrl(self, -1, "")
        boxR2C2.Add(self.txtColorName, 10, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
        boxRow2.Add(boxR2C2, 10, wx.EXPAND)

        # Add the Color Hex value as Column 3
        boxR2C3 = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(self, -1, _("Hex:"))
        boxR2C3.Add(lbl, 0, wx.LEFT, 10)
        self.txtColorHex = wx.TextCtrl(self, -1, "", size=(80, 20))
        # Bind a KillFocus event handler to validate the Hex value on leaving this field
        self.txtColorHex.Bind(wx.EVT_KILL_FOCUS, self.OnHexKillFocus)
        boxR2C3.Add(self.txtColorHex, 3, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
        boxRow2.Add(boxR2C3, 3, wx.EXPAND)

        # Add the Red value as Column 4
        boxR2C4 = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(self, -1, _("Red:"))
        boxR2C4.Add(lbl, 0, wx.LEFT, 10)
        self.txtColorRed = wx.TextCtrl(self, -1, "", size=(40, 20))
        # Bind a KillFocus event handler to validate the Red value on leaving this field
        self.txtColorRed.Bind(wx.EVT_KILL_FOCUS, self.OnRGBKillFocus)
        boxR2C4.Add(self.txtColorRed, 2, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
        boxRow2.Add(boxR2C4, 2, wx.EXPAND)

        # Add the Green value as Column 5
        boxR2C5 = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(self, -1, _("Green:"))
        boxR2C5.Add(lbl, 0, wx.LEFT, 10)
        self.txtColorGreen = wx.TextCtrl(self, -1, "", size=(40, 20))
        # Bind a KillFocus event handler to validate the Green value on leaving this field
        self.txtColorGreen.Bind(wx.EVT_KILL_FOCUS, self.OnRGBKillFocus)
        boxR2C5.Add(self.txtColorGreen, 2, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
        boxRow2.Add(boxR2C5, 2, wx.EXPAND)

        # Add the Blue value as Column 6
        boxR2C6 = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(self, -1, _("Blue:"))
        boxR2C6.Add(lbl, 0, wx.LEFT, 10)
        self.txtColorBlue = wx.TextCtrl(self, -1, "", size=(40, 20))
        # Bind a KillFocus event handler to validate the Blue value on leaving this field
        self.txtColorBlue.Bind(wx.EVT_KILL_FOCUS, self.OnRGBKillFocus)
        boxR2C6.Add(self.txtColorBlue, 2, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
        boxRow2.Add(boxR2C6, 2, wx.EXPAND)

        # Add the second row to the Main Sizer
        box.Add(boxRow2, 0, wx.EXPAND)

        # Create the boxButtons sizer, which will hold the dialog box's buttons
        boxButtons = wx.BoxSizer(wx.HORIZONTAL)

        # Get the image for File Open
        bmp = wx.ArtProvider_GetBitmap(wx.ART_FILE_OPEN, wx.ART_TOOLBAR, (16,16))
        # Create the File Open button
        self.btnFileOpen = wx.BitmapButton(self, -1, bmp)
        self.btnFileOpen.SetToolTip(wx.ToolTip(_('Open Color Configuration')))
        # Bind the Button to the appropriate Event handler
        self.btnFileOpen.Bind(wx.EVT_BUTTON, self.OnFileOpen)
        boxButtons.Add(self.btnFileOpen, 0, wx.ALIGN_LEFT | wx.LEFT | wx.TOP, 10)

        # Get the image for File Save
        bmp = wx.ArtProvider_GetBitmap(wx.ART_FILE_SAVE, wx.ART_TOOLBAR, (16,16))
        # Create the File Save button
        self.btnFileSave = wx.BitmapButton(self, -1, bmp)
        self.btnFileSave.SetToolTip(wx.ToolTip(_('Save Color Configuration')))
        # Bind the button to the appropriate event handler
        self.btnFileSave.Bind(wx.EVT_BUTTON, self.OnFileSave)
        boxButtons.Add(self.btnFileSave, 0, wx.ALIGN_LEFT | wx.LEFT | wx.TOP, 10)
        
        # Add an expandable spacer to the Button sizer
        boxButtons.Add((1, 1), 1, wx.ALIGN_CENTER | wx.EXPAND)

        # Create the Add button
        self.btnAdd = wx.Button(self, -1, _("Add"))
        # Bind the button to the approriate event handler
        self.btnAdd.Bind(wx.EVT_BUTTON, self.OnAdd)
        boxButtons.Add(self.btnAdd, 0, wx.LEFT | wx.TOP, 10)

        # Create the Edit button
        self.btnEdit = wx.Button(self, -1, _("Edit"))
        # bind the button to the appropriate event handler
        self.btnEdit.Bind(wx.EVT_BUTTON, self.OnEdit)
        boxButtons.Add(self.btnEdit, 0, wx.LEFT | wx.TOP, 10)

        # Create the Delete button
        self.btnDelete = wx.Button(self, -1, _("Delete"))
        # Bind the button to the appropriate event handler
        self.btnDelete.Bind(wx.EVT_BUTTON, self.OnDelete)
        boxButtons.Add(self.btnDelete, 0, wx.LEFT | wx.TOP, 10)

        # Add an expandable spacer to the Button sizer
        boxButtons.Add((1, 1), 1, wx.ALIGN_CENTER | wx.EXPAND)

        # Create an OK button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Set this as the default (for Enter)
        btnOK.SetDefault()
        # Bind the button to the appropriate event handler
        btnOK.Bind(wx.EVT_BUTTON, self.OnOK)
        boxButtons.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT, 10)

        # Create a Cancel button
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        btnCancel.Bind(wx.EVT_BUTTON, self.OnCancel)
        boxButtons.Add(btnCancel, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT, 10)

        # Create a Help button
        btnHelp = wx.Button(self, -1, _("Help"))
        btnHelp.Bind(wx.EVT_BUTTON, self.OnHelp)
        boxButtons.Add(btnHelp, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT, 10)

        # Add the boxButtons sizer to the main box sizer
        box.Add(boxButtons, 0, wx.ALIGN_BOTTOM | wx.EXPAND | wx.BOTTOM, 10)

        # Populate the Color List
        self.PopulateList()

        # Define box as the form's main sizer
        self.SetSizer(box)
        # Fit the form to the widgets created
        self.Fit()
        # Tell the form to maintain the layout and have it set the intitial Layout
        self.SetAutoLayout(True)
        self.Layout()
        # Set this as the minimum size for the form.
        self.SetSizeHints(minW = self.GetSize()[0], minH = self.GetSize()[1])

        # We need an Size event for the form for a little mainenance when the form size is changed
        self.Bind(wx.EVT_SIZE, self.OnSize)
        
        # Under wxPython 2.6.1.0-unicode, this form is throwing a segment fault when the color gets changed.
        # The following variable prevents that!
        self.closing = False

        # Position the form in the center of the screen
        self.CentreOnScreen()


    def PopulateList(self):
        """ Populate the Color List control """
        # Clear the color list
        self.colorList.ClearAll()
        # Create an ImageList object to hold our color sample graphics
        self.imageList = wx.ImageList(16,16)
        # initialize a counter
        counter = 0
        # Initialize a list to hold an index for the color sample images
        self.imageIndex = []
        # Iterate through all the defined colors
        for (colorName, (colorRed, colorGreen, colorBlue)) in self.colors:
            # Create a bitmap of the correct color
            bmp = self.CreateBmp(colorRed, colorGreen, colorBlue)
            # If we have a valid bitmap ...
            if bmp:
                # ... add that bitmap to the image list, simultaneously placing a referent to it in the image index ...
                self.imageIndex.append(self.imageList.Add(bmp))
                # ... and increment the counter
                counter += 1
        # Assign our image list to the List Control
        self.colorList.SetImageList(self.imageList, wx.IMAGE_LIST_SMALL)

        # (Based on the wxPython ListCtrl demo)
        # Since we want images on the column header we have to do it the hard way:
        info = wx.ListItem()
        info.m_mask = wx.LIST_MASK_TEXT | wx.LIST_MASK_IMAGE | wx.LIST_MASK_FORMAT
        info.m_image = -1
        info.m_format = 0
        info.m_text = _("Color")
        self.colorList.InsertColumnInfo(0, info)

        info.m_text = _("Hex")
        self.colorList.InsertColumnInfo(1, info)

        info.m_format = wx.LIST_FORMAT_RIGHT
        info.m_text = _("Red")
        self.colorList.InsertColumnInfo(2, info)
        info.m_text = _("Green")
        self.colorList.InsertColumnInfo(3, info)
        info.m_text = _("Blue")
        self.colorList.InsertColumnInfo(4, info)

        # Initialize a counter
        counter = 0
        # Iterate through our color list again
        for (colorName, (colorRed, colorGreen, colorBlue)) in self.colors:
            # and insert the items into the list this time
            index = self.colorList.InsertImageStringItem(sys.maxint, colorName, self.imageIndex[counter])
            self.colorList.SetStringItem(index, 1, "#%02X%02X%02X" % (colorRed, colorGreen, colorBlue))
            self.colorList.SetStringItem(index, 2, str(colorRed))
            self.colorList.SetStringItem(index, 3, str(colorGreen))
            self.colorList.SetStringItem(index, 4, str(colorBlue))
            # increment out counter
            counter += 1
        # Set our column widths
        self.colorList.SetColumnWidth(0, 170)
        self.colorList.SetColumnWidth(1, 69)
        self.colorList.SetColumnWidth(2, 80)
        self.colorList.SetColumnWidth(3, 80)
        self.colorList.SetColumnWidth(4, 80)

    def OnItemSelected(self, event):
        """ Selection of a Color in the Color List Ctrl """
        # Remember the current selection
        self.currentItem = event.GetIndex()
        # Set the Color Select button to the color of the selected item
        self.csColorSelect.SetColour((int(self.colorList.GetItem(self.currentItem, 2).GetText()),
                                      int(self.colorList.GetItem(self.currentItem, 3).GetText()),
                                      int(self.colorList.GetItem(self.currentItem, 4).GetText())))
        # Set the Color Name field to the color name of the selected item
        self.txtColorName.SetValue(event.GetText())
        # Set the color data values to the color of the selected item
        self.txtColorHex.SetValue(self.colorList.GetItem(self.currentItem, 1).GetText())
        self.txtColorRed.SetValue(self.colorList.GetItem(self.currentItem, 2).GetText())
        self.txtColorGreen.SetValue(self.colorList.GetItem(self.currentItem, 3).GetText())
        self.txtColorBlue.SetValue(self.colorList.GetItem(self.currentItem, 4).GetText())

    def OnUpDown(self, event):
        """ Up or Down Button Press """
        # Remember which button triggered this event, Up or Down
        btnID = event.GetId()
        # If we're moving UP ...
        if btnID == self.btnUp.GetId():
            # ... we want to work on the item ABOVE the current selection
            itemToChange = self.currentItem - 1
        # If we're moving DOWN ...
        elif btnID == self.btnDown.GetId():
            # ... we want to work on the item BELOW the current selection
            itemToChange = self.currentItem + 1
        else:
            print "Unknown button in OnUpDown()"
            return

        # To move items, we must meet the following conditions:
        #   Something in the list must be selected
        #   We can't move the first item UP
        #   We can't move the LAST item (White) UP!
        #   We can't move the last TWO items down.  (White must be LAST in the list.)
        if (self.colorList.GetSelectedItemCount() != 1) or \
           ((self.currentItem == 0) and (btnID == self.btnUp.GetId())) or \
           ((self.currentItem == self.colorList.GetItemCount() - 1) and (btnID == self.btnUp.GetId())) or \
           ((self.currentItem >= self.colorList.GetItemCount() - 2) and (btnID == self.btnDown.GetId())):
            # If we don't meet these conditions, exit this method.
            return

        # Remember the values for the current selection
        currentItemName = self.colorList.GetItem(self.currentItem, 0).GetText()
        currentItemHex = self.colorList.GetItem(self.currentItem, 1).GetText()
        currentItemRed = self.colorList.GetItem(self.currentItem, 2).GetText()
        currentItemGreen = self.colorList.GetItem(self.currentItem, 3).GetText()
        currentItemBlue = self.colorList.GetItem(self.currentItem, 4).GetText()
        # Delete the current selection.  (I tried swapping values in existing items, but couldn't
        # figure out how with the images in the first column.)
        self.colorList.DeleteItem(self.currentItem)

        # Create a new item in the desired new position and populate it with values stored above
        index = self.colorList.InsertImageStringItem(itemToChange, currentItemName, self.imageIndex[self.currentItem])
        self.colorList.SetStringItem(index, 1, currentItemHex)
        self.colorList.SetStringItem(index, 2, currentItemRed)
        self.colorList.SetStringItem(index, 3, currentItemGreen)
        self.colorList.SetStringItem(index, 4, currentItemBlue)

        # Need to swap imageIndex values!
        imgIndex = self.imageIndex[self.currentItem]
        self.imageIndex[self.currentItem] = self.imageIndex[itemToChange]
        self.imageIndex[itemToChange] = imgIndex

        # Now to select the moved position
        self.colorList.SetItemState(itemToChange, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)
        # The current item is now the new selection, where we moved to.
        self.currentItem = itemToChange

    def OnColorSelect(self, event):
        """ Color Select Button Press """
        # If no Color Name has been specified ...
        if self.txtColorName.GetValue() == "":
            # Get the Color Database
            cdb = wx.ColourDatabase()
            # Check to see if the defined value has a Color Name!  (Unlikely, but it happens occasionally.)
            colName = cdb.FindName(event.GetValue()).capitalize()
            # Enter the name in the Name Edit box
            self.txtColorName.SetValue(colName)
        # Enter the color's values into the appropriate edit boxes.
        self.txtColorHex.SetValue("#%02X%02X%02X" % (event.GetValue()[0], event.GetValue()[1], event.GetValue()[2]))
        self.txtColorRed.SetValue(str(event.GetValue()[0]))
        self.txtColorGreen.SetValue(str(event.GetValue()[1]))
        self.txtColorBlue.SetValue(str(event.GetValue()[2]))

    def OnHexKillFocus(self, event):
        """ Handle leaving the Hex Value field """
        # Assume we'll be successful
        success = True
        # Start exception handling
        try:
            # Check for a lenth of 7 with "#" as the first character or a length of 6
            if ((len(self.txtColorHex.GetValue()) == 7) and \
                (self.txtColorHex.GetValue()[0] == '#')) or \
               (len(self.txtColorHex.GetValue()) == 6):
                # Attempt to convert the hex (base 16) values to integers
                colorRed = int(self.txtColorHex.GetValue()[-6:-4], 16)
                colorGreen = int(self.txtColorHex.GetValue()[-4:-2], 16)
                colorBlue = int(self.txtColorHex.GetValue()[-2:], 16)
                # Place the converted values in the Red, Green, and Blue edit boxes
                self.txtColorRed.SetValue(str(colorRed))
                self.txtColorGreen.SetValue(str(colorGreen))
                self.txtColorBlue.SetValue(str(colorBlue))
            # If we don't meet the length criteria ...
            else:
                # ... then we cannot validate the Hex data
                success = False
        # If an exception is raised ...
        except:
            # ... then we cannot validate the Hex data.
            success = False
        # If we have not been successful ...
        if not success:
            # If the field is empty ... 
            if self.txtColorHex.GetValue() == '':
                # ... we have not been successful, but can still move on!
                return
            # ... create and display an error message
            msg = _("Please specify valid color hex value or clear this field.")
            if 'unicode' in wx.PlatformInfo:
                msg = unicode(msg, 'utf8')
            dlg = Dialogs.ErrorDialog(self, msg)
            dlg.ShowModal()
            dlg.Destroy()
            # Set the form focus on the Hex field
            self.txtColorHex.SetFocus()
        # If we have been successful ...
        else:
            # Update the Color Selector button
            self.csColorSelect.SetColour((colorRed, colorGreen, colorBlue))

    def OnRGBKillFocus(self, event):
        """ Handle leaving the Red, Green, or Blue Value fields """
        if event.GetId() == self.txtColorRed.GetId():
            ctrl = self.txtColorRed
        elif event.GetId() == self.txtColorGreen.GetId():
            ctrl = self.txtColorGreen
        elif event.GetId() == self.txtColorBlue.GetId():
            ctrl = self.txtColorBlue
        # Assume we'll be successful
        success = True
        # Start exception handling
        try:
            # Attempt to convert the value to an integer
            color = int(ctrl.GetValue())
            # See if the color is in the range of 0 - 255
            if (color < 0) or (color > 255):
                success = False
        # If an exception is raised ...
        except:
            # ... then we cannot validate the Hex data.
            success = False
        # If we have not been successful ...
        if not success:
            # If the field is empty ... 
            if ctrl.GetValue() == '':
                # ... we have not been successful, but can still move on!
                return
            # ... create and display an error message
            msg = _("Please specify valid color value (between 0 and 255) or clear this field.")
            if 'unicode' in wx.PlatformInfo:
                msg = unicode(msg, 'utf8')
            dlg = Dialogs.ErrorDialog(self, msg)
            dlg.ShowModal()
            dlg.Destroy()
            # Set the form focus on the Hex field
            ctrl.SetFocus()
        # If we have been successful ...
        else:
            try:
                # See if all three values have been defined
                if self.txtColorRed.GetValue() != '':
                    colorRed = int(self.txtColorRed.GetValue())
                else:
                    success = False
                if self.txtColorGreen.GetValue() != '':
                    colorGreen = int(self.txtColorGreen.GetValue())
                else:
                    success = False
                if self.txtColorBlue.GetValue() != '':
                    colorBlue = int(self.txtColorBlue.GetValue())
                else:
                    success = False
                if success:
                    # Update the Color Selector button
                    self.csColorSelect.SetColour((colorRed, colorGreen, colorBlue))
                    # Enter the hex value into the hex edit boxes.
                    self.txtColorHex.SetValue("#%02X%02X%02X" % (colorRed, colorGreen, colorBlue))
            except:
                pass

    def OnFileOpen(self, event):
        """ File Open Button Press """
       # File Filter definition
        fileTypesString = _("Text files (*.txt)|*.txt")
        # If no config file is defined (ie using default colors) ...
        if self.configFile == '':
            # ... use the video path and no filename
            lastPath = TransanaGlobal.configData.videoPath
            filename = ''
        # if a config file is defined ...
        else:
            # ... split it into path and filename
            (lastPath, filename) = os.path.split(self.configFile)
        # Create a File Open dialog.
        fdlg = wx.FileDialog(self, _('Select Color Definition file:'), lastPath, filename, fileTypesString,
                             wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        # Show the dialog and get user response.  If OK ...
        if fdlg.ShowModal() == wx.ID_OK:
            # Remember the file name locally (not globally yet)
            self.configFile = fdlg.GetPath()
            # Update the file name label on screen
            self.lblFileName.SetLabel(self.configFile)
            # Load the colors from the file
            self.colors = TransanaGlobal.getColorDefs(self.configFile)
            # Populate the dialog's color list
            self.PopulateList()
        # Destroy the File Dialog
        fdlg.Destroy()

    def OnFileSave(self, event):
        """ File Save Button Press """
        # If this method is called from the File Save button ...
        if event.GetId() == self.btnFileSave.GetId():
            # Show a File Save dialog.  (Otherwise, save is automatic.)
            # File Filter definition
            fileTypesString = _("Text files (*.txt)|*.txt")
            # If no config file is defined (ie using default colors) ...
            if self.configFile == '':
                # ... use the video path and no filename
                lastPath = TransanaGlobal.configData.videoPath
                filename = ''
            # if a config file is defined ...
            else:
                # ... split it into path and filename
                (lastPath, filename) = os.path.split(self.configFile)
            # Create a File Save dialog.
            fdlg = wx.FileDialog(self, _('Save Color Definition file:'), lastPath, filename, fileTypesString, 
                                 wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
            # Show the dialog and get user response.  If OK ...
            if fdlg.ShowModal() == wx.ID_OK:
                # ... remember the file name
                fs = fdlg.GetPath()
            # If the user cancels ...
            else:
                # signal this with an empty file name
                fs = ''
            # Destroy the File Dialog
            fdlg.Destroy()
        # If we're NOT asking for a file name ...
        else:
            # If there is an existing file name ...
            if self.configFile != '':
                # ... then use it.
                fs = self.configFile
            # If no file name exists ...
            else:
                # place the default file in the video root
                fs = ''   # os.path.join(TransanaGlobal.configData.videoPath, 'TransanaColorDefs.txt')
            
        # If user didn't cancel ..
        if fs != "":
            # If the file doesn't have a "TXT" extension ...
            if (fs[-4:].lower() != '.txt'):
                # ... give it one
                fs = fs + '.txt'
            # On the Mac, if no path is specified, the data is exported to a file INSIDE the application bundle, 
            # where no one will be able to find it.  Let's put it in the user's HOME directory instead.
            # I'm okay with not handling this on Windows, where it will be placed in the Program's folder
            # but it CAN be found.  (There's no easy way on Windows to determine the location of "My Documents"
            # especially if the user has moved it.)
            if "__WXMAC__" in wx.PlatformInfo:
                # if the specified file has no path specification ...
                if fs.find(os.sep) == -1:
                    # ... then prepend the HOME folder
                    fs = os.getenv("HOME") + os.sep + fs
            # Save the Color Defs file.  First, open the file for writing.
            f = file(fs, 'w')
            # Write the header text to the file
            f.write('# Transana Color Definitions file\n\n')
            f.write('# RULES for this file:\n')
            f.write('#\n')
            f.write('# "#" in first column makes a line into a comment, to be ignored entirely\n')
            f.write('#\n')
            f.write("# The first column is the color name.  Names don't matter, except that they MUST be unique.  If you duplicate a name,\n")
            f.write('# the color value of the first instance of that name will be used in all positions for that name.\n')
            f.write('#\n')
            f.write('# The second column is the RED value.  It must be between 0 and 255.\n')
            f.write('#\n')
            f.write('# The third column is the GREEN value.  It must be between 0 and 255.\n')
            f.write('#\n')
            f.write('# The fourth column is the BLUE value.  It must be between 0 and 255.\n')
            f.write('#\n')
            f.write('# The LAST ROW in this table must be "White, 255, 255, 255".  Don' + "'t mess with this.  It gets dropped in some places, \n")
            f.write('# and added back in in some to be the "Don' + "'t include this keyword" + '" value.\n');
            f.write('\n')

            # Initialize the local internal color list
            self.colors = []
            # Iterate through the color defs in the dialog's list item
            for item in range(self.colorList.GetItemCount()):
                # Get the color name from column 0
                colorName = self.colorList.GetItem(item, 0).GetText()
                # Write the color definition line to the data file
                f.write("%s, %s %3s, %3s, %3s\n" % (colorName, '                        '[len(colorName):], self.colorList.GetItem(item, 2).GetText(), self.colorList.GetItem(item, 3).GetText(), self.colorList.GetItem(item, 4).GetText()))
                # Add the color definition to the local internal color list
                self.colors.append((colorName, (int(self.colorList.GetItem(item, 2).GetText()), int(self.colorList.GetItem(item, 3).GetText()), int(self.colorList.GetItem(item, 4).GetText()))))
            # Flush the file's data buffer to disk
            f.flush()
            # Close the data file
            f.close()
            # Remember the file name locally (not globally yet)
            self.configFile = fs
            # Update the file name label on screen
            self.lblFileName.SetLabel(self.configFile)

    def ValidateColorDef(self, testName=False):
        """ Validate the color definition for a new color.  the testName parameter indicates whether we should test
            for duplicate color names. """
        # Begin exception handling
        try:
            # We need to make sure all fields are filled in.
            # Name is required.  EITHER HexValue or 3 color values are required.
            # Get the Color Name
            colorName = self.txtColorName.GetValue()
            # If the name is blank ...
            if (colorName == ''):
                # build and display an error message
                msg = _("Please add a unique color name.")
                if 'unicode' in wx.PlatformInfo:
                    msg = unicode(msg, 'utf8')
                dlg = Dialogs.ErrorDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()
                # Set the form focus to the Name field
                self.txtColorName.SetFocus()
                # Signal validation failure
                return (False, 0, 0, 0)
            # If we are supposed to test for duplicate color names ...
            if testName:
                # Initialize the duplicate name as NOT found
                nameFound = False
                # Iterate through the dialog's color list
                for item in range(self.colorList.GetItemCount()):
                    # If the entered name matches the list item's color Name...
                    if colorName.upper() == self.colorList.GetItem(item, 0).GetText().upper():
                        # ... signal that a duplicate was found ...
                        nameFound = True
                        # ... and stop looking
                        break
                # If a duplicate name was found ...
                if nameFound:
                    # ... build and display the error message
                    msg = _('A color named "%s" already exists.  Please change the color name.')
                    if 'unicode' in wx.PlatformInfo:
                        msg = unicode(msg, 'utf8')
                    dlg = Dialogs.ErrorDialog(self, msg % colorName)
                    dlg.ShowModal()
                    dlg.Destroy()
                    # Set the form focus to the Name field
                    self.txtColorName.SetFocus()
                    # Signal validation failure
                    return (False, 0, 0, 0)
            # We need either a Hex definition or (Red, Green, and Blue)
            if ((self.txtColorHex.GetValue() == '') and \
                ((self.txtColorRed.GetValue() == '') or \
                 (self.txtColorGreen.GetValue() == '') or \
                 (self.txtColorBlue.GetValue() == ''))):
                # If we don't have all the data we need to determine a color, raise a ValueError exception
                raise ValueError
            # Now let's make sure we have legal values.  Give precedence to Color Values
            if ((self.txtColorRed.GetValue() != '') and \
                (self.txtColorGreen.GetValue() != '') and \
                (self.txtColorBlue.GetValue() != '')):
                # Get the Red value.  If it's not a number, a ValueError exception will be triggered here.
                colorRed = int(self.txtColorRed.GetValue())
                # If Red is not in the range of 0 to 255, raise a ValueError
                if colorRed < 0 or colorRed > 255:
                    raise ValueError
                # Get the Green value.  If it's not a number, a ValueError exception will be triggered here.
                colorGreen = int(self.txtColorGreen.GetValue())
                # If Green is not in the range of 0 to 255, raise a ValueError
                if colorGreen < 0 or colorGreen > 255:
                    raise ValueError
                # Get the Blue value.  If it's not a number, a ValueError exception will be triggered here.
                colorBlue = int(self.txtColorBlue.GetValue())
                # If Blue is not in the range of 0 to 255, raise a ValueError
                if colorBlue < 0 or colorBlue > 255:
                    raise ValueError
                # Set the Hex value to match the Red, Gree, Blue values
                self.txtColorHex.SetValue("#%02X%02X%02X" % (colorRed, colorGreen, colorBlue))
                # Return Success, along with the color values
                return (True, colorRed, colorGreen, colorBlue)
            # If we're working from a Hex value ...
            else:
                # First, determine that the hex value is the proper length and format, either #RRGGBB or just RRGGBB
                if ((len(self.txtColorHex.GetValue()) == 7) and \
                    (self.txtColorHex.GetValue()[0] == '#')) or \
                   (len(self.txtColorHex.GetValue()) == 6):
                    # Convert number pairs in hex (base 16) to integers.  A ValueError exception will be raised if the
                    # conversion fails.  (It's easiest to pull digits by their distance to the END of the string!)
                    colorRed = int(self.txtColorHex.GetValue()[-6:-4], 16)
                    colorGreen = int(self.txtColorHex.GetValue()[-4:-2], 16)
                    colorBlue = int(self.txtColorHex.GetValue()[-2:], 16)
                # An improperly formatted data string will raise a ValueError exception
                else:
                    raise ValueError
                # Fill in the Red, Green, and Blue values on screen based on the Hex data
                self.txtColorRed.SetValue(str(colorRed))
                self.txtColorGreen.SetValue(str(colorGreen))
                self.txtColorBlue.SetValue(str(colorBlue))
                # Return Success, along with the color values
                return (True, colorRed, colorGreen, colorBlue)
        # If any exception is raised ... (probably a ValueError!)
        except:
            # ... build and display the error message.  This could be more specific, but does it really need to be??
            msg = _("Please specify valid color settings.")
            if 'unicode' in wx.PlatformInfo:
                msg = unicode(msg, 'utf8')
            dlg = Dialogs.ErrorDialog(self, msg)
            dlg.ShowModal()
            dlg.Destroy()
            # Set the form focus to the Hex field
            self.txtColorHex.SetFocus()
            # Signal validation failure
            return (False, 0, 0, 0)

    def OnAdd(self, event):
        """ Add Button Press -- Add a Color """
        # Validate the Color Definition, including making sure we don't have duplicate color names.
        (success, colorRed, colorGreen, colorBlue) = self.ValidateColorDef(True)
        # If not valid, we're done here.
        if not success:
            return
        
        # Create a bitmap of the correct color
        bmp = self.CreateBmp(colorRed, colorGreen, colorBlue)
        # Add the bitmap to the imagelist, capturing its index in the image index
        self.imageIndex.insert(self.currentItem, self.imageList.Add(bmp))

        # Create a new item in the desired new position in the list and populate it with values from the form
        index = self.colorList.InsertImageStringItem(self.currentItem, self.txtColorName.GetValue(), self.imageIndex[self.currentItem])
        self.colorList.SetStringItem(index, 1, self.txtColorHex.GetValue())
        self.colorList.SetStringItem(index, 2, self.txtColorRed.GetValue())
        self.colorList.SetStringItem(index, 3, self.txtColorGreen.GetValue())
        self.colorList.SetStringItem(index, 4, self.txtColorBlue.GetValue())

        # Now to select the new list position
        self.colorList.SetItemState(self.currentItem, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)

    def OnEdit(self, event):
        """ Edit Button Press -- Edit a Color """
        # Validate the Color Definition, including making sure we don't have duplicate color names.
        (success, colorRed, colorGreen, colorBlue) = self.ValidateColorDef()
        # If not valid, we're done here.
        if not success:
            return

        # NOTE:  We cannot edit the last item in the list, White.  Transana requires this to be the last value.
        if self.currentItem < self.colorList.GetItemCount() - 1:
            # NOTE:  I tried just replacing the value in the list, but could not figure out how to update the
            #        color bitmap in Column 0.  It's easiest to delete the entry and insert a new one.
            # Delete the current selection.
            self.colorList.DeleteItem(self.currentItem)
            # Remove the image index as well!
            self.imageIndex = self.imageIndex[:self.currentItem] + self.imageIndex[self.currentItem + 1:]

            # Create a bitmap of the correct color
            bmp = self.CreateBmp(colorRed, colorGreen, colorBlue)
            # Add the bitmap to the imagelist, capturing its index in the image index
            # (insert is OK because we removed it above!)
            self.imageIndex.insert(self.currentItem, self.imageList.Add(bmp))

            # Create a new item in the desired new position in the list and populate it with values stored above
            index = self.colorList.InsertImageStringItem(self.currentItem, self.txtColorName.GetValue(), self.imageIndex[self.currentItem])
            self.colorList.SetStringItem(index, 1, self.txtColorHex.GetValue())
            self.colorList.SetStringItem(index, 2, self.txtColorRed.GetValue())
            self.colorList.SetStringItem(index, 3, self.txtColorGreen.GetValue())
            self.colorList.SetStringItem(index, 4, self.txtColorBlue.GetValue())

            # Now to select the moved position
            self.colorList.SetItemState(self.currentItem, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)

    def OnDelete(self, event):
        """ Delete Button Press -- Delete a Color """
        # NOTE:  We cannot delete the last entry in the list.  Transana requires White at the end of the color list.
        if self.currentItem < self.colorList.GetItemCount() - 1:
            # Delete the current selection.
            self.colorList.DeleteItem(self.currentItem)
            # Remove the image index as well!
            self.imageIndex = self.imageIndex[:self.currentItem] + self.imageIndex[self.currentItem + 1:]
            # Now to select the item in that same position
            self.colorList.SetItemState(self.currentItem, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)

    def OnOK(self, event):
        """ OK Button Press """
        # Save the current values if OK is pressed
        self.OnFileSave(event)
        # If there was no SAVE because there is no file name, we still want to update the colors.  (Normally,
        # this happens within FileSave.)
        if self.configFile == '':
            # Initialize the local internal color list
            self.colors = []
            # Iterate through the color defs in the dialog's list item
            for item in range(self.colorList.GetItemCount()):
                # Get the color name from column 0
                colorName = self.colorList.GetItem(item, 0).GetText()
                # Add the color definition to the local internal color list
                self.colors.append((colorName, (int(self.colorList.GetItem(item, 2).GetText()), int(self.colorList.GetItem(item, 3).GetText()), int(self.colorList.GetItem(item, 4).GetText()))))
        # Update the colors for all of Transana
        TransanaGlobal.transana_graphicsColorList = self.colors
        # Update all of the color-related data values needed to make use of the color scheme
        (TransanaGlobal.transana_colorNameList, TransanaGlobal.transana_colorLookup, TransanaGlobal.keywordMapColourSet) = \
            TransanaGlobal.SetColorVariables()
        # Save the config filename in the configuration structure 
        TransanaGlobal.configData.colorConfigFilename = self.configFile
        # indicate that we are closing the form
        self.closing = True
        # Allow the form's OK event to fire to close the form
        event.Skip()
        
    def OnCancel(self, event):
        """ Cancel Button Press """
        # If you hit Cancel on the Mac, you get a Segment Fault!!  This is an attempt to fix that.
        # indicate that we are closing the form
        self.closing = True
        # Allow the form's Cancel event to fire to close the form
        event.Skip()

    def OnHelp(self, event):
        """ Help Button Press """
        # If menuWindow exists (and it better!) ...
        if TransanaGlobal.menuWindow != None:
            # ... call the Transana Help system
            TransanaGlobal.menuWindow.ControlObject.Help('Graphics Color Configuration')
        
    def OnSize(self, event):
        # Allow the form's base Size event to perform its duties, such as calling Layout() to update the sizers.
        event.Skip()

    def CreateBmp(self, colorRed, colorGreen, colorBlue):
        """ Create a small bitmap of the color specified """
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
        brush = wx.Brush(wx.Colour(colorRed, colorGreen, colorBlue))
        # Set the Brush for the Device Context
        dc.SetBrush(brush)
        # Draw a black border around the color graphic, leaving a little white space
        dc.DrawRectangle(0, 2, 14, 14)
        # End the drawing process
        dc.EndDrawing()
        # Select a different object into the Device Context, which allows the bmp to be used.
        dc.SelectObject(wx.EmptyBitmap(5,5))
        # Return the bitmap
        return bmp
        

# For testing purposes, this module can run stand-alone.
if __name__ == '__main__':
    
    # Create a simple app for testing.
    app = wx.PySimpleApp()
    
    print wx.PlatformInfo

    frame = ColorConfig(None)

    # Show the Dialog Box and process the result.
    # Note that both the wx.FontData object and the TransanaFontDef object
    # can return results regardless of which was used to start the process.
    if frame.ShowModal() == wx.ID_OK:
        print "OK pressed."
    else:
        print "Cancel pressed."

    # Destroy the dialog box.
    frame.Destroy()
    # Call the app's MainLoop()
    app.MainLoop()
