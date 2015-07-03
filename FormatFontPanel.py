# Copyright (C) 2003 - 2015 The Board of Regents of the University of Wisconsin System 
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

""" This implements the Transana Font Formatting Panel class. The wxPython wxFontDialog 
    proved inadequate for my needs, especially on the Mac.  It is modeled after wxFontDialog.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Enable (True) or Disable (False) debugging messages
DEBUG = False
if DEBUG:
    print "FormatFontPanel DEBUG is ON"

# For testing purposes, this module can run stand-alone.
if __name__ == '__main__':
    import wxversion
    wxversion.select(['2.6-unicode'])

# import wxPython
import wx

# Import the FormatDialog object (for constant definitions)
import FormatDialog

if __name__ == '__main__':
    # This module expects i18n.  Enable it here.
    __builtins__._ = wx.GetTranslation

# import the TransanaGlobal variables
import TransanaGlobal


class FormatFontPanel(wx.Panel):
    """ A custom Font Formatting Panel.  Pass in a wxFontData object (to maintain compatability with the wxFontDialog) or
        a FormatDialog.FormatDef object designed to allow for ambiguity in the font specification.  """
    
    def __init__(self, parent, fontData, sampleText='AaBbCc ... XxYyZz'):
        """ Initialize the Font Dialog Box.  fontData should be a FormatDialog.FormatDef object so that some values can be
            ambiguous due to conflicting settings in the selected text.  """
        
        # Capture the Font Data
        self.font = fontData

        # Remember the original FormatDialog.FormatDef settings in case the user presses Cancel
        self.originalFont = self.font.copy()

        # Define our Sample Text's Text.
        self.sampleText = sampleText
        # Create the Font Dialog
        wx.Panel.__init__(self, parent, -1)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the main Sizer, which will hold the boxTop, boxMiddle, and boxButton sizers
        box = wx.BoxSizer(wx.VERTICAL)
        # Create the boxTop sizer, which will hold the boxFont and boxSize sizers
        boxTop = wx.BoxSizer(wx.HORIZONTAL)
        # Create the boxFont sizer, which will hold the Font Face widgets
        boxFont = wx.BoxSizer(wx.VERTICAL)

        # Add Font Face widgets.
        # Create the label
        lblFont = wx.StaticText(self, -1, _('Font:'))
        boxFont.Add(lblFont, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP)
        boxFont.Add((0, 5))  # Spacer

        # Create a text control for the font face name.
        # First, determine the initial value.
        if self.font.fontFace == None:
            fontFace = ''
        else:
            fontFace = self.font.fontFace.strip()
        self.txtFont = wx.TextCtrl(self, -1, fontFace, style=wx.TE_LEFT)
        self.txtFont.Bind(wx.EVT_TEXT, self.OnTxtFontChange)
        self.txtFont.Bind(wx.EVT_KILL_FOCUS, self.OnTxtFontKillFocus)
        boxFont.Add(self.txtFont, 0, wx.ALIGN_LEFT | wx.EXPAND) 
        
        # Create a list box control of all available font face names.  The user can type into the text control, but the
        # font should only change when what is typed matches an entry in the list box.
        # First, let's get a list of all available fonts using a wxFontEnumerator.
        fontEnum = wx.FontEnumerator()
        fontEnum.EnumerateFacenames()
        fontList = fontEnum.GetFacenames()

        # Sort the Font List
        fontList.sort()

        # Now create the Font Face Names list box.  We can't use the LB_SORT parameter on OS X with wxPython 2.9.5.0.b
        self.lbFont = wx.ListBox(self, -1, choices=fontList, style=wx.LB_SINGLE | wx.LB_ALWAYS_SB ) # | wx.LB_SORT)
        # Make sure the initial font is in the list ...
        if self.font.fontFace != None:
            # If the font name IS found in the dropdown ...
            if self.font.fontFace in fontList:  # self.lbFont.FindString(self.font.fontFace) != wx.NOT_FOUND:
                # ... then select that font name.
                self.lbFont.SetStringSelection(self.font.fontFace.strip())
            # If not ...
            else:
                # Try to use the Default Font instead.  If the default font IS found in the dropdown ...
                if TransanaGlobal.configData.defaultFontFace in fontList:  # self.lbFont.FindString(TransanaGlobal.configData.defaultFontFace) != wx.NOT_FOUND:
                    # ... then select that font name in the dropdown and update the text control.
                    self.txtFont.SetValue(TransanaGlobal.configData.defaultFontFace.strip())
                    self.lbFont.SetStringSelection(TransanaGlobal.configData.defaultFontFace.strip())
                # If neither the current font nor the default font are in the list ...
                else:
                    # ... select the first font in the list in both the dropdown and in the text control.
                    self.lbFont.SetSelection(0)
                    self.txtFont.SetValue(self.lbFont.GetStringSelection())

        # Bind the List Box Change event
        self.lbFont.Bind(wx.EVT_LISTBOX, self.OnLbFontChange)
        boxFont.Add(self.lbFont, 3, wx.ALIGN_LEFT | wx.ALIGN_BOTTOM | wx.EXPAND | wx.GROW)

        # Add the boxFont sizer to the boxTop sizer
        boxTop.Add(boxFont, 5, wx.ALIGN_LEFT | wx.EXPAND)
        # Create the boxSize sizer, which will hold the Font Size widgets
        boxSize = wx.BoxSizer(wx.VERTICAL)

        # Add Font Size widgets.
        # Create the label
        lblSize = wx.StaticText(self, -1, _('Size:  (points)'))
        boxSize.Add(lblSize, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP)
        boxSize.Add((0, 5))  # Spacer

        # Create a text control for the font size
        # First, determine the initial value.
        if self.font.fontSize == None:
            fontSize = ''
        else:
            fontSize = str(self.font.fontSize)
        self.txtSize = wx.TextCtrl(self, -1, fontSize, style=wx.TE_LEFT)
        self.txtSize.Bind(wx.EVT_TEXT, self.OnTxtSizeChange)
        boxSize.Add(self.txtSize, 0, wx.ALIGN_LEFT | wx.EXPAND)

        # Create a list box control of available font sizes, though the user can type other options in the text control.
        # First, let's make a list of available font sizes.  It doesn't have to be complete.  The blank first entry
        # is used when the user types in something that isn't in the list.
        sizeList = ['', '6', '7', '8', '9', '10', '11', '12', '14', '16', '18', '20', '22', '24', '26', '28', '30', '32', '36', '40', '44', '48', '54', '64', '72']

        # Now create the Font Sizes list box        
        self.lbSize = wx.ListBox(self, -1, choices=sizeList, style=wx.LB_SINGLE | wx.LB_ALWAYS_SB)
        if (self.font.fontSize != None) and (str(self.font.fontSize) in sizeList):
            self.lbSize.SetStringSelection(str(self.font.fontSize))
        self.lbSize.Bind(wx.EVT_LISTBOX, self.OnLbSizeChange)
        boxSize.Add(self.lbSize, 3, wx.ALIGN_LEFT | wx.ALIGN_BOTTOM | wx.EXPAND | wx.GROW)
        boxTop.Add((15, 0))  # Spacer

        # Add the boxSize sizer to the boxTop sizer
        boxTop.Add(boxSize, 3, wx.ALIGN_RIGHT | wx.EXPAND | wx.GROW)
        # Add the boxTop sizer to the main box sizer
        box.Add(boxTop, 3, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.EXPAND | wx.GROW | wx.ALL, 10)

        # Create the boxMiddle sizer, which will hold the boxStyle and boxSample sizers
        boxMiddle = wx.BoxSizer(wx.HORIZONTAL)
        # Create the boxStyle sizer, which will hold the Style and Color widgets
        boxStyle = wx.BoxSizer(wx.VERTICAL)

        # Add the Font Style and Font Color widgets
        # Start with the Style label
        lblStyle = wx.StaticText(self, -1, _('Style:'))
        boxStyle.Add(lblStyle, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP)
        boxStyle.Add((0, 5))  # Spacer

        # Add a checkbox for Bold
        self.checkBold = wx.CheckBox(self, -1, _('Bold'), style=wx.CHK_3STATE)
        # Determine and set the initial value.
        if self.font.fontWeight == FormatDialog.fd_OFF:
            checkValue = wx.CHK_UNCHECKED
        elif self.font.fontWeight == FormatDialog.fd_BOLD:
            checkValue = wx.CHK_CHECKED
        elif self.font.fontWeight == FormatDialog.fd_AMBIGUOUS:
            checkValue = wx.CHK_UNDETERMINED
        self.checkBold.Set3StateValue(checkValue)
        self.checkBold.Bind(wx.EVT_CHECKBOX, self.OnBold)

        boxStyle.Add(self.checkBold, 0, wx.ALIGN_LEFT | wx.LEFT, 5)
        boxStyle.Add((0, 5))  # Spacer

        # Add a checkbox for Italics
        self.checkItalics = wx.CheckBox(self, -1, _('Italics'), style=wx.CHK_3STATE)
        # Determine and set the initial value.
        if self.font.fontStyle == FormatDialog.fd_OFF:
            checkValue = wx.CHK_UNCHECKED
        elif self.font.fontStyle == FormatDialog.fd_ITALIC:
            checkValue = wx.CHK_CHECKED
        elif self.font.fontStyle == FormatDialog.fd_AMBIGUOUS:
            checkValue = wx.CHK_UNDETERMINED
        self.checkItalics.Set3StateValue(checkValue)
        self.checkItalics.Bind(wx.EVT_CHECKBOX, self.OnItalics)
            
        boxStyle.Add(self.checkItalics, 0, wx.ALIGN_LEFT | wx.LEFT, 5)
        boxStyle.Add((0, 5))  # Spacer

        # Add a checkbox for Underline
        self.checkUnderline = wx.CheckBox(self, -1, _('Underline'), style=wx.CHK_3STATE)
        # Determine and set the initial value.
        if self.font.fontUnderline == FormatDialog.fd_OFF:
            checkValue = wx.CHK_UNCHECKED
        elif self.font.fontUnderline == FormatDialog.fd_UNDERLINE:
            checkValue = wx.CHK_CHECKED
        elif self.font.fontUnderline == FormatDialog.fd_AMBIGUOUS:
            checkValue = wx.CHK_UNDETERMINED
        self.checkUnderline.Set3StateValue(checkValue)
        self.checkUnderline.Bind(wx.EVT_CHECKBOX, self.OnUnderline)

        boxStyle.Add(self.checkUnderline, 0, wx.ALIGN_LEFT | wx.LEFT, 5)
        boxStyle.Add((0, 10))  # Spacer

        # Add a label for Color
        lblColor = wx.StaticText(self, -1, _('Color:'))
        boxStyle.Add(lblColor, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP)
        boxStyle.Add((0, 5))  # Spacer

        # We want enough colors, but not too many.  This list seems about right to me.  I doubt my color names are standard.
        # But then, I'm often perplexed by the colors that are included and excluded by most programs.  (Excel for example.)
        # Each entry is made up of a color name and a tuple of the RGB values for the color.
        self.colorList = TransanaGlobal.transana_textColorList

        # We need to create a list of the colors to be included in the control.
        choiceList = []
        # Default to undefined to allow for ambiguous colors, if the original color isn't included in the list.
        # NOTE:  This dialog will only support the colors in this list at this point.
        initialColor = ''  # _('Black')
        initialBgColor = ''  # _('White')
        # Iterate through the list of colors ...
        for (color, colDef) in self.colorList:
            # ... adding each color name to the list of what should be displayed ...
            choiceList.append(_(color))
            # ... and checking to see if the color in the list matches the initial color sent to the dialog.
            if colDef == self.font.fontColorDef:
                # If the current color matches a color in the list, remember it's name.
                initialColor = _(color)
            if colDef == self.font.fontBackgroundColorDef:
                initialBgColor = _(color)

        # Now create a Choice box listing all the colors in the color list
        self.cbColor = wx.Choice(self, -1, choices=[''] + choiceList)
        # Set the initial value of the Choice box to the default value determined above.
        self.cbColor.SetStringSelection(initialColor)

        self.cbColor.Bind(wx.EVT_CHOICE, self.OnCbColorChange)
        boxStyle.Add(self.cbColor, 1, wx.ALIGN_LEFT | wx.ALIGN_BOTTOM)

        boxStyle.Add((0, 10))  # Spacer

        # Add a label for Color
        lblBgColor = wx.StaticText(self, -1, _('Background Color:'))
        boxStyle.Add(lblBgColor, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP)
        boxStyle.Add((0, 5))  # Spacer

        # Now create a Choice box listing all the colors in the color list
        self.cbBgColor = wx.Choice(self, -1, choices=choiceList + [''])
        # Set the initial value of the Choice box to the default value determined above.
        self.cbBgColor.SetStringSelection(initialBgColor)
        self.cbBgColor.Bind(wx.EVT_CHOICE, self.OnCbColorChange)
        boxStyle.Add(self.cbBgColor, 1, wx.ALIGN_LEFT | wx.ALIGN_BOTTOM)

        # Add the boxStyle sizer to the boxMiddle sizer
        boxMiddle.Add(boxStyle, 1, wx.ALIGN_LEFT | wx.RIGHT, 10)

        # Create the boxSample sizer, which will hold the Text Sample widgets
        boxSample = wx.BoxSizer(wx.VERTICAL)

        # Create the Text Sample widgets.
        # Start with a label
        lblSample = wx.StaticText(self, -1, _('Sample:'))
        boxSample.Add(lblSample, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP)
        boxSample.Add((0, 5))  # Spacer

        # We'll use a StaticBitmap for the sample text, painting directly on its Device Context.  The TextCtrl 
        # on the Mac can't handle all we need it to for this task.
        # We create this here, as it may get called while creating other controls.  We'll add it to the sizers later.
        self.txtSample = wx.StaticBitmap(self, -1)

        # We added the txtSample control earlier.  Set its background color and add it to the Sizers here.
        self.txtSample.SetBackgroundColour(self.font.fontBackgroundColorDef)
        boxSample.Add(self.txtSample, 1, wx.ALIGN_RIGHT | wx.EXPAND | wx.GROW) 

        # Add the boxSample sizer to the boxMiddle sizer
        boxMiddle.Add(boxSample, 3, wx.ALIGN_RIGHT | wx.EXPAND | wx.GROW)
        # Add the boxMiddle sizer to the main box sizer
        box.Add(boxMiddle, 2, wx.ALIGN_LEFT | wx.EXPAND | wx.GROW | wx.ALL, 10)

        # Define box as the form's main sizer
        self.SetSizer(box)
        # Fit the form to the widgets created
        self.Fit()
        # Set this as the minimum size for the form.
        self.SetSizeHints(minW = self.GetSize()[0], minH = self.GetSize()[1])
        # Tell the form to maintain the layout and have it set the intitial Layout
        self.SetAutoLayout(True)
        self.Layout()

        # We need an Size event for the form for a little mainenance when the form size is changed
        self.Bind(wx.EVT_SIZE, self.OnSize)
        
        # Under wxPython 2.6.1.0-unicode, this form is throwing a segment fault when the color gets changed.
        # The following variable prevents that!
        self.closing = False


    def SetSampleFont(self):
        """ Update the Sample Text to reflect the dialog's current selections """
        # The Font Sample has been having some trouble on OS X for some users, but I can't recreate it.
        # We'll just put this whole method in a try ... except block to prevent it from causing problems.
        try:
            # To get a truly accurate font sample, we need to create a graphic object, paint the
            # font sample on it, and place that on the screen.
            # Get the size of the sample area
            (bmpWidth, bmpHeight) = self.txtSample.GetSize()
            # Create a BitMap that is the right size
            bmp = wx.EmptyBitmap(bmpWidth, bmpHeight)
            # Create a Device Context for the BitMap as a BufferedDC
            dc = wx.BufferedDC(None, bmp)
            # if the background color is not ambiguous ...
            if self.font.fontBackgroundColorDef != None:
                # Set the background Color
                dc.SetBackground(wx.Brush(self.font.fontBackgroundColorDef))
            # If the background color is ambiguous ...
            else:
                # ... just use White
                dc.SetBackground(wx.Brush(wx.NamedColour("White")))
            # Clear the Device Context
            dc.Clear()
            # Begin drawing on the Device Context
            dc.BeginDrawing()
            # Draw the boundary lines.  (DrawRect wipes out the colored background!)
            dc.DrawLines([wx.Point(0, 0),
                          wx.Point(bmpWidth-1, 0),
                          wx.Point(bmpWidth-1, bmpHeight-1),
                          wx.Point(0, bmpHeight-1),
                          wx.Point(0, 0)])
            if self.font.fontSize == None:
                tmpFontSize = TransanaGlobal.configData.defaultFontSize
            else:
                tmpFontSize = self.font.fontSize
            if self.font.fontFace == None:
                tmpFontFace = TransanaGlobal.configData.defaultFontFace.strip()
            else:
                tmpFontFace = self.font.fontFace.strip()
            if self.font.fontStyle == FormatDialog.fd_ITALIC:
                tmpStyle = wx.FONTSTYLE_ITALIC
            else:
                tmpStyle = wx.FONTSTYLE_NORMAL
            if self.font.fontWeight == FormatDialog.fd_BOLD:
                tmpWeight = wx.FONTWEIGHT_BOLD
            else:
                tmpWeight = wx.FONTWEIGHT_NORMAL
            if self.font.fontUnderline != FormatDialog.fd_AMBIGUOUS:
                tmpUnderline = self.font.fontUnderline
            else:
                tmpUnderline = False
            # Create the specified font
            tmpFont = wx.Font(tmpFontSize, wx.FONTFAMILY_DEFAULT, tmpStyle, tmpWeight, tmpUnderline, tmpFontFace)

            # If the font has been successfully created ...
            if tmpFont.IsOk():
                # Set the Device Context's Font
                dc.SetFont(tmpFont)
                # Set the Device Context's Foreground Color
                dc.SetTextForeground(self.font.fontColorDef)
                # Determine how bit our Sample Text is going to be
                (w, h) = dc.GetTextExtent(self.sampleText)
                # Center the Sample Text horizontally and vertically
                x = int(bmpWidth / 2) - int(w / 2)
                y = int(bmpHeight / 2) - int(h / 2)
                # Draw the Sample Text on the Device Context
                dc.DrawText(self.sampleText, x, y)
            # Signal that we're done drawing on the Device Context
            dc.EndDrawing()
            # Put the bitmap on the txtSample StaticBitmap
            self.txtSample.SetBitmap(bmp)
            # Update the window so the bitmap will be updated
            self.Refresh()
        except:

            if DEBUG:
                import sys
                print "FormatFontPanel.SetSampleFont():", sys.exc_info()[0], sys.exc_info()[1]
                import traceback
                print traceback.print_exc()
                
            pass
        
    def OnTxtFontChange(self, event):
        """ txtFont Change Event.  As the user types in the Font Name, the font ListBox should try to match it. """
        # NOTE:  This method can be improved by having it perform matching with incomplete strings.
        
        # If the typed string matches one of the items in the Font Name List Box ...
        if self.lbFont.FindString(self.txtFont.GetValue()) != wx.NOT_FOUND:
            # ... then update the ListBox to match the typing.  This triggers the font face change.
            self.lbFont.SetStringSelection(self.txtFont.GetValue())
        # If the typed string doesn't match a full font ...
        else:
            # ... get a list of all fonts
            fonts = self.lbFont.GetStrings()
            # Iterate through this list 
            for font in fonts:
                # If the text in the TextCtrl matches the item in the list ...
                if self.txtFont.GetValue().upper() == font.upper()[:len(self.txtFont.GetValue())]:
                    # ... set the list box to this item
                    self.lbFont.SetStringSelection(font)
                    # ... and stop iterating
                    break
        # Update the Font Face Name
        self.font.fontFace = self.lbFont.GetStringSelection().strip()
        # Update the Font Sample 
        self.SetSampleFont()

    def OnTxtFontKillFocus(self, event):
        """ txtFont Kill Focus Event. """
        # If the user leaves the txtFont widget, we need to make sure it has a valid value.
        # Check to see if the current value is a legal Font Face name by comparing it to the Font Face
        # List Box values.  If it's not a legal value ...
        if (not self.closing) and (self.txtFont.GetValue() != self.lbFont.GetStringSelection()):
            # ... revert the text to the last selected font face name.
            self.txtFont.SetValue(self.lbFont.GetStringSelection().strip())
        
    def OnLbFontChange(self, event):
        """ Change Font based on selection in the Font List Box """
        # Update the Font Face text control to match the list box selection
        self.txtFont.SetValue(self.lbFont.GetStringSelection())
        # Update the Current Font Face name to match the Font Name selection
        self.font.fontFace = self.lbFont.GetStringSelection().strip()
        # Update the Font Sample 
        self.SetSampleFont()
        if DEBUG:
            print "Font face set to", self.lbFont.GetStringSelection()
        
    def OnTxtSizeChange(self, event):
        """ txtSize Change Event.  As the user type in a Font Size, adjust the Sample if the value is legal. """
        # NOTE:  This could probably be improved by using a wx.MaskedEditCtrl instead of a wx.TextCtrl
        
        # If the size value is blank, don't bother with this method.
        if self.txtSize.GetValue == '':
            return
        # Let's make sure we've got only numbers in the field
        try:
            # Attempt to convert the text to an integer.  NOTE:  half-point values are valid!!
            size = int(self.txtSize.GetValue())
        except:
            # If we get an exception here, we don't have a valid integer in the text control.
            # If we have a valid selection in the Font Size ListBox, ...
            if self.lbSize.GetStringSelection() != '':
                # ... update the text control to match that value.
                self.txtSize.SetValue(self.lbSize.GetStringSelection())
            # If we don't have a valid font size in the Font Size List Box ...
            else:
                # Update the text to match the last legal value.
                self.txtSize.SetValue(str(self.font.fontSize))
        # Point Sizes need to be between 1 and 255.
        if (size > 0) and (size < 256):
            # If we have a valid value, adjust the font to match the text entry.
#            self.currentFont.SetPointSize(size)
            self.font.fontSize = size
            # If the text entry matches a value in the list box ...
            if self.lbSize.FindString(self.txtSize.GetValue()) != wx.NOT_FOUND:
                # ... select that value ...
                self.lbSize.SetStringSelection(self.txtSize.GetValue())
            else:
                # ... otherwise, select the first value, which is the blank.
                self.lbSize.SetSelection(0)
            # Set the font's point size to match the text entry
            # self.currentFont.SetPointSize(size)
            # Update the Font Sample 
            self.SetSampleFont()
            if DEBUG:
                print "font size set to", size
        # If the point size is not valid ...
        else:
            # If we have a valid selection in the Font Size ListBox, ...
            if self.lbSize.GetStringSelection() != '':
                # ... update the text control to match that value.
                self.txtSize.SetValue(self.lbSize.GetStringSelection())
            # If we don't have a valid font size in the Font Size List Box ...
            else:
                # Update the text to match the last legal value.
                self.txtSize.SetValue(str(self.font.fontSize))
        
    def OnLbSizeChange(self, event):
        """ Change Font Size based on selection in the Font Size Box """
        # If the selection is not the first, empty item ...
        if self.lbSize.GetStringSelection() != '':
            # Update the Font Size text control to match the list box selection
            self.txtSize.SetValue(self.lbSize.GetStringSelection())
            self.font.fontSize = int(self.lbSize.GetStringSelection())
            # Update the Font Sample 
            self.SetSampleFont()
            if DEBUG:
                print "Font size set to", self.lbSize.GetStringSelection()
        
    def OnBold(self, event):
        """ Change Bold based on selection in Bold checkbox """
        if self.checkBold.Is3State():
            # If the box is checked ...
            if self.checkBold.Get3StateValue() == wx.CHK_CHECKED:
                # ... the Font Weight should be wx.BOLD ...
                style = wx.BOLD
                fontStyle = FormatDialog.fd_BOLD
                if DEBUG:
                    print "Bold ON"
            elif self.checkBold.Get3StateValue() == wx.CHK_UNCHECKED:
                # ... otherwise it should be wx.NORMAL
                style = wx.NORMAL
                fontStyle = FormatDialog.fd_OFF
                if DEBUG:
                    print "Bold OFF"
            else:
                fontStyle = FormatDialog.fd_AMBIGUOUS
                if DEBUG:
                    print "Bold is AMBIGUOUS"
            self.font.fontWeight = fontStyle
        else:
            # If the box is checked ...
            if self.checkBold.IsChecked():
                # ... the Font Weight should be wx.BOLD ...
                style = wx.BOLD
                fontStyle = FormatDialog.fd_BOLD
                if DEBUG:
                    print "Bold ON"
            else:
                # ... otherwise it should be wx.NORMAL
                style = wx.NORMAL
                fontStyle = FormatDialog.fd_OFF
                if DEBUG:
                    print "Bold OFF"
            self.font.fontWeight = fontStyle
        # Update the Font Sample 
        self.SetSampleFont()
        
    def OnItalics(self, event):
        """ Change Italics based on selection in Italics checkbox """
        if self.checkItalics.Is3State():
            # If the box is checked ...
            if self.checkItalics.Get3StateValue() == wx.CHK_CHECKED:
                # ... the Font Style should be wx.ITALIC ...
                style = wx.ITALIC
                fontStyle = FormatDialog.fd_ITALIC
                if DEBUG:
                    print "Italics ON"
            elif self.checkItalics.Get3StateValue() == wx.CHK_UNCHECKED:
                # ... otherwise it should be wx.NORMAL
                style = wx.NORMAL
                fontStyle = FormatDialog.fd_OFF
                if DEBUG:
                    print "Italics OFF"
            else:
                fontStyle = FormatDialog.fd_AMBIGUOUS
                if DEBUG:
                    print "Style is AMBIGUOUS"
            self.font.fontStyle = fontStyle
        else:
            # If the box is checked ...
            if self.checkItalics.IsChecked():
                # ... the font style should be wx.ITALIC
                style = wx.ITALIC
                fontStyle = FormatDialog.fd_ITALIC
                if DEBUG:
                    print "Italics ON"
            else:
                # ... otherwise it should be wx.NORMAL
                style = wx.NORMAL
                fontStyle = FormatDialog.fd_OFF
                if DEBUG:
                    print "Italics OFF"
            self.font.fontStyle = fontStyle
        # Update the Font Sample 
        self.SetSampleFont()
        
    def OnUnderline(self, event):
        """ Change Underline based on selection in Underline checkbox """
        if self.checkUnderline.Is3State():
            # If the box is checked ...
            if self.checkUnderline.Get3StateValue() == wx.CHK_CHECKED:
                fontStyle = FormatDialog.fd_UNDERLINE
            else:
                fontStyle = FormatDialog.fd_OFF
            self.font.fontUnderline = fontStyle
        else:
            if self.checkUnderline.IsChecked():
                fontStyle = FormatDialog.fd_UNDERLINE
            else:
                fontStyle = FormatDialog.fd_OFF
            self.font.fontUnderline = fontStyle
        # Update the Font Sample 
        self.SetSampleFont()

    def OnCbColorChange(self, event):
        """ cbColor Change Event.  Change the font color. """
        if event.GetId() == self.cbColor.GetId():
            ctrl = self.cbColor
        elif event.GetId() == self.cbBgColor.GetId():
            ctrl = self.cbBgColor
        else:
            print "TransanaFontDialog.OnCbColorChange() FAILURE"
            return
        # Iterate through the color list ...
        for (color, colDef) in self.colorList:
            # ... and find the color that matches the choice box selection.
            if 'unicode' in wx.PlatformInfo:
                color = unicode(_(color), 'utf8')
            else:
                color = _(color)
            if color == ctrl.GetStringSelection():
                if DEBUG:
                    print "Color set to:", color, colDef
                if event.GetId() == self.cbColor.GetId():
                    # When you have a match, use that color's definition as the current color.
                    self.currentColor = wx.Colour(colDef[0], colDef[1], colDef[2])
                    self.font.fontColorDef = wx.Colour(colDef[0], colDef[1], colDef[2])
                elif event.GetId() == self.cbBgColor.GetId():
                    self.font.fontBackgroundColorDef = wx.Colour(colDef[0], colDef[1], colDef[2])
                # Update the Font Sample 
                self.SetSampleFont()

    def OnSize(self, event):
        # Allow the form's base Size event to perform its duties, such as calling Layout() to update the sizers.
        event.Skip()
        # When making the dialog larger, the check box widgets get spoiled.  Let's fix that.
        self.checkBold.Refresh()
        self.checkItalics.Refresh()
        self.checkUnderline.Refresh()
        # Update the Font Sample 
        self.SetSampleFont()
