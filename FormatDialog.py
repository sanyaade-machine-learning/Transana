# Copyright (C) 2003 - 2012 The Board of Regents of the University of Wisconsin System 
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

"""This file implements the FormatDialog class, which defines the
   Font and Paragraph formatting Dialog Box. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx

# import the Font Formatting Panel
import FormatFontPanel
# import the Paragraph Formatting Panel
import FormatParagraphPanel
# import the Tabs Panel
import FormatTabsPanel
# Import Transana's Globals
import TransanaGlobal

# Define Formatting Definition Constants
fd_OFF = 0           # The "unselected" state for font Weight, Style, and Underline
fd_BOLD = 1          # The "selected" state for Weight
fd_ITALIC = 1        # The "selected" state for Style
fd_UNDERLINE = 1     # The "selected" state for Underline
fd_AMBIGUOUS = 2     # The "ambiguous" (mixed unselected and selected) state for font Weight, Style, and Underline

class FormatDef(object):
    """ The Transana Font Dialog uses its own custom font definition object so that it is able
        to handle ambiguous font information.  For example, if some of the text to be set by
        this font specification is Bold and some is not, the Bold setting should be unset.
        The traditional wx.formatData object can't do this.  """

    def __init__(self):
        """ Initialize the FormatDef object """
        self._fontFace = None
        self._fontSize = None
        self._fontWeight = fd_OFF       # Options are fd_OFF, fd_BOLD,      fd_AMBIGUOUS  
        self._fontStyle = fd_OFF        # Options are fd_OFF, fd_ITALIC,    fd_AMBIGUOUS
        self._fontUnderline = fd_OFF    # Options are fd_OFF, fd_UNDERLINE, fd_AMBIGUOUS
        self._fontColorName = None
        self._fontColorDef = None
        self._fontBackgroundColorName = None
        self._fontBackgroundColorDef = None
        self._paragraphAlignment = None
        self._paragraphLeftIndent = None
        self._paragraphLeftSubIndent = None
        self._paragraphRightIndent = None
        self._paragraphLineSpacing = None
        self._paragraphSpaceBefore = None
        self._paragraphSpaceAfter = None
        self._tabs = []

    def __repr__(self):
        """ String Representation of the contents of the FormatDef object """
        st = 'FormatDialog.FormatDef object:\n'
        st += 'fontFace: %s\n' % self.fontFace
        st += 'fontSize: %s\n' % self.fontSize
        st += 'fontWeight: '
        if self.fontWeight == fd_OFF:
            st += 'fd_OFF'
        elif self.fontWeight == fd_BOLD:
            st += 'fd_BOLD'
        elif self.fontWeight == fd_AMBIGUOUS:
            st += 'fd_AMBIGUOUS'
        else:
            st += 'ILLEGAL SETTING "%s"' % self.fontWeight
        st += '\n'
        st += 'fontStyle: '
        if self.fontStyle == fd_OFF:
            st += 'fd_OFF'
        elif self.fontStyle == fd_ITALIC:
            st += 'fd_ITALIC'
        elif self.fontStyle == fd_AMBIGUOUS:
            st += 'fd_AMBIGUOUS'
        else:
            st += 'ILLEGAL SETTING "%s"' % self.fontStyle
        st += '\n'
        st += 'fontUnderline: '
        if self.fontUnderline == fd_OFF:
            st += 'fd_OFF'
        elif self.fontUnderline == fd_UNDERLINE:
            st += 'fd_UNDERLINE'
        elif self.fontUnderline == fd_AMBIGUOUS:
            st += 'fd_AMBIGUOUS'
        else:
            st += 'ILLEGAL SETTING "%s"' % self.fontUnderline
        st += '\n'
        st += 'fontColorName: %s\n' % self.fontColorName
        st += 'fontColorDef: %s\n' % (self.fontColorDef,)
        st += 'fontBackgroundColorName: %s\n' % self.fontBackgroundColorName
        st += 'fontBackgroundColorDef: %s\n\n' % (self.fontBackgroundColorDef,)
        st +=  "Alignment: %s  (Left = %s, Center = %s, Right = %s\n" % (self.paragraphAlignment, wx.TEXT_ALIGNMENT_LEFT, wx.TEXT_ALIGNMENT_CENTRE, wx.TEXT_ALIGNMENT_RIGHT)
        st +=  "Left Indent: %s\n" % self.paragraphLeftIndent
        st +=  "Left SubIndent: %s\n" % self.paragraphLeftSubIndent
        st +=  "   ... producing Left: %d, First Line: %d\n" % (self.paragraphLeftIndent + self.paragraphLeftSubIndent, 0 - self.paragraphLeftSubIndent)
        st +=  "Right Indent: %s\n" % self.paragraphRightIndent
        st +=  "Line Spacing: %s\n" % self.paragraphLineSpacing
        st +=  "Space Before: %s\n" % self.paragraphSpaceBefore
        st +=  "Space After: %s\n" % self.paragraphSpaceAfter
        st +=  "Tabs: %s\n\n" % self.tabs
        return st

    def copy(self):
        """ Create a copy of a FormatDef object """
        # Create a new FormatDef object
        fdCopy = FormatDef()
        # Copy the existing data values to the new Object
        fdCopy.fontFace = self.fontFace
        fdCopy.fontSize = self.fontSize
        fdCopy.fontWeight = self.fontWeight
        fdCopy.fontStyle = self.fontStyle
        fdCopy.fontUnderline = self.fontUnderline
        # We don't need to copy fontColorName.  Copying fontColorDef will take care of it.
        fdCopy.fontColorDef = self.fontColorDef
        fdCopy.fontBackgroundColorDef = self.fontBackgroundColorDef
        fdCopy.paragraphAlignment = self.paragraphAlignment
        fdCopy.paragraphLeftIndent = self.paragraphLeftIndent
        fdCopy.paragraphLeftSubIndent = self.paragraphLeftSubIndent
        fdCopy.paragraphRightIndent = self.paragraphRightIndent
        fdCopy.paragraphLineSpacing = self.paragraphLineSpacing
        fdCopy.paragraphSpaceBefore = self.paragraphSpaceBefore
        fdCopy.paragraphSpaceAfter = self.paragraphSpaceAfter
        fdCopy.tabs = self.tabs
        # Return the new Object
        return fdCopy

    # Property Getter and Setter functions
    def _getFontFace(self):
        return self._fontFace
    def _setFontFace(self, fontFace):
        self._fontFace = fontFace
    def _delFontFace(self):
        self._fontFace = None

    def _getFontSize(self):
        return self._fontSize
    def _setFontSize(self, fontSize):
        # If the parameter cannot be converted to an integer, don't change the value.
        try:
            if fontSize == None:
                self._fontSize = None
            else:
                self._fontSize = int(fontSize)
        except:
            pass
    def _delFontSize(self):
        self._fontSize = None
        
    def _getFontWeight(self):
        return self._fontWeight
    def _setFontWeight(self, fontWeight):
        if fontWeight in [fd_OFF, fd_BOLD, fd_AMBIGUOUS]:
            self._fontWeight = fontWeight
    def _delFontWeight(self):
        self._fontWeight = fd_OFF

    def _getFontStyle(self):
        return self._fontStyle
    def _setFontStyle(self, fontStyle):
        if fontStyle in [fd_OFF, fd_ITALIC, fd_AMBIGUOUS]:
            self._fontStyle = fontStyle
    def _delFontStyle(self):
        self._fontStyle = fd_OFF

    def _getFontUnderline(self):
        return self._fontUnderline
    def _setFontUnderline(self, fontUnderline):
        if fontUnderline in [fd_OFF, fd_UNDERLINE, fd_AMBIGUOUS]:
            self._fontUnderline = fontUnderline
    def _delFontUnderline(self):
        self._fontUnderline = fd_OFF

    def _getFontColorName(self):
        return self._fontColorName
    def _setFontColorName(self, fontColorName):
        if fontColorName in TransanaGlobal.transana_colorNameList:
            self._fontColorName = fontColorName
            # Set fontColorDef to match fontColorName
            for (colorName, colorDef) in TransanaGlobal.transana_textColorList:
                if colorName == fontColorName:
                    self._fontColorDef = wx.Colour(colorDef[0], colorDef[1], colorDef[2])
                    break
    def _delFontColorName(self):
        self._fontColorName = None
        self._fontColorDef = None

    def _getFontColorDef(self):
        return self._fontColorDef
    def _setFontColorDef(self, fontColorDef):
        self._fontColorDef = fontColorDef
        # Set fontColorName to match fontColorDef
        for (colorName, colorDef) in TransanaGlobal.transana_textColorList:
            if colorDef == fontColorDef:
                self._fontColorName = colorName
                break
    def _delFontColorDef(self):
        self._fontColorDef = None
        self._fontColorName = None

    def _getFontBackgroundColorName(self):
        return self._fontBackgroundColorName
    def _setFontBackgroundColorName(self, fontBackgroundColorName):
        if fontBackgroundColorName in TransanaGlobal.transana_colorNameList:
            self._fontBackgroundColorName = fontColorName
            # Set fontBackgroundColorDef to match fontColorName
            for (colorName, colorDef) in TransanaGlobal.transana_textColorList:
                if colorName == fontColorName:
                    self._fontBackgroundColorDef = wx.Colour(colorDef[0], colorDef[1], colorDef[2])
                    break
    def _delFontBackgroundColorName(self):
        self._fontBackgroundColorName = None
        self._fontBackgroundColorDef = None

    def _getFontBackgroundColorDef(self):
        return self._fontBackgroundColorDef
    def _setFontBackgroundColorDef(self, fontBackgroundColorDef):
        self._fontBackgroundColorDef = fontBackgroundColorDef
        # Set fontBackgroundColorName to match fontBackgroundColorDef
        for (colorName, colorDef) in TransanaGlobal.transana_textColorList:
            if colorDef == fontBackgroundColorDef:
                self._fontBackgroundColorName = colorName
                break
    def _delFontBackgroundColorDef(self):
        self._fontBackgroundColorDef = None
        self._fontBackgroundColorName = None

    def _getParagraphAlignment(self):
        return self._paragraphAlignment
    def _setParagraphAlignment(self, paragraphAlignment):
        self._paragraphAlignment = paragraphAlignment
    def _delParagraphAlignment(self):
        self._paragraphAlignment = None

    def _getParagraphLeftIndent(self):
        return self._paragraphLeftIndent
    def _setParagraphLeftIndent(self, paragraphLeftIndent):
        self._paragraphLeftIndent = paragraphLeftIndent
    def _delParagraphLeftIndent(self):
        self._paragraphLeftIndent = None

    def _getParagraphLeftSubIndent(self):
        return self._paragraphLeftSubIndent
    def _setParagraphLeftSubIndent(self, paragraphLeftSubIndent):
        self._paragraphLeftSubIndent = paragraphLeftSubIndent
    def _delParagraphLeftSubIndent(self):
        self._paragraphLeftSubIndent = None

    def _getParagraphRightIndent(self):
        return self._paragraphRightIndent
    def _setParagraphRightIndent(self, paragraphRightIndent):
        self._paragraphRightIndent = paragraphRightIndent
    def _delParagraphRightIndent(self):
        self._paragraphRightIndent = None

    def _getParagraphLineSpacing(self):
        return self._paragraphLineSpacing
    def _setParagraphLineSpacing(self, paragraphLineSpacing):
        self._paragraphLineSpacing = paragraphLineSpacing
    def _delParagraphLineSpacing(self):
        self._paragraphLineSpacing = None

    def _getParagraphSpaceBefore(self):
        return self._paragraphSpaceBefore
    def _setParagraphSpaceBefore(self, paragraphSpaceBefore):
        self._paragraphSpaceBefore = paragraphSpaceBefore
    def _delParagraphSpaceBefore(self):
        self._paragraphSpaceBefore = None

    def _getParagraphSpaceAfter(self):
        return self._paragraphSpaceAfter
    def _setParagraphSpaceAfter(self, paragraphSpaceAfter):
        self._paragraphSpaceAfter = paragraphSpaceAfter
    def _delParagraphSpaceAfter(self):
        self._paragraphSpaceAfter = None

    def _getTabs(self):
        return self._tabs
    def _setTabs(self, tabs):
        self._tabs = tabs
    def _delTabs(self):
        self._tabs = []
        
        
    # Public properties
    fontFace      = property(_getFontFace,      _setFontFace,      _delFontFace,      """ Font Face """)
    fontSize      = property(_getFontSize,      _setFontSize,      _delFontSize,      """ Font Size """)
    fontWeight    = property(_getFontWeight,    _setFontWeight,    _delFontWeight,    """ Font Weight [fd_OFF (0), fd_BOLD (1), fd_AMBIGUOUS (2)] """)
    fontStyle     = property(_getFontStyle,     _setFontStyle,     _delFontStyle,     """ Font Style [fd_OFF (0), fd_ITALIC (1), fd_AMBIGUOUS (2)] """)
    fontUnderline = property(_getFontUnderline, _setFontUnderline, _delFontUnderline, """ Font Underline [fd_OFF (0), fd_UNDERLINE (1), fd_AMBIGUOUS (2)] """)
    fontColorName = property(_getFontColorName, _setFontColorName, _delFontColorName, """ Font Color Name """)
    fontColorDef  = property(_getFontColorDef,  _setFontColorDef,  _delFontColorDef,  """ Font Color Definition """)
    fontBackgroundColorName = property(_getFontBackgroundColorName, _setFontBackgroundColorName, _delFontBackgroundColorName, """ Background Color Name """)
    fontBackgroundColorDef  = property(_getFontBackgroundColorDef,  _setFontBackgroundColorDef,  _delFontBackgroundColorDef,  """ Background Color Definition """)
    
    paragraphAlignment = property(_getParagraphAlignment, _setParagraphAlignment, _delParagraphAlignment, "Paragraph Alignment")
    paragraphLeftIndent = property(_getParagraphLeftIndent, _setParagraphLeftIndent, _delParagraphLeftIndent, "Paragraph Left Indent")
    paragraphLeftSubIndent = property(_getParagraphLeftSubIndent, _setParagraphLeftSubIndent, _delParagraphLeftSubIndent, "Paragraph Left SubIndent")
    paragraphRightIndent = property(_getParagraphRightIndent, _setParagraphRightIndent, _delParagraphRightIndent, "Paragraph Right Indent")
    paragraphLineSpacing = property(_getParagraphLineSpacing, _setParagraphLineSpacing, _delParagraphLineSpacing, "Paragraph Line Spacing")
    paragraphSpaceBefore = property(_getParagraphSpaceBefore, _setParagraphSpaceBefore, _delParagraphSpaceBefore, "Paragraph Space Before")
    paragraphSpaceAfter = property(_getParagraphSpaceAfter, _setParagraphSpaceAfter, _delParagraphSpaceAfter, "Paragraph Space After")
    tabs = property(_getTabs, _setTabs, _delTabs, "Tabs")


class FormatDialog(wx.Dialog):
    """ Format Font and Paragraph properties """

    def __init__(self, parent, formatData, tabToShow=0):
        self.parent = parent

        wx.Dialog.__init__(self, parent, -1,  _('Format'), style=wx.CAPTION | wx.SYSTEM_MENU | wx.THICK_FRAME)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        self.formatData = formatData

        # Create the main Sizer, which will hold the boxTop, boxMiddle, and boxButton sizers
        box = wx.BoxSizer(wx.VERTICAL)
        notebook = wx.Notebook(self, -1)
        box.Add(notebook, 1, wx.EXPAND, 2)

        self.panelFont = FormatFontPanel.FormatFontPanel(notebook, formatData)

        self.panelFont.SetAutoLayout(True)
        self.panelFont.Layout()

        notebook.AddPage(self.panelFont, _("Font"), True)


        self.panelParagraph = FormatParagraphPanel.FormatParagraphPanel(notebook, formatData)

        self.panelParagraph.SetAutoLayout(True)
        self.panelParagraph.Layout()

        notebook.AddPage(self.panelParagraph, _("Paragraph"), True)

        self.panelTabs = FormatTabsPanel.FormatTabsPanel(notebook, formatData)

        self.panelTabs.SetAutoLayout(True)
        self.panelTabs.Layout()

        notebook.AddPage(self.panelTabs, _("Tabs"), True)

        if tabToShow != notebook.GetSelection():
            notebook.SetSelection(tabToShow)

        # Create the boxButtons sizer, which will hold the dialog box's buttons
        boxButtons = wx.BoxSizer(wx.HORIZONTAL)

        # Create an OK button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        btnOK.SetDefault()
        btnOK.Bind(wx.EVT_BUTTON, self.OnOK)
        boxButtons.Add(btnOK, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT, 20)

        # Create a Cancel button
        btnCancel = wx.Button(self, wx.ID_CANCEL, _("Cancel"))
        btnCancel.Bind(wx.EVT_BUTTON, self.OnCancel)
        boxButtons.Add(btnCancel, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM)

        # Add the boxButtons sizer to the main box sizer
        box.Add(boxButtons, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.ALL, 10)

        # Define box as the form's main sizer
        self.SetSizer(box)
        # Fit the form to the widgets created
        self.Fit()
        # Set this as the minimum size for the form.
        self.SetSizeHints(minW = self.GetSize()[0], minH = self.GetSize()[1])
        # Tell the form to maintain the layout and have it set the intitial Layout
        self.SetAutoLayout(True)
        self.Layout()

        # Position the form in the center of the screen
        self.CentreOnScreen()

        # Update the Sample Font to match the initial settings
        wx.FutureCall(50, self.panelFont.SetSampleFont)


    def OnOK(self, event):
        """ OK Button Press """
        # When the OK button is pressed, we take the local data (currentFont and currentColor) and
        # put that information into the formatData structure.
        self.formatData.fontFace = self.panelFont.txtFont.GetValue()
        self.formatData.fontSize = self.panelFont.txtSize.GetValue()
        self.formatData.fontWeight = self.panelFont.font.fontWeight
        self.formatData.fontStyle = self.panelFont.font.fontStyle
        self.formatData.fontUnderline = self.panelFont.font.fontUnderline
        self.formatData.fontColorDef = self.panelFont.font.fontColorDef
        self.formatData.fontBackgroundColorDef = self.panelFont.font.fontBackgroundColorDef
        self.formatData.paragraphAlignment = self.panelParagraph.formatData.paragraphAlignment

        # Dealing with Margins is tricky.
        try:
            # If AMBIGUOUS ...
            if self.panelParagraph.txtLeftIndent.GetValue() == '':
                # ... set the value to NONE
                leftVal = None
            # Otherwise ...
            else:
                # ... get the value from the panel control
                leftVal = float(self.panelParagraph.txtLeftIndent.GetValue())
        # If an exception is raised in the float() conversion ...
        except:
            # ... assume illegal values equal 0
            leftVal = 0.0
        try:
            # If AMBIGUOUS ...
            if self.panelParagraph.txtFirstLineIndent.GetValue() == '':
                # ... set the value to NONE
                firstLineVal = None
            # Otherwise ...
            else:
                # ... get the value from the panel control
                firstLineVal = float(self.panelParagraph.txtFirstLineIndent.GetValue())
        # If an exception is raised in the float() conversion ...
        except:
            # ... assume illegal values equal 0
            firstLineVal = 0.0
        try:
            # If AMBIGUOUS ...
            if self.panelParagraph.txtRightIndent.GetValue() == '':
                # ... set the value to NONE
                rightVal = None
            # Otherwise ...
            else:
                # ... get the value from the panel control
                rightVal = float(self.panelParagraph.txtRightIndent.GetValue())
        # If an exception is raised in the float() conversion ...
        except:
            # ... assume illegal values equal 0
            rightVal = 0.0

        # The calculations are straight-forward, but the possibility of ambiguous values make that much more complicated.
#        leftSubIndent = 0.0 - firstLineVal
#        leftIndent = leftVal + firstLineVal
#        rightIndent = rightVal

#        print
#        print "FormatDialog:"

        # If the First Line Value is ambiguous ...
        if firstLineVal == None:
            # ... pass the ambiguity on to the calling routine
            self.formatData.paragraphLeftSubIndent = None
        # If the First Line Value is known ...
        else:
            # ... translate the units into tenths of a milimeter, which is what the RTC uses.
            # So if we are using INCHES ...
            if self.panelParagraph.rbUnitsInches.GetValue():
                # ... convert to Left Sub and convert to 10ths of a milimeter
                self.formatData.paragraphLeftSubIndent = (0.0 - firstLineVal) * 254.0
            # If we are using CENTIMETERS ...
            else:
                # ... convert to Left Sub and convert to 10ths of a milimeter
                self.formatData.paragraphLeftSubIndent = (0.0 - firstLineVal) * 100.0

#            print " Converted First Line Value:", self.formatData.paragraphLeftSubIndent

        # If the Left Value is ambiguous ...
        if leftVal == None:
            # ... pass the ambiguity on to the calling routine
            self.formatData.paragraphLeftIndent = None
            self.formatData.paragraphLeftValue = None
        # If the Left Value is known ...
        else:
            # ... but we don't know the first line values ...
            if firstLineVal == None:
                # ... then the Left Indent is STILL ambiguous!
                self.formatData.paragraphLeftIndent = None
                # ... but we DO know HALF of the value of the calculation and need to take that into account.
                # Thus, let's add it to the object.

                # ... translate the units into tenths of a milimeter, which is what the RTC uses.
                # So if we are using INCHES ...
                if self.panelParagraph.rbUnitsInches.GetValue():
                    # ... convert to 10ths of a milimeter
                    self.formatData.paragraphLeftValue =  leftVal * 254.0
                # If we are using CENTIMETERS ...
                else:
                    # ... convert to 10ths of a milimeter
                    self.formatData.paragraphLeftValue = leftVal * 100.0
            # ... and the First Line Value IS known ...
            else:
                # ... then the Left Indent value is also known

                # ... translate the units into tenths of a milimeter, which is what the RTC uses.
                # So if we are using INCHES ...
                if self.panelParagraph.rbUnitsInches.GetValue():
                    # ... convert to 10ths of a milimeter
                    self.formatData.paragraphLeftIndent = (leftVal + firstLineVal) * 254.0
                # If we are using CENTIMETERS ...
                else:
                    # ... convert to 10ths of a milimeter
                    self.formatData.paragraphLeftIndent = (leftVal + firstLineVal) * 100.0

#                print " Converted Left Indent:", self.formatData.paragraphLeftIndent, leftVal, firstLineVal, leftVal + firstLineVal
#        print "Line Spacing:", self.formatData.paragraphLineSpacing
#        print

        # If the Right Value is ambiguous ...
        if rightVal == None:
            # ... pass the ambiguity on to the calling routine
            self.formatData.paragraphRightIndent = None
        # If the Right Value is known ...
        else:
            # ... translate the units into tenths of a milimeter, which is what the RTC uses.
            # So if we are using INCHES ...
            if self.panelParagraph.rbUnitsInches.GetValue():
                # ... convert to 10ths of a milimeter
                self.formatData.paragraphRightIndent = rightVal * 254.0
            # If we are using CENTIMETERS ...
            else:
                # ... convert to 10ths of a milimeter
                self.formatData.paragraphRightIndent = rightVal * 100.0
        
        self.formatData.paragraphLineSpacing = self.panelParagraph.formatData.paragraphLineSpacing

#        try:
#            self.formatData.paragraphSpaceBefore = int(self.panelParagraph.txtSpacingBefore.GetValue())
#        except:
#            self.formatData.paragraphSpaceBefore = None

        # If the Space Before is ambiguous ...
        if self.formatData.paragraphSpaceBefore == None:
            # ... pass the ambiguity on to the calling routine
            self.formatData.paragraphSpaceBefore = None
        # If the Space Before is known ...
        else:
            # ... translate the units into tenths of a milimeter, which is what the RTC uses.
            # So if we are using INCHES ...
            if self.panelParagraph.rbUnitsInches.GetValue():
                # ... convert to 10ths of a milimeter
                self.formatData.paragraphSpaceBefore = float(self.panelParagraph.txtSpacingBefore.GetValue()) * 254.0
            # If we are using CENTIMETERS ...
            else:
                # ... convert to 10ths of a milimeter
                self.formatData.paragraphSpaceBefore = float(self.panelParagraph.txtSpacingBefore.GetValue()) * 100.0

#        try:
#            self.formatData.paragraphSpaceAfter = int(self.panelParagraph.txtSpacingAfter.GetValue())
#        except:
#            self.formatData.paragraphSpaceAfter = None

        # If the Space After is ambiguous ...
        if self.formatData.paragraphSpaceAfter == None:
            # ... pass the ambiguity on to the calling routine
            self.formatData.paragraphSpaceAfter = None
        # If the Space After is known ...
        else:
            # ... translate the units into tenths of a milimeter, which is what the RTC uses.
            # So if we are using INCHES ...
            if self.panelParagraph.rbUnitsInches.GetValue():
                # ... convert to 10ths of a milimeter
                self.formatData.paragraphSpaceAfter = float(self.panelParagraph.txtSpacingAfter.GetValue()) * 254.0
            # If we are using CENTIMETERS ...
            else:
                # ... convert to 10ths of a milimeter
                self.formatData.paragraphSpaceAfter = float(self.panelParagraph.txtSpacingAfter.GetValue()) * 100.0

        tmpTabs = self.panelTabs.lbTabStops.GetItems()
        try:
            # See if there's a value in the "To Add" entry box that didn't actually get entered.
            # First, get the value.
            val = float(self.panelTabs.txtAdd.GetValue())
            # IF the value is > 0 and not already in the list ...
            if (val > 0) and (not "%4.2f" % val in tmpTabs):
                # ... add it to the list ...
                tmpTabs.append("%4.2f" % val)
                # ... and sort the list.
                tmpTabs.sort()
        except:
            pass
        newTabs = []
        for tab in tmpTabs:
            if self.panelTabs.rbUnitsInches.GetValue():
                newTabs.append(float(tab) * 254.0)
            else:
                newTabs.append(float(tab) * 100.0)
        # If we sent tabs in OR we got tabs out ...
        if (self.formatData.tabs != None) or (len(newTabs) > 0):
            # ... update the tab stops
            self.formatData.tabs = newTabs
                            
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
        
    def GetFormatDef(self):
        """ This method allows the calling routine to access the FormatDialog.FormatDef. This includes background color. """
        # If the user pressed "OK", the originalFont has been updated to reflect the changes.
        # Otherwise, it is the unchanged original data.
        return self.formatData
        

        
