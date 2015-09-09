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

""" This dialog implements the Transana Format Paragraph Panel class.   """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Enable (True) or Disable (False) debugging messages
DEBUG = False
if DEBUG:
    print "FormatParagraphPanel DEBUG is ON"

# import wxPython
import wx
# import the wxPython RichTextCtrl
import wx.richtext as rt

# import the TransanaGlobal variables
import TransanaGlobal


class FormatParagraphPanel(wx.Panel):
    """ Transana's custom Paragraph Dialog Box.  Pass in a wxFontData object (to maintain compatability with the wxFontDialog) or
        a TransanaFontDef object to allow for ambiguity in the font specification.  """
    
    def __init__(self, parent, formatData):
        """ Initialize the Paragraph Panel. """

        self.formatData = formatData.copy()

        # Create the Font Dialog
        wx.Panel.__init__(self, parent, -1)   # , _('Font'), style=wx.CAPTION | wx.SYSTEM_MENU | wx.THICK_FRAME)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the main Sizer, which will hold the boxTop, boxMiddle, and boxButton sizers
        box = wx.BoxSizer(wx.VERTICAL)

        # Paragraph Alignment.
        # Create the label
        lblAlign = wx.StaticText(self, -1, _('Alignment:'))
        box.Add(lblAlign, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.LEFT | wx.TOP, 15)
        box.Add((0, 5))  # Spacer
        # Create a list of alignment options.  (Justified is not yet supported by the RichTextCtrl.
        alignList = [_('Left'), _("Center"), _("Right")]

        # Now create the Font Sizes list box        
        self.lbAlign = wx.Choice(self, -1, choices=alignList)
        if self.formatData.paragraphAlignment == wx.TEXT_ALIGNMENT_LEFT:
            self.lbAlign.SetStringSelection(_("Left"))
        elif self.formatData.paragraphAlignment == wx.TEXT_ALIGNMENT_CENTRE:
            self.lbAlign.SetStringSelection(_("Center"))
        elif self.formatData.paragraphAlignment == wx.TEXT_ALIGNMENT_RIGHT:
            self.lbAlign.SetStringSelection(_("Right"))
        self.lbAlign.Bind(wx.EVT_CHOICE, self.OnAlignSelect)

        # Add the boxTop sizer to the main box sizer
        box.Add(self.lbAlign, 0, wx.ALIGN_LEFT | wx.EXPAND | wx.LEFT | wx.RIGHT, 15)

        # Create the label
        lblIndent = wx.StaticText(self, -1, _('Indentation:'))
        box.Add(lblIndent, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.LEFT | wx.TOP, 15)
        box.Add((0, 5))  # Spacer

        indentSizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left Indent
        leftIndentSizer = wx.BoxSizer(wx.VERTICAL)
        lblLeftIndent = wx.StaticText(self, -1, _("Left:"))
        leftIndentSizer.Add(lblLeftIndent, 0, wx.BOTTOM, 5)
        self.txtLeftIndent = wx.TextCtrl(self, -1, "")
        self.txtLeftIndent.Bind(wx.EVT_CHAR, self.OnNumOnly)
        self.txtLeftIndent.Bind(wx.EVT_KEY_UP, self.OnKeyUp)
        leftIndentSizer.Add(self.txtLeftIndent, 0, wx.EXPAND)
        indentSizer.Add(leftIndentSizer, 1, wx.EXPAND | wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT, 15)

        # First Line Indent
        firstLineIndentSizer = wx.BoxSizer(wx.VERTICAL)
        lblFirstLineIndent = wx.StaticText(self, -1, _("First Line:"))
        firstLineIndentSizer.Add(lblFirstLineIndent, 0, wx.BOTTOM, 5)
        self.txtFirstLineIndent = wx.TextCtrl(self, -1, "")
        self.txtFirstLineIndent.Bind(wx.EVT_CHAR, self.OnNumOnly)
        self.txtFirstLineIndent.Bind(wx.EVT_KEY_UP, self.OnKeyUp)
        firstLineIndentSizer.Add(self.txtFirstLineIndent, 0, wx.EXPAND)
        indentSizer.Add(firstLineIndentSizer, 1, wx.EXPAND | wx.ALIGN_LEFT | wx.RIGHT, 15)

        # Right Indent
        rightIndentSizer = wx.BoxSizer(wx.VERTICAL)
        lblRightIndent = wx.StaticText(self, -1, _("Right:"))
        rightIndentSizer.Add(lblRightIndent, 0, wx.BOTTOM, 5)
        self.txtRightIndent = wx.TextCtrl(self, -1, "")
        self.txtRightIndent.Bind(wx.EVT_CHAR, self.OnNumOnly)
        self.txtRightIndent.Bind(wx.EVT_KEY_UP, self.OnKeyUp)
        rightIndentSizer.Add(self.txtRightIndent, 0, wx.EXPAND)
        indentSizer.Add(rightIndentSizer, 1, wx.EXPAND | wx.ALIGN_LEFT | wx.RIGHT, 15)

        box.Add(indentSizer, 0, wx.EXPAND)

        # Line Spacing
        # Create the label
        lblSpacing = wx.StaticText(self, -1, _('Line Spacing:'))
        box.Add(lblSpacing, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.LEFT | wx.TOP, 15)
        box.Add((0, 5))  # Spacer
        # Create the list of line spacing options
        spacingList = [_('Single'), _('11 point'), _('12 point'), _("One and a half"), _("Double"), _("Two and a half"), _("Triple")]
        # Now create the Line Spacing list box        
        self.lbSpacing = wx.Choice(self, -1, choices=spacingList)
        # if line spacing <= 10 ...
        # If line spacing is 0, let's reset it to 11!
        if self.formatData.paragraphLineSpacing == 0:
            self.formatData.paragraphLineSpacing = 11
        if (self.formatData.paragraphLineSpacing <= wx.TEXT_ATTR_LINE_SPACING_NORMAL):
            self.lbSpacing.SetStringSelection(_("Single"))
        # if line spacing <= 11 ...
        elif (self.formatData.paragraphLineSpacing <= 11):
            self.lbSpacing.SetStringSelection(_("11 point"))
        # if line spacing <= 12 ...
        elif (self.formatData.paragraphLineSpacing <= 12):
            self.lbSpacing.SetStringSelection(_("12 point"))
        # if line spacing <= 15 ...
        elif (self.formatData.paragraphLineSpacing <= wx.TEXT_ATTR_LINE_SPACING_HALF):
            self.lbSpacing.SetStringSelection(_("One and a half"))
        # if line spacing <= 20 ...
        elif (self.formatData.paragraphLineSpacing <= wx.TEXT_ATTR_LINE_SPACING_TWICE):
            self.lbSpacing.SetStringSelection(_("Double"))
        # if line spacing <= 25 ...
        elif (self.formatData.paragraphLineSpacing <= 25):
            self.lbSpacing.SetStringSelection(_("Two and a half"))
        # if line spacing <= 30 ...
        elif (self.formatData.paragraphLineSpacing <= 30):
            self.lbSpacing.SetStringSelection(_("Triple"))
        # if line spacing > 30, something's probably wrong ...
        else:
            # ... so reset line spacing to single spaced.
            self.lbSpacing.SetStringSelection(_("Single"))
            self.formatData.paragraphLineSpacing = 12  # wx.TEXT_ATTR_LINE_SPACING_NORMAL
        # Bind the event for setting line spacing.
        self.lbSpacing.Bind(wx.EVT_CHOICE, self.OnLineSpacingSelect)

        # Add the boxTop sizer to the main box sizer
        box.Add(self.lbSpacing, 0, wx.ALIGN_LEFT | wx.EXPAND | wx.LEFT | wx.RIGHT, 15)

        # Space Before
        paragraphSpacingSizer = wx.BoxSizer(wx.HORIZONTAL)
        spacingBeforeSizer = wx.BoxSizer(wx.VERTICAL)
        lblSpacingBefore = wx.StaticText(self, -1, _("Spacing Before:"))
        spacingBeforeSizer.Add(lblSpacingBefore, 0, wx.TOP, 15)
        spacingBeforeSizer.Add((0, 5))  # Spacer
        self.txtSpacingBefore = wx.TextCtrl(self, -1, "")
        self.txtSpacingBefore.Bind(wx.EVT_CHAR, self.OnNumOnly)
        spacingBeforeSizer.Add(self.txtSpacingBefore, 0, wx.EXPAND)
        paragraphSpacingSizer.Add(spacingBeforeSizer, 1, wx.EXPAND | wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT, 15)

        # Space After
        spacingAfterSizer = wx.BoxSizer(wx.VERTICAL)
        lblSpacingAfter = wx.StaticText(self, -1, _("Spacing After:"))
        spacingAfterSizer.Add(lblSpacingAfter, 0, wx.TOP, 15)
        spacingAfterSizer.Add((0, 5))  # Spacer
        self.txtSpacingAfter = wx.TextCtrl(self, -1, "")
        self.txtSpacingAfter.Bind(wx.EVT_CHAR, self.OnNumOnly)
        spacingAfterSizer.Add(self.txtSpacingAfter, 0, wx.EXPAND)
        paragraphSpacingSizer.Add(spacingAfterSizer, 1, wx.EXPAND | wx.ALIGN_LEFT | wx.RIGHT, 15)

        box.Add(paragraphSpacingSizer, 0, wx.EXPAND)
        box.Add((1, 1), 1, wx.EXPAND)  # Expandable Spacer

        # Error Message Text
        self.errorTxt = wx.StaticText(self, -1, "")
        box.Add(self.errorTxt, 0, wx.EXPAND | wx.GROW | wx.ALIGN_BOTTOM | wx.LEFT | wx.RIGHT | wx.BOTTOM, 15)

        # Units Header
        unitSizer = wx.BoxSizer(wx.HORIZONTAL)
        lblUnits = wx.StaticText(self, -1, _("Units:"))
        unitSizer.Add(lblUnits, 0, wx.RIGHT, 20)

        # Inches 
        self.rbUnitsInches = wx.RadioButton(self, -1, _("inches"), style=wx.RB_GROUP)
        unitSizer.Add(self.rbUnitsInches, 0, wx.RIGHT, 10)

        # Centimeters
        self.rbUnitsCentimeters = wx.RadioButton(self, -1, _("cm"))
        unitSizer.Add(self.rbUnitsCentimeters, 0)
        box.Add(unitSizer, 0, wx.EXPAND | wx.ALIGN_LEFT | wx.LEFT | wx.TOP | wx.RIGHT | wx.BOTTOM, 15)

        if TransanaGlobal.configData.formatUnits == 'cm':
            self.rbUnitsCentimeters.SetValue(True)
        else:
            self.rbUnitsInches.SetValue(True)

        # Bind the event for selecting Units
        self.Bind(wx.EVT_RADIOBUTTON, self.OnIndentUnitSelect)
        # Call the event based on the initial units
        self.OnIndentUnitSelect(None)

        # Define box as the form's main sizer
        self.SetSizer(box)
        # Fit the form to the widgets created
        self.Fit()
        # Set this as the minimum size for the form.
        self.SetSizeHints(minW = self.GetSize()[0], minH = self.GetSize()[1])
        # Tell the form to maintain the layout and have it set the intitial Layout
        self.SetAutoLayout(True)
        self.Layout()

        # Under wxPython 2.6.1.0-unicode, this form is throwing a segment fault when the color gets changed.
        # The following variable prevents that!
        self.closing = False

    def ConvertValueToStr(self, value):
        """  Get a string representation of a value in the appropriate measurement units  """
        # if we have an empty value ...
        if (value == '') or (value == None):
            # ... set the value string to blank
            valStr = ''
        # if we have a value ...
        else:
            # ... convert to the appropriate units
            if self.rbUnitsInches.GetValue():
                value = float(value) / 254.0
            else:
                value = float(value) / 100.0
            # Now convert the floar to a string
            valStr = "%4.3f" % value
        return valStr

    def ConvertStr(self, valStr):
        """  Convert a string representation of a number into the approriate value for the RTC, including converting
             the units appropriately  """
        # If Units is INCHES, the string value is in CM, since this gets called as part of conversion!
        if self.rbUnitsInches.GetValue():
            value = float(valStr) * 100.0
        else:
            value = float(valStr) * 254.0
        return self.ConvertValueToStr(value)

    def OnAlignSelect(self, event):
        """ Handle the Select event for Paragraph Alignment """
        # Convert the control's label into the appropriate wx alignment constant
        if self.lbAlign.GetStringSelection() == unicode(_("Left"), 'utf8'):
            self.formatData.paragraphAlignment = wx.TEXT_ALIGNMENT_LEFT
        elif self.lbAlign.GetStringSelection() == unicode(_("Center"), 'utf8'):
            self.formatData.paragraphAlignment = wx.TEXT_ALIGNMENT_CENTRE
        elif self.lbAlign.GetStringSelection() == unicode(_("Right"), 'utf8'):
            self.formatData.paragraphAlignment = wx.TEXT_ALIGNMENT_RIGHT

    def OnLineSpacingSelect(self, event):
        """ Handle the Select event for Line Spacing """
        # Convert the control's label into the appropriate wx.RTC Line Spacing constant or integer value
        if self.lbSpacing.GetStringSelection() == unicode(_("Single"), 'utf8'):
            self.formatData.paragraphLineSpacing = wx.TEXT_ATTR_LINE_SPACING_NORMAL
        elif self.lbSpacing.GetStringSelection() == unicode(_("11 point"), 'utf8'):
            self.formatData.paragraphLineSpacing = 11
        elif self.lbSpacing.GetStringSelection() == unicode(_("12 point"), 'utf8'):
            self.formatData.paragraphLineSpacing = 12
        elif self.lbSpacing.GetStringSelection() == unicode(_("One and a half"), 'utf8'):
            self.formatData.paragraphLineSpacing = wx.TEXT_ATTR_LINE_SPACING_HALF
        elif self.lbSpacing.GetStringSelection() == unicode(_("Double"), 'utf8'):
            self.formatData.paragraphLineSpacing = wx.TEXT_ATTR_LINE_SPACING_TWICE
        elif self.lbSpacing.GetStringSelection() == unicode(_("Two and a half"), 'utf8'):
            self.formatData.paragraphLineSpacing = 25
        elif self.lbSpacing.GetStringSelection() == unicode(_("Triple"), 'utf8'):
            self.formatData.paragraphLineSpacing = 30

    def OnIndentUnitSelect(self, event):
        """  Handle the selection of one of the Units radio buttons  """
        # The Left Indent from the formatting point of view is the sum of the LeftIndent and LeftSubIndent values!
        if (self.formatData.paragraphLeftIndent != None) and (self.formatData.paragraphLeftSubIndent != None):
            if self.txtLeftIndent.GetValue() == '':
                leftIndentVal = self.formatData.paragraphLeftIndent + self.formatData.paragraphLeftSubIndent
                leftIndentValStr = self.ConvertValueToStr(leftIndentVal)
            else:
                leftIndentValStr = self.ConvertStr(self.txtLeftIndent.GetValue())
        else:
            leftIndentValStr = ''
        self.txtLeftIndent.SetValue(leftIndentValStr)

        # The First Line Indent from the formatting point of view is the negative of the LeftSubIndent values!
        if (self.formatData.paragraphLeftSubIndent != None):
            if self.txtFirstLineIndent.GetValue() == '':
                firstLineIndentVal = 0 - self.formatData.paragraphLeftSubIndent
                firstLineIndentValStr = self.ConvertValueToStr(firstLineIndentVal)
            else:
                firstLineIndentValStr = self.ConvertStr(self.txtFirstLineIndent.GetValue())
        else:
            firstLineIndentValStr = ''
        self.txtFirstLineIndent.SetValue(firstLineIndentValStr)

        # Right Indent
        if (self.formatData.paragraphRightIndent != None):
            if self.txtRightIndent.GetValue() == '':
                # The Right Indent is just the RightIndent value!
                rightIndentVal = self.formatData.paragraphRightIndent
                rightIndentValStr = self.ConvertValueToStr(rightIndentVal)
            else:
                rightIndentValStr = self.ConvertStr(self.txtRightIndent.GetValue())
        else:
            rightIndentValStr = ''
        self.txtRightIndent.SetValue(rightIndentValStr)

        # Spacing Before
        if (self.formatData.paragraphSpaceBefore != None):
            if self.txtSpacingBefore.GetValue() == '':
                spaceBeforeVal = self.formatData.paragraphSpaceBefore
                spaceBeforeValStr = self.ConvertValueToStr(spaceBeforeVal)
            else:
                spaceBeforeValStr = self.ConvertStr(self.txtSpacingBefore.GetValue())
        else:
            spaceBeforeValStr = ''
        self.txtSpacingBefore.SetValue(spaceBeforeValStr)

        # Spacing After
        if (self.formatData.paragraphSpaceAfter != None):
            if self.txtSpacingAfter.GetValue() == '':
                spaceAfterVal = self.formatData.paragraphSpaceAfter
                spaceAfterValStr = self.ConvertValueToStr(spaceAfterVal)
            else:
                spaceAfterValStr = self.ConvertStr(self.txtSpacingAfter.GetValue())
        else:
            spaceAfterValStr = ''
        self.txtSpacingAfter.SetValue(spaceAfterValStr)

        # Update the Configuration data to reflect the selected unit type
        if self.rbUnitsInches.GetValue():
            TransanaGlobal.configData.formatUnits = 'in'
        else:
            TransanaGlobal.configData.formatUnits = 'cm'
        # Save the configuration change immediately
        TransanaGlobal.configData.SaveConfiguration()

    def OnNumOnly(self, event):
        """ EVT_CHAR handler for controls that MUST be numeric values """
        # Determine which control sent the event
        ctrl = event.GetEventObject()

        # Assume we should NOT skip to the control's parent's event handler unless proven otherwise.
        # (The character will NOT be processed unless we call Skip().)
        shouldSkip = False

        # If the ALT, CMD, CTRL, META, or SHIFT key is down ...
        if event.AltDown() or event.CmdDown() or event.ControlDown() or event.MetaDown() or event.ShiftDown():
            # ... call event.Skip()
            event.Skip()

        # If MINUS is pressed ...
        elif event.GetKeyCode() == ord('-'):

            # ... this key is only valid for the FIRST POSITION of the FIRST LINE INDENT control
            if (ctrl.GetId() == self.txtFirstLineIndent.GetId()) and (ctrl.GetInsertionPoint() == 0):
                # ... so if those are the conditions, Skip() is okay.  It's okay to add the character.
                shouldSkip = True

        # if DECIMAL is pressed ...
        elif event.GetKeyCode() == ord('.'):
            # If there is no decimal point already, OR if the decimal is inside the current selection, which will be over-written ...
            if (ctrl.GetValue().find('.') == -1) or (ctrl.GetStringSelection().find('.') > -1):
                # ... then it's okay to add a decimal point
                shouldSkip = True

        # if a DIGIT is pressed ...
        elif event.GetKeyCode() in [ord('0'), ord('1'), ord('2'), ord('3'), ord('4'), ord('5'), ord('6'), ord('7'), ord('8'), ord('9')]:
            # ... then it's okay to add the character
            shouldSkip = True

        # if cursor left, cursor right, backspace, or Delete are pressed ...
        elif event.GetKeyCode() in [wx.WXK_LEFT, wx.WXK_RIGHT, wx.WXK_BACK, wx.WXK_DELETE]:
            # ... then it's okay to add the character
            shouldSkip = True

        # If we should process the character ...
        if shouldSkip:
            # ... then process the character!
            event.Skip()

    def OnKeyUp(self, event):
        """ Handle the EVT_KEY_UP event for margins -- provides error checking """
        # Process the Key Up event at the parent level
        event.Skip()

        # Convert Left Indent to a float
        try:
            leftVal = float(self.txtLeftIndent.GetValue())
        except:
            leftVal = 0.0

        # Convert First Line Indent to a float
        try:
            firstLineVal = float(self.txtFirstLineIndent.GetValue())
        except:
            firstLineVal = 0.0

        # Convert Right Indent to a float
        try:
            rightVal = float(self.txtRightIndent.GetValue())
        except:
            rightVal = 0.0

        # Convert Form Values to RichTextCtrl values
        # ... left subindent is 0 minus first line indent!
        leftSubIndent = 0.0 - firstLineVal
        # ... left indent is left indent plus first line indent!
        leftIndent = leftVal + firstLineVal
        # ... right indent is right indent.
        rightIndent = rightVal

        # Initialize the error message
        errMsg = ''

        # if left indent > 4 inches ...
        if (self.rbUnitsInches.GetValue() and (leftIndent > 4.0)) or \
           (self.rbUnitsCentimeters.GetValue() and (leftIndent > 10.0)):
            # ... suggest that the left margin may be too big!
            errMsg = _("Left Margin may be too large.\n")

        # If left indent < 0 ...
        if (leftIndent < 0.0):
            # ... report that the first line indent is too big.
            errMsg += _("First Line Indent exceeds Left Margin.\n")

        # if right indent > 4 inches ...
        if (self.rbUnitsInches.GetValue() and (rightIndent > 4.0)) or \
           (self.rbUnitsCentimeters.GetValue() and (rightIndent > 10.0)):
            # ... suggest that the right margin may be too big!
            errMsg += _("Right Margin may be too large.\n")

        # Display the Error Message
        self.errorTxt.SetLabel(errMsg)
