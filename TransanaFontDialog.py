# Copyright (C) 2003 - 2014 The Board of Regents of the University of Wisconsin System 
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

""" This dialog implements the Transana Font Dialog class. The wxPython wxFontDialog 
    proved inadequate for our needs, especially on the Mac.  It is modeled after wxFontDialog.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Enable (True) or Disable (False) debugging messages
DEBUG = False
if DEBUG:
    print "TransanaFontDialog DEBUG is ON"

# For testing purposes, this module can run stand-alone.
if __name__ == '__main__':
    import wxversion
    wxversion.select(['2.6-unicode'])

# import wxPython
import wx

if __name__ == '__main__':
    # This module expects i18n.  Enable it here.
    __builtins__._ = wx.GetTranslation

# Import Transana's Constants
import TransanaConstants
# import the TransanaGlobal variables
import TransanaGlobal

# Define Transana Font Definition Constants
tfd_OFF = 0           # The "unselected" state for font Weight, Style, and Underline
tfd_BOLD = 1          # The "selected" state for Weight
tfd_ITALIC = 1        # The "selected" state for Style
tfd_UNDERLINE = 1     # The "selected" state for Underline
tfd_AMBIGUOUS = 2     # The "ambiguous" (mixed unselected and selected) state for font Weight, Style, and Underline


class TransanaFontDef(object):
    """ The Transana Font Dialog uses its own custom font definition object so that it is able
        to handle ambiguous font information.  For example, if some of the text to be set by
        this font specification is Bold and some is not, the Bold setting should be unset.
        The traditional wx.FontData object can't do this.  """

    def __init__(self):
        """ Initialize the TransanaFontDef object """
        self._fontFace = None
        self._fontSize = None
        self._fontWeight = tfd_OFF       # Options are tfd_OFF, tfd_BOLD,      tfd_AMBIGUOUS  
        self._fontStyle = tfd_OFF        # Options are tfd_OFF, tfd_ITALIC,    tfd_AMBIGUOUS
        self._fontUnderline = tfd_OFF    # Options are tfd_OFF, tfd_UNDERLINE, tfd_AMBIGUOUS
        self._fontColorName = None
        self._fontColorDef = None
        self._backgroundColorName = None
        self._backgroundColorDef = None

    def __repr__(self):
        """ String Representation of the contents of the TransanaFontDef object """
        st = 'TransanaFontDef object:\n'
        st += 'fontFace: %s\n' % self.fontFace
        st += 'fontSize: %s\n' % self.fontSize
        st += 'fontWeight: '
        if self.fontWeight == tfd_OFF:
            st += 'tfd_OFF'
        elif self.fontWeight == tfd_BOLD:
            st += 'tfd_BOLD'
        elif self.fontWeight == tfd_AMBIGUOUS:
            st += 'tfd_AMBIGUOUS'
        else:
            st += 'ILLEGAL SETTING "%s"' % self.fontWeight
        st += '\n'
        st += 'fontStyle: '
        if self.fontStyle == tfd_OFF:
            st += 'tfd_OFF'
        elif self.fontStyle == tfd_ITALIC:
            st += 'tfd_ITALIC'
        elif self.fontStyle == tfd_AMBIGUOUS:
            st += 'tfd_AMBIGUOUS'
        else:
            st += 'ILLEGAL SETTING "%s"' % self.fontStyle
        st += '\n'
        st += 'fontUnderline: '
        if self.fontUnderline == tfd_OFF:
            st += 'tfd_OFF'
        elif self.fontUnderline == tfd_UNDERLINE:
            st += 'tfd_UNDERLINE'
        elif self.fontUnderline == tfd_AMBIGUOUS:
            st += 'tfd_AMBIGUOUS'
        else:
            st += 'ILLEGAL SETTING "%s"' % self.fontUnderline
        st += '\n'
        st += 'fontColorName: %s\n' % self.fontColorName
        st += 'fontColorDef: %s\n' % (self.fontColorDef,)
        st += 'backgroundColorName: %s\n' % self.backgroundColorName
        st += 'backgroundColorDef: %s\n\n' % (self.backgroundColorDef,)
        return st

    def copy(self):
        """ Create a copy of a TransanaFontDef object """
        # Create a new TransanaFontDef object
        tfdCopy = TransanaFontDef()
        # Copy the existing data values to the new Object
        tfdCopy.fontFace = self.fontFace
        tfdCopy.fontSize = self.fontSize
        tfdCopy.fontWeight = self.fontWeight
        tfdCopy.fontStyle = self.fontStyle
        tfdCopy.fontUnderline = self.fontUnderline
        # We don't need to copy fontColorName.  Copying fontColorDef will take care of it.
        tfdCopy.fontColorDef = self.fontColorDef
        tfdCopy.backgroundColorDef = self.backgroundColorDef
        # Return the new Object
        return tfdCopy

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
            self._fontSize = int(fontSize)
        except:
            pass
    def _delFontSize(self):
        self._fontSize = None
        
    def _getFontWeight(self):
        return self._fontWeight
    def _setFontWeight(self, fontWeight):
        if fontWeight in [tfd_OFF, tfd_BOLD, tfd_AMBIGUOUS]:
            self._fontWeight = fontWeight
    def _delFontWeight(self):
        self._fontWeight = tfd_OFF

    def _getFontStyle(self):
        return self._fontStyle
    def _setFontStyle(self, fontStyle):
        if fontStyle in [tfd_OFF, tfd_ITALIC, tfd_AMBIGUOUS]:
            self._fontStyle = fontStyle
    def _delFontStyle(self):
        self._fontStyle = tfd_OFF

    def _getFontUnderline(self):
        return self._fontUnderline
    def _setFontUnderline(self, fontUnderline):
        if fontUnderline in [tfd_OFF, tfd_UNDERLINE, tfd_AMBIGUOUS]:
            self._fontUnderline = fontUnderline
    def _delFontUnderline(self):
        self._fontUnderline = tfd_OFF

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

    def _getBackgroundColorName(self):
        return self._backgroundColorName
    def _setBackgroundColorName(self, backgroundColorName):
        if backgroundColorName in TransanaGlobal.transana_colorNameList:
            self._backgroundColorName = fontColorName
            # Set backgroundColorDef to match fontColorName
            for (colorName, colorDef) in TransanaGlobal.transana_textColorList:
                if colorName == fontColorName:
                    self._backgroundColorDef = wx.Colour(colorDef[0], colorDef[1], colorDef[2])
                    break
    def _delBackgroundColorName(self):
        self._backgroundColorName = None
        self._backgroundColorDef = None

    def _getBackgroundColorDef(self):
        return self._backgroundColorDef
    def _setBackgroundColorDef(self, backgroundColorDef):
        self._backgroundColorDef = backgroundColorDef
        # Set backgroundColorName to match backgroundColorDef
        for (colorName, colorDef) in TransanaGlobal.transana_textColorList:
            if colorDef == backgroundColorDef:
                self._backgroundColorName = colorName
                break
    def _delBackgroundColorDef(self):
        self._backgroundColorDef = None
        self._backgroundColorName = None

    # Public properties
    fontFace      = property(_getFontFace,      _setFontFace,      _delFontFace,      """ Font Face """)
    fontSize      = property(_getFontSize,      _setFontSize,      _delFontSize,      """ Font Size """)
    fontWeight    = property(_getFontWeight,    _setFontWeight,    _delFontWeight,    """ Font Weight [tfd_OFF (0), tfd_BOLD (1), tfd_AMBIGUOUS (2)] """)
    fontStyle     = property(_getFontStyle,     _setFontStyle,     _delFontStyle,     """ Font Style [tfd_OFF (0), tfd_ITALIC (1), tfd_AMBIGUOUS (2)] """)
    fontUnderline = property(_getFontUnderline, _setFontUnderline, _delFontUnderline, """ Font Underline [tfd_OFF (0), tfd_UNDERLINE (1), tfd_AMBIGUOUS (2)] """)
    fontColorName = property(_getFontColorName, _setFontColorName, _delFontColorName, """ Font Color Name """)
    fontColorDef  = property(_getFontColorDef,  _setFontColorDef,  _delFontColorDef,  """ Font Color Definition """)
    backgroundColorName = property(_getBackgroundColorName, _setBackgroundColorName, _delBackgroundColorName, """ Background Color Name """)
    backgroundColorDef  = property(_getBackgroundColorDef,  _setBackgroundColorDef,  _delBackgroundColorDef,  """ Background Color Definition """)
    

class TransanaFontDialog(wx.Dialog):
    """ Transana's custom Font Dialog Box.  Pass in a wxFontData object (to maintain compatability with the wxFontDialog) or
        a TransanaFontDef object to allow for ambiguity in the font specification.  """
    
    def __init__(self, parent, fontData, bgColor=wx.Colour(255, 255, 255), sampleText='AaBbCc ... XxYyZz'):
        """ Initialize the Font Dialog Box.  fontData can either be a wxFontData object or a TransanaFontDef object.
            Use a TransanaFontDef object if some values are ambiguous due to conflicting settings in the selected text.  """
        # Set the initial font data values, depending on the type of object passed in.
        # At this point, if wx.FontData is passed in, use that to populate a parallel TransanaFontData
        # object, and if a TransanaFontData object is passed in, use it to populate a wx.FontData
        # object.  Later, we may eliminate the wx.FontData object, but not yet.
        if type(fontData) == type(wx.FontData()):
            # Define the initial font data
            self.font = TransanaFontDef()
            # Make the fontData available to the whole object
            self.fontData = fontData
            # Let's pull out a local copy of the font to manipulate
            self.currentFont = fontData.GetInitialFont()
            # Populate the TransanaFontDef object with the same information.
            self.font.fontFace = self.currentFont.GetFaceName()
            self.font.fontSize = self.currentFont.GetPointSize()
            if self.currentFont.GetWeight() == wx.BOLD:
                self.font.fontWeight = tfd_BOLD
            else:
                self.font.fontWeight = tfd_OFF
            if self.currentFont.GetStyle() == wx.ITALIC:
                self.font.fontStyle = tfd_ITALIC
            else:
                self.font.fontStyle = tfd_OFF
            if self.currentFont.GetUnderlined():
                self.font.fontUnderline = tfd_UNDERLINE
            else:
                self.font.fontUnderline = tfd_OFF
            # Let's pull out a local copy of the font color to manipulate
            self.currentColor = fontData.GetColour()
            self.font.fontColorDef = fontData.GetColour()
            self.font.backgroundColorDef = bgColor
        elif type(fontData) == type(TransanaFontDef()):
            self.font = fontData
            # TransanaFontData may have undefined data, where wxFontData is always known.
            # Undefined data indicates that the setting is ambiguous, that it has more than
            # one value in the text to be manipulated.  We need to detect this and substitute
            # the best values we can here, as the wxFontData object can't be missing values.

            # If the Font Size if ambiguous ...
            if self.font.fontSize == None:
                # ... subsitute the Transana Default Font Size
                fontSize = TransanaGlobal.configData.defaultFontSize
            else:
                fontSize = self.font.fontSize
            # If the Font Face if ambiguous ...
            if self.font.fontFace == None:
                # ... subsitute the Transana Default Font Face
                fontFace = TransanaGlobal.configData.defaultFontFace
            else:
                fontFace = self.font.fontFace
            # If Weight, Style, or Underline are not explicitly specified, assume they are absent
            # for the wxFontData, as it can't handle the ambiguous state.
            if self.font.fontWeight == tfd_BOLD:
                fontWeight = wx.BOLD
            else:
                fontWeight = wx.NORMAL
            if self.font.fontStyle == tfd_ITALIC:
                fontStyle = wx.ITALIC
            else:
                fontStyle = wx.NORMAL
            if self.font.fontUnderline == tfd_UNDERLINE:
                fontUnderline = True
            else:
                fontUnderline = False
            self.currentColor = self.font.fontColorDef
            self.font.backgroundColor = bgColor
            # Create the appropriate wx.Font object
            self.currentFont = wx.Font(fontSize, wx.FONTFAMILY_DEFAULT, style=fontStyle, weight=fontWeight, underline=fontUnderline, faceName=fontFace)
            # now that we have a wx.Font object, let's create a wx.FontData object
            self.fontData = wx.FontData()
            self.fontData.EnableEffects(True)
            self.fontData.SetInitialFont(self.currentFont)
            self.fontData.SetColour(self.font.fontColorDef)

        # Remember the original TransanaFontDef settings in case the user presses Cancel
        self.originalFont = self.font.copy()

        # Define our Sample Text's Text.
        self.sampleText = sampleText
        # Create the Font Dialog
        wx.Dialog.__init__(self, parent, -1, _('Font'), style=wx.CAPTION | wx.SYSTEM_MENU | wx.THICK_FRAME)

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
            fontFace = self.font.fontFace
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

        # Now create the Font Face Names list box        
        self.lbFont = wx.ListBox(self, -1, choices=fontList, style=wx.LB_SINGLE | wx.LB_ALWAYS_SB | wx.LB_SORT)
        # Make sure the initial font is in the list ...
        if self.font.fontFace != None:
            if self.lbFont.FindString(self.currentFont.GetFaceName()) != wx.NOT_FOUND:
                self.lbFont.SetStringSelection(self.currentFont.GetFaceName())
            else:
                # If not, substitute the platform's default font.   (This is more generic than using Transana's Default Font.)
                tmpFont = wx.Font(self.currentFont.GetPointSize(), wx.DEFAULT, wx.NORMAL, wx.NORMAL)
                # Made sure the default font is in the list.
                if self.lbFont.FindString(tmpFont.GetFaceName()) != wx.NOT_FOUND:
                    self.lbFont.SetStringSelection(tmpFont.GetFaceName())
                else:
                    # If the default font's face can't be found, just pick the first one in the list.
                    self.lbFont.SetSelection(0)
        self.lbFont.Bind(wx.EVT_LISTBOX, self.OnLbFontChange)
        boxFont.Add(self.lbFont, 3, wx.ALIGN_LEFT | wx.ALIGN_BOTTOM | wx.EXPAND | wx.GROW)

        # Add the boxFont sizer to the boxTop sizer
        boxTop.Add(boxFont, 5, wx.ALIGN_LEFT | wx.EXPAND)
        # Create the boxSize sizer, which will hold the Font Size widgets
        boxSize = wx.BoxSizer(wx.VERTICAL)

        # Add Font Size widgets.
        # Create the label
        lblSize = wx.StaticText(self, -1, _('Size:'))
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
        if self.font.fontWeight == tfd_OFF:
            checkValue = wx.CHK_UNCHECKED
        elif self.font.fontWeight == tfd_BOLD:
            checkValue = wx.CHK_CHECKED
        elif self.font.fontWeight == tfd_AMBIGUOUS:
            checkValue = wx.CHK_UNDETERMINED
        self.checkBold.Set3StateValue(checkValue)
        self.checkBold.Bind(wx.EVT_CHECKBOX, self.OnBold)

        boxStyle.Add(self.checkBold, 0, wx.ALIGN_LEFT | wx.LEFT, 5)
        boxStyle.Add((0, 5))  # Spacer

        # Add a checkbox for Italics
        self.checkItalics = wx.CheckBox(self, -1, _('Italics'), style=wx.CHK_3STATE)
        # Determine and set the initial value.
        if self.font.fontStyle == tfd_OFF:
            checkValue = wx.CHK_UNCHECKED
        elif self.font.fontStyle == tfd_ITALIC:
            checkValue = wx.CHK_CHECKED
        elif self.font.fontStyle == tfd_AMBIGUOUS:
            checkValue = wx.CHK_UNDETERMINED
        self.checkItalics.Set3StateValue(checkValue)
        self.checkItalics.Bind(wx.EVT_CHECKBOX, self.OnItalics)
            
        boxStyle.Add(self.checkItalics, 0, wx.ALIGN_LEFT | wx.LEFT, 5)
        boxStyle.Add((0, 5))  # Spacer

        # Add a checkbox for Underline
        self.checkUnderline = wx.CheckBox(self, -1, _('Underline'), style=wx.CHK_3STATE)
        # Determine and set the initial value.
        if self.font.fontUnderline == tfd_OFF:
            checkValue = wx.CHK_UNCHECKED
        elif self.font.fontUnderline == tfd_UNDERLINE:
            checkValue = wx.CHK_CHECKED
        elif self.font.fontUnderline == tfd_AMBIGUOUS:
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
        # Default to Black, if the original color isn't included in the list.
        # NOTE:  This dialog will only support the colors in this list at this point.
        initialColor = _('Black')
        initialBgColor = _('White')
        # Iterate through the list of colors ...
        for (color, colDef) in self.colorList:
            # ... adding each color name to the list of what should be displayed ...
            choiceList.append(_(color))
            # ... and checking to see if the color in the list matches the initial color sent to the dialog.
            if colDef == self.font.fontColorDef:
                # If the current color matches a color in the list, remember it's name.
                initialColor = _(color)
            if colDef == self.font.backgroundColorDef:
                initialBgColor = _(color)

        # Now create a Choice box listing all the colors in the color list
        self.cbColor = wx.Choice(self, -1, choices=choiceList)
        # Set the initial value of the Choice box to the default value determined above.
        if self.font.fontColorName != None:
            self.cbColor.SetStringSelection(initialColor)
        self.cbColor.Bind(wx.EVT_CHOICE, self.OnCbColorChange)
        boxStyle.Add(self.cbColor, 1, wx.ALIGN_LEFT | wx.ALIGN_BOTTOM)

        # If we are using the Rich Text Control, which supports background colors ...
        if TransanaConstants.USESRTC:

            boxStyle.Add((0, 10))  # Spacer

            # Add a label for Color
            lblBgColor = wx.StaticText(self, -1, _('Background Color:'))
            boxStyle.Add(lblBgColor, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP)
            boxStyle.Add((0, 5))  # Spacer

            # Now create a Choice box listing all the colors in the color list
            self.cbBgColor = wx.Choice(self, -1, choices=choiceList)
            # Set the initial value of the Choice box to the default value determined above.
            if self.font.backgroundColorName != None:
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

        # We'll use a Panel for the sample text, painting directly on its Device Context.  The TextCtrl 
        # on the Mac can't handle all we need it to for this task.
        self.txtSample = wx.Panel(self, -1, style=wx.SIMPLE_BORDER)
        self.txtSample.SetBackgroundColour(self.font.backgroundColorDef)
        boxSample.Add(self.txtSample, 1, wx.ALIGN_RIGHT | wx.EXPAND | wx.GROW) 

        # Add the boxSample sizer to the boxMiddle sizer
        boxMiddle.Add(boxSample, 3, wx.ALIGN_RIGHT | wx.EXPAND | wx.GROW)
        # Add the boxMiddle sizer to the main box sizer
        box.Add(boxMiddle, 2, wx.ALIGN_LEFT | wx.EXPAND | wx.GROW | wx.ALL, 10)

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

        # We need an Size event for the form for a little mainenance when the form size is changed
        self.Bind(wx.EVT_SIZE, self.OnSize)
        
        # Under wxPython 2.6.1.0-unicode, this form is throwing a segment fault when the color gets changed.
        # The following variable prevents that!
        self.closing = False

        # Position the form in the center of the screen
        self.CentreOnScreen()
        # Update the Sample Font to match the initial settings
        wx.FutureCall(50, self.SetSampleFont)
        

    def OnOK(self, event):
        """ OK Button Press """
        # When the OK button is pressed, we take the local data (currentFont and currentColor) and
        # put that information into the fontData structure.
        self.fontData.SetChosenFont(self.currentFont)
        self.fontData.SetColour(self.currentColor)
        # If the user presses OK, they want the ALTERED self.font, not the unchanged self.originalFont
        self.originalFont = self.font.copy()
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
        
    def GetFontData(self):
        """ This method allows the calling routine to access the fontData.  This maintains compatibility with wxFontDialog. """
        # If the user pressed "OK", fontData has been updated to reflect the changes.
        # Otherwise, it is the unchanged original data.
        return self.fontData

    def GetFontDef(self):
        """ This method allows the calling routine to access the TransanaFontDef. This includes background color. """
        # If the user pressed "OK", the originalFont has been updated to reflect the changes.
        # Otherwise, it is the unchanged original data.
        return self.originalFont
        
    def SetSampleFont(self):
        """ Update the Sample Text to reflect the dialog's current selections """
        # To get a truly accurate font sample, we need to create a graphic object, paint the
        # font sample on it, and place that on the screen.
        # Get the size of the sample area
        (bmpWidth, bmpHeight) = self.txtSample.GetSize()
        # Create a BitMap that is the right size
        bmp = wx.EmptyBitmap(bmpWidth, bmpHeight)
        # Get the Sample Text Control's Device Context
        cdc = wx.ClientDC(self.txtSample)
        # Prepare the Device Context
        self.PrepareDC(cdc)
        # Link the Control Device Context and the BitMap as a BufferedDC
        dc = wx.BufferedDC(cdc, bmp)
        # if the background color is not ambiguous ...
        if self.font.backgroundColorDef != None:
            # Set the background Color
            dc.SetBackground(wx.Brush(self.font.backgroundColorDef))
        # If the background color is ambiguous ...
        else:
            # ... just use White
            dc.SetBackground(wx.Brush(wx.NamedColour("White")))
        # Clear the Device Context
        dc.Clear()
        # Begin drawing on the Device Context
        dc.BeginDrawing()
        # Set the Device Context's Font
        dc.SetFont(self.currentFont)
        # Set the Device Context's Foreground Color
        dc.SetTextForeground(self.currentColor)
        # Determine how bit our Sample Text is going to be
        (w, h) = dc.GetTextExtent(self.sampleText)
        # Center the Sample Text horizontally and vertically
        x = int(bmpWidth / 2) - int(w / 2)
        y = int(bmpHeight / 2) - int(h / 2)
        # Draw the Sample Text on the Device Context
        dc.DrawText(self.sampleText, x, y)
        # Signal that we're done drawing on the Device Context
        dc.EndDrawing()
        
        if DEBUG:
            print "Displaying Sample Text in %s font" % self.currentFont.GetFaceName()

    def OnTxtFontChange(self, event):
        """ txtFont Change Event.  As the user types in the Font Name, the font ListBox should try to match it. """
        # NOTE:  This method can be improved by having it perform matching with incomplete strings.
        
        # If the typed string matches one of the items in the Font Name List Box ...
        if self.lbFont.FindString(self.txtFont.GetValue()) != wx.NOT_FOUND:
            # ... then update the ListBox to match the typing.  This triggers the font face change.
            self.lbFont.SetStringSelection(self.txtFont.GetValue())
            if DEBUG:
                print "Font changed to %s based on text." % self.txtFont.GetValue()
            # Update the Current Font Face name to match the Font Name selection
            self.currentFont.SetFaceName(self.lbFont.GetStringSelection())
            self.font.fontFace = self.lbFont.GetStringSelection()
            # Update the Font Sample 
            self.SetSampleFont()
            
    def OnTxtFontKillFocus(self, event):
        """ txtFont Kill Focus Event. """
        # If the user leaves the txtFont widget, we need to make sure it has a valid value.
        # Check to see if the current value is a legal Font Face name by comparing it to the Font Face
        # List Box values.  If it's not a legal value ...
        if (not self.closing) and (self.txtFont.GetValue() != self.lbFont.GetStringSelection()):
            # ... revert the text to the last selected font face name.
            self.txtFont.SetValue(self.lbFont.GetStringSelection())
        
    def OnLbFontChange(self, event):
        """ Change Font based on selection in the Font List Box """
        # Update the Font Face text control to match the list box selection
        self.txtFont.SetValue(self.lbFont.GetStringSelection())
        # Update the Current Font Face name to match the Font Name selection
        self.currentFont.SetFaceName(self.lbFont.GetStringSelection())
        self.font.fontFace = self.lbFont.GetStringSelection()
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
                self.txtSize.SetValue(str(self.currentFont.GetPointSize()))
        # Point Sizes need to be between 1 and 255.
        if (size > 0) and (size < 256):
            # If we have a valid value, adjust the font to match the text entry.
            self.currentFont.SetPointSize(size)
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
                self.txtSize.SetValue(str(self.currentFont.GetPointSize()))
        
    def OnLbSizeChange(self, event):
        """ Change Font Size based on selection in the Font Size Box """
        # If the selection is not the first, empty item ...
        if self.lbSize.GetStringSelection() != '':
            # Update the Font Size text control to match the list box selection
            self.txtSize.SetValue(self.lbSize.GetStringSelection())
            # Update the current font to match the list box selection
            self.currentFont.SetPointSize(int(self.lbSize.GetStringSelection()))
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
                fontStyle = tfd_BOLD
                if DEBUG:
                    print "Bold ON"
                # Set the font weight accordingly.
                self.currentFont.SetWeight(style)
            elif self.checkBold.Get3StateValue() == wx.CHK_UNCHECKED:
                # ... otherwise it should be wx.NORMAL
                style = wx.NORMAL
                fontStyle = tfd_OFF
                if DEBUG:
                    print "Bold OFF"
                # Set the font weight accordingly.
                self.currentFont.SetWeight(style)
            else:
                fontStyle = tfd_AMBIGUOUS
                if DEBUG:
                    print "Bold is AMBIGUOUS"
            self.font.fontWeight = fontStyle
        else:
            # If the box is checked ...
            if self.checkBold.IsChecked():
                # ... the Font Weight should be wx.BOLD ...
                style = wx.BOLD
                fontStyle = tfd_BOLD
                if DEBUG:
                    print "Bold ON"
            else:
                # ... otherwise it should be wx.NORMAL
                style = wx.NORMAL
                fontStyle = tfd_OFF
                if DEBUG:
                    print "Bold OFF"
            # Set the font weight accordingly.
            self.currentFont.SetWeight(style)
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
                fontStyle = tfd_ITALIC
                if DEBUG:
                    print "Italics ON"
                # Set the font style accordingly.
                self.currentFont.SetStyle(style)
            elif self.checkItalics.Get3StateValue() == wx.CHK_UNCHECKED:
                # ... otherwise it should be wx.NORMAL
                style = wx.NORMAL
                fontStyle = tfd_OFF
                if DEBUG:
                    print "Italics OFF"
                # Set the font style accordingly.
                self.currentFont.SetStyle(style)
            else:
                fontStyle = tfd_AMBIGUOUS
                if DEBUG:
                    print "Style is AMBIGUOUS"
            self.font.fontStyle = fontStyle
        else:
            # If the box is checked ...
            if self.checkItalics.IsChecked():
                # ... the font style should be wx.ITALIC
                style = wx.ITALIC
                fontStyle = tfd_ITALIC
                if DEBUG:
                    print "Italics ON"
            else:
                # ... otherwise it should be wx.NORMAL
                style = wx.NORMAL
                fontStyle = tfd_OFF
                if DEBUG:
                    print "Italics OFF"
            # Set the font style accordingly
            self.currentFont.SetStyle(style)
            self.font.fontStyle = fontStyle
        # Update the Font Sample 
        self.SetSampleFont()
        
    def OnUnderline(self, event):
        """ Change Underline based on selection in Underline checkbox """
        if self.checkUnderline.Is3State():
            # If the box is checked ...
            if self.checkUnderline.Get3StateValue() in [wx.CHK_CHECKED, wx.CHK_UNCHECKED]:
                # Set the font underline status to match the checkbox
                self.currentFont.SetUnderlined(self.checkUnderline.Get3StateValue() == wx.CHK_CHECKED)
            if self.checkUnderline.Get3StateValue() == wx.CHK_CHECKED:
                fontStyle = tfd_UNDERLINE
            else:
                fontStyle = tfd_OFF
            self.font.fontUnderline = fontStyle
            if DEBUG:
                print "Underlined is now ", self.currentFont.GetUnderlined()
            else:
                if DEBUG:
                    print "Underlined is AMBIGUOUS"
        else:
            # Set the font underline status to match the checkbox
            self.currentFont.SetUnderlined(self.checkUnderline.IsChecked())
            if self.checkUnderline.IsChecked():
                fontStyle = tfd_UNDERLINE
            else:
                fontStyle = tfd_OFF
            self.font.fontUnderline = fontStyle
            if DEBUG:
                print "Underlined is now ", self.currentFont.GetUnderlined()
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
                    self.bgColor = wx.Colour(colDef[0], colDef[1], colDef[2])
                    self.font.backgroundColorDef = wx.Colour(colDef[0], colDef[1], colDef[2])
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
        

# For testing purposes, this module can run stand-alone.
if __name__ == '__main__':
    
    def PrintwxFontInfo(font, color):
        """ print the base information of a font """
        print "Font data based on wx.FontData information:"
        print "  Font Face: %s" % font.GetFaceName()
        print "  Font Size: %d" % font.GetPointSize()
        if font.GetWeight() == wx.BOLD:
            print "  Font Weight:  Bold"
        else:
            print "  Font Weight:  Normal"
        if font.GetStyle() != wx.NORMAL:
            print "  Font Style:  Italics"
        else:
            print "  Font Style:  Normal"
        print "  Font Underline:  %s" % font.GetUnderlined()
        print "  Font Color:  %s" % color
        print

    # Create a simple app for testing.
    app = wx.PySimpleApp()
    
    print wx.PlatformInfo

    # The TransanaFontDialog can work based on a wxFontData Object if everything is known, or based
    # on a TransanaFontDef Object if there is some ambiguity in the Font Specification.
    # This setting determines which option we're testing.
    testType = 'TransanaFontDef' # 'TransanaFontDef' or 'wx.FontData'

    if testType == 'wx.FontData':

        # Create a Font Dialog using the wx.FontData approach

        # create a default font
        font = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        # Alter the font as needed for testing
        # font.SetFaceName('Impact')
        # font.SetPointSize(32)
        # font.SetUnderlined(True)
        # font.SetStyle(wx.ITALIC)
        # font.SetWeight(wx.BOLD)
        # create a wxFontData object, as that is what the Font Dialog works with.
        fontData = wx.FontData()
        # At present, the dialog always displays the effects.
        fontData.EnableEffects(True)
        # Set the fontData initial font to the test font created above
        fontData.SetInitialFont(font)
        # Set the fontData color.  There are very limited options that work here.
        # Valid options are black, blue, cyan, green, grey, light magenta (displays as magenta), red, yellow, and white.
        fontData.SetColour(wx.NamedColour('black'))
        # create the main Font Dialog, passing the fontData we just created

        print
        print "Original font information:"
        PrintwxFontInfo(fontData.GetInitialFont(), fontData.GetColour())

        frame = TransanaFontDialog(None, fontData)

    elif testType == 'TransanaFontDef':

        # Create a Font Dialog using the TransanaFontDef approach

        print
        print
        tfd = TransanaFontDef()
        # Alter the font as needed for testing
        tfd.fontFace = 'Courier New'  # 'Impact'  # 'Times New Roman'
        tfd.fontSize = 12
        tfd.fontWeight = tfd_OFF     # tfd_BOLD
        tfd.fontStyle = tfd_OFF      # tfd_ITALIC
        tfd.fontUnderline = tfd_OFF  # tfd_UNDERLINE
        # Font Color can be set either by giving it a Name or by giving it a Color Definition.
        # We only need one of the following:
        tfd.fontColorName = _('Black')
        # tfd.fontColorDef = wx.NamedColour('red')

        # Create Ambiguous States as needed for testing
#        del(tfd.fontFace)
#        del(tfd.fontSize)
#        tfd.fontWeight = tfd_AMBIGUOUS
#        tfd.fontStyle = tfd_AMBIGUOUS
#        tfd.fontUnderline = tfd_AMBIGUOUS
#        del(tfd.fontColorDef)
        print "Creating frame with this font definition:"
        print tfd

        frame = TransanaFontDialog(None, tfd)

    else:
        print "No testType Specification."

    
    if testType in ['wx.FontData', 'TransanaFontDef']:

        # Show the Dialog Box and process the result.
        # Note that both the wx.FontData object and the TransanaFontDef object
        # can return results regardless of which was used to start the process.
        if frame.ShowModal() == wx.ID_OK:
            print "OK pressed.  Font information should be changed."
            fontData = frame.GetFontData()
            PrintwxFontInfo(fontData.GetChosenFont(), fontData.GetColour())
        else:
            print "Cancel pressed.  Font information should be unchanged."
            # fontData has no GetChosenFont() if the user pressed Cancel!
            fontData = frame.GetFontData()
            PrintwxFontInfo(fontData.GetInitialFont(), fontData.GetColour())
        font = frame.GetFontDef()
        print font

        # Destroy the dialog box.
        frame.Destroy()
        # Call the app's MainLoop()
        app.MainLoop()
