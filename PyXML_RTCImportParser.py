# -*- coding: cp1252 -*-
# Copyright (C) 2011-2014 The Board of Regents of the University of Wisconsin System 
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

""" An XML import Parser for the wxPython RichTextCtrl """

# This module was necessary because the wxRichTextCtrl doesn't offer a way to combine XML data.  When
# I went to add a Clip transcript from the database to a Text Report, the RTC would CLEAR out the report
# contents as each transcript was added.  I tried a variety of things, including cut-and-paste of formatted
# text, but that caused formatting problems for the rest of the report.  

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False

# import wxPython
import wx
# Import wxPython's Rich Text Ctrl
import wx.richtext as richtext
# import the Python XML SAX handler
import xml.sax.handler
# import Python's cStringIO for faster string handling
import cStringIO


class XMLToRTCHandler(xml.sax.handler.ContentHandler):
    """ This handler ADDS XML data to an existing RichTextCtrl """

    def __init__(self, txtCtrl, additionalIndents = (0, 0), encoding='utf8'):
        """ Initialize the XML to RichTextCtrl import handler
              Parameters:  txtCtrl            the Rich Text Ctrl that should receive the data
                           additionalIndents  Adjustments to the Left and Right Indent values,
                                              passed as a tuple eg. (127, 127) increases both margins
                                               (Allows changing of original margins for Reports)
                           encoding           Encoding.  Only UTF-8 is tested. """

        # Remember the RTC to modify
        self.txtCtrl = txtCtrl
        # Remember the Additional Indent values
        self.additionalIndents = additionalIndents
        # Remember the encoding
        self.encoding = encoding

        # Define a variable for tracking what element we are changing
        self.element = ''

        # Define an initial Fonts.  We define multiple levels of fonts to handle cascading styles.
        self.fontAttributes = {}
        self.fontAttributes[u'text'] = {u'bgcolor' : '#FFFFFF',
                                    u'fontface' : 'Courier New',
                                    u'fontpointsize' : 12,
                                    u'fontstyle' : wx.FONTSTYLE_NORMAL,
                                    u'fontunderlined' : u'0',
                                    u'fontweight' : wx.FONTSTYLE_NORMAL,
                                    u'textcolor' : '#000000'}

        self.fontAttributes[u'symbol'] = {u'bgcolor' : '#FFFFFF',
                                    u'fontface' : 'Courier New',
                                    u'fontpointsize' : 12,
                                    u'fontstyle' : wx.FONTSTYLE_NORMAL,
                                    u'fontunderlined' : u'0',
                                    u'fontweight' : wx.FONTSTYLE_NORMAL,
                                    u'textcolor' : '#000000'}

        self.fontAttributes[u'paragraph'] = {u'bgcolor' : '#FFFFFF',
                                         u'fontface' : 'Courier New',
                                         u'fontpointsize' : 12,
                                         u'fontstyle' : wx.FONTSTYLE_NORMAL,
                                         u'fontunderlined' : u'0',
                                         u'fontweight' : wx.FONTSTYLE_NORMAL,
                                         u'textcolor' : '#000000'}

        self.fontAttributes[u'paragraphlayout'] = {u'bgcolor' : '#FFFFFF',
                                               u'fontface' : 'Courier New',
                                               u'fontpointsize' : 12,
                                               u'fontstyle' : wx.FONTSTYLE_NORMAL,
                                               u'fontunderlined' : u'0',
                                               u'fontweight' : wx.FONTSTYLE_NORMAL,
                                               u'textcolor' : '#000000'}

        # Define the initial Paragraph attributes.  We define mulitple levels to handle cascading styles.
        self.paragraphAttributes = {}
        self.paragraphAttributes[u'paragraph'] = {u'alignment' : u'1',
                                                  u'linespacing' : u'10',
                                                  u'leftindent' : u'0',
                                                  u'rightindent' : u'0',
                                                  u'leftsubindent' : u'0',
                                                  u'parspacingbefore' : u'0',
                                                  u'parspacingafter' : u'0',
                                                  u'bulletnumber' : None,
                                                  u'bulletstyle' : None,
                                                  u'bulletfont' : None,
                                                  u'bulletsymbol' : None,
                                                  u'bullettext' : None,
                                                  u'tabs' : None}

        self.paragraphAttributes[u'paragraphlayout'] = {u'alignment' : u'1',
                                                      u'linespacing' : u'10',
                                                      u'leftindent' : u'0',
                                                      u'rightindent' : u'0',
                                                      u'leftsubindent' : u'0',
                                                      u'parspacingbefore' : u'0',
                                                      u'parspacingafter' : u'0',
                                                      u'bulletnumber' : None,
                                                      u'bulletstyle' : None,
                                                      u'bulletfont' : None,
                                                      u'bulletsymbol' : None,
                                                      u'bullettext' : None,
                                                      u'tabs' : None}

        # Define an initial font table
        self.fontTable = [u'Courier New']

        # define an initial color table
        self.colorTable = ['#000000', '#FF0000', '#00FF00', '#0000FF', '#FFFFFF']

        # Define a variable for tracking what element we are changing
        self.element = ''

        # Track whether we're inside a Transana time code
        self.inTimeCode = False

        # Handling a URL
        self.url = ''

        # Handling an image
        self.imageData = ''

    def startElement(self, name, attributes):
        """ XML SAX - Handle the start of an XML element """

        # Remember the element's name
        self.element = name

        # If the element is a paragraphlayout, paragraph, symbol, or text element ...
        if name in [u'paragraphlayout', u'paragraph', u'symbol', u'text']:

            # Let's cascade the font and paragraph settings from a level up BEFORE we change things to reset the font and  
            # paragraph settings to the proper initial state.  First, let's create empty character and paragraph cascade lists
            charcascade = paracascade = []
            # Initially, assume we will cascade from our current object for character styles
            cascadesource = name
            # If we're in a Paragraph spec ...
            if name == u'paragraph':
                # ... we need to cascase paragraph, symbol, and text styles for characters ...
                charcascade = [u'paragraph', u'symbol', u'text']
                # ... from the paragraph layout style for characters ...
                cascadesource = u'paragraphlayout'
                # ... and we need to cascare paragraph styles for paragraphs
                paracascade = [u'paragraph']
            # If we're in a Text spec ...
            elif name == u'text':
                # ... we need to cascase text styles for characters ...
                charcascade = [u'text']
                # ... from the paragraph style for characters ...
                cascadesource = u'paragraph'
            # If we're in a Symbol spec ...
            elif name == u'symbol':
                # ... we need to cascase symbol styles for characters ...
                charcascade = [u'symbol']
                # ... from the paragraph style for characters ...
                cascadesource = u'paragraph'
            # For each type of character style we need to cascade ...
            for x in charcascade:
                # ... iterate through the dictionary elements ...
                for y in self.fontAttributes[x].keys():
                    # ... and assign the character cascade source styles (cascadesource) to the destination element (x)
                    self.fontAttributes[x][y] = self.fontAttributes[cascadesource][y]
            # For each type of paragraph style we need to cascade ...
            for x in paracascade:
                # ... iterate through the dictionary elements ...
                for y in self.paragraphAttributes[x].keys():
                    # ... and assign the paragraph cascade source styles (cascadesource) to the destination element (x)
                    self.paragraphAttributes[x][y] = self.paragraphAttributes[cascadesource][y]

            # If the element is a paragraph element or a paragraph layout element, there is extra processing to do at the start
            if name in [u'paragraph', u'paragraphlayout']:
                # ... iterate through the element attributes looking for paragraph attributes
                for x in attributes.keys():
                    # If the attribute is a paragraph format attribute ...
                    if x in [u'alignment',
                             u'linespacing',
                             u'leftindent',
                             u'rightindent',
                             u'leftsubindent',
                             u'parspacingbefore',
                             u'parspacingafter',
                             u'bulletnumber',
                             u'bulletstyle',
                             u'bulletfont',
                             u'bulletsymbol',
                             u'bullettext',
                             u'tabs']:
                        # ... update the current paragraph dictionary
                        self.paragraphAttributes[name][x] = attributes[x]
            # ... iterate through the element attributes looking for font attributes
            for x in attributes.keys():
                if x == u'fontsize':
                    x = u'fontpointsize'
                    # ... update the current font dictionary
                    self.fontAttributes[name][x] = attributes[u'fontsize']
                # If the attribute is a font format attribute ...
                elif x in [u'bgcolor',
                         u'fontface',
                         u'fontpointsize',
                         u'fontstyle',
                         u'fontunderlined',
                         u'fontweight',
                         u'textcolor']:
                    # ... update the current font dictionary
                    self.fontAttributes[name][x] = attributes[x]

                # If the attribute is a font name ...
                if x == u'fontface':
                    # ... that is not already in the font table ...
                    if not(attributes[x] in self.fontTable):
                        # ... add the font name to the font table list
                        self.fontTable.append(attributes[x])

                # If the element is a text element and the attribute is a url attribute ...
                if (name == u'text') and (x == u'url'):
                    # ... capture the URL data.
                    self.url = attributes[x]

            # Let's cascade the font and paragraph settings we've just changed.
            # First, let's create empty character and paragraph cascade lists
            charcascade = paracascade = []
            # Initially, assume we will cascade from our current object for character styles
            cascadesource = name
            # If we're in a Paragraph Layout spec ...
            if name == u'paragraphlayout':
                # ... we need to cascase paragraph, symbol, and text styles for characters ...
                charcascade = [u'paragraph', u'symbol', u'text']
                # ... we need to cascase paragraph styles for paragraphs ...
                paracascade = [u'paragraph']
            # If we're in a Paragraph spec ...
            elif name == u'paragraph':
                # ... we need to cascase symbol and text styles for characters ...
                charcascade = [u'symbol', u'text']
            # For each type of character style we need to cascade ...
            for x in charcascade:
                # ... iterate through the dictionary elements ...
                for y in self.fontAttributes[x].keys():
                    # ... and assign the character cascade source styles (cascadesource) to the destination element (x)
                    self.fontAttributes[x][y] = self.fontAttributes[cascadesource][y]
            for x in paracascade:
                # ... iterate through the dictionary elements ...
                for y in self.paragraphAttributes[x].keys():
                    # ... and assign the paragraph cascade source styles (cascadesource) to the destination element (x)
                    self.paragraphAttributes[x][y] = self.paragraphAttributes[cascadesource][y]

            if DEBUG:
                # List unknown elements
                for x in attributes.keys():
                    if not x in [u'bgcolor',
                                 u'fontface',
                                 u'fontsize',
                                 u'fontpointsize',
                                 u'fontstyle',
                                 u'fontunderlined',
                                 u'fontweight',
                                 u'textcolor',
                                 u'alignment',
                                 u'linespacing',
                                 u'leftindent',
                                 u'rightindent',
                                 u'leftsubindent',
                                 u'parspacingbefore',
                                 u'parspacingafter',
                                 u'url',
                                 u'tabs',
                                 u'bulletnumber',
                                 u'bulletstyle',
                                 u'bulletfont',
                                 u'bulletsymbol',
                                 u'bullettext']:
                        print "Unknown %s attribute:  %s  %s" % (name, x, attributes[x])

        # If the element is an image element ...
        elif name in [u'image']:
            # ... if we have a PNG graphic ...
            if attributes[u'imagetype'] == u'15':  # wx.BITMAP_TYPE_PNG = 15
                # initialize the Image Data
                self.imageData = ''
                # ... signal that we have a PNG image to process ...
                self.elementType = "ImagePNG"
            # It appears to me that all images will be PNG images coming from the RichTextCtrl.
            else:
                # if not, signal a unknown image type
                self.elementType = 'ImageUnknown'

                print "Image of UNKNOWN TYPE!!", attributes.keys()

        # If the element is a data or richtext element ...
        elif name in [u'data', u'richtext']:
            # ... we should do nothing here at this time 
            pass
        # If we have an unhandled element ...
        else:
            # ... output a message and the element attributes.
            print "PyRTFParser.XMLToRTFHandler.startElement():  Unknown XML tag:", name
            for x in attributes.keys():
                print x, attributes[x]
            print

        # If the element is a paragraph element ...
        if name in [u'paragraph']:

            # Code for handling bullet lists and numbered lists is preliminary and probably very buggy

#            print "Bullet Number:", self.paragraphAttributes[u'paragraph'][u'bulletnumber'], type(self.paragraphAttributes[u'paragraph'][u'bulletnumber'])
#            print "Bullet Style:", self.paragraphAttributes[u'paragraph'][u'bulletstyle'],
#            if self.paragraphAttributes[u'paragraph'][u'bulletstyle'] != None:
#                print "%04x" % int(self.paragraphAttributes[u'paragraph'][u'bulletstyle'])
#            else:
#                print
#            print "Bullet Font:", self.paragraphAttributes[u'paragraph'][u'bulletfont']
#            print "Bullet Symbol:", self.paragraphAttributes[u'paragraph'][u'bulletsymbol']
#            print "Bullet Text:", self.paragraphAttributes[u'paragraph'][u'bullettext']
#            print

            # If we have a bullet or numbered list specification ...
            if self.paragraphAttributes[u'paragraph'][u'bulletstyle'] != None:

                print "PyXML_RTCImportParser.startElement(): Paragraph BulletStyle NOT HANDLED"

##            # ... indicate that in the RTF output string
##                self.outputString.write('{\\listtext\\pard\\plain')
##
##                # Convert the Bullet Style to a hex string so we can interpret it correctly.
##                # (I'm sure there's a better way to do this!)
##                styleHexStr = "%04x" % int(self.paragraphAttributes[u'paragraph'][u'bulletstyle'])
##
##                # If we have a known symbol bullet (TEXT_ATTR_BULLET_STYLE_SYMBOL and defined bulletsymbol) ...
##                if (styleHexStr[2] == '2') and (self.paragraphAttributes[u'paragraph'][u'bulletsymbol'] != None):
##                    # ... add that to the RTF Output String
##                    self.outputString.write("\\f%s %s\\tab}" % (self.fontTable.index(self.fontAttributes[name][u'fontface']), chr(int(self.paragraphAttributes[u'paragraph'][u'bulletsymbol']))))
##
##                # if the second characters is a "2", we have richtext.TEXT_ATTR_BULLET_STYLE_STANDARD
##                elif (styleHexStr[1] == '2'):
##                    # If Symbol font is not yet in the Font Table ...
##                    if not 'Symbol' in self.fontTable:
##                        # ... then add it now.
##                        self.fontTable.append('Symbol')
##                    # add the bullet symbol in Symbol font to the RTF Output String
##                    self.outputString.write("\\f%s \\'b7\\tab}" % self.fontTable.index('Symbol'))
##
##                # If we have a know bullet NUMBER (i.e. a numbered list) ...
##                elif self.paragraphAttributes[u'paragraph'][u'bulletnumber'] != None:
##                    # Initialize variables used for presenting the proper "number" style and punctuation
##                    numberChar = ''
##                    numberLeadingChar = ''
##                    numberTrailingChar = ''
##
##                    # Put the bullet "number" into the correct format
##                    # TEXT_ATTR_BULLET_STYLE_ARABIC
##                    if styleHexStr[3] == '1':
##                        numberChar = self.paragraphAttributes[u'paragraph'][u'bulletnumber']
##                    # TEXT_ATTR_BULLET_STYLE_LETTERS_UPPER
##                    elif styleHexStr[3] == '2':
##                        bulletChars = string.uppercase[:26]
##                        numberChar = bulletChars[int(self.paragraphAttributes[u'paragraph'][u'bulletnumber']) - 1]
##                    # TEXT_ATTR_BULLET_STYLE_LETTERS_LOWER
##                    elif styleHexStr[3] == '4':
##                        bulletChars = string.lowercase[:26]
##                        numberChar = bulletChars[int(self.paragraphAttributes[u'paragraph'][u'bulletnumber']) - 1]
##                    # TEXT_ATTR_BULLET_STYLE_ROMAN_UPPER
##                    elif styleHexStr[3] == '8':
##                        numberChar = int2roman(int(self.paragraphAttributes[u'paragraph'][u'bulletnumber']))
##                    # TEXT_ATTR_BULLET_STYLE_ROMAN_LOWER
##                    elif styleHexStr[2] == '1':
##                        numberChar = int2roman(int(self.paragraphAttributes[u'paragraph'][u'bulletnumber'])).lower()
##                        
##                    # Put the bullet "number" into the correct punctuation structure
##                    # TEXT_ATTR_BULLET_STYLE_PERIOD
##                    if styleHexStr[1] == '1':
##                        numberTrailingChar = '.'
##                    # TEXT_ATTR_BULLET_STYLE_RIGHT_PARENTHESIS  
##                    elif styleHexStr[1] == '4':
##                        numberTrailingChar = ')'
##                    # TEXT_ATTR_BULLET_STYLE_PARENTHESIS  
##                    elif styleHexStr[2] == '8':
##                        numberLeadingChar = '('
##                        numberTrailingChar = ')'
##
##                    # ... add that to the RTF Output String
##                    self.outputString.write("\\f%s %s%s%s\\tab}" % (self.fontTable.index(self.fontAttributes[name][u'fontface']), numberLeadingChar, numberChar, numberTrailingChar))
##
##                # If we have a know bullet symbol ...
##                elif self.paragraphAttributes[u'paragraph'][u'bulletsymbol'] != None:
##                    # ... add that to the RTF Output String
##                    self.outputString.write("\\f%s %s\\tab}" % (self.fontTable.index(self.fontAttributes[name][u'fontface']), unichr(int(self.paragraphAttributes[u'paragraph'][u'bulletsymbol']))))
##
##                # If we still don't know what kind of bullet we have, we're in trouble.
##                else:
##                    print "PyRTFParser.startElement() SYMBOL INSERTION FAILURE"
                    
            # Paragraph alignment left is u'1'
            if self.paragraphAttributes[u'paragraph'][u'alignment'] == u'1':
                self.txtCtrl.SetTxtStyle(parAlign = wx.TEXT_ALIGNMENT_LEFT)
            # Paragraph alignment centered is u'2'
            elif self.paragraphAttributes[u'paragraph'][u'alignment'] == u'2':
                self.txtCtrl.SetTxtStyle(parAlign = wx.TEXT_ALIGNMENT_CENTER)
            # Paragraph alignment right is u'3'
            elif self.paragraphAttributes[u'paragraph'][u'alignment'] == u'3':
                self.txtCtrl.SetTxtStyle(parAlign = wx.TEXT_ALIGNMENT_RIGHT)
            else:
                print "Unknown alignment:", self.paragraphAttributes[u'paragraph'][u'alignment'], type(self.paragraphAttributes[u'paragraph'][u'alignment'])

            # line spacing u'10' is single line spacing, which is NOT included in the RTF as it is the default.
            if self.paragraphAttributes[u'paragraph'][u'linespacing'] in [u'0', u'10']:
                self.txtCtrl.SetTxtStyle(parLineSpacing = wx.TEXT_ATTR_LINE_SPACING_NORMAL)
            # 1.5 line spacing is u'15'
            elif self.paragraphAttributes[u'paragraph'][u'linespacing'] == u'15':
                self.txtCtrl.SetTxtStyle(parLineSpacing = wx.TEXT_ATTR_LINE_SPACING_HALF)
            # double line spacing is u'20'
            elif self.paragraphAttributes[u'paragraph'][u'linespacing'] == u'20':
                self.txtCtrl.SetTxtStyle(parLineSpacing = wx.TEXT_ATTR_LINE_SPACING_TWICE)
            else:

##                print "Unknown linespacing:", self.paragraphAttributes[u'paragraph'][u'linespacing'], type(self.paragraphAttributes[u'paragraph'][u'linespacing'])

                self.txtCtrl.SetTxtStyle(parLineSpacing = int(self.paragraphAttributes[u'paragraph'][u'linespacing']))

            # Paragraph Margins and first-line indents
            # First, let's convert the unicode strings we got from the XML to integers and translate from wxRichTextCtrl's
            # system to RTF's system.
            # Left Indent in RTF is the sum of wxRichTextCtrl's left indent and the additional indent adjustment
            leftindent = int(self.paragraphAttributes[u'paragraph'][u'leftindent']) + self.additionalIndents[0]
            # The First Line Indent in RTF is the wxRichTextCtrl's left subindent and the additional indent adjustment.
            firstlineindent = int(self.paragraphAttributes[u'paragraph'][u'leftsubindent'])
            # The Right Indent is the sum of the right indent and the additional indent adjustment
            rightindent = int(self.paragraphAttributes[u'paragraph'][u'rightindent']) + self.additionalIndents[1]
            self.txtCtrl.SetTxtStyle(parLeftIndent = (leftindent, firstlineindent), parRightIndent = rightindent)

            # Add non-zero Spacing before and after paragraphs to the RTF output String
            if int(self.paragraphAttributes[u'paragraph'][u'parspacingbefore']) != 0:
                self.txtCtrl.SetTxtStyle(parSpacingBefore = int(self.paragraphAttributes[u'paragraph'][u'parspacingbefore']))
            if int(self.paragraphAttributes[u'paragraph'][u'parspacingafter']) != 0:
                self.txtCtrl.SetTxtStyle(parSpacingAfter = int(self.paragraphAttributes[u'paragraph'][u'parspacingafter']))

            # If Tabs are defined ...
            if self.paragraphAttributes[u'paragraph'][u'tabs'] != None:
                # ... break the tab data into its component pieces to convert it to a list of intergers
                tabStops = []
                # For each tab stop ...
                for x in self.paragraphAttributes[u'paragraph'][u'tabs'].split(','):
                    # ... (assuming the data isn't empty) ...
                    if x != u'':
                        # ... add the tab stop data to the RTF output string, adjusting for the margin adjustments factor
                        tabStops.append(int(x) + self.additionalIndents[0])
                self.txtCtrl.SetTxtStyle(parTabs = tabStops)

        # Add Font formatting when we process text or symbol tags, as text and symbol specs can modify paragraph-level font specifications
        if name in [u'text', u'symbol']:
            # Add Font Face information
            self.txtCtrl.SetTxtStyle(fontFace = self.fontAttributes[name][u'fontface'],
            # Add Font Size information
            fontSize = int(self.fontAttributes[name][u'fontpointsize']))
            # If bold, add Bold
            if self.fontAttributes[name][u'fontweight'] == str(wx.FONTWEIGHT_BOLD):
                self.txtCtrl.SetTxtStyle(fontBold = True)
            else:
                self.txtCtrl.SetTxtStyle(fontBold = False)
            # If Italics, add Italics
            if self.fontAttributes[name][u'fontstyle'] == str(wx.FONTSTYLE_ITALIC):
                self.txtCtrl.SetTxtStyle(fontItalic = True)
            else:
                self.txtCtrl.SetTxtStyle(fontItalic = False)
            # If Underline, add Underline
            if self.fontAttributes[name][u'fontunderlined'] == u'1':
                self.txtCtrl.SetTxtStyle(fontUnderline = True)
            else:
                self.txtCtrl.SetTxtStyle(fontUnderline = False)
            # ... Add text foreground color
            self.txtCtrl.SetTxtStyle(fontColor = self.fontAttributes[name][u'textcolor'])
            # ... Add text background color
            self.txtCtrl.SetTxtStyle(fontBgColor = self.fontAttributes[name][u'bgcolor'])


    def characters(self, data):
        """ XML SAX - Handle the text between XML element tags """

        # If the characters come from a text element ...
        if self.element in ['text']:

            if DEBUG:
                print "PyXML_RTCImportParser.characters():"
                print '"%s"' % data.encode('utf8')

            # We've had some problems with quotation marks showing up inappropriately.  This attempts to fix that.
            # If we have a single Quotation Mark character in the data specification ...
            if data == '"':
                # ... skip it.
                data = ''
            # If the data is JUST a quotation mark and a space ...
            elif data == '" ':
                # ... drop the quotation mark!
                data = ' '
            # If the data is just a period and a quotation mark ...
            elif data == '."':
                # ... drop the quotation mark!
                data = '.'

            # If the text has leading or trailing spaces, it gets enclosed in quotation marks in the XML.
            # Otherwise, not.  We have to detect this and remove the quotes as needed.  Unicode characters
            # make this a bit more complicated, as in " 137 ë 137 ".
            if ((len(data) >= 2) and ((data[0] == '"') or (data[-1] == '"')) and \
                ((data[0] == ' ') or (data[1] == ' ') or (data[-2] == ' ') or (data[-1] == ' '))):
                # If (quote)(space), check to see if the LAST character inserted was a space.  There is an issue
                # with formatting such that, for example, "(space)<BOLD>WE</BOLD>(space)know" was missing the space, while
                # "(space)<BOLD>WE(space)</BOLD>know" was rendering correctly.
                if (data[0:2] == '" ') and (self.txtCtrl.GetCharAt(self.txtCtrl.GetLastPosition() - 1) == 32):
                    data = data[2:]
                elif data[0] == '"':
                    data = data[1:]
                if data[-1] == '"':
                    data = data[:-1]

            # If data isn't empty ...
            if data != "":
                # Add the text to the RTC
                self.txtCtrl.WriteText(data)

            # If we have a value in self.URL, populated in startElement, ...
            if self.url != '':
                # ... then we need to at least mention the hyperlink.
                self.txtCtrl.WriteText('(%s} ' % self.url)

        # If the characters come from a symbol element ...
        elif self.element == 'symbol':
            # Check that we don't have only whitespace, we don't have a multi-character string, and
            # we don't have a newline character.
            if (len(data.strip()) > 0) and ((len(data) != 1) or (ord(data) != 10)):
                # Convert the symbol data to the appropriate unicode character
                data = unichr(int(data))
                # Add that unicode character to the RTC
                self.txtCtrl.WriteText(data.encode(self.encoding))
        
        # If the characters come from a data element ...
        elif self.element == 'data':
            # If we're expecting a PNG Image ...
            if self.elementType == 'ImagePNG':
                # Image Data may come in several segments.  We accumulate all the data in the imageData
                # variable until we have it all
                self.imageData += data
                                       
            # I haven't seen anything but PNG image data in this data structure from the RichTextCtrl's XML data
            else:
                # If we're dealing with an image, we could convert the image to PNG, then do a Hex conversion.
                # RTF can also JPEG images directly, as well as Enhanced Metafiles, Windows Metafiles, QuickDraw
                # pictures, none of which I think wxPython can handle.
                print "I don't know how to handle the data!!"

        # We can ignore whitespace here, which will be made up of the spaces added by XML and newline characters
        # that are part of the XML file but not part of the data.
        elif data.strip() != '':
            # Otherwise, print a message to the developer
            print "PyXML_RTCImportParser.characters():  Unhandled text."
            print "Element:", self.element
            print 'Data: "%s"' % data

    def endElement(self, name):
        """ XML SAX - Handle the end (close) of an XML element """

        if name in [u'paragraph']:
            # Insert a Newline() to end the paragraph in the RTC
            self.txtCtrl.Newline()

        elif name in [u'image']:
            # When the image tag is closed, we have all the image data and can now insert the image into the RTC

            # Create a StringIO stream from the HEX-converted image data
            stream = cStringIO.StringIO(self.hex2int(self.imageData))
            # Now convert that stream to an image
            img = wx.ImageFromStream(stream, wx.BITMAP_TYPE_PNG)
            # If we were successful in creating a valid image ...
            if img.IsOk():
                # ... add that image to the wxRichTextEdit control
                self.txtCtrl.WriteImage(img)
            else:
                self.txtCtrl.WriteText( "pyXML_RTCImportParser.characters():  Image Import failed")
                self.txtCtrl.Newline()

            # Destroy the Image
            img.Destroy()
            # Destroy the Stream Object
            stream.close()

        # If we have a text, data, paragraph, paragraphlayout, or richtext end tag ...
        if name in [u'text', u'data', u'paragraph', u'paragraphlayout', u'richtext']:
            # ... we need to clear the element type, as we're no longer processing that type of element!
            self.element = None

    def hex2int(self, data):
        """ Image data is stored in a file-friendly Hex format.  We need to convert it to an image-friendly binary format. """
        # Initialize the conversion result variable
        result = ''
        # For each PAIR of characters in the hex data string ...
        for x in range(0, len(data), 2):
            # ... convert the hex pair into a integer, find that character, and add it to the result variable
            result += chr(int(data[x : x + 2], 16))
        # Return the converted data
        return result
