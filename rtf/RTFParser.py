# Copyright (C) 2002-2014 The Board of Regents of the University of Wisconsin System
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

"""This module implements a primitive RTF parser."""

__author__ = 'Nathaniel Case, David K. Woods <dwoods@wcer.wisc.edu>, Jonathan Beavers <jonathan.beavers@gmail.com>'

# Okay, poor form, but I have to have the Global Encoding here.
# import Transana's Globals
import TransanaGlobal
import array
import wx
import exceptions
import string
import copy

DEBUG = False
if DEBUG:
    print "RTFParser DEBUG on"
DEBUG2 = False
if DEBUG2:
    print "RTFParser DEBUG2 on"


#
# The parser will convert a RTF document file into a list of DocObjects.
# Each DocObject in the list is either an attribute or a text object.
# Attribute objects set things like font, style, etc.  Text objects are
# strings of data in between the attribute changes.  An attribute object
# sets the "current" attribute and affects all text following it, presumably
# until another attribute object is encountered to change it.
#
# This is a very primitive/simple way of dealing with the RTF documents,
# and is probably not the best way to do things.  It ignores many features
# of RTF.
#

class RTFParseError(exceptions.Exception):
    """Raised when an error is encountered while parsing an RTF document."""
    def __init__(self, msg="Unspecified RTF parsing error"):
        self.explanation = msg
    

class DocAttribute:
    """Text/font attribute setting in a document."""

    def __init__(self):
        self.bold = 0
        self.italic = 0
        self.underline = 0
        # Default to Times New Roman.  This matches Word for Windows' assumption.
        self.font = "Times New Roman"
        # Default to 12 point.  This matches Word for Windows' assumption.
        self.fontsize = 12
        self.fg = 0
        self.bg = 0xffffff  # Default bg color = WHITE

    def __repr__(self):
        attrs = []
        if self.bold:
            attrs.append("Bold")
        if self.italic:
            attrs.append("Italic")
        if self.underline:
            attrs.append("Underline")
        
        if attrs == []:
            attrs = ["Normal"]
            
        return "'%s' '%s' %d [fg: #%06x, bg: #%06x]" % (string.join(attrs), self.font, self.fontsize, self.fg, self.bg)
        
class DocObject:
    """An object in a document stream.  An object is either an attribute
    or text."""
    
    def __init__(self):
        self.attr = None
        self.text = ""
   
    def __repr__(self):
        return "Attr: %s, Text: %s" % (self.attr, self.text)

class RTFParser:
    def __init__(self, fname=None):
        """Initialize a RTFParser object."""
        self.buf = ""
        self.index = 0
        self.stream = []
        self.nest = 0       # {} stack level
        self.curattr = DocAttribute()
        self.attr_stack = [self.curattr]     # document attribute stack
        self.update_func = None

        # Parsing state information
        self.in_font_table = 0
        self.in_font_block = 0
        self.in_color_table = 0
#        self.fonts = []     # Document font table

        # DKW There is a problem with fonts on import.  They are not getting assigned properly.  This is attempting to fix that.
        self.fontName = ""
        self.fontNumber = -1
        self.fontTable = {}
        
        self.colors = [0x000000]    # Document color table
        self.coloridx = 0
        self.new_para = 0       # Got \par command for new paragraph?

        # Read in buffer if file given
        if fname != None:
            f = open(fname, "r")
            self.buf = f.read()
            f.close()

    def init_progress_update(self, update_func, start_val, end_val):
        """If called before parsing, will set up the RTF parser class to
        call an update function while parsing, for use in progress reporting.
        start and end values should reflect the start and end percentage
        desired for the parsing."""
        self.update_func = update_func
        self.update_start = float(start_val)
        self.update_end = float(end_val)
        self.update_value = self.update_start
        self.update_func(self.update_value)
        
    def check_header(self):
        if not self.buf.startswith("{\\rtf1"):
            if DEBUG:
                print "RTFParser.check_header(): This is not an RTF Doc!!"
            raise RTFParseError, "Not a RTF document"

    def seek_eob(self):
        """Seek to the end of the current block.  Assume index after
        the start of the block.  End up on character after '}'"""
        desired_nest = self.nest - 1
        x = self.index
        l = len(self.buf)
        while x < l:
            if self.buf[x] == "\\":
                x = x + 1
            elif self.buf[x] == "{":
                self.nest = self.nest + 1
            elif self.buf[x] == "}":
                self.nest = self.nest - 1
                if self.nest == desired_nest:
                    break
            x = x + 1
            
        self.index = x + 1

    def seek_next_block(self):
        """Seek to the next { char."""
        i = self.buf[self.index:].find("{")
        self.index = self.index + i
        self.nest = self.nest + 1

    def seek_next_token(self):
        """Seek to the next token."""
        inwhitespace = 0
        while self.index < len(self.buf):
            if string.whitespace.count(self.buf[self.index]) > 0:
                inwhitespace = 1
            else:
                if inwhitespace:
                    break
            self.index += 1

    def process_new_block(self):
        """Index should point to '{'.  Pushes attribute settings onto the
        stack.  Returns with index pointing to next character after the '{'."""
        self.nest += 1
        self.index += 1
        do = DocObject()
        self.attr_stack.append(copy.deepcopy(self.curattr))
        self.curattr = DocAttribute()
        do.attr = self.curattr
        self.stream.append(do)
        if DEBUG:
            print "Begin block, Appending docstream object: '%s'" % str(do)
        
    def process_end_block(self):
        """Index should point to '}'.  Pops the attribute settings for the
        upper level block back from the stack. Returns with index pointing to
        next character after the '}'."""
        self.nest -= 1
        self.index += 1
        try:
            self.curattr = self.attr_stack.pop()
        except:
            return    # Flaky RTF file with extra unmatched } ?
        
        # Color table definition ends at end of block 
        if self.in_color_table:
            self.in_color_table = 0
        
        do = DocObject()
        do.attr = self.curattr
        self.stream.append(do)
        if DEBUG:
            print "End block, Appending docstream object: '%s'" % str(do)
   
    def process_text(self, txt):
	"""Process a plain text string."""
	if len(txt) > 0:
	    # Got \par control word so next text is on new line
	    if self.new_para:
		# There might have been more than one \par so we insert
		# newlines for however many we counted
		txt = ("\n" * self.new_para) + txt
		self.new_para = 0

	    # Since timecode, and up/down intonation symbols are different between
	    # OSX and windows, we need to perform substitution while loading a document.
	    if ('unicode' in wx.PlatformInfo) and ('wxMac' in wx.PlatformInfo):

                if DEBUG:
                    print "RTFParser.process_text(): '%s' .. " % txt,
                    for x in txt:
                        print ord(x),
                    print
                
		# This (array.array('B', txt).tolist()) is supposed to be the most efficient way 
		# of converting a string to a list of integers.  JB
		# Unfortunately, it requires a string, not a Unicode object.  We'll need a less
		# efficient alternative for Unicode.  DKW

                # Check the type of the txt object.  If it's a string ...
		if isinstance(txt, str):
                    # ... convert it efficiently to a list of integers
                    intList = array.array('B', txt).tolist()
                # If txt is Unicode ...
                else:
                    # ... create an empty list ...
                    intList = []
                    # ... iterate through the characters in the Unicode string ...
                    for x in txt:
                        # ... and add the integer value for each character to the integer list
                        intList.append(ord(x))
		
		try:
		    # all the windows special characters begin with \xc2 (#194)
		    position = intList.index(194)
		    
		    # test for timecode
# It appears that with wxPython 2.8.0.1, the Mac can now handle the proper timecode character!
#		    if intList[position+1] == 164:
			
#			if DEBUG:
#			    print "RTFParser.process_text():  Mac Unicode Substitution - Time Code."

#			newString = txt[0:position] + unicode('\xc2\xa7', 'utf8')
#			txt = newString
		    # Test for up intonation
#		    elif intList[position+1] == 173:
		    if intList[position+1] == 173:
			if DEBUG:
			    print "RTFParser.process_text(): Mac Unicode Substitution - Up Arrow."

			newString = txt[0:position] + unicode('\xe2\x89\xa0', 'utf8')
			txt = newString
		    # Test for down intonation
		    elif intList[position+1] == 175:
			
			if DEBUG:
			    print "RTFParser.process_text(): Mac Unicode Substitution - Down Arrow."

			newString = txt[0:position] + unicode('\xc3\x98', 'utf8')
			txt = newString
		    # Test for closed dot (hi dot)  (194 149 is 1.24 encoding, 194 183 is 2.05 encoding)
		    # but the 2.05 encoding doesn't work here because we don't want the text to be in Symbol font.
		    # That is handled in RichTextEditCtrl.py's __ParseRTFText() method.
		    elif (intList[position+1] == 149):
			
			if DEBUG:
			    print "RTFParser.process_text(): Mac Unicode Substitution - closed dot."

			newString = unicode('\xe2\x80\xa2', 'utf8')  # txt[0:position] + unicode('\xe2\x80\xa2', 'utf8')
			txt = newString
		    # DON'T Test for open dot (whisper)  (194 176 is 2.05 encoding)
		    # but the 2.05 encoding doesn't work here because we don't want the text to be in Symbol font.
		    # That is handled in RichTextEditCtrl.py's __ParseRTFText() method.

		except ValueError:
		    i = 1
	
	    do = DocObject()
            do.text = txt

            if DEBUG and False:
                print "RTFParser.process_text()", txt, do.text, do
                
            self.stream.append(do)
        
    def process_doc(self):
        """Process the document buffer."""

        # We process plain text in chunks, and append them each time
        # we encounter a non-text token in the document.  Current
        # plain text chunk goes into this 'txt' variable.
        txt = ""
        if self.update_func:
            range_len = float(self.update_end - self.update_start)
        
        while self.index < len(self.buf):
            c = self.buf[self.index]

            if DEBUG:
                print "RTFParser.process_doc() processing:", c
                
            if "{}\\".count(c) > 0:

                # DKW  We need to process "escaped" brace characters and the backslash character
                if self.buf[self.index : self.index+2] in ['\{', '\}', '\\\\']:

                    if DEBUG:
                        print "Processing an escaped character : '%s'" % txt

                    # Add the escaped character to the text buffer, making it text rather than a control character
                    txt += self.buf[self.index+1]
                    # Increment the index to point to after the escaped character
                    self.index += 2

                # DKW We also need to process unicode characters
                elif (self.buf[self.index : self.index+2] in ["\\'", "\\\'"]) and ('unicode' in wx.PlatformInfo):

                    if DEBUG:
                        print "RTFParser.process_doc(): Processing a Unicode Character",

                    val = int(self.buf[self.index+2:self.index+4], 16)
                    
                    # Word Smart Quotes could cause problem.  These come across as "\'93" and "\'94"
                    #(hex for 147 and 148) and need to be replaced with a normal quote character.
                    if (val in [147, 148]):

                        if DEBUG:
                            print "Smart Quotes detected:  before:", self.buf[self.index - 10:self.index + 10],
                            
                        # 22 is the HEX value for chr 34, the quote character.
                        val = 34
                        # We need to replace them in the self.buf text.  
                        self.buf = self.buf[:self.index+2] + '22' + self.buf[self.index+4:]

                        if DEBUG:
                            print "after:", self.buf[self.index - 10:self.index + 10]

# It appears with wxPython 2.8.0.1 that the time code translation is no longer needed!
#                    elif (val == 164) and ('wxMac' in wx.PlatformInfo):

#                        if DEBUG:
#                            print "OLD STYLE TIME CODE detected.  before:", self.buf[self.index - 10:self.index + 10],
                            
                        # We need to replace them in the self.buf text.  
#                        self.buf = self.buf[:self.index+2] + 'a7' + self.buf[self.index+4:]
                        
#                        if DEBUG:
#                            print "OLD STYLE TIME CODE detected.  after:", self.buf[self.index - 10:self.index + 10],

                    s = unicode(chr(val), 'latin1')
                    # s = u'%s' % self.buf[self.index:self.index+4]

                    if DEBUG:
                        print "%d" % (val)

                    # If there's text in the buffer ...
                    if txt != '':
                        # ... process that text first
                        self.process_text(txt)

                    # replace the text in the buffer
                    txt = s.encode(TransanaGlobal.encoding)

                    if DEBUG:
                        print "Calling process_control_word()"
                        
                    # process the unicode character as a Control Word
                    self.process_control_word()

                    if DEBUG:
                        print "AFTER process_control_word call"

                    # Clear the text buffer
                    txt = ''

                # If we're not dealing with an escaped character, continue normal processing
                else:
                    
                    # If non-text token encountered
                    if self.update_func:
                        self.update_func(self.update_start + \
                                range_len * self.index / len(self.buf) + 1)
                   
                    if self.in_font_block and len(txt) > 0:
                        # Add new font face to font table
    #                    self.fonts.append(txt[:-1])

                        if DEBUG:
                            print "Adding %s to FontTable as %s" % (txt[:-1], self.fontNumber)

                        if txt[:-1] != '':
                            self.fontName = txt[:-1]
                        elif DEBUG:
                            print "Actually, the font was named", self.fontName
                        self.fontTable[self.fontNumber] = self.fontName
                    else:

                        if DEBUG and (txt.strip() != ''):
                            print "Processing text '%s'" % txt

                        self.process_text(txt)

                        if DEBUG:
                            print "Done processing '%s'  :  c = '%s'" % (txt, c)
                            
                    # Word is doing something weird with Batang on Windows, where it has a "{\*\falt" spec
                    if self.in_font_table and self.in_font_block and txt != '':
                        self.fontName = txt
                        if DEBUG:
                            print "This is that weird font thing", self.fontName, "***************************************"
                    txt = ""
                    if c == '{':
                        if (self.in_font_table):
                            self.in_font_block = 1
                            self.index += 1 # Point to next control word
                        else:
                            self.process_new_block()
                    elif c == '}':
                        if (self.in_font_block):
                            self.index += 1 # Point to next font def
                            self.in_font_block = 0
                        elif (self.in_font_table):
                            self.in_font_table = 0
                            self.process_end_block()
                        else:
                            self.process_end_block()
                    elif c == '\\':

                        if DEBUG:
                            print "Processing Control Word"
                            
                        self.process_control_word()

                        if DEBUG:
                            print "Done Processing Control Word"
                        
            else:
                if (c == ";") and (self.in_color_table):
                    self.coloridx += 1
                else:
                    # Certain ASCII characters aren't meant literally, and
                    # should be inserted in the document through control
                    # words rather than through the ASCII characters.
                    if (c != '\r' and c != '\n'):
                        txt = txt + c

                self.index += 1
                
        if self.update_func:
                    self.update_func(self.update_end)

    def process_control_word(self):
        """Assume index points to \ character.  Will point to first
        character that isn't part of the control word when returned."""
        if self.buf[self.index] != "\\":
            raise RTFParseError, "Expected \\ (programming error?)"
  
        try:
            cw = ""
            self.index += 1
            c = self.buf[self.index]
            # Extract control word
            while string.ascii_letters.count(c) > 0:
                cw = cw + c
                self.index += 1
                c = self.buf[self.index]

            # Now index is at the first 'delimiter' character (c)
            if DEBUG:
                print "Found control word '%s' (first delim is %s)  " % (cw, c),

            # Check for numeric parameters
            if (string.digits.count(c) > 0 or (c == '-')):
                # Control word contains a numeric parameter
                numstr = c
                self.index += 1
                c = self.buf[self.index]
                while string.digits.count(c) > 0:
                    numstr = numstr + c
                    self.index += 1
                    c = self.buf[self.index]
                # Now index is at first non-digit character (c)
                try:
                    num = int(numstr)
                except:
                    num = 0
                if DEBUG:
                    print "Parameter is %d, first non-digit char is %s" % (num, c)
            else:
                num = None
            
            if (c == ' '):
                # Spaces are considered to be grouped with the control word
                # and not real text.
                self.index += 1

            if (c == '*' and cw == ''):
                # This is a special case I haven't completely figured out yet.
                # For now we assume that any block that has a \* in it
                # is a block containing user property definitions that we
                # don't care about, so we jump to the end of the block.
                self.seek_eob()
                return
            
            if (c == "'" and cw == ''):
                # \'hh is a hex value for the current charset
                self.index += 1
                try:
                    value = int(self.buf[self.index:self.index+2], 16)

                    if DEBUG:
                        print "RTFParser.process_control_word()", self.buf[self.index:self.index+2], "value =", value

                    if ('unicode' in wx.PlatformInfo):
                        # Try a straight conversion to UTF-8, the DefaultPyEncoding
                        try:
                            # CHANGED for 2.30 because 201 is now the back-accented capital E character!
                            if value == 133:  # 201
                                
                                if DEBUG:
                                    print "Elipsis substitution"
                                    
                                self.process_text('...')
                            else:
                                tempChar = unichr(value)
                                # I'm not sure why passing through latin-1 is needed, but it appears to be necessary.
                                # tempChar = unicode(chr(value), 'latin-1')
                                self.process_text(tempChar.encode(TransanaGlobal.encoding))
                            
                        except UnicodeEncodeError:
                            # If we get a UnicodeEncodeError, as we do for Time Codes, let's try going through
                            # Latin-1 encoding to translate the single-byte character into a 2-byte character.
                            # The default encoding for RTF is Transana at least if ansicpg1252, which I think 
                            # is equivalent to Latin-1
                            tempc = unicode(chr(value), 'latin1')

                            if DEBUG:
                                print "Passing through Latin-1"

                            self.process_text(tempc.encode(TransanaGlobal.encoding))

                            if DEBUG:
                                print "Should now be encoded as %s" % TransanaGlobal.encoding
                                
                    else:
                        self.process_text(chr(value))

                finally:
                    self.index += 2
                return 

            if cw == 'u':   # Unicode Character Processing added by DKW

                if DEBUG and num not in [164, 8232]:
                    print "Processing Unicode Character Code %d" % num

                try:
                    tempChar = unichr(num)
                    
                    if num == 8232:
                        
                        if DEBUG:
                            print "\line substitution"
                            
                        self.process_text("\n")
#                        self.index += 6   # self.buf.find(' ', self.index)  # Skip past the unicode character digits
                    elif num == 164:
                        
                        if DEBUG:
                            print "Time Code Substitution"
                            
                        if 'wxMac' in wx.PlatformInfo:
                            val = 167
                        else:
                            val = 164
                        tempChar = unichr(val)
                        self.process_text(tempChar.encode(TransanaGlobal.encoding))
#                        self.index += 6   # self.buf.find(' ', self.index)  # Skip past the unicode character digits
                    else:
                        # We don't use the global encoding, but UTF-8 here, as Python, and therefore the wxSTC, are using
                        # UTF-8 regardless of what the database is using.
                        self.process_text(tempChar)  # .encode('utf8')  (TransanaGlobal.encoding)
                        self.index += 4   # self.buf.find(' ', self.index)  # Skip past the unicode character digits

                    if DEBUG and num not in [164, 8232]:
                        print "Now we need to deal with the alternate character specification, usually \'f3"

                        print self.index, self.buf[self.index:self.index + 20], self.buf.find(' ', self.index)
                        
                except ValueError:
                    if DEBUG:
                        print "ValueError in RTF Processing for Unicode.  Control Word 'u', num =", num
                    pass
            
            elif cw == "fonttbl":
                self.in_font_table = 1
            elif cw == "rtf":
                if DEBUG:
                    print "Document uses RTF version %d" % num
                pass
            elif cw == "f":
                # Set current font, num = font table index number
                do = DocObject()
                do.attr = copy.deepcopy(self.curattr)
                # DKW If the font number is not in the fontTable, we need to add it.  (This should only
                #     occur when the Font Table is being read.)  We don't have all the data yet, such as
                #     the font name.  So let's just remember the font number that we just found out for
                #     now.
                if not self.fontTable.has_key(num):
                    if DEBUG:
                        print "Font %d needs to be added to the Font Table" % num
                    self.fontNumber = num
                else:
                    if DEBUG:
                        print "Font set to %d, %s" % (num, self.fontTable[num])
                    do.attr.font = self.fontTable[num]
                    self.stream.append(do)
                    self.curattr = copy.deepcopy(do.attr)

            elif cw in ["froman", "fcharset", "fprq", "fswiss", "fmodern"]:
                if DEBUG:
                    print "Ignoring Control Word '%s'" % cw
                else:
                    pass

            elif cw == "fs":
                # Set current font size
                do = DocObject()
                do.attr = copy.deepcopy(self.curattr)
                if DEBUG:
                    print "Font Size set to %3.1f" % (num / 2)
                do.attr.fontsize = num / 2
                self.stream.append(do)
                self.curattr = copy.deepcopy(do.attr)
            elif cw == "cf":
                # Set the foreground color
                do = DocObject()
                do.attr = copy.deepcopy(self.curattr)
                #print "cf #%d (0x%08x)" % (num, self.colors[num])
                do.attr.fg = self.colors[num]
                self.stream.append(do)
                self.curattr = copy.deepcopy(do.attr)
            elif cw == "cb":
                # Set the background color
                do = DocObject()
                do.attr = copy.deepcopy(self.curattr)
                #print "cb #%d (0x%08x)" % (num, self.colors[num])
                do.attr.bg = self.colors[num]
                self.stream.append(do)
                self.curattr = copy.deepcopy(do.attr)
            elif cw == "info":
                # This block contains metadata such as author/title, ignore it
                self.seek_eob()
            elif cw == "line":
                self.process_text("\n")
            elif cw == "tab":
                self.process_text("\t")
            elif cw == "lquote":
                self.process_text("`")
            elif cw == "rquote":
                self.process_text("'")
            elif cw == "ldblquote" or cw == "rdblquote":
                self.process_text('"')
            elif cw == "par":
                if DEBUG:
                    print "(New paragraph, process newline for next text)"
                self.new_para += 1
            elif cw == "colortbl":
                if DEBUG:
                    print "Processing color table.."
                self.colors = [0x00000]     # Default entry 0
                self.in_color_table = 1
            elif cw == "red":
                # Check to see if entry needs to be added
                if self.coloridx == len(self.colors):
                    self.colors.append(num << 16)
                else:
                    self.colors[self.coloridx] = num << 16
            elif cw == "green":
                self.colors[self.coloridx] |= (num << 8)
            elif cw == "blue":
                self.colors[self.coloridx] |= num
            elif cw == "stylesheet":
                # Ignore stylesheet data as it messes things up if not properly
                # supported
                self.seek_eob()
            elif cw == "ulnone":
                do = DocObject()
                do.attr = copy.deepcopy(self.curattr)
                do.attr.underline = 0
                self.stream.append(do)
                self.curattr = copy.deepcopy(do.attr)
            elif cw == "ul" or cw == "b" or cw == "i" or cw == "plain":
                # Underline/Bold/Italic/Plain
                do = DocObject()
                do.attr = copy.deepcopy(self.curattr)
                
                if num:
                    val = num != 0
                else:
                    # If no parameter passed, assume to turn it on
                    val = 1
                    
                if cw == "ul":
                    do.attr.underline = val
                elif cw == "b":
                    do.attr.bold = val
                elif cw == "i":
                    do.attr.italic = val
                elif cw == "plain":
                    do.attr.bold = do.attr.italic = do.attr.underline = 0

                if DEBUG:
                    print "Attr change, Appending docstream object: '%s'" % str(do)
       
                self.stream.append(do)
                self.curattr = copy.deepcopy(do.attr)
            # Sometimes, the closed dot is encodes as a "bullet" in RTF.
            elif cw.lower() == 'bullet':
                # We're using Unicode Character 183
                tempChar = unichr(183)
                # And we need to process it at Text
                self.process_text(tempChar.encode(TransanaGlobal.encoding))
            else:
                if DEBUG:
                    numstr = ''
                    if num:
                        numstr = '(' + str(num) + ')'
                    print "Ignoring unknown control word %s%s" % (cw, numstr)
                    
        except IndexError:
            if DEBUG:
                print "Caught IndexError exception (aborting control word)"
                import sys, traceback
                print "Exception %s: %s" % (sys.exc_info()[0], sys.exc_info()[1])
                traceback.print_exc(file=sys.stdout)
                # Display the Exception Message, allow "continue" flag to remain true
            return
    
    def read_stream(self):
        # Verify that it's a valid RTF
        self.check_header()
        self.process_doc()

        if DEBUG:
            print
            print "Fonts:"
            keys = self.fontTable.keys()
            keys.sort()
            for key in keys:
                print key, self.fontTable[key]
            print
            #print "Color table = %s" % self.colors
