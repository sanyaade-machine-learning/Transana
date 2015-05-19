# Copyright (C) 2000  Donald N. Allingham
#
# Modified August 2002 by Gary Shao
#
#   Removed Gramps dependencies.
#
#   Made a more explicit distinction of whether paragraph properties
#   were being emitted in a table or not. Changing the location of
#   \qc and \qr commands in cell text enabled text alignment in table
#   cells to work properly.
#
#   Improved the appearance of output by adding some default spacing
#   before and after paragraphs using \sa and \sb commands. This lets
#   the output more closely mimic that of the other document generators
#   given the same document parameters.
#
#   Improved the appearance of tables by adding some default cell
#   padding using the \trgaph command. This causes the output to
#   more closely resemble that of the other document generators.
#
#   Removed unnecessary \par command at the end of embedded image.
#
#   Removed self-test function in favor of a general testing program
#   that can be run for all supported document generators. Simplifies
#   testing when the test program only has to be changed in one location
#   instead of in each document generator.
#
#   Modified open() and close() methods to allow the filename parameter
#   passed to open() to be either a string containing a file name, or
#   a Python file object. This allows the document generator to be more
#   easily used with its output directed to stdout, as may be called for
#   in a CGI script.
#
# Modified September 2002 by Gary Shao
#
#   Added start_listing() and end_listing() methods to allow showing
#   text blocks without filling or justifying.
#
#   Added line_break() method to allow forcing a line break in a text
#   block.
#
#   Added new methods start_italic() and end_italic() to enable
#   italicizing parts of text within a paragraph
#
#   Added method show_link() to display in text the value of a link.
#   This method really only has an active role in the HTML generator,
#   but is provided here for interface consistency.
#
#   Changed paragraph leader behavior to be like that of other generators.
#   Removed the use of a \tab entry in the leader segment that is
#   emitted.
#
#   Added a space after newline characters in write_text() text stream
#   since it appears that RTF simply concatenates the text from one
#   line to another in multi-line paragraphs.
#
#   Added ability to make leader bold text by having the 1st character
#   be a @ character. Seems like a simple way to convey emphasis without
#   changing the paragraph interface.
#
#   Added support for borderWidth and borderColor attributes of paragraph
#   style
#
# Modified 2004-2006 by Nate Case and David Woods
#
#   Made minor changes to make this compatible with Transana
#   and to fix Unicode support
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
#

#------------------------------------------------------------------------
#
# Load the base TextDoc class
#
#------------------------------------------------------------------------
from TextDoc import *

import TransanaConstants
import sys
import types


DEBUG = False
if DEBUG:
    print "RTFDoc DEBUG is True"


#try:
#    import Plugins
#    import intl
#    _ = intl.gettext
#except:
#    withGramps = 0
#else:
#    withGramps = 1

#------------------------------------------------------------------------
#
# RTF uses a unit called "twips" for its measurements. According to the 
# RTF specification, 1 point is 20 twips. This routines converts 
# centimeters to twips
#
# 2.54 cm/inch 72pts/inch, 20twips/pt
#
#------------------------------------------------------------------------
def twips(cm):
    return int(((cm/2.54)*72)+0.5)*20

#------------------------------------------------------------------------
#
# Rich Text Format Document interface. The current inteface does not
# use style sheets. Instead it writes raw formatting.
#
#------------------------------------------------------------------------
class RTFDoc(TextDoc):

    #--------------------------------------------------------------------
    #
    # Opens the file, and writes the header. Builds the color and font
    # tables.  Fonts are chosen using the MS TrueType fonts, since it
    # is assumed that if you are generating RTF, you are probably 
    # targeting Word.  This generator assumes a Western Europe character
    # set.
    #
    # NAC: extfonts is a list of font face strings, extcolors is a list of
    #      RGB tuples to be added to the color table
    #--------------------------------------------------------------------
    def open(self,filename,extfonts=[],extcolors=[]):
        #if type(filename) == type(""):
        if isinstance(filename, types.StringTypes):
            if filename[-4:] != ".rtf":
                self.filename = filename + ".rtf"
            else:
                self.filename = filename

            self.f = open(self.filename,"w")
            self.alreadyOpen = 0
        elif hasattr(filename, "write"):
            self.f = filename
            self.alreadyOpen = 1
        # Let's differentiate between Windows and Mac   DKW
        if sys.platform == 'darwin':  # Use python rather than wxPython to detect platform info
            self.f.write('{\\rtf1\\mac\\ansicpg1252\\deff0\n')
        else:
            self.f.write('{\\rtf1\\ansi\\ansicpg1252\\deff0\n')
            
        # Let's add a Default Font specifier, to try to deal with the Mac Font problem  DKW
        self.f.write('\stshfloch2 \stshfhich2\n')
        
        self.f.write('{\\fonttbl\n')
        self.f.write('{\\f0\\froman\\fcharset1\\fprq0 Times New Roman;}\n')
        self.f.write('{\\f1\\fswiss\\fcharset1\\fprq0 Arial;}\n')
        self.f.write('{\\f2\\fmodern\\fcharset1\\fprq1 Courier New;}\n')   # fprq1 indicates fixed pitch font  DKW
        
        # NAC: Write the extended font table -- Nate
        i = 3
        for face in extfonts:
            # Let's be a little more subtle than \\fswiss\\fcharset0\\fprq0  DKW
            if face == 'Symbol':
                family = 'ftech'  # ftech is Technical, Symbol. and, Mathematical font family
                charset = '2'     # Symbol character set
                prq = '0'         # Default font pitch
            elif face == 'Courier New':
                family = 'fmodern'  # fmodern is Courier New's font family
                charset = '1'       # Default character set
                prq = '1'           # Fixed font pitch
            elif face == 'Times New Roman':
                family = 'froman'  # fmodern is Courier New's font family
                charset = '1'       # Default character set
                prq = '0'           # Fixed font pitch
            else:
                family = 'fnil'   # fnil indicates unknown font family
                charset = '1'     # Default character set
                prq = '0'         # Default font pitch
            self.f.write("{\\f%d\\%s\\fcharset%s\\fprq%s %s;}\n" % (i,family,charset,prq,face))
            i = i + 1

        self.fontTable = ["Times New Roman", "Arial", "Courier New"] + extfonts
        
        self.f.write('}\n')

        self.f.write('{\colortbl\n')
        self.color_map = {}
        index = 2
        self.color_map[(0,0,0)] = 0
        self.f.write('\\red0\\green0\\blue0;')
        self.color_map[(0,0,255)] = 1
        self.f.write('\\red0\\green0\\blue255;')
        for style_name in self.style_list.keys():
            style = self.style_list[style_name]
            fgcolor = style.get_font().get_color()
            bgcolor = style.get_background_color()
            bordercolor = style.get_border_color()
            if not self.color_map.has_key(fgcolor):
                self.color_map[fgcolor] = index
                self.f.write('\\red%d\\green%d\\blue%d;' % fgcolor)
                index = index + 1
            if not self.color_map.has_key(bgcolor):
                self.f.write('\\red%d\\green%d\\blue%d;' % bgcolor)
                self.color_map[bgcolor] = index
                index = index + 1
            if not self.color_map.has_key(bordercolor):
                self.f.write('\\red%d\\green%d\\blue%d;' % bordercolor)
                self.color_map[bordercolor] = index
                index = index + 1

        # NAC: Write the extended color table -- Nate
        for col in extcolors:
            if not self.color_map.has_key(col):
                self.color_map[col] = index
                self.f.write("\\red%d\\green%d\\blue%d;" % col)
                index = index + 1
        
        self.f.write('}\n')
        self.f.write('\\kerning0\\cf0\\viewkind1')
        self.f.write('\\paperw%d' % twips(self.width))
        self.f.write('\\paperh%d' % twips(self.height))
        self.f.write('\\margl%d' % twips(self.lmargin))
        self.f.write('\\margr%d' % twips(self.rmargin))
        self.f.write('\\margt%d' % twips(self.tmargin))
        self.f.write('\\margb%d' % twips(self.bmargin))
        self.f.write('\\widowctl\n')
        self.in_table = 0
        self.in_listing = 0
        self.text = ""
        self.fontFace = ""
        self.fontSize = 12
        self.fgColor = (0,0,0)
        self.bgColor = (255,255,255)

    #--------------------------------------------------------------------
    #
    # Write the closing brace, and close the file.
    #
    #--------------------------------------------------------------------
    def close(self):
        self.f.write('}\n')
        if not self.alreadyOpen:
            self.f.close()

    #--------------------------------------------------------------------
    #
    # Force a section page break
    #
    #--------------------------------------------------------------------
    def end_page(self):
        self.f.write('\\sbkpage\n')

    #--------------------------------------------------------------------
    #
    # Starts a listing. Instead of using a style sheet, generate the
    # the style for each paragraph on the fly. Not the ideal, but it 
    # does work.
    #
    #--------------------------------------------------------------------
    def start_listing(self,style_name):
        self.opened = 0

        if DEBUG:
            print "self.start_listing opened = 0 (1)"
        
        p = self.style_list[style_name]

        # build font information

        f = p.get_font()
        size = f.get_size()*2
        bgindex = self.color_map[p.get_background_color()]
        fgindex = self.color_map[f.get_color()]
        if f.get_type_face() == FONT_MONOSPACE:
            self.font_type = '\\f2\\fs%d\\cf%d\\cb%d' % (size,fgindex,bgindex)
        elif f.get_type_face() == FONT_SERIF:
            self.font_type = '\\f0\\fs%d\\cf%d\\cb%d' % (size,fgindex,bgindex)
        else:
            self.font_type = '\\f1\\fs%d\\cf%d\\cb%d' % (size,fgindex,bgindex)
        if f.get_bold():
            self.font_type = self.font_type + "\\b"
        if f.get_underline():
            self.font_type = self.font_type + "\\ul"
        if f.get_italic():
            self.font_type = self.font_type + "\\i"

        # build listing block information

        self.f.write('\\pard')
        self.f.write('\\ql')
        self.para_align = '\\ql'
        self.f.write('\\nowidctlpar')
        self.f.write('\\ri%d' % twips(p.get_right_margin()))
        self.f.write('\\li%d' % twips(p.get_left_margin()))
        self.f.write('\\fi%d' % twips(p.get_first_indent()))
        if p.get_padding():
            self.f.write('\\sa%d' % twips((0.25 + p.get_padding())/2.0))
            self.f.write('\\sb%d' % twips((0.25 + p.get_padding())/2.0))
        else:
            self.f.write('\\sa%d' % twips(0.125))
            self.f.write('\\sb%d' % twips(0.125))
        if p.get_top_border():
            self.f.write('\\brdrt\\brdrs')
            if p.get_padding():
                self.f.write('\\brsp%d' % twips(p.get_padding()))
            else:
                self.f.write('\\brsp%d' % twips(0.125))
        if p.get_bottom_border():
            self.f.write('\\brdrb\\brdrs')
            if p.get_padding():
                self.f.write('\\brsp%d' % twips(p.get_padding()))
            else:
                self.f.write('\\brsp%d' % twips(0.125))
        if p.get_left_border():
            self.f.write('\\brdrl\\brdrs')
            if p.get_padding():
                self.f.write('\\brsp%d' % twips(p.get_padding()))
            else:
                self.f.write('\\brsp%d' % twips(0.125))
        if p.get_right_border():
            self.f.write('\\brdrr\\brdrs')
            if p.get_padding():
                self.f.write('\\brsp%d' % twips(p.get_padding()))
            else:
                self.f.write('\\brsp%d' % twips(0.125))
        self.in_listing = 1
 
    #--------------------------------------------------------------------
    #
    # Ends a listing. Care has to be taken to make sure that the 
    # braces are closed properly. The self.opened flag is used to indicate
    # if braces are currently open. If the last write was the end of 
    # a bold-faced phrase, braces may already be closed.
    #
    #--------------------------------------------------------------------
    def end_listing(self):
        self.f.write(self.text)
        if self.opened:
            self.f.write('}')
            self.opened = 0

            if DEBUG:
                print "self.end_listing() opened = 0 (2)"
        
        self.text = ""
        self.f.write('\n\\par')
        self.in_listing = 0

    #--------------------------------------------------------------------
    #
    # Starts a paragraph. Instead of using a style sheet, generate the
    # the style for each paragraph on the fly. Not the ideal, but it 
    # does work.
    #
    #--------------------------------------------------------------------
    def start_paragraph(self,style_name,leader=None):
        self.opened = 0

        if DEBUG:
            print "self.start_paragraph opened = 0 (3)"
        
        p = self.style_list[style_name]

        # build font information

        f = p.get_font()
        size = f.get_size()*2
        bgindex = self.color_map[p.get_background_color()]
        fgindex = self.color_map[f.get_color()]
        borderindex = self.color_map[p.get_border_color()]
        if f.get_type_face() == FONT_SERIF:
            self.font_type = '\\f0\\fs%d\\cf%d\\cb%d' % (size,fgindex,bgindex)
        else:
            self.font_type = '\\f1\\fs%d\\cf%d\\cb%d' % (size,fgindex,bgindex)
        if f.get_bold():
            self.font_type = self.font_type + "\\b"
        if f.get_underline():
            self.font_type = self.font_type + "\\ul"
        if f.get_italic():
            self.font_type = self.font_type + "\\i"

        # build paragraph information

        if not self.in_table:
            self.f.write('\\pard')
        if p.get_alignment() == PARA_ALIGN_RIGHT:
            if not self.in_table:
                self.f.write('\\qr')
            self.para_align = '\\qr'
        elif p.get_alignment() == PARA_ALIGN_CENTER:
            if not self.in_table:
                self.f.write('\\qc')
            self.para_align = '\\qc'
        else:
            self.para_align = '\\ql'
        self.f.write('\\ri%d' % twips(p.get_right_margin()))
        self.f.write('\\li%d' % twips(p.get_left_margin()))
        self.f.write('\\fi%d' % twips(p.get_first_indent()))
        if p.get_alignment() == PARA_ALIGN_JUSTIFY:
            self.f.write('\\qj')
        if self.in_table:
            self.f.write('\\trgaph80')
        else:
            if p.get_padding():
                self.f.write('\\sa%d' % twips((0.25 + p.get_padding())/2.0))
                self.f.write('\\sb%d' % twips((0.25 + p.get_padding())/2.0))
            else:
                self.f.write('\\sa%d' % twips(0.125))
                self.f.write('\\sb%d' % twips(0.125))
            borderwidth = p.get_border_width()
            if borderwidth > 3:
                widthMod = '\\brdrth'
            else:
                widthMod = '\\brdrs'
            if borderwidth > 7:
                borderwidth = 7
            widthMod = '\\brdrw%d' % (borderwidth * 20) + widthMod
            if p.get_top_border():
                self.f.write('\\brdrt%s' % widthMod)
                if p.get_padding():
                    self.f.write('\\brsp%d' % twips(p.get_padding()))
                else:
                    self.f.write('\\brsp%d' % twips(0.125))
                self.f.write('\\brdrcf%d' % borderindex)
            if p.get_bottom_border():
                self.f.write('\\brdrb%s' % widthMod)
                if p.get_padding():
                    self.f.write('\\brsp%d' % twips(p.get_padding()))
                else:
                    self.f.write('\\brsp%d' % twips(0.125))
                self.f.write('\\brdrcf%d' % borderindex)
            if p.get_left_border():
                self.f.write('\\brdrl%s' % widthMod)
                if p.get_padding():
                    self.f.write('\\brsp%d' % twips(p.get_padding()))
                else:
                    self.f.write('\\brsp%d' % twips(0.125))
                self.f.write('\\brdrcf%d' % borderindex)
            if p.get_right_border():
                self.f.write('\\brdrr%s' % widthMod)
                if p.get_padding():
                    self.f.write('\\brsp%d' % twips(p.get_padding()))
                else:
                    self.f.write('\\brsp%d' % twips(0.125))
                self.f.write('\\brdrcf%d' % borderindex)
        # This is redundant. Why was it here?
        #if p.get_first_indent():
        #    self.f.write('\\fi%d' % twips(p.get_first_indent()))
        #if p.get_left_margin():
        #    self.f.write('\\li%d' % twips(p.get_left_margin()))
        #if p.get_right_margin():
        #    self.f.write('\\ri%d' % twips(p.get_right_margin()))

        if leader:
            self.opened = 1

            if DEBUG:
                print "self.start_paragraph opened = 1 (1)"
        
            self.f.write('\\tx%d' % twips(p.get_left_margin()))
            if leader[0] == '@':
                leader = leader[1:]
                self.f.write('{%s\\b ' % self.font_type)
            else:
                self.f.write('{%s ' % self.font_type)
            self.f.write(leader)
            self.f.write('}')
            #self.f.write('\\tab}')
            self.opened = 0

            if DEBUG:
                print "self.start_paragraph opened = 0 (4)"
        

        self.bold_on = 0
        self.italic_on = 0
        self.underline_on = 0
    
    #--------------------------------------------------------------------
    #
    # Ends a paragraph. Care has to be taken to make sure that the 
    # braces are closed properly. The self.opened flag is used to indicate
    # if braces are currently open. If the last write was the end of 
    # a bold-faced phrase, braces may already be closed.
    #
    #--------------------------------------------------------------------
    def end_paragraph(self):
        if not self.in_table:
            self.f.write(self.text)
            if self.opened:
                self.f.write('}')
                self.opened = 0

                if DEBUG:
                    print "self.end_paragraph opened = 0 (5)"
        
            self.text = ""
            self.f.write('\n\\par')
        else:
            if self.text == "":
                self.write_text(" ")
            self.text = self.text + '}'

    #--------------------------------------------------------------------
    #
    # Starts italicized text, enclosed the braces
    #
    #--------------------------------------------------------------------
    def start_italic(self):
        self.italic_on = 1
        emph = self.get_emph()
        if self.opened:
            self.text = self.text + '}'
        self.text = self.text + '{%s%s ' % (self.font_type, emph)
        self.opened = 1

        if DEBUG:
            print "self.start_italic opened = 1 (2)"
        

    #--------------------------------------------------------------------
    #
    # Ends italicized text, closing the braces
    #
    #--------------------------------------------------------------------
    def end_italic(self):
        self.italic_on = 0
        if self.opened:
            self.text = self.text + '}'
        self.opened = 0

        if DEBUG:
            print "self.end_italic opened = 0 (6)"
        

    def get_emph(self):
        emph = ''
        if self.bold_on:
            emph = emph + '\\b'
        if self.italic_on:
            emph = emph + '\\i'
        if self.underline_on:
            emph = emph + '\\ul'
        return emph
        
    #--------------------------------------------------------------------
    #
    # Starts boldfaced text, enclosed the braces
    #
    #--------------------------------------------------------------------
    def start_bold(self):
        self.bold_on = 1
        emph = self.get_emph()
        if self.opened:
            self.text = self.text + '}'
        self.text = self.text + '{%s%s ' % (self.font_type, emph)
        self.opened = 1

        if DEBUG:
            print "self.start_bold opened = 1 (3)"
        

    #--------------------------------------------------------------------
    #
    # Ends boldfaced text, closing the braces
    #
    #--------------------------------------------------------------------
    def end_bold(self):
        self.bold_on = 0
        if self.opened:
            self.text = self.text + '}'
        self.opened = 0

        if DEBUG:
            print "self.end_bold opened = 0 (7)"
        

# FIXME: Resume here.. see above and the other under code.. use \ul ??
    #--------------------------------------------------------------------
    #
    # Starts underlined text, enclosed the braces
    #
    #--------------------------------------------------------------------
    def start_underline(self):
        self.underline_on = 1
        emph = self.get_emph()
        if self.opened:
            self.text = self.text + '}'
        self.text = self.text + '{%s%s ' % (self.font_type, emph)
        self.opened = 1

        if DEBUG:
            print "self.start_underline opened = 1 (4)"
        

    #--------------------------------------------------------------------
    #
    # Ends underlined text, closing the braces
    #
    #--------------------------------------------------------------------
    def end_underline(self):
        self.underline_on = 0
        if self.opened:
            self.text = self.text + '}'
        self.opened = 0

        if DEBUG:
            print "self.end_underline opened = 0 (8)"
        

    def set_ext_font(self, face, size):
        """Set to an extended font face/size."""
        # Should we deal with self.opened here?  Ignoring for now
        try:
            num = self.fontTable.index(face)
            # Space added to the end of this string to fix a formatting bug in Transana
            # where trailing spaces were disappearing if you changed a word's font.
            self.text = self.text + "\\f%d\\fs%d " % (num, size*2)
            self.fontFace = face
            self.fontSize = size
            #self.font_type = "\\f%d\\fs%d\\cf%d\\cb%d" % (num, size*2)
            #self.font_type = "\\f%d\\fs%d\\cf%d\\cb%d" % (num, size*2,0,2)
            self.font_type = "\\f%d\\fs%d" % (num, size*2)
        except:
            print "exception, invalid font spec'd?"
            return  # Invalid font specified, ignore call

    def set_ext_color(self, fgcol, bgcol):
        """Set to an extended font color.  fgcol and bgcol are RGB tuples"""
        try:
            fgindex = self.color_map[fgcol]
            bgindex = self.color_map[bgcol]
            # Space added to the end of this string to fix a formatting bug in Transana
            # where trailing spaces were disappearing if you changed a word's color.
            self.text = self.text + "\\cf%d\\cb%d " % (fgindex, bgindex)
            self.fgColor = fgcol
            self.bgColor = bgcol
            self.font_type = "\\f%d\\fs%d\\cf%d\\cb%d" % (self.fontTable.index(self.fontFace), self.fontSize*2, fgindex, bgindex)
        except:
            print "exception, invalid color spec'd?"
            return # Invalid color passed probably
           

    #--------------------------------------------------------------------
    #
    # Start a table. Grab the table style, and store it. Keep a flag to
    # indicate that we are in a table. This helps us deal with paragraphs
    # internal to a table. RTF does not require anything to start a 
    # table, since a table is treated as a bunch of rows.
    #
    #--------------------------------------------------------------------
    def start_table(self,name,style_name):
        self.in_table = 1
        self.tbl_style = self.table_styles[style_name]

    #--------------------------------------------------------------------
    #
    # End a table. Turn off the table flag
    #
    #--------------------------------------------------------------------
    def end_table(self):
        self.in_table = 0

    #--------------------------------------------------------------------
    #
    # Start a row. RTF uses the \trowd to start a row. RTF also specifies
    # all the cell data after it has specified the cell definitions for
    # the row. Therefore it is necessary to keep a list of cell contents
    # that is to be written after all the cells are defined.
    #
    #--------------------------------------------------------------------
    def start_row(self):
        self.contents = []
        self.cell = 0
        self.prev = 0
        self.cell_percent = 0.0
        self.f.write('\\trowd\n')

    #--------------------------------------------------------------------
    #
    # End a row. Write the cell contents, separated by the \cell marker,
    # then terminate the row
    #
    #--------------------------------------------------------------------
    def end_row(self):
        self.f.write('{')
        for line in self.contents:
            self.f.write(line)
            self.f.write('\\cell ')
        self.f.write('}\\pard\\intbl\\row\n')

    #--------------------------------------------------------------------
    #
    # Start a cell. Dump out the cell specifics, such as borders. Cell
    # widths are kind of interesting. RTF doesn't specify how wide a cell
    # is, but rather where it's right edge is in relationship to the 
    # left margin. This means that each cell is the cumlative of the 
    # previous cells plus its own width.
    #
    #--------------------------------------------------------------------
    def start_cell(self,style_name,span=1):
        s = self.cell_styles[style_name]
        self.remain = span -1
        if s.get_top_border():
            self.f.write('\\clbrdrt\\brdrs\\brdrw10\n')
        if s.get_bottom_border():
            self.f.write('\\clbrdrb\\brdrs\\brdrw10\n')
        if s.get_left_border():
            self.f.write('\\clbrdrl\\brdrs\\brdrw10\n')
        if s.get_right_border():
            self.f.write('\\clbrdrr\\brdrs\\brdrw10\n')
        table_width = float(self.get_usable_width())
        for cell in range(self.cell,self.cell+span):
            self.cell_percent = self.cell_percent + float(self.tbl_style.get_column_width(cell))
        cell_width = twips((table_width * self.cell_percent)/100.0)
        self.f.write('\\cellx%d\\pard\intbl\n' % cell_width)
        self.cell = self.cell+1

    #--------------------------------------------------------------------
    #
    # End a cell. Save the current text in the content lists, since data
    # must be saved until all cells are defined.
    #
    #--------------------------------------------------------------------
    def end_cell(self):
        self.contents.append(self.text)
        self.text = ""

    #--------------------------------------------------------------------
    #
    # Add a photo. Embed the photo in the document. Use the Python 
    # imaging library to load and scale the photo. The image is converted
    # to JPEG, since it is smaller, and supported by RTF. The data is
    # dumped as a string of HEX numbers.
    #
    #--------------------------------------------------------------------
    def add_photo(self,name,pos,x_cm,y_cm):
        # NOT SUPPORTED
        #        im = ImgManip.ImgManip(name)
        nx,ny = im.size()

        ratio = float(x_cm)*float(ny)/(float(y_cm)*float(nx))

        if ratio < 1:
            act_width = x_cm
            act_height = y_cm*ratio
        else:
            act_height = y_cm
            act_width = x_cm/ratio

        buf = im.jpg_scale_data(int(act_width*40),int(act_height*40))

        act_width = twips(act_width)
        act_height = twips(act_height)

        self.f.write('{\*\shppict{\\pict\\jpegblip')
        self.f.write('\\picwgoal%d\\pichgoal%d\n' % (act_width,act_height))
        index = 1
        for i in buf:
            self.f.write('%02x' % ord(i))
            if index%32==0:
                self.f.write('\n')
            index = index+1
        self.f.write('}}\n')
    
    #--------------------------------------------------------------------
    #
    # Writes text. If braces are not currently open, open them. Loop 
    # character by character (terribly inefficient, but it works). If a
    # character is 8 bit (>127), NO NO NO --> convert it to a hex representation in 
    # the form of \`XX. <-- NO NO NO  Let's write it to UTF-8 characters, not a
    # single-byte character.  This supports a wider variety of languages. DKW
    # Make sure to escape braces.
    # (escaping backslash added by DKW)
    #
    #--------------------------------------------------------------------
    def write_text(self,text):
        if self.opened == 0:
            emph = self.get_emph()
            self.opened = 1

            if DEBUG:
                print "self.write_text opened = 1 (5)"
        
            if self.in_table:
                self.text = self.text + self.para_align

            if DEBUG:
                print 'self.text = "%s"' % self.text
                
            self.text = self.text + '{%s%s ' % (self.font_type, emph)

            if DEBUG:
                print '{%s%s ' % (self.font_type, emph)
            
        for i in text:
            if ord(i) > 127:
                if type(i).__name__ == 'unicode':
                    if DEBUG:
                        print "RTFDoc.write_text():  ord(i) > 127:  ord(i) = %d, type(i) = %s" % (ord(i), type(i))
                        print i.encode('utf8')
                        print "RTF representation:  \\\'%2x   \u%04d" % (ord(i), ord(i))

                    if (sys.platform == 'darwin') and (i == unicode('\xc2\xa7', 'utf8')):

                        if DEBUG:
                            print "RTFDoc.write_text():  Mac Unicode Time Code Substitution"
                        
                        i = unicode('\xc2\xa4', 'utf8')

                    # Unicode characters need to be represented differently depending on their length.
                    # (at least in theory)
                    if ord(i) > 255:     # len(i.encode('utf8')) > 2:
                        self.text = self.text + "\\u%04d\\'3f" % ord(i)
                        if DEBUG:
                            print "Using \\u%04d\\'3f" % ord(i)
                    else:
                        self.text = self.text + '\\\'%2x' % ord(i)
                        if DEBUG:
                            print 'Using \\\'%2x' % ord(i)
                else:
                    self.text = self.text + '\\\'%2x' % ord(i)

            elif i == '{' or i == '}' or i == '\\':
                self.text = self.text + '\\%s' % i

                # print "RTFDoc.write_text() : Escaping %s character" % i
                
            elif self.in_listing and i == '\n':
                self.text = self.text + '\\line\n'
            elif i == '\n':
                self.text = self.text + '\n '
            else:
                self.text = self.text + i

    #--------------------------------------------------------------------
    #
    # Inserts a required line break into the text.
    #
    #--------------------------------------------------------------------
    def line_break(self):
        self.text = self.text + '\\line\n'

    #--------------------------------------------------------------------
    #
    # Shows link text.
    #
    #--------------------------------------------------------------------
    def show_link(self, text, href):
        #self.write_text("%s (" % text)
        #self.start_italic()
        #self.write_text(href)
        #self.end_italic()
        #self.write_text(") ")
        self.text = self.text + '{\\field{\\*\\fldinst{\ul\cf1  HYPERLINK "%s" }} {\\fldrslt{\ul\cf1 %s}}}' % (href, text)

#------------------------------------------------------------------------
#
# Register the document generator with the system if in Gramps
#
#------------------------------------------------------------------------
#if withGramps:
#    Plugins.register_text_doc(
#        name=_("Rich Text Format (RTF)"),
#        classref=RTFDoc,
#        table=1,
#        paper=1,
#        style=1
#        )
