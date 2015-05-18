# Copyright (C) 2003 - 2005 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the RichTextEditCtrl class, a Rich Text editor
based on StyledTextCtrl."""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'

import TransanaGlobal
import sys, os, string, re

import wx
from wx import stc

# RTF Module Path - This may not be necessary in the future but shouldn't
# harm anything.
#RTFModulePath = "..%s..%s..%srtf" % (os.sep, os.sep, os.sep)
RTFModulePath = "rtf"
if sys.path.count(RTFModulePath) == 0:
    sys.path.append(RTFModulePath)  # Add to path if not already there

import RTFParser
import RTFDoc, TextDoc
import cStringIO
import traceback

DEFAULT_TAB_SIZE = 4

# This mask determines the number of style bits to use.  The default is
# 31, which uses 5 bits (2^5) for 32 possible styles.  We increase this
# to 128 possible styles since we don't use the indicators feature that
# is normally used for the other 3 bits.  We have to use at least
# one indicator so we use 0x7f instead of 0xff.
STYLE_MASK = 0x7f

DEBUG = False
if DEBUG:
    print "RichTextEditCtrl DEBUG is ON."

class RichTextEditCtrl(stc.StyledTextCtrl):
    """This class is a Rich Text edit control."""

    def __init__(self, parent, id=-1):
        """Initialize an StyledTextCtrl object."""
        stc.StyledTextCtrl.__init__(self, parent, id)
        
        # The text display on the Mac is simply TOO SMALL.  This is an attempt to make it a bit bigger.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetZoom(3)
        # We need as many styles as possible, and will not use Indicators.
        self.SetStyleBits(7)
       
        # STYLE_HIDDEN is meant for things that are sometimes hidden and
        # sometimes visible.  STYLE_HIDDEN_ALWAYS is intended for things that
        # should never be visible to the user.

        # Any style assigned here needs to be added to __ResetBuffer() as well.
        self.STYLE_HIDDEN = -1
        self.STYLE_TIMECODE = -1
        self.STYLE_HIDDEN_ALWAYS = -1
        self.HIDDEN_CHARS = []
        self.HIDDEN_REGEXPS = []
        self.__ResetBuffer()

        # Set the Default Style
        self.StyleSetSpec(stc.STC_STYLE_DEFAULT, "size:" + str(TransanaGlobal.configData.defaultFontSize) + ",face:" + TransanaGlobal.configData.defaultFontFace)
        self.StyleSetSpec(stc.STC_STYLE_LINENUMBER, "size:10,face:" + TransanaGlobal.configData.defaultFontFace)
        self.SetWrapMode(stc.STC_WRAP_WORD)
        # Let's set the default Tab Width to 4 instead of 8 to better match Transana 1.24
        self.SetTabWidth(DEFAULT_TAB_SIZE)

        self.SetMarginType(0, stc.STC_MARGIN_NUMBER)
        if "__WXMAC__" in wx.PlatformInfo:
            lineNumberWidth = self.TextWidth(stc.STC_STYLE_LINENUMBER, '88888 ')
        else:
            lineNumberWidth = self.TextWidth(stc.STC_STYLE_LINENUMBER, '8888 ')

        self.SetMarginWidth(0, lineNumberWidth)
        self.SetMargins(6, 6)

        # Set the Selection Colors to white text on a blue background
        self.SetSelForeground(1, "white")
        self.SetSelBackground(1, "blue")

        stc.EVT_STC_CHARADDED(self, id, self.OnCharAdded)
        #stc.EVT_STC_POSCHANGED(self, id, self.OnPosChanged)
        stc.EVT_STC_UPDATEUI(self, id, self.OnPosChanged)
        stc.EVT_STC_MODIFIED(self, id, self.OnModified)
        
        # Remove Drag-and-Drop reference on the mac due to the Quicktime Drag-Drop bug
        if "__WXMAC__" in wx.PlatformInfo:
            stc.EVT_STC_START_DRAG(self, id, self.OnKillDrag)

        self.stylechange = 0
        
        # If set to point to a wxProgressDialog, this will be updated as
        # a document is loaded.
        self.ProgressDlg = None

        # There was a problem with the font being changed inappropriately during
        # data loading.  This helps prevent that.
        self.loadingData = False

# New Interfacesys.path.append("..%s..%s..%srtf" % (os.sep, os.sep, os.sep))

    # Remove Drag-and-Drop reference on the mac due to the Quicktime Drag-Drop bug
    if "__WXMAC__" in wx.PlatformInfo:
        def OnKillDrag(self, event):
            """ This method KILLS Drag-and-Drop functionality on the Mac. """
            # According to Robin Dunn, the way to kill Drag-and-Drop in the wxSTC is to blank out the Drag Text
            # He notes that this is an undocumented "hidden feature" of the START_DRAG event
            event.SetDragText("")

    def __GetStyle(self, spec):
        """Get a style number from the given StyleSpec string and
        return the Style index number.  If such a style does not exist
        then it will be created and a new number is returned.  This is
        limited in that you have to specify the spec string exactly
        how it was originally specified or else it won't recognize it as
        being the same (even if it is functionally the same)."""
        # FIXME: We should do a split() on it and see if all the elements
        # are compared rather than doing a string comparison.  With this way
        # it will waste some styles.

        if self.style_specs.count(spec) == 0:
            # Style doesn't exist yet
            if DEBUG:
                print "Allocating new style (spec %s)" % spec
            if (self.num_styles >= STYLE_MASK):
                if DEBUG:
                    print "No more styles left.  Using Default"
                # No more styles left, must use default
                return stc.STC_STYLE_DEFAULT

            # Okay, the stc.STC_STYLE_LINENUMBER style is getting over-written, which changes the style of the line numbers.
            # We have to do a trick here to avoid that.  HACK HACK HACK.
            # essentially, what this code does is insert the appropriate style when we get to this point in the
            # style array.  Apparently, the information is stored internally in the wxSTC until this 33rd style is created.
            if self.num_styles == stc.STC_STYLE_LINENUMBER:
                if DEBUG:
                    print "Skipping %d because it equals %d (stc.STC_STYLE_LINENUMBER)" % (self.num_styles, stc.STC_STYLE_LINENUMBER)
                self.StyleSetSpec(self.num_styles, "size:10,face:" + TransanaGlobal.configData.defaultFontFace+',fore:#000000')
                self.style_specs.append("size:10,face:" + TransanaGlobal.configData.defaultFontFace+',fore:#000000')
                attr = StyleSettings()
                attr.font_face = TransanaGlobal.configData.defaultFontFace
                attr.font_size = 10
                attr.font_fg = 0x000000
                self.style_attrs.append(attr)
                self.num_styles += 1

            self.StyleSetSpec(self.num_styles, spec)
            
            # Hidden text, when shown, is Red by default
            if spec == "hidden":
                self.StyleSetForeground(self.num_styles, wx.NamedColour("red"))
                
            self.style_specs.append(spec)
            
            attr = StyleSettings()
            props = spec.split(",")
            attr.bold = props.count("bold") > 0
            attr.italic = props.count("italic") > 0
            attr.underline = props.count("underline") > 0
            for x in props:
                if x[:5] == "face:":
                    attr.font_face = x[5:]
                if x[:5] == "fore:":
                    attr.font_fg = int(x[6:],16)
                if x[:5] == "back:":
                    attr.font_bg = int(x[6:],16)
                if x[:5] == "size:":
                    attr.font_size = int(x[5:])
            self.style_attrs.append(attr)
            
            self.num_styles += 1
            if DEBUG:
                print "New style #%d = %s" % (self.num_styles - 1, spec)
            return self.num_styles - 1
        else:
            return self.style_specs.index(spec) 

    def __CurrentSpec(self):
        """Return the current style spec string."""
        # We could have events that update self.style each time
        # the cursor is moved or text is selected rather than
        # "polling" it like this, could perform better.
        startpos = self.GetSelectionStart()
        endpos = self.GetSelectionEnd()
        if startpos == endpos:
            # No selection
            style = self.GetStyleAt(self.GetCurrentPos())
        else:
            style = self.GetStyleAt(startpos)

        return self.style_specs[style]
    
    def __StyledText(self, text):
        """Adds the style bytes to text using the current style."""
        stylebyte = self.style << 3
        data = ""
        for c in text:
            data = data + c + chr(stylebyte)
        return data
        
    def __ParseRTFStream(self, stream):
        
        # Assume that this stage of parsing takes up 50% of progress dialog,
        # so each element in the stream counts for 50/len(stream).
        prog_level = 50.0
        if len(stream) > 0:
            inc_level = 50.0/len(stream)
        else:
            inc_level = 0
        if self.ProgressDlg:
                self.ProgressDlg.Update(prog_level, _("Processing RTF stream"))

        for obj in stream:
            if self.ProgressDlg:
                self.ProgressDlg.Update(prog_level)
            if obj.attr == None:
                #self.AddText(obj.text)
                #self.AddStyledText(self.__StyledText(obj.text))
                startpos = self.GetCurrentPos()
                self.StartStyling(startpos, STYLE_MASK)
                self.AddText(obj.text)

                if DEBUG:
                    print "Adding text '%s' with style %s" % (obj.text, self.style)
                    
                self.SetStyling(len(obj.text), self.style)

                # Do a second pass through the text to check for hidden
                # characters.
                #i = startpos
                #for c in obj.text:
                #    if self.HIDDEN_CHARS.count(c):
                #        self.StartStyling(i, STYLE_MASK)
                #        self.SetStyling(1, self.STYLE_HIDDEN)
                #    i += 1
                                
                if DEBUG:
                    #print "Adding text len %d" % len(obj.text)
                    print "***TEXT =", obj.text
            else:
                # Everytime there's an attribute change, we change the
                # 'current' attribute completely to the new one.
                if DEBUG:
                    print "Attribute change: current style=%d" % self.style
                    print "Desired attribute: %s" % obj.attr
                self.__SetAttr("bold", obj.attr.bold)
                self.__SetAttr("underline", obj.attr.underline)
                self.__SetAttr("italic", obj.attr.italic)
                #self.SetFont(obj.attr.font, obj.attr.fontsize, obj.attr.fg, obj.attr.bg)
                self.__SetFontFace(obj.attr.font)
                self.__SetFontColor(obj.attr.fg, obj.attr.bg)
                self.__SetFontSize(obj.attr.fontsize)

                if DEBUG:
                    print "New style after attribute change: %d" % self.style
            prog_level = prog_level + inc_level
    
        # Do a second pass through the text to check for hidden
        # expressions.
        for hidden_re in self.HIDDEN_REGEXPS:
            if DEBUG:
                print "Checking for hidden expression"
            i = 0 
            # Get list of expanded timecodes in text
            hidden_seqs = hidden_re.findall(self.GetText())
            if DEBUG:
                print "Hidden seqs found = %s" % hidden_seqs
            # Hide each one found
# FIXME: It seems to be messing up on unicode strings that it gets here?
# I have to style for len(seq) + 1 if it's unicode maybe?  figure this
# out...
            for seq in hidden_seqs:
                if DEBUG:
                    print "type = %s" % type(seq)
                seq_start = self.FindText(i, 999999, seq, 0)
                if DEBUG:
                    print "Sequence starts at index %d" % seq_start
                i = seq_start + 1
                self.StartStyling(seq_start, STYLE_MASK)
                if DEBUG:
                    print "Styling for len = %d" % len(seq)
                # UNICODE?? self.SetStyling(len(seq) + 1, self.STYLE_HIDDEN)
                self.SetStyling(len(seq), self.STYLE_HIDDEN)
                    
        self.ProgressDlg.Update(100)

    def __ResetBuffer(self):
        if DEBUG:
            print "RichTextEditCtrl: ResetBuffer()"
        self.SetText("")
        self.EmptyUndoBuffer()
        self.StyleSetSpec(stc.STC_STYLE_DEFAULT, "size:" + str(TransanaGlobal.configData.defaultFontSize) + ",face:" + TransanaGlobal.configData.defaultFontFace)
        self.StyleClearAll()

        self.style_specs = []
        self.style_attrs = []
        self.num_styles = 0

        self.last_pos = 0

        self.__GetStyle("size:" + str(TransanaGlobal.configData.defaultFontSize) + ",face:" + TransanaGlobal.configData.defaultFontFace)
        self.style = 0          # "Current" style
        self.STYLE_HIDDEN = self.__GetStyle("hidden")
        self.STYLE_TIMECODE = self.__GetStyle("timecode")
        self.StyleSetVisible(self.STYLE_HIDDEN, False)
        self.StyleSetSpec(self.STYLE_TIMECODE, "fore:#FF0000,size:" + str(TransanaGlobal.configData.defaultFontSize) + ",face:" + TransanaGlobal.configData.defaultFontFace)

    def InsertHiddenText(self, text):
        """Insert hidden text at current cursor position."""
        # FIXME: Unicode issues with this?  Seems like len() for
        # SetStyling is 1 too short?
        curpos = self.GetCurrentPos()
        self.StartStyling(curpos, STYLE_MASK)
        self.InsertText(curpos, text)
        self.GotoPos(curpos+len(text))
        self.SetStyling(len(text), self.STYLE_HIDDEN)
        # Ensure that the OnPosChanged event isn't fooled into switching
        # to this new style
        self.last_pos = curpos + len(text)

    def InsertStyledText(self, text):
        """Insert text with the current style."""
        curpos = self.GetCurrentPos()
        self.StartStyling(curpos, STYLE_MASK)
        self.InsertText(curpos, text)
        self.GotoPos(curpos+len(text))
        self.SetStyling(len(text), self.style)

    def ClearDoc(self):
        """Clear the document buffer."""
        self.__ResetBuffer()

    def LoadRTFData(self, data):
        """Load a RTF document into the editor with the document as a
        buffer (string)."""
        
        # print "ReadOnly? %s" % str(self.GetReadOnly())

        # Signal that we ARE loading data, so that the problem with font specification is avoided.
        self.loadingData = True
        
        RichTextEditCtrl.ClearDoc(self) # Don't want to call any parent method too

        try:
            parse = RTFParser.RTFParser()
            parse.init_progress_update(self.ProgressDlg.Update, 0, 50)
            parse.buf = data
            parse.read_stream()
        except RTFParser.RTFParseError:
            # If the length of the data passed is > 0, display the error message from the Parser
            if DEBUG and len(data) > 0:
                print sys.exc_info()[0], sys.exc_info()[1]
            # If the length of the data is 0, we can safely ignore the parser error.
            else:
                pass
        except:
            print "Unhandled/ignored exception in RTF stream parsing"
            print sys.exc_info()[0], sys.exc_info()[1]
            traceback.print_exc()
        else:
            self.__ParseRTFStream(parse.stream)

        if DEBUG:
            print
            print "self.style_specs:"
            for i in range(len(self.style_specs)):
                print "%3d %s" % (i, self.style_specs[i])

        # Okay, we're done loading the data now.
        self.loadingData = False

    def LoadRTFFile(self, filename):

        # print "LoadRTFFile", filename
        
        """Load a RTF file into the editor."""
        RichTextEditCtrl.ClearDoc(self) # Don't want to call any parent method too
        try:
            parse = RTFParser.RTFParser(filename)
            parse.init_progress_update(self.ProgressDlg.Update, 0, 50)
            parse.read_stream()
        except:
            print "Unhandled/ignored exception in RTF parsing"
            traceback.print_exc()
        else:
            self.__ParseRTFStream(parse.stream)

    def LoadDocument(self, filename):
        """Load a file into the editor."""
        if filename.upper().endswith(".RTF"):
            self.LoadRTFFile(filename)

    def __ApplyStyle(self, style):
        """Apply the style to the current selection, if any."""
        startpos = self.GetSelectionStart()
        endpos = self.GetSelectionEnd()
        if startpos != endpos:
            self.StartStyling(startpos, STYLE_MASK)
            self.SetStyling(endpos-startpos, self.style)

    def GetBold(self):
        """Get bold state for current font or for the selected text."""
        return self.style_attrs[self.style].bold
        #props = self.__CurrentSpec().split(",")
        #return props.count("bold") > 0

    def GetItalic(self):
        """Get italic state for current font or for the selected text."""
        return self.style_attrs[self.style].italic
        #props = self.__CurrentSpec().split(",")
        #return props.count("italic") > 0

    def GetUnderline(self):
        """Get underline state for current font or for the selected text."""
        return self.style_attrs[self.style].underline
        #props = self.__CurrentSpec().split(",")
        #return props.count("underline") > 0

    def __SetAttrValue(self, attr, value):
        """Set the value of an attribute in the spec.  Never remove it.
        For example, use attr='face', and value='Times' to change the font
        face to Times."""
        props = self.style_specs[self.style].split(",")
        newprops = []
        hadface = 0
        for x in props:
            if x[:5] == attr + ":":
                newprops.append("%s:%s" % (attr, value))
                hadface = 1
            else:
                newprops.append(x)
        if not hadface:
            newprops.append("%s:%s" % (attr, value))
        self.style = self.__GetStyle(string.join(newprops, ","))

    def __SetFontFace(self, face):
        self.__SetAttrValue("face", face)
        return
        # DEAD CODE FOLLOWS, was replaced with the above (cleaner)  
        props = self.style_specs[self.style].split(",")
        newprops = []
        hadface = 0
        for x in props:
            if x[:5] == "face:":
                newprops.append("face:%s" % face)
                hadface = 1
            else:
                newprops.append(x)
        if not hadface:
            newprops.append("face:%s" % face)
        self.style = self.__GetStyle(string.join(newprops, ","))


    def __SetFontColor(self, fg, bg):
        self.__SetAttrValue("fore", "#%06x" % fg)
        self.__SetAttrValue("back", "#%06x" % bg)

    def __SetFontSize(self, size):
        self.__SetAttrValue("size", "%d" % size)
        
    def __SetAttr(self, attr, state=1):
        """Set a style attribute state for the current style."""
        props = self.style_specs[self.style].split(",")
        if DEBUG:
            print "SetAttr: Existing props: %s" % str(props)
        has_attr = props.count(attr) > 0
        if state:
            if not has_attr:
                if DEBUG:
                    print "SetAttr: Have to add attr to style"
                #print "SetAttr: CurrentSpec = %s" % self.__CurrentSpec()
                # The problem with using CurrentSpec() is that it depends
                # on the current selection or position.  I think we want
                # to just use the spec from the current style though, and
                # let whoever calls this worry about setting it to the
                # one that's at the current position or selection or whatever.
                #self.style = self.__GetStyle(self.__CurrentSpec() + "," + attr)
                self.style = self.__GetStyle(self.style_specs[self.style] + "," + attr)
        else:
            if has_attr:
                if DEBUG:
                    print "SetAttr: Have to remove attr from style"
                props.remove(attr)
                self.style = self.__GetStyle(string.join(props, ","))
    
    def SetBold(self, state=1):
        """Set bold state for current font or for the selected text."""
        self.__SetAttr("bold", state)
        self.__ApplyStyle(self.style)
        if DEBUG:
            print "SetBold: now self.style = %d" % self.style
    
    def SetItalic(self, state=1):
        """Set italic state for current font or for the selected text."""
        self.__SetAttr("italic", state)
        self.__ApplyStyle(self.style)
        if DEBUG:
            print "SetItalic: now self.style = %d" % self.style

    def SetUnderline(self, state=1):
        """Set underline state for current font or for the selected text."""
        self.__SetAttr("underline", state)
        self.__ApplyStyle(self.style)
        if DEBUG:
            print "SetUnderline: now self.style = %d" % self.style

    def SetFont(self, face, size, fg_color, bg_color):
        """Set the font."""
        #print "SetFont: face: %s, size:%d, color fg:#%06x, bg: #%06x" % (face, size, fg_color, bg_color)
        if len(face) > 0:
            self.__SetFontFace(face)
        self.__SetFontColor(fg_color, bg_color)
        self.__SetFontSize(size)
        #print "After SetATtr: Style=%d" % self.style
        self.__ApplyStyle(self.style)
        # This ensures that the OnPosChanged event doesn't mistakingly
        # change back the font.  Probably should also add this to the
        # SetBold/Italic/Underline?
        self.last_pos = self.GetCurrentPos()
        #self.__SetAttr("size:%d" % size, state)

    def GetFont(self):
        """Get the current (font face, size, color) in a tuple."""
        attr = self.style_attrs[self.style]
        return (attr.font_face, attr.font_size, attr.font_fg)

    def __NewRTFDoc(self, file_or_fname):
        """Create a new RTF output stream from a given file handle or
        filename."""
        ss = TextDoc.StyleSheet()
        ps = TextDoc.ParagraphStyle()
        ss.add_style("Default", ps)
        pps = TextDoc.PaperStyle("Transcript", 27.94, 21.59)

        doc = RTFDoc.RTFDoc(ss, pps, None)

        # Build a list of the fonts/colors used in this document
        faces = []
        colors = []
        for attr in self.style_attrs:
            if faces.count(attr.font_face) == 0:
                faces.append(attr.font_face)
            if colors.count(attr.fg_tuple()) == 0:
                colors.append(attr.fg_tuple())
            if colors.count(attr.bg_tuple()) == 0:
                colors.append(attr.bg_tuple())
        
        doc.open(file_or_fname, faces, colors)
        return doc
    
    def __ProcessDocAsRTF(self, doc, select_only=0):
        """Process the document data for output as a RTF document.
        If select_only=1, then it will only process the current selection
        instead of the whole document."""
        doc.start_paragraph("Default")
        

        if select_only:
            start = self.GetSelectionStart()
            # You need the "-1" here to prevent an extra character from being included in the Clip Transcript.
            end = self.GetSelectionEnd() - 1
        else:
            start = 0
            end = self.GetLength() - 1

        for x in range(start, end + 1):
            c = chr(self.GetCharAt(x))
            if ord(c) == 0xc2:
                # FIXME: For some reason we're seeing a 0xC2 (194) in the 
                # document everytime there's a timecode.  Ignore them as a
                # temporary fix.
                continue
            style = self.GetStyleAt(x)

            if self.style_attrs[style].bold:
                if not doc.bold_on:
                    doc.start_bold()
            else:
                if doc.bold_on:
                    doc.end_bold()
            
            if self.style_attrs[style].italic:
                if not doc.italic_on:
                    doc.start_italic()
            else:
                if doc.italic_on:
                    doc.end_italic()
            
            if self.style_attrs[style].underline:
                if not doc.underline_on:
                    doc.start_underline()
            else:
                if doc.underline_on:
                    doc.end_underline() 
          
            if self.style_attrs[style].font_face != doc.fontFace:
                doc.set_ext_font(self.style_attrs[style].font_face, \
                        self.style_attrs[style].font_size)

            fg = self.style_attrs[style].fg_tuple()
            bg = self.style_attrs[style].bg_tuple()
            if (fg != doc.fgColor) or (bg != doc.bgColor):
                doc.set_ext_color(fg, bg)

            if self.style_attrs[style].font_size != doc.fontSize:
                doc.set_ext_font(self.style_attrs[style].font_face, \
                        self.style_attrs[style].font_size)
                
            if c == '\r':
                # Do nothing.
                pass
            elif c == '\n':
                doc.line_break()
            else:
                doc.write_text(c)
       
        # This was the old way without attribute support.
        #text = self.GetText()
        #lines = text.splitlines()
        #for line in lines:
        #    doc.write_text(line)
        #    doc.line_break()
        
        doc.end_paragraph()
        
    def GetRTFBuffer(self, select_only=0):
        """Get a string of data in RTF format of the current document in
        memory.  If select_only = 1, then only the current selection will be
        processed instead of the whole document."""
        # Use a "virtual file" with StringIO since the TextDoc module
        # only supports RTF document creation with a file handle.
        f = cStringIO.StringIO()
        doc = self.__NewRTFDoc(f)
        self.__ProcessDocAsRTF(doc, select_only)
        doc.close()
        f.seek(0)
        data = f.read()
        f.close()
        return data

    def SaveRTFDocument(self, filename):
        """Save the document in memory to the given file."""
        doc = self.__NewRTFDoc(filename)
        self.__ProcessDocAsRTF(doc)
        doc.close()

    def SetDragEvent(self, id, func):
        stc.EVT_STC_START_DRAG(self, id, func)

    def OnCharAdded(self, event):
        """Called when a character is added to the document."""
        if DEBUG:
            print "OnCharAdded(): %s (style=%d)" % (event.GetKey(), self.style)
        # Don't do anything if in read-only mode
        if self.GetReadOnly():
            return
        self.StartStyling(self.GetCurrentPos() - 1, STYLE_MASK)
        self.SetStyling(1, self.style)
        if DEBUG:
            #print "dir(event): %s" % dir(event)
            print "event: %s" % str(event)

    def OnPosChanged(self, event):
        """Called when the position in the document has changed."""
        # Actually we are using UPDATEUI to trigger this because
        # POSCHANGED doesn't seem to work, so this will be called for
        # a few other events too (repainting, etc).
        pos = self.GetCurrentPos()
        # Only do this when the cursor moved
        if pos != self.last_pos:
            self.last_pos = pos
            # UPDATE: This logic seemed problematic.  When you would type
            # admist characters of another style it would favor the one
            # after it.  Not what the user expects.  For now we will just
            # always use the character before the current one.
            # Logic for determining which style to use when the cursor
            # position changes:  Use the style of the character the cursor
            # is at, unless the character is a newline, meaning it's at
            # the end of the line.  In this case, we take the style of
            # the character before the current character.
            #if self.GetCharAt(pos) == 10:
            #    self.style = self.GetStyleAt(pos-1)
            #else:
            #    self.style = self.GetStyleAt(pos)
            if pos == 0:
                style = self.GetStyleAt(pos)
            else:
                style = self.GetStyleAt(pos-1)

            # We never want to do manual user input as hidden text.
            # DKW -- However, this was interfering with loading data properly, so I added the
            #        self.loadingData parameter
            if (not self.loadingData) and (style != self.STYLE_HIDDEN):
                prevstyle = self.style
                self.style = style
                if self.style != prevstyle:
                    #print "Emitting StyleChanged event (style = %d, prevstyle = %d)" % (style, prevstyle)
                    # Emit a "style changed" event since the style has changed
                    # automatically.
                    # NOTE:  This statement causes an exception if pos = 0 and the first character is
                    # a timecode and you press the CURSOR_LEFT key.  I don't understand why, so I added
                    # the "pos != 0" test here to prevent the problem -- DKW
                    if (self.StyleChanged) and (pos != 0):
                        self.StyleChanged(self)
       
            
    def SetSavePoint(self):
        """Override the existing SetSavePoint to account for style changes."""
        wx.stc.StyledTextCtrl.SetSavePoint(self)
        self.stylechange = 0

    def GetModify(self):
        """Override the existing GetModify to account for style changes."""
        return wx.stc.StyledTextCtrl.GetModify(self) or self.stylechange
        
    def OnModified(self, event):
        """Triggered when the document is modified, including style
        changes."""
        # Note: No modifications may be performed when servicing this event!
        if event.GetModificationType() & wx.stc.STC_MOD_CHANGESTYLE:
            self.stylechange = 1

        
    # Many of these won't be implemented.  Delete as you find that they're
    # unnecessary.
    def set_default_font(self, font):
        """Change the default font."""
        
    def set_font(self, font):
        """Change the current font or the font for the selected text."""
        
    def cut_selected_text(self):
        """Delete selected text and place in clipboard."""
        
    def copy_seleted_text(self):
        """Copy selected text to clipboard."""
        
    def paste_text(self):
        """Paste text from clipboard."""
        
    def select_all(self):
        """Select all document text."""
        
    def find_text(self, text, matchcase, wraparound):
        """Find text in document."""
        
    def cursor_find(self, text):
        """Move the cursor to the next occurrence of given text in the
        transcript (for word tracking)."""
        
    def select_find(self, text):
        """Select the text from the current cursor position to the next
        occurrence of given text (for word tracking)."""
        
    def spell_check(self):
        """Interactively spell-check document."""
        
    def undo(self):
        """Undo last operation(s)."""
        
    def redo(self):
        """Redo last undone operation(s)."""
        
# Events    
    def EVT_DOC_CHANGED(self, win, id, func):
        """Set function to be called when document is modified."""

class StyleSettings:

    def __init__(self):
        self.bold = 0
        self.italic = 0
        self.underline = 0
        self.font_face = TransanaGlobal.configData.defaultFontFace
        self.font_size = TransanaGlobal.configData.defaultFontSize
        self.font_fg = 0x000000
        self.font_bg = 0xffffff

    def fg_tuple(self):
        """Return RGB tuple of font fg for convenience."""
        return ((self.font_fg & 0xff0000) >> 16, (self.font_fg & 0x00ff00) >> 8, self.font_fg & 0x0000ff)

    def bg_tuple(self):
        """Return RGB tuple of font bg for convenience."""
        return ((self.font_bg & 0xff0000) >> 16, (self.font_bg & 0x00ff00) >> 8, self.font_bg & 0x0000ff)
