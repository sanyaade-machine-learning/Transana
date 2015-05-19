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

""" This module implements the RichTextEditCtrl class, a Rich Text editor based on StyledTextCtrl. """

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>, Jonathan Beavers <jonathan.beavers@gmail.com>'

DEBUG = False
if DEBUG:
    print "RichTextEditCtrl DEBUG is ON."
DEBUG2 = False
if DEBUG2:
    print "RichTextEditCtrl DEBUG2 is ON."
SHOWHIDDEN = False
if SHOWHIDDEN:
    print "RichTextEditCtrl SHOWHIDDEN is ON."

import TransanaConstants        # import Transana's Constants
import TransanaGlobal           # import Transana's Globals
import Dialogs                  # import Transana's Dialogs for the ErrorDialog
import sys, os, string, re      # import Python modules that are needed
import pickle                   # the pickle module enables us to fast-save

import wx                       # import wxPython
from wx import stc              # import wxPython's StyledTextCtrl

# Add the "rtf" path to the Python Path so that the RTF Parser can be found.
RTFModulePath = "rtf"
if sys.path.count(RTFModulePath) == 0:
    sys.path.append(RTFModulePath)  # Add to path if not already there

import RTFParser                # import our RTF Parser for reading RTF files
import RTFDoc, TextDoc          # import RTFDoc and TextDoc for writing docs to RTF format
# NOTE:  cStringIO is faster, but does not support Unicode strings!!!   DKW
if 'unicode' in wx.PlatformInfo:
    import StringIO             # If we're using Unicode, we must use the Unicode-compatible StringIO
else:
    import cStringIO            # If we're using ansi, we can use the faster cStringIO
import traceback                # Import Python's traceback module for error reporting in DEBUG mode

# This mask determines the number of style bits to use.  The default is
# 31, which uses 5 bits (2^5) for 32 possible styles.  We increase this
# to 128 possible styles since we don't use the indicators feature that
# is normally used for the other 3 bits.  We have to use at least
# one indicator so we use 0x7f instead of 0xff.  This is the maximum that
# the Styled Text Control allows at this time.
STYLE_MASK = 0x7f
# Although we can technically support 127 styles, because of the way styles are created and stored, which
# can result in significant changes in the number defined, we'd better set a lower limit for now.
NUM_STYLES_SUPPORTED = 100   # 127

class RichTextEditCtrl(stc.StyledTextCtrl):
    """ This class is a Rich Text edit control implemented for Transana """

    def __init__(self, parent, id=-1):
        """Initialize a StyledTextCtrl object."""
        stc.StyledTextCtrl.__init__(self, parent, id)

        # The text display on the Mac is simply TOO SMALL.  This makes it a bit bigger by default.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetZoom(3)
        # We need as many styles as possible, and will not use Indicators.  This goes hand-in-hand with the STYLE_MASK defined above.
        self.SetStyleBits(7)
       
        # STYLE_HIDDEN is meant for things that are sometimes hidden and
        # sometimes visible, depending on the ShowTimeCodes value.
        # STYLE_HIDDEN_ALWAYS is intended for things that should never be
        # visible to the user.  (It's not actually used.)

        # Any style assigned here needs to be added to __ResetBuffer() as well.
        self.STYLE_HIDDEN = -1              # Define the Hidden Style for Time Codes and Time Code Data
        self.STYLE_TIMECODE = -1            # Define the Time Code Style for Time Codes when not hidden
        # self.STYLE_HIDDEN_ALWAYS = -1     NOT USED??
        # self.HIDDEN_CHARS = []            NOT USED??
        self.HIDDEN_REGEXPS = []            # Initialize a list for Regular Expressions for hidden data
        self.__ResetBuffer()                # Initialize the STC, including defining default font(s)

        # Set the Default Style based on the Default Font information in the Configuration object
        self.StyleSetSpec(stc.STC_STYLE_DEFAULT, "size:%s,face:%s,fore:#000000,back:#ffffff" % (str(TransanaGlobal.configData.defaultFontSize), TransanaGlobal.configData.defaultFontFace))
        # Set the Line Number style based on the Default Font Face
        self.StyleSetSpec(stc.STC_STYLE_LINENUMBER, "size:10,face:%s" % TransanaGlobal.configData.defaultFontFace)
        # Set Word Wrap as configured
        self.SetWrapMode(TransanaGlobal.configData.wordWrap)  # (stc.STC_WRAP_WORD)

        # Setting the LayoutCache to the whole document seems to reduce the typing lag problem on the PPC Mac
        self.SetLayoutCache(stc.STC_CACHE_DOCUMENT)
        # Additional suggestions from wxPython-mac user's mailing list:
        # Limit the events that trigger the EVT_STC_MODIFIED handler
        self.SetModEventMask(stc.STC_PERFORMED_UNDO | stc.STC_PERFORMED_REDO | stc.STC_MOD_DELETETEXT | stc.STC_MOD_INSERTTEXT)
        # Turn off Anti-Aliasing
        self.SetUseAntiAliasing(False)

        # Set the Tab Width to the configured size
        self.SetTabWidth(int(TransanaGlobal.configData.tabSize))
        # Indicate that we would like Line Numbers in the STC Margin
        self.SetMarginType(0, stc.STC_MARGIN_NUMBER)
        # We need a slighly wider Line Number section on the Mac.  This code determines an optimum width by platform
        if "__WXMAC__" in wx.PlatformInfo:
            self.lineNumberWidth = self.TextWidth(stc.STC_STYLE_LINENUMBER, '88888 ')
        else:
            self.lineNumberWidth = self.TextWidth(stc.STC_STYLE_LINENUMBER, '8888 ')
        # Apply the Line Number section width determined above
        self.SetMarginWidth(0, self.lineNumberWidth)
        # Set the STC right and left Margins
        self.SetMargins(6, 6)

        # Set the STC Selection Colors to white text on a blue background
        self.SetSelForeground(1, "black")   # "white"
        self.SetSelBackground(1, "cyan")   # "blue"

        # We need to capture key strokes so we can intercept the Paste command.
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)

        # Define the STC Events we need to implement or modify.
        # Character Added event
        stc.EVT_STC_CHARADDED(self, id, self.OnCharAdded)
        # User Interface Update event
        stc.EVT_STC_UPDATEUI(self, id, self.OnUpdateUI)
        # STC Modified event
        stc.EVT_STC_MODIFIED(self, id, self.OnModified)
        # Remove Drag-and-Drop reference on the Mac due to the Quicktime Drag-Drop bug.
        # (On the Mac, any use of Drag and Drop causes the Video Component to jump to the wrong
        #  Window!  Therefore, we must disable Drag and Drop wherever we can on the Mac.)
        if (not TransanaConstants.macDragDrop) and ("__WXMAC__" in wx.PlatformInfo):
            stc.EVT_STC_START_DRAG(self, id, self.OnKillDrag)

        # Initialzations:
        # We need to know if a style has been changed, as this can indicate the need for a Save.
        self.stylechange = 0
        # If set to point to a wxProgressDialog, this dialog will be updated as a document is loaded.
        # Initialize to None.
        self.ProgressDlg = None
        # There was a problem with the font being changed inappropriately during
        # data loading.  This flag helps prevent that.
        self.loadingData = False

    def OnKeyDown(self, event):
        """ We need to capture the Paste call to intercept it and remove Time Codes from the
            Clipboard Data """
        # We only need to intercept Ctrl-V on Windows or Command-V on the Mac.  First, see if the
        # platform-appropriate modified key is currently pressed
        if (('wxMac' not in wx.PlatformInfo) and event.ControlDown()) or \
           (('wxMac' in wx.PlatformInfo) and event.CmdDown()):
            # Look for "C" (Copy)
            if event.GetKeyCode() == 67:    # Ctrl-C
                # If so, call our modified Copy() method
                self.Copy()
            # Look for "V" (Paste)
            elif event.GetKeyCode() == 86:  # Ctrl-V
                # If so, call our modified Paste() method
                self.Paste()
            # Look for "X" (Cut)
            elif event.GetKeyCode() == 88:  # Ctrl-X
                # If so, call our modified Cut() method
                self.Cut()
            # For anything other than Ctrl/Cmd- C, V, or X ...
            else:
                # ... we let processing drop to the STC
                event.Skip()
        # If the keystroke isn't modified platform-approptiately ...
        else:       
            # ... we let processing drop to the STC
            event.Skip()

    def PutEditedSelectionInClipboard(self):
        """ Put the TEXT for the current selection into the Clipboard """
        tempTxt = self.GetSelectedText()
        # Initialize an empty string for the modified data
        newSt = ''
        # Created a TextDataObject to hold the text data from the clipboard
        tempDataObject = wx.TextDataObject()
        # Track whether we're skipping characters or not.  Start out NOT skipping
        skipChars = False
        # Now let's iterate through the characters in the text
        for ch in tempTxt:
            # Detect the time code character
            if ch == TransanaConstants.TIMECODE_CHAR:
                # If Time Code, start skipping characters
                skipChars = True
            # if we're skipping characters and we hit the ">" end of time code data symbol ...
            elif (ch == '>') and skipChars:
                # ... we can stop skipping characters.
                skipChars = False
            # If we're not skipping characters ...
            elif not skipChars:
                # ... add the character to the new string.
                newSt += ch
        # Save the new string in the Text Data Object
        tempDataObject.SetText(newSt)
        # Open the Clipboard
        wx.TheClipboard.Open()
        # Place the Text Data Object in the Clipboard
        wx.TheClipboard.SetData(tempDataObject)
        # Close the Clipboard
        wx.TheClipboard.Close()

    def Cut(self):
        if 'wxMac' in wx.PlatformInfo:
            self.PutEditedSelectionInClipboard()
            self.ReplaceSelection('')
        else:
            stc.StyledTextCtrl.Cut(self)

    def Copy(self):
        if 'wxMac' in wx.PlatformInfo:
            self.PutEditedSelectionInClipboard()
        else:
            stc.StyledTextCtrl.Copy(self)

    def Paste(self):
        """ We need to intercept the STC's Paste() method to deal with time codes.  We need to strip
            time codes out of the clipboard before pasting, as they don't paste properly. """
        # Initialize an empty string for the modified data
        newSt = ''
        # Created a TextDataObject to hold the text data from the clipboard
        tempDataObject = wx.TextDataObject()
        # Open the Clipboard
        wx.TheClipboard.Open()
        # Read the text data from the clipboard
        wx.TheClipboard.GetData(tempDataObject)
        # Track whether we're skipping characters or not.  Start out NOT skipping
        skipChars = False
        # Now let's iterate through the characters in the text
        for ch in tempDataObject.GetText():
            # Detect the time code character
            if ch == TransanaConstants.TIMECODE_CHAR:
                # If Time Code, start skipping characters
                skipChars = True
            # if we're skipping characters and we hit the ">" end of time code data symbol ...
            elif (ch == '>') and skipChars:
                # ... we can stop skipping characters.
                skipChars = False
            # If we're not skipping characters ...
            elif not skipChars:
                # ... add the character to the new string.
                newSt += ch
        # Save the new string in the Text Data Object
        tempDataObject.SetText(newSt)
        # Place the Text Data Object in the Clipboard
        wx.TheClipboard.SetData(tempDataObject)
        # if we're on Mac, we need to close the Clipboard to call Paste.  We don't need this on
        # Windows, but the Mac requires it.
        if 'wxMac' in wx.PlatformInfo:
            wx.TheClipboard.Close()
        # Call the STC's Paste method to complete the paste.
        stc.StyledTextCtrl.Paste(self)
        # If the Clipboard is open ...
        if wx.TheClipboard.IsOpened():
            # Close the Clipboard
            wx.TheClipboard.Close()

    def __ResetBuffer(self):
        """ Reset the wxSTC buffer -- that is, re-initialize the Rich Text Control's data """
        
        if DEBUG:
            print "RichTextEditCtrl: ResetBuffer()"

        # Set the wxSTC Text to blank
        self.SetText("")
        # Empty out the Undo Buffer, removing old Undo data
        self.EmptyUndoBuffer()

        # Redefine the Default Style, just in case
        self.StyleSetSpec(stc.STC_STYLE_DEFAULT, "size:%s,face:%s,fore:#000000,back:#ffffff" % (str(TransanaGlobal.configData.defaultFontSize), TransanaGlobal.configData.defaultFontFace))
        # Clear out all defined styles.  Internally, the wxSTC resets all styles to the default style.
        self.StyleClearAll()
        # Reset the data structures that keep track of the defined styles.
        # The style_specs list shows style information in a human-readable, yet parsable, way
        self.style_specs = []
        # The style_attrs list shows style information in a machine-readable way
        self.style_attrs = []
        # num_styles tracks the number of styles that have been defined
        self.num_styles = 0

        # Reset the Last Position indicator
        self.last_pos = 0

        # Define the initial style, Style 0, based on the user-specified defaults
        self.__GetStyle("size:%s,face:%s,fore:#000000,back:#ffffff" % (str(TransanaGlobal.configData.defaultFontSize), TransanaGlobal.configData.defaultFontFace))
        # Reset the global style currently being used to this default style
        self.style = 0
        # Redefine the Hidden and Timecode styles, which Transana always requires.
        self.STYLE_HIDDEN = self.__GetStyle("hidden")
        self.STYLE_TIMECODE = self.__GetStyle("timecode")
        # The RTF Parser seems to insist on adding these two "default" styles.  Let's add them here to reduce conflicts
        # with the number of styles that exist in new Transcripts vs. saved Transcripts.
        # NOTE:  These could be removed if the Style Definition code is cleaned up.
        self.__GetStyle("size:12,face:Times New Roman,fore:#000000,back:#ffffff")
        self.__GetStyle("size:12,face:Courier New,fore:#000000,back:#ffffff")
        
        # If we want to show the normally hidden time code data for debugging purposes, this does it!
        if not SHOWHIDDEN:
            self.StyleSetVisible(self.STYLE_HIDDEN, False)

        # Define the particulars of the Timecode style, which is to use the default font in red.            
        self.StyleSetSpec(self.STYLE_TIMECODE, "size:%s,face:%s,fore:#FF0000,back:#ffffff" % (str(TransanaGlobal.configData.defaultFontSize), TransanaGlobal.configData.defaultFontFace))

    def GetStyleAccessor(self, spec):
        """This method is simply a public wrapper around __GetStyle()"""
    	return self.__GetStyle(spec)

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

        if DEBUG:
            print "RichTextEditCtrl.__GetStyle(): spec = '%s', count = %d" % (spec, self.style_specs.count(spec))

        # If the desired style isn't already defined ...
        if self.style_specs.count(spec) == 0:

            if DEBUG:
                print "RichTextEditCtrl.__GetStyle():  Allocating new style (spec %s)" % spec

            # NOTE:  This is really messed up.  Given the current (Aug. 11, 2005) infrastructure, specifying a new
            #        style can cause multiple styles to be created along the way.  Saving and reloading a document
            #        can cause the number of styles to change, as some unused styles created above get dropped, but some
            #        style combinations that aren't used and weren't created above get created after a save.  And
            #        saving a second time seems to have the effect of reducing the number of styles considerably.
            #        In one experiment, I had 31 defined styles.  I added 5 styles, which took my total up to 45.
            #        Then I saved it and reloaded, and was faced with 65 styles.  Then I saved again, and upon reload
            #        found that I had 49 styles defined.

            #        Because of this, trapping the number of styles won't really work until the style creation mechanism
            #        for the RichTextEditCtrl is rewritten.
            
            # Check to see if we can support an additional style.
            if (self.num_styles >= NUM_STYLES_SUPPORTED):   # STYLE_MASK
                if DEBUG:
                    print "No more styles left.  Using Default"
                # We need the ErrorDialog from the Dialogs module.
                import Dialogs
                # Build and display an error message
                msg = _("Transana is only able to handle %d different font styles.  You have exceeded that capacity.\nYour text will revert to a previously-used style.\nPlease reduce the number of unique styles used in this transcript.")
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(msg, 'utf8')
                dlg = Dialogs.ErrorDialog(self, msg % NUM_STYLES_SUPPORTED)
                dlg.ShowModal()
                dlg.Destroy()
                # No more styles left, must use default
                return stc.STC_STYLE_DEFAULT

            # Okay, the stc.STC_STYLE_LINENUMBER style is getting over-written, which changes the style of the line numbers.
            # We have to do a trick here to avoid that.  HACK HACK HACK.
            # Essentially, what this code does is insert the appropriate style when we get to this point in the
            # style array.  Apparently, the information is stored internally in the wxSTC until this 33rd style is created.
            # This is necessary because we are using 7 style bits instead of 5, which is the wx.STC default.
            while (self.num_styles == stc.STC_STYLE_LINENUMBER) or (self.num_styles == stc.STC_STYLE_DEFAULT):
                
                if DEBUG:
                    print "Skipping %d because it equals %d (stc.STC_STYLE_LINENUMBER) or %s (stc.STC_STYLE_DEFAULT)" % (self.num_styles, stc.STC_STYLE_LINENUMBER, stc.STC_STYLE_DEFAULT)

                attr = StyleSettings()
                if (self.num_styles == stc.STC_STYLE_LINENUMBER):
                    # Define the StyleSpec appropriate for stc.STC_STYLE_LINENUMBER
                    self.StyleSetSpec(self.num_styles, "size:10,face:%s,fore:#000000,back:#ffffff" % TransanaGlobal.configData.defaultFontFace)
                    # Also place that style in style_specs
                    self.style_specs.append("size:10,face:%s,fore:#000000,back:#ffffff" % TransanaGlobal.configData.defaultFontFace)
                    attr.font_size = 10
                else:
                    # Define the StyleSpec appropriate for stc.STC_STYLE_DEFAULT
                    self.StyleSetSpec(self.num_styles, 'size:%s,face:%s,fore:#000000,back:#ffffff' % (TransanaGlobal.configData.defaultFontSize, TransanaGlobal.configData.defaultFontFace))
                    # Also place that style in style_specs
                    self.style_specs.append('size:%s,face:%s,fore:#000000,back:#ffffff' % (TransanaGlobal.configData.defaultFontSize, TransanaGlobal.configData.defaultFontFace))
                    attr.font_size = TransanaGlobal.configData.defaultFontSize
                # And also define the Attributes for the style ...
                attr.font_face = TransanaGlobal.configData.defaultFontFace
                attr.font_size = 10
                attr.font_fg = 0x000000
                # ... and store the style in the style_attrs list.
                self.style_attrs.append(attr)
                # Increment the style counter.
                self.num_styles += 1

            if DEBUG:
                print "RichTextEditCtrl.__GetStyles():  StyleSetSpec(%s, %s)" % (self.num_styles, spec)
                
            # Define the style for wx.STC
            self.StyleSetSpec(self.num_styles, spec)
            
            # Hidden text, when shown, is Red by default
            if spec == "hidden":
                self.StyleSetForeground(self.num_styles, wx.NamedColour("red"))

            # Add the style to the style_specs list, which allows us to keep track of what styles have been
            # defined internally by wx.STC.
            self.style_specs.append(spec)

            # We also need to gather the current attributes ...
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
            # ... and store them in the style_attrs list
            self.style_attrs.append(attr)

            # Increment the style counter
            self.num_styles += 1
            
            if DEBUG:
                print "New style #%d = %s" % (self.num_styles - 1, spec)

            # Return the style number, using -1 because we just incremented the counter!
            return self.num_styles - 1
        else:
            # Return the number that matches the specified style
            return self.style_specs.index(spec)

    def CurrentSpec(self):
        return self.__CurrentSpec()

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

            if DEBUG and (obj.text != ''):
                print
                print "RichTextEditCtrl.__ParseRTFStream(): next object in stream is:", obj
                print "RichTextEditCtrl.__ParseRTFStream(): type(obj.text):", type(obj.text)
                if type(obj.text).__name__ == 'str':
                    for c in obj.text:
                        print c, ord(c)

            # If we are opening a transcript from a non-unicode version of Transana, we need to convert
            # the text to unicode.
            if ('unicode' in wx.PlatformInfo) and isinstance(obj.text, str):

                # There is a problem importing transcripts on single-user Win version.  Specifically,
                # the closed dot doesn't import correctly.  This aims to fix that using character substitution.
                if obj.text == chr(149):
                    obj.text = unicode('\xe2\x80\xa2', 'utf8')
                else:
                    obj.text = unicode(obj.text, TransanaGlobal.encoding)
                
            if self.ProgressDlg:
                self.ProgressDlg.Update(prog_level)
            if obj.attr == None:

                if DEBUG:
                    print "RichTextCtrl.__ParseRTFStream():  '%s'" % obj.text.encode('utf8'),
                    for x in obj.text:
                        print ord(x),
                        if ord(x) < 127:
                            print x,
                    print

                if (('wxMac' in wx.PlatformInfo) and (ord(obj.text[0]) == 183)) or \
                   (('wxMSW' in wx.PlatformInfo) and (ord(obj.text[0]) == 183)):
                    if DEBUG:
                        print "RichTextCtrl.__ParseRTFStream():  Closed Dot"
                        
                    self.InsertInBreath()
                    
                elif (('wxMac' in wx.PlatformInfo) and (ord(obj.text[0]) == 176)) or \
                     (('wxMSW' in wx.PlatformInfo) and (ord(obj.text[0]) == 176)):

                    if DEBUG:
                        print "RichTextCtrl.__ParseRTFStream():  Open Dot"
                        
                    self.InsertWhisper()

                else:

                    # I've seen some sporadic problems converting RTF.  Specifically, chunks of the transcript get loaded
                    # in the wrong order, or in the wrong place in the new document.
                    #
                    # This problem can be detected by comparing GetCurrentPos() to GetLength(), since new text should
                    # ALWAYS be added to the end of the document.  If this problem is detected ...
                    if self.GetCurrentPos() != self.GetLength():
                        # ... correct it by resetting the CurrentPos to the last character in the document!
                        self.SetCurrentPos(self.GetLength())
                    
                    startpos = self.GetCurrentPos()

                    self.AddText(obj.text)

                    # If we receive Unicode objects, we need to decode them so we can figure out their correct length, which
                    # wxSTC needs to know for the purpose of styling the right number of bytes.
                    if ('unicode' in wx.PlatformInfo) and (type(obj.text).__name__ == 'unicode'):
                        # Don't use the current encoding, but always use UTF8, as that's what Python (and hence the
                        # wx.STC) is set for.  This fixes a bug with importing transcripts into Latin1 on single-user Win version.
                        txt = obj.text.encode('utf8') # (TransanaGlobal.encoding)
                    else:
                        txt = obj.text

                    if DEBUG:
                        print "RichTextEditCtrl.__ParseRTFStream(): Adding text '%s' with style %s, len =%d /%d" % (txt, self.style, len(obj.text), len(txt))
                        
                    self.StartStyling(startpos, STYLE_MASK)
                    self.SetStyling(len(txt), self.style)

            else:
                # Everytime there's an attribute change, we change the
                # 'current' attribute completely to the new one.

                # Let's not change the attributes one at a time.  That causes too many unused styles to be created.
                # Since we have a limit on the number of styles we can support, let's not waste any.
                # self.__SetAttr("bold", obj.attr.bold)
                # self.__SetAttr("underline", obj.attr.underline)
                # self.__SetAttr("italic", obj.attr.italic)
                # self.__SetFontFace(obj.attr.font)
                # self.__SetFontColor(obj.attr.fg, obj.attr.bg)
                # self.__SetFontSize(obj.attr.fontsize)

                # I'm not sure this is actually any better.  I don't have any more time to explore this right now.
                styleStr = "size:%d" % obj.attr.fontsize
                styleStr += ",face:%s" % obj.attr.font
                styleStr += ",fore:#%06x" % obj.attr.fg
                styleStr += ",back:#%06x" % obj.attr.bg
                if obj.attr.bold:
                    styleStr += ',bold'
                if obj.attr.underline:
                    styleStr += ',underline'
                if obj.attr.italic:
                    styleStr += ',italic'

                if DEBUG:
                    print "RichTextEditCtrl.__ParseRTFStream():  styleStr = '%s'" % styleStr

                self.style = self.__GetStyle(styleStr)
                
                if DEBUG:
                    print "RichTextEditCtrl.__ParseRTFStream(): New style after attribute change: %d" % self.style

            prog_level = prog_level + inc_level
    
        # Do a second pass through the text to check for hidden expressions.
        # This hides all of the Time Codes and the Time Code Data.
        for hidden_re in self.HIDDEN_REGEXPS:
            if DEBUG:
                print "Checking for hidden expression"
            i = 0 
            # Get list of expanded timecodes in text
            hidden_seqs = hidden_re.findall(self.GetText())
            if DEBUG:
                print "Hidden seqs found = %s" % hidden_seqs
            # Hide each one found
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
                if 'unicode' in wx.PlatformInfo:
                    self.SetStyling(len(seq) + 1, self.STYLE_HIDDEN)
                else:
                    self.SetStyling(len(seq), self.STYLE_HIDDEN)
                    
        if self.ProgressDlg:
            self.ProgressDlg.Update(100)

    def InsertRisingIntonation(self):
        """ Insert the Rising Intonation (Up Arrow) symbol """
        # NOTE:  Ugly HACK warning.
        #        The wxSTC on Mac isn't displaying Unicode characters correctly.  Therefore, we
        #        need to send a character that the wxSTC can display, even if it's not actually
        #        the character we want stored in the database.
        if 'unicode' in wx.PlatformInfo:
            
            # This is the "Unicode" way, which works well on Windows and inserts a character that can be saved
            # and loaded without difficulty.  Unfortunately, the Mac cannot display this character.
            # This has been abandoned for now, though might be resurrected if they fix the wxSTC.
            
            ch = unicode('\xe2\x86\x91', 'utf8')  # \u2191
            len = 3
            self.InsertUnicodeChar(ch, len)

            # The new approach, the "Symbol" way, is to insert the appropriate character in the Symbol
            # Font to display the up arrow character.  This is a different character on Windows than it
            # is on Mac.  Therefore, there needs to be code in loading and saving files that translates
            # this character so that files will work cross-platform.
#            if 'wxMac' in wx.PlatformInfo:
#                ch = unicode('\xe2\x89\xa0', 'utf8')
#                len = 3
#            else:
#                ch = unicode('\xc2\xad', 'utf8')
#                len = 2
            # We insert the specified character in Symbol Font.
#            self.InsertSymbol(ch, len)
        else:
            ch = '\xAD'
            self.InsertSymbol(ch)

    def InsertFallingIntonation(self):
        """ Insert the Falling Intonation (Down Arrow) symbol """
        # NOTE:  Ugly HACK warning.
        #        The wxSTC on Mac isn't displaying Unicode characters correctly.  Therefore, we
        #        need to send a character that the wxSTC can display, even if it's not actually
        #        the character we want stored in the database.
        if 'unicode' in wx.PlatformInfo:

            # This is the "Unicode" way, which works well on Windows and inserts a character that can be saved
            # and loaded without difficulty.  Unfortunately, the Mac cannot display this character.
            # This has been abandoned for now, though might be resurrected if they fix the wxSTC.
            
            ch = unicode('\xe2\x86\x93', 'utf8')
            len = 3
            self.InsertUnicodeChar(ch, len)

            # The new approach, the "Symbol" way, is to insert the appropriate character in the Symbol
            # Font to display the up arrow character.  This is a different character on Windows than it
            # is on Mac.  Therefore, there needs to be code in loading and saving files that translates
            # this character so that files will work cross-platform.
#            if 'wxMac' in wx.PlatformInfo:
#                ch = unicode('\xc3\x98', 'utf8')
#                len = 2
#            else:
#                ch = unicode('\xc2\xaf', 'utf8')
#                len = 2
            # We insert the specified character in Symbol Font.
#            self.InsertSymbol(ch, len)
        else:
            ch = '\xAF'
            self.InsertSymbol(ch)

    def InsertInBreath(self):
        """ Insert the In Breath (Closed Dot) symbol """
        if 'unicode' in wx.PlatformInfo:
            ch = unicode('\xe2\x80\xa2', 'utf8')
            len = 3
            self.InsertUnicodeChar(ch, len)
        else:
            ch = '\xB7'
            self.InsertSymbol(ch)

    def InsertWhisper(self):
        """ Insert the Whisper (Open Dot) symbol """
        if 'unicode' in wx.PlatformInfo:
            ch = unicode('\xc2\xb0', 'utf8')
            len = 2
            self.InsertUnicodeChar(ch, len)
        else:
            ch = chr(176)
            self.InsertSymbol(ch)

    def InsertSymbol(self, ch, length=1):
        """ Insert a character in Symbol Font into the text """
        # Find the cursor position
        curpos = self.GetCurrentPos()
        # Save the current font
        f = self.get_font()
        # Use symbol font
        self.set_font("Symbol", TransanaGlobal.configData.defaultFontSize, 0x000000)
        self.InsertStyledText(ch, length)
        # restore previous font
        self.set_font(f[0], f[1], f[2])
        # Position the cursor after the inserted character.  Unicode characters may have a length > 1, even though
        # they show up as only a single character.
        self.GotoPos(curpos + length)

    def InsertUnicodeChar(self, ch, length=1):
        """ Insert a Unicode character, which may not have a length of 1, into the text """
        # Save the current font.  If we are at the end of a document and we're not using the default font,
        # the font gets changed to the default font by this routine if we don't control for it!
        f = self.get_font()
        # Find the cursor position
        curpos = self.GetCurrentPos()
        # Insert the specified character at the cursor position
        self.InsertText(curpos, ch)
        # Unicode characters are wider than 1 character. Move the cursor the appropriate number of characters.
        self.GotoPos(curpos + length)
        # restore previous font
        self.set_font(f[0], f[1], f[2])

    def InsertHiddenText(self, text):
        """Insert hidden text at current cursor position."""
        # Determine the current cursor position
        curpos = self.GetCurrentPos()
        # Insert the text
        self.InsertText(curpos, text)
        # Determine how much to move the cursor.  Start with the length of the text.
        offset = len(text)
        # Iterate through the characters looking for Unicode characters
        for char in text:
            if ('unicode' in wx.PlatformInfo) and (ord(char) > 128):
                # I think we can get away with adding 1 here because the Time Code character, with a
                # width of 2, should be the only Unicode character that gets hidden.
                offset += 1
        self.StartStyling(curpos, STYLE_MASK)
        self.SetStyling(offset, self.STYLE_HIDDEN)
        self.GotoPos(curpos+offset)
        # Ensure that the OnPosChanged event isn't fooled into switching
        # to this new style
        self.last_pos = curpos + offset

    def InsertStyledText(self, text, length=0):
        """Insert text with the current style."""
        # Determine the length if needed
        if length == 0:
            if isinstance(text, unicode):
                length = len(text.encode(TransanaGlobal.encoding))  # Maybe 'utf8' instead of TransanaGlobal.encoding??
            else:
                length = len(text)
        # Determine the current cursor position
        curpos = self.GetCurrentPos()
        # Insert the text
        self.InsertText(curpos, text)
        # Determine how much to move the cursor.  Start with the length of the text.
        if length == 1:
            offset = len(text)
            # if we're using Unicode and we have a single character and it's a Unicode (double-width) character ...
            # Iterate through the characters looking for Unicode characters

            # NOTE:  This will only work for 2-byte Unicode characters, not 3- or 4-byte characters.
            #        Those characters MUST pass in a "length" value.
            for char in text:
                if ('unicode' in wx.PlatformInfo) and (ord(char) > 128):
                    offset += 1
        else:
            offset = length
        # Apply the desired style to the inserted text
        self.StartStyling(curpos, STYLE_MASK)
        self.SetStyling(offset, self.style)
        # Move the cursor to after the inserted text
        self.GotoPos(curpos+offset)

    def ClearDoc(self):
        """Clear the document buffer."""
        self.__ResetBuffer()

    def LoadRTFData(self, data):
        """Load a RTF document into the editor with the document as a
        buffer (string)."""
        # Signal that we ARE loading data, so that the problem with font specification is avoided.
        self.loadingData = True
        
        RichTextEditCtrl.ClearDoc(self) # Don't want to call any parent method too

        if DEBUG:
            print "RichTextEditCtrl.LoadRTFData()", type(data)

        try:
            parse = RTFParser.RTFParser()
            parse.init_progress_update(self.ProgressDlg.Update, 0, 50)
            parse.buf = data

            if DEBUG:
                print "RichTextEditCtrl.LoadRTFData():  calling parse.read_stream()"
                
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

            if DEBUG:
                print "RichTextEditCtrl.LoadRTFData():  calling self.__ParseRTFStream()"
                
            self.__ParseRTFStream(parse.stream)

        if DEBUG:
            print
            print "self.style_specs:"
            for i in range(len(self.style_specs)):
                print "%3d %s" % (i, self.style_specs[i])

        # Okay, we're done loading the data now.
        self.loadingData = False

    def InsertRTFText(self, text):
        """ Add RTF text at the cursor, without clearing the whole document.  Used to insert
            Clip Transcripts into TextReports. """
        # Start exception trapping
        try:
            # Create an instance of the RTF Parser
            parse = RTFParser.RTFParser()
            # Assign the RTF Text String to the Parser's buffer
            parse.buf = text
            # Process the RTF Stream
            parse.read_stream()
        # Handle exceptions if needed
        except:
            print "Unhandled/ignored exception in RTF parsing"
            traceback.print_exc()
        # If no exception occurs ...
        else:
            # ... Actually parse the RTF Stream and put the text in the RTFTextEditCtrl
            self.__ParseRTFStream(parse.stream)

    def LoadRTFFile(self, filename):
        """Load a RTF file into the editor."""
        RichTextEditCtrl.ClearDoc(self) # Don't want to call any parent method too

        if DEBUG:
            print "RichTextEditCtrl.LoadRTFFile()", type(data)

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
        
    def GetItalic(self):
        """Get italic state for current font or for the selected text."""
        return self.style_attrs[self.style].italic

    def GetUnderline(self):
        """Get underline state for current font or for the selected text."""
        return self.style_attrs[self.style].underline

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
        self.SetAttributeSkipTimeCodes("bold", state)
        if DEBUG:
            print "SetBold: now self.style = %d" % self.style
    
    def SetItalic(self, state=1):
        """Set italic state for current font or for the selected text."""
        self.SetAttributeSkipTimeCodes("italic", state)
        if DEBUG:
            print "SetItalic: now self.style = %d" % self.style

    def SetUnderline(self, state=1):
        """Set underline state for current font or for the selected text."""
        self.SetAttributeSkipTimeCodes("underline", state)
        if DEBUG:
            print "SetUnderline: now self.style = %d" % self.style

    def SetAttributeSkipTimeCodes(self, attribute, state):
        """ Set the supplied attribute for the current text selection, ignoring Time Codes and their hidden data """
        # Let's try to remember the cursor position
        self.SaveCursor()
        # If we have a selection ...
        if self.GetSelection()[0] != self.GetSelection()[1]:
            # Set the Wait cursor
            self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
            # Now we need to iterate through the selection and update the font information.
            # It doesn't work to try to apply formatting to the whole block, as time codes get spoiled.
            for selPos in range(self.GetSelection()[0], self.GetSelection()[1]):
                # We don't want to update the formatting of Time Codes or of hidden Time Code Data.  
                if not (self.GetStyleAt(selPos) in [self.STYLE_TIMECODE, self.STYLE_HIDDEN]):
                    # Select the character we want to work on from the larger selection
                    self.SetSelection(selPos, selPos + 1)
                    # Set the style appropriately
                    self.__SetAttr(attribute, state)
                    self.__ApplyStyle(self.style)
            # Set the cursor back to normal
            self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
        # If we do NOT have a selection ...
        else:
            # Set the style appropriately
            self.__SetAttr(attribute, state)
            self.__ApplyStyle(self.style)
        # Let's try restoring the Cursor Position when all is said and done.
        self.RestoreCursor()
        # Signal that the transcript has changed, so that the Save prompt will be displayed if this format
        # change is the only edit.
        self.stylechange = 1
        
    def SetFont(self, face, size, fg_color, bg_color):
        """Set the font."""
        
        if DEBUG:
            print "SetFont: face: %s, size:%d, color fg:#%06x, bg: #%06x" % (face, size, fg_color, bg_color)
            
        if len(face) > 0:
            self.__SetFontFace(face)
        self.__SetFontColor(fg_color, bg_color)
        self.__SetFontSize(size)
        
        if DEBUG:
            print "After SetATtr: Style=%d" % self.style
            
        self.__ApplyStyle(self.style)
        # This ensures that the OnPosChanged event doesn't mistakingly
        # change back the font.  Probably should also add this to the
        # SetBold/Italic/Underline?
        self.last_pos = self.GetCurrentPos()

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

        # print "RichTextEditCtrl.__ProcessDocAsRTF():  A ", self.GetCurrentPos(), self.GetSelectionStart(), self.GetSelectionEnd(), self.GetSelection()

        doc.start_paragraph("Default")
        
        if select_only:
            start = self.GetSelectionStart()
            if 'unicode' in wx.PlatformInfo:
                end = self.GetSelectionEnd()
            else:
                # You need the "-1" here to prevent an extra character from being included in the Clip Transcript.
                end = self.GetSelectionEnd() - 1

        else:
            start = 0
            if 'unicode' in wx.PlatformInfo:
                end = self.GetLength()
            else:
                # You need the "-1" here to prevent an extra line break character from being included in the Transcript.
                end = self.GetLength() - 1

        if 'unicode' in wx.PlatformInfo:
            # We need to get the styled text out of the STC, and it must be Unicode.
            # Unfortunately, there doesn't appear to be an EASY way to do this.

            # First, get the text.  This returns a string object, not a Unicode object.
            text = self.GetStyledText(start, end)

            # Create a blank Unicode string for building a Unicode version of the text
            uniText = unicode('', 'utf-8')

            # Because some Unicode characters are more than one byte wide, but the STC doesn't recognize this,
            # we will need to skip the processing of the later parts of multi-byte characters.  skipNext allows this.
            skipNext = 0
            # process each character in the StyledText.  The GetStyledText() call has returned a string
            # with the data in two-character chunks, the text char and the styling char.  This for loop
            # allows up to process these character pairs.
            for x in range(0, len(text) - 1, 2):
                # If we are looking at the second character of a Unicode character pair, we can skip
                # this processing, as it has already been handled.
                if skipNext > 0:
                    # We need to reset the skipNext flag so we won't skip too many characters.
                    skipNext -= 1
                else:
                    # Check for a Unicode character pair by looking to see if the first character is above 128
                    if ord(text[x]) > 127:
                        
                        if DEBUG:
                            print "RichTextEditCtrl.__ProcessDocasRTF(): ord(%s) > 128 (%d) ..." % (text[x], ord(text[x]))

                            print "x =", x
                            if text > 100:
                                print "text="
                                print text[x-100:x-50]
                                print text[x-49:x]
                                print " =>", text[x],"<=="
                                print text[x:x+50]
                                print text[x+51:x+100]

                        # UTF-8 characters are variable length.  We need to figure out the correct number of bytes.

                        # Note the current position
                        pos = x
                        # Initialize the final character variable
                        c = ''

                        # Begin processing of unicode characters, continue until we have a legal character.
                        while (pos < len(text) - 1):

                            if DEBUG:
                                print "RichTextEditCtrl.__ProcessDocasRTF(): '%s'" % text, pos, c
                            
                            # Add the current character to the character variable
                            c += text[pos]

                            if DEBUG:
                                print 'RichTextEditCtrl.__ProcessDocasRTF(): c = ', c, type(c)

                            # Because the Mac wxSTC doesn't display the up and down arrow characters properly, we need to
                            # substitute the appropriate characters here for the save.
                            if 'wxMac' in wx.PlatformInfo:
                                if (len(c) == 3) and (ord(c[0]) == 226) and (ord(c[1]) == 137) and (ord(c[2]) == 160):
                                    c = chr(194) + chr(173)
                                elif (len(c) == 2) and (ord(c[0]) == 195) and (ord(c[1]) == 152):
                                    c = chr(194) + chr(175)
				# We also need to substitute the time code character here
                                elif (len(c) == 2) and (ord(c[0]) == 194) and (ord(c[1]) == 167):
                                    c = chr(194) + chr(164)

                            # Correct for mal-formed time codes, which can occur with transcripts that come up
                            # through older versions of Transana.
                            if c == '\xa4':
                                c = chr(194) + chr(164)
                                
                            # Note the style of the current character.  Multi-byte characters should always have
                            # the same style.
                            style = ord(text[x + 1])
                            # Try to encode the character.
                            try:
                                # See if we have a legal character yet.
                                # NOTE:  We need to use UTF8 encoding, not the global encoding, as
                                # python, and therefore the wxSTC, uses UTF8 internally.
                                d = unicode(c, 'utf8')  # (c, TransanaGlobal.encoding)
                                # If so, break out of the while loop
                                break
                            # If we don't have a legal UTF-8 character, we'll get a UnicodeDecodeError exception
                            except UnicodeDecodeError:
                                # We need to signal the need to skip a charater in overall processing
                                skipNext += 1
                                # We need to update the current position and keep processing until we have a legal UTF-8 character
                                pos += 2

                        if DEBUG:
                            print "RichTextEditCtrl.__ProcessDocasRTF():  Unicode Charcter Processing done.\nc = '%s' type = %s" % (c, type(c))
                    else:
                        c = text[x]
                        style = ord(text[x + 1])

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

                        if DEBUG:
                            print "RichTextEditCtrl.__ProcessDocAsRTF(): calling write_text with UTF-8"
                            print "c = '%s' " % c,
                            for x in c:
                                print ord(x),
                            print

                            
                        # There is a situation where we end up with an incomplete Unicode character for the time code.
                        # Let's trap this error and ignore the character.
                        if (len(c) == 1) and (ord(c) == 194):
                            if DEBUG:
                                print
                                print "************************ TRAPPED IT ***************************************"
                                print
                            pass
                        else:
                            try:
                                doc.write_text(unicode(c, 'utf8'))
                            except UnicodeDecodeError:
                                doc.write_text(c)
                                msg = _("Encoding error during RTF Export.  Some transcript encoding may incorrect.")
                                errDlg = Dialogs.ErrorDialog(self, msg)
                                errDlg.ShowModal()
                                errDlg.Destroy()

                        if DEBUG:
                            print "RichTextEditCtrl.__ProcessDocAsRTF(): write_text called with UTF-8"
       
        else:
            text = self.GetText()

            if DEBUG:
                print "RichTextEditCtrl.__ProcessDocAsRTF(): text =", text

            x = 0
            for c in text:
                style = self.GetStyleAt(x)

                x += len(c)

                if DEBUG:
                    print "RichTextEditCtrl.__ProcessDocAsRTF(): c = %s, len(c) = %d, style = %d, x = %s" % (c, len(c), style, x)
                    
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
        # NOTE:  cStringIO is faster but does not support Unicode strings!!!   DKW
        if 'unicode' in wx.PlatformInfo:
            f = StringIO.StringIO()
        else:
            f = cStringIO.StringIO()
        doc = self.__NewRTFDoc(f)
        self.__ProcessDocAsRTF(doc, select_only)
        doc.close()
        f.seek(0)
        data = f.read()
        f.close()
        return data

    def GetPickledBuffer(self, select_only=0):
        """Get the current document as a pickled string. This is useful for
        saving to the database."""
 
	if DEBUG:
	    print "RichTextEditCtrl.GetPickledBuffer()"

	bufferContents = self.GetStyledText(0, self.GetLength())

	if DEBUG:
            print "bufferContents[:20] ="
            print bufferContents[:20]
            for x in bufferContents[:20]:
                print x, ord(x)
        
        # Due to wonderful unicode issues, the special characters are 
        # different between Windows and OSX. So what happens here is if we are
        # running on OSX, we must convert all the existing special 
        # characters into a character that wxSTC on Windows understands.
	# Sorry, Mac folks, you get the performance hit.
	if 'wxMac' in wx.PlatformInfo:
	    # this little dictionary contains Mac special character sequences
	    # along with their Windows counterparts.
	    sequenceList = {
		'\xc2\xa7':'\xc2\xa4',
		'\xe2\x89\xa0':'\xc2\xad',
		'\xc3\x98':'\xc2\xaf'
	    }
	    todo = {}
	    inSeq = False
	    gatherStyle = False
	    sequence = ''
	    # look for one of the Mac sequences specified in sequenceList
	    # and try to collect the "style character".
	    for char in bufferContents:
		if char == '\xe2' or char == '\xc3' or char == '\xc2':
		    inSeq = True
		    gatherStyle = True
		    sequence = char
		elif inSeq and gatherStyle:
		    styleChar = char
		    gatherStyle = False
		elif inSeq and not gatherStyle and char != styleChar:
		    sequence = sequence + char

		# we might have gathered a valid sequence, so now we need
		# to check it against the known sequences.
                if inSeq:
                    try:	
                        # It is important for the following statement to be first, because
                        # if the sequence is incomplete, we still need inSeq to be true.
                        replacement = sequenceList[sequence]
                        inSeq = False
                        # found a valid sequence, now insert the style character after
                        # each sequence character.
                        replacement = styleChar.join([x for x in replacement]) + styleChar
                        sequence = styleChar.join([x for x in sequence]) + styleChar
                        todo[sequence] = replacement
                    except KeyError:
                        pass
	
    	    # Go ahead and replace the existing sequences with their Windows compadres.
            for k, v in todo.iteritems():
                # The problem with Jonathan's logic above is that the style character PRECEEDS rather than 
                # follows the text character, so the final style character may differ from the one recorded above.
                # To fix that, let's strip the final style character from both strings.  DKW
                bufferContents = bufferContents.replace(k[:-1], v[:-1])
	
            if DEBUG:
                print "post-replacement bufferContents[:20] ="
                print bufferContents[:20]
                for x in bufferContents[:20]:
                    print x, ord(x)
        
        # Create a tuple containing the state of the buffer. If you don't 
        # include the style specs, you will not be able to correctly load the 
        # buffer again.
        data = (bufferContents, self.style_specs, self.style_attrs)

        # Since we're saving to a database server, we want the pickled data as
        # a string, rather than throwing it into some file.
        return pickle.dumps(data)

    def SaveRTFDocument(self, filename):
        """Save the document in memory to the given file."""
        doc = self.__NewRTFDoc(filename)
        self.__ProcessDocAsRTF(doc)
        doc.close()

    def SetDragEvent(self, id, func):
        stc.EVT_STC_START_DRAG(self, id, func)

    def OnCharAdded(self, event):
        """Called when a character is added to the document."""

        if TransanaConstants.demoVersion and (self.GetLength() > 10000):
            self.Undo()
            prompt = _("The Transana Demonstration limits the size of Transcripts.\nYou have reached the limit and cannot edit this transcript further.")
            tempDlg = Dialogs.InfoDialog(self, prompt)
            tempDlg.ShowModal()
            tempDlg.Destroy()
        
        if DEBUG:
            print "RichTextEditCtrl.OnCharAdded(): %s (style=%d)" % (event.GetKey(), self.style)
        # Don't do anything if in read-only mode
        if self.GetReadOnly():
            return
        # Unicode characters can have variable lengths.  We need to know the length of the character just entered.
        if 'unicode' in wx.PlatformInfo:
            # Get the Unicode character.
            tempChar = unichr(event.GetKey())
            # Determine the length of the character in the appropriate encoding.
            length = len(tempChar.encode(TransanaGlobal.encoding))
        # ANSI characters have a fixed length of 1
        else:
            length = 1
        # The current position is AFTER the character we've just entered.  Start styling at the beginning of that character.
        self.StartStyling(self.GetCurrentPos() - length, STYLE_MASK)
        # Apply the current style to the full length of the character just entered.
        self.SetStyling(length, self.style)

    def OnUpdateUI(self, event):
        """Called when the the User Interface needs to be updated."""
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

    def OnModified(self, event):
        """Triggered when the document is modified, including style changes."""
        # Note: No modifications may be performed when servicing this event!
        if event.GetModificationType() & wx.stc.STC_MOD_CHANGESTYLE:
            self.stylechange = 1

    # Remove Drag-and-Drop reference on the Mac due to the Quicktime Drag-Drop bug
    if not TransanaConstants.macDragDrop and ("__WXMAC__" in wx.PlatformInfo):
        def OnKillDrag(self, event):
            """ This method KILLS Drag-and-Drop functionality on the Mac. """
            # According to Robin Dunn, the way to kill Drag-and-Drop in the wxSTC is to blank out the Drag Text
            # He notes that this is an undocumented "hidden feature" of the START_DRAG event
            event.SetDragText("")

    def SetSavePoint(self):
        """Override the existing SetSavePoint to account for style changes."""
        wx.stc.StyledTextCtrl.SetSavePoint(self)
        self.stylechange = 0

    def GetModify(self):
        """Override the existing GetModify to account for style changes."""
        return wx.stc.StyledTextCtrl.GetModify(self) or self.stylechange
        
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
