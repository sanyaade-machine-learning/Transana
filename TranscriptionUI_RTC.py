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

""" This module implements the TranscriptionUI class for the Transcript Editor component of Transana. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "TranscriptionUI_RTC DEBUG is ON!!"

# import wxPython
import wx
# import the Splitter Window module
import wx.lib.splitter

# Define the Color for the Active Document when there are several Panels shown
ACTIVE_COLOR = wx.Colour(51, 153, 255)    
INACTIVE_COLOR = wx.WHITE

# import Python's gettext module
import gettext
# import Python's os module
import os
# import Python's sys module
import sys
# import Transana's Constants
import TransanaConstants
# import Transana's Exceptions
import TransanaExceptions
# Import Transana's Global variables
import TransanaGlobal
# Import the Transana Clip object
import Clip
# Import Transana's Dialogs
import Dialogs
# Import Transana's Document Object
import Document
# Import Transana's Episode Object
import Episode
# import Transana's Keyword List Edit Form
import KeywordListEditForm
# Import Transana's Quote Object
import Quote
# Import Transana's Images
import TransanaImages
# Import Transana's Transcript Object
import Transcript
# import Transana's Transcript Toolbar
## from TranscriptToolbar import TranscriptToolbar
# Import Transana's Transcript Editor for the wx.RichTextCtrl
import TranscriptEditor_RTC

import time
#print "TranscriptionUI_RTC -- import time"

# NOTE:  TrancriptionUI (self) is a FRAME.  It holds a TOOLBAR (self.toolbar) and a NOTEBOOK (self.nb).
#        Each page (self.nb.GetPage(x) or self.nb.GetCurrentPage()) of that NOTEBOOK holds a MultiSplitterWindow.
#        Each pane (self.nb.GetCurrentPage().GetChildren()) of that MultiSplitterWindow holds a TranscriptPanel.
#        The TranscriptPanel is a PANEL that holds a TranscriptEditor_RTC object.
#          (self.GetCurrentPage().GetChildren()[self.GetCurrentPage().activePanel].editor)
#        This allows multiple pages for multiple Documents and Transcripts, as
#        well as multiple adjustible panes on each page for multiple simultaneous
#        documents and transcripts.
#
#        self.dlg gets the nb.CurrentPage()'s activePanel automatically.  It's a property to retain compatibility with
#          pre-splitter code, when there were different Frames for different Editors.

class TranscriptionUI(wx.Frame):
    """This class manages the graphical user interface for the transcription
    editors component.  It creates the transcript window containing a
    TranscriptToolbar and one or more TranscriptEditor objects."""

    def __init__(self, parent, includeClose=False):
        """Initialize an TranscriptionUI object."""

        # The ControlObject handles all inter-object communication, initialized to None
        self.ControlObject = None

        # If we're including an optional Close button ...
        if includeClose:
            # ... define a style that includes the Close Box.  (System_Menu is required for Close to show on Windows in wxPython.)
            style = wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS | wx.SYSTEM_MENU | wx.CLOSE_BOX
        # If we don't need the close box ...
        else:
            # ... then we don't need that defined in the style
            style = wx.CAPTION | wx.RESIZE_BORDER | wx.WANTS_CHARS
        # Create the Frame with the appropriate style
        wx.Frame.__init__(self, parent, -1, _("Document"), self.__pos(), self.__size(), style=style)

        # if we're not on Linux ...
        if not 'wxGTK' in wx.PlatformInfo:
            # Set the Background Colour to the standard system background (not sure why this is necessary here.)
            self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_FRAMEBK))
        else:
            self.SetBackgroundColour(wx.WHITE)

        # Set "Window Variant" to small only for Mac to use small icons
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create a vertical Sizer for the Form
        vSizer = wx.BoxSizer(wx.VERTICAL)

##        # The Toolbar MUST be on a Panel, or the Search Box doesn't show up on OS X.
##        # So let's create a panel for it!!
##        self.toolbarPanel = _ToolbarPanel(self)

        # Define the Transcript Toolbar object
##        self.toolbar = TranscriptToolbar(self.toolbarPanel)
        self.toolbar = self.CreateToolBar(wx.TB_HORIZONTAL | wx.BORDER_SIMPLE | wx.TB_FLAT)

        # Set the Toolbar Bitmap size
        if 'wxMac' in wx.PlatformInfo:
            tbSize = (16, 16)
        else:
            tbSize = (16, 16)
        self.toolbar.SetToolBitmapSize(tbSize)
        # Keep a list of the tools placed on the toolbar so they're more easily manipulated
        self.tools = []

        # Create an Undo button
        self.CMD_UNDO_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_UNDO_ID, TransanaImages.Undo16.GetBitmap(),
                        shortHelpString=_('Undo action')))
        wx.EVT_MENU(self, self.CMD_UNDO_ID, self.OnUndo)

        if not 'wxMac' in wx.PlatformInfo:
            self.toolbar.AddSeparator()

        # Bold, Italic, Underline buttons
        self.CMD_BOLD_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_BOLD_ID, TransanaGlobal.GetImage(TransanaImages.Bold),
                        isToggle=1, shortHelpString=_('Bold text')))
        wx.EVT_MENU(self, self.CMD_BOLD_ID, self.OnBold)

        self.CMD_ITALIC_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_ITALIC_ID, TransanaGlobal.GetImage(TransanaImages.Italic),
                        isToggle=1, shortHelpString=_("Italic text")))
        wx.EVT_MENU(self, self.CMD_ITALIC_ID, self.OnItalic)
       
        self.CMD_UNDERLINE_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_UNDERLINE_ID, TransanaGlobal.GetImage(TransanaImages.Underline),
                        isToggle=1, shortHelpString=_("Underline text")))
        wx.EVT_MENU(self, self.CMD_UNDERLINE_ID, self.OnUnderline)

        if not 'wxMac' in wx.PlatformInfo:
            self.toolbar.AddSeparator()

        # Jeffersonian Symbols
        self.CMD_RISING_INT_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_RISING_INT_ID, TransanaImages.ArtProv_UP.GetBitmap(),
                        shortHelpString=_("Rising Intonation")))
        wx.EVT_MENU(self, self.CMD_RISING_INT_ID, self.OnInsertChar)
        
        self.CMD_FALLING_INT_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_FALLING_INT_ID, TransanaImages.ArtProv_DOWN.GetBitmap(),
                        shortHelpString=_("Falling Intonation")))
        wx.EVT_MENU(self, self.CMD_FALLING_INT_ID, self.OnInsertChar) 
       
        self.CMD_AUDIBLE_BREATH_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_AUDIBLE_BREATH_ID, TransanaGlobal.GetImage(TransanaImages.AudibleBreath),
                        shortHelpString=_("Audible Breath")))
        wx.EVT_MENU(self, self.CMD_AUDIBLE_BREATH_ID, self.OnInsertChar)
    
        self.CMD_WHISPERED_SPEECH_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_WHISPERED_SPEECH_ID, TransanaGlobal.GetImage(TransanaImages.WhisperedSpeech),
                        shortHelpString=_("Whispered Speech")))
        wx.EVT_MENU(self, self.CMD_WHISPERED_SPEECH_ID, self.OnInsertChar)
      
        if not 'wxMac' in wx.PlatformInfo:
            self.toolbar.AddSeparator()

        # Add show / hide timecodes button
        self.CMD_SHOWHIDE_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_SHOWHIDE_ID, TransanaGlobal.GetImage(TransanaImages.TimeCode16),
                        isToggle=1, shortHelpString=_("Show/Hide Time Code Indexes")))
        wx.EVT_MENU(self, self.CMD_SHOWHIDE_ID, self.OnShowHideCodes)

        # Add show / hide timecodes button
        self.CMD_SHOWHIDETIME_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_SHOWHIDETIME_ID, TransanaGlobal.GetImage(TransanaImages.TimeCodeData16),
                                       TransanaGlobal.GetImage(TransanaImages.TimeCodeData16),
                                       isToggle=1, shortHelpString=_("Show/Hide Time Code Values")))
        wx.EVT_MENU(self, self.CMD_SHOWHIDETIME_ID, self.OnShowHideValues)

        # Add read only / edit mode button
        self.CMD_READONLY_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_READONLY_ID, TransanaGlobal.GetImage(TransanaImages.ReadOnly16),
                                       TransanaGlobal.GetImage(TransanaImages.ReadOnly16),
                                       isToggle=1, shortHelpString=_("Edit/Read-only select")))
        wx.EVT_MENU(self, self.CMD_READONLY_ID, self.OnReadOnlySelect)

        # Add Formating button
        self.CMD_FORMAT_ID = wx.NewId()
        # ... and create a Format button on the tool bar.
        self.tools.append(self.toolbar.AddTool(self.CMD_FORMAT_ID, TransanaImages.ArtProv_HELPSETTINGS.GetBitmap(), shortHelpString=_("Format")))
        wx.EVT_MENU(self, self.CMD_FORMAT_ID, self.OnFormat)

        if not 'wxMac' in wx.PlatformInfo:
            self.toolbar.AddSeparator()

        # Add QuickClip button
        self.CMD_QUICKCLIP_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_QUICKCLIP_ID, TransanaGlobal.GetImage(TransanaImages.QuickClip16),
                        shortHelpString=_("Create Quick Quote or Clip")))
        wx.EVT_MENU(self, self.CMD_QUICKCLIP_ID, self.OnQuickClip)

        # Add Edit keywords button
        self.CMD_KEYWORD_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_KEYWORD_ID, TransanaGlobal.GetImage(TransanaImages.KeywordRoot16),
                        shortHelpString=_("Edit Keywords")))
        wx.EVT_MENU(self, self.CMD_KEYWORD_ID, self.OnEditKeywords)

        # Add Save Button
        self.CMD_SAVE_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_SAVE_ID, TransanaGlobal.GetImage(TransanaImages.Save16),
                        shortHelpString=_("Save Transcript")))
        wx.EVT_MENU(self, self.CMD_SAVE_ID, self.OnSave)

        if not 'wxMac' in wx.PlatformInfo:
            self.toolbar.AddSeparator()

        # Add Propagate Changes Button
        # First, define the ID for this button
        self.CMD_PROPAGATE_ID = wx.NewId()
        # Now create the button and add it to the Tools list
        self.tools.append(self.toolbar.AddTool(self.CMD_PROPAGATE_ID, TransanaGlobal.GetImage(TransanaImages.Propagate),
                        shortHelpString=_("Propagate Changes")))
        # Link the button to the appropriate event handler
        wx.EVT_MENU(self, self.CMD_PROPAGATE_ID, self.OnPropagate)

        if not 'wxMac' in wx.PlatformInfo:
            self.toolbar.AddSeparator()
        
        # Add Multi-Select Button
        # First, define the ID for this button
        self.CMD_MULTISELECT_ID = wx.NewId()
        # Now create the button and add it to the Tools list
        self.tools.append(self.toolbar.AddTool(self.CMD_MULTISELECT_ID, TransanaGlobal.GetImage(TransanaImages.MultiSelect),
                        shortHelpString=_("Match Selection in Other Transcripts")))
        # Link the button to the appropriate event handler
        wx.EVT_MENU(self, self.CMD_MULTISELECT_ID, self.OnMultiSelect)

        # Add Multiple Transcript Play Button
        # First, define the ID for this button
        self.CMD_PLAY_ID = wx.NewId()
        # Now create the button and add it to the Tools list
        self.tools.append(self.toolbar.AddTool(self.CMD_PLAY_ID, TransanaImages.Play.GetBitmap(),
                        shortHelpString=_("Play Multiple Transcript Selection")))
        # Link the button to the appropriate event handler
        wx.EVT_MENU(self, self.CMD_PLAY_ID, self.OnMultiPlay)

        if not 'wxMac' in wx.PlatformInfo:
            self.toolbar.AddSeparator()

        # SEARCH moved to TranscriptionUI because you can't put a TextCtrl on a Toolbar on the Mac!
        # Set the Initial State of the Editing Buttons to "False"
        for x in (self.CMD_UNDO_ID, self.CMD_BOLD_ID, self.CMD_ITALIC_ID, self.CMD_UNDERLINE_ID, \
                  self.CMD_RISING_INT_ID, self.CMD_FALLING_INT_ID, \
                  self.CMD_AUDIBLE_BREATH_ID, self.CMD_WHISPERED_SPEECH_ID, self.CMD_FORMAT_ID, \
                  self.CMD_PROPAGATE_ID, self.CMD_MULTISELECT_ID, self.CMD_PLAY_ID):
            self.toolbar.EnableTool(x, False)

        # Add Quick Search tools
        # Start with the Search Backwards button
        self.CMD_SEARCH_BACK_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_SEARCH_BACK_ID, TransanaImages.ArtProv_BACK.GetBitmap(),
                        shortHelpString=_("Search backwards")))

        wx.EVT_MENU(self, self.CMD_SEARCH_BACK_ID, self.OnSearch)

        # Add the Search Text box
        self.searchText = wx.TextCtrl(self.toolbar, -1, size=(100, 20), style=wx.TE_PROCESS_ENTER)
        self.toolbar.AddControl(self.searchText)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnSearch, self.searchText)

        # Add the Search Forwards button
        self.CMD_SEARCH_NEXT_ID = wx.NewId()
        self.tools.append(self.toolbar.AddTool(self.CMD_SEARCH_NEXT_ID, TransanaImages.ArtProv_FORWARD.GetBitmap(),
                        shortHelpString=_("Search forwards")))
        wx.EVT_MENU(self, self.CMD_SEARCH_NEXT_ID, self.OnSearch)

        if not 'wxMac' in wx.PlatformInfo:
            self.toolbar.AddSeparator()

        self.CMD_EXIT_ID = wx.NewId()
        # Add an Exit button to the Toolbar
        self.tools.append(self.toolbar.AddTool(self.CMD_EXIT_ID, TransanaImages.Exit.GetBitmap(), shortHelpString=_('Close Current')))
        # Link the button to the appropriate event handler
        wx.EVT_MENU(self, self.CMD_EXIT_ID, self.OnCloseCurrent)

        # Add the Selection Label, which indicates the time position of the current selection
        self.selectionText = wx.StaticText(self.toolbar, -1, "", size=wx.Size(200, 20))
        self.toolbar.AddControl(self.selectionText)

        # Call toolbar.Realize() to initialize the toolbar            
        self.toolbar.Realize()

        hSizer = wx.BoxSizer(wx.HORIZONTAL)
##        hSizer.Add(self.toolbarPanel, 1, wx.EXPAND)

        # Add the Toolbar to the Sizer
        vSizer.Add(hSizer, 0, wx.EXPAND)

        # add a Notebook Control to the Dialog Box
        self.nb = TranscriptionNotebook(self, _('(No Document Loaded)'))
        # Set the Notebook's background to White.  Otherwise, we get a visual anomoly on OS X with wxPython 2.9.4.0.
        self.nb.SetBackgroundColour(wx.Colour(255, 255, 255))
        
        # Let the notebook remember it's parent
        self.nb.parent = self
        vSizer.Add(self.nb, 1, wx.EXPAND | wx.TOP, 3)

        # Disable the Frame's toolbar
        self.toolbar.Enable(0)

        # Activate the Frame's Sizers
        self.SetSizer(vSizer)
        self.SetAutoLayout(True)
        self.Layout()

        # AFTER the frame is shown, we need to resize all the Splitter Panels on all Notebook Pages
        # For each Notebook Page ...
        for x in range(self.nb.GetPageCount()):
            # ... get the Notebook Page's Splitter object ...
            splitter = self.nb.GetPage(x)
            # ... and space the Panels on it evenly
            splitter.SpacePanelsEvenly()

        try:
            # Define the EVT_SIZE event, as we have to adjust the Toolbar Panel manually
            self.Bind(wx.EVT_SIZE, self.OnSize)
            # Define the Close event (for THIS Transcript Window)
            self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)

        except:
            print "TranscriptionUI._TranscriptPanel.__init__():"
            print sys.exc_info()[0]
            print sys.exc_info()[1]
            print

        # Set the FOCUS in the Editor.  (Required on Mac so that CommonKeys work.)
        wx.CallAfter(self.dlg.editor.SetFocus)
        
    # Public methods
    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        # Assign the passed-in control object to the dlg's ControlObject property
        self.ControlObject=ControlObject
        # Register the ControlObject with the Toolbar Panel as well
##        self.toolbarPanel.ControlObject = self.ControlObject
        # Iterate through all defined pages/tabs
        for pageNum in range(self.nb.GetPageCount()):
            # Iterate through the Panes on each Page
            for child in self.nb.GetPage(pageNum).GetChildren():
                # ... and assign the Control Object to each!
                child.ControlObject = ControlObject

    def CloseCurrent(self, event):
        """ Close the Current Notebook Page / Splitter Pane """
        # Stop the Line Numbers timer
        self.dlg.LineNumTimer.Stop()
        # Clear the Line Numbers
        self.dlg.ClearLineNum()
        self.dlg.EditorPaint(None)

        # Get a pointer to the Current Object
        tmpObj = self.GetCurrentObject()
        # If the tab points to a Transcript ...
        if isinstance(tmpObj, Transcript.Transcript):
            # Remove the Visualization Info for the item being closed
            self.ControlObject.DeleteVisualizationInfo((type(tmpObj), tmpObj.number))
        # If there are multiple Splitter Panes on this Notebook Page ...
        if len(self.nb.GetCurrentPage().GetChildren()) > 1:
            # Update the ControlObject's ActivePanel
            self.ControlObject.activeTranscript = self.nb.GetCurrentPage().activePanel
            
            # Save the changes, if needed
            self.ControlObject.SaveTranscript(1, transcriptToSave=self.nb.GetCurrentPage().activePanel)
            # If the Transcript is locked ...
            if self.dlg.editor.TranscriptObj != None and self.dlg.editor.TranscriptObj.isLocked:
                # ... unlock it
                self.dlg.editor.TranscriptObj.unlock_record()
            # ... then we can just close THIS pane!
            self.DeletePanel()
            # If we only have one Pane left ...
            if len(self.nb.GetCurrentPage().GetChildren()) == 1:
                # ... we need to disable the MultiTranscript Buttons
                self.UpdateMultiTranscriptButtons(False)
        # If there's one Pane but more than one Notebook Pages ...
        elif self.nb.GetPageCount() > 1:
            # Save the changes, if needed
            self.ControlObject.SaveTranscript(1, transcriptToSave=0)
            # If the Transcript is locked ...
            if self.dlg.editor.TranscriptObj != None and self.dlg.editor.TranscriptObj.isLocked:
                # ... unlock it
                self.dlg.editor.TranscriptObj.unlock_record()
            if tmpObj != None:
                # Remove the Visualization Info for the item being closed
                self.ControlObject.DeleteVisualizationInfo((type(tmpObj), tmpObj.number))
                # If we're closing the Transcript for a media-based object ... 
                if isinstance(tmpObj, Transcript.Transcript):
                    # ... clear the media file out of the media window (since there can only be one)
                    self.ControlObject.ClearMediaWindow()
            
            # ... delete the Notebook Page
            self.DeleteNotebookPage(event)
            # Other Window interaction is handled automatically as the Notebook Page changes on this page closing!!
        # If there's only one page with only one pane ...
        else:
            # Save the changes, if needed
            self.ControlObject.SaveTranscript(1, transcriptToSave=0)
            # If the Transcript is locked ...
            if self.dlg.editor.TranscriptObj != None and self.dlg.editor.TranscriptObj.isLocked:
                # ... unlock it
                self.dlg.editor.TranscriptObj.unlock_record()
            # Remove the Visualization Info for the item being closed, if there is one
            if tmpObj != None:
                self.ControlObject.DeleteVisualizationInfo((type(tmpObj), tmpObj.number))
            # Clear the current Document but leave the Control in place.
            self.ClearDoc()
            # Clear the Visualization Window
            self.ControlObject.ClearVisualization()
            # Clear the Media Window
            self.ControlObject.ClearMediaWindow()

    def OnSize(self, event):
        """ Resize method for Transcription UI.  Toolbar requires it. """
        # Get the Panel's size
#        (w1, h1) = self.toolbarPanel.GetSize()
        # Get the Frame's size
        (w2, h2) = self.GetSize()
        # Adjust the Toolbar to the Frame's with, but retaining the panel's height
        # OS X requires a slightly wider width.  Dunno why.
#        if 'wxMac' in wx.PlatformInfo:
#            self.toolbar.SetSize((w2 + 4, h1 + 4))
#        else:
#            self.toolbar.SetSize((w2, h1))
        # Call layout so the Sizers can do their magic
        self.Layout()

    def OnCloseWindow(self, event):
        """ Event for the Transcript Window Close button, which should only exist on secondary transcript windows """
        # Stop the Line Numbers timer
        self.dlg.LineNumTimer.Stop()
        # Clear the Line Numbers
        self.dlg.ClearLineNum()
        # Make the line numbers actually disappear
        self.dlg.EditorPaint(None)
        # Clear all the Transcript Windows (to make sure everything has been properly saved!)
        self.ControlObject.ClearAllWindows(clearAllTabs = True)

    def get_editor(self):
        """Get a reference to its TranscriptEditor object."""
        return self.dlg.editor
    
    def get_toolbar(self):
        """Get a reference to its TranscriptToolbar object."""
        return self.toolbar

    def ClearDoc(self):
        """Clear the Transcript window, unload any transcript."""
        # Clear the Transcription Editor
        self.dlg.editor.ClearDoc()
        # Update the Notebook Page (tab) header
        self.nb.SetPageText(self.nb.GetSelection(), _('(No Document Loaded)'))
        # Set the Transcript Window Title
        self.SetTitle(_('Document'))
        # Clear the Line Numbers
        self.dlg.ClearLineNum()
        self.dlg.EditorPaint(None)
        # Reset the Toolbar
        self.ClearToolbar()
        # Disable the toolbar
        self.toolbar.Enable(False)
        # Disable the Search
        self.dlg.EnableSearch(False)
        # If Line Numbers have been disabled ...
        if not self.dlg.lineNum.IsShown():
            # ... then make them re-appear!
            self.dlg.lineNum.Show(True)
            self.dlg.Layout()
            self.dlg.showLineNumbers = True

    def LoadTranscript(self, transcriptObj):
        """Load a transcript object."""
        if self.nb.GetCurrentPage().activePanel == 0:
            # Determine the appropriate text for the Notebook Tab label ...
            if isinstance(transcriptObj, Transcript.Transcript) and transcriptObj.clip_num > 0:
                tmpClip = Clip.Clip(transcriptObj.clip_num)
                pageLbl = tmpClip.id
            else:
                pageLbl = transcriptObj.id
            # Set the Notebook Tab Text            
            self.nb.SetPageText(self.nb.GetSelection(), pageLbl)

        # Transcripts should always be loaded in a Read-Only editor.  Set the Editor to Read Only.
        # This triggers a save prompt if the current transcript needs to be saved.
        self.dlg.editor.set_read_only(True)
        # Clear the toolbar
        self.ClearToolbar()

        # Load the transcript
        self.dlg.editor.load_transcript(transcriptObj)

        # if we're in the Demo version ...
        if TransanaConstants.demoVersion:

            # ... and if the Document / Transcript length is over 10,000 characters ...
            if self.dlg.editor.GetLength() > 10000:
                # ... go into Edit mode ...
                self.dlg.editor.SetReadOnly(False)
                # ... select everything beyond the 10,000th character ...
                self.dlg.editor.SetSelection(9999, self.dlg.editor.GetLength() - 1)
                # ... delete it from the control ...
                self.dlg.editor.DeleteSelection()
                # ... go back to Read Only mode ...
                self.dlg.editor.SetReadOnly(True)
                # ... and forget we've made any edits!
                self.dlg.editor.DiscardEdits()

                # Tell the user we've truncated the document and that they shouldn't save!
                msg = _("The Transana Demonstration limits the size of Documents and Transcripts.\nThis document has been truncated.  You may lose data if you try to edit it.")
                dlg = Dialogs.InfoDialog(self.dlg, msg)
                dlg.ShowModal()
                dlg.Destroy()

        # Enable the Toolbar
        self.dlg.toolbar.Enable(True)
        # Enable the Search
        self.dlg.EnableSearch(True)

    def UpdateGUI(self):
        """ This method should handle updating the GUI based on what data object is selected in the TranscriptWindow
            infrastructure """
        # Don't update the GUI if this is None.  This occurs as we are loading multiple transcripts, and causes problems.
        if self.nb.GetCurrentPage().GetChildren()[self.nb.GetCurrentPage().activePanel].editor.TranscriptObj <> None:
            # Let the Control Object know about the change (to update Data Window tabs, for one thing.)
            self.ControlObject.UpdateCurrentObject(self.nb.GetCurrentPage().GetChildren()[self.nb.GetCurrentPage().activePanel].editor.TranscriptObj)

    def GetCurrentTranscriptObject(self):
        """ Return the current Transcript Object, with the edited text even if it hasn't been saved. """
        # Make a copy of the Transcript Object, since we're going to be changing it.
        tempTranscriptObj = self.dlg.editor.TranscriptObj.duplicate()
        # Update the Transcript Object's text to reflect the edited state
        # STORE XML IN THE TEXT FIELD  (This shouldn't be necessary, but the time codes don't show up without it!)
        tempTranscriptObj.text = self.dlg.editor.GetFormattedSelection('XML')
        # Now return the copy of the Transcript Object
        return tempTranscriptObj

    def OnSearch(self, event):
        """ Pass-through event handler.  We don't know WHICH Document or Transcript to search in when the Toolbar is
            created, so this gives a place for the Toolbar to go which will figure out what to search when called! """
        # Pass through to the current Transcript Dialog's OnSearch method
        self.dlg.OnSearch(event)
        
    def GetDimensions(self):
        """ Get the position and size of the current TranscriptionUI window """
        # Get current Position information
        (left, top) = self.GetPositionTuple()  # .dlg
        # Get current Size information
        (width, height) = self.GetSizeTuple()   # .dlg
        # Return the values in a tuple
        return (left, top, width, height)

    def GetTranscriptDims(self):
        """Return dimensions of transcript editor component."""
        # Get current Position information
        (left, top) = self.dlg.editor.GetPositionTuple()
        # Get current Size information
        (width, height) = self.dlg.editor.GetSizeTuple()
        # Return the values in a tuple
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
        """ Set the position and size of the current TranscriptionUI window """
        self.SetDimensions(left, top, width, height)  # .dlg

    def UpdatePosition(self, positionMS):
        """Update the Transcript position given a time in milliseconds."""
        # Don't do anything unless auto word-tracking is enabled AND
        # edit mode is disabled.
        if TransanaGlobal.configData.wordTracking:
            # Assume we will NOT have movement
            returnVal = False
            # For each Editor Panel on the Current Notebook Tab ...
            for child in self.nb.GetCurrentPage().GetChildren():
                # ... if the editor is NOT in Edit Mode ...
                if child.editor.get_read_only():
                    # ... request the position update, changing the return value IF it moves something!
                    returnVal = returnVal or child.editor.scroll_to_time(positionMS)
            # Pass the return value (if ANY editor moved) to the calling routine
            return returnVal
        else:
            # Return true so that database tree tab updating that depends
            # on this return value still works (see ControlObject)
            return 1

    def InsertText(self, text):
        """Insert text at the current cursor position."""
        # Pass through to the Editor
        self.dlg.editor.InsertStyledText(text)
        
    def InsertTimeCode(self):
        """Insert a timecode at the current transcript position."""
        # Pass through to the Editor
        self.dlg.editor.insert_timecode()

    def InsertSelectionTimeCode(self, start_ms, end_ms):
        """Insert a timecode for the currently selected period in the
        Waveform.  start_ms and end_ms should contain the start and
        end time positions of the selected time period, in milliseconds."""
        # Pass through to the Editor
        self.dlg.editor.insert_timed_pause(start_ms, end_ms)

    def SetReadOnly(self, readOnly=True):
        """ Change the Read Only / Edit Mode status of a transcript """

        # ** THIS DOESN'T APPEAR TO EVER GET CALLED.  The RichTextEditCtrl_RTC method gets called instead! **
        
        # Change the Read Only Button status in the Toolbar
        self.dlg.toolbar.ToggleTool(self.CMD_READONLY_ID, not readOnly)
        # Fire the Read Only Button's event, like the button had been pressed manually.
        self.dlg.toolbar.OnReadOnlySelect(None)

    def TranscriptModified(self):
        """Return TRUE if transcript was modified since last save."""
        # Pass through to the Editor
        return self.dlg.editor.modified()
        
    def SaveTranscript(self, continueEditing=True):
        """Save the Transcript to the database.
           continueEditing is only used for Partial Transcript Editing."""
        # Pass through to the Editor
        if TransanaConstants.partialTranscriptEdit:
            self.dlg.editor.save_transcript(continueEditing)
        else:
            self.dlg.editor.save_transcript()

    def SaveTranscriptAs(self, fname):
        """Export the Transcript to an RTF or XML file."""
        # Pass through to the Editor
        self.dlg.editor.export_transcript(fname)
        # If saving an RTF file on a Mac, print a user warning
        if False and (fname[-4:].lower() == '.rtf') and ("__WXMAC__" in wx.PlatformInfo):
            msg = _('If you load this RTF file into Word on the Macintosh, you need to select "Format" > "AutoFormat...",\nmake sure the "AutoFormat now" option is selected, and press "OK".  Otherwise you will\nlose some Font formatting information from the file when you save it.\n(Courier New font will be changed to Times font anyway.)')
            msg = msg + '\n\n' + _('Also, Word on the Macintosh appears to handle the Whisper (Open Dot) Character for Jeffersonian \nNotation improperly.  You will need to convert this character to Symbol font within Word, but \nconvert it back to Courier New font prior to re-import into Transana.')
            dlg = Dialogs.InfoDialog(self.dlg, msg)
            dlg.ShowModal()
            dlg.Destroy()
        
    def TranscriptUndo(self, event):
        """ Make Transcript Undo available outside the TranscriptWindow """
        # Implement this by emulating an Undo press in the tool bar
        self.get_editor().OnUndo(event)

    def TranscriptCut(self, event):
        """  Pass-through for the Cut() method """
        # Pass through to the Editor
        self.get_editor().OnCutCopy(event)

    def TranscriptCopy(self, event):
        """  Pass-through for the Copy() method """
        # Pass through to the Editor
        self.get_editor().OnCutCopy(event)

    def TranscriptPaste(self, event):
        """  Pass-through for the Paste() method """
        # Pass through to the Editor
        self.get_editor().OnPaste(event)

    def CallFormatDialog(self, tabToOpen=0):
        """  Pass-through for the CallFormatDialog() method """
        # Pass through to the Editor
        self.dlg.editor.CallFormatDialog(tabToOpen)

    def InsertHyperlink(self, linkType, objNum):
        """ Insert a Hyperlink into a Document or Transcript """
        # If the transcript is NOT read-only ... (needed for Snapshot auto-insertion)
        if not self.dlg.editor.get_read_only():
            # ... then signal the Editor to insert an image
            self.dlg.editor.InsertHyperlink(linkType, objNum)

    def InsertImage(self, fileName = None, snapshotNum = -1):
        """ Insert an Image into the Transcript """
        # If the transcript is NOT read-only ... (needed for Snapshot auto-insertion)
        if not self.dlg.editor.get_read_only():
            # ... then signal the Editor to insert an image
            self.dlg.editor.InsertImage(fileName, snapshotNum)

    def AdjustIndexes(self, adjustmentAmount):
        """ Adjust Transcript Time Codes by the specified amount """
        # Check to see if Human Readable Time Code Values are displayed
        tcHRValues = self.dlg.editor.timeCodeDataVisible
        # If Human Readable Time Code Values are visible ...
        if tcHRValues:
            # ... hide them!
            self.dlg.editor.changeTimeCodeValueStatus(False)
        # If the transcript is "Read-Only", it must be put into "Edit" mode.  Let's do
        # this automatically.  ** THIS NEVER HAPPENS, AS THE MENU IS ONLY ENABLED IN EDIT MODE **
        if self.dlg.editor.get_read_only():
            # First, we "push" the Edit Mode button ...
            self.dlg.toolbar.ToggleTool(self.CMD_READONLY_ID, True)
            # ... then we call the button's method as if it really had been pushed.
            self.dlg.toolbar.OnReadOnlySelect(None)
        # Now adjust the indexes
        self.dlg.editor.AdjustIndexes(adjustmentAmount)
        # If Human Readable Time Code Values were visible ...
        if tcHRValues:
            # ... show them again!
            self.dlg.editor.changeTimeCodeValueStatus(True)

    def TextTimeCodeConversion(self):
        """ Convert Text (H:MM:SS.hh) Time Codes to Transana's format """
        # Call the Editor's Text Time Code Conversion Method
        self.dlg.editor.TextTimeCodeConversion()

    def UpdateSelectionText(self, text):
        """ Update the text indicating the start and end points of the current selection """
        # Pass through to the Dialog
        self.dlg.UpdateSelectionText(text)

    def ChangeLanguages(self):
        """ Change all on-screen prompts to the new language. """
        # Instruct the Toolbar to change languages
#        self.dlg.toolbar.ChangeLanguages()
        # Update the Speed Button Tool Tips
        self.toolbar.SetToolShortHelp(self.CMD_UNDO_ID, _('Undo action'))
        self.toolbar.SetToolShortHelp(self.CMD_BOLD_ID, _('Bold text'))
        self.toolbar.SetToolShortHelp(self.CMD_ITALIC_ID, _("Italic text"))
        self.toolbar.SetToolShortHelp(self.CMD_UNDERLINE_ID, _("Underline text"))
        self.toolbar.SetToolShortHelp(self.CMD_RISING_INT_ID, _("Rising Intonation"))
        self.toolbar.SetToolShortHelp(self.CMD_FALLING_INT_ID, _("Falling Intonation"))
        self.toolbar.SetToolShortHelp(self.CMD_AUDIBLE_BREATH_ID, _("Audible Breath"))
        self.toolbar.SetToolShortHelp(self.CMD_WHISPERED_SPEECH_ID, _("Whispered Speech"))
        self.toolbar.SetToolShortHelp(self.CMD_SHOWHIDE_ID, _("Show/Hide Time Code Indexes"))
        self.toolbar.SetToolShortHelp(self.CMD_SHOWHIDETIME_ID, _("Show/Hide Time Code Values"))
        self.toolbar.SetToolShortHelp(self.CMD_READONLY_ID, _("Edit/Read-only select"))
        self.toolbar.SetToolShortHelp(self.CMD_FORMAT_ID, _("Format"))
        self.toolbar.SetToolShortHelp(self.CMD_QUICKCLIP_ID, _("Create Quick Clip"))
        self.toolbar.SetToolShortHelp(self.CMD_KEYWORD_ID, _("Edit Keywords"))
        self.toolbar.SetToolShortHelp(self.CMD_SAVE_ID, _("Save Transcript"))
        self.toolbar.SetToolShortHelp(self.CMD_PROPAGATE_ID, _("Propagate Changes"))
        self.toolbar.SetToolShortHelp(self.CMD_MULTISELECT_ID, _("Match Selection in Other Transcripts"))
        self.toolbar.SetToolShortHelp(self.CMD_PLAY_ID, _("Play Multiple Transcript Selection"))
        self.toolbar.SetToolShortHelp(self.CMD_SEARCH_BACK_ID, _("Search backwards"))
        self.toolbar.SetToolShortHelp(self.CMD_SEARCH_NEXT_ID, _("Search forwards"))
        # Instruct the Dialog (Editor) to change languages
        self.dlg.ChangeLanguages()

    def GetNewRect(self):
        """ Get (X, Y, W, H) for initial positioning """
        pos = self.__pos()
        size = self.__size()
        return (pos[0], pos[1], size[0], size[1])
        
    def ChangeOrientation(self, event):
        """ Change the Orientation of the Splitter Panes """
        # Pass the orientation change request on to the Notebook to pass on to the Splitter
        self.nb.ChangeOrientation()
        
    def AddNotebookPage(self, pageTitle):
        """ Add a Notebook Page """
        # The new Notebook Page (Tab) will be given 5 panels in pre-defined colors
        self.nb.AddNotebookPage(pageTitle)

    def DeleteNotebookPage(self, event):
        """ Delete a Notebook Page """
        # If there's more than one Notebook Page ...
        if self.nb.GetPageCount() > 1:
            # ... delete the current Page
            self.nb.DeleteNotebookPage(self.nb.GetSelection())

    def AddPanel(self):
        """ Add a Splitter Pane to the curently-selected Notebook Tab """
        # Tell the Notebook we want to add a Panel
        self.nb.AddPanel()

    def DeletePanel(self, tab=-1, panel=-1):
        """ Delete the specified Splitter Pane from the specified Notebook Tab.
            if tab or panel is "-1", the current selection will be used.  """
        # Get the correct Notebook Page if one's not passed in
        if (tab <= -1) or (tab > self.nb.GetPageCount() - 1):
            tab = self.nb.GetSelection()
        if (panel <= -1) or (panel > len(self.nb.GetPage(tab).GetChildren()) - 1):
            panel = self.nb.GetPage(tab).activePanel
        # If the current Notebook Page has moer than ONE Splitter Pane ...
        if len(self.nb.GetPage(tab).GetChildren()) > 1:
            # ... tell the Notebook to delete the specified Panel
            self.nb.DeletePanel(tab, panel)
        # If the current Notebook Page has more than one Notebook Page ...
        elif self.nb.GetPageCount() > 1:
            # ... then Delete the Notebook Page!
            self.nb.DeletePage(tab)
        # If there's only one Notebook Page and only one Splitter Panel ...
        else:
            # Clear it!
            self.ClearDoc()

    def GetCurrentObject(self):
        """ Get the underlying OBJECT for the currently-selected text panel.  Could be a Document, a Transcript, a Quote,
            or None. """
        return self.nb.GetCurrentPage().GetChildren()[self.nb.GetCurrentPage().activePanel].editor.TranscriptObj

    def BringTranscriptToFront(self):
        """ The Document Window can contain MANY document tabs, but only one Transcript tab.  Find the Transcript and bring
            it to the front. """
        # For each Page in the Notebook ...
        for tabNum in range(self.nb.GetPageCount()):
            # ... if the text contained is a Transcript ...
            if isinstance(self.nb.GetPage(tabNum).GetChildren()[0].editor.TranscriptObj, Transcript.Transcript):
                # ... then select THIS page ...
                self.nb.SetSelection(tabNum)
                # ... and stop looking
                break

    # Toolbar Methods

    def ClearToolbar(self):
        """Clear buttons to default state."""
        # Reset toggle buttons to OFF.  This does not cause any events
        # to be emitted (only affects GUI state)
        self.toolbar.ToggleTool(self.CMD_BOLD_ID, False)
        self.toolbar.ToggleTool(self.CMD_ITALIC_ID, False)
        self.toolbar.ToggleTool(self.CMD_UNDERLINE_ID, False)
        self.toolbar.ToggleTool(self.CMD_READONLY_ID, False)
        self.toolbar.ToggleTool(self.CMD_SHOWHIDE_ID, False)
        self.toolbar.ToggleTool(self.CMD_SHOWHIDETIME_ID, False)
        self.UpdateEditingButtons()
        # Clear the Search Text
        self.dlg.ClearSearch()
        
    def OnUndo(self, evt):
        """ Implement Undo Button """
        self.dlg.editor.undo()

    def OnBold(self, evt):
        """ Implement Bold Button """
        bold_state = self.toolbar.GetToolState(self.CMD_BOLD_ID)
        self.dlg.editor.set_bold(bold_state)
        
    def OnItalic(self, evt):
        """ Implement Italics Button """
        italic_state = self.toolbar.GetToolState(self.CMD_ITALIC_ID)
        self.dlg.editor.set_italic(italic_state)
 
    def OnUnderline(self, evt):
        """ Implement Underline Button """
        underline_state = self.toolbar.GetToolState(self.CMD_UNDERLINE_ID)
        self.dlg.editor.set_underline(underline_state)

    def OnInsertChar(self, evt):
        """ Insert Jeffersonian Special Character """
        # Determine which button was pressed
        idVal = evt.GetId()
        # Rising Intonation
        if idVal == self.CMD_RISING_INT_ID:
            # Insert the Up Arrow / Rising Intonation symbol
            self.dlg.editor.InsertRisingIntonation()
        # Falling Intonation
        elif idVal == self.CMD_FALLING_INT_ID:
            # Insert the Down Arrow / Falling Intonation symbol
            self.dlg.editor.InsertFallingIntonation()
        # Audible Breath
        elif idVal == self.CMD_AUDIBLE_BREATH_ID:
            # Insert the High Dot / Inbreath symbol
            self.dlg.editor.InsertInBreath()
        # Whisper
        elif idVal == self.CMD_WHISPERED_SPEECH_ID:
            # Insert the Open Dot / Whispered Speech symbol
            self.dlg.editor.InsertWhisper()
        
    def OnShowHideCodes(self, evt):
        """ Implement Show / Hide Button """
        # Get the button's "indent" state
        show_codes = self.toolbar.GetToolState(self.CMD_SHOWHIDE_ID)
        # Depending on the indent state, show or hide the time codes
        if show_codes:
            self.dlg.editor.show_codes(showPopup=True)
        else:
            self.dlg.editor.hide_codes(showPopup=True)

    def OnShowHideValues(self, event):
        """ Implement Show / Hide Time Code Values """
        self.dlg.editor.show_timecodevalues(self.toolbar.GetToolState(self.CMD_SHOWHIDETIME_ID))

    def OnReadOnlySelect(self, evt):
        """ Implement Read Only / Edit Mode """
        # If there's no object currently loaded in the interface ...
        if self.dlg.editor.TranscriptObj == None:
            # ... then we can't edit it, can we?  Let's get out of here.
            return

##        print "TranscriptionUI_RTC.OnReadOnlySelect():"
##        print "  ", self.dlg.editor.TranscriptObj.id, self.dlg.editor.get_read_only()
##        print "  ", type(self), type(self.dlg), self.dlg.pageNum, self.dlg.panelNum, self.dlg.parent.activePanel
##        print
        
        # Get the button's "indent" state
        can_edit = self.toolbar.GetToolState(self.CMD_READONLY_ID)
        # If leaving edit mode, prompt for save if necessary.
        if not can_edit:
            if ((TransanaConstants.partialTranscriptEdit) and \
                (not self.ControlObject.SaveTranscript(1, transcriptToSave=self.nb.GetCurrentPage().activePanel, continueEditing=False))) or \
               ((not TransanaConstants.partialTranscriptEdit) and \
                (not self.ControlObject.SaveTranscript(1, transcriptToSave=self.nb.GetCurrentPage().activePanel))):
                # Reset the Toolbar
                self.ClearToolbar()
                # User chose to not save, revert back to database version
                tobj = self.dlg.editor.TranscriptObj
                # If our object is a DOCUMENT ...
                if isinstance(tobj, Document.Document):
                    # ... we need to re-load the Document.  tobj may have spoiled QuotePosition data based on
                    # edits that are currently being abandoned.
                    tobj = Document.Document(tobj.number)
                    # In this case, we also need to refresh the DocumentQuotes Tab
                    self.ControlObject.UpdateDataWindowOnDocumentEdit()
                # Set to None so that it doesn't ask us to save twice, since
                # normally when we load a Transcript with one already loaded,
                # it prompts to save for changes.
                self.dlg.editor.TranscriptObj = None                
                if tobj:
                    # Determine whether we have a Clip Transcript (not pickled) or an Episode Transcript (pickled)
                    # and load it.
                    if isinstance(tobj, Transcript.Transcript) and (tobj.clip_num > 0):
                        # If we're using the Rich Text Ctrl ...
                        if TransanaConstants.USESRTC:
                            # Load the Clip Transcript
                            self.dlg.editor.load_transcript(tobj)
                        # If we're using the Styled Text Ctrl ...
                        else:
                            # The Clip Transcript is never pickled.
                            self.dlg.editor.load_transcript(tobj)
                    else:
                        # If we're using the Rich Text Ctrl ...
                        if TransanaConstants.USESRTC:
                            # Load the Episode Transcript
                            self.dlg.editor.load_transcript(tobj)
                        else:
                            # The Episode Transcript will always have been pickled in this circumstance.
                            self.dlg.editor.load_transcript(tobj, dataType='pickle')
                # If our object is a DOCUMENT ...
                if isinstance(tobj, Document.Document):
                    # Let the Visualization Window know what Document to visualize ...
                    self.ControlObject.ReplaceVisualizationWindowTextObject(tobj)
                    # ... and the Keyword Visualization
                    self.ControlObject.UpdateKeywordVisualization()
            # Drop the Record Lock
            self.dlg.editor.TranscriptObj.unlock_record()
            # Set the Read only state based on the button's indent
            self.dlg.editor.set_read_only(not can_edit)
            # update the Toolbar's state
            self.UpdateEditingButtons()
        # if entering Edit Mode ...
        else:

            if TransanaConstants.partialTranscriptEdit:
                action = 'EnterEditMode'
                # Send the action to the Editor's UpdateCurrentContents method, which handles
                # loading and unloading data for editing large transcripts
                self.dlg.editor.UpdateCurrentContents(action)

            try:
                # Remember the original Last Save time
                oldLastSaveTime = self.dlg.editor.TranscriptObj.lastsavetime
                # Try to get a record lock
                self.dlg.editor.TranscriptObj.lock_record()
                # If the Transcript Object was updated during the Record Lock (due to having been
                # edited in the interim by another user) we need to refresh the editor!
                # Check the new LastSaveTime with the original one.
                if oldLastSaveTime != self.dlg.editor.TranscriptObj.lastsavetime:
                    if isinstance(self.dlg.editor.TranscriptObj, Document.Document):
                        msg = _('This Document has been updated since you originally loaded it!\nYour copy of the record will be refreshed to reflect the changes.')
                    elif isinstance(self.dlg.editor.TranscriptObj, Transcript.Transcript):
                        msg = _('This Transcript has been updated since you originally loaded it!\nYour copy of the record will be refreshed to reflect the changes.')
                    elif isinstance(self.dlg.editor.TranscriptObj, Quote.Quote):
                        msg = _('This Quote has been updated since you originally loaded it!\nYour copy of the record will be refreshed to reflect the changes.')
                    dlg = Dialogs.InfoDialog(self, msg)
                    dlg.ShowModal()
                    dlg.Destroy()
                    # If Time Code Data is displayed ...
                    if self.toolbar.GetToolState(self.CMD_SHOWHIDETIME_ID):
                        # ... signal that it was being shown ...
                        timeCodeDataShowing = True
                        # ... and HIDE it by toggling the button and calling the event!  We don't want it to be propagated.
                        self.toolbar.ToggleTool(self.CMD_SHOWHIDETIME_ID, False)
                        self.OnShowHideValues(evt)
                    # If Time Code Data is NOT displayed ...
                    else:
                        # ... then note that it isn't so we know not to re-display it later.
                        timeCodeDataShowing = False
                    # Determine whether we have a Clip Transcript (not pickled) or an Episode Transcript (pickled)
                    # and load it.
                    if isinstance(self.GetCurrentObject, Transcript.Transcript) and \
                       (self.GetCurrentObject.clip_num > 0):
                        # The Clip Transcript is never pickled.
                        self.dlg.editor.load_transcript(self.dlg.editor.TranscriptObj)
                    else:
                        if TransanaConstants.USESRTC:
                            # Load the Episode Transcript
                            self.dlg.editor.load_transcript(self.dlg.editor.TranscriptObj)
                        else:
                            # The Episode Transcript will always have been pickled in this circumstance.
                            self.dlg.editor.load_transcript(self.dlg.editor.TranscriptObj, dataType='pickle')
                        # ... and the Keyword Visualization
                        self.ControlObject.UpdateKeywordVisualization()
                    # reloading the Transcript unfortunately unlocks the record.  I can't figure out
                    # a clever way to avoid this, so let's just re-lock the record.
                    self.dlg.editor.TranscriptObj.lock_record()
                    # If Time Code Data was being displayed ...
                    if timeCodeDataShowing:
                        # ... RE-DISPLAY it by toggling the button and calling the event!
                        self.toolbar.ToggleTool(self.CMD_SHOWHIDETIME_ID, True)
                        self.OnShowHideValues(evt)
                # Set the Read only state based on the button's indent
                self.dlg.editor.set_read_only(not can_edit)
                # update the Toolbar's state
                self.UpdateEditingButtons()

                # If we're using the Rich Text Ctrl ...
                if TransanaConstants.USESRTC:
                    # There's a bit of weirdness with the RTC, or more likely with my implementation of it.
                    # New transcripts don't have proper values for Paragraph Alignment or Line Spacing.
                    # I tried to add them when initializing the control, but that caused problems with
                    # RTF import.  The following code tries to fix that problem when the USER first opens
                    # a new transcript.  (That's why it's here instead of elsewhere.)
                    if self.dlg.editor.GetLastPosition() == 0:
                        self.dlg.editor.SetTxtStyle(parAlign = wx.TEXT_ALIGNMENT_LEFT, parLineSpacing = wx.TEXT_ATTR_LINE_SPACING_NORMAL)
                    # Call CheckFormatting() to set the Default and Basic Styles correctly
                    self.dlg.editor.CheckFormatting()
                    # Set the Cursor Focus
                    self.dlg.editor.SetFocus()

            # Process Record Lock exception
            except TransanaExceptions.RecordLockedError, e:
                self.dlg.editor.TranscriptObj.lastsavetime = oldLastSaveTime
                self.toolbar.ToggleTool(self.CMD_READONLY_ID, not self.toolbar.GetToolState(self.CMD_READONLY_ID))
                if self.dlg.editor.TranscriptObj.id != '':
                    rtype = _('Transcript')
                    idVal = self.dlg.editor.TranscriptObj.id
                    TransanaExceptions.ReportRecordLockedException(rtype, idVal, e)
                else:
                    if 'unicode' in wx.PlatformInfo:
                        msg = unicode(_('You cannot proceed because you cannot obtain a lock on the Clip Transcript.\n'), 'utf8') + \
                              unicode(_('The record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = msg % (e.user)
                    else:
                        msg = _('You cannot proceed because you cannot obtain a lock on the Clip Transcript.\n') + \
                              _('The record is currently locked by %s.\nPlease try again later.') 
                    dlg = Dialogs.ErrorDialog(self, msg)
                    dlg.ShowModal()
                    dlg.Destroy()
            # Process Record Not Found exception
            except TransanaExceptions.RecordNotFoundError, e:
                # Reject the attempt to go into Edit mode by toggling the button back to its original state
                self.toolbar.ToggleTool(self.CMD_READONLY_ID, not self.toolbar.GetToolState(self.CMD_READONLY_ID))
                if self.parent.dlg.editor.TranscriptObj.id != '':
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(_('You cannot proceed because %s "%s" cannot be found.'), 'utf8') + \
                              unicode(_('\nIt may have been deleted by another user.'), 'utf8')
                        msg = msg % (unicode(_('Transcript'), 'utf8'), self.dlg.editor.TranscriptObj.id)
                    else:
                        msg = _('You cannot proceed because %s "%s" cannot be found.') + \
                              _('\nIt may have been deleted by another user.') % (_('Transcript'), self.dlg.editor.TranscriptObj.id)
                else:
                    msg = _('You cannot proceed because the Clip Transcript cannot be found.\n') + \
                          _('It may have been deleted by another user.')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8')
                dlg = Dialogs.ErrorDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()
                # Clear the deleted objects from the Transana Interface.  Otherwise, problems arise.
                self.ControlObject.ClearAllWindows()

    def OnFormat(self, event):
        """ Format Text """
        # Call the ControlObject's Format Dialog method
        self.ControlObject.TranscriptCallFormatDialog()

    def OnCloseCurrent(self, event):
        self.CloseCurrent(event)
        
    def UpdateEditingButtons(self):
        """ Update the Toolbar Buttons depending on the Edit State """
        # Enable/Disable editing buttons
        can_edit = not self.dlg.editor.get_read_only()
        for x in (self.CMD_UNDO_ID, self.CMD_BOLD_ID, self.CMD_ITALIC_ID, self.CMD_UNDERLINE_ID, \
                  self.CMD_RISING_INT_ID, self.CMD_FALLING_INT_ID, \
                  self.CMD_AUDIBLE_BREATH_ID, self.CMD_WHISPERED_SPEECH_ID, self.CMD_FORMAT_ID, \
                  self.CMD_PROPAGATE_ID):
            self.toolbar.EnableTool(x, can_edit)
        # Enable/Disable Transcript menu Items
        self.ControlObject.SetTranscriptEditOptions(can_edit)

    def UpdateMultiTranscriptButtons(self, enable):
        """ Update the Toolbar Buttons depending on the enable parameter """
        for x in (self.CMD_MULTISELECT_ID, self.CMD_PLAY_ID):
            self.toolbar.EnableTool(x, enable)

    def UpdateToolbarAppearance(self):
        """ When changing between Transcript Panels in the Splitter, we need to update ALL of the Toolbar's buttons for
            whether they are enabled and whether they are pressed """

        # Enable/Disable editing buttons
        can_edit = not self.dlg.editor.get_read_only()

        # Set the Editing Options
        self.UpdateEditingButtons()
        # Enable/Disable Transcript menu Items
        self.ControlObject.SetTranscriptEditOptions(can_edit)

# We need to figure out what else we need to do with these buttons!!
##          self.CMD_BOLD_ID,
##          self.CMD_ITALIC_ID,
##          self.CMD_UNDERLINE_ID,

        # Let's try to get a text sample!
        isBold = self.dlg.editor.get_bold()
        isItalic = self.dlg.editor.get_italic()
        isUnderline = self.dlg.editor.get_underline()

#        print "Bold:", isBold
#        print "Italic:", isItalic
#        print "Underline:", isUnderline

        # If we are Showing Time Codes and the Time Code button is NOT pressed ...
        if (self.dlg.editor.codes_vis and not self.toolbar.GetToolState(self.CMD_SHOWHIDE_ID)) or \
            (not(self.dlg.editor.codes_vis) and self.toolbar.GetToolState(self.CMD_SHOWHIDE_ID)):
            # ... press it!
            self.toolbar.ToggleTool(self.CMD_SHOWHIDE_ID, self.dlg.editor.codes_vis)

        # If we are Showing Time Code DATA and the Time Code DATA button is NOT pressed ...
        if (self.dlg.editor.timeCodeDataVisible and not self.toolbar.GetToolState(self.CMD_SHOWHIDETIME_ID)) or \
            (not(self.dlg.editor.timeCodeDataVisible) and self.toolbar.GetToolState(self.CMD_SHOWHIDETIME_ID)):
            # ... press it!
            self.toolbar.ToggleTool(self.CMD_SHOWHIDETIME_ID, self.dlg.editor.timeCodeDataVisible)
        
        # If we are in Edit mode and the Edit Mode button is NOT pressed ...
        if (can_edit and not self.toolbar.GetToolState(self.CMD_READONLY_ID)) or \
            (not(can_edit) and self.toolbar.GetToolState(self.CMD_READONLY_ID)):
            # ... press it!
            self.toolbar.ToggleTool(self.CMD_READONLY_ID, can_edit)
            
        # If this is called from a Transcript ...
        if (isinstance(self.dlg.editor.TranscriptObj, Transcript.Transcript) and \
            (len(self.nb.GetCurrentPage().GetChildren()) > 1)):
            # ... we need to disable the MultiTranscript Buttons
            self.UpdateMultiTranscriptButtons(True)
        else:
            # ... we need to disable the MultiTranscript Buttons
            self.UpdateMultiTranscriptButtons(False)

    def OnQuickClip(self, event):
        """ Create a Quick Quote or Clip """
        # If this is called from a Transcript ...
        if isinstance(self.dlg.editor.TranscriptObj, Transcript.Transcript):
            # Call upon the Control Object to create the Quick Clip
            self.ControlObject.CreateQuickClip()
        # If this is called from a Document or Quote ...
        elif isinstance(self.dlg.editor.TranscriptObj, Document.Document) or \
             isinstance(self.dlg.editor.TranscriptObj, Quote.Quote):
            # Call upon the Control Object to create the Quick Quote
            self.ControlObject.CreateQuickQuote()

    def OnEditKeywords(self, evt):
        """ Implement the Edit Keywords button """
        # Determine if a Transcript is loaded, and if so, what kind
        if self.dlg.editor.TranscriptObj != None:
            # Initialize a list where we can keep track of clip transcripts that are locked because they are in Edit mode.
            TranscriptLocked = []
            try:
                # If an episode/clip has multiple transcripts, we could run into lock problems.  Let's try to detect that.
                # (This is probably poor form from an object-oriented standpoint, but I don't know a better way.)
                # For each currently-open Transcript window ...
                for trWin in self.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren():
                    # ... note if the transcript is currently locked.
                    TranscriptLocked.append(trWin.editor.TranscriptObj.isLocked)
                    # If it is locked ...
                    if trWin.editor.TranscriptObj.isLocked:
                        # Leave Edit Mode, which will prompt about saving the Transcript.
                        # a) toggle the button
                        trWin.toolbar.ToggleTool(trWin.parent.parent.parent.CMD_READONLY_ID, not trWin.toolbar.GetToolState(trWin.parent.parent.parent.CMD_READONLY_ID))
                        # b) call the event that responds to the button state change
                        trWin.parent.parent.parent.OnReadOnlySelect(evt)
                # If we have a Document ...
                if isinstance(self.dlg.editor.TranscriptObj, Document.Document):
                    obj = Document.Document(self.dlg.editor.TranscriptObj.number)
                # If we have a Quote ...
                elif isinstance(self.dlg.editor.TranscriptObj, Quote.Quote):
                    obj = Quote.Quote(self.dlg.editor.TranscriptObj.number)
                # If the underlying Transcript object has a clip number, we're working with a CLIP.
                elif self.dlg.editor.TranscriptObj.clip_num > 0:
                    # Finally, we can load the Clip object.
                    obj = Clip.Clip(self.dlg.editor.TranscriptObj.clip_num)
                # Otherwise ...
                else:
                    # ... load the Episode
                    obj = Episode.Episode(self.dlg.editor.TranscriptObj.episode_num)
            # Process Record Not Found exception
            except TransanaExceptions.RecordNotFoundError, e:
                msg = _('You cannot proceed because the %s cannot be found.')
                if isinstance(self.dlg.editor.TranscriptObj, Document.Document):
                    prompt = _('Document')
                # If the Transcript does not have a clip number, 
                elif self.dlg.editor.TranscriptObj.clip_num == 0:
                    prompt = _('Episode')
                else:
                    prompt = _('Clip')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(msg, 'utf8') % unicode(prompt, 'utf8') + \
                          unicode(_('\nIt may have been deleted by another user.'), 'utf8')
                else:
                    msg = msg % prompt + _('\nIt may have been deleted by another user.')
                dlg = Dialogs.ErrorDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()
                # Clear the deleted objects from the Transana Interface.  Otherwise, problems arise.
                self.ControlObject.ClearAllWindows()
                return
            try:
                # Lock the data record
                obj.lock_record()
                # Determine the title for the KeywordListEditForm Dialog Box
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("Keywords for %s"), 'utf8')
                else:
                    prompt = _("Keywords for %s")
                dlgTitle = prompt % obj.id
                # Extract the keyword List from the Data object
                kwlist = []
                for kw in obj.keyword_list:
                    kwlist.append(kw)
                    
                # Create/define the Keyword List Edit Form
                dlg = KeywordListEditForm.KeywordListEditForm(self, -1, dlgTitle, obj, kwlist)
                # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
                contin = True
                # While the "continue" flag is True ...
                while contin:
                    # if the user pressed "OK" ...
                    try:
                        # Show the Keyword List Edit Form and process it if the user selects OK
                        if dlg.ShowModal() == wx.ID_OK:
                            # Clear the local keywords list and repopulate it from the Keyword List Edit Form
                            kwlist = []
                            for kw in dlg.keywords:
                                kwlist.append(kw)

                            # Copy the local keywords list into the appropriate object
                            obj.keyword_list = kwlist

                            # If we are dealing with a Document ...
                            if isinstance(obj, Document.Document):
                                # Check to see if there are keywords to be propagated
                                self.ControlObject.PropagateObjectKeywords(_('Document'), obj.number, obj.keyword_list)
                            # If we are dealing with an Episode ...
                            elif isinstance(obj, Episode.Episode):
                                # Check to see if there are keywords to be propagated
                                self.ControlObject.PropagateObjectKeywords(_('Episode'), obj.number, obj.keyword_list)

                            # Save the Data object
                            obj.db_save()

                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                if isinstance(obj, Episode.Episode):
                                    msg = 'Episode %d' % obj.number
                                    msgObjType = 'Episode'
                                    msgObjClipEpNum = 0
                                elif isinstance(obj, Clip.Clip):
                                    msg = 'Clip %d' % obj.number
                                    msgObjType = 'Clip'
                                    msgObjClipEpNum = obj.episode_num
                                else:
                                    msg = ''
                                    msgObjType = 'None'
                                    msgObjClipEpNum = 0
                                if msg != '':
                                    if TransanaGlobal.chatWindow != None:
                                        # Send the "Update Keyword List" message
                                        TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                            # If any Keyword Examples were removed, remove them from the Database Tree
                            for (keywordGroup, keyword, clipNum) in dlg.keywordExamplesToDelete:
                                self.ControlObject.RemoveDataWindowKeywordExamples(keywordGroup, keyword, clipNum)

                            # Update the Data Window Keywords Tab (this must be done AFTER the Save)
                            self.ControlObject.UpdateDataWindowKeywordsTab()

                            # Update the Keyword Visualizations
                            self.ControlObject.UpdateKeywordVisualization()

                            # Notify other computers to update the Keyword Visualization as well.
                            if not TransanaConstants.singleUserVersion:
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % (msgObjType, obj.number, msgObjClipEpNum))

                            # If we do all this, we don't need to continue any more.
                            contin = False

                        # If the user pressed Cancel ...
                        else:
                            # ... then we don't need to continue any more.
                            contin = False

                    # Handle "SaveError" exception
                    except TransanaExceptions.SaveError:
                        # Display the Error Message, allow "continue" flag to remain true
                        errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # Refresh the Keyword List, if it's a changed Keyword error
                        dlg.refresh_keywords()
                        # Highlight the first non-existent keyword in the Keywords control
                        dlg.highlight_bad_keyword()

                    # Handle other exceptions
                    except:
                        if DEBUG:
                            import traceback
                            traceback.print_exc(file=sys.stdout)
                        # Display the Exception Message, allow "continue" flag to remain true
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_("Exception %s: %s"), 'utf8')
                        else:
                            prompt = _("Exception %s: %s")
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                        errordlg.ShowModal()
                        errordlg.Destroy()

                # Unlock the Data Object
                obj.unlock_record()

            except TransanaExceptions.RecordLockedError, e:
                """Handle the RecordLockedError exception."""
                if isinstance(obj, Episode.Episode):
                    rtype = _('Episode')
                elif isinstance(obj, Clip.Clip):
                    rtype = _('Clip')
                idVal = obj.id
                TransanaExceptions.ReportRecordLockedException(rtype, idVal, e)

            # Process Record Not Found exception
            except TransanaExceptions.RecordNotFoundError, e:
                msg = _('You cannot proceed because the %s cannot be found.')
                prompt = _('Transcript')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(msg, 'utf8') % unicode(prompt, 'utf8') + \
                          unicode(_('\nIt may have been deleted by another user.'), 'utf8')
                dlg = Dialogs.ErrorDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()
                # Clear the deleted objects from the Transana Interface.  Otherwise, problems arise.
                self.ControlObject.ClearAllWindows()

    def OnSave(self, evt):
        """ Implement the Save Button """
        self.ControlObject.SaveTranscript(continueEditing=TransanaConstants.partialTranscriptEdit)

    def OnPropagate(self, event):
        """ Implement Propagate Changes button """
        # If Time Code Data is displayed ...
        if self.toolbar.GetToolState(self.CMD_SHOWHIDETIME_ID):
            # ... signal that it was being shown ...
            timeCodeDataShowing = True
            # ... and HIDE it by toggling the button and calling the event!  We don't want it to be propagated.
            self.toolbar.ToggleTool(self.CMD_SHOWHIDETIME_ID, False)
            self.OnShowHideValues(event)
        # If Time Code Data is NOT displayed ...
        else:
            # ... then note that it isn't so we know not to re-display it later.
            timeCodeDataShowing = False
        
        # Call the Propagate Changes method in the Control Object
        self.ControlObject.PropagateChanges(self.nb.GetCurrentPage().activePanel)

        # If Time Code Data was being displayed ...
        if timeCodeDataShowing:
            # ... RE-DISPLAY it by toggling the button and calling the event!
            self.toolbar.ToggleTool(self.CMD_SHOWHIDETIME_ID, True)
            self.OnShowHideValues(event)

    def OnMultiSelect(self, event):
        """ Implements the "Match Selection in Multiple Transcripts" button """
        self.ControlObject.MultiSelect(self.nb.GetCurrentPage().activePanel)

    def OnMultiPlay(self, event):
        """ Implements the "Play" button in the multi-transcript environment """
        self.ControlObject.MultiPlay()

    def OnStyleChange(self, editor):
        """This event handler is setup in the higher level Transcript Window,
        which instructs the Transcript editor to call this function when
        the current style changes automatically, but not programatically."""
        self.toolbar.ToggleTool(self.CMD_BOLD_ID, editor.get_bold())
        self.toolbar.ToggleTool(self.CMD_ITALIC_ID, editor.get_italic())
        self.toolbar.ToggleTool(self.CMD_UNDERLINE_ID, editor.get_underline())

    # End of Toolbar Methods

    def __size(self):
        """Determine the default size for the Transcript frame."""
        # Get the size of the correct monitor
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            primaryScreen = TransanaGlobal.configData.primaryScreen
        else:
            primaryScreen = 0
        rect = wx.Display(primaryScreen).GetClientArea()
        if not ('wxGTK' in wx.PlatformInfo):
            container = rect[2:4]
        else:
            screenDims = wx.Display(primaryScreen).GetClientArea()
            # screenDims2 = wx.Display(primaryScreen).GetGeometry()
            left = screenDims[0]
            top = screenDims[1]
            width = screenDims[2] - screenDims[0]  # min(screenDims[2], 1280 - self.left)
            height = screenDims[3]
            container = (width, height)

        # Transcript Compontent should be 71.5% of the WIDTH
        width = container[0] * .716  # rect[2] * .715
        # Transcript Compontent should be 74% of the HEIGHT, adjusted for the menu height
        height = (container[1] - TransanaGlobal.menuHeight) * .741  # (rect[3] - TransanaGlobal.menuHeight) * .74
        # Return the SIZE values
        return wx.Size(width, height)

    def __pos(self):
        """Determine default position of Transcript Frame."""
        # Get the size of the correct monitor
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            primaryScreen = TransanaGlobal.configData.primaryScreen
        else:
            primaryScreen = 0
        rect = wx.Display(primaryScreen).GetClientArea()
        if not ('wxGTK' in wx.PlatformInfo):
            container = rect[2:4]
        else:
            # Linux rect includes both screens, so we need to use an alternate method!
            container = TransanaGlobal.menuWindow.GetSize()
        # Get the adjusted default SIZE of the Transcription UI area of the screen
        (width, height) = self.__size()
        # rect[0] compensates if Start menu is on Left
        x = rect[0] + 1
        # rect[1] compensates if Start menu is on Top
        if 'wxGTK' in wx.PlatformInfo:
            # rect2 = wx.Display(primaryScreen).GetGeometry()
            y = (rect[3] - rect[1] - 6) * .35 + 24
        else:
            y = rect[1] + container[1] - height  # rect[1] + rect[3] - height - 3

        if DEBUG:
            print "TranscriptUI_RTC.__pos(): Y = %d, H = %d, Total = %d" % (y, height, y + height)

        # Return the POSITION values
        return (x, y)

    def _get_dlg(self):
        # The notebook's Current Page is a splitter with multiple Panels containing Editors
        # The Notebook's Current Page's activePanel property tells us WHICH Panel we should return
        return self.nb.GetCurrentPage().GetChildren()[self.nb.GetCurrentPage().activePanel]

    def _set_dlg(self):

        print "TranscritionUI_RTC._set_dlg() call"

    def _del_dlg(self):

        print "TranscritionUI_RTC._del_dlg() call"

    dlg = property(_get_dlg, _set_dlg, _del_dlg, """ dlg is the CURRENT Notebook Page's last-clicked MultiSplitterWindow Pane """)


class TranscriptionNotebook(wx.Notebook):
    """ A wxPython Notebook Object class for use on Transana's Document window """

    def __init__(self, parent, pageName=None):
        """ Initialize the Notebook Object """
        # Remember the parent
        self.parent = parent
        # Create a Notebook Object
        wx.Notebook.__init__(self, parent, style=wx.CLIP_CHILDREN | wx.NB_MULTILINE)
        # If a Page Name has been specified ...
        if pageName is not None:
            # ... add a Page
            self.AddNotebookPage(pageName)

        # Define the Page Changed event, so we can update the Toolbar
        self.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.OnPageChanged)

    def ChangeOrientation(self):
        """ Change the Orientation of Splitter Panes in the current Notebook Page """
        # Pass the Orientation Change on to the Splitter Window object
        self.GetPage(self.GetSelection()).ChangeOrientation()

    def AddNotebookPage(self, pageName):
        """ Add a Notebook Page with a Splitter with multiple Panes. """
        # Create a MultiSplitterWindow with panes in each of the specified colors
        splitter = TranscriptionSplitter(self)
        # Add that MultiSplitterWindow to the Notebook with the appropriate Tab Title
        self.AddPage(splitter, pageName)
        # Only after the Page is added can we reset the Splitter Sashes evenly
        splitter.SpacePanelsEvenly()

    def DeleteNotebookPage(self, pageNum):
        """ Delete a Notebook Page / Tab """
        # Delete the specified Page / Tab
        self.DeletePage(pageNum)

        # On OS X (with wxPython 3.0.2.0) ...
        if 'wxMac' in wx.PlatformInfo:
            # ... closing the current page doesn't call the OnPageChanged event, which needs to be called!
            self.OnPageChanged(None)

    def AddPanel(self):
        """ Add a Splitter Panel to the currently-selected Notebook Page """
        # Get the current Page, which is a MultiSplitterWindow object, and add a Panel
        self.GetPage(self.GetSelection()).AddPanel()

    def DeletePanel(self, tab=-1, panel=-1):
        """ Delete the specified Splitter Panel from the specified Notebook Page """
        # If the notebook tab is not specified or is out of range ...
        if (tab <= -1) or (tab > self.GetPageCount() - 1):
            # ... use the currently-selected Notebook Tab
            tab = self.GetSelection()
        # If the splitter pane is not specified or is out of range ...
        if (panel <= -1) or (panel > len(self.GetPage(tab).GetChildren()) - 1):
            # ... use the last-selected panel
            panel = self.GetPage(tab).activePanel
        # Get the specified Page, which is a MultiSplitterWindow object, and delete the specified Panel
        self.GetPage(tab).DeletePanel(panel)

    def OnPageChanged(self, event):
        """ Handle the EVT_NOTEBOOK_PAGE_CHANGED event for the Notebook Control """
        # If we forget event.Skip(), tabs don't update on OS X!  But not if event is None!
        if event != None:
            event.Skip()
        # Because of a weird issue on OS X, we weren't getting the right information from the Notebook in wxPython 3.0.1.1.
        # This is awkward, but works.
        wx.CallAfter(self.OnPageChangedAfter)

    def OnPageChangedAfter(self):
        """ Handle the EVT_NOTEBOOK_PAGE_CHANGED event for the Notebook Control """
        # On OS X, Editor panels would not show up properly when changing from one tab to another.
        # For reasons unknown, the Splitter (Notebook Page contents) was getting hidden.  This fixes that.
        self.GetCurrentPage().Show()

        # If there is media playing ...
        if self.parent.ControlObject.IsPlaying():
            # ... pause it, since we're moving OFF the Transcript here.
            self.parent.ControlObject.Pause()
        # If the Toolbar is disabled (which can happen when closing Pages) ...
        if not self.parent.toolbar.IsEnabled():
            # ... enable the Toolbar
            self.parent.toolbar.Enable(True)
        # Adjust the Common Toolbar to the status of the new active document
        self.parent.UpdateToolbarAppearance()

        # Set the ControlObject's active transcript to the panel that is currently active on the new page
        self.parent.ControlObject.activeTranscript = self.GetCurrentPage().activePanel

        # Signal the main TranscriptWindow object to update Transana's GUI based on this Notebook change
        self.parent.UpdateGUI()

class TranscriptionSplitter(wx.lib.splitter.MultiSplitterWindow):
    """ A wxPython MultiSplitterWindow object class """
    def __init__(self, parent):
        # Remember the parent
        self.parent = parent
        # We need to keep track of which Panel / Editor on this Splitter Window is currently active.  Default to the first one.
        self.activePanel = 0
        # Define the minimum pane sizes
        self.minVertical = 50
        self.minHorizontal = 100
        # Create a MultiSplitterWindow Object
        wx.lib.splitter.MultiSplitterWindow.__init__(self, parent, style=wx.SP_LIVE_UPDATE)
        # Set the SplitterWindow's orientation to Vertical initially
        self.SetOrientation(wx.VERTICAL)
        # Create a Panel which contains a TranscriptEditor
        panel = _TranscriptPanel(self, self.parent.parent.ControlObject, showLineNumbers=True)
        # ... and add that Panel to the Splitter
        self.AppendWindow(panel)
        # We need to handle Splitter Sash Position Changes to maintain minimum Editor sizes when scrolling DOWN.
        # The control's default behavior is to let them scroll off the edge, but it's hard to get them back!
        self.Bind(wx.EVT_SPLITTER_SASH_POS_CHANGED, self.OnSplitterSashPosChanged)
        
    def AddPanel(self):
        """ Add a Panel (with an editor) to the Splitter """
        # Create a Transcript Panel
        panel = _TranscriptPanel(self, self.parent.parent.ControlObject, showLineNumbers=True)
        # Add the Transcript Panel to the Splitter
        self.AppendWindow(panel)
        # Space all Panels evenly, so they are the same size
        self.SpacePanelsEvenly()
        # Select the new panel as the Active Panel
        self.ActivatePanel(len(self.GetChildren()) - 1)

    def ActivatePanel(self, panelNum):
        """ Select which Panel in the Splitter Window is the "Active" Panel, the Editor currently being used """
        # Change the background (frame) color for the OLD active Panel from Active to Inactive 
        self.GetChildren()[self.activePanel].SetBackgroundColour(INACTIVE_COLOR)
        # Change the Active Panel specifier
        self.activePanel = panelNum

        # If there's more than one Panel ...
        if len(self.GetChildren()) > 1:
            # Change the background (frame) color for the NEW active Panel from Inactive to Active
            self.GetChildren()[self.activePanel].SetBackgroundColour(ACTIVE_COLOR)
            # Refresh the screen
            self.Refresh()

        # Set the ControlObject's activeTranscript
        self.parent.parent.ControlObject.activeTranscript = self.activePanel

        # Signal the main TranscriptWindow object to update Transana's GUI based on this Notebook change
        self.parent.parent.UpdateGUI()

    def DeletePanel(self, panel=-1):
        """ Delete the specified Splitter Panel """
        # If the splitter pane is not specified or is out of range ...
        if (panel <= -1) or (panel > len(self.GetChildren()) - 1):
            # ... use the last-selected panel
            panel = self.activePanel
        # Get the selected panel contents as a Window
        win = self.GetWindow(panel)
        # "Detach" the window from the Splitter
        self.DetachWindow(win)
        # Destroy the Window.  (Contents should already have been saved before this is called!)
        win.Destroy()

        # For each of the Splitter's remaining Panes ...
        for x in range(len(self.GetChildren())):
            # ... reset the Panel Number
            self.GetWindow(x).panelNum = x
        # Reset the Splitter Window's active panel to the first panel
        self.activePanel = 0
        # Activate the first panel
        self.ActivatePanel(0)
        # Re-space the remaining panes evenly in the display window
        self.SpacePanelsEvenly()

    def ChangeOrientation(self):
        """ Change the Orientation of the MultiSplitterWindow panes """
        # If we are currently Vertically oriented ...
        if self.GetOrientation() == wx.VERTICAL:
            # ... change the orientation to Horizontal
            self.SetOrientation(wx.HORIZONTAL)
        # If we are currently Horizontally oriented ...
        else:
            # ... change the orientation to Vertical
            self.SetOrientation(wx.VERTICAL)
        # Change the Sash Positions so the panes are equally sized
        self.SpacePanelsEvenly()

    def SpacePanelsEvenly(self):
        """ Make all Panels / Editors take an even proportion of the current visible Window """
        # If we are currently Vertically oriented ...
        if self.GetOrientation() == wx.VERTICAL:
            # ... the size we are concerned with is the HEIGHT
            size = self.GetSize()[1]
        # If we are currently Horizontally oriented ...
        else:
            # ... the size we are concerned with is the WIDTH
            size = self.GetSize()[0]
        # For each Splitter SASH (Panes - 1) ....
        for x in range(len(self.GetChildren()) - 1):
            # ... set the Sash Position to the correct proportion of the correct dimension
            self.SetSashPosition(x, size / len(self.GetChildren()))
        # Call the Splitter Window's SizeWindows() method to make the changes show up
        self.SizeWindows()

    def OnSplitterSashPosChanged(self, event):
        """ Handle the MultiSplitterWindow's Sash Position Change Method to keep all panes always visible on screen """
        # Initialize a variable to keep track of the sum of the sizes of the Panes.  
        sumOfSizes = 0
        # If we are currently Vertically oriented ...
        if self.GetOrientation() == wx.VERTICAL:
            # For each Splitter SASH (Panes - 1) ....
            for x in range(len(self.GetChildren()) - 1):
                # If our sash is SMALLER than allowed ...
                if self.GetSashPosition(x) < self.minVertical:
                    # ... then set the sash to the minimum allowed ...
                    self.SetSashPosition(x, self.minVertical)
                    # ... and add that amount to the sum of pane sizes
                    sumOfSizes += self.minVertical
                # If the sum of pane sizes PLUS the next pane size exceeds the size of the Panel ...
                elif sumOfSizes + self.GetSashPosition(x) > self.GetSize()[1] - (self.minVertical * (len(self.GetChildren()) - (x+1))):
                    # ... determine the maximum position for the start of this pane, taking into account the total number of panes
                    # and the minimum pane size ...
                    val = self.GetSize()[1] - (self.minVertical * (len(self.GetChildren()) - (x+1))) - sumOfSizes
                    # ... the set the sash appropriately ...
                    self.SetSashPosition(x, val)
                    # ... and add that amount to the sum of pane sizes
                    sumOfSizes += val
                # If the pane size does not require any adjustments ...
                else:
                    # ... then add that amount to the sum of pane sizes
                    sumOfSizes += self.GetChildren()[x].GetSize()[1]

        # If we are currently Horizontally oriented ...
        else:
            # For each Splitter SASH (Panes - 1) ....
            for x in range(len(self.GetChildren()) - 1):
                # If our sash is SMALLER than allowed ...
                if self.GetSashPosition(x) < self.minHorizontal:
                    # ... then set the sash to the minimum allowed ...
                    self.SetSashPosition(x, self.minHorizontal)
                    # ... and add that amount to the sum of pane sizes
                    sumOfSizes += self.minHorizontal
                # If the sum of pane sizes PLUS the next pane size exceeds the size of the Panel ...
                elif sumOfSizes + self.GetSashPosition(x) > self.GetSize()[0] - (self.minHorizontal * (len(self.GetChildren()) - (x+1))):
                    # ... determine the maximum position for the start of this pane, taking into account the total number of panes
                    # and the minimum pane size ...
                    val = self.GetSize()[0] - (self.minHorizontal * (len(self.GetChildren()) - (x+1))) - sumOfSizes
                    # ... the set the sash appropriately ...
                    self.SetSashPosition(x, val)
                    # ... and add that amount to the sum of pane sizes
                    sumOfSizes += val
                # If the pane size does not require any adjustments ...
                else:
                    # ... then add that amount to the sum of pane sizes
                    sumOfSizes += self.GetChildren()[x].GetSize()[0]
        # Call the Splitter Window's SizeWindows() method to make the changes show up
        self.SizeWindows()

# import the wxRTC-based RichTextEditCtrl
import RichTextEditCtrl_RTC

class _TranscriptPanel(wx.Panel):
    """ Implement a wx.Panel-based Editor control for the private use of the TranscriptionUI object """

    def __init__(self, parent, controlObject, showLineNumbers=False):
        # Remember the parent
        self.parent = parent

        # Give each Transcript Panel (Splitter Pane) a unique number
#        self.pageNum = parent.parent.GetSelection()
        # Set the Panel Number, used to keep track of which panel is active
        self.panelNum = len(parent.GetChildren())

        # Assign the ControlObject
        self.ControlObject = controlObject
        # Remember whether we're including line numbers
        self.showLineNumbers = showLineNumbers

        # Get the size of the parent window
        psize = parent.GetSizeTuple()
        # reduce the width by 13 pixels (borders)
        width = psize[0] - 13
        # reduce the height by 45 pixels (toolbar, borders)
        height = psize[1] - 45

        # Create a Panel.  Use WANTS_CHARS style so the panel doesn't eat the Enter key.
        wx.Panel.__init__(self, parent, -1, style=wx.WANTS_CHARS, size=(width, height), name='TranscriptPanel')
        # If there's more than one Panel ...
        if self.panelNum > 1:
            # Set the Background Color to the pre-defined Active Color
            self.SetBackgroundColour(ACTIVE_COLOR)
        else:
            # Set the Background Color to the pre-defined Active Color
            self.SetBackgroundColour(INACTIVE_COLOR)

        # add the widgets to the panel
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Let's get a variable that indicates what Toolbar is associated
        self.toolbar = self.parent.parent.parent.toolbar

        # Add a blank Bitmap for creating Line Numbers
        self.lineNumBmp = wx.EmptyBitmap(100, 100)

        # Create a horizontal Sizer to hold the Line Numbers and the Edit Control
        hsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        
        # Create a Panel to hold Line Numbers
        self.lineNum = wx.Panel(self, -1, style = wx.BORDER_DOUBLE)
        # Set the background color for the Line Numbers panel to a pale gray
        self.lineNum.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNFACE))
        # Set the minimum and maximum width of the Line Number Panel to fix its size
        self.lineNum.SetSizeHints(minW = 50, minH = 0, maxW = 50)
        # Add the Line Number Panel to the Horizontal Sizer
        hsizer2.Add(self.lineNum, 0, wx.EXPAND | wx.TOP | wx.LEFT | wx.BOTTOM, 2)
        # Bind the PAINT event used to display line numbers
        self.lineNum.Bind(wx.EVT_PAINT, self.LineNumPaint)
        # Bind Mouse Up so clicking on the Line Numbers selects the panel
        self.lineNum.Bind(wx.EVT_LEFT_UP, self.OnLineNumLeftUp)

        # Create the Rich Text Edit Control itself, using Transana's TranscriptEditor object
        self.editor = TranscriptEditor_RTC.TranscriptEditor(self, -1, self.parent.parent.parent.OnStyleChange, updateSelectionText=True)
        # Place the editor on the horizontal sizer
        hsizer2.Add(self.editor, 1, wx.EXPAND | wx.ALL, 2)

        # Signal that we do NOT need to draw the Line Numbers initially
        self.redraw = False
        # Create a timer to update the Line Numbers
        self.LineNumTimer = wx.Timer()
        # Bind the method to the Timer event
        self.LineNumTimer.Bind(wx.EVT_TIMER, self.OnTimer)  # EditorPaint)

        # Disable the Search components initially
        self.EnableSearch(False)

        # Put the horizontal sizer in the main vertical sizer
        sizer.Add(hsizer2, 1, wx.EXPAND)
        if "__WXMAC__" in wx.PlatformInfo:
            # This adds a space at the bottom of the frame on Mac, so that the scroll bar will get the down-scroll arrow.
            sizer.Add((0, 15))
        # Finish the Size implementation
        self.SetSizer(sizer)
        self.SetAutoLayout(True)
        self.Layout()
        # Set the focus in the Editor
        self.editor.SetFocus()

        # Call Line Number Paint just once so the screen looks right
        self.EditorPaint(None)

        # Set a minimum size (THIS DOES NOT WORK!!)
        self.SetMinSize((parent.minHorizontal, parent.minVertical))
        self.editor.SetMinSize((parent.minHorizontal, parent.minVertical))
        self.SetSizeHints(parent.minHorizontal, parent.minVertical)
        self.editor.SetSizeHints(parent.minHorizontal, parent.minVertical)

        # Capture Size Changes
        wx.EVT_SIZE(self, self.OnSize)

        wx.EVT_IDLE(self, self.OnIdle)

        if DEBUG:
            print "TranscriptionUI_RTC._TranscriptPanel.__init__():  Initial size:", self.GetSize()

    def ActivatePanel(self):
        """ Capture when a Panel is selected by the user """
        # Let the parent know which Panel is selected or active
        self.parent.ActivatePanel(self.panelNum)

##        print "TranscriptionUI_RTC._TranscriptPanel.ActivatePanel():", self.panelNum
##        print "  ", self.editor.TranscriptObj.id, self.editor.get_read_only()
##        print
        
        # Adjust the Common Toolbar to the status of the new active document
        self.parent.parent.parent.UpdateToolbarAppearance()

    def TranscriptModified(self):
        """Return TRUE if transcript was modified since last save."""
        # Pass through to the Editor
        return self.editor.modified()

    def OnSize(self, event):
        """ Transcription Window Resize Method """
        # If we are not doing global resizing of all windows ...
        if not TransanaGlobal.resizingAll:
            # Get the position of the Transcript window
            (left, top) = self.parent.parent.parent.GetPositionTuple()
            # Get the size of the Transcript window
            (width, height) = self.parent.parent.parent.GetSize()
            # Call the ControlObject's routine for adjusting all windows

            if DEBUG:
                print
                print "TranscritionUI_RTC.OnSize(): Call 4", 'Transcript', width, left, width + left, top - 1

            # As long as the ControlObject is defined ...
            if (self.ControlObject != None):
                # ... update all Transana Window Positions
                self.ControlObject.UpdateWindowPositions('Transcript', width + left, YUpper = top - 1)
        # Call the Transcript Window's Layout.
        self.Layout()
        # We may need to scroll to keep the current selection in the visible part of the window.
        # Find the start of the selection.
        start = self.editor.GetSelectionStart()
        # Make sure the current selection is visible
        self.editor.ShowCurrentSelection()

    def LineNumPaint(self, event):
        """ Paint Event Handler for the Line Number control.  This is used to display Line Numbers. """
        # Call the parent event handler
        event.Skip()
        # Create a Buffered Paint Device Context for the Line Number panel which will draw the Line Number bitmap
        # created in the EditorPaint() method
        dc = wx.BufferedPaintDC(self.lineNum, self.lineNumBmp, style=wx.BUFFER_VIRTUAL_AREA)

    def ClearLineNum(self):
        """ Clear Line Numbers Control """
        # Get the size of the Line Number control
        (w, h) = self.lineNum.GetSize()
        # Create an empty bitmap the same size as the lineNum control
        self.lineNumBmp = wx.EmptyBitmap(w, h)
        # Create a Buffered Device Context for manipulating the Bitmap
        dc = wx.BufferedDC(None, self.lineNumBmp)
        # Paint the background of the Bitmap the color the Control's Background is supposed to be
        dc.SetBackground(wx.Brush(self.lineNum.GetBackgroundColour()))
        # Clear the Bitmap
        dc.Clear()
        # Call Refresh so the Control will re-paint
        self.lineNum.Refresh()

    def AddLineNum(self, num, yPos):
        """ Add Line Numbers to the Line Number Control """
        # Get a buffered Device Context based on the line number bitmap
        dc = wx.BufferedDC(None, self.lineNumBmp)
        # Specify the default font
        font = wx.Font(10, wx.ROMAN, wx.NORMAL, wx.NORMAL)
        # Set the font for the device context
        dc.SetFont(font)
        # Set the text color for the device context
        dc.SetTextForeground(wx.BLACK)
        # Get the size of the line number text to be drawn on the device context
        (w, h) = dc.GetTextExtent(str(num))
        # Make the Device Context editable
        dc.BeginDrawing()
        # Start exception handling
        try:
            # Place the Line Number on the Device Context, right-justified
            dc.DrawText(str(num), self.lineNum.GetSize()[0] - w - 5, yPos)
        # if an excepction arises ...
        except:
            # ... ignore it.
            import sys
            print "TranscriptionUI._TranscriptPanel.AddLineNum():"
            print sys.exc_info()[0]
            print sys.exc_info()[1]
            print
            pass
        # The Device Context won't be edited any more here.
        dc.EndDrawing()

    def OnLineNumLeftUp(self, event):
        """ Left Click Up event for the Line Number Panel """
        # Clicking on the Line Numbers should select the Panel without changing the Selection
        self.ActivatePanel()

    def OnTimer(self, event):
        # Instead of drawing the numbers, just signal that can be done on idle
        self.redraw = True

    def OnIdle(self, event):
        """ IDLE event handler """
        # If we have line numbers to draw ...
        if self.redraw:
            # if the media file is NOT playing ...
            if not self.ControlObject.IsPlaying():
                # ... update the line numbers
                self.EditorPaint(event)
                # ... and signal that they've been re-drawn
                self.redraw = False
            # If the media file IS playing ...
            else:
                # ... Clear the line numbers.  (We don't have time to redraw them during HD playback!)
                self.ClearLineNum()

    def EditorPaint(self, event):
        """ Paint Event Handler for the Editor control.  This is used to display Line Numbers. """
        # If we're not showing line numbers, we can skip this!
        if not self.showLineNumbers:
            return

        # We need to know how long this is taking.  Get the start time.
        start = time.time()
        
        # Call the parent event handler, if there is one
        if event != None:
            event.Skip()
        # Clear all line numbers
        self.ClearLineNum()
        # Initialize the current document position to before the start of the document
        curPos = -1
        # Initialize the current line to 0, before the first line of the document
        curLine = 0
        # Initialize the positioning offset to zero
        offset = 0
        # Depending on platform, set the positioning adjustment factor for line spacing
        if 'wxMac' in wx.PlatformInfo:
            adjustmentFactor = 1.5
            stepSize = 6
        else:
            adjustmentFactor = 1.75
            stepSize = 1

        # For each stepSize vertical pixels in the Transcript display ...
        # (Using a START of 10 prevents overlapping first line numbers.  Using a STEP of stepSize prevents typing lag on the Mac.)
        for y in range(10, self.editor.GetSize()[1], stepSize):
            # This raises an exception sometimes on OS X.  I can't recreate it, but I have gotten a couple reports from the field.
            try:
                # If we're using wxPython 2.8.x.x ...
                if wx.VERSION[:2] == (2, 8):
                    # ... get the character position (pos) of the character at the (10, y) pixel
                    (result, pos) = self.editor.HitTest((10, y))
                else:
                    # ... get the character position (pos) of the character at the (10, y) pixel
                    (result, pos) = self.editor.HitTestPos((10, y))
            # If there's an exception ...
            except:
                # ... we probably don't actually HAVE a transcript yet, so initial values should do
                result = 0
                pos = -1
            # if we have move down on screen to a new character position ...
            if pos > curPos:
                # ... get the formatting for the current character
                textAttr = self.editor.GetStyleAt(pos)
                # Adjust the positioning offset for Paragraph Space Before.
                # (Dividing by three approximately translates centimeters to pixels!)
                offset += int(textAttr.GetParagraphSpacingBefore() / 3)
                # Calculate the Line Number
                tmpLine = self.editor.PositionToXY(pos + 1)[1] + 1
                # Let's not get carried away here!
                if tmpLine > 1000000:
                    tmpLine = 1
                # Use PositionToXY to translate character position into text row/column, then
                # see if we've moved on to the next LINE.  (Line is paragraph number, not physical
                # screen line!)  If so ...
                if curLine < tmpLine:
                    # ... add the line number to to the Line Number display at the proper (adjusted) vertical value
                    self.AddLineNum(tmpLine, y + offset)
                    # Since we've added the line number, update the current line number value
                    curLine = tmpLine
                # Update the current character position value
                curPos = pos
                # We need to adjust the NEXT line number position for Paragraph Space After.  Note that this
                # re-initialized the offset value.
                offset = int(textAttr.GetParagraphSpacingAfter() / 3)
                # We also need to adjust the NEXT line number position for Line Spacing.
                offset += int((textAttr.GetLineSpacing() - 10) * adjustmentFactor)
        # Now that the line numbers have been determined, cause the Line Number control to be updated
        self.lineNum.Refresh()

        # note the End Time
        end = time.time()

        val = 0

        # If updating line numbers takes more than an acceptable amount of time, adjust the frequency of updates.
        if (end - start > 0.36) and (self.LineNumTimer.GetInterval() < 10000):
            # Reset to 10 seconds
            val = 10000
        elif (end - start > 0.30) and (self.LineNumTimer.GetInterval() < 7000):
            # Reset to 7 seconds
            val = 7000
        elif (end - start > 0.22) and (self.LineNumTimer.GetInterval() < 4000):
            # Reset to 4 seconds
            val = 4000
        elif (end - start < 0.22) and (self.LineNumTimer.GetInterval() > 1000):
            # Reset to 1 second
            val = 1000

        if val > 0:
            self.LineNumTimer.Stop()
            self.LineNumTimer.Start(val)

    def OnSearch(self, event):
        """ Implement the Toolbar's QuickSearch """
        # Get the text for the search
        txt = self.parent.parent.parent.searchText.GetValue()
        # If there is text ...
        if txt != '':
            # Determine whether we're searching forward or backward
            if event.GetId() == self.parent.parent.parent.CMD_SEARCH_BACK_ID:
                direction = "back"
            # Either CMD_SEARCH_FORWARD_ID or ENTER in the text box indicate forward!
            else:
                direction = "next"
            # Set the focus back on the editor component, rather than the button, so Paste or typing work.
            self.editor.SetFocus()
            # Perform the search in the Editor
            self.editor.find_text(txt, direction)

    def EnableSearch(self, enable):
        """ Change the "Enabled" status of the Search controls """
        # Enable / Disable the Back Button
        self.toolbar.EnableTool(self.parent.parent.parent.CMD_SEARCH_BACK_ID, enable)
        # Enable / Disable the Search Text Box
        self.parent.parent.parent.searchText.Enable(enable)
        # Enable / Disable the Forward Button
        self.toolbar.EnableTool(self.parent.parent.parent.CMD_SEARCH_NEXT_ID, enable)
        # If we're enabling Search and are showing Line Numbers ...
        if enable and self.showLineNumbers:
            # update the Line Numbers every second.  (This isn't Search specific, but it's convenient!)
            self.LineNumTimer.Start(1000)
        else:
            # Stop the Line Numbers timer
            self.LineNumTimer.Stop()

    def ClearSearch(self):
        """ Clear the Search Box """
        # Clear the Search Text box
        self.parent.parent.parent.searchText.SetValue('')
        
    def UpdateSelectionText(self, text):
        """ Update the text indicating the start and end points of the current selection """
        # Update the selectionText label with the supplied text
        if self.parent.parent.parent.selectionText.GetLabel() != text:
            self.parent.parent.parent.selectionText.SetLabel(text)

    def ChangeLanguages(self):
        """ Change Languages """
        pass

##class _ToolbarPanel(wx.Panel):
##    """ A subclass of wxPanel designed to hold the Document Toolbar. """
##    def __init__(self, parent):
##        # Trying to make the Toolbar larger on OS X.
##        if 'wxMac' in wx.PlatformInfo:
##            toolbarHeight = 32
##        else:
##            toolbarHeight = 24
##        # Initialize the Panel
##        wx.Panel.__init__(self, parent, -1, size=(780, toolbarHeight), style=wx.WANTS_CHARS)
##        # Remember the Panel's parent
##        self.parent = parent
##
##    def OnSearch(self, event):
##        """ Toolbar's Search function """
##        # Pass this on to the parent
##        self.parent.OnSearch(event)
##
##    def CloseCurrent(self, event):
##        """ Toolbar's Close Current Document function """
##        # Pass this on to the parent
##        self.parent.CloseCurrent(event)
##
##    # When I moved the Toolbar to the ToolbarPanel, some things broke.  "dlg" references didn't work any more.
##    # This fixes that.
##    def _get_dlg(self):
##        # The notebook's Current Page is a splitter with multiple Panels containing Editors
##        # The Notebook's Current Page's activePanel property tells us WHICH Panel we should return
##        return self.parent.nb.GetCurrentPage().GetChildren()[self.parent.nb.GetCurrentPage().activePanel]
##
##    def _set_dlg(self):
##
##        print "TranscritionUI_RTC._ToolbarPanel._set_dlg() call"
##
##    def _del_dlg(self):
##
##        print "TranscritionUI_RTC._ToolbarPanel._del_dlg() call"
##
##    dlg = property(_get_dlg, _set_dlg, _del_dlg, """ dlg is the CURRENT Notebook Page's last-clicked MultiSplitterWindow Pane """)
##
##    def _get_nb(self):
##        return self.parent.nb
##
##    def _set_nb(self):
##
##        print "TranscriptionUI_RTC._ToolbarPanel._set_nb() call"
##
##    def _del_nb(self):
##
##        print "TranscriptionUI_RTC._ToolbarPanel._del_nb() call"
##
##    nb = property(_get_nb, _set_nb, _del_nb, """ nb is the NOTEBOOK """)
