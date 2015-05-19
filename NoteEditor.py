# Copyright (C) 2003 - 2007 The Board of Regents of the University of Wisconsin System 
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

""" This module implements the Note Editor.  It's used for editing plain text Notes """

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

# import wxPython
import wx
# import Python's os module
import os
# import Python's time module
import time
# import Transana's Dialogs
import Dialogs
# import Transana's Constants
import TransanaConstants
# import Transana's Global variables
import TransanaGlobal
# import Transana's Printout Class
import TranscriptPrintoutClass

# Define button constants
T_DATETIME       =  wx.NewId()
T_SAVEAS         =  wx.NewId()
T_PAGESETUP      =  wx.NewId()
T_PRINTPREVIEW   =  wx.NewId()
T_PRINT          =  wx.NewId()
T_HELP           =  wx.NewId()
T_EXIT           =  wx.NewId()
CMD_SEARCH_BACK_ID = wx.NewId()
CMD_SEARCH_NEXT_ID = wx.NewId()

class NoteEditor(wx.Dialog):
    """ This class provides a simple editor for notes that are attached to various Transana objects.  """

    def __init__(self, parent, default_text=""):
        """Initialize an NoteEditor object."""
        # Determine the screen size
        rect = wx.ClientDisplayRect()
        # We'll make our default window 60% of the size of the screen
        self.width = rect[2] * .60
        self.height = rect[3] * .60
        # Initialize a Dialog object to form our basic window
        wx.Dialog.__init__(self, parent, -1, _("Note"), wx.DefaultPosition,
                           wx.Size(self.width, self.height), wx.CAPTION | wx.CLOSE_BOX | wx.SYSTEM_MENU | wx.RESIZE_BORDER)
        # Specify the minimium size of this window.
        self.SetMinSize((400, 220))
        # Add a Note Editor Panel to the dialog
        self.pnl = _NotePanel(self, default_text)
        # Create a Sizer for the Dialog
        dlgSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Place the Panel on the Sizer, making it expandable in all directions
        dlgSizer.Add(self.pnl, 1, wx.EXPAND | wx.ALL, 2)
        # Set the Sizer to the Dialog
        self.SetSizer(dlgSizer)
        # Center the Dialog on the screen
        self.CenterOnScreen()

    def get_text(self):
        """Run the note editor and return the note string."""
        # Show the Note Editor Dialog modally
        self.ShowModal()
        # Return the text from the Editor Panel
        return self.pnl.get_text()

    # The _NotePanel requires a couple of  methods in the parent object so that its toolbar buttons
    # can function correctly in different contexts
    def OnHelp(self, event):
        """ Implement the Help function """
        # If a global Menu Window is defined defined ...
        if (TransanaGlobal.menuWindow != None):
            # ... call Help through its Control Object!  The Help Context is the Notes Editor
            TransanaGlobal.menuWindow.ControlObject.Help("Notes")

    def OnClose(self, event):
        """ Close the Notes Dialog based on instructions from the embedded Panel """
        # Just close.  That's all we need to do here.  The save is handled by the calling routine.
        self.Close()


class _NotePanel(wx.Panel):
    """ This is a Panel-based embeddable plain text editor.
          The calling routine needs to implement an OnHelp() method to operationalize the Help Button.
          The calling routine needs to implement an OnClose() method to close the panel's container window. """

    def __init__(self, parent, default_text=""):
        """ Initialize and populate the Note Editor panel """
        # Due to an odd behavior on the part of wxTextCtrl (it selects all text upon initialization and ignores
        # SetSelection requests), we need to track to see if we have done the initial SetSelection.  We haven't.
        self.initialized = False
        # Remember your ancestors.
        self.parent = parent
        # Get the global print data
        self.printData = TransanaGlobal.printData
        # Create the Panel
        wx.Panel.__init__(self, parent, style = wx.RAISED_BORDER)
        # Give the Panel a Sizer
        pnlSizer = wx.BoxSizer(wx.VERTICAL)
        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)
        # Place a Tool Bar on the Panel
        self.toolbar = wx.ToolBar(self, style = wx.TB_HORIZONTAL | wx.TB_TEXT)   # wx.RAISED_BORDER | 
        # Add an Insert Date/Time button to the Toolbar
        self.toolbar.AddTool(T_DATETIME, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Time16.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Insert Date / Time'))
        # Add a Save As Text button to the Toolbar
        self.toolbar.AddTool(T_SAVEAS, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "SaveTXT16.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Save As'))
        # Add a Page Setup button to the toolbar
        self.toolbar.AddTool(T_PAGESETUP, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "PrintSetup.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Page Setup'))
        # Add a Print Preview button to the Toolbar
        self.toolbar.AddTool(T_PRINTPREVIEW, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "PrintPreview.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Print Preview'))

        # Disable Print Preview on the Mac
        if 'wxMac' in wx.PlatformInfo:
            self.toolbar.EnableTool(T_PRINTPREVIEW, False)
            
        # Add a Print button to the Toolbar
        self.toolbar.AddTool(T_PRINT, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Print.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Print'))
        # Get the graphic for Help ...
        bmp = wx.ArtProvider_GetBitmap(wx.ART_HELP, wx.ART_TOOLBAR, (16,16))
        # ... and create a bitmap button for the Help button
        self.toolbar.AddTool(T_HELP, bmp, shortHelpString=_("Help"))
        # Add an Exit button to the Toolbar
        self.toolbar.AddTool(T_EXIT, wx.Bitmap(os.path.join(TransanaGlobal.programDir, "images", "Exit.xpm"), wx.BITMAP_TYPE_XPM), shortHelpString=_('Exit'))
        # Adding a separator here helps things look better on the Mac.
        self.toolbar.AddSeparator()
        # Cause the Toolbar to be built
        self.toolbar.Realize()

        # Create a horizontal Sizer for the toolbar
        hsizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the Toolbar to the Sizer
        hsizer.Add(self.toolbar)
        # Add a space to the Sizer
        hsizer.Add((20, 1))
        # Add Quick Search tools
        # Get the icon for the Search Backwards button
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_BACK, wx.ART_TOOLBAR, (16,16))
        # Create the Bitmap Button for Search Back
        self.searchBack = wx.BitmapButton(self, CMD_SEARCH_BACK_ID, bmp, style=wx.NO_BORDER)
        # Add this button to the Toolbar Sizer
        hsizer.Add(self.searchBack, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        # Connect the button to the OnSearch Method
        wx.EVT_BUTTON(self, CMD_SEARCH_BACK_ID, self.OnSearch)
        # Add a spacer to the Toolbar Sizer
        hsizer.Add((10, 1))
        # Create a Search Text control
        self.searchText = wx.TextCtrl(self, -1, size=(100, 20), style=wx.TE_PROCESS_ENTER)
        # Add it to the Toolbar Sizer
        hsizer.Add(self.searchText, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        # Call OnSearch on Enter from within the searchText box
        self.Bind(wx.EVT_TEXT_ENTER, self.OnSearch, self.searchText)
        # Add a spacer
        hsizer.Add((10, 1))
        # Get the icon for the Search Forwards button
        bmp = wx.ArtProvider_GetBitmap(wx.ART_GO_FORWARD, wx.ART_TOOLBAR, (16,16))
        # Create the Bitmap Button for Search Back
        self.searchNext = wx.BitmapButton(self, CMD_SEARCH_NEXT_ID, bmp, style=wx.NO_BORDER)
        # Add this button to the Toolbar Sizer
        hsizer.Add(self.searchNext, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        # Connect the button to the OnSearch Method
        wx.EVT_BUTTON(self, CMD_SEARCH_NEXT_ID, self.OnSearch)
        # Add the Toolbar Sizer to the Panel Sizer
        pnlSizer.Add(hsizer)

        # add the note editing widget to the panel.  User a multi-line TextCtrl, and TE_RICH style to enable
        # font size change on Windows.
        self.txt = wx.TextCtrl(self, -1, style=wx.TE_MULTILINE | wx.TE_RICH)
        # Get the Default Style
        txtStyle = self.txt.GetDefaultStyle()
        # Get the Default Font from the Default Style
        self.txtFont = txtStyle.GetFont()
        # On Windows ...
        if 'wxMSW' in wx.PlatformInfo:
            # ... 10 point looks about right.  12 is too big.  (default is 8)
            fontSize = 10
        # On Mac ...
        else:
            # ... 12 point looks about right.  (default is 11)
            fontSize = 12
        # If that doesn't work, as it doesn't on Windows ...
        if not self.txtFont.IsOk():
            # ... just create a Default Font with point size 10 on Windows, 12 on Mac to look right.
            self.txtFont = wx.Font(pointSize=fontSize, family = wx.DEFAULT, style = wx.NORMAL, weight = wx.NORMAL)
        # If we did get the default font ...
        else:
            # ... change it to the desired size
            self.txtFont.SetPointSize(fontSize)
        # Apply the Font to the Style
        txtStyle.SetFont(self.txtFont)
        # Set the Style in the Text Control
        self.txt.SetDefaultStyle(txtStyle)
        # If there is default text ...    (Windows requires this conditional.  Otherwise, the font size is
        #                                  wrong on new Notes!)
        if default_text != "":
            # ... add the existing text to the note's text control
            self.txt.WriteText(default_text)

        # We want to trap a couple of key combinations
        self.txt.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        # Add the Text Ctrl to the Panel Sizer
        pnlSizer.Add(self.txt, 1, wx.EXPAND | wx.ALL, 2)
        # Define the Panel Sizer as the main Sizer
        self.SetSizer(pnlSizer)
        # Fit the controls to the Panel
        self.Fit()

        # Define the Methods for the Toolbar Buttons
        wx.EVT_MENU(self, T_DATETIME, self.OnDateTime)
        wx.EVT_MENU(self, T_SAVEAS, self.OnSaveAs)
        wx.EVT_MENU(self, T_PAGESETUP, self.OnPageSetup)
        wx.EVT_MENU(self, T_PRINTPREVIEW, self.OnPrintPreview)
        wx.EVT_MENU(self, T_PRINT, self.OnPrint)
        wx.EVT_MENU(self, T_HELP, self.OnHelp)
        wx.EVT_MENU(self, T_EXIT, self.OnClose)

        # Define the SetFocus Event for the Text Control.  This is where the SetSelection can get called so it will work.
        wx.EVT_SET_FOCUS(self.txt, self.OnSetFocus)
        # Put the focus on the Text Control (needed for Mac)
        self.txt.SetFocus()

    def set_text(self, text):
        """ Place text in the editor control. """
        self.txt.SetValue(text)

    def get_text(self):
        """ Get the note text. """
        return self.txt.GetValue()

    def isChanged(self):
        """ Determine whether the Text has been edited. """
        return self.txt.IsModified()

    def EnableControls(self, enable):
        """ Change the "Enable" status of Note Editor Panel controls """
        # Change the status of the Text Ctrl
        self.txt.Enable(enable)
        # Change the status of the Toolbar Buttons
        self.toolbar.EnableTool(T_DATETIME, enable)
        self.toolbar.EnableTool(T_SAVEAS, enable)
        self.toolbar.EnableTool(T_PAGESETUP, enable)

        # Disable Print Preview on the Mac.
        if not 'wxMac' in wx.PlatformInfo:
            self.toolbar.EnableTool(T_PRINTPREVIEW, enable)
            
        self.toolbar.EnableTool(T_PRINT, enable)
        # Change the status of the Search tool elements
        self.searchBack.Enable(enable)
        self.searchText.Enable(enable)
        self.searchNext.Enable(enable)

    def SetSearchText(self, searchText):
        """ Insert a string into the Search Text field """
        # We only do this if searchText is ENABLED
        if self.searchText.IsEnabled():
            # Insert the specified value into the text control
            self.searchText.SetValue(searchText)
            # Then set the focus to the text control, ready to search
            self.searchText.SetFocus()

    def OnSearch(self, event):
        """ Search the Note Text for a string """
        # Get the string to search in
        text = self.txt.GetValue().upper()
        # On Windows ...  THIS IS NO LONGER NEEDED AS OF 2.30 RELEASE
#        if 'wxMSW' in wx.PlatformInfo:
            # ... we need to adjust for the 2-character newline character.
#            text = text.replace('\n', '  ')
        # Get the string to search for
        searchText = self.searchText.GetValue().upper()
        # Set the program focus to the Note Text
        self.txt.SetFocus()
        # Get the current insertion point position
        pos = self.txt.GetInsertionPoint()
        # If we're supposed to search backwards ...
        if event.GetId() == CMD_SEARCH_BACK_ID:
            # First determine if there is an instance of the search string prior to the current insertion point.
            if text.rfind(searchText, 0, pos) >= 0:
                # If so, select the HIGHEST value of the search string (using rfind(), not find())
                self.txt.SetSelection(text.rfind(searchText, 0, pos),
                                      text.rfind(searchText, 0, pos) + len(searchText))
        # If we're supposed to search forwards ...
        elif (event.GetId() == CMD_SEARCH_NEXT_ID) or (event.GetId() == self.searchText.GetId()):
            # If we don't already have a search term selected ...
            if self.txt.GetStringSelection().upper() != searchText:
                # ... start at the Insertion Point
                startPos = pos
            # Otherwise, we probably have the search text selected ...
            else:
                # ... so we need to start on past the insertion point to find the NEXT instance
                startPos = pos + 1
            # First determine if there is an instance of the search string following (but not including) the current insertion point.
            if text.find(searchText, startPos) >= 0:
                # if so, select the next instance of the search string
                self.txt.SetSelection(text.find(searchText, startPos),
                                      text.find(searchText, startPos) + len(searchText))
        

    def OnKeyDown(self, event):
        """ Capture key strokes within the Note Editor """
        # Assume we want to call event.Skip() unless otherwise noted
        callSkip = True
        # Get the Key Code of the key being pressed
        c = event.GetKeyCode()
        # Trap exceptions raised by non-ASCII keys being pressed
        try:
            # Detect that the Ctrl key is being held down
            if event.ControlDown():
                # Ctrl-T should insert a Date / Time Stamp
                if c == ord('T'):
                    self.OnDateTime(event)
                    # We skip event.Skip(), which beeps.
                    callSkip = False
            else:
                # ESC should exit the Note Edit control
                if c == wx.WXK_ESCAPE:
                    self.OnClose(event)
                    # We skip event.Skip(), which beeps.
                    callSkip = False
        except:
            pass  # Non-ASCII value key pressed
        # If we still want to call event.Skip() so that non-handled keystrokes are processed ...
        if callSkip:
            # ... then call it already!
            event.Skip()
            
    def OnDateTime(self, event):
        """ Insert the current Date & Time into a Note """
        # Get the current Date / Time information from the system
        (year, month, day, hour, minute, second, weekday, yearday, dst) = time.localtime()
        # Are we in the morning?
        ampm = 'am'
        # Let's use 12-hour time.  If it's afternoon ...
        if hour > 12:
            # ... decrement the hour value and signal that it's afternoon
            hour -= 12
            ampm = 'pm'
        # Add the Date / Time stamp to the Note Text
        if TransanaConstants.singleUserVersion:
            # TODO:  Localize this!
            self.txt.WriteText("%s/%s/%s  %s:%02d:%02d %s\n" % (month, day, year, hour, minute, second, ampm))
        else:
            # If multi-user, include the username!
            # TODO:  Localize this!
            self.txt.WriteText("%s/%s/%s  %s:%02d:%02d %s - %s\n" % (month, day, year, hour, minute, second, ampm, TransanaGlobal.userName))

    def OnSaveAs(self, event):
        """Export the note to a TXT file."""
        # Create a File Dialog for saving a TXT file
        dlg = wx.FileDialog(self, wildcard="*.txt", style=wx.SAVE)
        # Display the dialog and get the user input
        if dlg.ShowModal() == wx.ID_OK:
            # Get the File name
            fname = dlg.GetPath()
            # Mac doesn't automatically append the file extension.  Do it if necessary.
            if not fname.upper().endswith(".TXT"):
                fname += '.txt'
            # Check to see if the file already exists ...
            if os.path.exists(fname):
                # ... and if so, build an error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('A file named "%s" already exists.  Do you want to replace it?'), 'utf8')
                else:
                    prompt = _('A file named "%s" already exists.  Do you want to replace it?')
                # Build an error dialog
                dlg2 = Dialogs.QuestionDialog(None, prompt % fname)
                # If the user chooses to overwrite ...
                if dlg2.LocalShowModal() == wx.ID_YES:
                    # ... Export the data to the file
                    self.txt.SaveFile(fname)
                # Destroy the error dialog
                dlg2.Destroy()
            # If the specified file doesn't already exist ...
            else:
                # ... export the data to the file
                self.txt.SaveFile(fname)
        # Destroy the File Dialog
        dlg.Destroy()
        
    def OnPageSetup(self, event):
        """ Page Setup method """
        # Let's use PAGE Setup here ('cause you can do Printer Setup from Page Setup.)  It's a better system
        # that allows Landscape on Mac.
        # Create a PageSetupDialogData object based on the global printData defined in __init__
        pageSetupDialogData = wx.PageSetupDialogData(self.printData)
        # Calculate the paper size from the paper ID (Obsolete?)
        pageSetupDialogData.CalculatePaperSizeFromId()
        # Create a Page Setup Dialog based on the Page Setup Dialog Data
        pageDialog = wx.PageSetupDialog(self, pageSetupDialogData)
        # Show the Page Dialog box
        pageDialog.ShowModal()
        # Extract the print data from the page dialog
        self.printData = wx.PrintData(pageDialog.GetPageSetupData().GetPrintData())
        # reflect the print data changes globally
        TransanaGlobal.printData = self.printData
        # Destroy the Page Dialog
        pageDialog.Destroy()

    def OnPrintPreview(self, event):
        """ Define the method that implements Print Preview """
        # Convert the Note Text (in plain text format) into the form needed for the
        # TranscriptPrintoutClass's Print Preview display.  (This creates graphic and pageData)
        (graphic, pageData) = TranscriptPrintoutClass.PrepareData(TransanaGlobal.printData, noteTxt = self.txt.GetValue())
        # Pass the graph can data obtained from PrepareData() to the Print Preview mechanism TWICE,
        # once for the preview and once for the printer.
        printout = TranscriptPrintoutClass.MyPrintout('', graphic, pageData)
        printout2 = TranscriptPrintoutClass.MyPrintout('', graphic, pageData)
        # use wxPython's PrintPreview object to display the Print Preview.
        self.preview = wx.PrintPreview(printout, printout2, self.printData)
        # Check to see if the Print Preview was properly created.  
        if not self.preview.Ok():
            # If not, display an error message and exit
            self.SetStatusText(_("Print Preview Problem"))
            return
        # Calculate the best size for the Print Preview window
        theWidth = max(wx.ClientDisplayRect()[2] - 180, 760)
        theHeight = max(wx.ClientDisplayRect()[3] - 200, 560)
        # Create the dialog to hold the wx.PrintPreview object
        frame2 = wx.PreviewFrame(self.preview, TransanaGlobal.menuWindow, _("Print Preview"), size=(theWidth, theHeight))
        frame2.Centre()
        # Initialize the frame so it will display correctly
        frame2.Initialize()
        # Finally, we actually show the frame!
        frame2.Show(True)

    def OnPrint(self, event):
        """ Define the method that implements Print """
        # Convert the Note Text (in plain text format) into the form needed for the
        # TranscriptPrintoutClass's Print Preview display.  (This creates graphic and pageData)
        (graphic, pageData) = TranscriptPrintoutClass.PrepareData(TransanaGlobal.printData, noteTxt = self.txt.GetValue())
        # Pass the graph can data obtained from PrepareData() to the Print Preview mechanism ONCE
        printout = TranscriptPrintoutClass.MyPrintout('', graphic, pageData)
        # Create a Print Dialog Data object
        pdd = wx.PrintDialogData()
        # Populate the Print Dialog Data with the global print data
        pdd.SetPrintData(self.printData)
        # Define a wxPrinter object with the Print Dialog Data
        printer = wx.Printer(pdd)
        # Send the output to the printer.  If there's a problem ...
        if not printer.Print(self, printout):
            # ... create and display an error message
            dlg = Dialogs.ErrorDialog(None, _("There was a problem printing this report."))
            dlg.ShowModal()
            dlg.Destroy()
        # Finally, destroy the printout object.
        printout.Destroy()
        
    def OnHelp(self, event):
        """ Implement the Help button """
        # Pass processing this request up to the Parent, as this panel doesn't know which Help context to use.
        self.parent.OnHelp(event)
        
    def OnClose(self, event):
        """ Implement the Close button """
        # pass processing this request up to the Parent.  This panel doesn't know how to close the parent object.
        self.parent.OnClose(event)

    def OnSetFocus(self, event):
        """ Handle the TextCtrl receiving program focus """
        # This method was suggested by Robin Dunn via the wxPython-users mailing list.
        # Essentially, to get the SetSelection to work, we have to call it after the control receives focus.

        # If we have not done this before ...
        if not self.initialized:
            # Set up the SetSelection command for the initial entry into the control
            wx.CallAfter(self.txt.SetSelection, 0, 0)
            # Then indicate that we've done this so we won't do it again.  Otherwise, the
            # cursor moves to the beginning of the text every time we focus on this control
            # instead of only the first.
            self.initialized = True
