#Copyright (C) 2002-2014  The Board of Regents of the University of Wisconsin System

#This program is free software; you can redistribute it and/or
#modify it under the terms of the GNU General Public License
#as published by the Free Software Foundation; either version 2
#of the License, or (at your option) any later version.

#This program is distributed in the hope that it will be useful,
#but WITHOUT ANY WARRANTY; without even the implied warranty of
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#GNU General Public License for more details.

#You should have received a copy of the GNU General Public License
#along with this program; if not, write to the Free Software
#Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

"""This module implements a utility class underlying all text-based reports in Transana.
   This infrastructure provides report editing, saving to RTF, and print/print preview functionality.

   To use this class, create a TextReport object with the following parameters:
       title          The title of the report
       displayMethod  The parent method that populates the text for the report
       filterMethod   The parent method that implements the Filter Dialog call  (If skipped, the Filter will be unavailable)
       helpContext    The Help Context flag for the Help system (if skipped, Help will not be available)

       When the object has been created, you THEN need to call the CallDisplay() method to populate the report.

       For example:
        
            self.report = TextReport(self, title=_("Test Report"), displayMethod=self.OnDisplay, filterMethod=self.OnFilter,
                                     helpContext="Keyword Map")
            self.report.CallDisplay()

       See the stand-alone code at the bottom of this file for a working example. (Help doesn't work in the example.)  """

__author__ = "David K. Woods <dwoods@wcer.wisc.edu>"

DEBUG = False
if DEBUG:
    print "TextReport DEBUG is ON!!"

# import Python's os and sys modules
import os, sys
# import Python's time module
import time
# load wxPython for GUI
import wx
# import the wx.richtext framework
import wx.richtext as richtext

# If running stand-alone ...
if __name__ == '__main__':
    # This module expects i18n.  Enable it here.
    __builtins__._ = wx.GetTranslation

# Import Transana's Dialogs
import Dialogs
# Import Transana's Filter Dialog
import FilterDialog
# import Transana Miscellaneous functions
import Misc
# Import Transana's Constants
import TransanaConstants
# import Transana's Globals
import TransanaGlobal
# Import Transana's Images
import TransanaImages
# import Transana's Transcript Object to facilitate printing
import Transcript
# import Transana's Transcript Editor.  This forms the base of the editable report.
if TransanaConstants.USESRTC:
    import TranscriptEditor_RTC
else:
    import TranscriptEditor_STC
# Load the Printout Class
import TranscriptPrintoutClass
# import python's platform module
import platform

# Declare Control IDs
# Menu Item and Toolbar Item for File > Filter
M_FILE_FILTER        =  wx.NewId()
T_FILE_FILTER        =  wx.NewId()
# Menu Item and Toolbar Item for File > Edit
M_FILE_EDIT          =  wx.NewId()
T_FILE_EDIT          =  wx.NewId()
# Menu Item and Toolbar Item for File > Font
M_FILE_FONT          =  wx.NewId()
T_FILE_FONT          =  wx.NewId()
# Menu Item and Toolbar Item for File > Save As
M_FILE_SAVEAS        =  wx.NewId()
T_FILE_SAVEAS        =  wx.NewId()
# Menu Item and Toolbar Item for File > Printer Setup
M_FILE_PRINTSETUP    =  wx.NewId()
T_FILE_PRINTSETUP    =  wx.NewId()
# Menu Item and Toolbar Item for File > Print Preview
M_FILE_PRINTPREVIEW  =  wx.NewId()
T_FILE_PRINTPREVIEW  =  wx.NewId()
# Menu Item and Toolbar Item for File > Print 
M_FILE_PRINT         =  wx.NewId()
T_FILE_PRINT         =  wx.NewId()
# Menu Item and Toolbar Item for File > Exit
M_FILE_EXIT          =  wx.NewId()
T_FILE_EXIT          =  wx.NewId()
# Menu Item and Toolbar Item for Help > Help
M_HELP_HELP          =  wx.NewId()
T_HELP_HELP          =  wx.NewId()
# Toolbar Items for Search buttons
T_SEARCH_BACK        =  wx.NewId()
T_SEARCH_FORWARD     =  wx.NewId()

class TextReport(wx.Frame):
    """ This is the main class for the Text Report infrastrucure.
        This infrastructure provides report editing, RTF export, and print preview/print services for
        text-based reports to be used within Transana. """
    def __init__(self, parent, id=-1, title="", displayMethod=None, filterMethod=None, helpContext=None):
        """ Initialize the report framework.  Parameters are:
              parent               The report frame's parent
              id=-1                The report frame's ID
              title=""             The report title
              displayMethod=None   The method in the parent object that implements populating the report.
                                   (If left off, the report cannot function!)
              filterMethod=None    The method in the parent object that implements the Filter Dialog call.
                                   (If left off, no Filter option will be displayed.)
              helpContext=None     The Transana Help Context text for context-sensitive Help.
                                   (If left off, no Help option will be displayed.) """
        # It's always important to remember your ancestors.  (And the passed parameters!)
        self.parent = parent
        self.title = title
        self.reportNumber = 0
        self.helpContext = helpContext
        self.displayMethod = displayMethod
        self.filterMethod = filterMethod
        # Default the Control Object to None
        self.controlObject = None
        # Determine the screen size for setting the initial dialog size
        rect = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()  # wx.ClientDisplayRect()
        width = rect[2] * .80
        height = rect[3] * .80
        # Create the basic Frame structure with a white background
        wx.Frame.__init__(self, parent, id, title, size=wx.Size(width, height), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL | wx.NO_FULL_REPAINT_ON_RESIZE)
        self.SetBackgroundColour(wx.WHITE)
        
        # Set the report's icon
        transanaIcon = wx.Icon(os.path.join(TransanaGlobal.programDir, "images", "Transana.ico"), wx.BITMAP_TYPE_ICO)
        self.SetIcon(transanaIcon)
        # You can't have a separate menu on the Mac, so we'll use a Toolbar
        self.toolBar = self.CreateToolBar(wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_TEXT)
        # If a Filter Method is defined ...
        if self.filterMethod != None:
            # ... get the graphic for the Filter button ...
            bmp = wx.ArtProvider_GetBitmap(wx.ART_LIST_VIEW, wx.ART_TOOLBAR, (16,16))
            # ... and create a Filter button on the tool bar.
            self.toolBar.AddTool(T_FILE_FILTER, bmp, shortHelpString=_("Filter"))
        # Add an Edit button  to the Toolbar
        self.toolBar.AddTool(T_FILE_EDIT, TransanaImages.ReadOnly16.GetBitmap(), isToggle=True, shortHelpString=_('Edit/Read-only select'))
        # ... get the graphic for the Format button ...
        bmp = wx.ArtProvider_GetBitmap(wx.ART_HELP_SETTINGS, wx.ART_TOOLBAR, (16,16))
        # ... and create a Format button on the tool bar.
        self.toolBar.AddTool(T_FILE_FONT, bmp, shortHelpString=_("Format"))
        # Disable the Font button
        self.toolBar.EnableTool(T_FILE_FONT, False)
        # Add a Save button to the Toolbar
        self.toolBar.AddTool(T_FILE_SAVEAS, TransanaImages.Save16.GetBitmap(), shortHelpString=_('Save As'))
        # Add a Print (page) Setup button to the toolbar
        self.toolBar.AddTool(T_FILE_PRINTSETUP, TransanaImages.PrintSetup.GetBitmap(), shortHelpString=_('Set up Page'))
        # Add a Print Preview button to the Toolbar
        self.toolBar.AddTool(T_FILE_PRINTPREVIEW, TransanaImages.PrintPreview.GetBitmap(), shortHelpString=_('Print Preview'))

        # Disable Print Preview on the PPC Mac and for Right-To-Left languages
        if (platform.processor() == 'powerpc') or (TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft):
            self.toolBar.EnableTool(T_FILE_PRINTPREVIEW, False)

        # Add a Print button to the Toolbar
        self.toolBar.AddTool(T_FILE_PRINT, TransanaImages.Print.GetBitmap(), shortHelpString=_('Print'))
        # If a help context is defined ...
        if self.helpContext != None:
            # ... get the graphic for Help ...
            bmp = wx.ArtProvider_GetBitmap(wx.ART_HELP, wx.ART_TOOLBAR, (16,16))
            # ... and create a bitmap button for the Help button
            self.toolBar.AddTool(T_HELP_HELP, bmp, shortHelpString=_("Help"))
        # Add an Exit button to the Toolbar
        self.toolBar.AddTool(T_FILE_EXIT, TransanaImages.Exit.GetBitmap(), shortHelpString=_('Exit'))
        # Add a toolbar separator
        self.toolBar.AddSeparator()
        # Add the Search Backwards button
        self.toolBar.AddTool(T_SEARCH_BACK, wx.ArtProvider_GetBitmap(wx.ART_GO_BACK, wx.ART_TOOLBAR, (16,16)), shortHelpString=_('Search backwards'))
        # Create a text box for search terms, with the toolbar as its parent.
        self.searchText = wx.TextCtrl(self.toolBar, -1, size=(100, 20), style=wx.TE_PROCESS_ENTER)
        # Add the text box to the toolbar.
        self.toolBar.AddControl(self.searchText)
        # Add the Search Forwards button
        self.toolBar.AddTool(T_SEARCH_FORWARD, wx.ArtProvider_GetBitmap(wx.ART_GO_FORWARD, wx.ART_TOOLBAR, (16,16)), shortHelpString=_('Search forwards'))
        # Actually create the Toolbar
        self.toolBar.Realize()
        # Let's go ahead and keep the menu for non-Mac platforms
        if not '__WXMAC__' in wx.PlatformInfo:
            # Add a Menu Bar
            menuBar = wx.MenuBar()
            # Create the File Menu
            self.menuFile = wx.Menu()
            # If a Filter Method is defined ...
            if self.filterMethod != None:
                # ... add a Filter item to the File menu
                self.menuFile.Append(M_FILE_FILTER, _("&Filter"), _("Filter report contents"))
            # Add "Edit" to the File Menu
            self.menuFile.Append(M_FILE_EDIT, _("&Edit"), _("Edit the report manually"))
            # Add "Font" to the File Menu
            self.menuFile.Append(M_FILE_FONT, _("Format"), _("Change the current formatting"))
            # Disable the Font Menu Option
            self.menuFile.Enable(M_FILE_FONT, False)
            # Add "Save As" to File Menu
            self.menuFile.Append(M_FILE_SAVEAS, _("Save &As"), _("Save As"))
            # Add "Page Setup" to the File Menu
            self.menuFile.Append(M_FILE_PRINTSETUP, _("Page Setup"), _("Set up Page"))
            # Add "Print Preview" to the File Menu
            self.menuFile.Append(M_FILE_PRINTPREVIEW, _("Print Preview"), _("Preview your printed output"))
            # Add "Print" to the File Menu
            self.menuFile.Append(M_FILE_PRINT, _("&Print"), _("Send your output to the Printer"))
            # Add "Exit" to the File Menu
            self.menuFile.Append(M_FILE_EXIT, _("E&xit"), _("Exit the Keyword Map program"))
            # Add the File Menu to the Menu Bar
            menuBar.Append(self.menuFile, _('&File'))

            # If a Help Context is defined ...
            if self.helpContext != None:
                # ... create a Help menu ...
                self.menuHelp = wx.Menu()
                # ... add a Help item to the Help menu ...
                self.menuHelp.Append(M_HELP_HELP, _("&Help"), _("Help"))
                # ... and add the Help menu to the menu bar
                menuBar.Append(self.menuHelp, _("&Help"))
            # Connect the Menu Bar to the Frame
            self.SetMenuBar(menuBar)
        # Link menu items and toolbar buttons to the appropriate methods
        if self.filterMethod != None:
            wx.EVT_MENU(self, M_FILE_FILTER, self.OnFilter)                           # Attach File > Filter to a method
            wx.EVT_MENU(self, T_FILE_FILTER, self.OnFilter)                           # Attach Toolbar Filter to a method
        wx.EVT_MENU(self, M_FILE_EDIT, self.OnEdit)                                   # Attach OnEdit to File > Edit
        wx.EVT_MENU(self, T_FILE_EDIT, self.OnEdit)                                   # Attach OnEdit to Toolbar Edit
        wx.EVT_MENU(self, M_FILE_FONT, self.OnFont)                                   # Attach OnFont to File > Font
        wx.EVT_MENU(self, T_FILE_FONT, self.OnFont)                                   # Attach OnFont to Toolbar Font button
        wx.EVT_MENU(self, M_FILE_SAVEAS, self.OnSaveAs)                               # Attach File > Save As to a method
        wx.EVT_MENU(self, T_FILE_SAVEAS, self.OnSaveAs)                               # Attach Toolbar Save As to a method
        wx.EVT_MENU(self, M_FILE_PRINTSETUP, self.OnPrintSetup)                       # Attach File > Print Setup to a method
        wx.EVT_MENU(self, T_FILE_PRINTSETUP, self.OnPrintSetup)                       # Attach Toolbar Print Setup to a method
        wx.EVT_MENU(self, M_FILE_PRINTPREVIEW, self.OnPrintPreview)                   # Attach File > Print Preview to a method
        wx.EVT_MENU(self, T_FILE_PRINTPREVIEW, self.OnPrintPreview)                   # Attach Toolbar Print Preview to a method
        wx.EVT_MENU(self, M_FILE_PRINT, self.OnPrint)                                 # Attach File > Print to a method
        wx.EVT_MENU(self, T_FILE_PRINT, self.OnPrint)                                 # Attach Toolbar Print to a method
        wx.EVT_MENU(self, M_FILE_EXIT, self.CloseWindow)                              # Attach CloseWindow to File > Exit
        wx.EVT_MENU(self, T_FILE_EXIT, self.CloseWindow)                              # Attach CloseWindow to Toolbar Exit
        wx.EVT_MENU(self, T_SEARCH_BACK, self.OnSearch)                               # Attach Toolbar Search Backwards to Search method
        wx.EVT_MENU(self, T_SEARCH_FORWARD, self.OnSearch)                            # Attach Toolbar Search Forwards to Search method
        self.searchText.Bind(wx.EVT_TEXT_ENTER, self.OnSearch)                        # Attach Search Box Enter event to Search method
        if self.helpContext != None:
            wx.EVT_MENU(self, M_HELP_HELP, self.OnHelp)                               # Attach OnHelp to Help > Help
            wx.EVT_MENU(self, T_HELP_HELP, self.OnHelp)                               # Attach OnHelp to Toolbar Help
        # Define the Window's Close event, so we can remove the window from the Transana Interface
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # Add a Status Bar
        self.CreateStatusBar()
        # If we're using the RTC 
        if TransanaConstants.USESRTC:
            # Add a Rich Text Edit control to the Report Frame.  This is where the actual report text goes.
            self.reportText = TranscriptEditor_RTC.TranscriptEditor(self)
            # Clear / initialize the document
            self.reportText.ClearDoc()
##        # If we're using STC
##        else:
##            # Add a Styled Text Edit control to the Report Frame.  This is where the actual report text goes.
##            self.reportText = TranscriptEditor_STC.TranscriptEditor(self)
##            # Set report margins, the left margin to 1 inch, the right margin to 0 to prevent premature word wrap.
##            self.reportText.SetMargins(TranscriptPrintoutClass.DPI, 0)
        # We need to over-ride the reportText's EVT_RIGHT_UP method
        wx.EVT_RIGHT_UP(self.reportText, self.OnRightUp)
        # Initialize a variable to indicate whether a custom edit has occurred
        self.reportEdited = False
        # Get the global print data
        self.printData = TransanaGlobal.printData

        if parent == None:
            TransanaGlobal.CenterOnPrimary(self)

        # Show the Frame
        self.Show(True)


    def CallDisplay(self):
        """ Call the parent method (passed in during initialization) that populates the report """
        # Get the paper size from the TranscriptPrintoutClass
        (paperWidth, paperHeight) = TranscriptPrintoutClass.GetPaperSize()
        # Get the current height of the report window
        height = self.GetSize()[1]
        # Turn off the size hints, so we can resize the window
        self.SetSizeHints(-1, -1, -1, -1)
        # We need to adjust the size of the display slightly depending on platform.
        if '__WXMAC__' in wx.PlatformInfo:
            # GetPaperSize() is based on PRINTER, not SCREEN resolution.  We need to adjust this on the mac, changing from 72 DPI to 96 DPI
            paperWidth = int(paperWidth * 4 / 3)
        # Adjust the page size for the right margin and scroll bar
        sizeAdjust = 48
        # Set the size of the report so that it matches the width of the paper being used (plus the
        # line number area and the scroll bar)
        self.SetSize((paperWidth + self.reportText.lineNumberWidth + sizeAdjust, height))
        # Set Size Hints, so the report can't be resized.  (This may not be practical for large paper on small monitors.)
        self.SetSizeHints(self.GetSize()[0], int(self.GetSize()[1] * 0.8), self.GetSize()[0], -1)
        # Center on the screen
        self.CenterOnScreen()

        # CallDisplay() ALWAYS makes the report Read Only, so change the EDIT button state to NO EDIT ...
        self.toolBar.ToggleTool(T_FILE_EDIT, False)
        # ... and make the report read-only.
        self.reportText.SetReadOnly(True)
        # Disable the Font items
        self.toolBar.EnableTool(T_FILE_FONT, False)
        # The Menu does not exist on the Mac ...
        if not '__WXMAC__' in wx.PlatformInfo:
            # ... but if we're not on the Mac, disable it!
            self.menuFile.Enable(M_FILE_FONT, False)

        # Clear the Report
        self.reportText.ClearDoc()
        # If a Display Method has been defined ...
        if self.displayMethod != None:
            # ... call it.
            self.displayMethod(self.reportText)
        # Move the cursor to the beginning of the report
        self.reportText.GotoPos(0)
        # Bring the Report to the top so it will definitely be visible
        self.Raise()

    def OnFilter(self, event):
        """ Call the parent method (passed in during initialization) that implements the Filter Dialog """
        # If a Filter Method has been defined ...
        if (self.filterMethod != None):
            # Set a variable that signals the desire to continue (or not)
            contin = True
            # ... if the user has edited the report ...
            if self.reportEdited:
                # ... post a warning that edits will be lost ...
                msg = _("Please note that you will lose custom report edits when you use the Filter Dialog.")
                dlg = Dialogs.QuestionDialog(self, msg, _("Transana Information"), useOkCancel=True)
                # If the user DOESN'T press OK ...
                if dlg.LocalShowModal() != wx.ID_OK:
                    # ... signal that we don't want to continue with showing the Filter Dialog
                    contin = False
                dlg.Destroy()
            # ... and call the FilterMethod() if we're supposed to.  If it returns True ...
            if contin and self.filterMethod(event):
                # Disable GUI updates while populating the report text
                self.reportText.Freeze()
                # ... call the DisplayMethod() to update the report contents.
                self.CallDisplay()
                # Re-enable GUI Updates
                self.reportText.Thaw()
                # Since edits were lost, we can reset the report edit indicator.
                self.reportEdited = False

    def OnEdit(self, event):
        """ Toggle the report's Editability.  It is read-only initially. """
        # If the report is currently read-only ...
        if self.reportText.get_read_only():
            # ... change the button state ... 
            self.toolBar.ToggleTool(T_FILE_EDIT, True)
            # ... and make the report editable.
            self.reportText.SetReadOnly(False)
            # Enable the Font items
            self.toolBar.EnableTool(T_FILE_FONT, True)
            # The Menu does not exist on the Mac ...
            if not '__WXMAC__' in wx.PlatformInfo:
                # ... but if we're not on the Mac, enable it!
                self.menuFile.Enable(M_FILE_FONT, True)
            # Indicate that the report's been edited
            self.reportEdited = True
        # If the report is currently editable ...
        else:
            # ... change the button state ...
            self.toolBar.ToggleTool(T_FILE_EDIT, False)
            # ... and make the report read-only.
            self.reportText.SetReadOnly(True)
            # Disable the Font items
            self.toolBar.EnableTool(T_FILE_FONT, False)
            # The Menu does not exist on the Mac ...
            if not '__WXMAC__' in wx.PlatformInfo:
                # ... but if we're not on the Mac, disable it!
                self.menuFile.Enable(M_FILE_FONT, False)

    def OnFont(self, event):
        """ Change Font Characteristics for editing """
        # If we're using RTC ...
        if TransanaConstants.USESRTC:
            # ... use the RTC Format Dialog (font, paragraph, and tabs)
            self.reportText.CallFormatDialog()
        # If we're using STC ...
        else:
            # ... use the STC Font Dialog
            self.reportText.CallFontDialog()
            
    def OnSaveAs(self, event):
        """Export the report to an RTF file."""
        # Create a File Dialog for saving an RTF file
        dlg = wx.FileDialog(self, wildcard="*.rtf", style=wx.SAVE)
        # Display the dialog and get the user input
        if dlg.ShowModal() == wx.ID_OK:
            # Get the File name
            fname = dlg.GetPath()
            # Mac doesn't automatically append the file extension.  Do it if necessary.
            if not fname.upper().endswith(".RTF"):
                fname += '.rtf'
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
                dlg2.CentreOnScreen()
                # If the user chooses to overwrite ...
                if dlg2.LocalShowModal() == wx.ID_YES:
                    # ... Export the data to the file
                    self.reportText.export_transcript(fname)
                # Destroy the error dialog
                dlg2.Destroy()
            # If the specified file doesn't already exist ...
            else:
                # ... export teh data to the file
                self.reportText.export_transcript(fname)
        # Destroy the File Dialog
        dlg.Destroy()

    # Define the Method that implements Printer Setup
    def OnPrintSetup(self, event):
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
        # If no custom edits have been performed (which would be spoiled/lost by a call to CallDisplay()) ...
        if not self.reportEdited:
            self.reportText.Freeze()
            # ... adjust the present display for a potential new paper size
            self.CallDisplay()
            self.reportText.Thaw()

    def OnPrintPreview(self, event):
        """ Define the method that implements Print Preview """

        # If we're using the Rich Text Control ...
##        if TransanaConstants.USESRTC:
##            data = wx.PrintDialogData(self.printData)
##            printout = richtext.RichTextPrintout()
##            printout.SetRichTextBuffer(self.reportText.GetBuffer())
##            printout2 = richtext.RichTextPrintout()
##            printout2.SetRichTextBuffer(self.reportText.GetBuffer())
##            self.preview = wx.PrintPreview(printout, printout2, data)
##
##            if not self.preview.Ok():
##                self.log.WriteText("Houston, we have a problem...\n")
##                return
##
##            pfrm = wx.PreviewFrame(self.preview, self, "This is a print preview")
##
##            pfrm.Initialize()
##            pfrm.SetPosition(self.GetPosition())
##            pfrm.SetSize(self.GetSize())
##            pfrm.Show(True)





        # If we're using the Rich Text Control ...
        if TransanaConstants.USESRTC:
##            # The RichTextPrinting.PreviewBuffer() has a bug such that it crashes Transana.
##            # This is the way things SHOULD work, but of course, they don't.
##            if False:
##                # Define the FONT for the Header and Footer text
##                headerFooterFont = wx.Font(10, wx.FONTFAMILY_ROMAN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Times New Roman")
##                # Get the current date and time
##                t = time.localtime()
##                # Format the Footer Date as M/D/Y
##                footerDate = "%d/%d/%d" % (t.tm_mon, t.tm_mday, t.tm_year)
##                # Format the Footer Page Number
##                footerPage = _("Page") + " @PAGENUM@ " + _("of") + " @PAGESCNT@"
##                # Create a RichTextPrinting Object
##                printout = richtext.RichTextPrinting(_("Transana Report"), self)
##                # Let the printout know about the default printer settings
##                printout.SetPrintData(self.printData)
##                # Specify the Header / Footer Font
##                printout.SetHeaderFooterFont(headerFooterFont)
##                # Add the Report Title to the top right
##                printout.SetHeaderText(self.title, location=richtext.RICHTEXT_PAGE_RIGHT)
##                # Add the date to the bottom left
##                printout.SetFooterText(footerDate, location=richtext.RICHTEXT_PAGE_LEFT)
##                # Add the page number to the bottom right
##                printout.SetFooterText(footerPage, location=richtext.RICHTEXT_PAGE_RIGHT)
##                # Do NOT show Header and Footer on the First Page
##                printout.SetShowOnFirstPage(False)
##                # Preview the RTC Buffer Contents
##                printout.PreviewBuffer(self.reportText.GetBuffer())
##                # Destroy the RichTextPrinting Object
##                printout.Destroy()
##            # Because of the Preview Bug in RichTextPrinting, we have to do things a different,
##            # slightly harder way.  This way does not allow for Headers and Footers because
##            # richtext.RichTextHeaderFooterData() has not been wrapped in wxPython.
##            else:
                # Create the first Printout object using a RichTextPrintout object
                printout1 = richtext.RichTextPrintout()
                # Put the RichTextCtrl's Buffer into the RichTextPrintout object
                printout1.SetRichTextBuffer(self.reportText.GetBuffer())
                # Create the second Printout object using a RichTextPrintout object
                printout2 = richtext.RichTextPrintout()
                # Put the RichTextCtrl's Buffer into the RichTextPrintout object
                printout2.SetRichTextBuffer(self.reportText.GetBuffer())

                # Get a PrintDialogData object
                data = wx.PrintDialogData()
                # Populate the PrintDialogData Object with print data from Transana's configuration info
                data.SetPrintData(self.printData)

                # Create a PrintPreview Object using the printout objects and print data created above
                preview = wx.PrintPreview(printout1, printout2, data)
                # If this didn't work ...
                if not preview.Ok():

                    print "NEED ERROR MESSAGE HERE!!!!  PREVIEW FAILURE !!!!"

                    # ... just exit here.
                    return

                # Now create a Print Preview WINDOW using the Print Preview Object created above
                previewFrame = wx.PreviewFrame(preview, self, _("Transana Report"))
                # Initialize the Print Preview Window
                previewFrame.Initialize()

                # Okay, here's another wrinkle.  If you choose Print from the Print Preview, THAT messes up the
                # RTC formatting too.  Here's a workaround for that.

                # First, identify the "Print" button using its label
                printButton = previewFrame.FindWindowByLabel("&Print...")
                # If that failed ...
                if printButton == None:
                    # ... try again using its ID, which appears to be 4 pretty consistently for me.
                    printButton = previewFrame.FindWindowById(4)

                # If the Print button is found
                if printButton != None:
                    # hijack the button press, pointing it to the local OnPrint method
                    printButton.Bind(wx.EVT_BUTTON, self.OnPrint)
                # If the Print Button is NOT found ... (It's almost certainly due to i18n issues.  Is there another way??)
                else:

                    print "NEED ERROR MESSAGE HERE!!!!  DO NOT USE PRINT BUTTON!!!!"

                # Position the Print Preview Frame on top of the TextReport frame
                previewFrame.SetPosition(self.GetPosition())
                # Size the Print Preview Frame to match the TextReport Frame
                previewFrame.SetSize(self.GetSize())
                # We need to capture the Close Event for the Preview Window to compensate for a
                # problem with RichTextCtrl formatting.
##                previewFrame.Bind(wx.EVT_CLOSE, self.OnClosePreview)
                # Finally, we can show the Print Preview Frame
                previewFrame.Show(True)

####        # If we're using the Styled Text Control ...
####        else:
####            # Create a temporary Transcript Object for the print preview
####            tempTranscript = Transcript.Transcript()
####            # Put the RichTextEditCtrl's contents, in RTF from, into the Transcript Object
####            tempTranscript.text = self.reportText.GetRTFBuffer()
####            # Convert the temporary transcript object (and its RTF contents) into the form needed for the
####            # TranscriptPrintoutClass's Print Preview display.  (This creates graphic and pageData)
####            (graphic, pageData) = TranscriptPrintoutClass.PrepareData(TransanaGlobal.printData, tempTranscript)
####            # Pass the graph can data obtained from PrepareData() to the Print Preview mechanism TWICE,
####            # once for the preview and once for the printer.
####            printout = TranscriptPrintoutClass.MyPrintout('', graphic, pageData)
####            printout2 = TranscriptPrintoutClass.MyPrintout('', graphic, pageData)
####            # use wxPython's PrintPreview object to display the Print Preview.
####            self.preview = wx.PrintPreview(printout, printout2, self.printData)
####            # Check to see if the Print Preview was properly created.  
####            if not self.preview.Ok():
####                # If not, display an error message and exit
####                self.SetStatusText(_("Print Preview Problem"))
####                return
####            # Calculate the best size for the Print Preview window
####            theWidth = max(wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()[2] - 180, 760)  # wx.ClientDisplayRect()
####            theHeight = max(wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()[3] - 200, 560)  # wx.ClientDisplayRect()
####            # Create the dialog to hold the wx.PrintPreview object
####            frame2 = wx.PreviewFrame(self.preview, self, _("Print Preview"), size=(theWidth, theHeight))
####            frame2.Centre()
####            # Initialize the frame so it will display correctly
####            frame2.Initialize()
####            # Finally, we actually show the frame!
####            frame2.Show(True)
##
##    def OnClosePreview(self, event):
##        """ Handles processing when the Print Preview window is closed """
##        # Call parent event handler
##        event.Skip()
##        # The Print Preview has caused a problem in the display.  We need to tweak the frame size
##        # to get it to re-draw itself correctly.
##        # Get the current size
##        (w, h) = self.GetSizeTuple()
##        # Increase height by 1 pixel
##        self.SetSize(wx.Size(w, h+1))
##        # Return to original height.
##        self.SetSize(wx.Size(w, h))

    def OnPrint(self, event):
        """ Event Handler for Printing """
        # If we're using the Rich Text Control ...
        if TransanaConstants.USESRTC:
            # Define the FONT for the Header and Footer text
            headerFooterFont = wx.Font(10, wx.FONTFAMILY_ROMAN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Times New Roman")
            # Get the current date and time
            t = time.localtime()
            # Format the Footer Date as M/D/Y
            footerDate = "%d/%d/%d" % (t.tm_mon, t.tm_mday, t.tm_year)
            # Format the Footer Page Number
            footerPage = _("Page") + " @PAGENUM@ " + _("of") + " @PAGESCNT@"
            # Create a RichTextPrinting Object
            printout = richtext.RichTextPrinting(_("Transana Report"), self)
            # Let the printout know about the default printer settings
            printout.SetPrintData(self.printData)
            # Specify the Header / Footer Font
            printout.SetHeaderFooterFont(headerFooterFont)
            # Add the Report Title to the top right
            printout.SetHeaderText(self.title, location=richtext.RICHTEXT_PAGE_RIGHT)
            # Add the date to the bottom left
            printout.SetFooterText(footerDate, location=richtext.RICHTEXT_PAGE_LEFT)
            # Add the page number to the bottom right
            printout.SetFooterText(footerPage, location=richtext.RICHTEXT_PAGE_RIGHT)
            # Do NOT show Header and Footer on the First Page
            printout.SetShowOnFirstPage(False)
            # print the RTC Buffer Contents
            printout.PrintBuffer(self.reportText.GetBuffer())
            # Destroy the RichTextPrinting Object
            printout.Destroy()

##        # If we're using the Styled Text Control ...
##        else:
##            # Create a temporary Transcript Object for the print preview
##            tempTranscript = Transcript.Transcript()
##            # Put the RichTextEditCtrl's contents, in RTF from, into the Transcript Object
##            tempTranscript.text = self.reportText.GetRTFBuffer()
##            # Convert the temporary transcript object (and its RTF contents) into the form needed for the
##            # TranscriptPrintoutClass's Print display.  (This creates graphic and pageData)
##            (graphic, pageData) = TranscriptPrintoutClass.PrepareData(TransanaGlobal.printData, tempTranscript)
##            # Pass the graph can data obtained from PrepareData() to the Print Preview mechanism ONCE
##            printout = TranscriptPrintoutClass.MyPrintout('', graphic, pageData)
##            # Create a Print Dialog Data object
##            pdd = wx.PrintDialogData()
##            # Populate the Print Dialog Data with the global print data
##            pdd.SetPrintData(self.printData)
##            # Define a wxPrinter object with the Print Dialog Data
##            printer = wx.Printer(pdd)
##            # Send the output to the printer.  If there's a problem ...
##            if not printer.Print(self, printout):
##                # ... create and display an error message
##                dlg = Dialogs.ErrorDialog(None, _("There was a problem printing this report."))
##                dlg.ShowModal()
##                dlg.Destroy()
##            # Finally, destroy the printout object.
##            printout.Destroy()

    def OnHelp(self, event):
        """ Implement the Help function """
        # If a Help Window and the Help Context are defined ...
        if (TransanaGlobal.menuWindow != None) and (self.helpContext != None):
            # ... call Help!
            TransanaGlobal.menuWindow.ControlObject.Help(self.helpContext)

    def OnClose(self, event):
        """ Handle the Close Event """
        # If the report has a defined Control Object ...
        if self.controlObject != None:
            # ... remove this report from the Menu Window's Window Menu
            self.controlObject.RemoveReportWindow(self.title, self.reportNumber)

        # Inherit the parent Close event so things will, you know, close.
        event.Skip()

    def CloseWindow(self, event):
        """ Close the Report on File Exit """
        # Close the form
        self.Close()

    def OnSearch(self, event):
        """ Implement the Toolbar's QuickSearch """
        # Get the text for the search
        txt = self.searchText.GetValue()
        # If there is text ...
        if txt != '':
            # If searching forward ...
            if event.GetId() in [T_SEARCH_FORWARD, self.searchText.GetId()]:
                # ... set the search direction to next
                direction = "next"
            # If searching backwards ...
            elif event.GetId() == T_SEARCH_BACK:
                # ... set the search direction to backwards
                direction = "back"
            # Perform the search in the report text
            self.reportText.find_text(txt, direction)
            # Set the focus back on the report text component, rather than the button, so Paste or typing work.
            self.reportText.SetFocus()

    def GetCenterSpacer(self, style, text):
        """ The wxSTC control doesn't support paragraph centering.  This method fakes it.
            it returns the number of spaces needed to approximately center the text sent in in the style passed. """
        # If our text parameter is blank ...
        if text == '':
            # we can just skip this method.
            return ""
        # Determine HALF the width of the text
        textWidth = int(self.reportText.TextWidth(style, ' ' + text) / 2.0)
        # Initialize the space string
        centerSpacer = ''
        # Add spaces to the space string until we have enough to center the text
        while self.reportText.TextWidth(style, centerSpacer) + textWidth < int(self.reportText.GetSizeTuple()[0] / 2.0) - (TranscriptPrintoutClass.DPI * 1.5):
            centerSpacer += ' '
        # Return the space string
        return centerSpacer

    def OnRightUp(self, event):
        """ Over-ride the self.textReport object's EVT_RIGHT_UP method """
        # We don't want to do anything.  This exists to prevent TranscriptEditor's EVT_RIGHT_UP from being triggered.
        pass


if __name__ == '__main__':

    class TextReportTest(wx.Dialog):
        """ The following does an automatic self test of the TextReport functionality to demonstrate how to use this module. """
        def __init__(self):
            wx.Dialog.__init__(self, None)
            btn = wx.Button(self, -1, "Push Me!")
            btn.Bind(wx.EVT_BUTTON, self.onBtn)
            self.ShowModal()

        def onBtn(self, event):
            self.report = TextReport(self, title=_("Test Report"), displayMethod=self.OnDisplay, filterMethod=self.OnFilter,
                                     helpContext="Keyword Map")
            self.report.CallDisplay()

        def OnDisplay(self, reportText):

            print "No OnDisplay Method is defined for __main__ since RTC conversion"

##            print "self.OnDisplay()"
##
##            reportText.SetReadOnly(False)
##
##            reportText.InsertStyledText('This is a test.\n\n')
##
##            reportText.SetBold(True)
##
##            reportText.InsertStyledText('This is also a test.')
##
##            reportText.SetBold(False)
##
##            reportText.InsertStyledText('\n\nNow the test is done.\n\n')
##
##            reportText.InsertStyledText('This is another test.\n\n')
##
##            reportText.SetFont('Arial Black', 12, 0x008800, 0xFF00FF)
##
##            reportText.InsertStyledText('This is also a test.')
##
##            reportText.SetFont('Courier New', 10, 0x000000, 0xFFFFFF)
##
##            reportText.InsertStyledText('\n\nNow the test is done.')
##            reportText.SetReadOnly(True)

        def OnFilter(self, event):
            print "self.OnFilter() is not yet defined.  No filter is appropriate for this example."

            
    
    class MyApp(wx.App):
       def OnInit(self):
          frame = TextReportTest()
          self.SetTopWindow(frame)
          return True
          

    app = MyApp(0)
    app.MainLoop()
