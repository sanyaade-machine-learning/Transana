# Copyright (C) 2004 - 2007 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Collection Summary Report. """

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

import wx
import DBInterface
import Dialogs
import TransanaGlobal
import TranscriptPrintoutClass

class CollectionSummaryReport(wx.Object):
    """ Collection Summary Report """
    def __init__(self, dbTree, sel):
        # Set the Cursor to the Hourglass while the report is assembled
        TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))

        try:

            # Build the title and subtitle
            title = _('Collection Summary Report')
            if dbTree.GetPyData(sel).nodetype == 'SearchCollectionNode':
                reportType = _('Search Result Collection')
            else:
                reportType = _('Collection')
            prompt = '%s: %s'
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                reportType = unicode(reportType, 'utf8')
                prompt = unicode(prompt, 'utf8')
            subtitle = prompt % (reportType, dbTree.GetItemText(sel))
            
            # Prepare the Transcript for printing
            (graphic, pageData) = TranscriptPrintoutClass.PrepareData(TransanaGlobal.printData, collectionTree=dbTree, collectionNode=sel, title=title, subtitle=subtitle)

            # If there's data inthe report...
            if pageData != []:
                # Send the results of the PrepareData() call to the MyPrintout object, once for the print preview
                # version and once for the printer version.  
                printout = TranscriptPrintoutClass.MyPrintout(title, graphic, pageData, subtitle=subtitle)
                printout2 = TranscriptPrintoutClass.MyPrintout(title, graphic, pageData, subtitle=subtitle)
                
                # Create the Print Preview Object
                printPreview = wx.PrintPreview(printout, printout2, TransanaGlobal.printData)
                
                # Check for errors during Print preview construction
                if not printPreview.Ok():
                    dlg = Dialogs.ErrorDialog(None, _("Print Preview Problem"))
                    dlg.ShowModal()
                    dlg.Destroy()
                else:
                    # Create the Frame for the Print Preview
                    # The parent Frame is the global Menu Window
                    theWidth = max(wx.ClientDisplayRect()[2] - 180, 760)
                    theHeight = max(wx.ClientDisplayRect()[3] - 200, 560)
                    printFrame = wx.PreviewFrame(printPreview, TransanaGlobal.menuWindow, _("Print Preview"), size=(theWidth, theHeight))
                    printFrame.Centre()
                    # Initialize the Frame for the Print Preview
                    printFrame.Initialize()
                    # Restore Cursor to Arrow
                    TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
                    # Display the Print Preview Frame
                    printFrame.Show(True)
            else:
                # Restore Cursor to Arrow
                TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
                # If there are no clips to report, display an error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Collection "%s" has no Clips for the Collection Summary Report.'), 'utf8')
                else:
                    prompt = _('Collection "%s" has no Clips for the Collection Summary Report.')
                dlg = wx.MessageDialog(None, prompt % dbTree.GetItemText(sel), style = wx.OK | wx.ICON_EXCLAMATION)
                dlg.ShowModal()
                dlg.Destroy()

        finally:
            # Restore Cursor to Arrow
            TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
            
