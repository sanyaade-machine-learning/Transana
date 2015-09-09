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

"""This module implements the DataWindow class for the Data Display Objects."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case <nacase@wisc.edu>'

DEBUG = False
if DEBUG:
    print "DataWindow DEBUG is ON!"

import wx
from DatabaseTreeTab import *
import DocumentQuotesTab
import DataItemsTab
from KeywordsTab import *
import TransanaConstants
import TransanaGlobal

class DataWindow(wx.Dialog):
    """This class implements the window containing all data display tabs."""

    def __init__(self, parent, id=-1):
        """Initialize a DataWindow object."""
        # Start with a Dialog Box (wxPython)
        wx.Dialog.__init__(self, parent, id, _("Data"), self.__pos(),
                            self.__size(),
                            style=wx.CAPTION | wx.RESIZE_BORDER)

        # Set "Window Variant" to small only for Mac to use small icons
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        vSizer = wx.BoxSizer(wx.VERTICAL)

        # add a Notebook Control to the Dialog Box
        # The wxCLIP_CHILDREN style allegedly reduces flicker.
        self.nb = wx.Notebook(self, -1, style=wx.CLIP_CHILDREN)
        # Set the Notebook's background to White.  Otherwise, we get a visual anomoly on OS X with wxPython 2.9.4.0.
        self.nb.SetBackgroundColour(wx.Colour(255, 255, 255))
        
        # Let the notebook remember it's parent
        self.nb.parent = self
        vSizer.Add(self.nb, 1, wx.EXPAND)

        # Create tabs for the Notebook Control.  These tabs are complex enough that they are
        # instantiated as separate objects.

        self.DBTab = DatabaseTreeTab(self.nb)

        # Initialize the remaining Tabs to None for the moment
        self.DataItemsTab = None
        self.SelectedDataItemsTab = None
        self.KeywordsTab = None
        
        # Add the tabs to the Notebook Control
        self.nb.AddPage(self.DBTab, _("Database"), True)

        # OSX requires this for some reason or else it won't have a default
        # page selected.
        self.nb.SetSelection(0)

        # Handle Key Press Events
        self.nb.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)

        self.SetSizer(vSizer)
        self.SetAutoLayout(True)
        self.Layout()

        # If we are using the Multi-User Version ...
        if not TransanaConstants.singleUserVersion:
            # ... check to see if the Database is using SSL
            #     and select the appropriate graphic file ...
            if TransanaGlobal.configData.ssl:
                bitmap = TransanaImages.locked.GetBitmap()
            else:
                bitmap = TransanaImages.unlocked.GetBitmap()
            # Place a graphic on the screen.  This is NOT placed in the Sizer, but is placed on top of the notebook tabs.
            self.sslDBImage = wx.StaticBitmap(self, -1, bitmap, pos=(self.nb.GetSizeTuple()[0] - 32, 0), size=(16, 16))
            # Make the SSL Indicator clickable
            self.sslDBImage.Bind(wx.EVT_LEFT_DOWN, self.OnSSLClick)

            # ... check to see if both the Database and the Message Server are using SSL
            #     and select the appropriate graphic file ...
            if TransanaGlobal.configData.ssl and TransanaGlobal.chatIsSSL:
                bitmap = TransanaImages.locked.GetBitmap()
            else:
                bitmap = TransanaImages.unlocked.GetBitmap()
            # Place a graphic on the screen.  This is NOT placed in the Sizer, but is placed on top of the notebook tabs.
            self.sslImage = wx.StaticBitmap(self, -1, bitmap, pos=(self.nb.GetSizeTuple()[0] - 16, 0), size=(16, 16))
            # Make the SSL Indicator clickable
            self.sslImage.Bind(wx.EVT_LEFT_DOWN, self.OnSSLClick)
            # Remember the current SSL Status of the Message Server
            self.sslStatus = TransanaGlobal.configData.ssl and TransanaGlobal.chatIsSSL

        self.ControlObject = None            # The ControlObject handles all inter-object communication, initialized to None

        # Capture Size Changes
        wx.EVT_SIZE(self, self.OnSize)
        wx.EVT_PAINT(self, self.OnPaint)

        if DEBUG:
            print "DataWindow.__init__():  Initial size:", self.GetSize()

        # Capture the selection of different Notebook Pages
        wx.EVT_NOTEBOOK_PAGE_CHANGED(self, self.nb.GetId(), self.OnNotebookPageSelect)


    def AddItemsTab(self, libraryObj=None, dataObj=None):
        if self.DataItemsTab == None:
            if isinstance(dataObj, Episode.Episode):
                self.DataItemsTab = DataItemsTab.DataItemsTab(self.nb, libraryObj, dataObj)
                self.nb.AddPage(self.DataItemsTab, _("Episode Items"), False)
            else:
                self.DataItemsTab = DocumentQuotesTab.DocumentQuotesTab(self.nb, libraryObj, dataObj)
                self.nb.AddPage(self.DataItemsTab, _("Document Items"), False)
            # If a ControlObject is defined, propagate it to the DataItemsTab so that Clips can be loaded via double-clicking
            if self.ControlObject != None:
                self.DataItemsTab.Register(self.ControlObject)
            # Allow the Database Tab to redraw, as adding the Episode Clips tab can interfere with the appearance of the Data Window
            self.DBTab.Refresh()

    def AddSelectedItemsTab(self, libraryObj=None, dataObj=None, timeCode=-1, textPos=-1, textSel=(-2, -2)):
        if self.SelectedDataItemsTab == None:
            if timeCode > -1:
                self.SelectedDataItemsTab = DataItemsTab.DataItemsTab(self.nb, libraryObj, dataObj, timeCode)
            else:
                self.SelectedDataItemsTab = DocumentQuotesTab.DocumentQuotesTab(self.nb, libraryObj, dataObj, textPos, textSel)
            # If a ControlObject is defined, propagate it to the DataItemsTab so that Clips can be loaded via double-clicking
            if self.ControlObject != None:
                self.SelectedDataItemsTab.Register(self.ControlObject)
            self.nb.AddPage(self.SelectedDataItemsTab, _("Selected Items"), False)
            # Allow the Database Tab to redraw, as adding the Episode Clips tab can interfere with the appearance of the Data Window
            self.DBTab.Refresh()

    def AddKeywordsTab(self, seriesObj=None, episodeObj=None, documentObj=None, collectionObj=None, clipObj=None, quoteObj=None):
        """ Add the Keywords Tab to the Data Window """
        # If there is not currently a Keywords Tab ...
        if self.KeywordsTab == None:
            # ... create a keywords Tab
            self.KeywordsTab = KeywordsTab(self.nb, seriesObj, episodeObj, documentObj, collectionObj, clipObj, quoteObj)
            # Add the Tab to the Data Window Notebook
            self.nb.AddPage(self.KeywordsTab, _("Keywords"), False)
            # Allow the Database Tab to redraw, as adding the Keywords tab can interfere with the appearance of the Data Window
            self.DBTab.Refresh()
        # If a Keywords Tab already exists ...
        else:
            # ... display an error message for the PROGRAMMER!
            dlg = Dialogs.ErrorDialog(self, _('Problem creating the Keywords Tab!'))
            dlg.ShowModal()
            dlg.Destroy()

    def DeleteTabs(self):
        """ Delete all tabs in the Data Window except the DatabaseTreeTab, which should always be retained. """
        # On the Mac, double-clicking a Clip from the Episode Clips Tab of the Selected Clips Tab was causing
        # Transana to crash with either a Segmentation Fault or a Bus Error.  This is due to deleting 
        # the tab before it's done processing it's double-click events.  Therefore, let's detect the active tab
        # and put off deleting it until later.
        currentTab = self.nb.GetSelection()
        # Set the selection to the DatabaseTreeTab tab, which won't be deleted
        self.nb.SetSelection(0)

        # Under rare circumstances, this is called before some tabs are RENDERED and thus were not
        # getting deleted.  This fixes that!
        wx.YieldIfNeeded()
        
        # Delete all tabs but the DatabaseTreeTab.
        # Start with the Keywords Tab.  Check to see if it exists ...
        if (self.KeywordsTab != None) or (self.nb.GetPageCount() > 1):
            # ... and if it does, delete it ...
            self.nb.DeletePage(self.nb.GetPageCount() - 1)
            # ... and remove the reference to it.
            self.KeywordsTab = None

        # Do the Selected Episode Clips Tab next.  Check to see if it exists ...
        if self.SelectedDataItemsTab != None:
            # ... and if it does, delete it.  If it's the current tab, delete it later.  Otherwise, just delete it now ...
            if currentTab == 2:
                wx.CallAfter(self.DeleteTabLater)
            else:
                self.nb.DeletePage(2)
            # ... and remove the reference to it.
            self.SelectedDataItemsTab = None

        # Do the Episode Clips Tab next.  Check to see if it exists ...
        if self.DataItemsTab != None:
            # ... and if it does, delete it.  If it's the current tab, delete it later.  Otherwise, just delete it now ...
            if currentTab == 1:
                wx.CallAfter(self.DeleteTabLater)
            else:
                self.nb.DeletePage(1)
            # ... and remove the reference to it.
            self.DataItemsTab = None

    def DeleteTabLater(self):
        """ On the Mac, selecting a clip from the Episode Clips tab or the Selected Clips tab caused a
            Segment Fault or a Bus Error.  That was because the DataItemsTab tab was being deleted by
            the DeleteTabs() call above before its double-click method was done.  This method allows us
            to delay the deletion of the appropriate tab until after the tab's method is done.  """
        # The tab to be deleted is always tab 1, as the other tabs have been already deleted and the 
        # DatabaseTreeTab isn't supposed to be deleted.
        self.nb.DeletePage(1)

    def OnSize(self, event):
        """ Data Window Resize Event """
        # If we're not resizing ALL the Transana Windows ...  (avoid recursive calls!)
        if not TransanaGlobal.resizingAll:
            # Get the size of the Data Window
            (left, top) = self.GetPositionTuple()
            # Ask the Control Object to resize all other windows

            if DEBUG:
                print
                print "Call 2", 'Data', left, -1, top - 1

            self.ControlObject.UpdateWindowPositions('Data', left, YLower = top - 1)
        # Call to Layout() is required so that the Notebook Control resizes properly
        self.Layout()
        if not TransanaConstants.singleUserVersion:
            # Keep the SSL Image in the upper right corner
            self.sslDBImage.SetPosition((self.nb.GetSizeTuple()[0] - 32, 0))
            self.sslImage.SetPosition((self.nb.GetSizeTuple()[0] - 16, 0))

    def OnPaint(self, event):
        """ Local Paint Event to keep the SSL Image on top """
        # Call the normal Paint Event
        event.Skip()
        # If we're in the Multi-user version ...
        if not TransanaConstants.singleUserVersion:
            # ... bring the SSL Image to the top
            self.sslDBImage.Raise()
            self.sslImage.Raise()

    def OnNotebookPageSelect(self, event):
        """ Detect which tab in the Notebook is selected and prepare that tab for display. """
        # If Episode Clips Tab is selected ...
        if self.nb.GetPageText(event.GetSelection()) in [unicode(_("Episode Items"), 'utf8'), unicode(_("Document Items"), 'utf8')]:
            # ... get the latest Data for the Episode/Document Items
            self.DataItemsTab.Refresh()
        # If the Selected Clips Tab is selected ...
        elif self.nb.GetPageText(event.GetSelection()) == unicode(_("Selected Items"), 'utf8'):
            # If we have a Transcript (Episode) ...
            if isinstance(self.ControlObject.currentObj, Transcript.Transcript):
                # ... get the latest Data based on the current Video Position
                self.SelectedDataItemsTab.Refresh(self.ControlObject.GetVideoPosition())
            # If we have a Document ...
            elif isinstance(self.ControlObject.currentObj, Document.Document):
                # ... get the current Document position / selection ...
                (pos, sel) = self.ControlObject.GetDocumentPosition()
                # ... and get the latest Quote data based on the current selection / position
                self.SelectedDataItemsTab.Refresh(pos, sel)
        # If the Keyword Tab is selected ...
        elif self.nb.GetPageText(event.GetSelection()) == unicode(_('Keywords'), 'utf8'):
            # update the Keywords Tab in case something has changed since the objects were first loaded.
            self.KeywordsTab.Refresh()

    def Register(self, ControlObject=None):
        """ Register a ControlObject """
        self.ControlObject=ControlObject
        self.DBTab.Register(ControlObject=self.ControlObject)  # Propagate the Control Object registration
        
    def ClearData(self):
        """Clear the display."""
        # Remove any extra tabs that are displayed
        self.DeleteTabs()

    def UpdateDataWindow(self):
        """ Update the DataWindow, if needed """
        # If the Episode Clip Tab (which also shows Document Quotes) exists ...
        if (self.DataItemsTab != None):
            # ... update it!
            self.DataItemsTab.Refresh()

    def UpdateSSLStatus(self, sslValue):
        """ Update the SSL Status Indicator """

        # If we are using the Multi-User Version ...
        if not TransanaConstants.singleUserVersion:
            # ... check to see if both the Database and the Message Server are using SSL
            #     and select the appropriate graphic file ...
            if TransanaGlobal.configData.ssl:
                bitmap = TransanaImages.locked.GetBitmap()
            else:
                bitmap = TransanaImages.unlocked.GetBitmap()
            # Place a graphic on the screen.  This is NOT placed in the Sizer, but is placed on top of the notebook tabs.
            self.sslDBImage.SetBitmap(bitmap)

            # ... check to see if both the Database and the Message Server are using SSL
            #     and select the appropriate graphic file ...
            if sslValue:
                bitmap = TransanaImages.locked.GetBitmap()
            else:
                bitmap = TransanaImages.unlocked.GetBitmap()
            # Place a graphic on the screen.  This is NOT placed in the Sizer, but is placed on top of the notebook tabs.
            self.sslImage.SetBitmap(bitmap)
            # Remember the current SSL status of the Message Server
            self.sslStatus = sslValue

    def OnSSLClick(self, event):
        """ Handle click on the SSL indicator image """
        # Determine whether SSL is in use with the Database connection
        dbIsSSL = TransanaGlobal.configData.ssl
        # Determine whether SSL is FULLY in use with the Message Server connection
        chatIsSSL = self.sslStatus
            
        # Start building user feedback based on SSL usage
        if dbIsSSL:
            prompt = _("You have a secure connection to the Database.  ")
        else:
            prompt = _("You do not have a secure connection to the Database.  ")
        if chatIsSSL:
            prompt += '\n' + _("You have a secure connection to the Message Server.  ")
        else:
            prompt += '\n' + _("You do not have a secure connection to the Message Server.  ")
        prompt += "\n\n"
        # Complete user feedback with a summary based on SSL usage
        if dbIsSSL:
            if chatIsSSL:
                prompt += _("Therefore, your Transana connection is as secure as we can make it.")
            else:
                prompt += _('To maintain data security, you should avoid using identifying\ninformation in object names, keywords, and chat messages.')
        else:
            prompt += _("Therefore, your data could be observed during transmission.\nYou may want to look into making your Transana connections more secure.")

        # Create and display a dialog to provide the user security feedback.
        tmpDlg = Dialogs.InfoDialog(self, prompt)
        tmpDlg.ShowModal()
        tmpDlg.Destroy()
        
    def OnKeyDown(self, event):
        """ Handle Key Press events """
        # See if the ControlObject wants to handle the key that was pressed.
        if self.ControlObject.ProcessCommonKeyCommands(event):
            # If so, we're done here.  (Actually, we're done anyway, but what the hell.)
            return

    def GetDimensions(self):
        (left, top) = self.GetPositionTuple()
        (width, height) = self.GetSizeTuple()
        return (left, top, width, height)

    def SetDims(self, left, top, width, height):
        self.SetDimensions(left, top, width, height)

    def ChangeLanguages(self):
        self.SetTitle(_("Data"))
        self.nb.SetPageText(0, _("Database"))
        self.DBTab.tree.refresh_tree()
        self.DBTab.tree.create_menus()

        # print "DataWindow.ChangeLanguages()  (%s)" % _("Data")

    def GetNewRect(self):
        """ Get (X, Y, W, H) for initial positioning """
        pos = self.__pos()
        size = self.__size()
        return (pos[0], pos[1], size[0], size[1])

    def __size(self):
        """Determine default size of Data Frame."""
        # Determine which monitor to use and get its size and position
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            primaryScreen = TransanaGlobal.configData.primaryScreen
        else:
            primaryScreen = 0
        rect = wx.Display(primaryScreen).GetClientArea()  # wx.ClientDisplayRect()
        if not 'wx.GTK' in wx.PlatformInfo:
            container = rect[2:4]
        else:
            screenDims = wx.Display(primaryScreen).GetClientArea()
            # screenDims2 = wx.Display(primaryScreen).GetGeometry()
            left = screenDims[0]
            top = screenDims[1]
            width = screenDims[2] - screenDims[0]  # min(screenDims[2], 1280 - self.left)
            height = screenDims[3]
            container = (width, height)
        
        width = container[0] * .282  # rect[2] * .28
        height = (container[1] - TransanaGlobal.menuHeight) * 0.931  # .66

        # Compensate in Linux.  I'm not sure why this is necessary, but it seems to be.
#        if 'wxGTK' in wx.PlatformInfo:
#            height -= 50

        return wx.Size(width, height)

    def __pos(self):
        """Determine default position of Data Frame."""
        # Determine which monitor to use and get its size and position
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            primaryScreen = TransanaGlobal.configData.primaryScreen
        else:
            primaryScreen = 0
        rect = wx.Display(primaryScreen).GetClientArea()  # wx.ClientDisplayRect()
        if not 'wxGTK' in wx.PlatformInfo:
            container = rect[2:4]
        else:
            # Linux rect includes both screens, so we need to use an alternate method!
            container = TransanaGlobal.menuWindow.GetSize()
        (width, height) = self.__size()
        # rect[0] compensates if the Start menu is on the Left
        if 'wxGTK' in wx.PlatformInfo:
            x = rect[0] + min((rect[2] - 10), (1280 - rect[0])) - width   # min(rect[2], 1440) - width - 3
        else:
            x = rect[0] + container[0] - width - 2  # rect[0] + rect[2] - width - 3
        # rect[1] compensates if the Start menu is on the Top
        if 'wxGTK' in wx.PlatformInfo:
            y = rect[1] + rect[3] - height + 24
        else:
            y = rect[1] + container[1] - height # rect[1] + rect[3] - height - 3

        if DEBUG:
            print "DataWindow.__pos(): Y = %d, H = %d, Total = %d" % (y, height, y + height)
        
        return wx.Point(x, y)
