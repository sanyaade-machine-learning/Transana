# Copyright (C) 2003 - 2006 The Board of Regents of the University of Wisconsin System 
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

""" This module implements the PlayAllClips class, which is a controller window that handles
    the "Play All Clips" in a Collection function. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Import wxPython
import wx
# Import the Transana Collection Object
import Collection
# import Transana's Dialogs
import Dialogs
# Import the DBInterface
import DBInterface
# Import Menu Constants
import MenuSetup
# import Transana's Exceptions
import TransanaExceptions

# Declare GUI Constants for the Play All Clips Dialog
ID_PLAYALLCLIPSTIMER = wx.NewId()
ID_BTNPLAYPAUSE      = wx.NewId()
ID_BTNCANCEL         = wx.NewId()


class PlayAllClips(wx.Dialog):
    """This object is responsible for controlling media playback when
    the "Play all Clips in a Collection" feature is selected.  It includes
    a user interface dialog and determines what clips are played in what
    order."""
    
    def __init__(self, collection=None, controlObject=None, searchColl=None, treeCtrl=None):
        """Initialize an PlayAllClips object."""
        # The controlObject parameter, pointing to the Transana ControlObject, is required.
        # If this routine is requested from a Collection,
        #   the Collection should be passed in using the collection parameter.
        # If this routine is requested from a Search Collection,
        #   the searchCollection Node should be passed in the searchColl parameter and
        #   the entire dbTree should be passed in using the treeCtrl parameter.
        
        # Remember Control Object information sent in by the calling routine
        self.ControlObject = controlObject

        # If a video is currently playing, it must be stopped!
        if self.ControlObject.IsPlaying() or self.ControlObject.IsPaused():
            self.ControlObject.Stop()
        
        # Get a list of all the clips that should be played in order

        # If PlayAllClips is requested for a Collection ...
        if collection != None:
            # Remember the Collection, making it available throughout this class
            self.collection = collection
            # Load a list of all Clips in the Collection.
            self.clipList = DBInterface.list_of_clips_by_collection(self.collection.id, self.collection.parent)

        # If PlayAllClips is requested for a Search Collection ...
        elif searchColl != None:
            # Get the Collection Node's Data
            itemData = treeCtrl.GetPyData(searchColl)
            # Load the Real Collection (not the search result collection) to get its data
            self.collection = Collection.Collection(itemData.recNum)
            # Initialize the Clip List to an empty list
            self.clipList = []
            # extracting data from the treeCtrl requires a "cookie" value, which is initialized to 0
            cookie = 0
            # Get the first Child node from the searchColl collection
            (item, cookie) = treeCtrl.GetFirstChild(searchColl)
            # Process all children in the searchColl collection
            while item.IsOk():
                # Get the item's Name
                itemText = treeCtrl.GetItemText(item)
                # Get the item's Node Data
                itemData = treeCtrl.GetPyData(item)
                # See if the item is a Clip
                if itemData.nodetype == 'SearchClipNode':
                    # If it's a Clip, add the Clip's Node Data to the clipList
                    self.clipList.append((itemData.recNum, itemText, itemData.parent))
                # When we get to the last Child Item, stop looping
                if item == treeCtrl.GetLastChild(searchColl):
                    break
                # If we're not at the Last Child Item, get the next Child Item and continue the loop
                else:
                    (item, cookie) = treeCtrl.GetNextChild(searchColl, cookie)


        # Register with the ControlObject.  (This allows coordination between the various components that make up Transana.)
        self.ControlObject.Register(PlayAllClips = self)

        # We need to make the Dialog Window a different size depending on whether we are in Presentation Mode or not.
        if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            # Determine the size of the Data Window.
            (left, top, width, height) = self.ControlObject.DataWindow.GetRect()
            desiredHeight = 122
        else:
            # Determine the size of the Client Window
            (left, top, width, height) = wx.ClientDisplayRect()
            desiredHeight = 56
            
        # Determine (and remember) the appropriate X and Y coordinates for the window
        self.xPos = left
        self.yPos = top + height - desiredHeight
        
        # This Dialog should cover the bottom edge of the Data Window, and should stay on top of it.
        wx.Dialog.__init__(self, self.ControlObject.DataWindow, -1, _("Play All Clips"),
                             pos = (self.xPos, self.yPos), size=(width, desiredHeight), style=wx.CAPTION | wx.STAY_ON_TOP)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Layout should be different if we're in "Video Only" or "Video and Transcript Only" presentation Modes
        # vs. Standard Transana Mode

        # Add a label that says "Now Playing:"
        lay = wx.LayoutConstraints()
        if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            lay.left.SameAs(self, wx.Left, 10)
            lay.top.SameAs(self, wx.Top, 10)
        else:
            lay.left.PercentOf(self, wx.Width, 30)
            lay.top.SameAs(self, wx.Top, 6)
        lay.width.AsIs()
        lay.height.AsIs()
        lblNowPlaying = wx.StaticText(self, -1, _("Now Playing:"))
        lblNowPlaying.SetConstraints(lay)

        # Add a label that identifies the Collection
        lay = wx.LayoutConstraints()
        if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            lay.left.SameAs(self, wx.Left, 10)
            lay.top.Below(lblNowPlaying, 5)
        else:
            lay.left.PercentOf(self, wx.Width, 40)
            lay.top.SameAs(self, wx.Top, 6)
        lay.width.AsIs()
        lay.height.AsIs()
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("Collection: %s"), 'utf8')
        else:
            prompt = _("Collection: %s")
        self.lblCollection = wx.StaticText(self, -1, prompt % self.collection.id)
        self.lblCollection.SetConstraints(lay)

        # Add a label that identifies the Clip
        lay = wx.LayoutConstraints()
        if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            lay.left.SameAs(self, wx.Left, 10)
            lay.top.Below(self.lblCollection, 5)
            lay.right.SameAs(self, wx.Right, 10)
        else:
            lay.left.PercentOf(self, wx.Width, 60)
            lay.top.SameAs(self, wx.Top, 6)
            lay.width.AsIs()
        lay.height.AsIs()
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("Clip: %s   (%s of %s)"), 'utf8')
        else:
            prompt = _("Clip: %s   (%s of %s)")
        self.lblClip = wx.StaticText(self, -1, prompt % (' ', 0, len(self.clipList)))
        self.lblClip.SetConstraints(lay)

        # Add a button for Pause/Play functioning
        lay = wx.LayoutConstraints()
        lay.left.SameAs(self, wx.Left, 10)
        if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            lay.top.Below(self.lblClip, 5)
            lay.right.PercentOf(self, wx.Width, 48)
            lay.height.AsIs()
        else:
            lay.top.SameAs(self, wx.Top, 1)
            lay.width.PercentOf(self, wx.Width, 12)
            # Layout was funky on the Mac.
            if "__WXMAC__" in wx.PlatformInfo:
                lay.height.AsIs()
            else:
                lay.bottom.SameAs(self, wx.Bottom, 1)
        self.btnPlayPause = wx.Button(self, ID_BTNPLAYPAUSE, _("Pause"))
        self.btnPlayPause.SetConstraints(lay)

        wx.EVT_BUTTON(self, ID_BTNPLAYPAUSE, self.OnPlayPause)

        # Add a button to Cancel the Playing of Clips
        lay = wx.LayoutConstraints()
        if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            lay.left.PercentOf(self, wx.Width, 52)
            lay.top.Below(self.lblClip, 5)
            lay.right.SameAs(self, wx.Right, 10)
            lay.height.AsIs()
        else:
            lay.left.RightOf(self.btnPlayPause, 10)
            lay.top.SameAs(self, wx.Top, 1)
            lay.width.SameAs(self.btnPlayPause, wx.Width)
            # Layout was funky on the Mac.
            if "__WXMAC__" in wx.PlatformInfo:
                lay.height.AsIs()
            else:
                lay.bottom.SameAs(self, wx.Bottom, 1)
        self.btnCancel = wx.Button(self, ID_BTNCANCEL, _("Cancel"))
        self.btnCancel.SetConstraints(lay)

        wx.EVT_BUTTON(self, ID_BTNCANCEL, self.OnClose)

        # Link to a method that handles window move attempts
        wx.EVT_MOVE(self, self.OnMove)

        # Tell the Dialog to Lay out the widgets, and to adjust them automatically
        self.Layout()
        self.SetAutoLayout(True)

        # Add a Timer.  The timer checks to see if the clip that is playing has stopped, which
        # is the signal that it is time to load the next clip
        self.playAllClipsTimer = wx.Timer(self, ID_PLAYALLCLIPSTIMER)
        wx.EVT_TIMER(self, ID_PLAYALLCLIPSTIMER, self.OnTimer)

        # Point to the first clip in the list as the clip that should be played
        self.clipNowPlaying = 0

        # If there are clips to play, play them
        if len(self.clipList) > 0:
            # Initialize the flag that says a clip has started playing to TRUE or the first video will never load!
            self.HasStartedPlaying = True
            # 0.5 second delay hopefully gives a clip enough time to load and play.
            # If clips are getting skipped, this increment may need to be increased.
            self.playAllClipsTimer.Start(500)

            # Show the Play All Clips Dialog
            self.ShowModal()
        else:
            # If there are no clips to play, display an error message.
            dlg = wx.MessageDialog(None, _("This Collection has no Clips to play."), style = wx.OK | wx.ICON_EXCLAMATION)
            dlg.ShowModal()
            dlg.Destroy()
            self.ControlObject.Register(PlayAllClips = self)
            


    def OnTimer(self, event):
        """ This method should be polled periodically when this dialog is displayed.  If the media player is paused or playing,
            everything is fine and nothing needs to be done.  If not, we need to load the next clip and play it. """

        # Occasionally, clips were getting skipped.  It seems that there can be up to a 0.1 second delay
        # between when a video is loaded and when it starts playing (at least on Windows) because of the
        # way video_msw.py works.  It has a timer that checks to see if the video is done loading before
        # it starts playing.  So I'm introducing this HasStartedPlaying variable to try to make sure no
        # clips get skipped

        # Check to see if the video is NOT playing and NOT paused ...
        if (not self.ControlObject.IsPlaying()) and (not self.ControlObject.IsPaused()):
            # If it is neither playing nor paused, see if we have reached the end of the list 
            if self.clipNowPlaying >= len(self.clipList):
                # If so, close the PlayAllClips window.
                self.OnClose(event)
            elif (not self.ControlObject.IsLoading()) and (self.HasStartedPlaying):
                # If we are neither paused nor playing, nor are we out of clips, then
                # we need to load the next clip!!
                # First, update the label to tell what clip is up
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("Clip: %s   (%s of %s)"), 'utf8')
                else:
                    prompt = _("Clip: %s   (%s of %s)")
                self.lblClip.SetLabel(prompt % (self.clipList[self.clipNowPlaying][1], self.clipNowPlaying + 1, len(self.clipList)))

                # Stop the timer long enough to load the Clip.  That way, if the MediaFile is bad, we don't
                # get repeated attempts to load the clip.
                self.playAllClipsTimer.Stop()

                # In MU, it's possible a clip in the Play All Clips list could get deleted by another user.
                # Therefore, we need to be prepared to catch the exception that is raised by failing to be
                # able to load the Clip.
                try:
                    # Try to Load the next clip into the ControlObject
                    if not self.ControlObject.LoadClipByNumber(self.clipList[self.clipNowPlaying][0]):
                        # If the Media File has been moved, this failed.  Try one more time.
                        if not self.ControlObject.LoadClipByNumber(self.clipList[self.clipNowPlaying][0]):
                            # if it fails a second time, signal that Play All Clips should be stopped
                            # by setting the Clip List Pointer to the end of the list
                            self.clipNowPlaying = len(self.clipList)
                    # Clip loaded.  Restart the timer.
                    self.playAllClipsTimer.Start(500)

                    # Play the next clip
                    self.ControlObject.Play()
                    # Increment the pointer to the next clip
                    self.clipNowPlaying = self.clipNowPlaying + 1
                    
                    self.HasStartedPlaying = False
                # If a Clip cannot be found ...  (This should only happen in MU if a clip is deleted by another user.)
                except TransanaExceptions.RecordNotFoundError:
                    # Build an error message
                    msg = 'Clip "%s" could not be found.\nPerhaps it was deleted by another user.\nPlay All Clips cannot continue.'
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8')
                    # Display the error message
                    dlg = Dialogs.ErrorDialog(self, msg % self.clipList[self.clipNowPlaying][1])
                    dlg.ShowModal()
                    dlg.Destroy()
                    # Get out of Play All Clips.
                    self.OnClose(event)

        # If a video has not yet been flagged as playing but it HAS started playing, flag it as having started.
        elif (not self.HasStartedPlaying) and (self.ControlObject.IsPlaying()):
            # Let's display the Keywords Tab during PlayAllClips
            self.ControlObject.ShowDataTab(1)
            self.HasStartedPlaying = True

    def OnPlayPause(self, event):
        """ If playing, then pause.  If paused, then play. """
        if self.btnPlayPause.GetLabel() == _("Pause"):
            self.ControlObject.Pause()
            # Change the label on the button
            self.btnPlayPause.SetLabel(_("Play"))
        else:
            self.ControlObject.Play()
            # Change the label on the button
            self.btnPlayPause.SetLabel(_("Pause"))

    def OnMove(self, event):
        """ Detect attempts to move the Play All Windows dialog, and block it if appropriate. """
        # If we are not in "Show All Windows" mode, block attempts to move the Play All Clips window
        if not self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            self.Move(wx.Point(self.xPos, self.yPos))

    def OnClose(self, event):
        """ Close the Play All Clips Dialog """
        # If a video is playing, stop it!
        if self.ControlObject.IsPlaying() or self.ControlObject.IsPaused():
            self.ControlObject.Stop()
        # If the timer is still active, stop it!
        if self.playAllClipsTimer.IsRunning():
            self.playAllClipsTimer.Stop()
        # Return the Data Window to the Database Tab
        self.ControlObject.ShowDataTab(0)
        # Un-Register with the ControlObject
        self.ControlObject.Register(PlayAllClips = None)
        # Close the dialog box
        self.Close()
