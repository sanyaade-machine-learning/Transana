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
#

""" This module implements the PlayAllClips class, which is a controller window that handles
    the "Play All Clips" in a Collection function. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "PlayAllClips DEBUG is ON!!"

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

TIMER_INTERVAL = 500
EXTRA_LOAD_TIME = 1500

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
        lblNowPlaying = wx.StaticText(self, -1, _("Now Playing:"))

        # Add a label that identifies the Collection
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("Collection: %s"), 'utf8')
        else:
            prompt = _("Collection: %s")
        self.lblCollection = wx.StaticText(self, -1, prompt % self.collection.id)

        # Add a label that identifies the Clip
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_("Clip: %s   (%s of %s)"), 'utf8')
        else:
            prompt = _("Clip: %s   (%s of %s)")
        self.lblClip = wx.StaticText(self, -1, prompt % (' ', 0, len(self.clipList)))

        # Add a button for Pause/Play functioning
        self.btnPlayPause = wx.Button(self, ID_BTNPLAYPAUSE, _("Pause"))

        wx.EVT_BUTTON(self, ID_BTNPLAYPAUSE, self.OnPlayPause)

        # Add a button to Cancel the Playing of Clips
        self.btnCancel = wx.Button(self, ID_BTNCANCEL, _("Cancel"))

        wx.EVT_BUTTON(self, ID_BTNCANCEL, self.OnClose)

        # Link to a method that handles window move attempts
        wx.EVT_MOVE(self, self.OnMove)

        # Define the layout differently depending on the Presentation Mode setting.
        if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            box = wx.BoxSizer(wx.VERTICAL)
            box.Add((1, 5))
            box.Add(lblNowPlaying, 0, wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT, 10)
            box.Add((1, 5))
            box.Add(self.lblCollection, 0, wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT, 10)
            box.Add((1, 5))
            box.Add(self.lblClip, 0, wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT, 10)
            box.Add((1, 5))
            btnSizer = wx.BoxSizer(wx.HORIZONTAL)
            btnSizer.Add(self.btnPlayPause, 1, wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT | wx.EXPAND, 10)
            btnSizer.Add(self.btnCancel, 1, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.EXPAND, 10)
            box.Add(btnSizer, 1, wx.EXPAND, 0)
            box.Add((1, 5))
        else:
            box = wx.BoxSizer(wx.HORIZONTAL)
            box.Add(self.btnPlayPause, 1, wx.ALIGN_LEFT | wx.LEFT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 10)
            box.Add(self.btnCancel, 1, wx.ALIGN_LEFT | wx.LEFT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 10)
            box.Add((10, 1), 1, wx.EXPAND)
            box.Add(lblNowPlaying, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
            box.Add((10,1), 1, wx.EXPAND)
            box.Add(self.lblCollection, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
            box.Add((10, 1), 1, wx.EXPAND)
            box.Add(self.lblClip, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
            box.Add((10, 1), 1, wx.EXPAND)

        self.SetSizer(box)
        self.Fit()


        # Tell the Dialog to Lay out the widgets, and to adjust them automatically
        self.Layout()
        self.SetAutoLayout(True)

        # Now that the size is determined, let's reposition the dialog.
        if not self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            (left, top, width, height) = self.ControlObject.VideoWindow.frame.GetRect()        
        self.xPos = left
        self.yPos = top + height - self.GetSize()[1]
        self.SetPosition((self.xPos, self.yPos))


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
            # 1.5 second additional delay hopefully gives a clip enough time to load and play.
            # If clips are getting skipped, this increment may need to be increased.
            self.playAllClipsTimer.Start(TIMER_INTERVAL + EXTRA_LOAD_TIME)
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

        if DEBUG:
            print self.clipNowPlaying,
            if self.clipNowPlaying < len(self.clipList):
                print self.clipList[self.clipNowPlaying]
            else:
                print
            print "IsLoading()", (self.ControlObject.IsLoading())
            print "IsPlaying()", (self.ControlObject.IsPlaying())
            print "IsPaused()", self.ControlObject.IsPaused()
            print "Getlabel() != Play", self.btnPlayPause.GetLabel() != _("Play")
            print "HasStartedPlaying:", (self.HasStartedPlaying)
            print

        # The original code doesn't work with the new video player infrastructure.  This is an attempt to start over.

        # If the clip is playing, we don't do anything!
        if self.ControlObject.IsPlaying():

            if DEBUG:
                print "Clip is playing."
            
            self.HasStartedPlaying = True

        # If a Clip is in the process of loading, we don't need to do anything but wait for it to finish loading
        elif self.ControlObject.IsLoading():

            if DEBUG:
                print "Clip is loading.  Don't do anything!", self.clipNowPlaying

        elif (self.clipNowPlaying > len(self.clipList) - 1) and self.HasStartedPlaying:

            if DEBUG:
                print "All clips have played.  (%s > %s).  We need to close." % (self.clipNowPlaying, len(self.clipList) - 1)

            self.OnClose(event)
            
        # If a clip has been loaded and has started playing, but isn't playing any more, then load another!
        elif self.HasStartedPlaying:

            if DEBUG:
                print "Clip isn't playing, paused, or loading.  We need to load the next clip!", self.clipNowPlaying, self.clipList[self.clipNowPlaying]

            # First, update the label to tell what clip is up
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Clip: %s   (%s of %s)"), 'utf8')
            else:
                prompt = _("Clip: %s   (%s of %s)")
            self.lblClip.SetLabel(prompt % (self.clipList[self.clipNowPlaying][1], self.clipNowPlaying + 1, len(self.clipList)))

            # In MU, it's possible a clip in the Play All Clips list could get deleted by another user.
            # Therefore, we need to be prepared to catch the exception that is raised by failing to be
            # able to load the Clip.
            try:
                # Loading a Clip is slow.  Let's stop the timer, so it doesn't cause problems.  (It was with QuickTime video on Windows.)
                self.playAllClipsTimer.Stop()

                # Try to Load the next clip into the ControlObject
                if not self.ControlObject.LoadClipByNumber(self.clipList[self.clipNowPlaying][0]):
                    # If the Media File has been moved, this failed.  Try one more time.
                    if not self.ControlObject.LoadClipByNumber(self.clipList[self.clipNowPlaying][0]):
                        # if it fails a second time, signal that Play All Clips should be stopped
                        # by setting the Clip List Pointer to the end of the list
                        self.clipNowPlaying = len(self.clipList)
                        
                # Increment the pointer to the next clip
                self.clipNowPlaying = self.clipNowPlaying + 1
                self.HasStartedPlaying = False

                # Let's display the Keywords Tab during PlayAllClips
                self.ControlObject.ShowDataTab(1)

                # Now that the clip is loading, re-start the timer.  the extra 1.5 seconds give the clip time to load.
                self.playAllClipsTimer.Start(TIMER_INTERVAL + EXTRA_LOAD_TIME)

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

        else:

            if DEBUG:
                print "Clip load has been called, but Play hasn't quite started yet!"

            self.PlayAfterLoading()

    def PlayAfterLoading(self):
        """ After a Clip is done loading, it needs to be told to Play. """
        # Play the next clip
        self.ControlObject.Play()

        # Clip loaded.  Restart the timer to the shorter interval.
        wx.CallAfter(self.playAllClipsTimer.Start, TIMER_INTERVAL)

        if DEBUG:
            print "Clip", self.clipNowPlaying, "has been told to play"

    def OnPlayPause(self, event):
        """ If playing, then pause.  If paused, then play. """
        if self.btnPlayPause.GetLabel() == _("Pause"):
            # Stop the time when we pause.  This is necessary to prevent clips from sometimes being dropped when we re-start.
            self.playAllClipsTimer.Stop()
            # Pause the video
            self.ControlObject.Pause()
            # Change the label on the button
            self.btnPlayPause.SetLabel(_("Play"))
        else:
            # Play the video
            self.ControlObject.Play()
            # Change the label on the button
            self.btnPlayPause.SetLabel(_("Pause"))
            # Restart the timer.
            wx.CallAfter(self.playAllClipsTimer.Start, TIMER_INTERVAL)

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
