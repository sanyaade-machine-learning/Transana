# Copyright (C) 2003 - 2014 The Board of Regents of the University of Wisconsin System 
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
# Import the Transana Clip Object
import Clip
# Import the Transana Collection Object
import Collection
# import Transana's Dialogs
import Dialogs
# Import the DBInterface
import DBInterface
# Import Menu Constants
import MenuSetup
# Import Transana's Constants
import TransanaConstants
# import Transana's Exceptions
import TransanaExceptions
# Import Transana's Global variables
import TransanaGlobal

# Increased from 500 to 1000 for multi-transcript clips for Transana 2.30.
TIMER_INTERVAL = 1000
EXTRA_LOAD_TIME = 1500

# Declare GUI Constants for the Play All Clips Dialog
ID_PLAYALLCLIPSTIMER = wx.NewId()
ID_BTNPREVIOUS       = wx.NewId()
ID_BTNPLAYPAUSE      = wx.NewId()
ID_BTNCANCEL         = wx.NewId()
ID_BTNNEXT           = wx.NewId()


class PlayAllClips(wx.Dialog):  # (wx.MDIChildFrame)
    """This object is responsible for controlling media playback when
    the "Play all Clips in a Collection" feature is selected.  It includes
    a user interface dialog and determines what clips are played in what
    order."""
    
    def __init__(self, collection=None, controlObject=None, searchColl=None, treeCtrl=None, singleObject=None):
        """Initialize an PlayAllClips object."""
        # The controlObject parameter, pointing to the Transana ControlObject, is required.
        # If this routine is requested from a Collection,
        #   the Collection should be passed in using the collection parameter.
        # If this routine is requested from a Search Collection,
        #   the searchCollection Node should be passed in the searchColl parameter and
        #   the entire dbTree should be passed in using the treeCtrl parameter.
        # If this routine is requested for a single object, as in Video Only Presentation Mode,
        #   the singleObject should be the Episode or Clip object to be displayed.
        #   (A collection CANNOT be passed in with a singleObject.)

        # Show Clips in Nested Collection.  (If we add a Filter Dialog, this will become optional, but
        # for now, it's just True.
        self.showNested = True
        
        # Remember Control Object information sent in by the calling routine
        self.ControlObject = controlObject
        # Remember the singleObject setting too
        self.singleObject = singleObject

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
            # If we're showing Nested Collections ...
            if self.showNested:
                # Get the initial list of nested Collections
                nestedCollections = DBInterface.list_of_collections(collection.number)
                # As long as there are entries in the list of nested collections that haven't been processed ...
                while len(nestedCollections) > 0:
                    # ... extract the data from the top of the nested collection list ...
                    (collNum, collName, parentCollNum) = nestedCollections[0]
                    # ... and remove that entry from the list.
                    del(nestedCollections[0])
                    # now get the clips for the new collection and add them to the Major (clip) list
                    self.clipList += DBInterface.list_of_clips_by_collection(collName, parentCollNum)
                    # Then get the nested collections under the new collection and add them to the Nested Collection list
                    nestedCollections += DBInterface.list_of_collections(collNum)

        # If PlayAllClips is requested for a Search Collection ...
        elif searchColl != None:
            # Get the Collection Node's Data
            itemData = treeCtrl.GetPyData(searchColl)
            # If we have a Search Collection, it has a recNum.  A Search Result node has recNum == 0.  See which we have.
            if itemData.recNum != 0:
                # Load the Real Collection (not the search result collection) to get its data
                self.collection = Collection.Collection(itemData.recNum)
            # If it's a search collection ...
            else:
                # ... we don't have a collection yet.
                self.collection = None
            # Initialize the Clip List to an empty list
            self.clipList = []
            # Create an empty list for Nested Collections so we can recurse through them
            nestedCollections = []
            # extracting data from the treeCtrl requires a "cookie" value, which is initialized to 0
            cookie = 0
            # Get the first Child node from the searchColl collection
            (item, cookie) = treeCtrl.GetFirstChild(searchColl)
            # While looking at children, we need a pointer to the parent node
            currentNode = searchColl
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
                # If we have a Collection Node ...
                elif self.showNested and (itemData.nodetype == 'SearchCollectionNode'):
                    # ... add it to the list of nested Collections to be processed
                    nestedCollections.append(item)
                    
                # When we get to the last Child Item for the current node, ...
                if item == treeCtrl.GetLastChild(currentNode):
                    # ... check to see if there are nested collections that need to be processed.  If so ...
                    if len(nestedCollections) > 0:
                        # ... set the current node pointer to the first nested collection ...
                        currentNode = nestedCollections[0]
                        # ... get the first child node of the nested collection ...
                        (item, cookie) = treeCtrl.GetFirstChild(nestedCollections[0])
                        # ... and remove the nested collection from the list waiting to be processed
                        del(nestedCollections[0])
                    # If there are no nested collections to be processed ...
                    else:
                        # ... stop looping.  We're done.
                        break
                # If we're not at the Last Child Item ...
                else:
                    # ... get the next Child Item and continue the loop
                    (item, cookie) = treeCtrl.GetNextChild(currentNode, cookie)

        # If a Single Object (Episode or Clip) is passed in ...
        elif singleObject != None:
            # Initialize the Clip List to an empty list
            self.clipList = []
            # Create an empty list for Nested Collections
            nestedCollections = []
            # Remember the collection (which should be None) to avoid processing problems.
            self.collection = collection

        # Register with the ControlObject.  (This allows coordination between the various components that make up Transana.)
        self.ControlObject.Register(PlayAllClips = self)

        # We need to make the Dialog Window a different size depending on whether we are in Presentation Mode or not.
        if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            # Determine the size of the Data Window.
            (left, top, width, height) = self.ControlObject.DataWindow.GetRect()
            desiredHeight = 122
        else:
            # Determine the size of the Client Window
            (left, top, width, height) = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()  # wx.ClientDisplayRect()
            desiredHeight = 56
            
        # Determine (and remember) the appropriate X and Y coordinates for the window
        self.xPos = left
        self.yPos = top + height - desiredHeight
        
        # This Dialog should cover the bottom edge of the Data Window, and should stay on top of it.
        wx.Dialog.__init__(self, self.ControlObject.MenuWindow, -1, _("Play All Clips"),
#        wx.MDIChildFrame.__init__(self, self.ControlObject.MenuWindow, -1, _("Play All Clips"),
                             pos = (self.xPos, self.yPos), size=(width, desiredHeight), style=wx.CAPTION | wx.STAY_ON_TOP)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Add the Play All Clips window to the Window Menu
        self.ControlObject.MenuWindow.AddWindowMenuItem(_("Play All Clips"), 0)

        self.SetMinSize((10, 10))

        # Layout should be different if we're in "Video Only" or "Video and Transcript Only" presentation Modes
        # vs. Standard Transana Mode

        # Add a label that says "Now Playing:"
        lblNowPlaying = wx.StaticText(self, -1, _("Now Playing:"))
        # If we have a defined collection (ie NOT a Search Result node) ...
        if self.collection != None:
            # Add a label that identifies the Collection
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Collection: %s"), 'utf8')
            else:
                prompt = _("Collection: %s")
            self.lblCollection = wx.StaticText(self, -1, prompt % self.collection.GetNodeString())
        # If we have a single object ...
        elif self.singleObject != None:
            # Determine the type of the passed-in object and set the prompt accordingly.
            if type(self.singleObject) == Clip.Clip:
                prompt = _("Clip: %s")
            else:
                prompt = _("Episode: %s")
            # Convert the prompt to Unicode
            if 'unicode' in wx.PlatformInfo:
                prompt = unicode(prompt, 'utf8')
            # Display the prompt
            self.lblCollection = wx.StaticText(self, -1, prompt % self.singleObject.id)
        # If we have a Search Result Node ...
        else:
            # ... create a blank label, which will get filled in in just a moment with the first clip's Collection name
            self.lblCollection = wx.StaticText(self, -1, "")

        # If we do NOT have a single object ...
        if self.singleObject == None:
            # Add a label that identifies the Clip
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Clip: %s   (%s of %s)"), 'utf8')
            else:
                prompt = _("Clip: %s   (%s of %s)")
            self.lblClip = wx.StaticText(self, -1, prompt % (' ', 0, len(self.clipList)))

            # Add a button for Previous
            self.btnPrevious = wx.Button(self, ID_BTNPREVIOUS, _("Previous"))
            self.btnPrevious.SetMinSize((10, self.btnPrevious.GetSize()[1]))
            self.btnPrevious.Bind(wx.EVT_BUTTON, self.OnChangeClip)
        # If we DO have a single object ...
        else:
            # ... then the Clip Label should be blank.  (The other label handles single Episodes and Clips)
            self.lblClip = wx.StaticText(self, -1, "")

        # Add a button for Pause/Play functioning
        self.btnPlayPause = wx.Button(self, ID_BTNPLAYPAUSE, _("Pause"))
        self.btnPlayPause.SetMinSize((10, self.btnPlayPause.GetSize()[1]))
        wx.EVT_BUTTON(self, ID_BTNPLAYPAUSE, self.OnPlayPause)

        # Add a button to Cancel the Playing of Clips
        self.btnCancel = wx.Button(self, ID_BTNCANCEL, _("Cancel"))
        self.btnCancel.SetMinSize((10, self.btnCancel.GetSize()[1]))
        wx.EVT_BUTTON(self, ID_BTNCANCEL, self.OnClose)

        # If we do NOT have a single object ...
        if self.singleObject == None:
            # Add a button for Next
            self.btnNext = wx.Button(self, ID_BTNNEXT, _("Next"))
            self.btnNext.SetMinSize((10, self.btnNext.GetSize()[1]))
            self.btnNext.Bind(wx.EVT_BUTTON, self.OnChangeClip)

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
            btnSizer.Add(self.btnPrevious, 1, wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT | wx.EXPAND, 5)
            btnSizer.Add(self.btnPlayPause, 1, wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT | wx.EXPAND, 5)
            btnSizer.Add(self.btnCancel, 1, wx.ALIGN_RIGHT | wx.LEFT | wx.RIGHT | wx.EXPAND, 5)
            btnSizer.Add(self.btnNext, 1, wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT | wx.EXPAND, 5)
            box.Add(btnSizer, 1, wx.EXPAND, 0)
            box.Add((1, 5))
        else:
            box = wx.BoxSizer(wx.HORIZONTAL)
            if self.singleObject == None:
                box.Add(self.btnPrevious, 1, wx.ALIGN_LEFT | wx.LEFT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)
            box.Add(self.btnPlayPause, 1, wx.ALIGN_LEFT | wx.LEFT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)
            box.Add(self.btnCancel, 1, wx.ALIGN_LEFT | wx.LEFT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)
            if self.singleObject == None:
                box.Add(self.btnNext, 1, wx.ALIGN_LEFT | wx.LEFT | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)
            box.Add((10, 1), 0, wx.EXPAND)
            box.Add(lblNowPlaying, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
            box.Add((10,1), 1, wx.EXPAND)
            box.Add(self.lblCollection, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
            box.Add((10, 1), 1, wx.EXPAND)
            box.Add(self.lblClip, 0, wx.ALIGN_LEFT | wx.ALIGN_CENTER_VERTICAL)
            box.Add((10, 1), 1, wx.EXPAND)

        self.SetSizer(box)
#        self.Fit()

        # Tell the Dialog to Lay out the widgets, and to adjust them automatically
        self.Layout()
        self.SetAutoLayout(True)

        # Now that the size is determined, let's reposition the dialog.
        if not self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            (left, top, width, height) = self.ControlObject.VideoWindow.GetRect()
        self.xPos = left
        self.yPos = top + height - self.GetSize()[1]
        self.SetPosition((self.xPos, self.yPos))

        # Point to the first clip in the list as the clip that should be played
        self.clipNowPlaying = 0

        # Add a Timer.  The timer checks to see if the clip that is playing has stopped, which
        # is the signal that it is time to load the next clip
        self.playAllClipsTimer = wx.Timer(self, ID_PLAYALLCLIPSTIMER)
        wx.EVT_TIMER(self, ID_PLAYALLCLIPSTIMER, self.OnTimer)

        # If there are clips to play, play them
        if len(self.clipList) > 0:
            # Initialize the flag that says a clip has started playing to TRUE or the first video will never load!
            self.HasStartedPlaying = True
            # 1.5 second additional delay hopefully gives a clip enough time to load and play.
            # If clips are getting skipped, this increment may need to be increased.
            self.playAllClipsTimer.Start(TIMER_INTERVAL + EXTRA_LOAD_TIME)
            # Show the Play All Clips Dialog
            self.Show()
        # If we have a single object ...
        elif singleObject != None:
            # Initialize the flag that says a clip has started playing to FALSE or the first video will not play!
            self.HasStartedPlaying = False
            # Start the timer, which will detect when the episode / clip is done.  Since the video is already loaded,
            # there's no need for additional time.
            self.playAllClipsTimer.Start(TIMER_INTERVAL)
            # Show the Play All Clips Dialog
            self.Show()
        else:
            # Hide the PlayAllClips Window
            self.Hide()
            # If there are no clips to play, display an error message.
            dlg = Dialogs.InfoDialog(None, _("This Collection has no Clips to play."))
            dlg.ShowModal()
            dlg.Destroy()
            self.ControlObject.Register(PlayAllClips = self)
            # We need to close the PlayAllClips Windows, but can't until this method is done processing.
            wx.CallAfter(self.Close)

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

            # Load the Collection the next Clip is from
            tempColl = Collection.Collection(self.clipList[self.clipNowPlaying][2])
            # Add a label that identifies the Collection
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Collection: %s"), 'utf8')
            else:
                prompt = _("Collection: %s")
            self.lblCollection.SetLabel(prompt % tempColl.GetNodeString())
            # If we're in PRESENT_ALL presentation mode ...
            if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
                # Let's wrap the Collection Name to fit inside the window.  (Needed with the GetNodeString() change.)
                self.lblCollection.Wrap(self.GetSizeTuple()[0] - 20)
            # If we're in PRESENT_ALL presentation mode ...
            if self.ControlObject.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
                # Now we need to reposition the window because of the changed size
                (left, top, width, height) = self.ControlObject.DataWindow.GetRect()
                self.xPos = left
                self.yPos = top + height - self.GetSize()[1]
                self.SetPosition((self.xPos, self.yPos))
                self.Raise()

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
            # If the video volume doesn't exist on a Mac, or the video cannot be found ...
            except wx._core.PyAssertionError:
                pass

        else:

            if DEBUG:
                print "Clip load has been called, but Play hasn't quite started yet!"

            self.PlayAfterLoading()

    def PlayAfterLoading(self):
        """ After a Clip is done loading, it needs to be told to Play. """
        # If we don't have a control object or a currently loaded object ...
        if (self.ControlObject == None) or (self.ControlObject.currentObj == None):
            # ... then skip this!
            return
        
        # I can't figure out why Vista with WMV files has lost the Clip Start position by this point during Play All Clips.
        # The correct start position gets set, but then something clears it.
        # Fortunately, this corrects it.

        # If the current video position is earlier than the clip start ...
        if self.ControlObject.VideoWindow.GetCurrentVideoPosition() < self.ControlObject.currentObj.clip_start:
            # ... then set the video start to the clip start.
            self.ControlObject.SetVideoStartPoint(self.ControlObject.currentObj.clip_start)

        try:
            # Play the next clip
            self.ControlObject.Play()
        except IndexError:
            self.Close()

        # Clip loaded.  Restart the timer to the shorter interval.
        wx.CallAfter(self.playAllClipsTimer.Start, TIMER_INTERVAL)

        if DEBUG:
            print "Clip", self.clipNowPlaying, "has been told to play"

    def OnChangeClip(self, event):
        """ Implement the "Previous" and "Next" buttons """
        # If Previous is pressed ...
        if event.GetId() == ID_BTNPREVIOUS:
            # See if we're later than the first clip.
            if self.clipNowPlaying > 1:
                # If so, go back TWO, as the number gets incremented forward by the Timer
                self.clipNowPlaying -= 2
            # If this is the first clip ...
            elif self.clipNowPlaying > 0:
                # ... go back only one (because of the Timer's increment), as there's no previous clip to go back to.
                self.clipNowPlaying -= 1
            # Make sure the Windows WON'T re-arrange if moving backwards
            self.ControlObject.shutdownPlayAllClips = False
        # If the Next button is pressed ...
        elif event.GetId() == ID_BTNNEXT:
            # ... we don't need to do anything here.  The Timer's increment will take care of it for us.
            pass
        # If the clip is actively playing ...
        if self.ControlObject.IsPlaying() and (self.btnPlayPause.GetLabel() == unicode(_("Pause"), 'utf8')):
            # Stop the current clip.  This will cause the Timer to move us on to the proper clip, previous or next.
            self.ControlObject.Stop()
        # If the clip is actually paused (ie is not in mid-load) ...
        elif self.ControlObject.IsPaused():
            # Change the label on the button
            self.btnPlayPause.SetLabel(_("Pause"))
            # Restart the timer so that the changed clip will load.
            wx.CallAfter(self.playAllClipsTimer.Start, TIMER_INTERVAL)

    def OnPlayPause(self, event):
        """ If playing, then pause.  If paused, then play. """
        if self.btnPlayPause.GetLabel() == unicode(_("Pause"), 'utf8'):
            # Prevent screen re-organization if the LAST clip is paused
            self.ControlObject.shutdownPlayAllClips = False
            # Stop the time when we pause.  This is necessary to prevent clips from sometimes being dropped when we re-start.
            self.playAllClipsTimer.Stop()
            # Pause the video
            self.ControlObject.Pause()
            # Change the label on the button
            self.btnPlayPause.SetLabel(_("Play"))
        else:
            # If we're on the LAST Clip ...
            if (self.clipNowPlaying == len(self.clipList)):
                # Prevent screen re-organization if the LAST clip is paused
                self.ControlObject.shutdownPlayAllClips = True
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
        # Signal the Control Object that we need to reset the Window Configuration
        self.ControlObject.shutdownPlayAllClips = True
        # If a video is playing, stop it!
        if self.ControlObject.IsPlaying() or self.ControlObject.IsPaused():
            self.ControlObject.Stop()
        # Sometimes, if paused, we don't see the windows restore correctly on CANCEL on Mac.
        # This fixes that.
        else:
            # Send a STOP signal, even if not playing!!
            self.ControlObject.UpdatePlayState(TransanaConstants.MEDIA_PLAYSTATE_STOP)
        # If the timer is still active, stop it!
        if self.playAllClipsTimer.IsRunning():
            self.playAllClipsTimer.Stop()
        # Delete the Play All Clips window from the Window Menu
        self.ControlObject.MenuWindow.DeleteWindowMenuItem(_("Play All Clips"), 0)
        # Return the Data Window to the Database Tab
        self.ControlObject.ShowDataTab(0)
        # Un-Register with the ControlObject
        self.ControlObject.Register(PlayAllClips = None)
        # Close the dialog box
        self.Close()
