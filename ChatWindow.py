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

""" A Message Client that processes inter-instance chat communication between Transana-MU
    clients.  This client utility requires connection to a Transana Message Server. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "ChatWindow DEBUG is ON!"


# import wxPython
import wx
# import Python socket module
import socket
# import Python os module
import os
# import Python sys module
import sys
# import Python time module
import time
# import Python threading module
import threading
# import Transana's Global module
import TransanaGlobal
# import Transana Exceptions
import TransanaExceptions
# import Transana's Database Interface
import DBInterface
# import Transana Dialogs
import Dialogs
# import Transana's Series object
import Series
# import Transana's Episode object
import Episode
# import Transana's Transcript object
import Transcript
# import Transana's Collection object
import Collection
# import Transana's Clip object
import Clip
# import Transana's Note object
import Note


# We create a thread to listen for messages from the Message Server.  However,
# only the Main program thread can interact with a wxPython GUI.  Therefore,
# we need to create a custom event to receive the message from the Message Server
# (via the socket connection) and transfer that data to the Main program thread
# where it will be displayed.
#
# We need a second event to signal that the socket connection has died and that
# the Chat Window needs to be closed.

# Get an ID for a custom Event for posting messages received in a thread
EVT_POST_MESSAGE_ID = wx.NewId()
# Get an ID for a custom Event for signalling that the Chat Window needs to be closed
EVT_CLOSE_MESSAGE_ID = wx.NewId()
# Get an ID for a custom Event for if the Message Server is lost
EVT_MESSAGESERVER_LOST_ID = wx.NewId()

# Define a custom "Post Message" event
def EVT_POST_MESSAGE(win, func):
    """ Defines the EVT_POST_MESSAGE event type """
    win.Connect(-1, -1, EVT_POST_MESSAGE_ID, func)

# Define a custom "Close Message" event
def EVT_CLOSE_MESSAGE(win, func):
    """ Defines the EVT_CLOSE_MESSAGE event type """
    win.Connect(-1, -1, EVT_CLOSE_MESSAGE_ID, func)

# Define a custom "Message Server Lost" event
def EVT_MESSAGESERVER_LOST(win, func):
    """ Defines the EVT_CLOSE_MESSAGE event type """
    win.Connect(-1, -1, EVT_MESSAGESERVER_LOST_ID, func)

# Create the actual Custom Post Message Event object
class PostMessageEvent(wx.PyEvent):
    """ This event is used to trigger posting a message in the GUI. It carries the data. """
    def __init__(self, data):
        # Initialize a wxPyEvent
        wx.PyEvent.__init__(self)
        # Link the event to the Event ID
        self.SetEventType(EVT_POST_MESSAGE_ID)
        # Store the message data in the Custom Event object
        self.data = data

        if DEBUG:
            print "PostMessageEvent created with data =", self.data.encode('latin1')

# Create the actual Custom Close Message Event object
class CloseMessageEvent(wx.PyEvent):
    """ This event is used to trigger closing the GUI. """
    def __init__(self):
        # Initialize a wxPyEvent
        wx.PyEvent.__init__(self)
        # Link the event to the Event ID
        self.SetEventType(EVT_CLOSE_MESSAGE_ID)

# Create the actual Custom Message Server Lost Event object
class MessageServerLostEvent(wx.PyEvent):
    """ This event is used to indicate that the Message Server has been lost to the GUI. """
    def __init__(self):
        # Initialize a wxPyEvent
        wx.PyEvent.__init__(self)
        # Link the event to the Event ID
        self.SetEventType(EVT_MESSAGESERVER_LOST_ID)

        if DEBUG:
            print "MessageServerLost Event created"

# Create a Thread Lock object so that we can use thread locking as needed
threadLock = threading.Lock()

# Create a custom Thread object for listening for messages from the Message Server
class ListenerThread(threading.Thread):
    """ custom Thread object for listening for Messages from the Message Server """
    def __init__(self, notificationWindow, socketObj):
        # parameters are:
        #   notificationWindow  --  the GUI window that must be notified that a message has been received
        #   socketObj           --  the socket object to listen to for messages.
        
        # Initialize the Thread object
        threading.Thread.__init__(self)
        # Remember the window that needs to be notified
        self.window = notificationWindow
        # Remember the socket object
        self.socketObj = socketObj
        # Signal that we don't yet want to abort the thread
        self._want_abort = False
        # prevent the application from hanging on Close
        self.setDaemon(1)
        # Start the thread
        self.start()

    def run(self):
        # We need to track overflow (incomplete messages) from the socket from when the socket buffer
        # gets full.  
        dataOverflow = ''
        # As long as the thread is running ...
        while not self._want_abort:
            try:
                # listen to the socket for a message from the Message Server
                newData = self.socketObj.recv(2048)
                # if we're using Unicode, we need to decode the socket message.
                # Either way, we need to add the socket message to whatever overflow is unprocessed.
                if 'unicode' in wx.PlatformInfo:
                    data = dataOverflow + newData.decode('utf8')
                else:
                    data = dataOverflow + newData
            except socket.error:

                if DEBUG:
                    print 'ChatWindow.ListenerThread.run() socket error'

                # NOTE:  No GUI from inside the thread.  We use a MessageServerLost Event instead.
                wx.PostEvent(self.window, MessageServerLostEvent())
                # Signal that the thread should close itself.
                self._want_abort = True
                # Post a Close Message
                wx.PostEvent(self.window, CloseMessageEvent())

                if DEBUG:
                    print "ChatWindow.ListenerThread.run() about to break"
                    
                # Break out of the while loop
                break

            except:

                if DEBUG:
                    print "ChatWindow.ListenerThread.run() other error"

            if DEBUG:
                print self.socketObj, 'received "%s"' % data.encode('latin1')

            # As long as data should be processed, and the data is not blank ...
            if not self._want_abort and (data != ''):
                # Lock the thread
                threadLock.acquire()

                # A single data element from the Message Server can contain multiple Transana messages, separated
                # by the ' ||| ' message separator.  Break a Message Server message into its component messages.
                messages = data.split(' ||| ')
                # The last segment is USUALLY '', but when we get to > 80 connections, that's not always
                # true.  It might be an incomplete message segment.  Therefore, we'll remember it and
                # stick it on the front of the NEXT message.  If it's blank, like it's supposed to be,
                # then it will make no difference, but if it's NOT blank, it would otherwise get lost.
                dataOverflow = messages[-1]
                
                # Process all the messages.  (The last message is always blank or overflow, so skip it!)
                for message in messages[:-1]:
                    # Post the Message Event with the message as data
                    wx.PostEvent(self.window, PostMessageEvent(message))
                # Unlock the thread when we're done processing all the messages.
                threadLock.release()
            else:

                if DEBUG:
                    print "Blank message received.  This may mean that the socket connection was lost!!\nOr it may not."
                    
                # NOTE:  No GUI from inside the thread.  We use a MessageServerLost Event instead.
                if self.window.reportSocketLoss:
                    wx.PostEvent(self.window, MessageServerLostEvent())

                return
                
        else:
            # If we want to abort the thread, post the message event with None as the data
            # NOTE:  This probably never gets called, as the code waits at the socket.recv()
            #        line until it gets a message of the connection is broken.
            # wx.PostEvent(self.window, PostMessageEvent(None))
            return

    def abort(self):
        # Signal that you want to abort the thread.  (Probably not effective, as the thread is
        # waiting on a socket.recv() call)
        self._want_abort = True


class ChatWindow(wx.Frame):
    """ This window displays the Chat form. """
    def __init__(self,parent,id,title, socketObj):
        # Remember the parent window
        self.parent = parent
        # Remember the window title
        self.title = title
        # remember the Socket object
        self.socketObj = socketObj
        # Get the username from the TransanaGlobal module
        self.userName = TransanaGlobal.userName
        # Define the main Frame for the Chat Window
        wx.Frame.__init__(self, parent, -1, title, size = (710,450), style=wx.DEFAULT_FRAME_STYLE|wx.NO_FULL_REPAINT_ON_RESIZE)
        # Set the background to White
        self.SetBackgroundColour(wx.WHITE)
        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)
        # Set the Chat Window Icon
        transanaIcon = wx.Icon(TransanaGlobal.programDir + os.sep + "images/Transana.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(transanaIcon)

        # Create a Sizer for the form
        box = wx.BoxSizer(wx.VERTICAL)

        # Create a Sizer for data the form receives
        boxRecv = wx.BoxSizer(wx.HORIZONTAL)

        # Create a Sizer for the Memo section
        boxMemo = wx.BoxSizer(wx.VERTICAL)
        # Create a label for the Memo section
        self.txtMemo = wx.StaticText(self, -1, _("Messages"))
        # Put the label in the Memo Sizer, with a little padding below
        boxMemo.Add(self.txtMemo, 0, wx.BOTTOM, 3)
        # Add a TextCtrl for the chat text.  This is read only, as it is filled programmatically.
        self.memo = wx.TextCtrl(self, -1, style = wx.TE_MULTILINE | wx.TE_WORDWRAP | wx.TE_READONLY)
        # Put the Memo control in the Memo Sizer
        boxMemo.Add(self.memo, 1, wx.EXPAND)
        # Add the Memo display to the Receiver Sizer
        boxRecv.Add(boxMemo, 8, wx.EXPAND)

        # Create a Sizer for the User section
        boxUser = wx.BoxSizer(wx.VERTICAL)
        # Create a label for the User section
        self.txtUser = wx.StaticText(self, -1, _("Current Users"))
        # Put the label in the User Sizer with a little padding below
        boxUser.Add(self.txtUser, 0, wx.BOTTOM, 3)
        # Add a ListBox to hold the names of active users
        self.userList = wx.ListBox(self, -1, choices=[self.userName], style=wx.LB_SINGLE)
        boxUser.Add(self.userList, 1, wx.BOTTOM | wx.EXPAND, 3)

        # Add a checkbox to enable/disable audio feedback
        self.useSound = wx.CheckBox(self, -1, _("Sound Enabled"))
        # Check the box to start
        self.useSound.SetValue(True)
        # Add the checkbox to the sizer
        boxUser.Add(self.useSound, 0)
        
        # Add the User List display to the Receiver Sizer
        boxRecv.Add(boxUser, 2, wx.EXPAND | wx.LEFT, 6)
        
        # Put the TextCtrl in the form sizer
        box.Add(boxRecv, 13, wx.EXPAND | wx.ALL, 4)

        # Create a sizer for the text entry and send portion of the form
        boxSend = wx.BoxSizer(wx.HORIZONTAL)
        # Add a TextCtrl where the user can enter messages.  It needs to process the Enter key.
        self.txtEntry = wx.TextCtrl(self, -1, style = wx.TE_PROCESS_ENTER)
        # Add the Text Entry control to the Text entry sizer
        boxSend.Add(self.txtEntry, 5, wx.EXPAND)
        # bind the OnSend event with the Text Entry control's enter key event
        self.txtEntry.Bind(wx.EVT_TEXT_ENTER, self.OnSend)

        # Add a "Send" button
        self.btnSend = wx.Button(self, -1, _("Send"))
        # Add the Send button to the Text Entry sizer
        boxSend.Add(self.btnSend, 0, wx.LEFT, 6)
        # Bind the OnSend event to the send button's press event
        self.btnSend.Bind(wx.EVT_BUTTON, self.OnSend)
        
        # Add a "Clear" button
        self.btnClear = wx.Button(self, -1, _("Clear"))
        # Add the Clear button to the Text Entry sizer
        boxSend.Add(self.btnClear, 0, wx.LEFT, 6)
        # Bind the OnClear event to the Clear button's press event
        self.btnClear.Bind(wx.EVT_BUTTON, self.OnClear)
        
        # Add the send Sizer to the Form Sizer
        box.Add(boxSend, 1, wx.EXPAND | wx.ALL, 4)

        # Define the Chat Sound
        flName = TransanaGlobal.programDir + os.sep + 'images' + os.sep + 'chatmessage.wav'
        # Create the player for the Chat Sound
        self.soundplayer = wx.media.MediaCtrl(self, -1, fileName = flName)
        # Hide the player for the Chat Sound.  It doesn't need to be visible.
        self.soundplayer.Show(False)

        # Define the form's OnClose handler
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # Attach the Form's Main Sizer to the form
        self.SetSizer(box)
        # Set AutoLayout on
        self.SetAutoLayout(True)
        # Lay out the form
        self.Layout()
        # Set the minimum size for the form.
        self.SetSizeHints(minW = 600, minH = 440)
        # Center the form on the screen
        self.CentreOnScreen()
        # Set the initial focus to the Text Entry control
        self.txtEntry.SetFocus()
        # We need to know if loss of socket connection is expected or should be reported
        self.reportSocketLoss = True
        # Start the Listener Thread to listen for messages from the Message Server
        self.listener = ListenerThread(self, self.socketObj)

        # Define the custom Post Message Event's handler
        EVT_POST_MESSAGE(self, self.OnPostMessage)
        # Define the custom Post Message Event's handler
        EVT_MESSAGESERVER_LOST(self, self.OnMessageServerLost)

        if DEBUG and False:
            print
            print "Threads:"
            for thr in threading.enumerate():
                if type(thr).__name__ == 'ListenerThread':
                    print thr, type(thr).__name__, thr.socketObj
                else:
                    print thr, type(thr).__name__
            print
            
        # Define the custom Close Message Event's handler
        EVT_CLOSE_MESSAGE(self, self.OnCloseMessage)

        # Register with the Control Object
        if TransanaGlobal.menuWindow.ControlObject != None:
            TransanaGlobal.menuWindow.ControlObject.Register(Chat=self)
            self.ControlObject = TransanaGlobal.menuWindow.ControlObject
        else:
            self.ControlObject = None

            if DEBUG:
                print "ChatWindow has NO CONTROLOBJECT *********************************************"

        if DEBUG:
            print "Sending Connection Message"
            
        # Inform the message server of user's identity and database connection
        if 'unicode' in wx.PlatformInfo:
            userName = self.userName.encode('utf8')
            host = TransanaGlobal.configData.host.encode('utf8')
            db = TransanaGlobal.configData.database.encode('utf8')
            
            if DEBUG:
                print 'C %s %s %s 220 ||| ' % (userName, host, db)

            # If we are running the Transana Client on the same computer as the MySQL server, we MUST refer to it as localhost.
            # In this circumstance, this copy of the Transana Client will not be recognized by the Transana Message Server
            # as being connected to the same database as other computers connecting to it.  To get around this, we need to
            # get the correct Server name from the user.

            # Detect the use of "localhost" 
            if host.lower() == 'localhost':
                # Create a Text Entry Dialog to get the proper server name from the user.
                dlg = wx.TextEntryDialog(self, _('What is the Host / Server name other computers use to connect to this MySQL Server?'),
                                         _('Transana Message Server connection'), 'localhost')
                # Show the Text Entry.  See if the user selects "OK".
                if dlg.ShowModal() == wx.ID_OK:
                    # If so, update the host name to pass to the Transana Message Server
                    host = dlg.GetValue()
                # Destroy the Text Entry Dialog.
                dlg.Destroy()
            
            self.socketObj.send('C %s %s %s 220 ||| ' % (userName, host, db))
        else:
            self.socketObj.send('C %s %s %s 220 ||| ' % (self.userName, TransanaGlobal.configData.host, TransanaGlobal.configData.database))

        # Create a Timer to check for Message Server validation.
        # Initialize to unvalidated state
        self.serverValidation = False
        # Create a Timer
        self.validationTimer = wx.Timer()
        # Assign the Timer's event
        self.validationTimer.Bind(wx.EVT_TIMER, self.OnValidationTimer)
        # 10 seconds should be sufficient for the connection to the message server to be established and confirmed
        self.validationTimer.Start(10000)

    def SendMessage(self, message):
        """ Send a message through the chatWindow's socket """
        try:
            msg = '%s ||| ' % message
            # If we're using Unicode, we need to encode the messages passed to the socket.
            if 'unicode' in wx.PlatformInfo:
                self.socketObj.send(msg.encode('utf8'))
            else:
                self.socketObj.send(msg)
            # I added this to try to prevent messages from bumping into one another and stacking up.
            # It *should* allow messages to be sent totally independently.
            try:
                wx.Yield()
                if DEBUG:
                    print "ChatWindow.SendMessage():  Yield called."
            except:
                if DEBUG:
                    print "ChatWindow.SendMessage():  Yield FAILED.", sys.exc_info()[0],sys.exc_info()[1]
                pass

            
        except socket.error:
            if DEBUG:
                print "ChatWindow.SendMessage() socket error."

            # NOTE:  No GUI from inside the thread.  We use a MessageServerLost Event instead.
            wx.PostEvent(self.window, MessageServerLostEvent())
            # Close the Chat Window if the connection's been broken.
            self.Close()

    def OnSend(self, event):
        """ Send Message handler """
            # indicate that this is a Text Message by prefacing the text with "M".
        self.SendMessage('M %s' % self.txtEntry.GetValue())
        # Clear the Text Entry control
        self.txtEntry.SetValue('')
        # Set the focus to the Text Entry Control
        self.txtEntry.SetFocus()

    def OnClear(self, event):
        """ Clear button handler """
        # Clear the previous chat text
        self.memo.Clear()
        # Set the focus to the Text Entry control
        self.txtEntry.SetFocus()
        
    def OnPostMessage(self, event):
        """ Post Message handler """

        def ConvertMessageToNodeList(message):
            """ Take a message from the Transana Message Server and convert it to a NodeList for use with the DB Tree
                add_Node method """
            nodelist = ()
            for m in message.split(' >|< '):
                nodelist += (m,)
            return nodelist
                
        # If there is data in the message event ...
        if event.data != None:

            if DEBUG:
                print 'event.data = "%s"' % event.data.encode('latin1')

            message = event.data
            messageHeader = message[:message.find(' ')]
            message = message[message.find(' ') + 1:].strip()
            messageSender = message[:message.find(' ') - 1]  # drop the ":"
            message = message[message.find(' ') + 1:].strip()

            if DEBUG:
                print 'messageHeader = "%s"' % messageHeader
                print 'messageSender = "%s"' % messageSender
                print 'message = "%s"' % message.encode('latin1')
                print
            
            # Determine what type of message it is by looking at the first character.                
            # Text Message ?
            if messageHeader == 'M':
                # If it's not visible ...
                if not self.IsShown():
                    # ... show the ChatWindow and ...
                    self.Show(True)
                # ... display the message (minus the Message Prefix) on the screen
                self.memo.AppendText("%s\n" % event.data[2:])
                # If sound is enabled...
                if self.useSound.IsChecked():
                    # ... play the message sound.
                    self.soundplayer.Play()
                
            # Connection Message ?
            elif messageHeader == 'C':
                # This signals that a UserName should be added to the list of current users.
                # Drop the Message Prefix.
                st = event.data[2:]
                # The first "word" is the username.  Text after the first space should be dropped.
                if st.find(' ') > -1:
                    st = st[:st.find(' ')]
                # Only add a user name if it's not redundant.  (The message server should prevent this
                # from occurring.)
                if st != self.userName:
                    self.userList.Append(st)
                    
            # Rename Message ?
            elif messageHeader == 'R':
                # If a user name is duplicated, the Message Server renames it to prevent confusion.
                # This code indicates that THIS USER's account has been renamed by the Message Server.
                # Remove the "old" user name
                self.userList.Delete(0)
                # Drop the Message Prefix
                st = event.data[2:]
                # The first "word" is the username.  Text after the first space should be dropped.
                if st.find(' ') > -1:
                    self.userName = st[:st.find(' ')]
                else:
                    self.userName = st
                # self.SetStatusText(_('Transana Message Server "%s" on port %d.  User "%s" on Database Server "%s", Database "%s".') % (SERVERHOST, SERVERPORT, self.userName, TransanaGlobal.configData.host, TransanaGlobal.configData.database))
                self.userList.Append(st)
                
            # Import Message
            elif messageHeader == 'I':
                # Another user has imported a database.  We need to refresh the whole Database Tree!
                # See if a Control Object has been defined.
                if self.ControlObject != None:
                    # Update the Data Window via the Control Object
                    self.ControlObject.UpdateDataWindow()
                    
            # Server Validation ?
            elif messageHeader == 'V':
                # Indicate that the server has been validated.  The Validation Timer processes this later.
                self.serverValidation = True
                
            # Disconnect Message ?
            elif messageHeader == 'D':
                # Remove the Message Prefix.
                st = event.data[2:]
                # The remainder of the message is the username to be removed.  Delete it from the User List.
                self.userList.Delete(self.userList.FindString(st))
                
            else:
                # The remaining messages should not be processed if this user was the message sender
                if self.userName != messageSender:
                    # We can't have the tree selection changing because of the activity of other users.  That creates all kinds of
                    # problems if we're in the middle of editing something.  So let's note the current selection
                    currentSelection = self.ControlObject.DataWindow.DBTab.tree.GetSelection()
                    # The Control Object MUST be defined (and always will be)
                    if self.ControlObject != None:
                        # Add Series Message
                        if messageHeader == 'AS':
                            tempSeries = Series.Series(message)
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('SeriesNode', (_('Series'), message), tempSeries.number, None, False, avoidRecursiveYields = True)
                            
                        # Add Episode Message
                        elif messageHeader == 'AE':
                            nodelist = ConvertMessageToNodeList(message)
                            tempEpisode = Episode.Episode(series=nodelist[0], episode=nodelist[1])
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('EpisodeNode', (_('Series'),) + nodelist, tempEpisode.number, tempEpisode.series_num, False, avoidRecursiveYields = True)

                        # Add Transcript Message
                        elif messageHeader == 'AT':
                            nodelist = ConvertMessageToNodeList(message)
                            tempEpisode = Episode.Episode(series=nodelist[0], episode=nodelist[1])
                            tempTranscript = Transcript.Transcript(nodelist[-1], ep=tempEpisode.number)
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('TranscriptNode', (_('Series'),) + nodelist, tempTranscript.number, tempEpisode.number, False, avoidRecursiveYields = True)

                        # Add Collection Message
                        elif messageHeader == 'AC':
                            nodelist = ConvertMessageToNodeList(message)
                            parentNum = 0
                            for coll in nodelist:
                                tempCollection = Collection.Collection(coll, parentNum)
                                parentNum = tempCollection.number
                            # avoidRecursiveYields added to try to prevent a problem on the Mac when converting Searches
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('CollectionNode', (_('Collections'),) + nodelist, tempCollection.number, tempCollection.parent, False, avoidRecursiveYields=True)

                        # Add Clip Message
                        elif messageHeader == 'ACl':
                            nodelist = ConvertMessageToNodeList(message)
                            parentNum = 0
                            for coll in nodelist[:-1]:
                                tempCollection = Collection.Collection(coll, parentNum)
                                parentNum = tempCollection.number
                            tempClip = Clip.Clip(nodelist[-1], tempCollection.id, tempCollection.parent)
                            # avoidRecursiveYields added to try to prevent a problem on the Mac when converting Searches
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('ClipNode', (_('Collections'),) + nodelist, tempClip.number, tempCollection.number, False, avoidRecursiveYields=True)

                            # If we are moving a Clip, the clip's Notes need to travel with the Clip.  The first step is to
                            # get a list of those Notes.
                            noteList = DBInterface.list_of_notes(Clip=tempClip.number)
                            # If there are Clip Notes, we need to make sure they travel with the Clip
                            if noteList != []:
                                insertNode = self.ControlObject.DataWindow.DBTab.tree.select_Node((_('Collections'),) + nodelist, 'ClipNode')
                                # We accomplish this using the TreeCtrl's "add_note_nodes" method
                                self.ControlObject.DataWindow.DBTab.tree.add_note_nodes(noteList, insertNode, Clip=tempClip.number)
                                self.ControlObject.DataWindow.DBTab.tree.Refresh()

                        # Add Clip in Sort Order Message
                        elif messageHeader == 'AClSO':
                            nodelist = ConvertMessageToNodeList(message)
                            parentNum = 0
                            for coll in nodelist[:-2]:
                                tempCollection = Collection.Collection(coll, parentNum)
                                parentNum = tempCollection.number
                            # We need the NODE for the Clip we should place the new clip in front of.  Let's get that here.
                            insertNode = self.ControlObject.DataWindow.DBTab.tree.select_Node((_('Collections'),) + nodelist[:-1], 'ClipNode')
                            tempClip = Clip.Clip(nodelist[-1], tempCollection.id, tempCollection.parent)
                            # Add new node, leaving the insertNode out of the nodeList.
                            # avoidRecursiveYields added to try to prevent a problem on the Mac when converting Searches
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('ClipNode', (_('Collections'),) + nodelist[:-2] + (nodelist[-1],), tempClip.number, tempCollection.number, False, insertNode, avoidRecursiveYields=True)
                            # If we are moving a Clip, the clip's Notes need to travel with the Clip.  The first step is to
                            # get a list of those Notes.
                            noteList = DBInterface.list_of_notes(Clip=tempClip.number)
                            # If there are Clip Notes, we need to make sure they travel with the Clip
                            if noteList != []:
                                insertNode = self.ControlObject.DataWindow.DBTab.tree.select_Node((_('Collections'),) + nodelist[:-2] + (nodelist[-1],), 'ClipNode')
                                # We accomplish this using the TreeCtrl's "add_note_nodes" method
                                self.ControlObject.DataWindow.DBTab.tree.add_note_nodes(noteList, insertNode, Clip=tempClip.number)
                                self.ControlObject.DataWindow.DBTab.tree.Refresh()

                        # Add Note Message
                        elif messageHeader in ['ASN', 'AEN', 'ATN', 'ACN', 'AClN']:
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            # Initialize variables
                            parentNum = 0
                            objectType = None
                            nodeType = None
                            tempObj = None
                            parentNum = 0
                            nodeCount = 0
                            # Iterate through the node list to figure out what kind of Note we're looking at
                            for node in nodelist[:-1]:
                                # Count how far into the list we are
                                nodeCount += 1
                                # If the first entry in the node list is the "Series" Root Node ...
                                if (objectType == None) and (node == 'Series'):
                                    # ... then we're climbing up the Series branch, and are at a Series record.
                                    objectType = 'Series'
                                # If we're already at a Series record ...
                                elif (objectType == 'Series'):
                                    # ... then we're moving on to an Episode next
                                    objectType = 'Episode'
                                    # We might have a Series Note, at least if we stop here!
                                    nodeType = 'SeriesNoteNode'
                                    # Let's load the Series record ...
                                    tempObj = Series.Series(node)
                                    # .. and note that the parent of the NEXT object is this series' number!
                                    parentNum = tempObj.number
                                # If we're already at an Episode record ...
                                elif (objectType == 'Episode'):
                                    # ... then we're moving on to a Transcript next
                                    objectType = 'Transcript'
                                    # we might have an Episode Note if we stop here!
                                    nodeType = 'EpisodeNoteNode'
                                    # Let's load the Episode Record
                                    tempObj = Episode.Episode(series=tempObj.id, episode=node)
                                    # .. and note that the parent of the NEXT object is this Episode's number!
                                    parentNum = tempObj.number
                                # If we're already at a Transcript record ...
                                elif (objectType == 'Transcript'):
                                    # ... then the only way to go is to a Transcript Note!
                                    objectType = 'Transcript Note'
                                    # We have a Transcript Note
                                    nodeType = 'TranscriptNoteNode'
                                    # Load the Transcript record ...
                                    tempObj = Transcript.Transcript(node, ep=parentNum)
                                    # ... and note that the parent of the Transcript Note is this Trasncript.
                                    parentNum = tempObj.number
                                # If our node is the Collections Root Node ...
                                elif (objectType == None) and (node == 'Collections'):
                                    # ... then the first level of object we're looking at is a Collection.
                                    objectType = 'Collections'
                                # if we're looking at a Collection and either we don't have a Clip Note or we're not at the end of the list yet...
                                elif (objectType == 'Collections') and ((messageHeader != 'AClN') or (nodeCount < len(nodelist) - 1)):
                                    # ... then we're still looking at a Collection
                                    objectType = 'Collections'
                                    # ... and if we stop here, we've got a Collection Note
                                    nodeType = 'CollectionNoteNode'
                                    # Load the Collection
                                    tempObj = Collection.Collection(node, parentNum)
                                    # ... and note that the collection is the parent of the NEXT object.
                                    parentNum = tempObj.number
                                # if we're looking at a Collection and we have a Clip Note and we're at the end of the list ...
                                elif (objectType == 'Collections') and (messageHeader == 'AClN') and (nodeCount == len(nodelist) - 1):
                                    # ... then we're looking at a Clip
                                    objectType = 'Clip'
                                    # ... and we're dealing with a Clip Note
                                    nodeType = 'ClipNoteNode'
                                    # Load the Clip ...
                                    tempObj = Clip.Clip(node, tempObj.id, tempObj.parent)
                                    # ... and note its number as the parent number of the Note
                                    parentNum = tempObj.number
                            # Initialize the Temporary Note object
                            tempNote = None
                            # Load the Note, which we do a bit differently based on what kind of parent object we have.
                            if nodeType == 'SeriesNoteNode':
                                tempNote = Note.Note(nodelist[-1], Series=tempObj.number)
                            elif nodeType == 'EpisodeNoteNode':
                                tempNote = Note.Note(nodelist[-1], Episode=tempObj.number)
                            elif nodeType == 'TranscriptNoteNode':
                                tempNote = Note.Note(nodelist[-1], Transcript=tempObj.number)
                            elif nodeType == 'CollectionNoteNode':
                                tempNote = Note.Note(nodelist[-1], Collection=tempObj.number)
                            elif nodeType == 'ClipNoteNode':
                                tempNote = Note.Note(nodelist[-1], Clip=tempObj.number)
                            # Add the Note to the Database Tree
                            self.ControlObject.DataWindow.DBTab.tree.add_Node(nodeType, nodelist, tempNote.number, tempObj.number, False, avoidRecursiveYields = True)
                            # If the Notes Browser is open ...
                            if self.ControlObject.NotesBrowserWindow != None:
                                # ... add the Note to the Notes Browser
                                self.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('A', tempNote)

                        # Add Keyword Group Message
                        elif messageHeader == 'AKG':
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('KeywordGroupNode', (_('Keywords'),) + (message, ), 0, 0, False, avoidRecursiveYields = True)
                            # Once we've added the Keyword Group, we need to update the Keyword Groups Data Structure
                            self.ControlObject.DataWindow.DBTab.tree.updateKWGroupsData()

                        # Add Keyword Message
                        elif messageHeader == 'AK':
                            nodelist = ConvertMessageToNodeList(message)
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('KeywordNode', (_('Keywords'),) + nodelist, 0, nodelist[0], False, avoidRecursiveYields = True)

                        # Add Keyword Example Message
                        elif messageHeader == 'AKE':
                            # The first message parameter for a Keyword Example is the Clip Number
                            nodelist = ConvertMessageToNodeList(message)
                            tempClip = Clip.Clip(int(nodelist[0]))
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('KeywordExampleNode', (_('Keywords'),) + nodelist[1:], tempClip.number, tempClip.collection_num, False, avoidRecursiveYields = True)

                        # Rename a Node
                        elif messageHeader == 'RN':
                            nodelist = ConvertMessageToNodeList(message)
                            # The first element in the nodelist is the nodeType, which we need for the rename_Node call.
                            # The second element in the nodelist is the UNTRANSLATED root node label.  This avoids problems
                            # in mixed-language environments.  But we now need to translate it.
                            # The last element is the name the Node should be changed to.
                            # One more wrinkle -- the Root Node label might be a string, or it might already be a Unicode
                            # object.  It needs to be handled differently.  (In English, it's unicode, otherwise it's a string.)
                            if type(_(nodelist[1])) == type(u''):
                                tmpRootNode = _(nodelist[1])
                            else:
                                tmpRootNode = unicode(_(nodelist[1]), 'utf8')
                            
                            nodelist = (nodelist[0], tmpRootNode) + nodelist[2:]
                            
                            if DEBUG:
                                tmpstr = "Calling rename_Node(%s, %s, %s) %s %s" % (nodelist[1:-1], nodelist[0], nodelist[-1], \
                                         type(nodelist[0]), type(nodelist[-1]))
                                print tmpstr.encode('latin1')
                                print
                                
                            self.ControlObject.DataWindow.DBTab.tree.rename_Node(nodelist[1:-1], nodelist[0], nodelist[-1])
                            
                            # If we're removing a Keyword Group ...
                            if nodelist[0] == 'KeywordGroupNode':
                                # ... we need to update the Keyword Groups Data Structure
                                self.ControlObject.DataWindow.DBTab.tree.updateKWGroupsData()

                            # If we're renaming  a Keyword ...
                            elif nodelist[0] == 'KeywordNode':
                                # ... see if we have an Episode or Clip object currently loaded ...
                                if isinstance(self.ControlObject.currentObj, Episode.Episode) or isinstance(self.ControlObject.currentObj, Clip.Clip):
                                    # ... let's see if the Keywords Tab is being shown ...
                                    if self.ControlObject.DataWindow.nb.GetPageText(self.ControlObject.DataWindow.nb.GetSelection()) == unicode(_('Keywords'), 'utf8'):
                                        # ... and if so, iterate through its keywords ...
                                        for kw in self.ControlObject.currentObj.keyword_list:
                                            # ... and see if it contains the keyword that was changed.
                                            if (nodelist[-3].upper() == kw.keywordGroup.upper()) and (nodelist[-2].upper() == kw.keyword.upper()):
                                                # If so, update it.  (Its Refresh() method updates data from the database.)
                                                self.ControlObject.DataWindow.KeywordsTab.Refresh()
                                                # ... and refresh the keyword list
                                                self.ControlObject.currentObj.refresh_keywords()
                                                break

                            # If we're renaming a Note ...
                            elif nodelist[0] in ['SeriesNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'CollectionNoteNode', 'ClipNoteNode']:
                                # ... if the Notes Browser is open, we need to update the note there as well.
                                if self.ControlObject.NotesBrowserWindow != None:
                                    # The first element in the nodelist is the NOTE Node Type.
                                    nodeType = nodelist[0]
                                    # The NOTE Node type can be dropped from the node list
                                    nodelist = nodelist[1:]
                                    # Initialize variables
                                    parentNum = 0
                                    objectType = None
                                    tempObj = None
                                    parentNum = 0
                                    nodeCount = 0
                                    # Iterate through the node list to figure out what kind of Note we're looking at,
                                    # (skipping the old and new note names, which aren't needed here!)
                                    for node in nodelist[:-2]:
                                        # Keep track of our position in the list.
                                        nodeCount += 1
                                        # If the first entry in the node list is the "Series" Root Node ...
                                        if (objectType == None) and (node == unicode(_('Series'), 'utf8')):
                                            # ... then we're climbing up the Series branch, and are at a Series record.
                                            objectType = 'Series'
                                        # If we're already at a Series record ...
                                        elif (objectType == 'Series'):
                                            # ... then we're moving on to an Episode next
                                            objectType = 'Episode'
                                            # Let's load the Series record ...
                                            tempObj = Series.Series(node)
                                            # .. and note that the parent of the NEXT object is this series' number!
                                            parentNum = tempObj.number
                                        # If we're already at an Episode record ...
                                        elif (objectType == 'Episode'):
                                            # ... then we're moving on to a Transcript next
                                            objectType = 'Transcript'
                                            # Let's load the Episode Record
                                            tempObj = Episode.Episode(series=tempObj.id, episode=node)
                                            # .. and note that the parent of the NEXT object is this Episode's number!
                                            parentNum = tempObj.number
                                        # If we're already at a Transcript record ...
                                        elif (objectType == 'Transcript'):
                                            # ... then the only way to go is to a Transcript Note!
                                            objectType = 'Transcript Note'
                                            # Load the Transcript record ...
                                            tempObj = Transcript.Transcript(node, ep=parentNum)
                                            # ... and note that the parent of the Transcript Note is this Trasncript.
                                            parentNum = tempObj.number
                                        # If our node is the Collections Root Node ...
                                        elif (objectType == None) and (node == unicode(_('Collections'), 'utf8')):
                                            # ... then the first level of object we're looking at is a Collection.
                                            objectType = 'Collections'
                                        # if we're looking at a Collection and either we don't have a Clip Note or we're not at the end of the list yet...
                                        elif (objectType == 'Collections') and ((nodeType == 'CollectionNoteNode') or (nodeCount < len(nodelist) - 2)):
                                            # ... then we're still looking at a Collection
                                            objectType = 'Collections'
                                            # Load the Collection
                                            tempObj = Collection.Collection(node, parentNum)
                                            # ... and note that the collection is the parent of the NEXT object.
                                            parentNum = tempObj.number
                                        # if we're looking at a Collection and we have a Clip Note and we're at the end of the list ...
                                        elif (objectType == 'Collections') and (nodeType == 'ClipNoteNode') and (nodeCount == len(nodelist) - 2):
                                            # ... then we're looking at a Clip
                                            objectType = 'Clip'
                                            # Load the Clip ...
                                            tempObj = Clip.Clip(node, tempObj.id, tempObj.parent)
                                            # ... and note its number as the parent number of the Note
                                            parentNum = tempObj.number
                                    # Initialize the Temporary Note object
                                    tempNote = None
                                    # Load the Note, which we do a bit differently based on what kind of parent object we have.
                                    if nodeType == 'SeriesNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Series=tempObj.number)
                                    elif nodeType == 'EpisodeNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Episode=tempObj.number)
                                    elif nodeType == 'TranscriptNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Transcript=tempObj.number)
                                    elif nodeType == 'CollectionNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Collection=tempObj.number)
                                    elif nodeType == 'ClipNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Clip=tempObj.number)
                                    # Rename the Note in the Database Tree
                                    self.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('R', tempNote, oldName=nodelist[-2])

                        # Delete Node
                        elif messageHeader == 'DN':
                            nodelist = ConvertMessageToNodeList(message)
                            # Check the TYPE of the translated second element.
                            if type(_(nodelist[1])).__name__ == 'str':
                                # If string, translate it and convert it to unicode
                                nodelist = (nodelist[0],) + (unicode(_(nodelist[1]), 'utf8'),) + nodelist[2:]
                            # If not string, it's unicode!
                            else:
                                # ... in which case, we just translate it.
                                nodelist = (nodelist[0],) + (_(nodelist[1]),) + nodelist[2:]
                            self.ControlObject.DataWindow.DBTab.tree.delete_Node(nodelist[1:], nodelist[0], sendMessage=False)
                            # If we're removing a Keyword Group ...
                            if nodelist[0] == 'KeywordGroupNode':
                                # ... we need to update the Keyword Groups Data Structure
                                self.ControlObject.DataWindow.DBTab.tree.updateKWGroupsData()
                                
                                # ... see if we have an Episode or Clip object currently loaded ...
                                if isinstance(self.ControlObject.currentObj, Episode.Episode) or isinstance(self.ControlObject.currentObj, Clip.Clip):
                                    # ... let's see if the Keywords Tab is being shown ...
                                    if self.ControlObject.DataWindow.nb.GetPageText(self.ControlObject.DataWindow.nb.GetSelection()) == unicode(_('Keywords'), 'utf8'):
                                        # ... and if so, iterate through its keywords ...
                                        for kw in self.ControlObject.currentObj.keyword_list:
                                            # ... and see if it contains the keyword that was changed.
                                            if (nodelist[-1].upper() == kw.keywordGroup.upper()):
                                                # If so, update it.  (Its Refresh() method updates data from the database.)
                                                self.ControlObject.DataWindow.KeywordsTab.Refresh()
                                                # ... and refresh the keyword list
                                                self.ControlObject.currentObj.refresh_keywords()
                                                break

                            # If we're deleting  a Keyword ...
                            elif nodelist[0] == 'KeywordNode':
                                # ... see if we have an Episode or Clip object currently loaded ...
                                if isinstance(self.ControlObject.currentObj, Episode.Episode) or isinstance(self.ControlObject.currentObj, Clip.Clip):
                                    # ... let's see if the Keywords Tab is being shown ...
                                    if self.ControlObject.DataWindow.nb.GetPageText(self.ControlObject.DataWindow.nb.GetSelection()) == unicode(_('Keywords'), 'utf8'):
                                        # ... and if so, iterate through its keywords ...
                                        for kw in self.ControlObject.currentObj.keyword_list:
                                            # ... and see if it contains the keyword that was changed.
                                            if (nodelist[-2].upper() == kw.keywordGroup.upper()) and (nodelist[-1].upper() == kw.keyword.upper()):
                                                # If so, update it.  (Its Refresh() method updates data from the database.)
                                                self.ControlObject.DataWindow.KeywordsTab.Refresh()
                                                # ... and refresh the keyword list
                                                self.ControlObject.currentObj.refresh_keywords()
                                                break
                            # If we're deleting a Note Node ...
                            elif nodelist[0] in ['SeriesNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'CollectionNoteNode', 'ClipNoteNode']:
                                # ... and the Notes Browser is open, we need to delete the Note from there too.
                                if self.ControlObject.NotesBrowserWindow != None:
                                    # Determine the Note Browser's root node based on the type of Note we're deleting
                                    if nodelist[0] == 'SeriesNoteNode':
                                        nodeType = 'Series'
                                    elif nodelist[0] == 'EpisodeNoteNode':
                                        nodeType = 'Episode'
                                    elif nodelist[0] == 'TranscriptNoteNode':
                                        nodeType = 'Transcript'
                                    elif nodelist[0] == 'CollectionNoteNode':
                                        nodeType = 'Collection'
                                    elif nodelist[0] == 'ClipNoteNode':
                                        nodeType = 'Clip'
                                    else:
                                        nodeType = None
                                    # The Note Object has already been DELETED, so we can't load the Note itself!
                                    # Therefore, we must build its NodeList here and pass it!  We pass the UNTRANSLATED
                                    # object type.
                                    if nodeType != None:
                                        self.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('D', (nodeType, nodelist[-1]))

                        # Update Keyword List
                        elif messageHeader == 'UKL':
                            # Parse the message at the space into object type and object number
                            msgData = message.split(' ')
                            # See if the currently loaded object matches the object described in the message.
                            if ((isinstance(self.ControlObject.currentObj, Episode.Episode) and \
                                 (msgData[0] == 'Episode')) or \
                                (isinstance(self.ControlObject.currentObj, Clip.Clip) and \
                                 (msgData[0] == 'Clip'))) and \
                               (self.ControlObject.currentObj.number == int(msgData[1])):
                                # Let's see if the Keywords Tab is being shown
                                if self.ControlObject.DataWindow.nb.GetPageText(self.ControlObject.DataWindow.nb.GetSelection()) == unicode(_('Keywords'), 'utf8'):
                                    # If so, update it.  (Its Refresh() method updates data from the database.)
                                    self.ControlObject.DataWindow.KeywordsTab.Refresh()
                                    # ... and refresh the keyword list
                                    self.ControlObject.currentObj.refresh_keywords()

                        # Update Keyword Visualization
                        elif messageHeader == 'UKV':
                            # Parse the message at the space into object type, object number, and possible Episode Number (if Clip)
                            msgData = message.split(' ')
                            # If no Object Type ...
                            if msgData[0] == 'None':
                                # ... we need to update the keyword visualization no matter what.
                                self.ControlObject.UpdateKeywordVisualization()
                            # if Object Type is Episode ...
                            elif msgData[0] == 'Episode':
                                # See if the currently loaded episode matches the episode number sent from the Message Server
                                if isinstance(self.ControlObject.currentObj, Episode.Episode) and \
                                   self.ControlObject.currentObj.number == int(msgData[1]):
                                    # ... we need to update the keyword visualization no matter what.
                                    self.ControlObject.UpdateKeywordVisualization()
                            # if Object Type is Clip ...
                            elif msgData[0] == 'Clip':
                                # See if the currently loaded episode matches the episode number sent from the Message Server
                                # or the currently loaded Clip matches the Clip Number sent from the Message Server
                                if (isinstance(self.ControlObject.currentObj, Episode.Episode) and \
                                   self.ControlObject.currentObj.number == int(msgData[2])) or \
                                   (isinstance(self.ControlObject.currentObj, Clip.Clip) and \
                                   self.ControlObject.currentObj.number == int(msgData[1])):
                                    # ... we need to update the keyword visualization.
                                    self.ControlObject.UpdateKeywordVisualization()
                                    
                        else:
                            if DEBUG:
                                print "Unprocessed Message: ", event.data.encode('latin1')
                                print
                                
                            # If it's not visible ...
                            if not self.IsShown():
                                # ... show the ChatWindow.
                                self.Show(True)
                            # Inform the user of the unknown message.  This should never occur.
                            self.memo.AppendText('Unprocessed Message: "%s"\n' % event.data)
                    else:
                        if DEBUG:
                            print "ChatWindow has NO CONTROLOBJECT *********************************************"

                    # Unless we've just deleted it ...
                    if messageHeader != 'DN':
                        # ... now that we're done, we should re-select the originally-selected tree item
                        self.ControlObject.DataWindow.DBTab.tree.SelectItem(currentSelection)
                else:
                    if DEBUG:
                        print "We DON'T need to add an object, as we created it in the first place."

    def OnMessageServerLost(self, event):
        dlg = Dialogs.ErrorDialog(None, _("Your connection to the Message Server has been lost.\nYou may have lost your connection to the network, or there may be a problem with the Server.\nPlease quit Transana immediately and resolve the problem."))
        dlg.ShowModal()
        dlg.Destroy()
        self.Close()
        TransanaGlobal.chatWindow = None
        wx.CallAfter(self.Destroy)

    def OnCloseMessage(self, event):
        """ Process the custom Close Message """
        # Close the Chat Window

        if DEBUG:
            print "ChatWindow.OnCloseMessage()"

        self.Close()
        
    def OnValidationTimer(self, event):
        """ Checks for the validity of the Message Server """
        # If the server's been validated ...
        if self.serverValidation:
            # ... Stop the validation timer.  It's done.
            self.validationTimer.Stop()
        # If the server has NOT been validated ...
        else:
            # If it's not visible ...
            if not self.IsShown():
                # ... show the ChatWindow.
                self.Show(True)
            # Bring this window to the top of the others
            # self.Raise()
            # Change the Timer interval to a minute
            self.validationTimer.Start(60000)
            # Display an error message to the user.
            self.memo.AppendText(_('Transana-MU:  The Transana Message Server has not been validated.\n'));
            self.memo.AppendText(_('Transana-MU:  Your connection to the Transana Message Server may have failed, or\n'));
            self.memo.AppendText(_('Transana-MU:  you may be using an improper version of the Transana Message Server,\n'));
            self.memo.AppendText(_('Transana-MU:  one that is different than this version of Transana-MU requires.\n'));
            self.memo.AppendText(_('Transana-MU:  Please report this problem to your system administrator.\n\n'));

    def OnClose(self, event):
        """ Intercept when the Close Button is selected """
        # When the Close Button is selected, we should HIDE the form, but not Close it entirely

        if DEBUG:
            print 'ChatWindow.OnClose()'
            
        self.Show(False)
    
    def OnFormClose(self, event):
        """ Form Close handler for when the form should be destroyed, not just hidden """
        try:

            if DEBUG:
                print "ChatWindow.FormClose()"
                
            # Inform the message server that you're disconnecting (Needed on Mac)
            self.socketObj.send('D %s ||| ' % self.userName)
            # Wait a second for this message to get through.
            time.sleep(1)
            # Close the socket connection
            if TransanaGlobal.socketConnection != None:
                TransanaGlobal.socketConnection.close()
                    
        except socket.error:

            if DEBUG:
                print 'ChatWindow.OnFormClose() -- socket.error'
            
        except:

            if DEBUG:
                print "ChatWindow.OnFormClose() -- not socket.error"
            print sys.exc_info()[0], sys.exc_info()[1]
            import traceback
            print traceback.print_exc(file=sys.stdout)
        # Try to tell the listener thread to abort (probably does nothing.)
        self.listener.abort()
        # Destroy the Chat Sound player
        self.soundplayer.Destroy()
        # Go on and close the form.
        self.Close()

    def ChangeLanguages(self):
        """ Handles the change of languages """
        # Change the Window Title
        self.SetTitle(_(self.title))
        # Change the prompts
        self.txtMemo.SetLabel(_("Messages"))
        self.txtUser.SetLabel(_("Current Users"))
        # Change the Buttons
        self.btnSend.SetLabel(_("Send"))
        self.btnClear.SetLabel(_("Clear"))


def ConnectToMessageServer():
    """ Create a connection to the Transana Message Server.
        This function returns a socket object if the connection is successful,
        or None if it is not.  """
    # If there is already a Chat Window open ...
    if TransanaGlobal.chatWindow != None:

        if DEBUG:
            print "Closing ChatWindow in ConnectToMessageServer"

        # Closing the form will cause an expected Socket Loss, which should not be reported.
        TransanaGlobal.chatWindow.reportSocketLoss = False
        # ... close the Chat Form, which will in turn break the socket connection.
        TransanaGlobal.chatWindow.OnFormClose(None)
        # Allow the final messages to be processed.
        # There can be an issue with recursive calls to wxYield, so trap the exception ...
        try:
            wx.Yield()
        # ... and ignore it!
        except:
            pass
        # Destroy the Chat Window
        TransanaGlobal.chatWindow.Destroy()
        TransanaGlobal.chatWindow = None

    # Initialize a variable to indicate that no connection has been made yet.
    ConnectedToMessageServer = False
    # Keep trying until a connection is made or the user gives up.
    while not ConnectedToMessageServer:
        # Create a socket connection to the Transana Message Server
        try:
            # With Unicode, not having a Message Server defined was raising a Unicode Error rather than a
            # socket error as the code expected.  Therefore, if the Message Server is undefined, let's just
            # raise a socket error now.
            if TransanaGlobal.configData.messageServer == '':
                raise socket.error
            # Define the Socket connection
            socketObj = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            # Connect to the socket.  Server and Port are pulled from the Transana Configuration File
            socketObj.connect((TransanaGlobal.configData.messageServer, TransanaGlobal.configData.messageServerPort))
            # Create the Chat Window Frame, passing in the socket object.
            # This window is NOT SHOWN at this time.
            TransanaGlobal.chatWindow = ChatWindow(None, -1, _("Transana Chat Window"), socketObj)
            # If we get this far, the connection was successful.
            ConnectedToMessageServer = True
        except socket.error:
            # Unable to connect to the specified Message Server.
            # Build the appropriate error message.
            if TransanaGlobal.configData.messageServer != '':
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(_('You were unable to connect to a Transana Message Server at "%s".\nWould you like to change this setting?'), 'utf8') % TransanaGlobal.configData.messageServer
                else:
                    msg = _('You were unable to connect to a Transana Message Server at "%s".\nWould you like to change this setting?') % TransanaGlobal.configData.messageServer
            else:
                msg = _('No Transana Message Server has been specified.\nWould you like to specify one now?')
            # Display the error message, seek user feedback
            dlg = Dialogs.QuestionDialog(TransanaGlobal.menuWindow, msg, _('Transana Message Server Connection'))
            # Note the user's feedback
            result = dlg.LocalShowModal()
            # Destroy the message dialog
            dlg.Destroy()
            # If the user says "Yes" ...
            if result == wx.ID_YES:
                # We need to access the Program Options dialog
                import OptionsSettings
                # Display the Message Server Tab of the Transana Options dialog
                OptionsSettings.OptionsSettings(TransanaGlobal.menuWindow, 2)
                # Update the video playback speed, if it is changed.
                TransanaGlobal.menuWindow.ControlObject.VideoWindow.SetPlayBackSpeed(TransanaGlobal.configData.videoSpeed)
            # If the user does not say "Yes" ...
            else:
                # Indicate that there has not been a socket Object created.
                socketObj = None
                # break out of the "while" loop
                break
        except:
            print sys.exc_info()[0], sys.exc_info()[1]
            import traceback
            print traceback.print_exc(file=sys.stdout)
            # Indicate that there has not been a socket Object created.
            socketObj = None
            # break out of the "while" loop
            break
    # Returnt the Socket Object that was established or which was set to None if it failed.
    return socketObj


# For testing purposes, this code allows the Chat Window to function as a stand-alone application.
# Well, it used to anyway.  Now it's broken.  I don't care to fix it at this time, as I can just run
# it in the context of Transana.
if __name__ == '__main__':
    # These values come from the Config object in Transana
    SERVERHOST = 'localhost'
    SERVERPORT = 17595

    # Declare the main App object
    app = wx.PySimpleApp()
    try:
        # Define the Socket connection
        socketObj = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # Connect to the socket
        socketObj.connect((SERVERHOST, SERVERPORT))
        # Create the Chat Window Frame, passing in the socket object
        frame = ChatWindow(None, -1, _("Transana Chat Window"), socketObj)
        frame.Show()
        # run the MainLoop
        app.MainLoop()
    except socket.error:
        print _('You were unable to connect to a Transana Message Server')
    except:
        print sys.exc_info()[0], sys.exc_info()[1]
        import traceback
        print traceback.print_exc(file=sys.stdout)
