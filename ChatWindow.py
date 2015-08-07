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

""" A Message Client that processes inter-instance chat communication between Transana-MU
    clients.  This client utility requires connection to a Transana Message Server. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "ChatWindow DEBUG is ON!"

VERSION = 300

# import wxPython
import wx

if __name__ == '__main__':
    __builtins__._ = wx.GetTranslation

# import Python ssl module
import ssl
# import Python socket module
import socket

import datetime

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
# import Transana's Images
import TransanaImages
# import Transana's Database Interface
import DBInterface
# import Transana Dialogs
import Dialogs
# import Transana's Document Object
import Document
# import Transana's Library object
import Library
# import Transana's Episode object
import Episode
# import Transana's Transcript object
import Transcript
# import Transana's Collection object
import Collection
# import Transana's Quote object
import Quote
# import Transana's Clip object
import Clip
# import Transana's Snapshot object
import Snapshot
# import Transana's Note object
import Note


# We create a thread to listen for messages from the Message Server.  However,
# only the Main program thread can interact with a wxPython GUI.  Therefore,
# we need to create a custom event to receive the event from the Message Server
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
            print "PostMessageEvent created with data =", self.data.encode('latin1'), "ChatWindow.EventIDS:", EVT_POST_MESSAGE_ID, EVT_CLOSE_MESSAGE_ID, EVT_MESSAGESERVER_LOST_ID


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
                print '"%s"' % data.encode('latin1')

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

                    if DEBUG:
                        print "ChatWindow.ListenerThread.run() posting a PostMessageEvent of '%s'" % message

                    tmpEvent = PostMessageEvent(message)
                    
                    # Post the Message Event with the message as data
                    wx.PostEvent(self.window, tmpEvent)
                # Unlock the thread when we're done processing all the messages.
                threadLock.release()
            else:

                if DEBUG:
                    print "Blank message received.  This may mean that the socket connection was lost!!\nOr it may not."
                    
                # NOTE:  No GUI from inside the thread.  We use a MessageServerLost Event instead.
                if self.window.reportSocketLoss:
                    try:
                        wx.PostEvent(self.window, MessageServerLostEvent())
                    except wx._core.PyDeadObjectError:
                        pass

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
    def __init__(self, parent, id, title, socketObj):
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
        # Add a ListBox to hold the names of active users.  Allow multiple selection for Private Chat specification.
        self.userList = wx.ListBox(self, -1, choices=[self.userName], style=wx.LB_MULTIPLE)
        boxUser.Add(self.userList, 1, wx.BOTTOM | wx.EXPAND, 3)
        # Create a dictionary that knows the SSL status of each connected user
        self.sslStatus = {}        

        # Add the SSL image
        # Create a Row Sizer
        infoSizer = wx.BoxSizer(wx.HORIZONTAL)
        # If Chat has an SSL connection ...
        if TransanaGlobal.chatIsSSL:
            # ... load the "locked" image
            image = TransanaImages.locked.GetBitmap()
        # If Chat has an un-encrypted connection ...
        else:
            # ... load the "unlocked" image
            image = TransanaImages.unlocked.GetBitmap()

        # Create a BitMap to display on screen
        self.sslImage = wx.StaticBitmap(self, -1, image, (16, 16))
        # Add the image to the Row sizer
        infoSizer.Add(self.sslImage, 0, wx.RIGHT, 4)
        # Make the image clickable
        self.sslImage.Bind(wx.EVT_LEFT_DOWN, self.OnSSLClick)

        # Add a checkbox to enable/disable audio feedback
        self.useSound = wx.CheckBox(self, -1, _("Sound Enabled"))
        # Check the box to start
        self.useSound.SetValue(True)
        # Add the checkbox to the Row sizer
        infoSizer.Add(self.useSound, 0)
        # Add the Row Sizer to the User column Sizer
        boxUser.Add(infoSizer, 0)
        
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
        # Bind the OnKeyUp event with the Text Entry control's wxEVT_KEY_UP event
        self.txtEntry.Bind(wx.EVT_KEY_UP, self.OnKeyUp)

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
        if __name__ == '__main__':
            self.CentreOnScreen()
        else:
            TransanaGlobal.CenterOnPrimary(self)
        
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
        if (TransanaGlobal.menuWindow != None) and (TransanaGlobal.menuWindow.ControlObject != None):
            TransanaGlobal.menuWindow.ControlObject.Register(Chat=self)
            self.ControlObject = TransanaGlobal.menuWindow.ControlObject
        else:
            self.ControlObject = None

        # Inform the message server of user's identity and database connection
        if 'unicode' in wx.PlatformInfo:
            if __name__ == '__main__':
                userName = 'Test_' + sys.platform
                host = '192.168.1.19'
                db = 'ChatTestDB'
                ssl = False
            else:
                userName = self.userName.encode('utf8')
                host = TransanaGlobal.configData.host.encode('utf8')
                db = TransanaGlobal.configData.database.encode('utf8')
                ssl = TransanaGlobal.chatIsSSL
            
            if DEBUG:
                print 'C %s %s %s %s %s ||| ' % (userName, host, db, ssl, VERSION)


            if ('wxGTK' in wx.PlatformInfo) and (host.lower() == 'localhost'):
                host = 'walnut-v.ad.education.wisc.edu'

                print
                print '***************************************************************************'
                print '*                      GTK Message Server Faked!                          *'
                print '***************************************************************************'
            
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

            self.socketObj.send('C %s %s %s %s %s ||| ' % (userName, host, db, ssl, VERSION))

            # Add this user to the SSL Status Dictionary
            if ssl:
                self.sslStatus[userName] = 'TRUE'
            else:
                self.sslStatus[userName] = 'FALSE'
        else:
            self.socketObj.send('C %s %s %s %s %s ||| ' % (self.userName, TransanaGlobal.configData.host, TransanaGlobal.configData.database, TransanaGlobal.chatIsSSL, VERSION))

            # Add this user to the SSL Status Dictionary
            if TransanaGlobal.chatIsSSL:
                self.sslStatus[self.userName] = 'TRUE'
            else:
                self.sslStatus[self.userName] = 'FALSE'

        # Okay, this is really annoying.  Here's the scoop.
        # July 21, 2015.  If I use OS X 10.10.3 (Yosemite), the current OS X release,
        # the Chat and MU Messaging infrastructure doesn't work on the Mac.
        # That's because EVT_POST_MESSAGE works to Post a Message, but for reasons that elude me,
        # it doesn't trigger ChatWindow.OnPostMessage() like it's supposed to.  So the message
        # coming in from the Message Server gets received and written to the Event Queue okay,
        # but it doesn't get "read" and posted in the Chat Window because the ChatWindow's
        # EVT_POST_MESSAGE handler doesn't get called in response to the message being queued
        # correctly.
        #
        # So what I've done is set up a time that's called every half second.  All it does is
        # call wx.YieldIfNeeded().  That seems to be enough!
        #
        # I know I shouldn't HAVE to do this, but at least for now, I do.
        #
##        if 'wxMac' in wx.PlatformInfo:
##            if DEBUG:
##                self.processMessageQueueTime = datetime.datetime.now()
##            self.processMessageQueueTimer = wx.Timer()
##            self.processMessageQueueTimer.Bind(wx.EVT_TIMER, self.OnProcessMessageQueue)
##            self.processMessageQueueTimer.Start(500)

        # Create a Timer to check for Message Server validation.
        # Initialize to unvalidated state
        self.serverValidation = False
        # Create a Timer
        self.validationTimer = wx.Timer()
        # Assign the Timer's event
        self.validationTimer.Bind(wx.EVT_TIMER, self.OnValidationTimer)
        # 10 seconds should be sufficient for the connection to the message server to be established and confirmed
        self.validationTimer.Start(10000)

##    def OnProcessMessageQueue(self, event):
##        
##        if DEBUG:
##            print "ChatWIndow.ProcessMessageQueue():", datetime.datetime.now() - self.processMessageQueueTime
##        
##        wx.YieldIfNeeded()

    def SendMessage(self, message):
        """ Send a message through the chatWindow's socket """
        try:
            # Process Windows Messages, if needed.  (Completes SAVES, in theory!)
            wx.YieldIfNeeded()
            msg = '%s ||| ' % message
            # If we're using Unicode, we need to encode the messages passed to the socket.
            if 'unicode' in wx.PlatformInfo:
                self.socketObj.send(msg.encode('utf8'))
            else:
                self.socketObj.send(msg)
            # This *should* allow messages to be sent totally independently.
            time.sleep(0.05)
            
        except socket.error:
            if DEBUG:
                print "ChatWindow.SendMessage() socket error."

            # NOTE:  No GUI from inside the thread.  We use a MessageServerLost Event instead.
            wx.PostEvent(self.window, MessageServerLostEvent())
            # Close the Chat Window if the connection's been broken.
            self.Close()
        except:
            print sys.exc_info()[0]
            print sys.exc_info()[1]
            import traceback
            traceback.print_exc(file=sys.stdout)
            

    def OnSend(self, event):
        """ Send Message handler """
        # Get the message from the text entry box
        message = self.txtEntry.GetValue()
        # if a User Report is requested ...
        if message.upper() == _('REPORT'):
            # ... get the user list
            userList = self.userList.GetItems()
            # Print the Report Header
            self.memo.AppendText(_('User Report:') + u'\n')
            # For each user ...
            for x in userList:
                # ... determine the user's SSL Status
                if self.sslStatus[x] == 'TRUE':
                    status = _('Secure')
                else:
                    status = _('NOT secure')
                # Add a line to the Report indicating the user name and SSL status
                self.memo.AppendText(unicode('%s\t\t%s\n', 'utf8') % (x, status))
            # Add a blank line to the Report
            self.memo.AppendText('\n')

        # Make sure there IS a message!
        elif message.strip() != "":
            # If there are user selections, we want PRIVATE MESSAGING.  If all or none are selected,
            # everyone gets to see the message.  If only the current user is selected, there are NO
            # RECIPIENTS, so don't sent the message!
            if (len(self.userList.GetSelections()) > 0) and \
               (self.userList.GetSelections() != (0, )):
                # (len(self.userList.GetSelections()) < self.userList.GetCount()) and \
                # Make sure THIS user is NOT selected
                if self.userList.IsSelected(0):
                    self.userList.Deselect(0)
                # If (and only if) not everyone in the list is selected ...
                if len(self.userList.GetSelections()) < self.userList.GetCount() - 1:
                    # Indicate a Private Message by adding on the intended recipients
                    message += ' >|<'
                    # Add the recipients list to the message
                    for index in self.userList.GetSelections():
                        message += ' ' + self.userList.GetString(index)
            # if more than just the current user is selected ...
            if (self.userList.GetSelections() != (0, )):
                # ... send the message.  Indicate that this is a Text Message by prefacing the text with "M".
                self.SendMessage('M %s' % message)
        # Clear the Text Entry control
        self.txtEntry.SetValue('')
        # Set the focus to the Text Entry Control
        self.txtEntry.SetFocus()

    def OnKeyUp(self, event):
        """ Text Entry Control's wxEVT_KEY_UP method """
        # if the message length exceeds 800 characters, and the user presses SPACE ...
        if (len(self.txtEntry.GetValue()) > 800) and (event.GetKeyCode() == wx.WXK_SPACE):
            # ... just send the damn message
            self.OnSend(event)

    def OnClear(self, event):
        """ Clear button handler """
        # Clear the previous chat text
        self.memo.Clear()
        # Set the focus to the Text Entry control
        self.txtEntry.SetFocus()

    def UpdateSSLStatus(self):
        """ Check and Update the SSL Connection Status of the Chat Window """
        # See if there are unsecured connections ...
        if "FALSE" in self.sslStatus.values():
            # ... load the "unlocked" image
            image = TransanaImages.unlocked.GetBitmap()
        # If Chat has only encrypted connections ...
        else:
            # ... load the "locked" image
            image = TransanaImages.locked.GetBitmap()

        # Update the BitMap on screen
        self.sslImage.SetBitmap(image)

        # Notify the Control Object to update other SSL indicator(s)
        if self.ControlObject != None:
            self.ControlObject.UpdateSSLStatus(not ("FALSE" in self.sslStatus.values()))

    def OnSSLClick(self, event):
        """ Handle click on the SSL indicator image """
        # Determine whether SSL is in use with the Database connection
        dbIsSSL = TransanaGlobal.configData.ssl
        # Determine whether SSL is FULLY in use with the Message Server connection
        chatIsSSL = not ("FALSE" in self.sslStatus.values())
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
        
    def OnPostMessage(self, event):
        """ Post Message handler """

        def ConvertMessageToNodeList(message):
            """ Take a message from the Transana Message Server and convert it to a NodeList for use with the DB Tree
                add_Node method """
            nodelist = ()
            for m in message.split(' >|< '):
                nodelist += (m,)
            return nodelist

        if DEBUG:
            print "ChatWindow.OnPostMessage():", event.data
                
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
                
            # Connection Message
            elif messageHeader == 'C':
                # This signals that a UserName should be added to the list of current users.
                # Drop the Message Prefix.
                st = event.data[2:]
                # Break the message into its components
                tmpData = st.split(' ')
                # See how many components were included.  If five (when OTHER users connect using 2.61 or later) ...
                if len(tmpData) == 5:
                    # ... then SSL is the 4th (index 3) of them.
                    tmpSSL = tmpData[3].upper()
                # If there are only 2 elements (OTHER users when WE first connect)
                elif len(tmpData) == 2:
                    tmpSSL = tmpData[1].upper()
                # If not five or 2 (OTHER users using 2.60 or earlier have 4, for example) ...
                else:
                    # ... then SSL is missing, so is FALSE!!
                    tmpSSL = 'FALSE'
                
                # Only add a user name if it's not redundant.  (The message server should prevent this
                # from occurring.)
                if (tmpData[0] != self.userName) and not (tmpData[0] in self.userList.GetItems()):
                    # Add the user to the User List
                    self.userList.Append(tmpData[0])
                    # Add the user's SSL Status to the dictionary
                    self.sslStatus[tmpData[0]] = tmpSSL
                    # Update the SSL indicators
                    self.UpdateSSLStatus()
                    
            # Rename Message ?
            elif messageHeader == 'R':
                # If a user name is duplicated, the Message Server renames it to prevent confusion.
                # This code indicates that THIS USER's account has been renamed by the Message Server.
                # Remove the "old" user name
                self.userList.Delete(0)
                # Start Exception Handling
                try:
                    # Try to remove the user from the SSL Status dictionary
                    del(self.sslStatus[self.userName])
                # If a KeyError is generated ...
                except KeyError:
                    # ... we can ignore it.
                    pass
                # Drop the Message Prefix
                st = event.data[2:]
                # Break the message into its parts
                tmpData = st.split(' ')
                # The first part will be the user name
                self.userName = tmpData[0]
                # The second part will be the user's SSL status
                SSL = tmpData[1]
                # Add the user to the User List
                self.userList.Append(self.userName)
                # Add this user's SSL Status to the dictionary
                self.sslStatus[tmpData[0]] = tmpData[1]
                # Update SSL Status indicators
                self.UpdateSSLStatus()
                
            # Import Message
            elif messageHeader == 'I':
                # Another user has imported a database.  We need to refresh the whole Database Tree!
                # See if a Control Object has been defined.
                if self.ControlObject != None:
                    # See if there's a Notes Browser open
                    if self.ControlObject.NotesBrowserWindow != None:
                        # If so, close it.
                        self.ControlObject.NotesBrowserWindow.Close()
                    # Update the Data Window via the Control Object
                    self.ControlObject.UpdateDataWindow()

            # Server Validation
            elif messageHeader == 'V':
                # Indicate that the server has been validated.  The Validation Timer processes this later.
                self.serverValidation = True
                
            # Disconnect Message ?
            elif messageHeader == 'D':
                # Remove the Message Prefix.
                st = event.data[2:]
                try:
                    # The remainder of the message is the username to be removed.  Delete it from the User List.
                    self.userList.Delete(self.userList.FindString(st))
                    # Remove the user from the SSL Status dictionary
                    del(self.sslStatus[st])
                    # Update the SSL Status indicators
                    self.UpdateSSLStatus()
                except:
                    pass
                
            else:
                # The remaining messages should not be processed if this user was the message sender
                if self.userName != messageSender:
                    # We can't have the tree selection changing because of the activity of other users.  That creates all kinds of
                    # problems if we're in the middle of editing something.  So let's note the current selection
                    currentSelection = self.ControlObject.DataWindow.DBTab.tree.GetSelections()
                    # The Control Object MUST be defined (and always will be)
                    if self.ControlObject != None:
                        # Add Library Message
                        if messageHeader == 'AS':
                            tempLibrary = Library.Library(message)
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('LibraryNode', (_('Libraries'), message), tempLibrary.number, None, expandNode=False, avoidRecursiveYields = True)
                            
                        # Add Episode Message
                        elif messageHeader == 'AE':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            tempEpisode = Episode.Episode(series=nodelist[0], episode=nodelist[1])
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('EpisodeNode', (_('Libraries'),) + nodelist, tempEpisode.number, tempEpisode.series_num, expandNode=False, avoidRecursiveYields = True)

                        # Add Transcript Message
                        elif messageHeader == 'AT':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            tempEpisode = Episode.Episode(series=nodelist[0], episode=nodelist[1])
                            # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                            tempTranscript = Transcript.Transcript(nodelist[-1], ep=tempEpisode.number, skipText=True)
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('TranscriptNode', (_('Libraries'),) + nodelist, tempTranscript.number, tempEpisode.number, expandNode=False, avoidRecursiveYields = True)

                        # Add Document Message
                        elif messageHeader == 'AD':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            tempDocument = Document.Document(libraryID=nodelist[0], documentID=nodelist[1])
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('DocumentNode', (_('Libraries'),) + nodelist, tempDocument.number, tempDocument.library_num, expandNode=False, avoidRecursiveYields = True)

                        # Add Collection Message
                        elif messageHeader == 'AC':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            parentNum = 0
                            for coll in nodelist:
                                tempCollection = Collection.Collection(coll, parentNum)
                                parentNum = tempCollection.number
                            # avoidRecursiveYields added to try to prevent a problem on the Mac when converting Searches
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('CollectionNode', (_('Collections'),) + nodelist, tempCollection.number, tempCollection.parent, expandNode=False, avoidRecursiveYields=True)

                        # Add Quote Message
                        elif messageHeader == 'AQ':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            parentNum = 0
                            for coll in nodelist[:-1]:
                                tempCollection = Collection.Collection(coll, parentNum)
                                parentNum = tempCollection.number
                            # Get a temporary copy of the Quote.  We don't need the quote's text, which speeds this up.
                            tempQuote = Quote.Quote(quoteID=nodelist[-1], collectionID=tempCollection.id, collectionParent=tempCollection.parent, skipText=True)
                            # avoidRecursiveYields added to try to prevent a problem on the Mac when converting Searches
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('QuoteNode', (_('Collections'),) + nodelist, tempQuote.number, tempCollection.number, sortOrder=tempQuote.sort_order, expandNode=False, avoidRecursiveYields=True)
                            # If the Quote's Document is open, it needs to be updated with the Quote information!
                            self.ControlObject.AddQuoteToOpenDocument(tempQuote)
                            # If we are moving a Quote, the quote's Notes need to travel with the Quote.  The first step is to
                            # get a list of those Notes.
                            noteList = DBInterface.list_of_notes(Quote=tempQuote.number)
                            # If there are Quote Notes, we need to make sure they travel with the Quote
                            if noteList != []:
                                insertNode = self.ControlObject.DataWindow.DBTab.tree.select_Node((_('Collections'),) + nodelist, 'QuoteNode', ensureVisible=False)
                                # We accomplish this using the TreeCtrl's "add_note_nodes" method
                                self.ControlObject.DataWindow.DBTab.tree.add_note_nodes(noteList, insertNode, Quote=tempQuote.number)
                                self.ControlObject.DataWindow.DBTab.tree.Refresh()

                        # Add Clip Message
                        elif messageHeader == 'ACl':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            parentNum = 0
                            for coll in nodelist[:-1]:
                                tempCollection = Collection.Collection(coll, parentNum)
                                parentNum = tempCollection.number
                            # Get a temporary copy of the Clip.  We don't need the clip's transcript, which speeds this up.
                            tempClip = Clip.Clip(nodelist[-1], tempCollection.id, tempCollection.parent, skipText=True)
                            # avoidRecursiveYields added to try to prevent a problem on the Mac when converting Searches
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('ClipNode', (_('Collections'),) + nodelist, tempClip.number, tempCollection.number, sortOrder=tempClip.sort_order, expandNode=False, avoidRecursiveYields=True)
                            # If we are moving a Clip, the clip's Notes need to travel with the Clip.  The first step is to
                            # get a list of those Notes.
                            noteList = DBInterface.list_of_notes(Clip=tempClip.number)
                            # If there are Clip Notes, we need to make sure they travel with the Clip
                            if noteList != []:
                                insertNode = self.ControlObject.DataWindow.DBTab.tree.select_Node((_('Collections'),) + nodelist, 'ClipNode', ensureVisible=False)
                                # We accomplish this using the TreeCtrl's "add_note_nodes" method
                                self.ControlObject.DataWindow.DBTab.tree.add_note_nodes(noteList, insertNode, Clip=tempClip.number)
                                self.ControlObject.DataWindow.DBTab.tree.Refresh()

                        # Add Clip in Sort Order Message
                        elif messageHeader == 'AClSO':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            parentNum = 0
                            for coll in nodelist[:-2]:
                                tempCollection = Collection.Collection(coll, parentNum)
                                parentNum = tempCollection.number
                            # We need the NODE for the Clip we should place the new clip in front of.  Let's get that here.
                            insertNode = self.ControlObject.DataWindow.DBTab.tree.select_Node((_('Collections'),) + nodelist[:-1], 'ClipNode', ensureVisible=False)
                            # Get a temporary copy of the Clip.  We don't need the clip's transcript, which speeds this up.
                            tempClip = Clip.Clip(nodelist[-1], tempCollection.id, tempCollection.parent, skipText=True)
                            # Add new node, leaving the insertNode out of the nodeList.
                            # avoidRecursiveYields added to try to prevent a problem on the Mac when converting Searches
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('ClipNode', (_('Collections'),) + nodelist[:-2] + (nodelist[-1],), tempClip.number, tempCollection.number, sortOrder=tempClip.sort_order, expandNode=False, insertPos=insertNode, avoidRecursiveYields=True)
                            # If we are moving a Clip, the clip's Notes need to travel with the Clip.  The first step is to
                            # get a list of those Notes.
                            noteList = DBInterface.list_of_notes(Clip=tempClip.number)
                            # If there are Clip Notes, we need to make sure they travel with the Clip
                            if noteList != []:
                                insertNode = self.ControlObject.DataWindow.DBTab.tree.select_Node((_('Collections'),) + nodelist[:-2] + (nodelist[-1],), 'ClipNode', ensureVisible=False)
                                # We accomplish this using the TreeCtrl's "add_note_nodes" method
                                self.ControlObject.DataWindow.DBTab.tree.add_note_nodes(noteList, insertNode, Clip=tempClip.number)
                                self.ControlObject.DataWindow.DBTab.tree.Refresh()

                        # Order Collection Message
                        elif messageHeader == 'OC':
                            # Convert the message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            # Get the Collection's Tree Node
                            node = self.ControlObject.DataWindow.DBTab.tree.select_Node((_('Collections'),) + nodelist[1:], 'CollectionNode')
                            # Update the Sort Information for that tree node
                            self.ControlObject.DataWindow.DBTab.tree.UpdateCollectionSortOrder(node, sendMessage=False)

                        # Add Snapshot Message
                        elif messageHeader == 'ASnap':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            # Initialize the Parent Number variable
                            parentNum = 0
                            # Get the appropriate Collection by interating through the node list ...
                            for coll in nodelist[:-1]:
                                # ... get the Collection for each node ...
                                tempCollection = Collection.Collection(coll, parentNum)
                                # ... and the parent number
                                parentNum = tempCollection.number
                            # Get a temporary copy of the Snapshot.
                            tempSnapshot = Snapshot.Snapshot(nodelist[-1], parentNum)
                            # avoidRecursiveYields added to try to prevent a problem on the Mac when converting Searches
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('SnapshotNode', (_('Collections'),) + nodelist, tempSnapshot.number, tempCollection.number, sortOrder=tempSnapshot.sort_order, expandNode=False, avoidRecursiveYields=True)
                            # If we are moving a Snapshot, the snapshot's Notes need to travel with the Snapshot.  The first step is to
                            # get a list of those Notes.
                            noteList = DBInterface.list_of_notes(Snapshot=tempSnapshot.number)
                            # If there are Snapshot Notes, we need to make sure they travel with the Snapshot
                            if noteList != []:
                                insertNode = self.ControlObject.DataWindow.DBTab.tree.select_Node((_('Collections'),) + nodelist, 'SnapshotNode', ensureVisible=False)
                                # We accomplish this using the TreeCtrl's "add_note_nodes" method
                                self.ControlObject.DataWindow.DBTab.tree.add_note_nodes(noteList, insertNode, Snapshot=tempSnapshot.number)
                                self.ControlObject.DataWindow.DBTab.tree.Refresh()

                        # Add Note Message
                        elif messageHeader in ['ASN', 'ADN', 'AEN', 'ATN', 'ACN', 'AQN', 'AClN', 'ASnN']:
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
                                # If the first entry in the node list is the "Library" Root Node ...
                                if (objectType == None) and (node == 'Libraries'):
                                    # ... then we're climbing up the Library branch, and are at a Library record.
                                    objectType = 'Library'
                                # If we're already at a Library record and we have an Episode or Transcript Note ...
                                elif (objectType == 'Library') and (messageHeader in ['ASN', 'AEN', 'ATN']):
                                    # ... then we're moving on to an Episode next
                                    objectType = 'Episode'
                                    # We might have a Library Note, at least if we stop here!
                                    nodeType = 'LibraryNoteNode'
                                    # Let's load the Library record ...
                                    tempObj = Library.Library(node)
                                    # .. and note that the parent of the NEXT object is this Library's number!
                                    parentNum = tempObj.number
                                # If we're already at a Library record and we have a Document Note ...
                                elif (objectType == 'Library') and (messageHeader in ['ADN']):
                                    # ... then we're moving on to an Document next
                                    objectType = 'Document'
                                    # We might have a Library Note, at least if we stop here!
                                    nodeType = 'LibraryNoteNode'
                                    # Let's load the Library record ...
                                    tempObj = Library.Library(node)
                                    # .. and note that the parent of the NEXT object is this Library's number!
                                    parentNum = tempObj.number
                                # If we're already at a Document record ...
                                elif (objectType == 'Document'):
                                    # ... then we're looking at a Document
                                    objectType = 'Document Note'
                                    # we have a Document Note if we stop here!
                                    nodeType = 'DocumentNoteNode'
                                    # Let's load the Document Record
                                    tempObj = Document.Document(libraryID=tempObj.id, documentID=node)
                                    # .. and note that the parent of the NEXT object is this Document's number!
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
                                    # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                                    tempObj = Transcript.Transcript(node, ep=parentNum, skipText=True)
                                    # ... and note that the parent of the Transcript Note is this Trasncript.
                                    parentNum = tempObj.number
                                # If our node is the Collections Root Node ...
                                elif (objectType == None) and (node == 'Collections'):
                                    # ... then the first level of object we're looking at is a Collection.
                                    objectType = 'Collections'
                                # if we're looking at a Collection and either we don't have a Quote / Clip / Snapshot Note
                                # or we're not at the end of the list yet...
                                elif (objectType == 'Collections') and (not (messageHeader in ['AQN', 'AClN', 'ASnN']) or (nodeCount < len(nodelist) - 1)):
                                    # ... then we're still looking at a Collection
                                    objectType = 'Collections'
                                    # ... and if we stop here, we've got a Collection Note
                                    nodeType = 'CollectionNoteNode'
                                    # Load the Collection
                                    tempObj = Collection.Collection(node, parentNum)
                                    # ... and note that the collection is the parent of the NEXT object.
                                    parentNum = tempObj.number
                                # if we're looking at a Collection and we have a Quote Note and we're at the end of the list ...
                                elif (objectType == 'Collections') and (messageHeader == 'AQN') and (nodeCount == len(nodelist) - 1):
                                    # ... then we're looking at a Quote
                                    objectType = 'Quote'
                                    # ... and we're dealing with a Quote Note
                                    nodeType = 'QuoteNoteNode'
                                    # Get a temporary copy of the Quote.  We don't need the Quote's transcript, which speeds this up.
                                    tempObj = Quote.Quote(quoteID=node, collectionID=tempObj.id, collectionParent=tempObj.parent, skipText=True)
                                    # ... and note its number as the parent number of the Note
                                    parentNum = tempObj.number
                                # if we're looking at a Collection and we have a Clip Note and we're at the end of the list ...
                                elif (objectType == 'Collections') and (messageHeader == 'AClN') and (nodeCount == len(nodelist) - 1):
                                    # ... then we're looking at a Clip
                                    objectType = 'Clip'
                                    # ... and we're dealing with a Clip Note
                                    nodeType = 'ClipNoteNode'
                                    # Get a temporary copy of the Clip.  We don't need the clip's transcript, which speeds this up.
                                    tempObj = Clip.Clip(node, tempObj.id, tempObj.parent, skipText=True)
                                    # ... and note its number as the parent number of the Note
                                    parentNum = tempObj.number
                                # if we're looking at a Collection and we have a Snapshot Note and we're at the end of the list ...
                                elif (objectType == 'Collections') and (messageHeader == 'ASnN') and (nodeCount == len(nodelist) - 1):
                                    # ... then we're looking at a Snapshot
                                    objectType = 'Snapshot'
                                    # ... and we're dealing with a Snapshot Note
                                    nodeType = 'SnapshotNoteNode'
                                    # Get a temporary copy of the Snapshot.
                                    tempObj = Snapshot.Snapshot(node, tempObj.number)
                                    # ... and note its number as the parent number of the Note
                                    parentNum = tempObj.number
                            # Initialize the Temporary Note object
                            tempNote = None
                            # Load the Note, which we do a bit differently based on what kind of parent object we have.
                            if nodeType == 'LibraryNoteNode':
                                tempNote = Note.Note(nodelist[-1], Library=tempObj.number)
                            elif nodeType == 'DocumentNoteNode':
                                tempNote = Note.Note(nodelist[-1], Document=tempObj.number)
                            elif nodeType == 'EpisodeNoteNode':
                                tempNote = Note.Note(nodelist[-1], Episode=tempObj.number)
                            elif nodeType == 'TranscriptNoteNode':
                                tempNote = Note.Note(nodelist[-1], Transcript=tempObj.number)
                            elif nodeType == 'CollectionNoteNode':
                                tempNote = Note.Note(nodelist[-1], Collection=tempObj.number)
                            elif nodeType == 'QuoteNoteNode':
                                tempNote = Note.Note(nodelist[-1], Quote=tempObj.number)
                            elif nodeType == 'ClipNoteNode':
                                tempNote = Note.Note(nodelist[-1], Clip=tempObj.number)
                            elif nodeType == 'SnapshotNoteNode':
                                tempNote = Note.Note(nodelist[-1], Snapshot=tempObj.number)
                            # Add the Note to the Database Tree
                            self.ControlObject.DataWindow.DBTab.tree.add_Node(nodeType, nodelist, tempNote.number, tempObj.number, expandNode=False, avoidRecursiveYields = True)
                            # If the Notes Browser is open ...
                            if self.ControlObject.NotesBrowserWindow != None:
                                # ... add the Note to the Notes Browser
                                self.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('A', tempNote)

                        # Add Keyword Group Message
                        elif messageHeader == 'AKG':
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('KeywordGroupNode', (_('Keywords'),) + (message, ), 0, 0, expandNode=False, avoidRecursiveYields = True)
                            # Once we've added the Keyword Group, we need to update the Keyword Groups Data Structure
                            self.ControlObject.DataWindow.DBTab.tree.updateKWGroupsData()

                        # Add Keyword Message
                        elif messageHeader == 'AK':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('KeywordNode', (_('Keywords'),) + nodelist, 0, nodelist[0], expandNode=False, avoidRecursiveYields = True)

                        # Add Keyword Example Message
                        elif messageHeader == 'AKE':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            # Get a temporary copy of the Clip.  We don't need the clip's transcript, which speeds this up.
                            tempClip = Clip.Clip(int(nodelist[0]), skipText=True)
                            self.ControlObject.DataWindow.DBTab.tree.add_Node('KeywordExampleNode', (_('Keywords'),) + nodelist[1:], tempClip.number, tempClip.collection_num, expandNode=False, avoidRecursiveYields = True)

                        # Rename a Node
                        elif messageHeader == 'RN':
                            # Convert the Message to a Node List
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
                            # Encode the first-level node name for the LOCAL language
                            nodelist = (nodelist[0], tmpRootNode) + nodelist[2:]
                            
                            if DEBUG:
                                tmpstr = "Calling rename_Node(%s, %s, %s) %s %s" % (nodelist[1:-1], nodelist[0], nodelist[-1], \
                                         type(nodelist[0]), type(nodelist[-1]))
                                print tmpstr.encode('latin1')
                                print

                            # Rename the tree node
                            self.ControlObject.DataWindow.DBTab.tree.rename_Node(nodelist[1:-1], nodelist[0], nodelist[-1])

                            # Let's see if the renamed Document, Quote, or (Episode or Clip) Transcript is currently OPEN.
                            # (Open Episodes don't need to be treated the same way because of the looser relationship to
                            #  the transcript!)
                            docType = None
                            if nodelist[0] == "DocumentNode":
                                docType = Document.Document
                                tmpObj = Document.Document(libraryID = nodelist[-3], documentID = nodelist[-1])
                            elif nodelist[0] == 'TranscriptNode':
                                docType = Transcript.Transcript
                                tmpEpisode = Episode.Episode(series=nodelist[-4], episode=nodelist[-3])
                                # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                                tmpObj = Transcript.Transcript(nodelist[-1], ep=tmpEpisode.number, skipText=True)
                            elif nodelist[0] == 'QuoteNode':
                                docType = Quote.Quote
                                parentNum = 0
                                for coll in nodelist[2:-2]:
                                    tmpCollection = Collection.Collection(coll, parentNum)
                                    parentNum = tmpCollection.number
                                # Get a temporary copy of the Quote.  We don't need the quote's text, which speeds this up.
                                tmpObj = Quote.Quote(quoteID=nodelist[-1], collectionID=tmpCollection.id, collectionParent=tmpCollection.parent, skipText=True)
                            elif nodelist[0] == 'ClipNode':
                                docType = Transcript.Transcript
                                parentNum = 0
                                for coll in nodelist[2:-2]:
                                    tmpCollection = Collection.Collection(coll, parentNum)
                                    parentNum = tmpCollection.number
                                # Get a temporary copy of the Clip.  We don't need the clip's transcript, which speeds this up.
                                tmpClip = Clip.Clip(nodelist[-1], tmpCollection.id, tmpCollection.parent)
                                tmpObj = tmpClip.transcripts[0]
                            if (docType != None) and \
                               self.ControlObject.GetOpenDocumentObject(docType, tmpObj.number) != None:
                                # Note the type and number of the currently opened object in the Document Window
                                currObjType = type(self.ControlObject.GetCurrentDocumentObject())
                                currObjNum = self.ControlObject.GetCurrentDocumentObject().number
                                # Close the Open Document Window for the changed object
                                self.ControlObject.CloseOpenTranscriptWindowObject(docType, tmpObj.number)
                                # Then open a new Document Tab for the changed object
                                if nodelist[0] == 'DocumentNode':
                                    self.ControlObject.LoadDocument(nodelist[-3], tmpObj.id, tmpObj.number)
                                elif nodelist[0] == 'TranscriptNode':
                                    self.ControlObject.LoadTranscript(nodelist[-4], tmpEpisode.id, tmpObj.id)
                                elif nodelist[0] == 'QuoteNode':
                                    self.ControlObject.LoadQuote(tmpObj.number)
                                elif nodelist[0] == 'ClipNode':
                                    self.ControlObject.LoadClipByNumber(tmpClip.number)
                                # Finally, restore the interface to the object that was showing when we started.
                                self.ControlObject.SelectOpenDocumentTab(currObjType, currObjNum)
                                
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
                                            
                                # We need to update open Snapshots that contain this keyword.
                                # Start with a list of all the open Snapshot Windows.
                                openSnapshotWindows = self.ControlObject.GetOpenSnapshotWindows()
                                # Interate through the Snapshot Windows
                                for win in openSnapshotWindows:
                                    # For ANY non-editable Snapshot ...
                                    if (not win.editTool.IsToggled()):
                                        # ... update the Snapshot Window
                                        win.FileClear(event)
                                        win.FileRestore(event)

                                        win.OnEnterWindow(event)

                            # If we're renaming a Note ...
                            elif nodelist[0] in ['LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode',
                                                 'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', 'SnapshotNoteNode']:
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
                                        # If the first entry in the node list is the "Library" Root Node ...
                                        if (objectType == None) and (node == unicode(_('Libraries'), 'utf8')):
                                            # ... then we're climbing up the Library branch, and are at a Library record.
                                            objectType = 'Library'
                                        # If we're already at a Library record and we're NOT looking for a Document Note ...
                                        elif (objectType == 'Library'):
                                            if (nodeType == 'DocumentNoteNode'):
                                                # ... then we're looking at a Document
                                                objectType = 'Document'
                                            else:
                                                # ... then we're moving on to an Episode next
                                                objectType = 'Episode'
                                            # Let's load the Library record ...
                                            tempObj = Library.Library(node)
                                            # .. and note that the parent of the NEXT object is this Library's number!
                                            parentNum = tempObj.number
                                        # if we're looking at a Library and we have a Document Note and we're at the end of the list ...
                                        elif (objectType == 'Document'):
                                            # ... then we're looking at a Document
                                            objectType = 'Document Note'
                                            # Get a temporary copy of the Document.  We don't need the Document's transcript, which speeds this up.
                                            tempObj = Document.Document(node, libraryID=tempObj.id, documentID=node, skipText=True)
                                            # ... and note its number as the parent number of the Note
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
                                            # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                                            tempObj = Transcript.Transcript(node, ep=parentNum, skipText=True)
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
                                        # if we're looking at a Collection and we have a Quote Note and we're at the end of the list ...
                                        elif (objectType == 'Collections') and (nodeType == 'QuoteNoteNode') and (nodeCount == len(nodelist) - 2):
                                            # ... then we're looking at a Quote
                                            objectType = 'Quote'
                                            # Get a temporary copy of the Quote.  We don't need the Quote's transcript, which speeds this up.
                                            tempObj = Quote.Quote(node, quoteID=node, collectionID=tempObj.id, collectionParent=tempObj.parent, skipText=True)
                                            # ... and note its number as the parent number of the Note
                                            parentNum = tempObj.number
                                        # if we're looking at a Collection and we have a Clip Note and we're at the end of the list ...
                                        elif (objectType == 'Collections') and (nodeType == 'ClipNoteNode') and (nodeCount == len(nodelist) - 2):
                                            # ... then we're looking at a Clip
                                            objectType = 'Clip'
                                            # Get a temporary copy of the Clip.  We don't need the clip's transcript, which speeds this up.
                                            tempObj = Clip.Clip(node, tempObj.id, tempObj.parent, skipText=True)
                                            # ... and note its number as the parent number of the Note
                                            parentNum = tempObj.number
                                        # if we're looking at a Collection and we have a Snapshot Note and we're at the end of the list ...
                                        elif (objectType == 'Collections') and (nodeType == 'SnapshotNoteNode') and (nodeCount == len(nodelist) - 2):
                                            # ... then we're looking at a Snapshot
                                            objectType = 'Snapshot'
                                            # Get a temporary copy of the Snapshot.
                                            tempObj = Snapshot.Snapshot(node, tempObj.number)
                                            # ... and note its number as the parent number of the Note
                                            parentNum = tempObj.number
                                    # Initialize the Temporary Note object
                                    tempNote = None
                                    # Load the Note, which we do a bit differently based on what kind of parent object we have.
                                    if nodeType == 'LibraryNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Library=tempObj.number)
                                    elif nodeType == 'DocumentNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Document=tempObj.number)
                                    elif nodeType == 'EpisodeNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Episode=tempObj.number)
                                    elif nodeType == 'TranscriptNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Transcript=tempObj.number)
                                    elif nodeType == 'CollectionNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Collection=tempObj.number)
                                    elif nodeType == 'QuoteNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Quote=tempObj.number)
                                    elif nodeType == 'ClipNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Clip=tempObj.number)
                                    elif nodeType == 'SnapshotNoteNode':
                                        tempNote = Note.Note(nodelist[-1], Snapshot=tempObj.number)
                                    # Rename the Note in the Database Tree
                                    self.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('R', tempNote, oldName=nodelist[-2])

                        # Move Collection Node
                        elif messageHeader == 'MCN':
                            # Convert the Message to a Node List
                            nodelist = ConvertMessageToNodeList(message)
                            # The first element in the node list needs translation.  Check it's type.
                            if type(_(nodelist[0])).__name__ == 'str':
                                # If string, translate it and convert it to unicode
                                nodelist = (unicode(_(nodelist[0]), 'utf8'),) + nodelist[1:]
                            # If not string, it's already unicode!
                            else:
                                # ... in which case, we just translate it.
                                nodelist = (_(nodelist[0]),) + nodelist[1:]
                            # Get a pointer to the Tree Control
                            tree = self.ControlObject.DataWindow.DBTab.tree
                            # Get the Collection node that has been moved
                            tmpNode = tree.select_Node(nodelist, 'CollectionNode', ensureVisible=False)
                            # Get the data underlying the tree node.
                            tmpPyData = tree.GetPyData(tmpNode)
                            # Load the collection underlying the node that has been moved
                            tmpCollection = Collection.Collection(tmpPyData.recNum)
                            # Now that we have the collection, we can build the node data for where it should be!
                            destNodeList = (_('Collections'),) + tmpCollection.GetNodeData()[:-1]
                            # Move the local copy of the node without sending MU Messaging
                            tree.copy_Node('CollectionNode', nodelist, destNodeList, True, sendMessage=False)

                        # Delete Node
                        elif messageHeader == 'DN':
                            # Extract the node list from the message
                            nodelist = ConvertMessageToNodeList(message)

                            # Check the TYPE of the translated second element.
                            if type(_(nodelist[1])).__name__ == 'str':
                                # If string, translate it and convert it to unicode
                                nodelist = (nodelist[0],) + (unicode(_(nodelist[1]), 'utf8'),) + nodelist[2:]
                            # If not string, it's unicode!
                            else:
                                # ... in which case, we just translate it.
                                nodelist = (nodelist[0],) + (_(nodelist[1]),) + nodelist[2:]
                            # Keyword Examples need a bit of extra processing.  If we have a keyword example ...
                            if nodelist[0] == 'KeywordExampleNode':
                                # ... pull the clip number off the end of the node list ...
                                exampleClipNum = int(nodelist[-1])
                                # ... remove the clip number from the node list ...
                                nodelist = nodelist[:-1]
                                # ... and call delete_Node, passing the clip number.  We don't want messages sent further.
                                self.ControlObject.DataWindow.DBTab.tree.delete_Node(nodelist[1:], nodelist[0], exampleClipNum = exampleClipNum, sendMessage=False)
                            # If we are removing any other kind of Node ...
                            else:

                                # ... delete the node without passing further messages
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

                                # We need to update open Snapshots that contain this keyword.
                                # Start with a list of all the open Snapshot Windows.
                                openSnapshotWindows = self.ControlObject.GetOpenSnapshotWindows()
                                # Interate through the Snapshot Windows
                                for win in openSnapshotWindows:
                                    # For ANY non-editable Snapshot ...
                                    if (not win.editTool.IsToggled()):
                                        # ... update the Snapshot Window
                                        win.FileClear(event)
                                        win.FileRestore(event)

                                        win.OnEnterWindow(event)

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

                                # We need to update open Snapshots that contain this keyword.
                                # Start with a list of all the open Snapshot Windows.
                                openSnapshotWindows = self.ControlObject.GetOpenSnapshotWindows()
                                # Interate through the Snapshot Windows
                                for win in openSnapshotWindows:
                                    # For ANY non-editable Snapshot ...
                                    if (not win.editTool.IsToggled()):
                                        # ... update the Snapshot Window
                                        win.FileClear(event)
                                        win.FileRestore(event)

                                        win.OnEnterWindow(event)

                            # If we're deleting a Note Node ...
                            elif nodelist[0] in ['LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'CollectionNoteNode',
                                                 'ClipNoteNode', 'SnapshotNoteNode', 'DocumentNoteNode', 'QuoteNoteNode']:
                                # ... and the Notes Browser is open, we need to delete the Note from there too.
                                if self.ControlObject.NotesBrowserWindow != None:
                                    # Determine the Note Browser's root node based on the type of Note we're deleting
                                    if nodelist[0] == 'LibraryNoteNode':
                                        nodeType = 'Library'
                                    elif nodelist[0] == 'EpisodeNoteNode':
                                        nodeType = 'Episode'
                                    elif nodelist[0] == 'TranscriptNoteNode':
                                        nodeType = 'Transcript'
                                    elif nodelist[0] == 'CollectionNoteNode':
                                        nodeType = 'Collection'
                                    elif nodelist[0] == 'ClipNoteNode':
                                        nodeType = 'Clip'
                                    elif nodelist[0] == 'SnapshotNoteNode':
                                        nodeType = 'Snapshot'
                                    elif nodelist[0] == 'DocumentNoteNode':
                                        nodeType = 'Document'
                                    elif nodelist[0] == 'QuoteNoteNode':
                                        nodeType = 'Quote'
                                    else:
                                        nodeType = None
                                    # Signal the Notes Browser to delete the Note.  Shorten the node list by 1 element
                                    # so it does not conflict with DatabaseTreeTab.py calls.
                                    if nodeType != None:
                                        self.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('D', (nodeType, nodelist[1:]))
                            # Otherwise, if a Library, Document, Episode, Transcript, Collection, Quote, Clip, or Snapshot node is deleted ...
                            elif nodelist[0] in ['LibraryNode', 'DocumentNode', 'EpisodeNode', 'TranscriptNode', 'CollectionNode', 'QuoteNode', 'ClipNode', 'SnapshotNode']:
                                # ... and if the Notes Browser is open, ...
                                if self.ControlObject.NotesBrowserWindow != None:
                                    # ... we need to CHECK to see if any notes were deleted.
                                    self.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
                                    
                        # If a Quote is being deleted on another computer, we need to
                        # Delete Quote Position from Open Document
                        elif messageHeader == 'DQPOD':
                            # Parse the message at the space into object type and object number
                            msgData = message.split(' ')
                            # ... we need to drop that Quote's Position from the source Document, if it's open.
                            self.ControlObject.RemoveQuoteFromOpenDocument(int(msgData[0]), int(msgData[1]))

                        # Update Keyword List
                        elif messageHeader == 'UKL':
                            # Parse the message at the space into object type and object number
                            msgData = message.split(' ')
                            # See if the currently loaded object matches the object described in the message.
                            if ((isinstance(self.ControlObject.currentObj, Episode.Episode) and \
                                 (msgData[0] == 'Episode')) or \
                                (isinstance(self.ControlObject.currentObj, Clip.Clip) and \
                                 (msgData[0] == 'Clip')) or \
                                (isinstance(self.ControlObject.currentObj, Document.Document) and \
                                 (msgData[0] == 'Document')) or \
                                (isinstance(self.ControlObject.currentObj, Quote.Quote) and \
                                 (msgData[0] == 'Quote'))) and \
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
                            # if Object Type is Document ...
                            elif msgData[0] == 'Document':
                                # Get the source document's object if it is open somewhere in this copy of Transana
                                openSourceDocument = self.ControlObject.GetOpenDocumentObject(Document.Document, int(msgData[2]))
                                # If it is open somewhere ...
                                if openSourceDocument != None:
                                    # ... then reload it to update its QuotePositions, which may have changed due to a Quote Merge.
                                    openSourceDocument.db_load_by_num(openSourceDocument.number)
                                
                                # See if the currently loaded document matches the document number sent from the Message Server
                                if isinstance(self.ControlObject.currentObj, Document.Document) and \
                                   self.ControlObject.currentObj.number == int(msgData[2]):
                                    # ... we need to update the keyword visualization no matter what.
                                    self.ControlObject.UpdateKeywordVisualization()
                            # if Object Type is Episode ...
                            elif msgData[0] == 'Episode':
                                # See if the currently loaded episode matches the episode number sent from the Message Server
                                if isinstance(self.ControlObject.currentObj, Episode.Episode) and \
                                   self.ControlObject.currentObj.number == int(msgData[1]):
                                    # ... we need to update the keyword visualization no matter what.
                                    self.ControlObject.UpdateKeywordVisualization()
                            # if Object Type is Quote ...
                            elif msgData[0] == 'Quote':
                                # See if the currently loaded Document matches the Document number sent from the Message Server
                                # or the currently loaded Quote matches the Quote Number sent from the Message Server
                                if (isinstance(self.ControlObject.currentObj, Document.Document) and \
                                   self.ControlObject.currentObj.number == int(msgData[2])) or \
                                   (isinstance(self.ControlObject.currentObj, Quote.Quote) and \
                                   self.ControlObject.currentObj.number == int(msgData[1])):
                                    # ... we need to update the keyword visualization.
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

                                print "ChatWindow.OnPostMessage():  UKV not processed for ", msgData[0]
                                print

                        # Update Snapshot
                        elif messageHeader == 'US':
                            # Start with a list of all the open Snapshot Windows.
                            openSnapshotWindows = self.ControlObject.GetOpenSnapshotWindows()
                            # Interate through the Snapshot Windows
                            for win in openSnapshotWindows:
                                # If the snapshot in question is open ...
                                if (win.obj.number == int(message)):
                                    # ... update the Snapshot Window
                                    win.FileClear(event)
                                    win.FileRestore(event)
                                    win.OnEnterWindow(event)
                                    
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

                    # Unless we've just deleted it ...
                    if messageHeader != 'DN':
                        # First, de-select all items
                        self.ControlObject.DataWindow.DBTab.tree.UnselectAll()
                        for currNode in currentSelection:
                            # ... now that we're done, we should re-select the originally-selected tree item
                            self.ControlObject.DataWindow.DBTab.tree.SelectItem(currNode)
                else:
                    if DEBUG:
                        print "We DON'T need to add an object, as we created it in the first place."

    def OnMessageServerLost(self, event):
        dlg = Dialogs.ErrorDialog(None, _("Your connection to the Message Server has been lost.\nYou may have lost your connection to the network, or there may be a problem with the Server.\nPlease quit Transana immediately and resolve the problem."))
        dlg.ShowModal()
        dlg.Destroy()
        # If Transana MU is left on overnight and loses connection to the Message Server, an exception can occur here.
        try:
            self.Close()
        except wx._core.PyDeadObjectError:
            pass
        TransanaGlobal.chatWindow = None
        try:
            wx.CallAfter(self.Destroy)
        except:
            pass

    def OnCloseMessage(self, event):
        """ Process the custom Close Message """
        # Close the Chat Window
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
            # Change the Timer interval to a minute
            self.validationTimer.Start(60000)
            # Display an error message to the user.
            self.memo.AppendText(_('Transana-MU:  The Transana Message Server has not been validated.\n'));
            self.memo.AppendText(_('Transana-MU:  Your connection to the Transana Message Server may have failed, or\n'));
            self.memo.AppendText(_('Transana-MU:  you may be using an improper version of the Transana Message Server,\n'));
            self.memo.AppendText(_('Transana-MU:  one that is different than this version of Transana-MU requires.\n'));
            self.memo.AppendText(_('Transana-MU:  Please report this problem to your system administrator.\n\n'));
            self.memo.AppendText(_('Transana-MU:  Please do not proceed.  Data corruption could result.\n\n'));

    def OnClose(self, event):
        """ Intercept when the Close Button is selected """
        # When the Close Button is selected, we should HIDE the form, but not Close it entirely
        self.Show(False)
    
    def OnFormClose(self, event):
        """ Form Close handler for when the form should be destroyed, not just hidden """
        try:
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
        self.useSound.SetLabel(_("Sound Enabled"))
        # Change the Buttons
        self.btnSend.SetLabel(_("Send"))
        self.btnClear.SetLabel(_("Clear"))


def ConnectToMessageServer():
    """ Create a connection to the Transana Message Server.
        This function returns a socket object if the connection is successful,
        or None if it is not.  """
    # If there is already a Chat Window open ...
    if TransanaGlobal.chatWindow != None:
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

            # If an SSL Certificate File has been defined ...
            if TransanaGlobal.configData.ssl and (TransanaGlobal.configData.sslMsgSrvCert != ''):

                if DEBUG:
                    print TransanaGlobal.configData.sslMsgSrvCert, os.path.exists(TransanaGlobal.configData.sslMsgSrvCert)

                # Define the Socket connection
                socketObj_plain = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                # Start Exception Handling to catch SSL errors
                try:

                    # Add the SSL wrapper to the socket connection
                    socketObj = ssl.wrap_socket(socketObj_plain,
                                                ca_certs=TransanaGlobal.configData.sslMsgSrvCert,
                                                cert_reqs=ssl.CERT_REQUIRED)
                    # Connect to the socket.  Server and Port are pulled from the Transana Configuration File
                    socketObj.connect((TransanaGlobal.configData.messageServer, TransanaGlobal.configData.messageServerPort + 1))

                    if DEBUG:
                        print "ChatWindow.ConnectToMessageServer():  SSL Connection established"

                    # Note that the Chat WIndow is using SSL
                    TransanaGlobal.chatIsSSL = True

                except ssl.SSLError:

                    prompt = _("SSL Connection Error:")
                    prompt += '\n\n%s\n\n' % sys.exc_info()[1]
                    prompt += _("Transana is establishing an un-encrypted connection to the Message Server.")

                    tmpDlg = Dialogs.ErrorDialog(None, prompt, _("SSL Connection Error"))
                    tmpDlg.ShowModal()
                    tmpDlg.Destroy()

                    # On SSL failure, we need to delete the sockets and start over to make an un-encrypted connection
                    try:
                        del(socketObj)
                        del(socketObj_plain)
                    except:
                        pass
                    
                    # Define the Socket connection
                    socketObj = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    # Connect to the socket.  Server and Port are pulled from the Transana Configuration File
                    socketObj.connect((TransanaGlobal.configData.messageServer, TransanaGlobal.configData.messageServerPort))

                    if DEBUG:
                        print "ChatWindow.ConnectToMessageServer():  Un-encoded Connection established"

                    # Note that the Chat Window is NOT using SSL
                    TransanaGlobal.chatIsSSL = False

                except socket.error:

                    prompt = _("SSL Socket Error:")
                    prompt += '\n\n%s\n\n' % sys.exc_info()[1]
                    prompt += _("Transana is establishing an un-encrypted connection to the Message Server.")

                    tmpDlg = Dialogs.ErrorDialog(None, prompt, _("SSL Connection Error"))
                    tmpDlg.ShowModal()
                    tmpDlg.Destroy()

                    # On SSL failure, we need to delete the sockets and start over to make an un-encrypted connection
                    try:
                        del(socketObj)
                        del(socketObj_plain)
                    except:
                        pass
                    
                    # Define the Socket connection
                    socketObj = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    # Connect to the socket.  Server and Port are pulled from the Transana Configuration File
                    socketObj.connect((TransanaGlobal.configData.messageServer, TransanaGlobal.configData.messageServerPort))

                    if DEBUG:
                        print "ChatWindow.ConnectToMessageServer():  Un-encoded Connection established"

                    # Note that the Chat Window is NOT using SSL
                    TransanaGlobal.chatIsSSL = False

            else:
                # Define the Socket connection
                socketObj = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

                # Connect to the socket.  Server and Port are pulled from the Transana Configuration File
                socketObj.connect((TransanaGlobal.configData.messageServer, TransanaGlobal.configData.messageServerPort))

                if DEBUG:
                    print "ChatWindow.ConnectToMessageServer():  Un-encoded Connection established"

                # Note that the Chat Window is NOT using SSL
                TransanaGlobal.chatIsSSL = False

            # Create the Chat Window Frame, passing in the socket object.
            # This window is NOT SHOWN at this time.
            TransanaGlobal.chatWindow = ChatWindow(None, -1, _("Transana Chat Window"), socketObj)
            # If we get this far, the connection was successful.
            ConnectedToMessageServer = True
        except socket.error:

            print sys.exc_info()[0], sys.exc_info()[1]
            import traceback
            print traceback.print_exc(file=sys.stdout)

            # Note that the Chat Window is NOT using SSL
            TransanaGlobal.chatIsSSL = False

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
    SERVERHOST = '192.168.1.19'
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
        print sys.exc_info()[0], sys.exc_info()[1]
        import traceback
        print traceback.print_exc(file=sys.stdout)
    except:
        print sys.exc_info()[0], sys.exc_info()[1]
        import traceback
        print traceback.print_exc(file=sys.stdout)
