# Copyright (C) 2003 - 2005 The Board of Regents of the University of Wisconsin System 
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
# import Python sys module
import sys
# import Python time module
import time
# import Python threading module
import threading

# These values will eventually come from the Config object
SERVERHOST = 'WCER401964'
SERVERPORT = 17595
USERNAME = 'Macintosh1'
DBHOSTNAME = 'Mesquite'
DATABASENAME = 'DatabaseName'


print "CHATWINDOW NEEDS I18N!!"


# Get an ID for a custom Event for posting messages received in a thread
EVT_POST_MESSAGE_ID = wx.NewId()
# Get an ID for a custom Event for signalling that the Chat Window needs to be closed
EVT_CLOSE_MESSAGE_ID = wx.NewId()

# We create a thread to listen for messages from the Message Server.  However,
# only the Main program thread can interact with a wxPython GUI.  Therefore,
# we need to create a custom event to receive the event from the Message Server
# (via the socket connection) and transfer that data to the Main program thread
# where it will be displayed.  We need a second event to signal that the socket
# connection has died and that the Chat Window needs to be closed.

# Define a custom "Post Message" event
def EVT_POST_MESSAGE(win, func):
    """ Defines the EVT_POST_MESSAGE event type """
    win.Connect(-1, -1, EVT_POST_MESSAGE_ID, func)

def EVT_CLOSE_MESSAGE(win, func):
    """ Defines the EVT_CLOSE_MESSAGE event type """
    win.Connect(-1, -1, EVT_CLOSE_MESSAGE_ID, func)

# Create the actual Custom Event object
class PostMessageEvent(wx.PyEvent):
    """ This event is used to trigger posting a message in the GUI. It carries the data. """
    def __init__(self, data):
        # Initialise a wxPyEvent
        wx.PyEvent.__init__(self)
        # Link the event to the Event ID
        self.SetEventType(EVT_POST_MESSAGE_ID)
        # Store the message data in the Custom Event object
        self.data = data

# Create the actual Custom Event object
class CloseMessageEvent(wx.PyEvent):
    """ This event is used to trigger closing the GUI. """
    def __init__(self):
        # Initialise a wxPyEvent
        wx.PyEvent.__init__(self)
        # Link the event to the Event ID
        self.SetEventType(EVT_CLOSE_MESSAGE_ID)


threadLock = threading.Lock()

# Create a custom Thread object for listening for messages from the Message Server
class ListenerThread(threading.Thread):
    """ custom Thread object for listening for Messages from the Message Server """
    def __init__(self, notificationWindow, socketObj):
        # parameters are the GUI window that must be notified that a message has been received
        # and the socket object to listen to for messages.
        
        # Initialize the Thread object
        threading.Thread.__init__(self)
        # Remember the window that needs to be notified
        self.window = notificationWindow
        # Remember the socket object
        self.socketObj = socketObj
        # Signal that we don't yet want to abort the thread (probably not used)
        self._want_abort = False
        # prevent the application from hanging on Close
        self.setDaemon(1)
        # Start the thread
        self.start()

    def run(self):
        # We need to track overflow (incomplete messages) from the socket.  This allows us to do this.
        dataOverflow = ''
        # As long as the thread is running ...
        while not self._want_abort:
            try:
                # listen to the socket for a message from the Message Server
                data = dataOverflow + self.socketObj.recv(2048)
                dataOverflow = ''
            except socket.error:
                print "Your connection to the Message Server has been lost. (1)"
                self._want_abort = True
                wx.PostEvent(self.window, CloseMessageEvent())
                break


            if DEBUG:
                print 'Received:', data

            threadLock.acquire()
            # If the message is blank ...
            if not data:
                # post the Message Event with None as the data
                wx.PostEvent(self.window, PostMessageEvent(None))
                return
            else:
                # A single data element from the Message Server can contain multiple messages, separated
                # by the ' ||| ' message separator.  
                # Break a message into its component messages
                messages = data.split(' ||| ')

                print messages[-1]
                dataOverflow = messages[-1]
                
                # Process all the messages.  (The last message is always blank, so skip it!)
                for message in messages[:-1]:
                    # Post the Message Event with the message as data
                    wx.PostEvent(self.window, PostMessageEvent(message))
            threadLock.release()
        else:
            # If we want to abort the thread, post the message event with None as the data
            # NOTE:  This probably never gets called, as the code waits at the socket.recv
            #        line until it gets a message of the connection is broken.
            wx.PostEvent(self.window, PostMessageEvent(None))
            return

    def abort(self):
        # Signal that you want to abort the thread.  (Probably not used.)
        self._want_abort = True
        

class ChatWindow(wx.Frame):
    """ This window displays the Chat form. """
    def __init__(self,parent,id,title, socketObj):
        # remember the Socket object
        self.socketObj = socketObj
        self.userName = USERNAME
        # Define the main Frame for the Chat Window
        wx.Frame.__init__(self,parent,-4, title, size = (760,400), style=wx.DEFAULT_FRAME_STYLE|wx.NO_FULL_REPAINT_ON_RESIZE)
        # Set the background to White
        self.SetBackgroundColour(wx.WHITE)
        
        # Create a Sizer for the form
        box = wx.BoxSizer(wx.VERTICAL)

        # Create a Sizer for data the form receives
        boxRecv = wx.BoxSizer(wx.HORIZONTAL)
        # Add a TextCtrl for the chat text.  This is read only, as it is filled programmatically.
        self.memo = wx.TextCtrl(self, -1, style = wx.TE_MULTILINE | wx.TE_WORDWRAP | wx.TE_READONLY)
        # Add the Chat display to the Receiver Sizer
        boxRecv.Add(self.memo, 5, wx.EXPAND)

        # Add a ListBox to hold the names of active users
        self.userList = wx.ListBox(self, -1, choices=[self.userName], style=wx.LB_SINGLE)
        # Add the User List display to the Receiver Sizer
        boxRecv.Add(self.userList, 1, wx.EXPAND | wx.LEFT, 6)
        
        # Put the TextCtrl in the form sizer
        box.Add(boxRecv, 9, wx.EXPAND | wx.ALL, 4)

        # Create a sizer for the text entry and send portion of the form
        boxSend = wx.BoxSizer(wx.HORIZONTAL)
        # Add a TextCtrl where the user can enter messages.  It needs to process the Enter key.
        self.txtEntry = wx.TextCtrl(self, -1, style = wx.TE_PROCESS_ENTER)
        # Add the Text Entry control to the Text entry sizer
        boxSend.Add(self.txtEntry, 6, wx.EXPAND)
        # bind the OnSend event with the Text Entry control's enter key event
        self.txtEntry.Bind(wx.EVT_TEXT_ENTER, self.OnSend)

        # Add a "Send" button
        self.btnSend = wx.Button(self, -1, "Send")
        # Add the Send button to the Text Entry sizer
        boxSend.Add(self.btnSend, 0, wx.LEFT, 6)
        # Bind the OnSend event to the send button's press event
        self.btnSend.Bind(wx.EVT_BUTTON, self.OnSend)
        
        # Add the send Sizer to the Form Sizer
        box.Add(boxSend, 1, wx.EXPAND | wx.ALL, 4)

        # Define the form's OnClose handler
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        # Define the custom Post Message Event's handler
        EVT_POST_MESSAGE(self, self.OnPostMessage)
        # Define the custom Close Message Event's handler
        EVT_CLOSE_MESSAGE(self, self.OnCloseMessage)

        # Status Bar
        self.CreateStatusBar()
        self.SetStatusText('Connected to Transana Message Server "%s" on port %d as %s, %s, %s.' % (SERVERHOST, SERVERPORT, self.userName, DBHOSTNAME, DATABASENAME))

        # Attach the Form's Main Sizer to the form
        self.SetSizer(box)
        # Set AutoLayout on
        self.SetAutoLayout(True)
        # Lay out the form
        self.Layout()
        # Center the form on the screen
        self.CentreOnScreen()
        # Set the initial focus to the Text Entry control
        self.txtEntry.SetFocus()
        # Start the Listener Thread to listen for messages from the Message Server
        self.listener = ListenerThread(self, self.socketObj)
        
        # Show the Form
        self.Show(True)
        # Inform the message server of user's identity and database connection
        self.socketObj.send('C %s %s %s 200 ||| ' % (self.userName, DBHOSTNAME, DATABASENAME))

        # Create a Timer to check for Message Server validation
        self.serverValidation = False
        self.validationTimer = wx.Timer()
        self.validationTimer.Bind(wx.EVT_TIMER, self.OnValidationTimer)
        self.validationTimer.Start(10000)


    def OnSend(self, event):
        """ Send Message handler """
        # self.memo.AppendText("Sent:     %s\n" % self.txtEntry.GetValue())
        # indicate that this is a Text message by prefacing the text with "M"
        self.socketObj.send('M %s ||| ' % self.txtEntry.GetValue())
        # Clear the Text Entry control
        self.txtEntry.SetValue('')
        # Set the focus to the Text Entry Control
        self.txtEntry.SetFocus()
        
    def OnPostMessage(self, event):
        """ Post Message handler """
        # If there is data in the message event ...
        if event.data != None:
            # Text Message ?
            if event.data[0] == 'M':
                # ... display the message on the screen
                self.memo.AppendText("%s\n" % event.data[2:])
            # Connection Message ?
            elif event.data[0] == 'C':
                st = event.data[2:]
                if st.find(' ') > -1:
                    st = st[:st.find(' ')]

                if DEBUG:
                    print 'PostMessage:', event.data, st, self.userName
                
                if st != self.userName:
                    self.userList.Append(st)
            # Rename Message ?
            elif event.data[0] == 'R':
                self.userList.Delete(0)
                st = event.data[2:]
                if st.find(' ') > -1:
                    self.userName = st[:st.find(' ')]
                else:
                    self.userName = st
                self.SetStatusText('Connected to Transana Message Server "%s" on port %d as %s, %s, %s.' % (SERVERHOST, SERVERPORT, self.userName, DBHOSTNAME, DATABASENAME))
                self.userList.Append(st)
            # Server Validation ?
            elif event.data[0] == 'V':
                self.serverValidation = True
            # Disconnect Message ?
            elif event.data[0] == 'D':
                st = event.data[2:]
                self.userList.Delete(self.userList.FindString(st))
            else:

                if DEBUG:
                    print "Unprocessed Message: ", event.data
                
                self.memo.AppendText('Unprocessed Message: "%s"\n' % event.data)
                

    def OnCloseMessage(self, event):
        self.Close()
        
    def OnValidationTimer(self, event):
        """ Checks for the validity of the Message Server """
        if self.serverValidation:
            self.validationTimer.Stop()
        else:
            if not self.IsShown():
                self.Show(True)
            # Bring this window to the top of the others
            self.Raise()
            self.validationTimer.Start(60000)
            self.memo.AppendText('Transana-MU:  The Transana Message Server has not been validated.\n');
            self.memo.AppendText('Transana-MU:  Your connection to the Transana Message Server may have failed, or\n');
            self.memo.AppendText('Transana-MU:  you may be using an improper version of the Transana Message Server,\n');
            self.memo.AppendText('Transana-MU:  one that is different than this version of Transana-MU requires.\n');
            self.memo.AppendText('Transana-MU:  Please report this problem to your system administrator.\n\n');

        
    def OnClose(self, event):
        """ Form Close handler """
        try:
            # Inform the message server that you're disconnecting (Needed on Mac)
            self.socketObj.send('D %s ||| ' % self.userName)
            time.sleep(1)
        except socket.error:
            print "Your connection to the Message Server has been lost. (2)"
        except:
            print sys.exc_info()[0], sys.exc_info()[1]
            import traceback
            print traceback.print_exc(file=sys.stdout)
        # Try to tell the listener thread to abort (probably does nothing.)
        self.listener.abort()
        # Go on and close the form.
        event.Skip()


# For testing purposes, this code allows the Chat Window to function as a stand-alone application
if __name__ == '__main__':
    # Declare the main App object
    app = wx.PySimpleApp()
    try:
        # Define the Socket connection
        socketObj = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # Connect to the socket
        socketObj.connect((SERVERHOST, SERVERPORT))
        # Create the Chat Window Frame, passing in the socket object
        frame = ChatWindow(None, -1, "Chat Window", socketObj)
        # run the MainLoop
        app.MainLoop()
    except socket.error:
        print 'Unable to connect to a Message Server'
    except:
        print sys.exc_info()[0], sys.exc_info()[1]
        import traceback
        print traceback.print_exc(file=sys.stdout)
        
