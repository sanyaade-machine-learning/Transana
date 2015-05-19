# Copyright (C) 2003 - 2009 The Board of Regents of the University of Wisconsin System 
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

""" A Message Server that processes inter-instance communication between Transana-MU
    clients.  This server utility enables synchronization of the Data Windows. """
# NOTE:  Major rewrite started on 12/19/2005.  To work as a Service on Windows, we really need all of this
#        instantiated in a class instead of just thrown out there in flat, linear code.

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

VERSION = 240
# NOTE:  Remember to update the version number for the _svc_display_name_ too!

# import Python's re (regular expression) module
import re
# import python's socket module
import socket
# import python's sys module
import sys
# import python's thread module
import threading
# import Python's time module
import time

# Indicate whether DEBUG messages should be shown or not.
DEBUG = False
if DEBUG:
    print "MessageServer DEBUG is ON!"

# We can run the MessageServer as a Stand-alone utility program for debugging
# or it can run as a Service on Windows.
RUNASWINSERVICE = True and (sys.platform == 'win32') and (not DEBUG)

# NOTE:  DEBUG cannot be used if running as a Service
if DEBUG and RUNASWINSERVICE:
    print "You can't have DEBUG turned on if you're running as a Service!"
    sys.exit(1)

# If we want to run as a service, we need to take some extra steps
if RUNASWINSERVICE:
    # Import Python's win32service module
    import win32serviceutil
    import win32service
    import win32event

# The MessageServer is ALWAYS run as localhost, as it can only run on the machine running it!!
myHost = ''   # '' indicates localhost
# The Transana message Server defaults to port 17595
myPort = 17595

# Allow a different port to be passed in as a command-line parameter
if len(sys.argv) > 1:
    try:
        # The port must be an integer
        myPort = int(sys.argv[1])
    except:
        if not RUNASWINSERVICE:
            print "usage:  MessageServer [port]   (the default port is 17595.)"
            print "        MessageServer 17590    would launch the Message Server using port 17590."
        
            # NOTE:  The "sys.exit" line kills the Service.  It won't even install with this line present in the code.
            sys.exit(1)

def now():
    """ function that returns the current date and time """
    return time.ctime(time.time())


# Establish the capacity to lock the threads
threadLock = threading.Lock()

# Create a subclass of Thread for socket connections
class connectionThread(threading.Thread):
    """ Subclass of threading.thread designed to handle socket connections in threads """
    def __init__(self, parent, id, connection, address):
        # Initialize the Thread object
        threading.Thread.__init__(self)

        # We need a way to signal to the thread that it's time to quit!
        self.keepRunning = True
        # remember the parent object, dispatcher, which is common for all connection threads.
        self.parent = parent
        # id is the thread identifier, probably not used for anything
        self.id = id
        # connection is the connection established by the socket
        self.connection = connection
        # address is the address of the socket connection
        self.address = address

        if DEBUG:
            print "New thread spawned for %s, %s" % (self.connection, self.address)
            
        # Let's keep track of the number of errors that arise
        self.errorCount = 0

    def run(self):
        # running this thread involves listening to the socket and processing messages that come in.
        # Put this in a loop so that multiple messages will be processed.
        while self.keepRunning:
            # Broken socket connections cause exceptions, so process everything in a "try .. except" block
            try:
                if DEBUG:
                    print "Waiting for data from ", self.address
                    
                # Get data from the socket connection
                data = self.connection.recv(1024)
                if not data:
                    break
                # Get a Thread Lock.  
                threadLock.acquire()

                if DEBUG:
                    print 'Server received  "%s" from %s' % (data, self.address)

                # Remove the data separator, if there is one at the end of the message
                if data[-5:] == ' ||| ':
                    data = data[:-5]

                # If the message is a Connection message ...
                # (Format: "C Username DatabaseHost DatabaseName[ Version]")
                if (len(data) > 1) and (data[:data.find(' ')] == 'C'):
                    # Strip the connection flag from the message
                    st = data[2:].strip()
                    # extract the User Name
                    userName = st[:st.find(' ')]
                    # remove the User Name from the data string
                    st = st[st.find(' ') + 1:]
                    # extract the Database Host 
                    dbHost = st[:st.find(' ')].upper()
                    # remove the Database Host from the data string
                    st = st[st.find(' ') + 1:]
                    # See if there is a "Version" parameter
                    if st.find(' ') > -1:
                        # Extract the Database Name
                        dbName = st[:st.find(' ')].upper()
                        version = st[st.find(' ') + 1:]
                    else:
                        # Extract the Database Name
                        dbName = st.upper()
                        version = '100'

                    if DEBUG:
                        print 'New Connection: Username = "%s", dbHost = "%s", dbName = "%s", version = "%s"' % (userName, dbHost, dbName, version)
                        print "%d users are currently logged on. (1)" % (len(self.parent.connections) + 1)

                    # Check the Transana version against the Message Server version.
                    # Detect old Transana versions...
                    if int(version) < VERSION:
                        self.connection.send('M MessageServer: The Transana Message Server you have connected to is newer than ||| ')
                        self.connection.send('M MessageServer: your copy of Transana-MU.  Please upgrade your copy of Transana-MU ||| ')
                        self.connection.send('M MessageServer: immediately. ||| ')
                        self.connection.send('M MessageServer: - ||| ')
                        self.connection.send('M MessageServer: Please do not proceed.  Data corruption could result. ||| ')
                        self.connection.send('M MessageServer: - ||| ')
                    # Detect new Transana versions...
                    elif int(version) > VERSION:
                        self.connection.send('M MessageServer: The Transana Message Server you have connected to is older than ||| ')
                        self.connection.send('M MessageServer: your copy of Transana-MU.  Please ask your system administrator ||| ')
                        self.connection.send('M MessageServer: to upgrade your copy of the Transana-MU Message Server immediately. ||| ')
                        self.connection.send('M MessageServer: - ||| ')
                        self.connection.send('M MessageServer: Please do not proceed.  Data corruption could result. ||| ')
                        self.connection.send('M MessageServer: - ||| ')
                    else:
                        self.connection.send('V MessageServer: ServerValidated ||| ')

                    # Add the new User information to the Connection List
                    self.parent.connections[self.address] = {'name' : userName, 'dbHost' : dbHost, 'dbName' : dbName, 'version' : version}

                    # Keep doing this until told to stop.
                    while True:
                        # Assume we will NOT find the userName in the list.
                        nameFound = False

                        # Check for duplicate user names in the list of existing threads and avoid them.
                        for thr in threading.enumerate():
                            # ... check to be sure we've got a connectionThread, not the Main program thread ...
                            if type(thr) == type(self):

                                if DEBUG:
                                    print thr.address, self.parent.connections.has_key(thr.address)
                                
                                # ... if a connection is not THIS connection AND ...
                                # ... the connection has not been dropped AND ...
                                # ... the connection has the same Username ...
                                if (thr.address != self.address) and \
                                   self.parent.connections.has_key(thr.address) and \
                                   self.parent.connections[thr.address]['name'] == userName:

                                    # Define a regular expression to find the " (#)" portion of a name
                                    regex = re.compile("\(\d+\)")
                                    # Find that regular expression in the userName
                                    regexResult = regex.findall(userName)

                                    # If the regex is not found ...
                                    if regexResult == []:
                                        # ... then we add "(2)" to the username
                                        userName = userName + '(2)'
                                    # However, if the regex is found ...
                                    else:
                                        # ... extract the number from the string.
                                        # (the regexResult is in the form ['(#)'], so we extract the list element,
                                        #  then strip the parentheses from the string, then convert to an integer.)
                                        n = int(regexResult[0][1:-1])
                                        # Update the user name with the new number, in parentheses
                                        userName = userName[:userName.rfind('(')] + "(%d)" % (n + 1)

                                    # Update the connection list with the new name
                                    self.parent.connections[self.address]['name'] = userName
                                    # Prepare the data string to be broadcase with the new username
                                    data = 'C %s' % userName
                                    # Signal that we found the userName, so we can now search for the new iteration too
                                    nameFound = True

                                    # Inform the Chat Client that the username was updated
                                    self.connection.send('R %s ||| ' % userName)

                        # If we DIDN'T change the name, we can stop searching to see if the name's in use
                        if not nameFound:
                            break

                    # Broadcast the existing connection information to the new user.  Loop through all existing connection threads ...
                    for thr in threading.enumerate():
                        # ... check to be sure we've got a connectionThread, not the Main program thread ...
                        if type(thr) == type(self):
                            # ... if a connection is not THIS connection but has the same Database Host and
                            # Database Name ...
                            if (thr.address != self.address) and \
                               self.parent.connections[thr.address]['dbHost'] == self.parent.connections[self.address]['dbHost'] and \
                               self.parent.connections[thr.address]['dbName'] == self.parent.connections[self.address]['dbName']:
                                # ... then broadcast "C Username" to signal the names of other users in the same
                                # database as the connecting user.  This tells the newly connected user who else
                                # was already connected when s/he joined the group.
                                self.connection.send('C %s ||| ' % self.parent.connections[thr.address]['name'])

                elif (len(data) > 1) and (data[:data.find('0')] == 'D'):
                    # Substitute the correct user name, the one the Message Server knows.
                    data = 'D %s' % self.parent.connections[self.address]['name']

                if (data == 'M SHOW USERS') and (self.parent.connections[self.address]['name'] == 'DavidW'):

                    message = "M MessageServer: Users: ||| "
                    self.connection.send(message)
                    for c in self.parent.connections:
                        message = "M MessageServer: %s %s %s ||| " % (self.parent.connections[c]['name'], self.parent.connections[c]['dbHost'], self.parent.connections[c]['dbName'])
                        self.connection.send(message)
                        
                else:
                    # Broadcast the incoming message to others.  Loop through all existing connection threads.
                    for thr in threading.enumerate():
                        # This "if" ensures that only communication threads are processed, not the main program thread
                        if type(thr) == type(self):
                            try:
                                # pass the message to other threads.
                                thr.BroadcastMessage('%s' % data, self.parent.connections[self.address]['name'], self.parent.connections[self.address]['dbHost'], self.parent.connections[self.address]['dbName'])
                            except:
                                if DEBUG:
                                    print "Key Error??"
                                continue
                    # release the Thread Lock, as we're done.
                    if (len(data) > 1) and (data[:data.find(' ')] == 'D'):

                        if DEBUG:
                            print "D", self.address

                        # Remove this address from our connections list
                        del(self.parent.connections[self.address])
                        # Close the socket connection
                        self.connection.close()

                        if DEBUG:
                            print "Disconnection message from ", userName
                            print "%d users are currently logged on. (2)" % len(self.parent.connections)

                # Release the thread lock
                threadLock.release()

            # A socket.error indicates that one of the clients has crashed.
            except socket.error:

		try:
                    # Broadcast the loss of the user to other users
                    data = 'D %s' % self.parent.connections[self.address]['name']
                    # Broadcast the incoming message to others.  Loop through all existing connection threads.
                    for thr in threading.enumerate():
                        # This "if" ensures that only communication threads are processed, not the main program thread
                        if type(thr) == type(self):
                            try:
                                # pass the message to other threads.
                                thr.BroadcastMessage('%s' % data, self.parent.connections[self.address]['name'], self.parent.connections[self.address]['dbHost'], self.parent.connections[self.address]['dbName'])
                            except:
                                if DEBUG:
                                    print "Exception within a socket error"
                                continue
                    # Remove this address from our connections list
                    del(self.parent.connections[self.address])
                except:
                    pass  # forget about sending the message if it won't be sent.
                
                # Close the socket connection
                self.connection.close()

                if DEBUG:
                    # print "except socket.error"
                    print "Connection %s lost." % (self.address,)
                    print "%d users are currently logged on. (3)" % len(self.parent.connections)
                    if len(self.parent.connections) == 0:
                        print
                        print
                        print
                        
                # end the "while" loop that receives socket data
                break
            except:
                if DEBUG:
                    print "except"
                    print sys.exc_info()[0], sys.exc_info()[1]
                    import traceback
                    print traceback.print_exc(file=sys.stdout)
                self.errorCount += 1
                # Release the thread lock
                threadLock.release()

                if self.errorCount > 50:
                    if DEBUG:
                        print "break.  Too many errors."
                    break

        if DEBUG:
            print "A thread is actually dying here?"

    def BroadcastMessage(self, message, sendingUsername, sendingDBHost, sendingDBName):
        """ Selectively broadcast a message if the sender and recipient are on the same
            Database Host and are using the same Database Name """
        # Only forward the message to people on the same dbHost in the same DB
        if (self.parent.connections[self.address]['dbHost'] == sendingDBHost) and \
           (self.parent.connections[self.address]['dbName'] == sendingDBName):
            # If it's a message other than Connect, Disconnect, and Rename, insert the
            # username into the message.
            if (len(message) > 0) and not (message[:message.find(' ')] in ['C', 'D', 'R']):
                message = '%s %s: %s' % (message[:message.find(' ')], sendingUsername, message[message.find(' ') + 1:])
            if message.find(' ||| ') == -1:
                # Add the Message Terminator
                message += ' ||| '
            # Send the message

            if DEBUG:
                print "sending %s to %s" % (message, self.address)

            self.connection.send(message)
            
    def CommitSuicide(self):
        """ For this to work as a Service, the threads need the capacity to kill themselves!  Otherwise, the
            Service is unable to Stop. """
        # How do we kill the thread?  I'm trying this!  It should signal the thread to quit.
        self.keepRunning = False


class dispatcher(object):
    """ Detects new connections and dispatches a thread to handle it """
    def __init__(self):
        # Initialize a Thread Counter
        threadCount = 0
        # Maintain a dictionary of connections
        self.connections = {}

        if DEBUG:
            self.Report()

        # Create a TCP Socket object
        sockobj = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # Bind the socket to the port
        sockobj.bind((myHost, myPort))
        # Allow a max of 100 connections
        sockobj.listen(100)

        # Wait for new connections indefinitely
        while True:
            # Detect a new socket connection
            connection, address = sockobj.accept()
        
            if DEBUG:
                print 'Server connection established from %s at %s\n' % (address, now())
                print 'There are currently %d connections to this Message Server.' % (len(self.connections) + 1)

            # increment the Thread Counter
            threadCount += 1
            # Spawn a Connection Thread Object to handle the connection
            thread = connectionThread(self, threadCount, connection, address)
            # Start the new Thread
            thread.start()

    def Report(self):
        print
        print "Report for %s:" % now()
        print "Treads:"
        for thr in threading.enumerate():
            print thr, type(thr).__name__, str(type(thr))
        print "Connections:"
        for c in self.connections:
            print c, self.connections[c]['name'], self.connections[c]['dbHost'], self.connections[c]['dbName']
        print
        t = threading.Timer(30.0, self.Report)
        t.start()


    def KillAllThreads(self):
        for thr in threading.enumerate():
            if isinstance(thr, connectionThread):
                # Tell the threads they should kill themselves when they get their next message
                thr.CommitSuicide()
        
    if DEBUG:
        print "Starting MessageServer on port %d" % myPort

if RUNASWINSERVICE:

    # See Chapter 18 of Hammond and Robinson's "Python Programming on Win32"

    class TransanaMessageService(win32serviceutil.ServiceFramework):
        """ A Class designed to manage the Transana Message Server as a Windows Service """
        _svc_name_ = "TransanaMessageServer"
        _svc_display_name_ = "Transana Message Server, v. 2.40"
        def __init__(self, args):
            """ Initialize the Service """
            # Create a Service Framework
            win32serviceutil.ServiceFramework.__init__(self, args)
            # Create an event to wait on.  Service Stop will use this event.
            self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)            

        def SvcStop(self):
            """ Stop the Service """
            # Kill the current message threads
            self.dispatcher.KillAllThreads()
            # and kill the dispatcher
            del(self.dispatcher)
            # Tell the SCM we are starting the Stop
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            # Set the Stop Event
            win32event.SetEvent(self.hWaitStop)

        def SvcDoRun(self):
            """ Run the Service """
            self.dispatcher = dispatcher()
            win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

    if __name__ == '__main__':
        win32serviceutil.HandleCommandLine(TransanaMessageService)
else:
    # The main action of this program is to launch the Connection Thread Dispatcher
    dispatcher()
