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

""" A Message Server that processes inter-instance communication between Transana-MU
    clients.  This server utility enables synchronization of the Data Windows. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

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

# Indicate whether DEBUG messages should be shown or not
DEBUG = True

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
        print "usage:  MessageServer200 [port]   (the default port is 17595.)"
        sys.exit(1)

def now():
    """ function that returns the current date and time """
    return time.ctime(time.time())

# Maintain a dictionary of connections
connections = {}

# Create a TCP Socket object
sockobj = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
# Bind the socket to the port
sockobj.bind((myHost, myPort))
# Allow a max of 100 connections
sockobj.listen(100)

# Establish the capacity to lock the threads
threadLock = threading.Lock()


# Create a subclass of Thread for socket connections
class connectionThread(threading.Thread):
    """ Subclass of threading.thread designed to handle socket connections in threads """
    def __init__(self, id, connection, address):
        # id is the thread identifier, probably not used for anything
        self.id = id
        # connection is the connection established by the socket
        self.connection = connection
        # address is the address of the socket connection
        self.address = address
        # Initialize the Thread object
        threading.Thread.__init__(self)
        # Let's keep track of the number of errors that arise
        self.errorCount = 0

    def run(self):
        # running this thread involves listening to the socket and processing messages that come in.
        # Put this in a loop so that multiple messages will be processed.
        while True:
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
                if (len(data) > 1) and (data[0] == 'C'):
                    # Strip the connection flag from the message
                    st = data[2:].strip()
                    # extract the User Name
                    userName = st[:st.find(' ')]
                    # remove the User Name from the data string
                    st = st[st.find(' ') + 1:]
                    # extract the Database Host 
                    dbHost = st[:st.find(' ')]
                    # remove the Database Host from the data string
                    st = st[st.find(' ') + 1:]
                    # See if there is a "Version" parameter
                    if st.find(' ') > -1:
                        # Extract the Database Name
                        dbName = st[:st.find(' ')]
                        version = st[st.find(' ') + 1:]
                    else:
                        # Extract the Database Name
                        dbName = st
                        version = '100'

                    if DEBUG:
                        print 'New Connection: Username = "%s", dbHost = "%s", dbName = "%s", version = "%s"' % (userName, dbHost, dbName, version)
                        print "%d users are currently logged on. (1)" % (len(connections) + 1)

                    # Check the Transana version against the Message Server version.
                    # Detect old Transana versions...
                    if int(version) < 200:
                        self.connection.send('M MessageServer: The Transana Message Server you have connected to is newer than your copy of Transana-MU. ||| ')
                        self.connection.send('M MessageServer: Please upgrade your copy of Transana-MU immediately. ||| ')
                    # Detect new Transana versions...
                    elif int(version) > 200:
                        self.connection.send('M MessageServer: The Transana Message Server you have connected to is older than your copy of Transana-MU. ||| ')
                        self.connection.send('M MessageServer: Please ask your system administrator to upgrade your copy of the Transana-MU Message Server immediately. ||| ')
                    else:
                        self.connection.send('V MessageServer: ServerValidated ||| ')

                    # Add the new User information to the Connection List
                    connections[self.address] = {'name' : userName, 'dbHost' : dbHost, 'dbName' : dbName, 'version' : version}

                    # Keep doing this until told to stop.
                    while True:
                        # Assume we will NOT find the userName in the list.
                        nameFound = False
                        # Check for duplicate user names in the list of existing threads and avoid them.
                        for thr in threading.enumerate():
                            # ... check to be sure we've got a connectionThread, not the Main program thread ...
                            if type(thr) == type(self):
                                # ... if a connection is not THIS connection but has the same Username ...
                                if (thr.address != self.address) and \
                                   connections[thr.address]['name'] == userName:
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
                                    connections[self.address]['name'] = userName
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
                               connections[thr.address]['dbHost'] == connections[self.address]['dbHost'] and \
                               connections[thr.address]['dbName'] == connections[self.address]['dbName']:
                                # ... then broadcast "C Username" to signal the names of other users in the same
                                # database as the connecting user.  This tells the newly connected user who else
                                # was already connected when s/he joined the group.
                                self.connection.send('C %s ||| ' % connections[thr.address]['name'])
                elif (len(data) > 1) and (data[0] == 'D'):
                    # Substitute the correct user name, the one the Message Server knows.
                    data = 'D %s' % connections[self.address]['name']
                    
                # Broadcast the incoming message to others.  Loop through all existing connection threads.
                for thr in threading.enumerate():
                    # This "if" ensures that only communication threads are processed, not the main program thread
                    if type(thr) == type(self):
                        # pass the message to other threads.
                        thr.BroadcastMessage('%s' % data, connections[self.address]['name'], connections[self.address]['dbHost'], connections[self.address]['dbName'])
                # release the Thread Lock, as we're done.
                if (len(data) > 1) and (data[0] == 'D'):
                    del(connections[self.address])

                    if DEBUG or True:
                        print "Disconnection message from ", userName
                        print "%d users are currently logged on. (2)" % len(connections)
                        
                threadLock.release()
            except socket.error:

                if DEBUG:
                    print "Connection %s lost." % (self.address,)
                    print "%d users are currently logged on. (3)" % len(connections)

                    print sys.exc_info()[0], sys.exc_info()[1]
                    import traceback
                    print traceback.print_exc(file=sys.stdout)

                # end the "while" loop that receives socket data
                break
            except:
                if DEBUG:
                    print sys.exc_info()[0], sys.exc_info()[1]
                    import traceback
                    print traceback.print_exc(file=sys.stdout)
                self.errorCount += 1
                threadLock.release()
                if self.errorCount > 50:
                    print "break.  Too many errors."
                    break

    def BroadcastMessage(self, message, sendingUsername, sendingDBHost, sendingDBName):
        """ Selectively broadcast a message if the sender and recipient are on the same
            Database Host and are using the same Database Name """
        # Only forward the message to people on the same dbHost in the same DB
        if (connections[self.address]['dbHost'] == sendingDBHost) and \
           (connections[self.address]['dbName'] == sendingDBName):
            # If it's a text message, insert the username into the message
            if (len(message) > 0) and (message[0] == 'M'):
                message = '%s %s: %s' % (message[0], sendingUsername, message[2:])
            if message.find(' ||| ') == -1:
                # Add the Message Terminator
                message += ' ||| '
            # Send the message

            if DEBUG and False:
                print "sending %s to %s" % (message, self.address)
                
            self.connection.send(message)
        

def dispatcher():
    """ Detects new connections and dispatches a thread to handle it """
    # Initialize a Thread Counter
    threadCount = 0
    # Wait for new connections indefinitely
    while True:
        # Detect a new socket connection
        connection, address = sockobj.accept()
        
        if DEBUG:
            print 'Server connection established from %s at %s\n' % (address, now())
            print "%d users are currently logged on. (4)" % (len(connections) + 1)

        # increment the Thread Counter
        threadCount += 1
        # Spawn a Connection Thread Object to handle the connection
        thread = connectionThread(threadCount, connection, address)
        # Start the new Thread
        thread.start()

if DEBUG:
    print "Starting MessageServer on port %d" % myPort

# The main action of this program is to launch the Connection Thread Dispatcher
dispatcher()
