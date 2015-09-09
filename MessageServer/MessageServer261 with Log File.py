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

""" A Message Server that processes inter-instance communication between Transana-MU
    clients.  This server utility enables synchronization of the Data Windows. """

# NOTE:  Major rewrite started on 12/19/2005.  To work as a Service on Windows, we really need all of this
#        instantiated in a class instead of just thrown out there in flat, linear code

# As of September 10, 2014, there are problems using SSL to connect to the Transana Message Server on OS X.
#
# I've experimented extensively with Python 2.6.7, Python 2.7.8, and Python 3.4.1.
# Python 2.6.7 and Python 2.7.8 have a problem in the Python ssl module that prevents them from connecting to the
# Transana Message Server correctly from OS X if the Certificate files were created on Ubuntu 14.04.  Python 3.4.1
# doesn't have this problem.  However, there is not currently a version of wxPython for Python 3.4.1 that supports
# all of the widgets that Transana requires.  Certificate files created on Windows 7 and on OS X 10.7.5 work fine.
#
# I have modified the Transana Message Server to handle both SSL and non-SSL connections at the same time.

# Daemon code based on http://www.jejik.com/articles/2007/02/a_simple_unix_linux_daemon_in_python/

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Indicate whether DEBUG messages should be shown or not.
DEBUG = False
DEBUG2 = False   # Report output
if DEBUG:
    print "MessageServer DEBUG is ON!"

VERSION = 261
# NOTE:  Remember to update the version number for the _svc_display_name_ too!

# import the Python os module
import os
# import Python's re (regular expression) module
import re
# import Python's SSL module
import ssl
# import python's socket module
import socket
# import python's sys module
import sys
# import python's thread module
import threading
# import Python's time module
import time

# We can run the MessageServer as a Stand-alone utility program for debugging
# or it can run as a Service on Windows
# or it can run as a daemon on OS X

# (NOTE:  Service and Daemon handling are coded separately and in parallel.)

RUNASWINSERVICE = True and (sys.platform == 'win32') and (not DEBUG)
RUNASDAEMON = True and (sys.platform in ['darwin', 'linux2']) and (not DEBUG)


# NOTE:  DEBUG cannot be used if running as a Service
if DEBUG and RUNASWINSERVICE:
    print "You can't have DEBUG turned on if you're running as a Service!"
    sys.exit(1)

# NOTE:  DEBUG cannot be used if running as a Daemon
if DEBUG and RUNASDAEMON:
    print "You can't have DEBUG turned on if you're running as a Daemon!"
    sys.exit(1)

# If we want to run as a service, we need to take some extra steps
if RUNASWINSERVICE:
    # Import Python's win32service module
    import win32serviceutil
    import win32service
    import win32event

# If we want to run as a daemon, we need to take some extra steps
if RUNASDAEMON:
    # import the daemon class
    import daemon

# Interval for checking client validation (in seconds)
checkInterval = 20.0
# The MessageServer is ALWAYS run as localhost, as it can only run on the machine running it!!
myHost = ''   # '' indicates localhost
# The Transana message Server defaults to port 17595
myPort = 17595
mySSLPort = 17596


# We need to know what directory the program is running from.  We use this to find
# the Certificate Files in the SSL directory.
# This has emerged as the "preferred" cross-platform method on the wxPython-users list.
programDir = os.path.abspath(sys.path[0])
# Okay, that doesn't work with wxversion, which adds to the path.  Here's the fix, I hope.
# This should over-ride the programDir with the first value that contains Transana in the path.
for path in sys.path:
    if 'messageserver' in path.lower():
        programDir = path
        break
if os.path.isfile(programDir):
    programDir = os.path.dirname(programDir)

# I've renamed the files from the example below
CERT_FILE = os.path.join(programDir, 'SSL', "TransanaMessageServer-cert.pem")  # public key
CERT_KEY = os.path.join(programDir, 'SSL', "TransanaMessageServer-key.pem")    # private key


# The MySQL web site describes creating the necessary files at http://dev.mysql.com/doc/refman/5.5/en/creating-ssl-certs.html
# I *think* I did something like this (on Linux, and sharing certificate files across platforms):
#
# openssl genrsa 2048 > ca-key.pem
# openssl req -new -x509 -nodes -days 3600 -key ca-key.pem -out ca-cert.pem
# openssl req -newkey rsa:2048 -days 3600 -nodes -keyout server-key.pem -out server-req.pem
# openssl rsa -in server-key.pem -out server-key.pem
# openssl x509 -req -in server-req.pem -days 3600 -CA ca-cert.pem -CAkey ca-key.pem -set_serial 01 -out server-cert.pem
# openssl req -newkey rsa:2048 -days 3600 -nodes -keyout client-key.pem -out client-req.pem
# openssl rsa -in client-key.pem -out client-key.pem
# openssl x509 -req -in client-req.pem -days 3600 -CA ca-cert.pem -CAkey ca-key.pem -set_serial 01 -out client-cert.pem

# These commands generate the necessary certificate files for using SSL, as described at
# http://carlo-hamalainen.net/blog/2013/1/24/python-ssl-socket-echo-test-with-self-signed-certificate :
#
# openssl genrsa -des3 -out transanamessageserver.orig.key 2048
# openssl rsa -in transanamessageserver.orig.key -out TransanaMessageServer-key.pem
# openssl req -new -key TransanaMessageServer-key.pem -out TransanaMessageServer-cert.pem
# openssl x509 -req -days 3600 -in TransanaMessageServer-cert.pem -signkey TransanaMessageServer-key.pem -out TransanaMessageServer-cert.pem


if DEBUG:
    print
    print 'Cert File:', CERT_FILE, os.path.exists(CERT_FILE)
    print ' Key File:', CERT_KEY, os.path.exists(CERT_KEY)
    print

# If you want a log file of logins, enable this path.  If blank, no log file is saved.
LOGFILE = ''  # os.path.join(os.sep, 'var', 'log', 'TMS_Logins.log')

if DEBUG:
    print 'Log File:', LOGFILE
    print

# Allow a different port to be passed in as a command-line parameter
if len(sys.argv) > 1:
    try:
        if sys.platform == 'win32':
            # The port must be an integer
            myPort = int(sys.argv[1])
            mySSLPort = myPort + 1
        else:
            cmd = sys.argv[1]
            if len(sys.argv) > 2:
                # The port must be an integer
                myPort = int(sys.argv[2])
                mySSLPort = myPort + 1
    except:
        if not RUNASWINSERVICE:
            if sys.platform == 'win32':
                print "usage:  MessageServer [port]   (the default port is 17595, 17596 for SSL.)"
                print "        MessageServer 17590    would launch the Message Server using port 17590, 17591 for SSL."
            else:
                print "usage:  MessageServer start|stop|restart [port]   (the default port is 17595, 17596 for SSL.)"
                print "        MessageServer start 17590    would launch the Message Server using port 17590, 17591 for SSL."
        
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

        # Initialize the User Name
        userName = ''
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
                # (Format: "C Username DatabaseHost DatabaseName [SSL] Version")
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
                    # See if there is another parameter after dbName
                    if st.find(' ') > -1:
                        # extract the Database Name 
                        dbName = st[:st.find(' ')].upper()
                        # remove the Database Name from the data string
                        st = st[st.find(' ') + 1:]
                        # See if there are both "SSL" and "Version" parameter
                        if st.find(' ') > -1:
                            # Extract the SSL value ...
                            SSL = st[:st.find(' ')].upper()
                            # ... and conver the Version value
                            version = st[st.find(' ') + 1:]
                        else:
                            # If there's no SSL value, then SSL is FALSE!!
                            SSL = 'FALSE'
                            # Extract the VERSION value
                            version = st.upper()
                    # If dbName is the last parameter
                    else:
                        # ... extract the Database Name
                        dbName = st.upper()
                        # ... set SSL to False
                        SSL = 'FALSE'
                        # ... and default Version to 1.00
                        version = '100'

                    if DEBUG:
                        print 'New Connection: Username = "%s", dbHost = "%s", dbName = "%s", SSL = "%s", version = "%s"' % (userName, dbHost, dbName, SSL, version)
                        print "%d users are currently logged on. (1)" % (len(self.parent.connections) + 1)

                    # Check the Transana version against the Message Server version.
                    # Detect Transana 2.60 on a Transana 2.61 server
                    if (int(version) == 260) and (VERSION == 261):
                        self.connection.send('M MessageServer: The Transana Message Server you have connected to is newer than ||| ')
                        self.connection.send('M MessageServer: your copy of Transana-MU.  Please upgrade your copy of Transana-MU ||| ')
                        self.connection.send('M MessageServer: as soon as you are able. ||| ')
                        self.connection.send('M MessageServer: - ||| ')
                        # But we can go ahead and validate the Message Server, as it WILL work!
                        self.connection.send('V MessageServer: ServerValidated ||| ')
                    # Detect old Transana versions...
                    elif int(version) < VERSION:
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

                    if LOGFILE != '':

                        try:

                            if DEBUG:
                                print "Opening LogFile:", LOGFILE
                            
                            f = file(LOGFILE, 'a')
                            t = time.localtime()

                            if DEBUG:
                                print "Writing to LogFile:", '%4d/%02d/%02d %02d:%02d:%02d\t%s\t%s\t%s\t%s\t%s\n' % \
                                      (t.tm_year, t.tm_mon, t.tm_mday, t.tm_hour, t.tm_min, t.tm_sec, userName, dbHost, dbName, SSL, version)
                            
                            f.write('%4d/%02d/%02d %02d:%02d:%02d\t%s\t%s\t%s\t%s\t%s\n' % \
                                    (t.tm_year, t.tm_mon, t.tm_mday, t.tm_hour, t.tm_min, t.tm_sec, userName, dbHost, dbName, SSL, version))

                            f.close()
                            
                        except:

                            if DEBUG:
                                print
                                print "Exception in opening log file:"
                                print sys.exc_info()[0]
                                print sys.exc_info()[1]
                                import traceback
                                traceback.print_exc(file=sys.stdout)
                                print

                            pass

                    # Add the new User information to the Connection List
                    self.parent.connections[self.address] = {'name' : userName, 'dbHost' : dbHost, 'dbName' : dbName, 'ssl' : SSL, 'version' : version}

                    # Keep doing this until told to stop.
                    while True:
                        # Assume we will NOT find the userName in the list.
                        nameFound = False

                        # Check for duplicate user names in the list of existing threads and avoid them.
                        for thr in threading.enumerate():
                            # ... check to be sure we've got a connectionThread, not the Main program thread ...
                            if isinstance(thr, connectionThread):

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
                                    data = 'C %s %s' % (userName, SSL)
                                    # Signal that we found the userName, so we can now search for the new iteration too
                                    nameFound = True

                                    # Inform the Chat Client that the username was updated
                                    self.connection.send('R %s %s ||| ' % (userName, SSL))

                        # If we DIDN'T change the name, we can stop searching to see if the name's in use
                        if not nameFound:
                            break

                    # Broadcast the existing connection information to the new user.  Loop through all existing connection threads ...
                    for thr in threading.enumerate():
                        # ... check to be sure we've got a connectionThread, not the Main program thread ...
                        if isinstance(thr, connectionThread):
                            # ... if a connection is not THIS connection but has the same Database Host and
                            # Database Name ...
                            if (thr.address != self.address) and \
                               self.parent.connections[thr.address]['dbHost'] == self.parent.connections[self.address]['dbHost'] and \
                               self.parent.connections[thr.address]['dbName'] == self.parent.connections[self.address]['dbName']:
                                # ... then broadcast "C Username" to signal the names of other users in the same
                                # database as the connecting user.  This tells the newly connected user who else
                                # was already connected when s/he joined the group.
                                self.connection.send('C %s %s ||| ' % (self.parent.connections[thr.address]['name'], self.parent.connections[thr.address]['ssl']))

                elif (len(data) > 1) and (data[:data.find('0')] == 'D'):
                    # Substitute the correct user name, the one the Message Server knows.
                    data = 'D %s' % self.parent.connections[self.address]['name']

                if (data == 'M SHOW USERS') and (self.parent.connections[self.address]['name'] in ['DavidW', 'DavidW(2)', 'DavidW(3)']):

                    message = "M MessageServer: Users: ||| "
                    self.connection.send(message)
                    for c in self.parent.connections:
                        message = "M MessageServer: %s %s %s %s ||| " % (self.parent.connections[c]['name'], self.parent.connections[c]['dbHost'], self.parent.connections[c]['dbName'], self.parent.connections[c]['ssl'])
                        self.connection.send(message)

                elif (data == 'M SHOW CERTIFICATES'):
                    
                    message = "M MessageServer: %s %s ||| " % (CERT_FILE, os.path.exists(CERT_FILE))
                    self.connection.send(message)
                    message = "M MessageServer: %s %s ||| " % (CERT_KEY, os.path.exists(CERT_KEY))
                    self.connection.send(message)

                elif (data == 'M RESET USERS') and (self.parent.connections[self.address]['name'] in ['DavidW', 'DavidW(2)', 'DavidW(3)']):
                    message = "M MessageServer: RESET -  Exit immediately.  %s %s %s ||| " % (self.parent.connections[self.address]['name'], self.parent.connections[self.address]['dbHost'], self.parent.connections[self.address]['dbName'])
                    self.connection.send(message)
                    self.parent.KillAllThreads()
                    self.parent.connections = {}

                else:
                    # Broadcast the incoming message to others.  Loop through all existing connection threads.
                    for thr in threading.enumerate():
                        # This "if" ensures that only communication threads are processed, not the main program thread
                        if isinstance(thr, connectionThread):
                            try:
                                # pass the message to other threads.
                                thr.BroadcastMessage('%s' % data, self.parent.connections[self.address]['name'], self.parent.connections[self.address]['dbHost'], self.parent.connections[self.address]['dbName'])
                            except:
                                if DEBUG:
                                    print "Key Error??"

                                    print sys.exc_info()[0], sys.exc_info()[1]
                                    import traceback
                                    traceback.print_exc(file=sys.stdout)
                                    
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

                if DEBUG:
                    print "Socket Error"
                    import traceback
                    traceback.print_exc(file=sys.stdout)

		try:
                    # Broadcast the loss of the user to other users
                    data = 'D %s' % self.parent.connections[self.address]['name']
                    # Broadcast the incoming message to others.  Loop through all existing connection threads.
                    for thr in threading.enumerate():
                        # This "if" ensures that only communication threads are processed, not the main program thread
                        if isinstance(thr, connectionThread):
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
            print "A thread is actually dying here.  %s has left the building" % userName

    def BroadcastMessage(self, message, sendingUsername, sendingDBHost, sendingDBName):
        """ Selectively broadcast a message if the sender and recipient are on the same
            Database Host and are using the same Database Name """
        # Detect Private Messages
        if (message[0:2] == 'M ') and (message.find(' >|< ') > -1):
            # Create a recipient list, starting with the originating user
            recipientList = [sendingUsername]
            # Split the data into the message and the recipient list
            messageParts = message.split(' >|< ')
            # Add the recipients to the recipient list
            for user in messageParts[1].split(' '):
                recipientList.append(user)
            # Save the recipient list as a string
            recipientString = messageParts[1]
            # Remove the recipient list from the message
            message = messageParts[0]
            # If this message is NOT terminated ...
            if message.find(' ||| ') == -1:
                # ... add the Private Message indicator and the terminator
                message += "  (" + "private message to " + recipientString + ") ||| "
            # If this message is terminated ...
            else:
                # ... insert the Private Message indicator before the terminator
                message = message[: message.find(' ||| ')] + "  (" + "private message for " + recipientString + ")  " + \
                            message[message.find(' ||| ') + 1:]
        # If this is NOT a Private Message ...
        else:
            # ... initialize the Recipient list to None to indicate a public message
            recipientList = None
                
        # Only forward the message to people on the same dbHost in the same DB
        if (self.parent.connections[self.address]['dbHost'] == sendingDBHost) and \
           (self.parent.connections[self.address]['dbName'] == sendingDBName) and \
           ((recipientList == None) or (self.parent.connections[self.address]['name'] in recipientList)):
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


# Create a subclass of Thread for socket listeners that accept socket connections
class listenerThread(threading.Thread):
    def __init__(self, parent, useSSL):
        # Initialize the Thread object
        threading.Thread.__init__(self)

        # We need a way to signal to the thread that it's time to quit!
        self.keepRunning = True

        self.parent = parent
        self.useSSL = useSSL
        
        # Create a TCP Socket object
        self.sockobj = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

        if useSSL:

            if DEBUG:
                print "Binding SSL port", mySSLPort
                print
                
            # Bind the socket to the port
            self.sockobj.bind((myHost, mySSLPort))
            self.port = mySSLPort
        else:

            if DEBUG:
                print "Binding port", myPort
                print
                
            # Bind the socket to the port
            self.sockobj.bind((myHost, myPort))
            self.port = myPort
            
        # Allow a max of 100 connections
        self.sockobj.listen(100)

    def run(self):
        # running this thread involves listening to the socket and processing messages that come in.
        # Put this in a loop so that multiple messages will be processed.

        while self.keepRunning:

            try:

                if self.useSSL:

                    if DEBUG:
                        print "Creating SSL Connection"

                    # Detect a new socket connection
                    connection_plain, address = self.sockobj.accept()


                    # Wrap the connection using SSL
                    connection = ssl.wrap_socket(connection_plain,
                                                 server_side=True,
                                                 certfile=CERT_FILE,
                                                 keyfile=CERT_KEY)

                else:

                    if DEBUG:
                        print "Creating Connection without SSL"


                    # Detect a new socket connection
                    connection, address = self.sockobj.accept()

                if DEBUG:
                    print 'Server connection established from %s at %s\n' % (address, now())
                    print 'There are currently %d connections to this Message Server.' % (len(self.parent.connections) + 1)

                # increment the Thread Counter
                self.parent.threadCount += 1
                # Spawn a Connection Thread Object to handle the connection
                thread = connectionThread(self.parent, self.parent.threadCount, connection, address)
                # Start the new Thread
                thread.start()
            except:
                if DEBUG:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    print

                pass

        
    def CommitSuicide(self):
        """ For this to work as a Service, the threads need the capacity to kill themselves!  Otherwise, the
            Service is unable to Stop. """
        # How do we kill the thread?  I'm trying this!  It should signal the thread to quit.
        self.keepRunning = False



class dispatcher(object):
    """ This starts the Transana Message Server.  It creates two Listener threads, one for
        un-encrypted Socket connections and one for SSL-encrypted SSL connections """
    def __init__(self):
        # Initialize a Thread Counter
        self.threadCount = 0
        # Maintain a dictionary of connections
        self.connections = {}

        # Call the Report.  (This initiates a Timer that keeps it going!)
        self.Report()

        # Create and start an un-encrypted Listener thread
        plainListener = listenerThread(self, False)
        plainListener.start()

        # Create and start an encrypted Listener thread
        sslListener = listenerThread(self, True)
        sslListener.start()

        if DEBUG:
            print "Starting MessageServer on ports %d (unencrypted) and %d (SSL-encrypted)" % (myPort, mySSLPort)

    def Report(self):
        """ Report status in DEBUG Mode, but do some connection maintenance whether the Report
            output is displayed or not """
        
        if DEBUG and DEBUG2:
            print
            print "Report for %s:" % now()
            print "Treads:"
            for thr in threading.enumerate():
                if isinstance(thr, listenerThread):
                    print "Listener Thread", thr.port, thr.useSSL
                if isinstance(thr, connectionThread):
                    print "Connection Thread:", thr, self.connections[thr.address]['name']
            print "Connections:"
            for c in self.connections:
                print c, self.connections[c]['name'], self.connections[c]['dbHost'], self.connections[c]['dbName']
            print

        # Create an empty list to track existing threads
        threadList = []
        # for each existing thread ...
        for thr in threading.enumerate():
            # If this is a connection thread (not the Timer or the program's main thread) ...
            if isinstance(thr, connectionThread):
                # add this thread's address to the thread list
                threadList.append(thr.address)

        # Since we can't change the self.connections dictionary while iterating through it,
        # initialize a list of the threads NOT found in the connections dictionary
        threadsNotFound = []
        # Iterate through the known connections
        for c in self.connections:
            # If we come across a connection not listed in the threads, we know a Mac crashed
            # and didn't clean up after itself correctly.
            if not c in threadList:

                if DEBUG  and DEBUG2:
                    print '------------------------------------'
                    print c, 'not in', threadList
                    print self.connections[c]
                    print '------------------------------------'

                # ... so add this connection to the list of connections not found
                threadsNotFound.append(c)

        # For each connection where we didn't find a thread ...
        for c in threadsNotFound:
            try:
                # Broadcast the loss of the user to other users
                data = 'D %s' % self.connections[c]['name']
                # Broadcast the incoming message to others.  Loop through all existing connection threads.
                for thr in threading.enumerate():
                    # This "if" ensures that only communication threads are processed, not the main program thread
                    if isinstance(thr, connectionThread):
                        try:
                            # pass the message to other threads.
                            thr.BroadcastMessage('%s' % data, self.connections[thr.address]['name'], self.connections[thr.address]['dbHost'], self.connections[thr.address]['dbName'])
                        except:
                            if DEBUG and DEBUG2:
                                print "Exception within a socket error"
                            continue
            except:
                pass  # forget about sending the message if it won't be sent.

            # ... delete the connetion record
            del(self.connections[c])

        if DEBUG and DEBUG2:
            print

        # Create a one-off Timer to call this method again at the appropriate interval
        t = threading.Timer(checkInterval, self.Report)
        # Start the timer
        t.start()

    def KillAllThreads(self):
        """ Kill all Connection and all Listener threads """
        for thr in threading.enumerate():
            if isinstance(thr, connectionThread):
                # Tell the threads they should kill themselves when they get their next message
                thr.CommitSuicide()
            if isinstance(thr, listenerThread):
                # Tell the threads they should kill themselves when they get their next message
                thr.CommitSuicide()
        
if RUNASWINSERVICE:

    # See Chapter 18 of Hammond and Robinson's "Python Programming on Win32"

    class TransanaMessageService(win32serviceutil.ServiceFramework):
        """ A Class designed to manage the Transana Message Server as a Windows Service """
        _svc_name_ = "TransanaMessageServer"
        _svc_display_name_ = "Transana Message Server, v. %.2f" % (float(VERSION) / 100.0)
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

elif RUNASDAEMON:


    # To COMPLETELY remove this daemon on OS X, take the following steps:
    #
    # 1.  Stop the daemon:
    #       sudo python /Applications/TransanaMessageServer/MessageServer.py stop
    # 2.  Delete /Applications/TransanaMessageServer
    # 3.  Delete the LaunchDaemon plist:
    #       sudo rm /Library/LaunchDaemons/org.transana.osx.TransanaMessageServer.261.plist
    # 4.  Have the OS X Package Utility "forget" the Message Server was ever installed:
    #       sudo pkgutil --forget org.transana.TransanaMessageServer
    

    class MyDaemon(daemon.Daemon):
        def run(self):
            while True:
                self.dispatcher = dispatcher()

    daemon = MyDaemon('/tmp/transanamessageserver.pid')
    if len(sys.argv) == 2:
        if 'start' == sys.argv[1]:
            daemon.start()
        elif 'stop' == sys.argv[1]:
            daemon.stop()
        elif 'restart' == sys.argv[1]:
            daemon.restart()
        else:
            print "Unknown command"
            sys.exit(2)
        sys.exit(0)
    else:
        print "usage: %s start|stop|restart" % sys.argv[0]
        sys.exit(2)
    
else:
    # The main action of this program is to launch the Connection Thread Dispatcher
    dispatcher()
