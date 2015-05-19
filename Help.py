# Copyright (C) 2003 - 2012 The Board of Regents of the University of Wisconsin System 
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

"""This file handles the Transana Help System.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "Help.py DEBUG is ON!!"

# import wxPython
import wx
# import the wxPython html module
import wx.html
# import Python's os and sys modules
import os, sys
# import Python's pickle module
import pickle


class Help(object):
    """ This class implements Help Calls to the Transana Manual """
    # Show the Manual's Welcome page if no context is supplied
    def __init__(self, HelpContext='Welcome'):
        self.help = wx.html.HtmlHelpController()
        
        # This has emerged as the "preferred" method on the wxPython-users list.
        programDir = os.path.abspath(sys.path[0])
        # Okay, that doesn't work with wxversion, which adds to the path.  Here's the fix, I hope.
        # This should over-ride the programDir with the first value that contains Transana in the path.
        for path in sys.path:
            if 'transana' in path.lower():
                programDir = path
                break
        if os.path.isfile(programDir):
            programDir = os.path.dirname(programDir)

        if DEBUG:
            msg = "Help.__init__():  programDir = %s" % programDir
            tmpDlg = wx.MessageDialog(None, msg)
            tmpDlg.ShowModal()
            tmpDlg.Destroy()

        # There's a problem with wxPython 2.8.3.0 that causes the Help Window to close when a Search is performed.
        # Robin Dunn, author of wxPython, writes:
        #
        #    The frame isn't created right away, so you need to do it later.
        #
        #    The problem I worked around is the fact that the frame used by the 
        #    HtmlHelpController is set to not prevent the exit of MainLoop, so
        #    if some other top level window closes (such as the search dialog)
        #    then there is no other TLW left besides the help frame, so the App
        #    will exit MainLoop.
        #
        #    So if you are using the HtmlHelpController in a situation where there
        #    are no other TLW's in the same app then you'll have the same problem.
        #    I worked around it in wx.tools.helpviewer by making a frame that is
        #    never shown and uses the help frame as it's parent so it will get
        #    destroyed when the help frame closes.  You can also call 
        #    wx.GetApp().SetExitOnFrameDelete(False) but then you'll have to worry 
        #    about exiting MainLoop yourself.

        # The following comments and line of code, plus the method it calls, are
        # based on Robin's code for implementing this.
        
        # The frame used by the HtmlHelpController is set to not prevent
        # app exit, so in the case of a standalone helpviewer like this
        # when the about box or search box is closed the help frame will
        # be the only one left and the app will close unexpectedly.  To
        # work around this we'll create another frame that is never shown,
        # but which will be closed when the helpviewer frame is closed.
        wx.CallAfter(self.makeOtherFrame, self.help)

        # Now we add the actual contents to the HelpCtrl
        self.help.AddBook(os.path.join(programDir, 'help', 'Manual.hhp'))
        self.help.AddBook(os.path.join(programDir, 'help', 'Tutorial.hhp'))
        self.help.AddBook(os.path.join(programDir, 'help', 'TranscriptNotation.hhp'))
        # And finally, we display the Help Control, showing the current contents.
        self.help.Display(HelpContext)

    def makeOtherFrame(self, helpctrl):
        """ See long comment above, with the CallAfter line.  This is code from Robin Dunn. """
        # Get the control's Frame
        parent = helpctrl.GetFrame()
        # Create another frame with the HelpCtrl as the parent.  This provides a Top Level
        # Window that is never shown but still prevents the program from exiting when the
        # Search Dialog is closed.  Because it has the HelpCtrl as a parent, it will close
        # automatically.
        otherFrame = wx.Frame(parent)

        # Now that we have access to the Frame (because this is in a CallAfter situation),
        # we can change the size of the Help Window.  Yea!
        if parent != None:
            # Get the size of the screen
            rect = wx.Display(0).GetClientArea()  # wx.ClientDisplayRect()
            # Set the top to 10% of the screen height
            top = int(0.1 * rect[3])
            # Set the height so that the top and bottom will have equal boundaries
            height = rect[3] - (2 * top)
            # Set the window width to 80% of the screen or 1000 pixels, whichever is smaller
            width = min(int(rect[2] * 0.8), 1000)
            # Set the left margin so that the window is centered
            left = int(rect[2] / 2) - int(width / 2)
            # Position the window based on the calculated position
            parent.SetPosition(wx.Point(left, top))
            # Size the window based on the calculated size
            parent.SetSize(wx.Size(width, height))

# We need the Help Control to run stand-alone
if __name__ == '__main__':
    # Import Python's sys module
    import sys

    class MyApp(wx.App):
        """ Application class for the Transana Help application """
        def OnInit(self):
            # Initialize the Help context
            helpContext = ''
            
            # NOTE:  The Mac lacks the capacity for command line parameters, so we pass the help context via
            #        a pickled string on the Mac and by command line elsewhere.

            # If we're on a Mac ...
            if "__WXMAC__" in wx.PlatformInfo:
                # ... start by looking for the pickle file containing the help context we should look for.
                if os.path.exists(os.getenv("HOME") + '/TransanaHelpContext.txt'):
                    # If the file exists, open it.
                    helpfile = open(os.getenv("HOME") + '/TransanaHelpContext.txt', 'r')
                    # use pickle to read the pickled contents of the file
                    helpContext = pickle.load(helpfile)
                    # Close the help file.
                    helpfile.close()
                    # Once we've read it, we can delete the file!
                    os.remove(os.getenv("HOME") + '/TransanaHelpContext.txt')
            # If we're NOT on a Mac ...
            else:
                # ... iterate through the command line parameters.  (Multiple word help contexts are passed as multiple parameters.)
                for paramCount in range(1, len(sys.argv)):
                    # Build the help context from the parameters
                    helpContext = helpContext + sys.argv[paramCount] + ' '
            # If no help context is provided ...
            if helpContext == '':
                # ... use the default Help Context
                helpContext = 'Welcome'
            # Create the frame, passing the help context
            self.frame = Help(helpContext.strip())
            return True

    # Define the Application Object
    app = MyApp(0)
    # Run the Application's Main Loop
    app.MainLoop()
