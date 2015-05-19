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
        
        # on Mac, GetFrame does not work at the moment!
        if self.help.GetFrame() != None:

            rect = wx.ClientDisplayRect()
            top = 20
            if rect[2] > 1060:  # Screen Width
                left = int((rect[2]-1040) / 2)
                width = 1040
            else:
                left = 0
                width = rect[2]
            height = rect[3] - 40
        
            self.help.GetFrame().SetPosition(wx.Point(left, top))
            self.help.GetFrame().SetSize(wx.Size(width, height))

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

        self.help.AddBook(os.path.join(programDir, 'help', 'Manual.hhp'))
        self.help.AddBook(os.path.join(programDir, 'help', 'Tutorial.hhp'))
        self.help.AddBook(os.path.join(programDir, 'help', 'TranscriptNotation.hhp'))
        self.help.AddBook(os.path.join(programDir, 'help', 'FileManagement.hhp'))
        self.help.Display(HelpContext)

        
if __name__ == '__main__':

    import sys

    class MyApp(wx.App):
        def OnInit(self):
            helpContext = ''
            
            # NOTE:  The Mac lacks the capacity for command line parameters, so we pass the help context via
            #        a pickled string on the Mac and by command line elsewhere.
            if "__WXMAC__" in wx.PlatformInfo:
                
                if os.path.exists(os.getenv("HOME") + '/TransanaHelpContext.txt'):
                    helpfile = open(os.getenv("HOME") + '/TransanaHelpContext.txt', 'r')
                    helpContext = pickle.load(helpfile)
                    helpfile.close()
                    os.remove(os.getenv("HOME") + '/TransanaHelpContext.txt')
                else:
                    helpContext = 'Welcome'
            else:
                for paramCount in range(1, len(sys.argv)):
                    helpContext = helpContext + sys.argv[paramCount] + ' '
            self.frame = Help(helpContext.strip())
            return True

    app = MyApp(0)
    app.MainLoop()
