# Copyright (C) 2003 The Board of Regents of the University of Wisconsin System 
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

"""This module contains Transana's global variables."""

__author__ = 'Nathaniel Case <nacase@wisc.edu>, David Woods <dwoods@wcer.wisc.edu>'

# Import wxPython
import wx
# import Transana's ConfigData
import ConfigData
# import Python's os and sys modules
import os
import sys

# We need to know what directory the program is running from.  We use this in several
# places in the program to be able to find things like images and help files.
#programDir, programName = os.path.split(sys.argv[0])
#if programDir == '':
#    programDir = os.getcwd()

# This has emerged as the "preferred" method on the wxPython-users list.
programDir = os.path.abspath(sys.path[0])
if os.path.isfile(programDir):
    programDir = os.path.dirname(programDir)

# Here's another possibility, described as "canonical" by the author.
# Bob Ippolito, 11/18/2004 Pythonmac-SIG "The where's-my-app problem"
# "For applications frozen with py2exe, there IS effectively no on-disk main script (the bytecode is  
#  shoved into the executable).  Finding your main script is not a useful thing to do from a packaged
#  application anyway, so only use this when find_packager() returns None."

# import __main__
# print os.path.dirname(os.path.abspath(__main__.__file__))



# Determine the height of the Menu Window.  Because the Mac has a common menu bar for all
# applications, we set it to 0 on the Mac!
if wx.Platform == '__WXMAC__':
    menuHeight = 0
else:
    # While we default to 44, this value actually can get altered elsewhere to reflect the height of
    # the title/header bar.  XP using Large Fonts, for example, needs a larger value.
    menuHeight = 44

# Menu Window, defined in Transana.py
menuWindow = None

# Prepare the wxPrintData object for use in Printing
printData = wx.PrintData()
# wxPython defaults to A4 paper.  Transana should default to Letter
printData.SetPaperId(wx.PAPER_LETTER)

# Create a Configuration Data Object.  This automatically load previously saved configuration information
configData = ConfigData.ConfigData()
