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

"""This module contains Transana's global constants."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare'

# import wxPython
import wx
# import the Python string module
import string

# Define all global program Constants here

# Define a Boolean to indicate Single- or Multi- user
# NOTE:  When you change this value, you MUST change the MySQL for Python installation you are using
#        to match.
singleUserVersion = True

# Program Version Number
versionNumber = '2.05'
# Modify for Multi-user if appropriate
if not singleUserVersion:
    versionNumber = versionNumber + '-MU'
# Add testing version information if appropriate.  (Set to "''" if not!)
# NOTE:  This will differ by Platform for a little while.
if '__WXMAC__' in wx.PlatformInfo:
    versionNumber = versionNumber +  '-Mac Alpha 1.02'
else:
    versionNumber = versionNumber + '-Win'

# IDs for the Visualization Window
VISUAL_BUTTON_ZOOMIN            =  wx.NewId()
VISUAL_BUTTON_ZOOMOUT           =  wx.NewId()
VISUAL_BUTTON_ZOOM100           =  wx.NewId()
VISUAL_BUTTON_CURRENT           =  wx.NewId()
VISUAL_BUTTON_SELECTED          =  wx.NewId()

# IDs for the Data Right-click Menus are automatically generated based on this seed.
# For the moment, we will leave 800 - 999 open for them
DATA_MENU_CMD_OFSET             =  800


# Define the IDs of wxPython GUI objects that need them

# Define the IDs for the Video Unit
CONTROL_READYTOPLAY             = wx.NewId()    # This identifies a timer used in the Windows-specific Video unit
CONTROL_POSITIONAFTERLOADING    = wx.NewId()    # This identifies a timer used in the Windows-specific Video unit
CONTROL_PROGRESSNOTIFICATION    = wx.NewId()    # This identifies a timer used in the Windows-specific Video unit


# Define Media Constants needed for inter-object communication
# NOTE:  These constants are different for different media players.
if "__WXMSW__" in wx.PlatformInfo:
    MEDIA_PLAYSTATE_NONE               = -1
    MEDIA_PLAYSTATE_STOP               =  0
    MEDIA_PLAYSTATE_PAUSE              =  1
    MEDIA_PLAYSTATE_PLAY               =  2
elif "__WXMAC__" in wx.PlatformInfo:
    MEDIA_PLAYSTATE_NONE               = -1
    MEDIA_PLAYSTATE_STOP               =  1
    MEDIA_PLAYSTATE_PAUSE              =  2
    MEDIA_PLAYSTATE_PLAY               =  0


fileTypesString = _("""All files (*.*)|*.*|All video files (*.mpg, *.avi)|*.mpg;*.mpeg;*.avi|All audio files (*.mp3, *.wav, *.au, *.snd)|*.mp3;*.wav;*.au;*.snd|MPEG files (*.mpg)|*.mpg;*.mpeg|AVI files (*.avi)|*.avi|MP3 files (*.mp3)|*.mp3|WAV files (*.wav)|*.wav""")

fileTypesList = [_("All files (*.*)"),
                 _("All video files (*.mpg, *.avi)"),
                 _("All audio files (*.mp3, *.wav, *.au, *.snd)"),
                 _("BMP, PNG, and WAV files (*.bmp, *.png, *.wav)"),
                 _("RTF files (*.rtf)"),
                 _("BMP and PNG files (*.bmp, *.png)"),
                 _("WAV files (*.wav)"),
                 _("MPEG files (*.mpg, *.mpeg)"),
                 _("AVI files (*.avi)"),
                 _("MOV files (*.mov)"),
                 _("MP3 files (*.mp3)")]
                 
legalFilenameCharacters = string.ascii_letters + string.digits + ":. -_$&@!%(){}[]~'#^+=" 
if "__WXMAC__" in wx.PlatformInfo:
    legalFilenameCharacters += '/'
else:
    legalFilenameCharacters += '\\'
