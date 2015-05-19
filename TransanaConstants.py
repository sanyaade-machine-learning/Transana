# Copyright (C) 2003 - 2010 The Board of Regents of the University of Wisconsin System 
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

"""This module contains Transana's global variables and constants."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare'

# import wxPython
import wx

# import the Python string module
import string


# Define all global program Variables and Constants here

import TransanaConfigConstants
# Define a Boolean to indicate Single- or Multi- user
# NOTE:  When you change this value, you MUST change the MySQL for Python installation you are using
#        to match.
singleUserVersion = TransanaConfigConstants.singleUserVersion
# Indicate if this is the Lab version
labVersion = TransanaConfigConstants.labVersion
# Set this flag to "True" to create the Demonstration version.  (But don't mix this with MU!)
demoVersion = TransanaConfigConstants.demoVersion
workshopVersion = TransanaConfigConstants.workshopVersion
if workshopVersion:
    startdate = TransanaConfigConstants.stdt
    expirationdate = TransanaConfigConstants.xpdt

# Program Version Number
versionNumber = '2.43a1'
# Build Number
buildNumber = '243a1'
# Modify for Multi-user if appropriate
if not singleUserVersion:
    versionNumber = versionNumber + '-MU'
# Add testing version information if appropriate.  (Set to "''" if not!)
# NOTE:  This will differ by Platform for a little while.
if '__WXMAC__' in wx.PlatformInfo:
    versionNumber = versionNumber +  '-Mac'
elif 'wxMSW' in wx.PlatformInfo:
    versionNumber = versionNumber + '-Win'
elif 'wxGTK' in wx.PlatformInfo:
    versionNumber += '-Linux Alpha 1.0'
else:
    versionNumber += '-Unknown Platform Alpha 1.0'

if labVersion:
    versionNumber += ' Lab'
    
# Define limits for the Demonstration Version
if demoVersion:
    versionNumber += ' Demonstration'
    maxEpisodes = 5
    maxEpisodeTranscripts = 5
    maxClips = 30
    maxKeywords = 15

# If this is the Workshop Version ...
if workshopVersion:
    # ... note that on the Version Number
    versionNumber += ' Workshop'

# Allow Drag and Drop on the Mac?  (There's a bug, but it could be fixed!)
macDragDrop = True
# Define the maximum number of Trasncript Windows that can be opened
maxTranscriptWindows = 5

# Define the Timecode Character
if 'unicode' in wx.PlatformInfo:
    TIMECODE_CHAR = unicode('\xc2\xa4', 'utf-8')
else:
    TIMECODE_CHAR = '\xa4'

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
import wx.media
MEDIA_PLAYSTATE_NONE               =  None
MEDIA_PLAYSTATE_STOP               =  wx.media.MEDIASTATE_STOPPED
MEDIA_PLAYSTATE_PAUSE              =  wx.media.MEDIASTATE_PAUSED
MEDIA_PLAYSTATE_PLAY               =  wx.media.MEDIASTATE_PLAYING

# Define the Media File type strings to be used throughout Transana
fileTypesString = _("""All files (*.*)|*.*|All supported media files (*.mpg, *.avi, *.mov, *.mp4, *.m4v, *.wmv, *.mp3, *.wav, *.wma, *.aac)|*.mpg;*.avi;*.mov;*.mp4;*.m4v;*.wmv;*.mp3;*.wav;*.wma;*.aac|All video files (*.mpg, *.avi, *.mov, *.mp4, *.m4v, *.wmv)|*.mpg;*.mpeg;*.avi;*.mov;*.mp4;*.m4v;*.wmv|All audio files (*.mp3, *.wav, *.wma, *.aac, *.au, *.snd)|*.mp3;*.wav;*.wma;*.aac;*.au;*.snd|MPEG files (*.mpg)|*.mpg;*.mpeg|AVI files (*.avi)|*.avi|QuickTime files (*.mov, *.mp4, *.m4v)|*.mov;*.mp4;*.m4v|Windows Media Video (*.wmv)|*wmv|MP3 files (*.mp3)|*.mp3|WAV files (*.wav)|*.wav|Windows Media Audio (*.wma)|*.wma|AAC Audio (*.aac)|*.aac""")
fileTypesList = [_("All files (*.*)"),
                 _("All supported media files (*.mpg, *.avi, *.mov, *.mp4, *.m4v, *.wmv, *.mp3, *.wav, *.wma, *.aac)"),
                 _("All video files (*.mpg, *.avi, *.mov, *.mp4, *.m4v, *.wmv)"),
                 _("All audio files (*.mp3, *.wav, *.wma, *.aac, *.au, *.snd)"),
                 _("MPEG files (*.mpg, *.mpeg)"),
                 _("AVI files (*.avi)"),
                 _("QuickTime files (*.mov, *.mp4, *.m4v)"),
                 _("Windows Media Video files (*.wmv)"),
                 _("MP3 files (*.mp3)"),
                 _("WAV files (*.wav)"),
                 _("Windows Media Audio files (*.wma)"),
                 _("AAC files (*.aac)"),
                 _("Rich Text Format files (*.rtf)"),
                 _("BMP, PNG, and WAV files (*.bmp, *.png, *.wav)")]
mediaFileTypes = ['mpg', 'mpeg', 'avi', 'mov', 'mp4', 'm4v', 'wmv', 'mp3', 'wav', 'wma', 'aac']

# We need to know what characters are legal in a file name!
legalFilenameCharacters = string.ascii_letters + string.digits + ":. -_$&@!%(){}[]~'#^+=/" 
if "__WXMSW__" in wx.PlatformInfo:
    legalFilenameCharacters += '\\'

# There are several different encodings that could be used with Chinese.  "gbk" seems to work bests
chineseEncoding = 'gbk'
