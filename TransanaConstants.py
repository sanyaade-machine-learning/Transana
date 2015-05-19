# Copyright (C) 2003 - 2006 The Board of Regents of the University of Wisconsin System 
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

# Use the New Video Player on Mac, but not on Windows yet.
NEWVIDEOPLAYER = ('wxMac' in wx.PlatformInfo)

# import the Python string module
import string

# Define all global program Constants here

# Define a Boolean to indicate Single- or Multi- user
# NOTE:  When you change this value, you MUST change the MySQL for Python installation you are using
#        to match.
singleUserVersion = True

# Program Version Number
versionNumber = '2.11'
# Modify for Multi-user if appropriate
if not singleUserVersion:
    versionNumber = versionNumber + '-MU'
# Add testing version information if appropriate.  (Set to "''" if not!)
# NOTE:  This will differ by Platform for a little while.
if '__WXMAC__' in wx.PlatformInfo:
    versionNumber = versionNumber +  '-Mac Alpha 1.15'
else:
    versionNumber = versionNumber + '-Win'

# Define the Timecode Character
if 'unicode' in wx.PlatformInfo:
    if 'wxMac' in wx.PlatformInfo:
        TIMECODE_CHAR = unicode('\xc2\xa7', 'utf-8')
    else:
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
if NEWVIDEOPLAYER:
    import wx.media
    MEDIA_PLAYSTATE_NONE               =  None
    MEDIA_PLAYSTATE_STOP               =  wx.media.MEDIASTATE_STOPPED
    MEDIA_PLAYSTATE_PAUSE              =  wx.media.MEDIASTATE_PAUSED
    MEDIA_PLAYSTATE_PLAY               =  wx.media.MEDIASTATE_PLAYING
elif "__WXMSW__" in wx.PlatformInfo:
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
                 
legalFilenameCharacters = string.ascii_letters + string.digits + ":. -_$&@!%(){}[]~'#^+=/" 
if "__WXMSW__" in wx.PlatformInfo:
    legalFilenameCharacters += '\\'

chineseEncoding = 'gbk'

# We want enough colors, but not too many.  This list seems about right to me.  I doubt my color names are standard.
# But then, I'm often perplexed by the colors that are included and excluded by most programs.  (Excel for example.)
# Each entry is made up of a color name and a tuple of the RGB values for the color.
transana_colorList = [('Black',       (  0,   0,   0)),
                      ('Dark Blue',   (  0,   0, 128)),
                      ('Blue',        (  0,   0, 255)),
                      ('Light Blue',  (  0, 128, 255)),
                      ('Cyan',        (  0, 255, 255)),
                      ('Light Aqua',  (128, 255, 255)),
                      ('Blue Green',  (  0, 128, 128)),
                      ('Dark Green',  (  0, 128,   0)),
                      ('Green Blue',  (  0, 255, 128)),
                      ('Green',       (  0, 255,   0)),
                      ('Chartreuse',  (128, 255,   0)),
                      ('Light Green', (128, 255, 128)),
                      ('Olive',       (128, 128,   0)),
                      ('Gray',        (128, 128, 128)),
                      ('Lavendar',    (128, 128, 255)),
                      ('Purple',      (128,   0, 255)),
                      ('Dark Purple', (128,   0, 128)),
                      ('Maroon',      (128,   0,   0)),
                      ('Magenta',     (255,   0, 255)),
                      ('Light Fuscia',(255, 128, 255)),
                      ('Rose',        (255,   0, 128)),
                      ('Red',         (255,   0,   0)),
                      ('Salmon',      (255, 128, 128)),
                      ('Orange',      (255, 128,   0)),
                      ('Yellow',      (255, 255,   0)),  
                      ('Light Yellow',(255, 255, 128)),  
                      ('White',       (255, 255, 255))]
# The following exists only to ensure that the color names are available for translation.
# (I had to take the translation code out of the above data structure, as color names were only showing up in
#  the initial language.)
tmpColorList = (_('Black'), _('Dark Blue'), _('Blue'), _('Light Blue'), _('Cyan'), _('Light Aqua'), _('Green Blue'),
                 _('Dark Green'), _('Blue Green'),_('Green'), _('Chartreuse'), _('Light Green'), _('Olive'), _('Gray'),
                _('Lavendar'), _('Purple'), _('Dark Purple'), _('Maroon'), _('Magenta'), _('Light Fuscia'), _('Rose'),
                _('Red'), _('Salmon'), _('Orange'), _('Yellow'), _('Light Yellow'), _('White'))
transana_colorNameList = []
transana_colorLookup = {}
for (colorName, colorDef) in transana_colorList:
    transana_colorNameList.append(colorName)
    transana_colorLookup[colorName] = colorDef
