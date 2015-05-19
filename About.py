# -*- coding: cp1252 -*-
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

"""This file implements the About Dialog Box for the Transana application."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx
# import Transana's Constants
import TransanaConstants
# Import Transana's Globals
import TransanaGlobal
# import Python's os module
import os

class AboutBox(wx.Dialog):
    """ Create and display the About Dialog Box """
    def __init__(self):
        """ Create and display the About Dialog Box """
        # Define the initial size
        if '__WXMAC__' in wx.PlatformInfo:
            width = 400
        else:
            width = 370
        height = 425

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Create a Dialog Box
        wx.Dialog.__init__(self, None, -1, _("About Transana"), size=(width, height), style=wx.CAPTION | wx.WANTS_CHARS)
        # OS X requires the Small Window Variant to look right
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)
        # Create a label for the Program Title
        title = wx.StaticText(self, -1, _("Transana"), style=wx.ALIGN_CENTRE)
        # Specify 16pt Bold font
        font = wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD)
        # Apply this font to the Title
        title.SetFont(font)
        # Add the title to the main sizer
        mainSizer.Add(title, 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, 12)

        # Create a label for the Program Version
        version = wx.StaticText(self, -1, _("Version %s") % TransanaConstants.versionNumber, style=wx.ALIGN_CENTRE)
        # Specify 10pt Bold font 
        font = wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD)
        # Apply the font to the Version Label
        version.SetFont(font)
        # Add the version to the main sizer
        mainSizer.Add(version, 0, wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT | wx.RIGHT, 12)

        # Create a label for the Program Copyright
        str = _("Copyright 2002-2012\nThe Board of Regents of the University of Wisconsin System")
        copyright = wx.StaticText(self, -1, str, style=wx.ALIGN_CENTRE)
        # Apply the last specified font (from Program Version) to the copyright label
        font = self.GetFont()
        copyright.SetFont(font)
        # Add the copyright to the main sizer
        mainSizer.Add(copyright, 0, wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT | wx.RIGHT, 12)

        # Create a label for the Program Description, including GNU License information
        self.description_str = _("Transana is written at the \nWisconsin Center for Education Research, \nUniversity of Wisconsin, Madison, and is released with \nno warranty under the GNU General Public License (GPL).  \nFor more information, see http://www.gnu.org.")
        self.description = wx.StaticText(self, -1, self.description_str, style=wx.ALIGN_CENTRE)
        # Add the description to the main sizer
        mainSizer.Add(self.description, 0, wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT | wx.RIGHT, 12)

        # Create a label for the Program Authoring Credits
        self.credits_str = _("Transana was originally conceived and written by Chris Fassnacht.\nCurrent development is being directed by David K. Woods, Ph.D.\nOther contributors include:  Jonathan Beavers, Nate Case, Mark Kim, \nRajas Sambhare and David Mandelin")
        self.credits_str += ".\n" + _("Methodology consultants: Paul Dempster, Chris Thorn,\nand Nicolas Sheon.")
        self.credits = wx.StaticText(self, -1, self.credits_str, style=wx.ALIGN_CENTRE)
        # Add the credits to the main sizer
        mainSizer.Add(self.credits, 0, wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT | wx.RIGHT, 12)

        # Create a label for the Translation Credits
        if TransanaGlobal.configData.language == 'en':
            str = 'Documentation written by David K. Woods\nwith assistance and editing by Becky Holmes.'
        elif TransanaGlobal.configData.language == 'ar':
            str = _("Arabic translation provided by\n.")
        elif TransanaGlobal.configData.language == 'da':
            str = _("Danish translation provided by\nChris Kjeldsen, Aalborg University.")
        elif TransanaGlobal.configData.language == 'de':
            str = _("German translation provided by Tobias Reu, New York University.\n.")
        elif TransanaGlobal.configData.language == 'el':
            str = _("Greek translation provided by\n.")
        elif TransanaGlobal.configData.language == 'es':
            str = _("Spanish translation provided by\nNate Case, UW Madison, and Paco Molinero, Barcelone, Spain")
        elif TransanaGlobal.configData.language == 'fi':
            str = _("Finnish translation provided by\nMiia Collanus, University of Helsinki, Finland")
        elif TransanaGlobal.configData.language == 'fr':
            str = _("French translation provided by\nDr. Margot Kaszap, Ph. D., Universite Laval, Quebec, Canada.")
        elif TransanaGlobal.configData.language == 'he':
            str = _("Hebrew translation provided by\n.")
        elif TransanaGlobal.configData.language == 'it':
            str = _("Italian translation provided by\nFabio Malfatti, Centro Ricerche EtnoAntropologiche www.creasiena.it\nPeri Weingrad, University of Michigan & Terenziano Speranza, Rome, Italy\nDorian Soru, Ph. D. e ICLab, www.iclab.eu, Padova, Italia")
        elif TransanaGlobal.configData.language == 'nl':
            str = _("Dutch translation provided by\nFleur van der Houwen.")
        elif TransanaGlobal.configData.language in ['nb', 'nn']:
            str = _("Norwegian translations provided by\nDr. Dan Yngve Jacobsen, Department of Psychology,\nNorwegian University of Science and Technology, Trondheim")
        elif TransanaGlobal.configData.language == 'pl':
            str = _("Polish translation provided by\n.")
        elif TransanaGlobal.configData.language == 'pt':
            str = _("Portuguese translation provided by\n.")
        elif TransanaGlobal.configData.language == 'ru':
            str = _("Russian translation provided by\nViktor Ignatjev.")
        elif TransanaGlobal.configData.language == 'sv':
            str = _("Swedish translation provided by\nJohan Gille, Stockholm University, Sweden")
        elif TransanaGlobal.configData.language == 'zh':
            str = _("Chinese translation provided by\nZhong Hongquan, Beijin Poweron Technology Co.Ltd.,\nmaintained by Bei Zhang, University of Wisconsin.")
        translations = wx.StaticText(self, -1, str, style=wx.ALIGN_CENTRE)
        # Add the transcription credit to the main sizer
        mainSizer.Add(translations, 0, wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT | wx.RIGHT, 12)

        # Create a label for the FFmpeg Credits
        self.ffmpeg_str = _("This software uses libraries from the FFmpeg project\nunder the LGPLv2.1 or GNU-GPL.  Transana's copyright\ndoes not extend to the FFmpeg libraries or code.\nPlease see http://www.ffmpeg.com.")
        ffmpeg = wx.StaticText(self, -1, self.ffmpeg_str, style=wx.ALIGN_CENTRE)
        # Add the FFmpeg Credits to the main sizer
        mainSizer.Add(ffmpeg, 0, wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT | wx.RIGHT, 12)

        # Create an OK button
        btnOK = wx.Button(self, wx.ID_OK, _("OK"))
        # Add the OK button to the main sizer
        mainSizer.Add(btnOK, 0, wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT | wx.RIGHT, 12)

        # Add the Keyboard Event for the Easter Egg screen
        self.Bind(wx.EVT_KEY_UP, self.OnKeyUp)
        btnOK.Bind(wx.EVT_KEY_UP, self.OnKeyUp)

        # Define the form's Sizer
        self.SetSizer(mainSizer)
        # Lay out the form
        self.SetAutoLayout(True)
        self.Layout()
        # Fit the form to the contained controls
        self.Fit()
        # Center on screen
	self.CentreOnScreen()

        # Show the About Dialog Box modally
        val = self.ShowModal()

        # Destroy the Dialog Box when done with it
        self.Destroy()

        # Dialog Box Initialization returns "None" as the result (see wxPython documentation)
        return None

    def OnKeyUp(self, event):
        """ OnKeyUp event captures the release of keys so they can be processed """
        # If ALT and SHIFT are pressed ...
        if event.AltDown() and event.ShiftDown():
            # ... get the key that was pressed
            key = event.GetKeyCode()
            # If F11 is pressed, show COMPONENT VERSION information
            if key == wx.WXK_F11:
                # Import Python's ctypes, Transana's DBInterface, MySQLdb, and Python's sys modules
                import Crypto, ctypes, DBInterface, MySQLdb, paramiko, sys
                # Build a string that contains the version information for crucial programming components
                str = '\n            Transana %s uses the following tools:\n\n'% (TransanaConstants.versionNumber)
                str = '%sPython:  %s\n' % (str, sys.version_info)
                if 'unicode' in wx.PlatformInfo:
                    str2 = 'unicode'
                else:
                    str2 = 'ansi'
                str = '%swxPython:  %s - %s\n' % (str, wx.VERSION_STRING, str2)
                str = '%sMySQL for Python:  %s\n' % (str, MySQLdb.__version__)
                if DBInterface._dbref != None:
                    # Get a Database Cursor
                    dbCursor = DBInterface._dbref.cursor()
                    # Query the Database about what Database Names have been defined
                    dbCursor.execute('SELECT VERSION()')
                    vs = dbCursor.fetchall()
                    for v in vs:
                        str = "%sMySQL:  %s\n" % (str, v[0])
                str = "%sctypes:  %s\n" % (str, ctypes.__version__)
                str = "%sCrypto:  %s\n" % (str, Crypto.__version__)
                str = "%sparamiko:  %s\n" % (str, paramiko.__version__)
                str = "%sEncoding:  %s\n" % (str, TransanaGlobal.encoding)
                str = "%sLanguage:  %s\n" % (str, TransanaGlobal.configData.language)
                # Replace the Description text with the version information text
                self.description.SetLabel(str)
                # Eliminate the Credits text
                self.credits.SetLabel('')
            # If F12 is pressed ...
            elif key == wx.WXK_F12:
                # Replace the Version information text with the original description text
                self.description.SetLabel(self.description_str)
                # Replace the blank credits text with the original credits text
                self.credits.SetLabel(self.credits_str)
            # Fit the window to the altered controls
            self.Fit()
        # If ALT and SHIFT aren't both pressed ...
        else:
            # ... then we don't do anything
            pass
