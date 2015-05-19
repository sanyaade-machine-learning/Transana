# -*- coding: cp1252 -*-
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

"""This file implements the About Dialog Box for the Transana application."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

import wx
import TransanaConstants
import TransanaGlobal
import os

class AboutBox(wx.Dialog):
    """ Create and display the About Dialog Box """
    def __init__(self):
        """ Create and display the About Dialog Box """
        if '__WXMAC__' in wx.PlatformInfo:
            width = 400
        else:
            width = 370
        height = 350
        # Create a Dialog Box, fixed size
        dlg = wx.Dialog(None, -1, _("About Transana"), size=(width, height), style=wx.CAPTION | wx.WANTS_CHARS)
        if "__WXMAC__" in wx.PlatformInfo:
            dlg.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)
        lay = wx.LayoutConstraints()
        lay.top.SameAs(dlg, wx.Top, 10)
        lay.centreX.SameAs(dlg, wx.CentreX)
        lay.height.Absolute(20)
        lay.width.Absolute(200)
        # Create a label for the Program Title
        title = wx.StaticText(dlg, -1, _("Transana"), style=wx.ALIGN_CENTRE)
        # Specify 16pt Bold font
        font = wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD)
        # Apply this font to the Title
        title.SetFont(font)
        title.SetConstraints(lay)

# This looks like crap.  Forget it for now.
        # Load the Transana icon
#        iconImage = wx.Image('images' + os.sep + 'Transana.ico', wx.BITMAP_TYPE_ICO).ConvertToBitmap()
        # Display the Transana Icon on the screen
#        icon = wx.StaticBitmap(dlg, -1, iconImage, pos=(40, 25), size=(iconImage.GetWidth(), iconImage.GetHeight()))
#        icon2= wx.StaticBitmap(dlg, -1, iconImage, pos=(width - iconImage.GetWidth() - 40, 25), size=(iconImage.GetWidth(), iconImage.GetHeight()))
        
        lay = wx.LayoutConstraints()
        lay.top.Below(title, 10)
        lay.centreX.SameAs(dlg, wx.CentreX)
        lay.height.AsIs()
        lay.width.AsIs()
        # Create a label for the Program Version
        version = wx.StaticText(dlg, -1, _("Version %s") % TransanaConstants.versionNumber, style=wx.ALIGN_CENTRE)
        # Specify 10pt Bold font 
        font = wx.Font(10, wx.SWISS, wx.NORMAL, wx.BOLD)
        # Apply the font to the Version Label
        version.SetFont(font)
        version.SetConstraints(lay)

        # Create a label for the Program Copyright
        str = _("Copyright 2002-2006\nThe Board of Regents of the University of Wisconsin System")
        lay = wx.LayoutConstraints()
        lay.top.Below(version, 10)
        lay.centreX.SameAs(dlg, wx.CentreX)
        lay.height.AsIs()
        lay.width.AsIs()
        copyright = wx.StaticText(dlg, -1, str, style=wx.ALIGN_CENTRE)
        # Apply the last specified font (from Program Version) to the copyright label
        font = dlg.GetFont()
        copyright.SetFont(font)
        copyright.SetConstraints(lay)

        # Create a label for the Program Description, including GNU License information
        self.description_str = _("Transana is free software written at the \nWisconsin Center for Education Research, \nUniversity of Wisconsin, Madison, and is released with \nno warranty under the GNU General Public License (GPL).  \nFor more information, see http://www.gnu.org.")
        lay = wx.LayoutConstraints()
        lay.top.Below(copyright, 10)
        lay.centreX.SameAs(dlg, wx.CentreX)
        lay.height.AsIs()
        lay.width.AsIs()
        self.description = wx.StaticText(dlg, -1, self.description_str, style=wx.ALIGN_CENTRE)
        self.description.SetConstraints(lay)

        # Create a label for the Program Authoring Credits
        self.credits_str = _("Transana was originally conceived and written by Chris Fassnacht.\nCurrent development is being directed by David K. Woods, Ph.D.\nOther contributors include:  Jonathan Beavers, Nate Case, Mark Kim, \nRajas Sambhare and David Mandelin")
        lay = wx.LayoutConstraints()
        lay.top.Below(self.description, 10)
        lay.centreX.SameAs(dlg, wx.CentreX)
        lay.height.AsIs()
        lay.width.AsIs()
        self.credits = wx.StaticText(dlg, -1, self.credits_str, style=wx.ALIGN_CENTRE)
        self.credits.SetConstraints(lay)

        # Create a label for the Translation Credits
        if TransanaGlobal.configData.language == 'en':
            str = 'Documentation written by David K. Woods\nwith assistance and editing by Becky Holmes.'
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
        elif TransanaGlobal.configData.language == 'it':
            str = _("Italian translation provided by\nPeri Weingrad, University of Michigan & Terenziano Speranza, Rome, Italy")
        elif TransanaGlobal.configData.language == 'nl':
            str = _("Dutch translation provided by\nFleur van der Houwen.")
        elif TransanaGlobal.configData.language == 'pl':
            str = _("Polish translation provided by\n.")
        elif TransanaGlobal.configData.language == 'ru':
            str = _("Russian translation provided by\nViktor Ignatjev.")
        elif TransanaGlobal.configData.language == 'sv':
            str = _("Swedish translation provided by\nJohan Gille, Stockholm University, Sweden")
        lay = wx.LayoutConstraints()
        lay.top.Below(self.credits, 10)
        lay.centreX.SameAs(dlg, wx.CentreX)
        lay.height.AsIs()
        lay.width.AsIs()
        translations = wx.StaticText(dlg, -1, str, style=wx.ALIGN_CENTRE)
        translations.SetConstraints(lay)

        # Create an OK button
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(dlg, wx.Bottom, 10)
        lay.centreX.SameAs(dlg, wx.CentreX)
        lay.height.AsIs()
        lay.width.AsIs()
        btnOK = wx.Button(dlg, wx.ID_OK, _("OK"))
        btnOK.SetConstraints(lay)

        # Add the Keyboard Event for the Easter Egg screen
        dlg.Bind(wx.EVT_KEY_UP, self.OnKeyUp)
        btnOK.Bind(wx.EVT_KEY_UP, self.OnKeyUp)

        dlg.Layout()
	dlg.CentreOnScreen()

        # Show the About Dialog Box modally
        val = dlg.ShowModal()

        # Destroy the Dialog Box when done with it
        dlg.Destroy()

        # Dialog Box Initialization returns "None" as the result (see wxPython documentation)
        return None

    def OnKeyUp(self, event):
        """ OnKeyUp event captures the release of keys so they can be processed """
        if event.AltDown() and event.ShiftDown():
            key = event.GetKeyCode()
            if key == wx.WXK_F11:
                import ctypes, DBInterface, MySQLdb, sys
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
                str = "%sEncoding:  %s\n" % (str, TransanaGlobal.encoding)
                str = "%sLanguage:  %s\n" % (str, TransanaGlobal.configData.language)
                self.description.SetLabel(str)
                self.credits.SetLabel('')
            elif key == wx.WXK_F12:
                self.description.SetLabel(self.description_str)
                self.credits.SetLabel(self.credits_str)
        else:
            pass
