# Copyright (C) 2003 - 2015 The Board of Regents of the University of Wisconsin System 
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

"""This file implements the OptionsSettings class, which defines the
   main Options > Settings Dialog Box and values. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare'

# import wxPython
import wx
# import the Styled Text Ctrl, to get some constants
from wx import stc
# import the Python os module
import os
# import the Python sys module
import sys
# import Transana's Database Interface
import DBInterface
# import Transana Dialogs
import Dialogs
# import Transana's Constants
import TransanaConstants
# Import the Transana Global Variables
import TransanaGlobal


class OptionsSettings(wx.Dialog):
    """ Options > Settings Dialog Box """

    def __init__(self, parent, tabToShow=0, lab=False):
        """ Initialize the Program Options Dialog Box """
        self.parent = parent
        self.lab = lab
        dlgWidth = 600
        # The Configuration Options dialog needs to be a different on different platforms, and
        # if we're showing the LAB version's initial configuration, we need a bit more room.
        if 'wxMSW' in wx.PlatformInfo:
            if self.lab:
                dlgHeight = 445
            else:
                dlgHeight = 445
        else:
            if self.lab:
                dlgHeight = 395
            else:
                dlgHeight = 345
        # Define the Dialog Box
        wx.Dialog.__init__(self, parent, -1, _("Transana Settings"), wx.DefaultPosition, wx.Size(dlgWidth, dlgHeight), style=wx.CAPTION | wx.SYSTEM_MENU | wx.THICK_FRAME)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the form's main VERTICAL Sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Define a wxNotebook for the Tab structure
        notebook = wx.Notebook(self, -1, size=self.GetSizeTuple())
        # Add the notebook to the main sizer
        mainSizer.Add(notebook, 1, wx.EXPAND | wx.ALL, 3)

        # Define the Directories Tab that goes in the wxNotebook
        panelDirectories = wx.Panel(notebook, -1, size=notebook.GetSizeTuple(), name='OptionsSettings.DirectoriesPanel')

        # Define the main VERTICAL sizer for the Notebook Page
        panelDirSizer = wx.BoxSizer(wx.VERTICAL)

        # The LAB version initial configuration dialog gets some introductory text that can be skipped otherwise.
        if lab:
            # Add the LAB Version configuration instructions Label to the Directories Tab
            instText = _("Transana needs to know where you store your data.  Please identify the location where you store your \nsource media files, where you want Transana to save your waveform data, and where you want your \ndatabase files stored.  ") + '\n\n'
            instText += _("None of this should be on the lab computer, where others may be able to access your confidential data, \nor where data may be deleted over night.")
            lblLabInst = wx.StaticText(panelDirectories, -1, instText, style=wx.ST_NO_AUTORESIZE)
            # Add the label to the Panel Sizer
            panelDirSizer.Add(lblLabInst, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)

        # Add the Media Library Directory Label to the Directories Tab
        lblVideoDirectory = wx.StaticText(panelDirectories, -1, _("Media Library Directory"), style=wx.ST_NO_AUTORESIZE)
        # Add the label to the Panel Sizer
        panelDirSizer.Add(lblVideoDirectory, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
        # Add a spacer
        panelDirSizer.Add((0, 3))

        # Create a Row Sizer
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the Media Library Directory TextCtrl to the Directories Tab
        # If the Media path is not empty, we should normalize the path specification
        if TransanaGlobal.configData.videoPath == '':
            videoPath = TransanaGlobal.configData.videoPath
        else:
            videoPath = os.path.normpath(TransanaGlobal.configData.videoPath)
        self.videoDirectory = wx.TextCtrl(panelDirectories, -1, videoPath)
        # Add the element to the Row Sizer
        r1Sizer.Add(self.videoDirectory, 6, wx.EXPAND | wx.RIGHT, 10)

        # Add the Media Library Directory Browse Button to the Directories Tab
        self.btnVideoBrowse = wx.Button(panelDirectories, -1, _("Browse"))
        # Add the element to the Row Sizer
        r1Sizer.Add(self.btnVideoBrowse, 0)
        wx.EVT_BUTTON(self, self.btnVideoBrowse.GetId(), self.OnBrowse)

        # Add the Row Sizer to the Panel Sizer
        panelDirSizer.Add(r1Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Waveform Directory Label to the Directories Tab
        lblWaveformDirectory = wx.StaticText(panelDirectories, -1, _("Waveform Directory"), style=wx.ST_NO_AUTORESIZE)
        # Add the label to the Panel Sizer
        panelDirSizer.Add(lblWaveformDirectory, 0, wx.LEFT | wx.RIGHT, 10)
        # Add a spacer
        panelDirSizer.Add((0, 3))
        
        # Create a Row Sizer
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the Waveform Directory TextCtrl to the Directories Tab
        # If the visualization path is not empty, we should normalize the path specification
        if TransanaGlobal.configData.visualizationPath == '':
            visualizationPath = TransanaGlobal.configData.visualizationPath
        else:
            visualizationPath = os.path.normpath(TransanaGlobal.configData.visualizationPath)
        self.waveformDirectory = wx.TextCtrl(panelDirectories, -1, visualizationPath)
        # Add the element to the Row Sizer
        r2Sizer.Add(self.waveformDirectory, 6, wx.EXPAND | wx.RIGHT, 10)

        # Add the Waveform Directory Browse Button to the Directories Tab
        self.btnWaveformBrowse = wx.Button(panelDirectories, -1, _("Browse"))
        # Add the element to the Row Sizer
        r2Sizer.Add(self.btnWaveformBrowse, 0)
        wx.EVT_BUTTON(self, self.btnWaveformBrowse.GetId(), self.OnBrowse)

        # Add the Row Sizer to the Panel Sizer
        panelDirSizer.Add(r2Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Database Directory Label to the Directories Tab
        lblDatabaseDirectory = wx.StaticText(panelDirectories, -1, _("Database Directory"), style=wx.ST_NO_AUTORESIZE)
        # Add the element to the Panel Sizer
        panelDirSizer.Add(lblDatabaseDirectory, 0, wx.LEFT | wx.RIGHT, 10)
        # Add a spacer
        panelDirSizer.Add((0, 3))
        
        # Create a Row Sizer
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the Database Directory TextCtrl to the Directories Tab
        # If the database path is not empty, we should normalize the path specification
        if TransanaGlobal.configData.databaseDir == '':
            databaseDir = TransanaGlobal.configData.databaseDir
        else:
            databaseDir = os.path.normpath(TransanaGlobal.configData.databaseDir)
        self.oldDatabaseDir = databaseDir
        self.databaseDirectory = wx.TextCtrl(panelDirectories, -1, databaseDir)
        # Add the element to the Row Sizer
        r3Sizer.Add(self.databaseDirectory, 6, wx.EXPAND | wx.RIGHT, 10)

        # Add the Database Directory Browse Button to the Directories Tab
        self.btnDatabaseBrowse = wx.Button(panelDirectories, -1, _("Browse"))
        # Add the element to the Row Sizer
        r3Sizer.Add(self.btnDatabaseBrowse, 0)
        wx.EVT_BUTTON(self, self.btnDatabaseBrowse.GetId(), self.OnBrowse)

        # Add the Row Sizer to the Panel Sizer
        panelDirSizer.Add(r3Sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # The Database Directory should not be visible for the Multi-user version of the program.
        # Let's just hide it so that the program doesn't crash for being unable to populate the control.
        if not TransanaConstants.singleUserVersion:
            lblDatabaseDirectory.Show(False)
            self.databaseDirectory.Show(False)
            self.btnDatabaseBrowse.Show(False)

        # Tell the Directories Panel to lay out now and do AutoLayout
        panelDirectories.SetSizer(panelDirSizer)
        panelDirectories.SetAutoLayout(True)
        panelDirectories.Layout()

        # If we're not doing the LAB version's initial configuration screen, we allow for a lot more configuration data
        if not self.lab:
            # Add the Transcriber Panel to the Notebook
            panelTranscriber = wx.Panel(notebook, -1, size=notebook.GetSizeTuple(), name='OptionsSettings.TranscriberPanel')

            # Define the main VERTICAL sizer for the Notebook Page
            panelTranSizer = wx.BoxSizer(wx.VERTICAL)

            # Add the Media Setback Label to the Transcriber Settings Tab
            lblTranscriptionSetback = wx.StaticText(panelTranscriber, -1, _("Transcription Setback:  (Auto-rewind interval for Ctrl-S)"),
                                                    style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Panel Sizer
            panelTranSizer.Add(lblTranscriptionSetback, 0, wx.LEFT | wx.TOP, 10)
            # Add a spacer
            panelTranSizer.Add((0, 3))

            # Add the Media Setback Slider to the Transcriber Settings Tab
            self.transcriptionSetback = wx.Slider(panelTranscriber, -1, TransanaGlobal.configData.transcriptionSetback, 0, 5,
                                                  style=wx.SL_HORIZONTAL | wx.SL_AUTOTICKS)
            # Add the element to the Panel Sizer
            panelTranSizer.Add(self.transcriptionSetback, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

            # Create a Row Sizer
            setbackSizer = wx.BoxSizer(wx.HORIZONTAL)
            # Add a spacer so the numbers are positioned correctly
            setbackSizer.Add((11, 0))
            # Add the Media Setback "0" Value Label to the Transcriber Settings Tab
            lblTranscriptionSetbackMin = wx.StaticText(panelTranscriber, -1, "0", style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            setbackSizer.Add(lblTranscriptionSetbackMin, 0)
            # Add a spacer
            setbackSizer.Add((1, 0), 1, wx.EXPAND)

            # Add the Media Setback "1" Value Label to the Transcriber Settings Tab
            lblTranscriptionSetback1 = wx.StaticText(panelTranscriber, -1, "1", style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            setbackSizer.Add(lblTranscriptionSetback1, 0)
            # Add a spacer
            setbackSizer.Add((1, 0), 1, wx.EXPAND)

            # Add the Media Setback "2" Value Label to the Transcriber Settings Tab
            lblTranscriptionSetback2 = wx.StaticText(panelTranscriber, -1, "2", style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            setbackSizer.Add(lblTranscriptionSetback2, 0)
            # Add a spacer
            setbackSizer.Add((1, 0), 1, wx.EXPAND)

            # Add the Media Setback "3" Value Label to the Transcriber Settings Tab
            lblTranscriptionSetback3 = wx.StaticText(panelTranscriber, -1, "3", style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            setbackSizer.Add(lblTranscriptionSetback3, 0)
            # Add a spacer
            setbackSizer.Add((1, 0), 1, wx.EXPAND)

            # Add the Media Setback "4" Value Label to the Transcriber Settings Tab
            lblTranscriptionSetback4 = wx.StaticText(panelTranscriber, -1, "4", style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            setbackSizer.Add(lblTranscriptionSetback4, 0)
            # Add a spacer
            setbackSizer.Add((1, 0), 1, wx.EXPAND)

            # Add the Media Setback "5" Value Label to the Transcriber Settings Tab
            lblTranscriptionSetbackMax = wx.StaticText(panelTranscriber, -1, "5", style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            setbackSizer.Add(lblTranscriptionSetbackMax, 0)
            # Add a spacer so the values are positioned correctly
            setbackSizer.Add((9, 0))

            # Add the Row Sizer to the Panel Sizer
            panelTranSizer.Add(setbackSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

            # On Windows, we can use a number of different media players.  There are trade-offs.
            #   wx.media.MEDIABACKEND_DIRECTSHOW allows speed adjustment, but not WMV or WMA formats (at least on XP).
            #   wx.media.MEDIABACKEND_WMP10 allows WMV and WMA formats, but speed adjustment is broken.
            # Let's allow the user to select which back end to use!
            if ('wxMSW' in wx.PlatformInfo):
                # Add the Media Player Option Label to the Transcriber Settings Tab
                lblMediaPlayer = wx.StaticText(panelTranscriber, -1, _("Media Player Selection"), style=wx.ST_NO_AUTORESIZE)
                # Add the element to the Panel Sizer
                panelTranSizer.Add(lblMediaPlayer, 0, wx.LEFT | wx.TOP, 10)
                # Add a spacer
                panelTranSizer.Add((0, 3))

                # Add the Media Player Option to the Transcriber Settings Tab
                self.chMediaPlayer = wx.Choice(panelTranscriber, -1,
                                               choices = [_('Windows Media Player back end'),
                                                          _('DirectShow back end')])
                self.chMediaPlayer.SetSelection(TransanaGlobal.configData.mediaPlayer)
                # Add the element to the Panel Sizer
                panelTranSizer.Add(self.chMediaPlayer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
                
            # Create a Row Sizer
            lblSpeedSizer = wx.BoxSizer(wx.HORIZONTAL)
            # Add the Media Speed Slider Label to the Transcriber Settings Tab
            lblVideoSpeed = wx.StaticText(panelTranscriber, -1, _("Media Playback Speed"), style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            lblSpeedSizer.Add(lblVideoSpeed, 0)
            # Add a spacer
            lblSpeedSizer.Add((1, 0), 1, wx.EXPAND)

            # Screen elements get a bit out of order here!  We put the CURRENT VALUE of the slider above the slider.
            # We'll use the order of adding things to sizers to get around the logic problems this presents.

            # Add the Media Speed Slider to the Transcriber Settings Tab
            self.videoSpeed = wx.Slider(panelTranscriber, -1, TransanaGlobal.configData.videoSpeed, 1, 20,
                                        style=wx.SL_HORIZONTAL | wx.SL_AUTOTICKS)

            # Add the Media Speed Slider Current Setting Label to the Transcriber Settings Tab
            self.lblVideoSpeedSetting = wx.StaticText(panelTranscriber, -1, "%1.1f" % (float(self.videoSpeed.GetValue()) / 10))
            # Add the element to the Row Sizer
            lblSpeedSizer.Add(self.lblVideoSpeedSetting, 0, wx.ALIGN_RIGHT)

            # Add the LABEL Row Sizer to the Panel Sizer
            panelTranSizer.Add(lblSpeedSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 10)
            # Add a spacer
            panelTranSizer.Add((0, 3))
            # Add the ELEMENT (Control) to the Panel Sizer
            panelTranSizer.Add(self.videoSpeed, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

            # Define the Scroll Event for the Slider to keep the Current Setting Label updated
            wx.EVT_SCROLL(self, self.OnScroll)

            # Create a Row Sizer
            speedSizer = wx.BoxSizer(wx.HORIZONTAL)
            # Add a spacer so the values are positioned correctly
            speedSizer.Add((6, 0))
            # Add the Media Speed Slider Minimum Speed Label to the Transcriber Settings Tab
            lblVideoSpeedMin = wx.StaticText(panelTranscriber, -1, "0.1", style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            speedSizer.Add(lblVideoSpeedMin, 0)
            # Add a spacer
            speedSizer.Add((1, 0), 8, wx.EXPAND)

            # Add the Media Speed Slider Normal Speed Label to the Transcriber Settings Tab
            lblVideoSpeed1 = wx.StaticText(panelTranscriber, -1, "1.0", style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            speedSizer.Add(lblVideoSpeed1, 0)
            # Add a spacer
            speedSizer.Add((1, 0), 9, wx.EXPAND)

            # Add the Media Speed Slider Maximum Speed Label to the Transcriber Settings Tab
            lblVideoSpeedMax = wx.StaticText(panelTranscriber, -1, "2.0", style=wx.ST_NO_AUTORESIZE)
            # Add the element to the Row Sizer
            speedSizer.Add(lblVideoSpeedMax, 0)
            # Add a spacer so the values are positioned properly
            speedSizer.Add((4, 0))

            # Add the Row Sizer to the Panel Sizer
            panelTranSizer.Add(speedSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

            # Create a Row Sizer
            fontSizer = wx.BoxSizer(wx.HORIZONTAL)

            # STC needs the Tab Size setting.  RTC does not support it.
            if not TransanaConstants.USESRTC:
                # Create an Element Sizer
                v1 = wx.BoxSizer(wx.VERTICAL)

                # Add Tab Size
                lblTabSize = wx.StaticText(panelTranscriber, -1, _("Tab Size"))
                # Add the element to the element Sizer
                v1.Add(lblTabSize, 0, wx.BOTTOM, 3)

                # Tab Size Box
                self.tabSize = wx.Choice(panelTranscriber, -1, choices=['0', '2', '4', '6', '8', '10', '12', '14', '16', '18', '20'])
                # Add the element to the element Sizer
                v1.Add(self.tabSize, 0, wx.EXPAND | wx.RIGHT, 10)

                # Add the element Sizer to the Row Sizer
                fontSizer.Add(v1, 1, wx.EXPAND)

                # Set the value to the default value provided by the Configuration Data
                self.tabSize.SetStringSelection(TransanaGlobal.configData.tabSize)

            # Create an Element Sizer
            v2 = wx.BoxSizer(wx.VERTICAL)
            # Add Default Transcript Font
            lblDefaultFont = wx.StaticText(panelTranscriber, -1, _("Default Font"))
            # Add the label to the element Sizer
            v2.Add(lblDefaultFont, 0, wx.BOTTOM, 3)

            # We need to figure out what options we have for the default font.
            # First, let's get a list of all available fonts.
            fontEnum = wx.FontEnumerator()
            fontEnum.EnumerateFacenames()
            fontList = fontEnum.GetFacenames()

            # Now let's set up a list of the fonts we'd like.
            defaultFontList = ['Arial', 'Comic Sans MS', 'Courier', 'Courier New', 'Futura', 'Geneva', 'Helvetica', 'Times', 'Times New Roman', 'Verdana']
            # Initialize the actual font list to nothing.
            choicelist = []
            # Now iterate through the list of fonts we'd like...
            for font in defaultFontList:
                # ... and see if each font is available ...
                if font in fontList:
                    # ... and if so, add it to the list.
                    choicelist.append(font)
                    
            # If the list is empty, let's at least put one real value in it.
            if len(choicelist) == 0:
                font = wx.Font(TransanaGlobal.configData.defaultFontSize, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
                choicelist.append(font.GetFaceName())
                   
            # As of wxPython 2.9.5.0, Mac doesn't support wx.CB_SORT and gives an ugly message about it!
            if 'wxMac' in wx.PlatformInfo:
                style = wx.CB_DROPDOWN
                choicelist.sort()
            else:
                style = wx.CB_DROPDOWN | wx.CB_SORT
            # Default Font Combo Box
            self.defaultFont = wx.ComboBox(panelTranscriber, -1, choices=choicelist, style = style)
            # Add the element to the element Sizer
            v2.Add(self.defaultFont, 0, wx.EXPAND | wx.RIGHT, 10)

            # Set the value to the default value provided by the Configuration Data
            self.defaultFont.SetValue(TransanaGlobal.configData.defaultFontFace)

            # Add the element Sizer to the Row Sizer
            fontSizer.Add(v2, 3, wx.EXPAND)

            # Create an Element Sizer
            v3 = wx.BoxSizer(wx.VERTICAL)
            # Add Default Transcript Font Size
            lblDefaultFontSize = wx.StaticText(panelTranscriber, -1, _("Default Font Size"))
            # Add the label to the element Sizer
            v3.Add(lblDefaultFontSize, 0, wx.BOTTOM, 3)

            # Set up the list of choices
            choicelist = ['8', '10', '11', '12', '14', '16', '20']
                   
            # Default Font Combo Box
            self.defaultFontSize = wx.ComboBox(panelTranscriber, -1, choices=choicelist, style = wx.CB_DROPDOWN)
            # Add the element to the element Sizer
            v3.Add(self.defaultFontSize, 0, wx.EXPAND)

            # Set the value to the default value provided by the Configuration Data
            self.defaultFontSize.SetValue(str(TransanaGlobal.configData.defaultFontSize))

            # Add the element sizer to the Row Sizer
            fontSizer.Add(v3, 2, wx.EXPAND)

            # Add the Row Sizer to the Panel Sizer
            panelTranSizer.Add(fontSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 10)

            # Create a Row Sizer
            fontSizer2 = wx.BoxSizer(wx.HORIZONTAL)

            # Create an Element Sizer
            v4 = wx.BoxSizer(wx.VERTICAL)
            # Add Special Transcript Font
            lblSpecialFont = wx.StaticText(panelTranscriber, -1, _("Special Symbol Font"))
            # Add the label to the element Sizer
            v4.Add(lblSpecialFont, 0, wx.BOTTOM, 3)

            # Now let's set up a list of the fonts we'd like.
            defaultFontList = ['Arial', 'Comic Sans MS', 'Courier', 'Courier New', 'Futura', 'Geneva', 'Helvetica', 'Times', 'Times New Roman', 'Verdana']
            # Initialize the actual font list to nothing.
            choicelist = []
            # Now iterate through the list of fonts we'd like...
            for font in defaultFontList:
                # ... and see if each font is available ...
                if font in fontList:
                    # ... and if so, add it to the list.
                    choicelist.append(font)
                    
            # If the list is empty, let's at least put one real value in it.
            if len(choicelist) == 0:
                font = wx.Font(TransanaGlobal.configData.defaultFontSize, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
                choicelist.append(font.GetFaceName())
                   
            # As of wxPython 2.9.5.0, Mac doesn't support wx.CB_SORT and gives an ugly message about it!
            if 'wxMac' in wx.PlatformInfo:
                style = wx.CB_DROPDOWN
                choicelist.sort()
            else:
                style = wx.CB_DROPDOWN | wx.CB_SORT
            # Special Font Combo Box
            self.specialFont = wx.ComboBox(panelTranscriber, -1, choices=choicelist, style = style)
            # Add the element to the element Sizer
            v4.Add(self.specialFont, 0, wx.EXPAND | wx.RIGHT, 10)

            # Set the value to the default value provided by the Configuration Data
            self.specialFont.SetValue(TransanaGlobal.configData.specialFontFace)

            # Add the element Sizer to the Row Sizer
            fontSizer2.Add(v4, 3, wx.EXPAND)

            # Create an Element Sizer
            v5 = wx.BoxSizer(wx.VERTICAL)
            # Add Special Transcript Font Size
            lblspecialFontSize = wx.StaticText(panelTranscriber, -1, _("Special Symbol Font Size"))
            # Add the label to the element Sizer
            v5.Add(lblspecialFontSize, 0, wx.BOTTOM, 3)

            # Set up the list of choices
            choicelist = ['8', '10', '11', '12', '14', '16', '18', '20', '22', '24']
                   
            # Special Font Size Combo Box
            self.specialFontSize = wx.ComboBox(panelTranscriber, -1, choices=choicelist, style = wx.CB_DROPDOWN)
            # Add the element to the element Sizer
            v5.Add(self.specialFontSize, 0, wx.EXPAND)

            # Set the value to the special value provided by the Configuration Data
            self.specialFontSize.SetValue(str(TransanaGlobal.configData.specialFontSize))

            # Add the element sizer to the Row Sizer
            fontSizer2.Add(v5, 2, wx.EXPAND)

            # Add the Row Sizer to the Panel Sizer
            panelTranSizer.Add(fontSizer2, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 10)

            # Create a Row Sizer
            checkboxSizer = wx.BoxSizer(wx.HORIZONTAL)

            # STC needs the Word Wrap setting.  RTC does not support it.
            if not TransanaConstants.USESRTC:
                # Word Wrap checkbox  NOT SUPPORTED BY RICH TEXT CTRL!!
                self.cbWordWrap = wx.CheckBox(panelTranscriber, -1, _("Word Wrap") + "  ", style=wx.ALIGN_RIGHT)
                # Add the element to the Row Sizer
                checkboxSizer.Add(self.cbWordWrap, 0)
                # Set the value to the configured value for Word Wrap
                self.cbWordWrap.SetValue((TransanaGlobal.configData.wordWrap == stc.STC_WRAP_WORD))
                # Add a spacer to the sizer
                checkboxSizer.Add((10, 0))

            # Auto Save checkbox
            self.cbAutoSave = wx.CheckBox(panelTranscriber, -1, _("Auto Save (10 min)") + "  ", style=wx.ALIGN_RIGHT)
            # Add the element to the Row Sizer
            checkboxSizer.Add(self.cbAutoSave, 0)
            # Set the value to the configured value for Word Wrap
            self.cbAutoSave.SetValue((TransanaGlobal.configData.autoSave))

            # Add a spacer
            checkboxSizer.Add((20, 1))

            # Max Transcript Image Width checkbox
            self.cbMaxTranscriptImageWidth = wx.CheckBox(panelTranscriber, -1, _("Limit Image Width in Transcripts") + "  ", style=wx.ALIGN_RIGHT)
            # Add the element to the Row Sizer
            checkboxSizer.Add(self.cbMaxTranscriptImageWidth, 0)
            # Set the value to the configured value for Word Wrap
            self.cbMaxTranscriptImageWidth.SetValue((TransanaGlobal.configData.maxTranscriptImageWidth))
            

            # Add the row sizer to the panel sizer
            panelTranSizer.Add(checkboxSizer, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)

            # Tell the Transcriber Panel to lay out now and do AutoLayout
            panelTranscriber.SetSizer(panelTranSizer)
            panelTranscriber.SetAutoLayout(True)
            panelTranscriber.Layout()

        # The Message Server Tab should only appear for the Multi-user version of the program.
        if not TransanaConstants.singleUserVersion:
            
            # Add the Message Server Tab to the Notebook
            self.panelMessageServer = MessageServerPanel(notebook, name='OptionsSettings.MessageServerPanel')





      
        # Add the three Panels as Tabs in the Notebook
        notebook.AddPage(panelDirectories, _("Directories"), True)
        # If we're not in the Lab version initial configuration screen ...
        if not self.lab:
            # ... add the Transcriber Settings tab.
            notebook.AddPage(panelTranscriber, _("Transcriber Settings"), False)
        # If we're in the Multi-user version ...
        if not TransanaConstants.singleUserVersion:
            # ... then add the Message Server tab.
            notebook.AddPage(self.panelMessageServer, _("MU Message Server"), False)

        # the tabToShow parameter is the NUMBER of the tab which should be shown initially.
        #   0 = Directories tab
        #   1 = Transcriber Settings tab
        #   2 = Message Server tab, if MU
        if tabToShow != notebook.GetSelection():
            notebook.SetSelection(tabToShow)
        # If the Directories Tab is showing ...
        if notebook.GetSelection() == 0:
            # ... the Media directory should receive initial focus
            self.videoDirectory.SetFocus()
        # If the Transcriber Settings tab is showing ...
        elif notebook.GetSelection() == 1:
            # ... the Transcription Setback slider should receive focus
            self.transcriptionSetback.SetFocus()
        # If the Message Server tab is showing ...
        elif notebook.GetSelection() == 2:
            # ... the Message Server field should recieve initial focus
            self.panelMessageServer.messageServer.SetFocus()

        # Create a Row Sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add a spacer
        btnSizer.Add((1, 0), 1, wx.EXPAND)
        # Define the buttons on the bottom of the form
        # Define the "OK" Button
        btnOK = wx.Button(self, -1, _('OK'))
        # Add the Button to the Row Sizer
        btnSizer.Add(btnOK, 0, wx.RIGHT, 10)

        # Define the Cancel Button
        btnCancel = wx.Button(self, -1, _('Cancel'))
        # Add the Button to the Row Sizer
        btnSizer.Add(btnCancel, 0, wx.RIGHT, 10)

        # Define the Help Button
        btnHelp = wx.Button(self, -1, _('Help'))
        # Add the Button to the Row Sizer
        btnSizer.Add(btnHelp, 0)

        # Add the Row Sizer to the main form sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 3)

        # Attach events to the Buttons
        wx.EVT_BUTTON(self, btnOK.GetId(), self.OnOK)
        wx.EVT_BUTTON(self, btnCancel.GetId(), self.OnCancel)
        wx.EVT_BUTTON(self, btnHelp.GetId(), self.OnHelp)

        # Bind the notebook page change event
        notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.OnPageChange)

        # Lay out the Window, and tell it to Auto Layout
        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()
        TransanaGlobal.CenterOnPrimary(self)

        # Get the form's current size
        (width, height) = self.GetSize()
        # Set the current size as the minimum size
        self.SetSizeHints(width, height)

        # Make OK the Default Button
        self.SetDefaultItem(btnOK)
        # Show the newly created Window as a modal dialog
        self.ShowModal()

    def OnOK(self, event):
        """ Method attached to the 'OK' Button """
        # Assign the values from this form to the Transana Global Variables
        
        # If the Waveform Directory does not end with the separator character, add one.
        # Then update the Global Waveform Directory
        if (len(self.waveformDirectory.GetValue()) > 0) and \
           self.waveformDirectory.GetValue()[-1] != os.sep:
            TransanaGlobal.configData.visualizationPath = self.waveformDirectory.GetValue() + os.sep
        else:
            TransanaGlobal.configData.visualizationPath = self.waveformDirectory.GetValue()
        # If the Media Directory does not end with the separator character, add one,
        # then update the Global Media Directory.  (But the lab version doesn't HAVE this value at start-up time.)
        if (len(self.videoDirectory.GetValue()) > 0) and \
           (self.videoDirectory.GetValue()[-1] != os.sep):
            tempVideoPath = self.videoDirectory.GetValue() + os.sep
        else:
            tempVideoPath = self.videoDirectory.GetValue()
        # If the Database Directory does not end with the separator character, add one.
        # Then update the Global Database Directory
        if (len(self.databaseDirectory.GetValue()) > 0) and \
           (self.databaseDirectory.GetValue()[-1] != os.sep):
            TransanaGlobal.configData.databaseDir = self.databaseDirectory.GetValue() + os.sep
        else:
            TransanaGlobal.configData.databaseDir = self.databaseDirectory.GetValue()
        # If we're not in the LAB version and the Media Library Path has changed ...
        if (not self.lab) and (tempVideoPath != TransanaGlobal.configData.videoPath):
            # First, find out if there are Episodes or Clips that need to be changed in the Database
            (episodeCount, clipCount) = DBInterface.VideoFilePaths(tempVideoPath)
            # If there are records to update ...
            if episodeCount > 0 or clipCount > 0:
                # Build a Dialog to prompt the user about updating the records
                msg = _("%d Episode Records and %s Clip Records will be updated.")
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(msg, 'utf8')
                # Can't use Dialogs.InfoDialog because Cancel is an option.  Added an OkCancel option to QuestionDialog.
                dlg = Dialogs.QuestionDialog(TransanaGlobal.menuWindow, msg % (episodeCount, clipCount), useOkCancel=True)
                # Display the Dialog
                result = dlg.LocalShowModal()
                # Destroy the Dialog
                dlg.Destroy
                # if the User says OK ...
                if result == wx.ID_OK:
                    # Set the cursor to the Wait Cursor
                    self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
                    # Actually UPDATE the paths by adding the update=True parameter
                    DBInterface.VideoFilePaths(tempVideoPath, True)
                    # Add the new Path to the Configuration Data  (This won't get updated if the user says "NO"!)
                    TransanaGlobal.configData.videoPath = tempVideoPath
                    # Set the cursor to the arrow again.
                    self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
            # If there are no records that would require updating...
            else:
                # Add the new Path to the Configuration Data
                TransanaGlobal.configData.videoPath = tempVideoPath

        # If we ARE in the LAB version ...
        else:
            # Add the new Path to the Configuration Data
            TransanaGlobal.configData.videoPath = tempVideoPath
            # Set the cursor to the arrow again.
            self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
        # If we're not in the LAB initial configuration ...  (The lab version doesn't HAVE these values at start-up time.)
        if not self.lab:
            # Update the Global Transcription Setback
            TransanaGlobal.configData.transcriptionSetback = self.transcriptionSetback.GetValue()
            # If on Windows ...
            if ('wxMSW' in wx.PlatformInfo):
                # Update the Media Player selection
                TransanaGlobal.configData.mediaPlayer = self.chMediaPlayer.GetSelection()
            # Update the Global Media Speed
            TransanaGlobal.configData.videoSpeed = self.videoSpeed.GetValue()

            # STC needs the Tab Size and Word Wrap settings.  RTC does not support them.
            if not TransanaConstants.USESRTC:
                # Update the tab size
                TransanaGlobal.configData.tabSize = self.tabSize.GetStringSelection()
                # Update the Word Wrap setting
                if self.cbWordWrap.GetValue():
                    wordWrapValue = stc.STC_WRAP_WORD
                else:
                    wordWrapValue = stc.STC_WRAP_NONE
                # Update the Word Wrap value in the configuration system
                TransanaGlobal.configData.wordWrap = wordWrapValue
            # Update the Auto Save value
            TransanaGlobal.configData.autoSave = self.cbAutoSave.GetValue()
            # Update the Max Transcript Image Width value
            TransanaGlobal.configData.maxTranscriptImageWidth = self.cbMaxTranscriptImageWidth.GetValue()
            # Update the Global Default Font
            TransanaGlobal.configData.defaultFontFace = self.defaultFont.GetValue()
            # Update the Global Default Font Size
            TransanaGlobal.configData.defaultFontSize = int(self.defaultFontSize.GetValue())
            # Update the Global Special Font
            TransanaGlobal.configData.specialFontFace = self.specialFont.GetValue()
            # Update the Global Special Font Size
            TransanaGlobal.configData.specialFontSize = int(self.specialFontSize.GetValue())

        # Make sure the current Media Library and visualization path settings are saved in the configuration under the (username, server, database) key.
        TransanaGlobal.configData.pathsByDB[(TransanaGlobal.userName.encode('utf8'), TransanaGlobal.configData.host.encode('utf8'), TransanaGlobal.configData.database.encode('utf8'))] = \
            {'videoPath' : TransanaGlobal.configData.videoPath.encode('utf8'),
             'visualizationPath' : TransanaGlobal.configData.visualizationPath.encode('utf8')}

        if not TransanaConstants.singleUserVersion:
            # TODO:  If Message Server is changed, disconnect and connect to new Message Server!
            # Update the Global Message Server Variable
            TransanaGlobal.configData.messageServer = self.panelMessageServer.messageServer.GetValue()
            # Update the Global Message Server Port
            TransanaGlobal.configData.messageServerPort = int(self.panelMessageServer.messageServerPort.GetValue())
        
        # Make sure the oldDatabaseDir ends with the proper seperator character.
        # (But the LAB version won't HAVE an oldDatabaseDir.)
        if (self.oldDatabaseDir <> '') and (self.oldDatabaseDir[-1] != os.sep):
            self.oldDatabaseDir = self.oldDatabaseDir + os.sep
        # If database directory was changed, ...  (But the LAB version won't HAVE an oldDatabaseDir.)
        if (self.oldDatabaseDir <> '') and (self.oldDatabaseDir != TransanaGlobal.configData.databaseDir):
            # ... signal need to shut down Transana
            self.parent.shutDown = True

        # Let's save the configuration data so it doesn't disappear if Transana crashes, not, of course, that Transana ever crashes.
        TransanaGlobal.configData.SaveConfiguration()
        self.Close()

    def OnCancel(self, event):
        """ Method attached to the 'Close' Button """
        # Close the Dialog without updating any Global Variables
        self.Close()

    def OnHelp(self, event):
        """Invoked when dialog Help button is activated."""
        # Normally, the Control Object is already defined.  Check, though, because it's not for the LAB version's
        # initial configuration dialog.
        if (TransanaGlobal.menuWindow != None) and (TransanaGlobal.menuWindow.ControlObject != None):
            TransanaGlobal.menuWindow.ControlObject.Help('Program Settings')
        # If we're using the Lab version and the user calls for help from the initial configuration dialog ...
        else:
            # ... import the ControlObject Class ...
            import ControlObjectClass
            # ... create a temporary Control Object ...
            tmpCO = ControlObjectClass.ControlObject()
            # ... and THEN call Help!
            tmpCO.Help('Program Settings')

    def OnBrowse(self, event):
        """ Implements the "Browse" button for the Waveform or Media Library Directories on the Directories Tab """
        # Set the prompt and current value for the correct Browse operation (Waveform or Media Library folder)
        if event.GetId() == self.btnWaveformBrowse.GetId():
            prompt = _("Choose a Waveforms directory:")
            currentValue = self.waveformDirectory.GetValue()
        elif event.GetId() == self.btnVideoBrowse.GetId():
            prompt = _("Choose a Media Library directory:")
            currentValue = self.videoDirectory.GetValue()
        elif event.GetId() == self.btnDatabaseBrowse.GetId():
            prompt = _("Choose a Database Directory:")
            currentValue = self.databaseDirectory.GetValue()
        # Create a Directory Dialog Box
        dlg = wx.DirDialog(self, prompt, currentValue, style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        # Display the Dialog modally and process if OK is selected.
        if dlg.ShowModal() == wx.ID_OK:
            # Assign the Directory selected to the appropriate field
            if event.GetId() == self.btnWaveformBrowse.GetId():
                self.waveformDirectory.SetValue(dlg.GetPath())
            elif event.GetId() == self.btnVideoBrowse.GetId():
                self.videoDirectory.SetValue(dlg.GetPath())
                # If no Waveform directory is assigned when the Media Library is selected (as will be true for the LAB version) ...
                if self.waveformDirectory.GetValue() == '':
                    # ... then auto-assign a waveforms subdirectory!
                    self.waveformDirectory.SetValue(os.path.join(dlg.GetPath(), 'waveforms'))
            elif event.GetId() == self.btnDatabaseBrowse.GetId():
                self.databaseDirectory.SetValue(dlg.GetPath())
        # Destroy the Dialog
        dlg.Destroy

    def OnScroll(self, event):
        """ Handle the Scroll Event for the Media Speed Slider. """
        # Update the Current Media Speed Label
        self.lblVideoSpeedSetting.SetLabel("%1.1f" % (float(self.videoSpeed.GetValue()) / 10))

    def OnPageChange(self, event):
        """ Notebook page change event """
        # Call the parent page change method so the tab is drawn
        event.Skip()
        # If the Directories tab is showing ...
        if event.GetSelection() == 0:
            # ... the Media Library should recieve focus
            wx.CallAfter(self.videoDirectory.SetFocus)
        # If the Transcriber Settings tab is showing ...
        elif event.GetSelection() == 1:
            # ... the Transcription Setback slider should receive focus
            wx.CallAfter(self.transcriptionSetback.SetFocus)
        # If the Message Server tab is showing ...
        elif event.GetSelection() == 2:
            # ... the Message Server should receive focus
            wx.CallAfter(self.panelMessageServer.messageServer.SetFocus)


class MessageServerPanel(wx.Panel):
    def __init__(self, parent, name):
            # Add the Message Server Tab to the Notebook
            wx.Panel.__init__(self, parent, -1, size=parent.GetSizeTuple(), name=name)
            
            # Define the main VERTICAL sizer for the Notebook Page
            panelMsgSizer = wx.BoxSizer(wx.VERTICAL)
            
            # Add the Message Server Label to the Message Server Tab
            lblMessageServer = wx.StaticText(self, -1, _("Transana-MU Message Server Host Name"), style=wx.ST_NO_AUTORESIZE)
            # Add the label to the Panel Sizer
            panelMsgSizer.Add(lblMessageServer, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
            # Add a spacer
            panelMsgSizer.Add((0, 3))
            
            # Add the Message Server TextCtrl to the Message Server Tab
            self.messageServer = wx.TextCtrl(self, -1, TransanaGlobal.configData.messageServer)
            # Add the element to the Panel Sizer
            panelMsgSizer.Add(self.messageServer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            # Add the Message Server Port Label to the Message Server Tab
            lblMessageServerPort = wx.StaticText(self, -1, _("Port"), style=wx.ST_NO_AUTORESIZE)
            # Add the label to the Panel Sizer
            panelMsgSizer.Add(lblMessageServerPort, 0, wx.LEFT | wx.RIGHT, 10)
            # Add a spacer
            panelMsgSizer.Add((0, 3))
            
            # Add the Message Server Port TextCtrl to the Message Server Tab
            self.messageServerPort = wx.TextCtrl(self, -1, str(TransanaGlobal.configData.messageServerPort))
            # Add the element to the Panel Sizer
            panelMsgSizer.Add(self.messageServerPort, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

            # Tell the Message Server Panel to lay out now and do AutoLayout
            self.SetSizer(panelMsgSizer)
            self.SetAutoLayout(True)
            self.Layout()
