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

"""This file implements the OptionsSettings class, which defines the
   main Options > Settings Dialog Box and values. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare'

# import wxPython
import wx
# import the Python os module
import os
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

    def __init__(self, parent, tabToShow=0):
        """ Initialize the Program Options Dialog Box """
        self.parent = parent
        if 'wxMSW' in wx.PlatformInfo:
            dlgHeight = 360
        else:
            dlgHeight = 300
        # Define the Dialog Box
        wx.Dialog.__init__(self, parent, -1, _("Transana Settings"), wx.DefaultPosition, wx.Size(550, dlgHeight), style=wx.CAPTION | wx.SYSTEM_MENU | wx.THICK_FRAME)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Define a wxNotebook for the Tab structure
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self, wx.Top, 0)
        lay.left.SameAs(self, wx.Left, 0)
        lay.right.SameAs(self, wx.Right, 0)
        lay.bottom.SameAs(self, wx.Bottom, 27)
        notebook = wx.Notebook(self, -1, size=self.GetSizeTuple())
        notebook.SetConstraints(lay)

        # Define the Directories Tab that goes in the wxNotebook
        panelDirectories = wx.Panel(notebook, -1, size=notebook.GetSizeTuple(), name='OptionsSettings.DirectoriesPanel')

        # Add the Waveform Directory Label to the Directories Tab
        lblWaveformDirectory = wx.StaticText(panelDirectories, -1, _("Waveform Directory"), style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.SameAs(panelDirectories, wx.Top, 20)
        lay.left.SameAs(panelDirectories, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        lblWaveformDirectory.SetConstraints(lay)
        
        # Add the Waveform Directory TextCtrl to the Directories Tab
        # If the visualization path is not empty, we should normalize the path specification
        if TransanaGlobal.configData.visualizationPath == '':
            visualizationPath = TransanaGlobal.configData.visualizationPath
        else:
            visualizationPath = os.path.normpath(TransanaGlobal.configData.visualizationPath)
        self.waveformDirectory = wx.TextCtrl(panelDirectories, -1, visualizationPath)
        lay = wx.LayoutConstraints()
        lay.top.Below(lblWaveformDirectory, 3)
        lay.left.SameAs(panelDirectories, wx.Left, 10)
        lay.right.SameAs(panelDirectories, wx.Right, 100)
        lay.height.AsIs()
        self.waveformDirectory.SetConstraints(lay)

        # Add the Waveform Directory Browse Button to the Directories Tab
        self.btnWaveformBrowse = wx.Button(panelDirectories, -1, _("Browse"))
        lay = wx.LayoutConstraints()
        lay.top.Below(lblWaveformDirectory, 3)
        lay.left.RightOf(self.waveformDirectory, 10)
        lay.right.SameAs(panelDirectories, wx.Right, 10)
        lay.height.AsIs()
        self.btnWaveformBrowse.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnWaveformBrowse.GetId(), self.OnBrowse)
        
        # Add the Video Root Directory Label to the Directories Tab
        lblVideoDirectory = wx.StaticText(panelDirectories, -1, _("Video Root Directory"), style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.waveformDirectory, 20)
        lay.left.SameAs(panelDirectories, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        lblVideoDirectory.SetConstraints(lay)
        
        # Add the Video Root Directory TextCtrl to the Directories Tab
        # If the video path is not empty, we should normalize the path specification
        if TransanaGlobal.configData.videoPath == '':
            videoPath = TransanaGlobal.configData.videoPath
        else:
            videoPath = os.path.normpath(TransanaGlobal.configData.videoPath)
        self.videoDirectory = wx.TextCtrl(panelDirectories, -1, videoPath)
        lay = wx.LayoutConstraints()
        lay.top.Below(lblVideoDirectory, 3)
        lay.left.SameAs(panelDirectories, wx.Left, 10)
        lay.right.SameAs(panelDirectories, wx.Right, 100)
        lay.height.AsIs()
        self.videoDirectory.SetConstraints(lay)

        # Add the Video Root Directory Browse Button to the Directories Tab
        self.btnVideoBrowse = wx.Button(panelDirectories, -1, _("Browse"))
        lay = wx.LayoutConstraints()
        lay.top.Below(lblVideoDirectory, 3)
        lay.left.RightOf(self.videoDirectory, 10)
        lay.right.SameAs(panelDirectories, wx.Right, 10)
        lay.height.AsIs()
        self.btnVideoBrowse.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnVideoBrowse.GetId(), self.OnBrowse)

        # Add the Database Directory Label to the Directories Tab
        lblDatabaseDirectory = wx.StaticText(panelDirectories, -1, _("Database Directory"), style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.videoDirectory, 20)
        lay.left.SameAs(panelDirectories, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        lblDatabaseDirectory.SetConstraints(lay)
        
        # Add the Database Directory TextCtrl to the Directories Tab
        # If the database path is not empty, we should normalize the path specification
        if TransanaGlobal.configData.databaseDir == '':
            databaseDir = TransanaGlobal.configData.databaseDir
        else:
            databaseDir = os.path.normpath(TransanaGlobal.configData.databaseDir)
        self.oldDatabaseDir = databaseDir
        self.databaseDirectory = wx.TextCtrl(panelDirectories, -1, databaseDir)
        lay = wx.LayoutConstraints()
        lay.top.Below(lblDatabaseDirectory, 3)
        lay.left.SameAs(panelDirectories, wx.Left, 10)
        lay.right.SameAs(panelDirectories, wx.Right, 100)
        lay.height.AsIs()
        self.databaseDirectory.SetConstraints(lay)

        # Add the Database Directory Browse Button to the Directories Tab
        self.btnDatabaseBrowse = wx.Button(panelDirectories, -1, _("Browse"))
        lay = wx.LayoutConstraints()
        lay.top.Below(lblDatabaseDirectory, 3)
        lay.left.RightOf(self.databaseDirectory, 10)
        lay.right.SameAs(panelDirectories, wx.Right, 10)
        lay.height.AsIs()
        self.btnDatabaseBrowse.SetConstraints(lay)
        wx.EVT_BUTTON(self, self.btnDatabaseBrowse.GetId(), self.OnBrowse)
        
        # The Database Directory should not be visible for the Multi-user version of the program.
        # Let's just hide it so that the program doesn't crash for being unable to populate the control.
        if not TransanaConstants.singleUserVersion:
            lblDatabaseDirectory.Show(False)
            self.databaseDirectory.Show(False)
            self.btnDatabaseBrowse.Show(False)

        # Tell the Directories Panel to lay out now and do AutoLayout
        panelDirectories.SetAutoLayout(True)
        panelDirectories.Layout()

        # Add the Transcriber Panel to the Notebook
        panelTranscriber = wx.Panel(notebook, -1, size=notebook.GetSizeTuple(), name='OptionsSettings.TranscriberPanel')

        # Add the Video Setback Label to the Transcriber Settings Tab
        lblTranscriptionSetback = wx.StaticText(panelTranscriber, -1, _("Transcription Setback:  (Auto-rewind interval for Ctrl-S)"), style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.SameAs(panelTranscriber, wx.Top, 20)
        lay.left.SameAs(panelTranscriber, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        lblTranscriptionSetback.SetConstraints(lay)

        # Add the Video Setback Slider to the Transcriber Settings Tab
        self.transcriptionSetback = wx.Slider(panelTranscriber, -1, TransanaGlobal.configData.transcriptionSetback, 0, 5, style=wx.SL_HORIZONTAL | wx.SL_AUTOTICKS)
        lay = wx.LayoutConstraints()
        lay.top.Below(lblTranscriptionSetback, 3)
        lay.left.SameAs(panelTranscriber, wx.Left, 10)
        lay.right.SameAs(panelTranscriber, wx.Right, 10)
        lay.height.AsIs()
        self.transcriptionSetback.SetConstraints(lay)

        # Add the Video Setback "0" Value Label to the Transcriber Settings Tab
        lblTranscriptionSetbackMin = wx.StaticText(panelTranscriber, -1, "0", style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.transcriptionSetback, 2)
        lay.left.SameAs(self.transcriptionSetback, wx.Left, 7)
        lay.width.AsIs()
        lay.height.AsIs()
        lblTranscriptionSetbackMin.SetConstraints(lay)

        # Add the Video Setback "1" Value Label to the Transcriber Settings Tab
        lblTranscriptionSetback1 = wx.StaticText(panelTranscriber, -1, "1", style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.transcriptionSetback, 2)
        # The "1" position is 20% of the way between 0 and 5.  However, 23% looks better on Windows.
        lay.left.PercentOf(self.transcriptionSetback, wx.Width, 23)
        lay.width.AsIs()
        lay.height.AsIs()
        lblTranscriptionSetback1.SetConstraints(lay)

        # Add the Video Setback "2" Value Label to the Transcriber Settings Tab
        lblTranscriptionSetback2 = wx.StaticText(panelTranscriber, -1, "2", style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.transcriptionSetback, 2)
        # The "2" position is 40% of the way between 0 and 5.  However, 42% looks better on Windows.
        lay.left.PercentOf(self.transcriptionSetback, wx.Width, 42)
        lay.width.AsIs()
        lay.height.AsIs()
        lblTranscriptionSetback2.SetConstraints(lay)

        # Add the Video Setback "3" Value Label to the Transcriber Settings Tab
        lblTranscriptionSetback3 = wx.StaticText(panelTranscriber, -1, "3", style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.transcriptionSetback, 2)
        # The "3" position is 60% of the way between 0 and 5.  However, 61% looks better on Windows.
        lay.left.PercentOf(self.transcriptionSetback, wx.Width, 61)
        lay.width.AsIs()
        lay.height.AsIs()
        lblTranscriptionSetback3.SetConstraints(lay)

        # Add the Video Setback "4" Value Label to the Transcriber Settings Tab
        lblTranscriptionSetback4 = wx.StaticText(panelTranscriber, -1, "4", style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.transcriptionSetback, 2)
        # The "4" position is 80% of the way between 0 and 5.
        lay.left.PercentOf(self.transcriptionSetback, wx.Width, 80)
        lay.width.AsIs()
        lay.height.AsIs()
        lblTranscriptionSetback4.SetConstraints(lay)

        # Add the Video Setback "5" Value Label to the Transcriber Settings Tab
        lblTranscriptionSetbackMax = wx.StaticText(panelTranscriber, -1, "5", style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.transcriptionSetback, 2)
        lay.right.SameAs(self.transcriptionSetback, wx.Right, 7)
        lay.width.AsIs()
        lay.height.AsIs()
        lblTranscriptionSetbackMax.SetConstraints(lay)

        # On Windows, we can use a number of different media players.  There are trade-offs.
        #   wx.media.MEDIABACKEND_DIRECTSHOW allows speed adjustment, but not WMV or WMA formats.
        #   wx.media.MEDIABACKEND_WMP10 allows WMV and WMA formats, but speed adjustment is broken.
        # Let's allow the user to select which back end to use!
        # This option is Windows only!
        if 'wxMSW' in wx.PlatformInfo:
            # Add the Media Player Option Label to the Transcriber Settings Tab
            lblMediaPlayer = wx.StaticText(panelTranscriber, -1, _("Media Player Selection"), style=wx.ST_NO_AUTORESIZE)
            lay = wx.LayoutConstraints()
            lay.top.Below(lblTranscriptionSetbackMin, 15)
            lay.left.SameAs(panelTranscriber, wx.Left, 10)
            lay.width.AsIs()
            lay.height.AsIs()
            lblMediaPlayer.SetConstraints(lay)

            # Add the Media Player Option to the Transcriber Settings Tab
            self.chMediaPlayer = wx.Choice(panelTranscriber, -1, choices = [_('Enable WMV and WMA formats, disable speed control'), _('Disable WMV and WMA formats, enable speed control')])
            self.chMediaPlayer.SetSelection(TransanaGlobal.configData.mediaPlayer)
            lay = wx.LayoutConstraints()
            lay.top.Below(lblMediaPlayer, 3)
            lay.left.SameAs(panelTranscriber, wx.Left, 10)
            lay.width.AsIs()
            lay.height.AsIs()
            self.chMediaPlayer.SetConstraints(lay)
# REMOVED!  When Speed Control is OFF for WMP, it still works for QuickTime!
#            self.chMediaPlayer.Bind(wx.EVT_CHOICE, self.OnMediaPlayerSelect)

            nextLabelPositioner = self.chMediaPlayer
        else:
            nextLabelPositioner = lblTranscriptionSetbackMin
            

        # Add the Video Speed Slider Label to the Transcriber Settings Tab
        lblVideoSpeed = wx.StaticText(panelTranscriber, -1, _("Video Playback Speed"), style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(nextLabelPositioner, 15)
        lay.left.SameAs(panelTranscriber, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        lblVideoSpeed.SetConstraints(lay)

        # Add the Video Speed Slider to the Transcriber Settings Tab
        self.videoSpeed = wx.Slider(panelTranscriber, -1, TransanaGlobal.configData.videoSpeed, 1, 20, style=wx.SL_HORIZONTAL | wx.SL_AUTOTICKS)
        lay = wx.LayoutConstraints()
        lay.top.Below(lblVideoSpeed, 3)
        lay.left.SameAs(panelTranscriber, wx.Left, 10)
        lay.right.SameAs(panelTranscriber, wx.Right, 10)
        lay.height.AsIs()
        self.videoSpeed.SetConstraints(lay)

        # Add the Video Speed Slider Current Setting Label to the Transcriber Settings Tab
        self.lblVideoSpeedSetting = wx.StaticText(panelTranscriber, -1, "%1.1f" % (float(self.videoSpeed.GetValue()) / 10))
        lay = wx.LayoutConstraints()
        lay.top.Below(nextLabelPositioner, 15)
        lay.right.SameAs(panelTranscriber, wx.Right, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        self.lblVideoSpeedSetting.SetConstraints(lay)

        # Define the Scroll Event for the Slider to keep the Current Setting Label updated
        wx.EVT_SCROLL(self, self.OnScroll)

# REMOVED!  When Speed Control is OFF for WMP, it still works for QuickTime!
        # Disable the slider if it should be disabled
#        if TransanaGlobal.configData.mediaPlayer == 0:
#            self.videoSpeed.Enable(False)
#            self.lblVideoSpeedSetting.SetLabel("%1.1f" % (1.0))

        # Add the Video Speed Slider Minimum Speed Label to the Transcriber Settings Tab
        lblVideoSpeedMin = wx.StaticText(panelTranscriber, -1, "0.1", style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.videoSpeed, 2)
        lay.left.SameAs(self.videoSpeed, wx.Left, 0)
        lay.width.AsIs()
        lay.height.AsIs()
        lblVideoSpeedMin.SetConstraints(lay)

        # Add the Video Speed Slider Normal Speed Label to the Transcriber Settings Tab
        lblVideoSpeed1 = wx.StaticText(panelTranscriber, -1, "1.0", style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.videoSpeed, 2)
        # The "center" (1.0) position is 47% (9 / 19) of the way between 0.1 and 2.0.  However, 48% looks better on Windows.
        lay.left.PercentOf(self.videoSpeed, wx.Width, 48)
        lay.width.AsIs()
        lay.height.AsIs()
        lblVideoSpeed1.SetConstraints(lay)

        # Add the Video Speed Slider Maximum Speed Label to the Transcriber Settings Tab
        lblVideoSpeedMax = wx.StaticText(panelTranscriber, -1, "2.0", style=wx.ST_NO_AUTORESIZE)
        lay = wx.LayoutConstraints()
        lay.top.Below(self.videoSpeed, 2)
        lay.right.SameAs(self.videoSpeed, wx.Right, 0)
        lay.width.AsIs()
        lay.height.AsIs()
        lblVideoSpeedMax.SetConstraints(lay)

        # Add Default Transcript Font
        lay = wx.LayoutConstraints()
        lay.top.Below(lblVideoSpeedMin, 15)
        lay.left.SameAs(panelTranscriber, wx.Left, 10)
        lay.width.AsIs()
        lay.height.AsIs()
        lblDefaultFont = wx.StaticText(panelTranscriber, -1, _("Default Font"))
        lblDefaultFont.SetConstraints(lay)

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
               
        # Default Font Combo Box
        lay = wx.LayoutConstraints()
        lay.top.Below(lblDefaultFont, 3)
        lay.left.SameAs(panelTranscriber, wx.Left, 10)
        lay.right.PercentOf(panelTranscriber, wx.Width, 45)
        lay.height.AsIs()
        self.defaultFont = wx.ComboBox(panelTranscriber, -1, choices=choicelist, style = wx.CB_DROPDOWN | wx.CB_SORT)
        self.defaultFont.SetConstraints(lay)

        # Set the value to the default value provided by the Configuration Data
        self.defaultFont.SetValue(TransanaGlobal.configData.defaultFontFace)

        # Add Default Transcript Font Size
        lay = wx.LayoutConstraints()
        lay.top.Below(lblVideoSpeedMin, 15)
        lay.left.PercentOf(panelTranscriber, wx.Width, 55)
        lay.width.AsIs()
        lay.height.AsIs()
        lblDefaultFontSize = wx.StaticText(panelTranscriber, -1, _("Default Font Size"))
        lblDefaultFontSize.SetConstraints(lay)

        # Set up the list of choices
        choicelist = ['8', '10', '11', '12', '14', '16', '20']
               
        # Default Font Combo Box
        lay = wx.LayoutConstraints()
        lay.top.Below(lblDefaultFont, 3)
        lay.left.SameAs(lblDefaultFontSize, wx.Left, 0)
        lay.right.SameAs(panelTranscriber, wx.Right, 10)
        lay.height.AsIs()
        self.defaultFontSize = wx.ComboBox(panelTranscriber, -1, choices=choicelist, style = wx.CB_DROPDOWN)
        self.defaultFontSize.SetConstraints(lay)

        # Set the value to the default value provided by the Configuration Data
        self.defaultFontSize.SetValue(str(TransanaGlobal.configData.defaultFontSize))

        # Tell the Transcriber Panel to lay out now and do AutoLayout
        panelTranscriber.SetAutoLayout(True)
        panelTranscriber.Layout()

        # The Message Server Tab should only appear for the Multi-user version of the program.
        if not TransanaConstants.singleUserVersion:
            # Add the Message Server Tab to the Notebook
            panelMessageServer = wx.Panel(notebook, -1, size=notebook.GetSizeTuple(), name='OptionsSettings.MessageServerPanel')
            
            # Add the Message Server Label to the Message Server Tab
            lblMessageServer = wx.StaticText(panelMessageServer, -1, _("Transana-MU Message Server Host Name"), style=wx.ST_NO_AUTORESIZE)
            lay = wx.LayoutConstraints()
            lay.top.SameAs(panelMessageServer, wx.Top, 20)
            lay.left.SameAs(panelMessageServer, wx.Left, 10)
            lay.width.AsIs()
            lay.height.AsIs()
            lblMessageServer.SetConstraints(lay)
            
            # Add the Message Server TextCtrl to the Message Server Tab
            self.messageServer = wx.TextCtrl(panelMessageServer, -1, TransanaGlobal.configData.messageServer)
            lay = wx.LayoutConstraints()
            lay.top.Below(lblMessageServer, 3)
            lay.left.SameAs(panelMessageServer, wx.Left, 10)
            lay.right.SameAs(panelMessageServer, wx.Right, 10)
            lay.height.AsIs()
            self.messageServer.SetConstraints(lay)

            # Add the Message Server Port Label to the Message Server Tab
            lblMessageServerPort = wx.StaticText(panelMessageServer, -1, _("Port"), style=wx.ST_NO_AUTORESIZE)
            lay = wx.LayoutConstraints()
            lay.top.Below(self.messageServer, 20)
            lay.left.SameAs(panelMessageServer, wx.Left, 10)
            lay.width.AsIs()
            lay.height.AsIs()
            lblMessageServerPort.SetConstraints(lay)
            
            # Add the Message Server Port TextCtrl to the Message Server Tab
            self.messageServerPort = wx.TextCtrl(panelMessageServer, -1, str(TransanaGlobal.configData.messageServerPort))
            lay = wx.LayoutConstraints()
            lay.top.Below(lblMessageServerPort, 3)
            lay.left.SameAs(panelMessageServer, wx.Left, 10)
            lay.width.Absolute(50)
            lay.height.AsIs()
            self.messageServerPort.SetConstraints(lay)

            # Tell the Message Server Panel to lay out now and do AutoLayout
            panelMessageServer.SetAutoLayout(True)
            panelMessageServer.Layout()
      
        # Add the three Panels as Tabs in the Notebook
        notebook.AddPage(panelDirectories, _("Directories"), True)
        notebook.AddPage(panelTranscriber, _("Transcriber Settings"), False)
        if not TransanaConstants.singleUserVersion:
            notebook.AddPage(panelMessageServer, _("MU Message Server"), False)

        # the tabToShow parameter is the NUMBER of the tab which should be shown initially.
        #   0 = Directories tab
        #   1 = Transcriber Settings tab
        #   2 = Message Server tab, if MU
        if tabToShow != notebook.GetSelection():
            notebook.SetSelection(tabToShow)

        # Define the buttons on the bottom of the form
        # Define the "OK" Button
        lay = wx.LayoutConstraints()
        lay.top.Below(notebook, 3)
        lay.width.Absolute(85)
        lay.left.SameAs(self, wx.Right, -268)
        lay.bottom.SameAs(self, wx.Bottom, 0)
        btnOK = wx.Button(self, -1, _('OK'))
        btnOK.SetConstraints(lay)

        # Define the Cancel Button
        lay = wx.LayoutConstraints()
        lay.top.Below(notebook, 3)
        lay.width.Absolute(85)
        lay.left.RightOf(btnOK, 6)
        lay.bottom.SameAs(self, wx.Bottom, 0)
        btnCancel = wx.Button(self, -1, _('Cancel'))
        btnCancel.SetConstraints(lay)

        # Define the Help Button
        lay = wx.LayoutConstraints()
        lay.top.Below(notebook, 3)
        lay.width.Absolute(85)
        lay.left.RightOf(btnCancel, 6)
        lay.bottom.SameAs(self, wx.Bottom, 0)
        btnHelp = wx.Button(self, -1, _('Help'))
        btnHelp.SetConstraints(lay)

        # Attach events to the Buttons
        wx.EVT_BUTTON(self, btnOK.GetId(), self.OnOK)
        wx.EVT_BUTTON(self, btnCancel.GetId(), self.OnCancel)
        wx.EVT_BUTTON(self, btnHelp.GetId(), self.OnHelp)

        # Lay out the Window, and tell it to Auto Layout
        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()
        # Show the newly created Window as a modal dialog
        self.ShowModal()

    def OnOK(self, event):
        """ Method attached to the 'OK' Button """
        # Assign the values from this form to the Transana Global Variables
        
        # If the Waveform Directory does not end with the separator character, add one.
        # Then update the Global Waveform Directory
        if self.waveformDirectory.GetValue()[-1] != os.sep:
            TransanaGlobal.configData.visualizationPath = self.waveformDirectory.GetValue() + os.sep
        else:
            TransanaGlobal.configData.visualizationPath = self.waveformDirectory.GetValue()
        # If the Video Directory does not end with the separator character, add one.
        # Then update the Global Video Directory
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
        # If the Video Root Path has changed ...
        if tempVideoPath != TransanaGlobal.configData.videoPath:
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
        # Update the Global Transcription Setback
        TransanaGlobal.configData.transcriptionSetback = self.transcriptionSetback.GetValue()
        # If on Windows ...
        if 'wxMSW' in wx.PlatformInfo:
            # Update the Media Player selection
            TransanaGlobal.configData.mediaPlayer = self.chMediaPlayer.GetSelection()
        # Update the Global Video Speed
        TransanaGlobal.configData.videoSpeed = self.videoSpeed.GetValue()
        # Update the Global Default Font
        TransanaGlobal.configData.defaultFontFace = self.defaultFont.GetValue()
        # Update the Global Default Font Size
        TransanaGlobal.configData.defaultFontSize = int(self.defaultFontSize.GetValue())
        if not TransanaConstants.singleUserVersion:
            # TODO:  If Message Server is changed, disconnect and connect to new Message Server!
            # Update the Global Message Server Variable
            TransanaGlobal.configData.messageServer = self.messageServer.GetValue()
            # Update the Global Message Server Port
            TransanaGlobal.configData.messageServerPort = int(self.messageServerPort.GetValue())
        
        # Make sure the oldDatabaseDir ends with the proper seperator character
        if self.oldDatabaseDir[-1] != os.sep:
            self.oldDatabaseDir = self.oldDatabaseDir + os.sep
        # If database directory was changed inform user
        if self.oldDatabaseDir != TransanaGlobal.configData.databaseDir:
            infoDlg = Dialogs.InfoDialog(self, _("Database directory change will take effect after you restart Transana."))
            infoDlg.ShowModal()
            infoDlg.Destroy()
        self.Close()

    def OnCancel(self, event):
        """ Method attached to the 'Close' Button """
        # Close the Dialog without updating any Global Variables
        self.Close()

    def OnHelp(self, event):
        """Invoked when dialog Help button is activated."""
        if TransanaGlobal.menuWindow != None:
            TransanaGlobal.menuWindow.ControlObject.Help('Program Settings')

    def OnBrowse(self, event):
        """ Implements the "Browse" button for the Waveform or Video Root Directories on the Directories Tab """
        # Set the prompt and current value for the correct Browse operation (Waveform or Video Root folder)
        if event.GetId() == self.btnWaveformBrowse.GetId():
            prompt = _("Choose a Waveforms directory:")
            currentValue = self.waveformDirectory.GetValue()
        elif event.GetId() == self.btnVideoBrowse.GetId():
            prompt = _("Choose a Video Root directory:")
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
            elif event.GetId() == self.btnDatabaseBrowse.GetId():
                self.databaseDirectory.SetValue(dlg.GetPath())
        # Destroy the Dialog
        dlg.Destroy


# REMOVED!  When Speed Control is OFF for WMP, it still works for QuickTime!
#    def OnMediaPlayerSelect(self, event):
#        """ Handle the OnChoice Event for the Media Player Choice box """
        # Disable the slider if it should be disabled
#        if self.chMediaPlayer.GetCurrentSelection() == 0:
#            self.videoSpeed.SetValue(10)
#            self.videoSpeed.Enable(False)
#            self.lblVideoSpeedSetting.SetLabel("%1.1f" % (1.0))
#        else:
#            self.videoSpeed.Enable(True)
#            self.lblVideoSpeedSetting.SetLabel("%1.1f" % (float(self.videoSpeed.GetValue()) / 10))

    def OnScroll(self, event):
        """ Handle the Scroll Event for the Video Speed Slider. """
        # Update the Current Video Speed Label
        self.lblVideoSpeedSetting.SetLabel("%1.1f" % (float(self.videoSpeed.GetValue()) / 10))
