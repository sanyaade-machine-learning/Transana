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

"""This module implements the Progress Dialog for Waveform Creation."""

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

import wx      # import wxPython
import os, sys

ID_BTNCANCEL    =  wx.NewId()

class WaveformProgress(wx.Dialog):
    """ This class implements the Progress Dialog for Waveform Creation. """

    def __init__(self, parent, waveFilename, label=''):
        """ Initialize the Progress Dialog """
        # Remember the filename
        self.waveFilename = waveFilename

        # Define Cancel variable
        self.userCancel = False
        
        # Define the Dialog Box.  wx.STAY_ON_TOP required because this dialog can't be modal, but shouldn't be hidden or worked behind.
        wx.Dialog.__init__(self, parent, -1, _('Wave Extraction Progress'), size=(400, 160), style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # For now, place an empty StaticText label on the Dialog
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self, wx.Top, 10)
        lay.height.AsIs()
        lay.left.SameAs(self, wx.Left, 10)
        lay.right.SameAs(self, wx.Right, 10)
        self.lbl = wx.StaticText(self, -1, label)
        self.lbl.SetConstraints(lay)

        # Progress Bar
        lay = wx.LayoutConstraints()
        lay.top.Below(self.lbl, 10)
        lay.left.SameAs(self, wx.Left, 10)
        lay.height.AsIs()
        lay.right.SameAs(self, wx.Right, 10)
        self.progressBar = wx.Gauge(self, -1, 100, style=wx.GA_HORIZONTAL | wx.GA_SMOOTH)
        self.progressBar.SetConstraints(lay)

        # Seconds Processed label
        lay = wx.LayoutConstraints()
        lay.top.Below(self.progressBar, 10)
        lay.left.SameAs(self, wx.Left, 10)
        lay.height.AsIs()
        lay.width.PercentOf(self.progressBar, wx.Width, 70)
        self.lblTime = wx.StaticText(self, -1, _("%d:%02d:%02d processed") % (0, 0, 0), style=wx.ST_NO_AUTORESIZE)
        self.lblTime.SetConstraints(lay)

        # Percent Processed label
        lay = wx.LayoutConstraints()
        lay.top.Below(self.progressBar, 10)
        lay.right.SameAs(self, wx.Right, 10)
        lay.height.AsIs()
        lay.width.PercentOf(self.progressBar, wx.Width, 30)
        self.lblPercent = wx.StaticText(self, -1, "%d %%" % 0, style=wx.ST_NO_AUTORESIZE | wx.ALIGN_RIGHT)
        self.lblPercent.SetConstraints(lay)

        # Place a Cancel button on the Dialog
        lay = wx.LayoutConstraints()
        lay.bottom.SameAs(self, wx.Bottom, 10)
        lay.height.AsIs()
        lay.centreX.SameAs(self, wx.CentreX, 0)
        lay.width.AsIs()
        self.btnCancel = wx.Button(self, ID_BTNCANCEL, _("Cancel"))
        self.btnCancel.SetConstraints(lay)

        wx.EVT_BUTTON(self, ID_BTNCANCEL, self.OnCancel)

        # Call Layout to "place" the widgits
        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()
        

    def Update(self, percent, seconds):
        """ This method allows the contents of this form to be "updated" in the Wave Extraction Callback process """
        # Update the Progress Bar
        self.progressBar.SetValue(percent)
        # Calculate the time values to be displayed
        timeProcessed = seconds
        secondsProcessed = (timeProcessed % 60)
        timeProcessed = timeProcessed - secondsProcessed
        hoursProcessed = (timeProcessed / (60 * 60))
        timeProcessed = timeProcessed - (hoursProcessed * 60 * 60)
        minutesProcessed = (timeProcessed / 60)
        # Update the amount of time processed
        self.lblTime.SetLabel(_("%d:%02d:%02d processed") % (hoursProcessed, minutesProcessed, secondsProcessed))
        # Display % processed
        self.lblPercent.SetLabel("%d %%" % percent)
        # wxYield tells the OS to process all messages in the queue, such as GUI updates, so the label change will show up
        wx.Yield()
        # User Cancel is implemented through the callback function return value
        if self.userCancel:
            return 0
        else:
            return 1

    def OnCancel(self, event):
        self.userCancel = True
