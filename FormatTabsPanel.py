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

""" This dialog implements the Transana Format Tabs Panel class.   """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Enable (True) or Disable (False) debugging messages
DEBUG = False
if DEBUG:
    print "FormatTabsPanel DEBUG is ON"

# import wxPython
import wx

# import the TransanaGlobal variables
import TransanaGlobal


class FormatTabsPanel(wx.Panel):
    """ Transana's custom Tabs Dialog Box.  Pass in a TransanaFontDef object to allow for ambiguity in the font specification.  """
    
    def __init__(self, parent, formatData):
        """ Initialize the Paragraph Panel. """

        self.formatData = formatData.copy()

        # Create the tabs Panel
        wx.Panel.__init__(self, parent, -1)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the main Sizer, which will hold the boxTop, boxMiddle, and boxButton sizers
        box = wx.BoxSizer(wx.VERTICAL)
        hBox = wx.BoxSizer(wx.HORIZONTAL)

        col1Box = wx.BoxSizer(wx.VERTICAL)
        # Create the label
        lblTabStops = wx.StaticText(self, -1, _('Defined Tabs:'))
        col1Box.Add(lblTabStops, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.LEFT | wx.TOP, 15)
        col1Box.Add((0, 5))  # Spacer
        self.lbTabStops = wx.ListBox(self, -1, style=wx.LB_EXTENDED)
        col1Box.Add(self.lbTabStops, 1, wx.EXPAND | wx.LEFT, 15)

        hBox.Add(col1Box, 1, wx.EXPAND | wx.GROW)

        hBox.Add((15, 1))

        col2Box = wx.BoxSizer(wx.VERTICAL)

        # Create the label
        lblAdd = wx.StaticText(self, -1, _('Tab Stop Value to Add:'))
        col2Box.Add(lblAdd, 0, wx.ALIGN_LEFT | wx.ALIGN_TOP | wx.LEFT | wx.TOP, 15)
        col2Box.Add((0, 5))  # Spacer

        self.txtAdd = wx.TextCtrl(self, -1)
        self.txtAdd.Bind(wx.EVT_CHAR, self.OnNumOnly)
        col2Box.Add(self.txtAdd, 0, wx.EXPAND | wx.TOP | wx.RIGHT, 15)

        self.btnAdd = wx.Button(self, -1, _("Add"))
        self.btnAdd.Bind(wx.EVT_BUTTON, self.OnAdd)
        col2Box.Add(self.btnAdd, 0, wx.EXPAND | wx.TOP | wx.RIGHT, 15)
        self.btnDelete = wx.Button(self, -1, _("Delete"))
        self.btnDelete.Bind(wx.EVT_BUTTON, self.OnDelete)
        col2Box.Add(self.btnDelete, 0, wx.EXPAND | wx.TOP | wx.RIGHT, 15)

        col2Box.Add((1, 1), 1, wx.EXPAND)


        hBox.Add(col2Box, 1, wx.EXPAND | wx.GROW)

        box.Add(hBox, 5, wx.EXPAND)

        unitSizer = wx.BoxSizer(wx.HORIZONTAL)
        lblUnits = wx.StaticText(self, -1, _("Units:"))
        unitSizer.Add(lblUnits, 0, wx.RIGHT, 20)

        self.rbUnitsInches = wx.RadioButton(self, -1, _("inches"), style=wx.RB_GROUP)
        unitSizer.Add(self.rbUnitsInches, 0, wx.RIGHT, 10)

        self.rbUnitsCentimeters = wx.RadioButton(self, -1, _("cm"))
        unitSizer.Add(self.rbUnitsCentimeters, 0)
        box.Add(unitSizer, 0, wx.EXPAND | wx.ALIGN_LEFT | wx.LEFT | wx.TOP | wx.RIGHT | wx.BOTTOM, 15)

        if TransanaGlobal.configData.formatUnits == 'cm':
            self.rbUnitsCentimeters.SetValue(True)
        else:
            self.rbUnitsInches.SetValue(True)

        # If an empty list is passed in ...
        if self.formatData.tabs == []:
            # create default tabs every half an inch
            for loop in range(100, 2160, 100):
                self.formatData.tabs.append(loop)

        tabStops = []
        for loop in self.formatData.tabs:
            tabStops.append(self.ConvertValueToStr(loop))
            
        self.lbTabStops.Set(tabStops)

        self.Bind(wx.EVT_RADIOBUTTON, self.OnIndentUnitSelect)

        # Define box as the form's main sizer
        self.SetSizer(box)
        # Fit the form to the widgets created
        self.Fit()
        # Set this as the minimum size for the form.
        self.SetSizeHints(minW = self.GetSize()[0], minH = self.GetSize()[1])
        # Tell the form to maintain the layout and have it set the intitial Layout
        self.SetAutoLayout(True)
        self.Layout()

    def ConvertValueToStr(self, value):
        if (value == '') or (value == None):
            valStr = ''
        else:
            # Now convert to the appropriate units
            if self.rbUnitsInches.GetValue():
                value = float(value) / 254.0
            else:
                value = float(value) / 100.0
            valStr = "%4.2f" % value
        return valStr

    def OnAdd(self, event):

        valStr = self.txtAdd.GetValue()
        try:
            val = float(valStr)
        except:
            val = 0
        if (val > 0) and (self.lbTabStops.FindString("%4.2f" % val) == -1):
            tmpItems = self.lbTabStops.GetItems()
            tmpItems.append("%4.2f" % val)
            tmpItems.sort()
            self.lbTabStops.SetItems(tmpItems)
            self.txtAdd.SetValue('')

    def OnDelete(self, event):
        sels = self.lbTabStops.GetSelections()
        # On Mac OS X, selections may be in a random order, which breaks the delete.
        # Convert them to a list (from a tuple)
        sels = list(sels)
        # and sort them!
        sels.sort()
        # Delete selections in REVERSE order!!
        for loop in range(len(sels), 0, -1):
            self.lbTabStops.Delete(sels[loop - 1])
        
    def OnIndentUnitSelect(self, event):
        newList = []
        for valStr in self.lbTabStops.GetItems():
            if self.rbUnitsInches.GetValue():
                val = float(valStr) / 2.54
            else:
                val = float(valStr) * 2.54
            newList.append("%4.2f" % val)
        self.lbTabStops.SetItems(newList)

    def OnNumOnly(self, event):
        ctrl = event.GetEventObject()

        shouldSkip = False

        if event.AltDown() or event.CmdDown() or event.ControlDown() or event.MetaDown() or event.ShiftDown():
            event.Skip()

        elif event.GetKeyCode() == ord('.'):

            # If there is no decimal point already, OR if the decimal is inside the current selection, which will be over-written ...
            if (ctrl.GetValue().find('.') == -1) or (ctrl.GetStringSelection().find('.') > -1):
                # ... then it's okay to add a decimal point
                shouldSkip = True

        elif event.GetKeyCode() in [ord('0'), ord('1'), ord('2'), ord('3'), ord('4'), ord('5'), ord('6'), ord('7'), ord('8'), ord('9')]:

            shouldSkip = True

        elif event.GetKeyCode() in [wx.WXK_LEFT, wx.WXK_RIGHT, wx.WXK_BACK, wx.WXK_DELETE]:

            event.Skip()

        if shouldSkip:

            event.Skip()

