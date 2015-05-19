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

"""This module implements the Core Data Properties Form."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Import Python's Internationalization Module
import gettext

# Import the CoreData Object
import CoreData
# Import Transana's Database Interface
import DBInterface
# Import the Dialogs module from which this form inherits
import Dialogs

# import wxPython
import wx
# import the Masked Edit Control (requires wxPython 2.4.2.4 or later)
import wx.lib.masked

# Define the Format Options for different File Types
imageOptions = ['', 'avi', 'mov', 'mp4', 'm4v', 'mpeg', 'mpeg2', 'wmv']
soundOptions = ['', 'mp3', 'wav', 'wma', 'aac']
allOptions   = ['', 'avi', 'mov', 'mp3', 'mp4', 'm4v', 'mpeg', 'mpeg2', 'wav', 'wma', 'wmv', 'aac']


class CoreDataPropertiesForm(Dialogs.GenForm):
    """ Form containing Core Data fields. """
    def __init__(self, parent, id, title, coredata_object):
        """ Initialize the Core Data Properties form """
        # Set the Size of the Form
        self.width = 550
        self.height = 510
        # Create the Base Form, based on GenForm in the Transana Dialogs module
        Dialogs.GenForm.__init__(self, parent, id, title, size=(self.width, self.height), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Core Data Properties')

        # Preserve the initial Core Data Object sent in from the calling routine
        self.obj = coredata_object

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Media File Name
        # NOTE:  The otherwise unused Comment Field is used to contain the full Media File Name.
        media_filename_edit = self.new_edit_box(_("Media File Name"), v1, self.obj.comment)
        # Add the element to the row sizer
        r1Sizer.Add(v1, 1, wx.EXPAND)
        # Media File Name Field is not editable
        media_filename_edit.Enable(False)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Title
        title_edit = self.new_edit_box(_("Title"), v2, self.obj.title, maxLen=255)
        # Add the element to the row sizer
        r2Sizer.Add(v2, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v3 = wx.BoxSizer(wx.VERTICAL)
        # Creator
        creator_edit = self.new_edit_box(_("Creator"), v3, self.obj.creator, maxLen=255)
        # Add the element to the row sizer
        r3Sizer.Add(v3, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r3Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        # Subject
        subject_edit = self.new_edit_box(_("Subject"), v4, self.obj.subject, maxLen=255)
        # Add the element to the row sizer
        r3Sizer.Add(v4, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r3Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r4Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v5 = wx.BoxSizer(wx.VERTICAL)
        # Description
        self.description_edit = self.new_edit_box(_("Description"), v5, self.obj.description, style=wx.TE_MULTILINE, maxLen=255)
        # Add the element to the row sizer
        r4Sizer.Add(v5, 1, wx.EXPAND)

        # Add a spacer to enforce the height of the Description item
        r4Sizer.Add((0, 80))

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r4Sizer, 5, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r5Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v6 = wx.BoxSizer(wx.VERTICAL)
        # Publisher
        publisher_edit = self.new_edit_box(_("Publisher"), v6, self.obj.publisher, maxLen=255)
        # Add the element to the row sizer
        r5Sizer.Add(v6, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r5Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v7 = wx.BoxSizer(wx.VERTICAL)
        # Contributor
        contributor_edit = self.new_edit_box(_("Contributor"), v7, self.obj.contributor, maxLen=255)
        # Add the element to the row sizer
        r5Sizer.Add(v7, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r5Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r6Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v8 = wx.BoxSizer(wx.VERTICAL)
        # Dialogs.GenForm does not provide a Masked text control, so the Date
        # Field is handled differently than other fields.
        
        # Date [label]
        date_lbl = wx.StaticText(self.panel, -1, _("Date (MM/DD/YYYY)"))
        # Add the element to the vertical sizer
        v8.Add(date_lbl, 0, wx.BOTTOM, 3)

        # Date layout
        # Use the Masked Text Control (Requires wxPython 2.4.2.4 or later)
        # TODO:  Make Date autoformat localizable
        self.date_edit = wx.lib.masked.TextCtrl(self.panel, -1, '', autoformat='USDATEMMDDYYYY/')
        # Add the element to the vertical sizer
        v8.Add(self.date_edit, 1, wx.EXPAND)
        # If a Date is know, load it into the control
        if (self.obj.dc_date != '') and (self.obj.dc_date != '01/01/0'):
            self.date_edit.SetValue(self.obj.dc_date)

        # Add the element to the row sizer
        r6Sizer.Add(v8, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r6Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v9 = wx.BoxSizer(wx.VERTICAL)
        # Type
        # Define legal options for the Combo Box
        options = ['', _('Image'), _('Sound')]
        self.type_combo = self.new_combo_box(_("Media Type"), v9, options, self.obj.dc_type)
        # Add the element to the row sizer
        r6Sizer.Add(v9, 1, wx.EXPAND)
        # Define a Combo Box Event.  When different Types are selected, different Format options should be displayed.
        wx.EVT_COMBOBOX(self, self.type_combo.GetId(), self.OnTypeChoice)

        # Add a horizontal spacer to the row sizer        
        r6Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v10 = wx.BoxSizer(wx.VERTICAL)
        # Format
        # The Format Combo Box has different options depending on the value of Type
        if self.obj.dc_type == unicode(_('Image'), 'utf8'):
            options = imageOptions
        elif self.obj.dc_type == unicode(_('Sound'), 'utf8'):
            options = soundOptions
        else:
            options = allOptions
        self.format_combo = self.new_combo_box(_("Format"), v10, options, self.obj.format)
        # Add the element to the row sizer
        r6Sizer.Add(v10, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r6Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r7Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v11 = wx.BoxSizer(wx.VERTICAL)
        # Identifier
        identifier_edit = self.new_edit_box(_("Identifier"), v11, self.obj.id)
        # Add the element to the row sizer
        r7Sizer.Add(v11, 1, wx.EXPAND)
        # Identifier is not editable
        identifier_edit.Enable(False)

        # Add a horizontal spacer to the row sizer        
        r7Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v12 = wx.BoxSizer(wx.VERTICAL)
        # Source
        source_edit = self.new_edit_box(_("Source"), v12, self.obj.source, maxLen=255)
        # Add the element to the row sizer
        r7Sizer.Add(v12, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r7Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r8Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v13 = wx.BoxSizer(wx.VERTICAL)
        # Language
        language_edit = self.new_edit_box(_("Language"), v13, self.obj.language, maxLen=25)
        # Add the element to the row sizer
        r8Sizer.Add(v13, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r8Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v14 = wx.BoxSizer(wx.VERTICAL)
        # Relation
        relation_edit = self.new_edit_box(_("Relation"), v14, self.obj.relation, maxLen=255)
        # Add the element to the row sizer
        r8Sizer.Add(v14, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r8Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the first row
        r9Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v15 = wx.BoxSizer(wx.VERTICAL)
        # Coverage
        coverage_edit = self.new_edit_box(_("Coverage"), v15, self.obj.coverage, maxLen=255)
        # Add the element to the row sizer
        r9Sizer.Add(v15, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r9Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v16 = wx.BoxSizer(wx.VERTICAL)
        # Rights
        rights_edit = self.new_edit_box(_("Rights"), v16, self.obj.rights, maxLen=255)
        # Add the element to the row sizer
        r9Sizer.Add(v16, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r9Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the buttons
        self.create_buttons(sizer=btnSizer)
        # Add the button sizer to the main sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            mainSizer.Add((0, 2))

        # Set the PANEL's main sizer
        self.panel.SetSizer(mainSizer)
        # Tell the PANEL to auto-layout
        self.panel.SetAutoLayout(True)
        # Lay out the Panel
        self.panel.Layout()
        # Lay out the panel on the form
        self.Layout()
        # Resize the form to fit the contents
        self.Fit()

        # Get the new size of the form
        (width, height) = self.GetSizeTuple()
        # Reset the form's size to be at least the specified minimum width
        self.SetSize(wx.Size(max(self.width, width), height))
        # Define the minimum size for this dialog as the current size
        self.SetSizeHints(max(self.width, width), height)
        # Center the form on screen
        self.CenterOnScreen()

        # Set focus to Title
        title_edit.SetFocus()

        
    def OnTypeChoice(self, event):
        """ This method changes the options in the Format field based on the selection made in the Type field. """
        # Get the current Text Value for the Format field
        currentString = self.format_combo.GetValue()

        # Determine the appropriate options for the Format Combo based on the current value of the Type Field
        if self.type_combo.GetStringSelection() == unicode(_('Image'), 'utf8'):
            options = imageOptions
        elif self.type_combo.GetStringSelection() == unicode(_('Sound'), 'utf8'):
            options = soundOptions
        else:
            options = allOptions

        # Clear the values from the Format Combo
        self.format_combo.Clear()

        # Insert the New Values for the Format Combo
        for option in options:
            self.format_combo.Append(option)

        # Reset the String Value for the Format Combo to its original value
        if currentString != '':
            self.format_combo.SetValue(currentString)
        


    def get_input(self):
        """Show the dialog and return the modified Core Data Object.  Result
        is None if user pressed the Cancel button."""
        d = Dialogs.GenForm.get_input(self)     # inherit parent method
        if d:
            self.obj.id = d[_('Identifier')]
            self.obj.comment = d[_('Media File Name')]
            self.obj.title = d[_('Title')]
            self.obj.creator = d[_('Creator')]
            self.obj.subject = d[_('Subject')]
            # Description Field does not utilize Dialogs.GenForm.get_input()
            self.obj.description = self.description_edit.GetValue()
            self.obj.publisher = d[_('Publisher')]
            self.obj.contributor = d[_('Contributor')]
            # TODO:  Block "Okay" if self.date_edit.IsValid() is false!
            # Date Field does not utilize Dialogs.GenForm.get_input()
            self.obj.dc_date = self.date_edit.GetValue()
            self.obj.dc_type = d[_('Media Type')]
            self.obj.format = d[_('Format')]
            self.obj.source = d[_('Source')]
            self.obj.language = d[_('Language')]
            self.obj.relation = d[_('Relation')]
            self.obj.coverage = d[_('Coverage')]
            self.obj.rights = d[_('Rights')]
        else:
            self.obj = None

        return self.obj
        
class AddCoreDataDialog(CoreDataPropertiesForm):
    """Dialog used when adding a new Core Data record."""

    def __init__(self, parent, id):
        # Create a blank Core Data Object
        obj = CoreData.CoreData()
        CoreDataPropertiesForm.__init__(self, parent, id, _("Add Core Data"), obj)


class EditCoreDataDialog(CoreDataPropertiesForm):
    """Dialog used when editing Core Data properties."""

    def __init__(self, parent, id, coredata_object):
        CoreDataPropertiesForm.__init__(self, parent, id, _("Core Data Properties"), coredata_object)
