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

# Code Review and Documentation completed by DKW on 11/5/2003

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
imageOptions = ['', 'avi', 'mpeg', 'mpeg2']
soundOptions = ['', 'mp3', 'wav']
allOptions   = ['', 'avi', 'mp3', 'mpeg', 'mpeg2', 'wav']


class CoreDataPropertiesForm(Dialogs.GenForm):
    """ Form containing Core Data fields. """


    # TODO:  Fix Tab Order
    
    def __init__(self, parent, id, title, coredata_object):
        """ Initialize the Core Data Properties form """
        # Set the Size of the Form
        self.width = 550
        self.height = 510
        # Create the Base Form, based on GenForm in the Transana Dialogs module
        Dialogs.GenForm.__init__(self, parent, id, title, size=(self.width, self.height), HelpContext='Core Data Properties')

        # Preserve the initial Core Data Object sent in from the calling routine
        self.obj = coredata_object

        # Media File Name layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)         # 10 from Form top
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from Form right
        lay.height.AsIs()
        # NOTE:  The otherwise unused Comment Field is used to contain the full Media File Name.
        media_filename_edit = self.new_edit_box(_("Media File Name"), lay, self.obj.comment)
        # Media File Name Field is not editable
        media_filename_edit.Enable(False)

        # Title layout
        lay = wx.LayoutConstraints()
        lay.top.Below(media_filename_edit, 10)     # 10 under Media File Name
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from Form right
        lay.height.AsIs()
        title_edit = self.new_edit_box(_("Title"), lay, self.obj.title)

        # Creator layout
        lay = wx.LayoutConstraints()
        lay.top.Below(title_edit, 10)              # 10 under Title
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form Left
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        creator_edit = self.new_edit_box(_("Creator"), lay, self.obj.creator)

        # Subject layout
        lay = wx.LayoutConstraints()
        lay.top.Below(title_edit, 10)              # 10 under Title
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from Form Right
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        subject_edit = self.new_edit_box(_("Subject"), lay, self.obj.subject)

        # Description layout
        lay = wx.LayoutConstraints()
        lay.top.Below(creator_edit, 10)   # 10 under Description
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from Form Right
        lay.height.AsIs()
        self.description_edit = self.new_edit_box(_("Description"), lay, self.obj.description, style=wx.TE_MULTILINE)

        # Publisher layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.description_edit, 10)   # 10 under Description
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form left
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        publisher_edit = self.new_edit_box(_("Publisher"), lay, self.obj.publisher)

        # Contributor layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.description_edit, 10)   # 10 under Description
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from Form Right
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        contributor_edit = self.new_edit_box(_("Contributor"), lay, self.obj.contributor)

        # Dialogs.GenForm does not provide a Masked text control, so the Date
        # Field is handled differently than other fields.
        
        # Date layout [label]
        lay = wx.LayoutConstraints()
        lay.top.Below(publisher_edit, 10)          # 10 under Publisher
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form left
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from Form Right
        lay.height.AsIs()
        date_lbl = wx.StaticText(self.panel, -1, _("Date (MM/DD/YYYY)"))
        date_lbl.SetConstraints(lay)

        # Date layout
        lay = wx.LayoutConstraints()
        lay.top.Below(date_lbl, 3)                 #  3 under Date Label
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form left
        lay.width.PercentOf(self.panel, wx.Width, 31)  # 31% width, allows for three across positioning
        lay.height.AsIs()
        # Use the Masked Text Control (Requires wxPython 2.4.2.4 or later)
        # TODO:  Make Date autoformat localizable
        self.date_edit = wx.lib.masked.TextCtrl(self.panel, -1, '', autoformat='USDATEMMDDYYYY/')
        # If a Date is know, load it into the control
        if (self.obj.dc_date != '') and (self.obj.dc_date != '01/01/0'):
            self.date_edit.SetValue(self.obj.dc_date)
        self.date_edit.SetConstraints(lay)

        # Type Layout
        lay = wx.LayoutConstraints()
        lay.top.Below(publisher_edit, 10)          # 10 under Publisher
        lay.left.RightOf(self.date_edit, 10)       # 10 to the right of Date
        lay.width.PercentOf(self.panel, wx.Width, 31)  # 31% width, allows for three across positioning
        lay.height.AsIs()
        # Define legal options for the Combo Box
        options = ['', _('Image'), _('Sound')]
        self.type_combo = self.new_combo_box(_("Media Type"), lay, options, self.obj.dc_type)
        # Define a Combo Box Event.  When different Types are selected, different Format options should be displayed.
        wx.EVT_COMBOBOX(self, self.type_combo.GetId(), self.OnTypeChoice)

        # Format Layout
        lay = wx.LayoutConstraints()
        lay.top.Below(publisher_edit, 10)          # 10 under Publisher
        lay.left.RightOf(self.type_combo, 10)      # 10 to the right of Type
        lay.width.PercentOf(self.panel, wx.Width, 31)  # 31% width, allows for three across positioning
        lay.height.AsIs()
        # The Format Combo Box has different options depending on the value of Type
        if self.obj.dc_type == _('Image'):
            options = imageOptions
        elif self.obj.dc_type == _('Sound'):
            options = soundOptions
        else:
            options = allOptions
        self.format_combo = self.new_combo_box(_("Format"), lay, options, self.obj.format)

        # Identifier layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.date_edit, 10)          # 10 under Date
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form left
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        identifier_edit = self.new_edit_box(_("Identifier"), lay, self.obj.id)
        # Identifier is not editable
        identifier_edit.Enable(False)

        # Source layout
        lay = wx.LayoutConstraints()
        lay.top.Below(self.date_edit, 10)          # 10 under Date
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from Form Right
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        source_edit = self.new_edit_box(_("Source"), lay, self.obj.source)

        # Language layout
        lay = wx.LayoutConstraints()
        lay.top.Below(identifier_edit, 10)         # 10 under Identifier
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form left
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        language_edit = self.new_edit_box(_("Language"), lay, self.obj.language)

        # Relation layout
        lay = wx.LayoutConstraints()
        lay.top.Below(source_edit, 10)             # 10 under Source
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from Form Right
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        relation_edit = self.new_edit_box(_("Relation"), lay, self.obj.relation)

        # Coverage layout
        lay = wx.LayoutConstraints()
        lay.top.Below(language_edit, 10)           # 10 under Language
        lay.left.SameAs(self.panel, wx.Left, 10)       # 10 from Form left
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        coverage_edit = self.new_edit_box(_("Coverage"), lay, self.obj.coverage)

        # Rights layout
        lay = wx.LayoutConstraints()
        lay.top.Below(language_edit, 10)           # 10 under Language
        lay.right.SameAs(self.panel, wx.Right, 10)     # 10 from Form Right
        lay.width.PercentOf(self.panel, wx.Width, 47)  # 47% width, allows for two across positioning
        lay.height.AsIs()
        rights_edit = self.new_edit_box(_("Rights"), lay, self.obj.rights)

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

        title_edit.SetFocus()

        
    def OnTypeChoice(self, event):
        """ This method changes the options in the Format field based on the selection made in the Type field. """
        # Get the current Text Value for the Format field
        currentString = self.format_combo.GetValue()

        # Determine the appropriate options for the Format Combo based on the current value of the Type Field
        if self.type_combo.GetStringSelection() == _('Image'):
            options = imageOptions
        elif self.type_combo.GetStringSelection() == _('Sound'):
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
