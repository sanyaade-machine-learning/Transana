# Copyright (C) 2002 - 2006 The Board of Regents of the University of Wisconsin System 
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

"""This module contains custom Exception classes for Transana.<BR><BR>

class RecordLockedError(exceptions.Exception):<BR>
    Raised when a database operation fails because the record is locked.<BR><BR>

class RecordNotFoundError(exceptions.Exception):<BR>
    Raised when the specified record was not found in the database.<BR><BR>

class SaveError(exceptions.Exception):<BR>
    Raised when a record save attempt fails.<BR><BR>

class DeleteError(exceptions.Exception):<BR>
    Raised when a record delete attempt fails.<BR><BR>

class InvalidLockError(SaveError):<BR>
    Raised when a record lock that is no longer valid halts the save.<BR><BR>

class NotImplementedError(exceptions.Exception):<BR>
    Raised when a feature is not yet implemented.<BR><BR>

class ImageLoadError(exceptions.Exception):<BR>
    Raised when a image file was not successfully loaded.<BR><BR>

class ProgrammingError(exceptions.Exception):<BR>
    Raised when program reaches invalid state due to programming error.<BR><BR>

class GeneralError(exceptions.Exception):<BR>
    General error message.<BR><BR>
"""

__author__ = 'Nathaniel Case <nacase@wisc.edu>, David K. Woods <dwoods@wcer.wisc.edu>'

# Import Python's gettext module
#import gettext
# Define the gettext function.  I'm not sure why this module requires it when
# others don't, but it does.  DKW
#_ = gettext.gettext
import wx
_ = wx.GetTranslation


import exceptions

class RecordLockedError(exceptions.Exception):
    """Raised when a database operation fails because the record is locked."""
    def __init__(self, user=None):
        self.user = user
        self.args = _("Database operation failed due to record lock by %s") % user

class RecordNotFoundError(exceptions.Exception):
    """Raised when the specified record was not found in the database."""
    def __init__(self, record, rowcount):
        self.record = record    # name, number, etc.
        self.rowcount = rowcount
        self.args = _("Record not found in database.")
        
class SaveError(exceptions.Exception):
    """Raised when a record save attempt fails."""
    def __init__(self, reason):
        if ('unicode' in wx.PlatformInfo) and isinstance(reason, str):
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            reason = unicode(reason, 'utf8')
        prompt = _("Unable to save.  %s")
        self.args = prompt % reason
        self.reason = reason

class DeleteError(exceptions.Exception):
    """Raised when a record delete attempt fails."""
    def __init__(self, reason):
        self.reason = reason
        self.args = _("Unable to delete.  %s") % reason

class InvalidLockError(SaveError):
    """Raised when a record lock that is no longer valid halts the save."""
    def __init__(self):
        SaveError.__init__(self, _("Record lock no longer valid."))

class NotImplementedError(exceptions.Exception):
    """Raised when a feature is not yet implemented."""
    def __init__(self):
        self.args = _("This feature is not yet implemented.")

class ImageLoadError(exceptions.Exception):
    """Raised when a image file was not successfully loaded."""
    def __init__(self, args=_("Unable to load image file.")):
        self.args = args

class ProgrammingError(exceptions.Exception):
    """Raised when program reaches invalid state due to programming error."""
    def __init__(self, args=_("Programming error.")):
        self.args = args

class GeneralError(exceptions.Exception):
    """General error message."""
    def __init__(self, args=_("General error")):
        self.args = args
