# Copyright (C) 2002 - 2009 The Board of Regents of the University of Wisconsin System 
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

def ReportRecordLockedException(rtype, id, e):<BR>
    Handles the reporting of Record Lock Exceptions consistently.<BR><BR>
"""

__author__ = 'Nathaniel Case <nacase@wisc.edu>, David K. Woods <dwoods@wcer.wisc.edu>'

import wx


import exceptions

class RecordLockedError(exceptions.Exception):
    """Raised when a database operation fails because the record is locked."""
    def __init__(self, user=None):
        self.user = user
        prompt = _("Database operation failed due to record lock by %s")
        if ('unicode' in wx.PlatformInfo) and isinstance(prompt, str):
            prompt = unicode(prompt, 'utf8')
        self.explanation = prompt % user

class RecordNotFoundError(exceptions.Exception):
    """Raised when the specified record was not found in the database."""
    def __init__(self, record, rowcount):
        self.record = record    # name, number, etc.
        self.rowcount = rowcount
        self.explanation = _("Record not found in database.")
        
class SaveError(exceptions.Exception):
    """Raised when a record save attempt fails."""
    def __init__(self, reason):
        if ('unicode' in wx.PlatformInfo):
            if isinstance(reason, str):
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                reason = unicode(reason, 'utf8')
            prompt = unicode(_("Unable to save.  %s"), 'utf8')
        else:
            prompt = _("Unable to save.  %s")
        self.explanation = prompt % reason
        self.reason = reason

class DeleteError(exceptions.Exception):
    """Raised when a record delete attempt fails."""
    def __init__(self, reason):
        self.reason = reason
        prompt = _("Unable to delete.  %s")
        if 'unicode' in wx.PlatformInfo:
            prompt = unicode(prompt, 'utf8')
        self.explanation = prompt % reason

class InvalidLockError(SaveError):
    """Raised when a record lock that is no longer valid halts the save."""
    def __init__(self):
        SaveError.__init__(self, _("Record lock no longer valid."))

class NotImplementedError(exceptions.Exception):
    """Raised when a feature is not yet implemented."""
    def __init__(self):
        self.explanation = _("This feature is not yet implemented.")

class ImageLoadError(exceptions.Exception):
    """Raised when a image file was not successfully loaded."""
    def __init__(self, explanation=_("Unable to load image file.")):
        self.explanation = explanation

class ProgrammingError(exceptions.Exception):
    """Raised when program reaches invalid state due to programming error."""
    def __init__(self, explanation=_("Programming error.")):
        self.explanation = explanation

class GeneralError(exceptions.Exception):
    """General error message."""
    def __init__(self, explanation=_("General error")):
        self.explanation = explanation

def ReportRecordLockedException(rtype, idVal, e):
    """ Report a RecordLocked exception """
    if 'unicode' in wx.PlatformInfo:
        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
        msg = unicode(_('You cannot proceed because you cannot obtain a lock on %s "%s"') + \
                      _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
        if isinstance(rtype, str):
            rtype = unicode(rtype, 'utf8')
        if isinstance(idVal, str):
            id = unicode(idVal, 'utf8')
    else:
        msg = _('You cannot proceed because you cannot obtain a lock on %s "%s"') + \
              _('.\nThe record is currently locked by %s.\nPlease try again later.')

    import Dialogs
    dlg = Dialogs.ErrorDialog(None, msg % (rtype, idVal, e.user))
    dlg.ShowModal()
    dlg.Destroy()
