#
# Copyright (C) 2002 The Board of Regents of the University of Wisconsin System 
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

"""This module implements miscellaneous convenience functions."""

__author__ = 'Nathaniel Case <nacase@wisc.edu>, David Woods <dwoods@wcer.wisc.edu>'

import mx.DateTime
import string

def datestr_to_dt(datestr):
    """Construct a DateTime object from a given date string.  This function
    is less strict in formatting requirements than the regular DateTime
    constructors, which is useful when depending on user input.  Return
    None if unable to parse the string given."""

    datestr = string.replace(datestr, "-", "/")
    return mx.DateTime.DateTimeFrom(datestr)
    
    # The following code has been abandoned because strptime() is not
    # available in the Windows C libraries.  So we use DateTimeFrom()
    # as shown above instead.
    
    # This list defines the format orders to attempt to parse.  It will
    # use the first one that matches.  See the strptime() manpage for
    # details on the format string syntax.
    # FIXME: l10n: This list should be localized (don't worry about this
    # since this code isn't actually used anymore, see above)
#    formats = ['%m/%d/%y', '%m-%d-%y', '%m/%d/%Y', '%m-%d-%Y', '%Y-%m-%d',
#               '%Y/%m/%d']

#    for format in formats:
#        try:
#            dt = mx.DateTime.strptime(datestr, format)
#        except mx.DateTime.Error, e:
#            # Given string doesn't match format
#            dt = None
#        else:
#            break
    
#    return dt

def dt_to_datestr(dt):
    """Return a localized date string for the given DateTime object."""
    # FIXME: l10n: Doesn't actually localize yet :)
    try:
        s = "%d/%d/%d" % (dt.month, dt.day, dt.year)
    except AttributeError:
        s = ""

    return s
    # return dt.strftime('%m/%d/%Y')

def time_in_ms_to_str(time_in_ms):
    """ Return a string representation of Time in Milliseconds to H:MM:SS.s format """
    # determine the total number of seconds in the time value
    seconds = int(time_in_ms / 1000)
    # determine the number of hours in that number of seconds
    hours = int(seconds / 3600)
    # determine the number of minutes in the number of seconds left once the hours have been removed
    minutes = int(seconds / 60) - (hours * 60)
    # determine the number of seconds left after the hours and minutes have been removed
    seconds = seconds - (minutes * 60) - (hours * 3600)
    # determine the number of milliseconds left after hours, minutes, and seconds have been removed,
    # and round to the number of TENTHS of a second
    tenthsofasecond = round((time_in_ms - (seconds * 1000) - (minutes * 60000) - (hours * 3600000)) / 100.0)
    # Adjust (round up) for values that are maxed out (These are unlikely, but this is necessary)
    if tenthsofasecond == 10:
        tenthsofasecond = 0
        seconds = seconds + 1
        if seconds == 60:
            seconds = 0
            minutes = minutes + 1
            if minutes == 60:
                minutes = 0
                hours = hours + 1
    # Now build the final string.  "%02d" converts values to 2-character, 0-padded strings.
    str = "%d:%02d:%02d.%d" % (hours, minutes, seconds, tenthsofasecond)
    return str
    
