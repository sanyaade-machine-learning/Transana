# -*- coding: cp1252 -*-
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

"""This module implements miscellaneous convenience functions."""

__author__ = 'Nathaniel Case, David Woods <dwoods@wcer.wisc.edu>'
# Patch sent by David Fraser to eliminate need for mx module

#import mx.DateTime
import string, sys

#def datestr_to_dt(datestr):
#    """Construct a DateTime object from a given date string.  This function
#    is less strict in formatting requirements than the regular DateTime
#    constructors, which is useful when depending on user input.  Return
#    None if unable to parse the string given."""
#    if datestr == '':
#        return None
#    else:
#        datestr = string.replace(datestr, "-", "/")
        # return mx.DateTime.DateTimeFrom(datestr)
#        timetuple = time.strptime(datestr, "%m/%d/%Y")
#        return datetime.datetime(*timetuple[:7])

def dt_to_datestr(dt):
    """Return a localized date string for the given DateTime object."""
    # FIXME: l10n: Doesn't actually localize yet :)
    try:
        s = "%d/%d/%d" % (dt.month, dt.day, dt.year)
    except AttributeError:
        s = ""
    return s

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


def TimeMsToStr(TimeVal):
    """ Converts Time in Milliseconds to a formatted string """
    # Convert from milliseconds to seconds
    seconds = int(TimeVal) / 1000
    # Determine how many whole hours there are in this number of seconds
    hours = seconds / (60 * 60)
    # Remove the hours from the total number of seconds
    seconds = seconds % (60 * 60)
    # Determine the number of minutes in the remaining seconds
    minutes = seconds / 60
    # Remove the minutes from the seconds, and round to the nearest second.
    seconds = round(seconds % 60)
    # Initialize a String
    TempStr = ''
    # Convert to string
    if hours > 0:
        TempStr = '%s:%02d:%02d' % (hours, minutes, seconds)
    else:
        TempStr = '%s:%02d' % (minutes, seconds)
    # Return the string representation to the calling function
    return TempStr

def time_in_str_to_ms(time_in_str):
    """ Returns the number of milliseconds represented by a string in H:M:S.s format,
        or -1 if the string is not properly formatted """

    # Initialize variables for conversion
    seconds = 0
    minutes = 0
    hours = 0

    # Divide the time string into components
    timeParts = time_in_str.split(':')
    try:
        # Convert seconds
        seconds = float(timeParts[-1])
        # If there are minutes ...
        if len(timeParts) > 1:
            # ... convert the minutes
            minutes = int(timeParts[-2])
        # If there are hours ...
        if len(timeParts) > 2:
            # ... convert the hours
            hours = int(timeParts[-3])
        # Convert values to milliseconds and return the result
        return (hours * 60 * 60  * 1000) + (minutes * 60 * 1000) + int((seconds * 1000))
    except:
        # If a problem arises in conversion, return -1 to signal a conversion failure
        # print "exception", sys.exc_info()[0], sys.exc_info()[1]
        return -1


def convertMacFilename(filename):
    """ Mac Filenames use a different encoding system.  We need to adjust the string returned by certain wx.Widgets.
        Surely there's an easier way, but I can't figure it out. """
    filename = filename.replace(unichr(65) + unichr(768), unichr(192))  # Capital A with Grave
    filename = filename.replace(unichr(65) + unichr(769), unichr(193))  # Capital A with Acute
    filename = filename.replace(unichr(65) + unichr(770), unichr(194))  # Capital A with Circumflex
    filename = filename.replace(unichr(65) + unichr(771), unichr(195))  # Capital A with Tilde
    filename = filename.replace(unichr(65) + unichr(776), unichr(196))  # Capital A with Umlaut
    filename = filename.replace(unichr(65) + unichr(778), unichr(197))  # Capital A with Ring Above

    filename = filename.replace(unichr(67) + unichr(807), unichr(199))  # Capital C with Cedila

    filename = filename.replace(unichr(69) + unichr(768), unichr(200))  # Capital E with Grave
    filename = filename.replace(unichr(69) + unichr(769), unichr(201))  # Capital E with Acute
    filename = filename.replace(unichr(69) + unichr(770), unichr(202))  # Capital E with Circumflex
    filename = filename.replace(unichr(69) + unichr(776), unichr(203))  # Capital E with Umlaut

    filename = filename.replace(unichr(73) + unichr(768), unichr(204))  # Capital I with Grave
    filename = filename.replace(unichr(73) + unichr(769), unichr(205))  # Capital I with Acute
    filename = filename.replace(unichr(73) + unichr(770), unichr(206))  # Capital I with Circumflex
    filename = filename.replace(unichr(73) + unichr(776), unichr(207))  # Capital I with Umlaut

    filename = filename.replace(unichr(78) + unichr(771), unichr(209))  # Capital N with Tilde

    filename = filename.replace(unichr(79) + unichr(768), unichr(210))  # Capital O with Grave
    filename = filename.replace(unichr(79) + unichr(769), unichr(211))  # Capital O with Acute
    filename = filename.replace(unichr(79) + unichr(770), unichr(212))  # Capital O with Circumflex
    filename = filename.replace(unichr(79) + unichr(771), unichr(213))  # Capital O with Tilde
    filename = filename.replace(unichr(79) + unichr(776), unichr(214))  # Capital O with Umlaut

    filename = filename.replace(unichr(85) + unichr(768), unichr(217))  # Capital U with Grave
    filename = filename.replace(unichr(85) + unichr(769), unichr(218))  # Capital U with Acute
    filename = filename.replace(unichr(85) + unichr(770), unichr(219))  # Capital U with Circumflex
    filename = filename.replace(unichr(85) + unichr(776), unichr(220))  # Capital U with Umlaut

    filename = filename.replace(unichr(97) + unichr(768), unichr(224))  # a with Grave
    filename = filename.replace(unichr(97) + unichr(769), unichr(225))  # a with Acute
    filename = filename.replace(unichr(97) + unichr(770), unichr(226))  # a with Circumflex
    filename = filename.replace(unichr(97) + unichr(771), unichr(227))  # a with Tilde
    filename = filename.replace(unichr(97) + unichr(776), unichr(228))  # a with Umlaut
    filename = filename.replace(unichr(97) + unichr(778), unichr(229))  # a with Ring Above

    filename = filename.replace(unichr(99) + unichr(807), unichr(231))  # c with Cedila

    filename = filename.replace(unichr(101) + unichr(768), unichr(232))  # e with Grave
    filename = filename.replace(unichr(101) + unichr(769), unichr(233))  # e with Acute
    filename = filename.replace(unichr(101) + unichr(770), unichr(234))  # e with Circumflex
    filename = filename.replace(unichr(101) + unichr(776), unichr(235))  # e with Umlaut

    filename = filename.replace(unichr(105) + unichr(768), unichr(236))  # i with Grave
    filename = filename.replace(unichr(105) + unichr(769), unichr(237))  # i with Acute
    filename = filename.replace(unichr(105) + unichr(770), unichr(238))  # i with Circumflex
    filename = filename.replace(unichr(105) + unichr(776), unichr(239))  # i with Umlaut

    filename = filename.replace(unichr(110) + unichr(771), unichr(241))  # n with Tilde

    filename = filename.replace(unichr(111) + unichr(768), unichr(242))  # o with Grave
    filename = filename.replace(unichr(111) + unichr(769), unichr(243))  # o with Acute
    filename = filename.replace(unichr(111) + unichr(770), unichr(244))  # o with Circumflex
    filename = filename.replace(unichr(111) + unichr(771), unichr(245))  # o with Tilde
    filename = filename.replace(unichr(111) + unichr(776), unichr(246))  # o with Umlaut

    filename = filename.replace(unichr(117) + unichr(768), unichr(249))  # u with Grave
    filename = filename.replace(unichr(117) + unichr(769), unichr(250))  # u with Acute
    filename = filename.replace(unichr(117) + unichr(770), unichr(251))  # u with Circumflex
    filename = filename.replace(unichr(117) + unichr(776), unichr(252))  # u with Umlaut

    filename = filename.replace(unichr(121) + unichr(776), unichr(255))  # y with Umlaut
    return filename


# There a problem in Python using string.strip() with strings encoded with UTF-8.
# To demonstrate, run the following lines:
#
#    s = u'Matti\xe0'
#    t = s.encode('utf8')
#
# s will equal u'Matti\xe0' and t will equal 'Matti\xc3\xa0'.
#
#    print s, t.decode('utf8')
#
# will print:
#
#    Mattià Mattià
#
# which is correct.
#
# The problem is that t.strip() will reduce 'Matti\xc3\xa0' to 'Matti\xc3'.  If you do that,
# you can no longer decode t with UTF-8.  The string is broken and irretrievable.
#
# This method is a replacement for string.strip() that avoids this problem with UTF-8 strings.
def unistrip(strng):
    """ A unicode-safe replacement for string.strip() """
    # We can safely use lstrip() to remove whitespace from the left side of the string.
    strng = strng.lstrip()
    # Now we can check the right side of the string character by character.  If the string still has
    # characters and the last character is whitespace ...
    while (len(strng) > 0) and (strng[-1] in string.whitespace):
        # ... drop the last character from the string
	strng = strng[:-1]
    # return the results
    return strng
