# Copyright (C) 2003 - 2009 The Board of Regents of the University of Wisconsin System 
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

"""This module contains Transana's configuration constants."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# Define Transana's Configuration Constants here.  This file can be easily substituted by an automated build process.

# Define a Boolean to indicate Single- or Multi- user
# NOTE:  When you change this value, you MUST change the MySQL for Python installation you are using
#        to match.
singleUserVersion = True
# Indicate if this is the Lab version
labVersion = False
# Set this flag to "True" to create the Demonstration version.  (But don't mix this with MU!)
demoVersion = False
# Indicate if this is the Workshop version.  (But don't mix this with MU or Lab!)
workshopVersion = True
# If we have the Workshop Version ...
if workshopVersion:
    # Import Python's datetime and time modules
    import datetime, time

    # Set a start date
    stdt = datetime.datetime(2010, 10, 12, 18, 00)
    # Set an expiration date
    xpdt = datetime.datetime(2010, 10, 15, 6, 00)
    # Determine the current date and time
    t2 = time.localtime()
    # Convert the current date and time to a datetime object so it can be compared
    t3 = datetime.datetime(t2[0], t2[1], t2[2], t2[3], t2[4])

    # If this version has expired ...
    if (stdt > t3) or (xpdt < t3):
        # ... alter the singleUserVersion value to sabbotage the database connection
        singleUserVersion = False
