# Copyright (C) 2003 - 2015 The Board of Regents of the University of Wisconsin System 
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

# import Python's sys module
import sys

# Define Transana's Configuration Constants here.  This file can be easily substituted by an automated build process.

# Define a Boolean to indicate Single- or Multi- user
# NOTE:  When you change this value, you MUST change the MySQL for Python installation you are using
#        to match.
singleUserVersion = False
# Different Python versions require different database engines!
if sys.version[:5] == '2.6.6':
    DBInstalled = 'MySQLdb-server'
elif sys.version[:6].strip() in ['2.7.7 ', '2.7.8', '2.7.9', '2.7.10']:
#    DBInstalled = 'PyMySQL'
    DBInstalled = 'MySQLdb-server'

##    print "TransanaConfigConstants.py"
##    print "Python", sys.version[:5]
##    print
##    print DBInstalled, "in use.  Need to fix SSL configuration for PyMySQL."
##    print

# Define Std vs. Pro feature set
proVersion = True
# Indicate if this is the Lab version
labVersion = False
# Set this flag to "True" to create the Demonstration version.  (But don't mix this with MU!)
demoVersion = False
# Indicate if this is the Workshop version.  (But don't mix this with MU or Lab!)
workshopVersion = False
