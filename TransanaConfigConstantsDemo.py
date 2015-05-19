# Copyright (C) 2003 - 2010 The Board of Regents of the University of Wisconsin System 
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
demoVersion = True
# Indicate if this is the Workshop version.  (But don't mix this with MU or Lab!)
workshopVersion = False
