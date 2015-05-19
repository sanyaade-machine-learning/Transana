# Copyright (c) 2003 - 2006 The Board of Regents of the University of Wisconsin System 
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
# Bundlebuilder script for Transana
# Creates a bundle for distribution on OS/X 10.3
# Author: Rajas A. Sambhare <rasambhare@wisc.edu>
# Modified extensively:  David K. Woods <dwoods@wcer.wisc.edu>

from bundlebuilder import buildapp
import os, glob

import sys
if (len(sys.argv) > 2) and (sys.argv[2] == 'MU'):
    MULTI_USER = True
    del(sys.argv[2])
else:
    MULTI_USER = False
if (len(sys.argv) > 2) and (sys.argv[2] == 'PPC'):
    PPC = True
    del(sys.argv[2])
else:
    PPC = False

# Some globals
wxPythonLib = '/usr/local/lib/wxPython-unicode-2.8.1.1/lib/'
resourceDir = 'Contents/Resources'

# Select the audio extraction and SRB connection files we need for the version we're building
if PPC:
    allfiles = [('./audioextract PPC', os.path.join(resourceDir, 'audioextract'))]
    alldylibs = [('./srbClient-3.3.1-PPC.dylib', os.path.join(resourceDir, 'srbClient.dylib'))]
else:
    allfiles = [('./audioextract Universal', os.path.join(resourceDir, 'audioextract'))]
    alldylibs = [('./srbClient-3.3.1-Universal.dylib', os.path.join(resourceDir, 'srbClient.dylib'))]
    
# Create basic list of files
allfiles = allfiles + \
           [('locale/en/LC_MESSAGES', os.path.join(resourceDir, 'locale/en/LC_MESSAGES')),
            ('locale/da/LC_MESSAGES', os.path.join(resourceDir, 'locale/da/LC_MESSAGES')),
            ('locale/de/LC_MESSAGES', os.path.join(resourceDir, 'locale/de/LC_MESSAGES')),
            ('locale/es/LC_MESSAGES', os.path.join(resourceDir, 'locale/es/LC_MESSAGES')),
            ('locale/fr/LC_MESSAGES', os.path.join(resourceDir, 'locale/fr/LC_MESSAGES')),
            ('locale/it/LC_MESSAGES', os.path.join(resourceDir, 'locale/it/LC_MESSAGES')),
            ('locale/nb/LC_MESSAGES', os.path.join(resourceDir, 'locale/nb/LC_MESSAGES')),
            ('locale/nn/LC_MESSAGES', os.path.join(resourceDir, 'locale/nn/LC_MESSAGES')),
            ('locale/nl/LC_MESSAGES', os.path.join(resourceDir, 'locale/nl/LC_MESSAGES')),
            ('locale/ru/LC_MESSAGES', os.path.join(resourceDir, 'locale/ru/LC_MESSAGES')),
            ('locale/sv/LC_MESSAGES', os.path.join(resourceDir, 'locale/sv/LC_MESSAGES')) ]
allfiles = allfiles + \
          alldylibs + \
          [('share', os.path.join(resourceDir, 'share') )] + \
          [('images', os.path.join(resourceDir, 'images') )]

if MULTI_USER:
    appname = 'Transana-MU'
else:
    appname = 'Transana'
if PPC:
   appname += '-PPC.app'
else:
    appname += '-Intel.app' 
buildapp (
    name = appname,
    mainprogram = 'Transana.py',
    standalone = 1,
    libs = [
            wxPythonLib + 'libwx_macud-2.8.0.dylib',
#            wxPythonLib + 'libwx_macud-2.8.0.0.0.dylib',
            wxPythonLib + 'libwx_macud_gizmos-2.8.0.dylib',
#            wxPythonLib + 'libwx_macud_gizmos-2.8.0.0.0.dylib',
            wxPythonLib + 'libwx_macud_stc-2.8.0.dylib',
#            wxPythonLib + 'libwx_macud_stc-2.8.0.0.0.dylib',
           ],
    includePackages = ['encodings'],
    iconfile = 'images/Transana.icns',
    files = allfiles,
)

buildapp (
    name = 'TransanaHelp.app',
    mainprogram = 'Help.py',
    standalone = 1,
    libs = [
            wxPythonLib + 'libwx_macud-2.8.0.dylib',
#            wxPythonLib + 'libwx_macud-2.8.0.0.0.dylib',
            wxPythonLib + 'libwx_macud_gizmos-2.8.0.dylib',
#            wxPythonLib + 'libwx_macud_gizmos-2.8.0.0.0.dylib',
            wxPythonLib + 'libwx_macud_stc-2.8.0.dylib',
#            wxPythonLib + 'libwx_macud_stc-2.8.0.0.0.dylib',
           ],
    includePackages = ['encodings'],
    iconfile = 'images/TransanaHelp.icns',
    files=[('help', os.path.join(resourceDir, 'help'))]
)

# Let's now set permissions for the application bundles
print "changing permissions for app bundles"
os.chdir('build')
os.system('chmod -R a+rw *.app')
os.chdir('..')
print "Build complete"
