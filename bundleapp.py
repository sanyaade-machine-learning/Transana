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

# Some globals
wxPythonLib = '/usr/local/lib/wxPython-unicode-2.6.1.0/lib/'
resourceDir = 'Contents/Resources'

# Add all dylibs not added automatically
alldylibs = glob.glob('*.dylib')
for x in alldylibs[:]:
    alldylibs.remove(x)
    alldylibs.append( (x, os.path.join(resourceDir, x)) )

# Create basic list of files
allfiles = [('locale/en/LC_MESSAGES/Transana.mo', os.path.join(resourceDir, 'locale/en/LC_MESSAGES/Transana.mo')),
            ('locale/da/LC_MESSAGES/Transana.mo', os.path.join(resourceDir, 'locale/da/LC_MESSAGES/Transana.mo')),
            ('locale/de/LC_MESSAGES/Transana.mo', os.path.join(resourceDir, 'locale/de/LC_MESSAGES/Transana.mo')),
            ('locale/es/LC_MESSAGES/Transana.mo', os.path.join(resourceDir, 'locale/es/LC_MESSAGES/Transana.mo')),
            ('locale/fr/LC_MESSAGES/Transana.mo', os.path.join(resourceDir, 'locale/fr/LC_MESSAGES/Transana.mo')),
            ('locale/it/LC_MESSAGES/Transana.mo', os.path.join(resourceDir, 'locale/it/LC_MESSAGES/Transana.mo')),
            ('locale/nl/LC_MESSAGES/Transana.mo', os.path.join(resourceDir, 'locale/nl/LC_MESSAGES/Transana.mo')),
            ('locale/sv/LC_MESSAGES/Transana.mo', os.path.join(resourceDir, 'locale/sv/LC_MESSAGES/Transana.mo')) ]
allfiles = allfiles + \
          alldylibs + \
          [('share', os.path.join(resourceDir, 'share') )] + \
          [('images', os.path.join(resourceDir, 'images') )]

if MULTI_USER:
    appname = 'Transana-MU.app'
else:
    appname = 'Transana.app'
buildapp (
    name = appname,
    mainprogram = 'Transana.py',
    standalone = 1,
    libs = [
#            wxPythonLib + 'libwx_macd_core-2.5.1.dylib',
            wxPythonLib + 'libwx_macud-2.6.0.dylib',
#            wxPythonLib + 'libwx_macd_adv-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd_gizmos-2.5.1.dylib',
            wxPythonLib + 'libwx_macud_gizmos-2.6.0.dylib',
#            wxPythonLib + 'libwx_macd_gl-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd_html-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd_ogl-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd_stc-2.5.1.dylib',
            wxPythonLib + 'libwx_macud_stc-2.6.0.dylib',
#            wxPythonLib + 'libwx_macd_xrc-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd-2.5.1.rsrc',
#            wxPythonLib + 'libwx_macd-2.5.3.rsrc'
#            wxPythonLib + 'libwx_base_carbond-2.5.1.dylib',
#            wxPythonLib + 'libwx_base_carbond_net-2.5.1.dylib',
#            wxPythonLib + 'libwx_base_carbond_xml-2.5.1.dylib',
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
#            wxPythonLib + 'libwx_macd_core-2.5.1.dylib',
            wxPythonLib + 'libwx_macud-2.6.0.dylib',
#            wxPythonLib + 'libwx_macd_adv-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd_gizmos-2.5.1.dylib',
            wxPythonLib + 'libwx_macud_gizmos-2.6.0.dylib',
#            wxPythonLib + 'libwx_macd_gl-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd_html-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd_ogl-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd_stc-2.5.1.dylib',
            wxPythonLib + 'libwx_macud_stc-2.6.0.dylib',
#            wxPythonLib + 'libwx_macd_xrc-2.5.1.dylib',
#            wxPythonLib + 'libwx_macd-2.5.1.rsrc',
#            wxPythonLib + 'libwx_macd-2.5.3.rsrc'
#            wxPythonLib + 'libwx_base_carbond-2.5.1.dylib',
#            wxPythonLib + 'libwx_base_carbond_net-2.5.1.dylib',
#            wxPythonLib + 'libwx_base_carbond_xml-2.5.1.dylib',
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
