#!/usr/bin/env python
'''
beaverAutoCAD is a software for calculating data recieved from AutoCAD DWG,
and creates a MS-Excel file of this data.

Copyright 2013 Shai Efrati

This file is part of beaverAutoCAD.

beaverAutoCAD is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

beaverAutoCAD is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with beaverAutoCAD.  If not, see <http://www.gnu.org/licenses/>.
'''

try:
    import gtk

    if gtk.pygtk_version < (2, 4, 0):
        print "Please install an updated version of pygtk.\nCurrently using pygtk version " + str(
            gtk.pygtk_version[0]) + "." + str(gtk.pygtk_version[1]) + "." + str(gtk.pygtk_version[2])
        print "Using textual interface"
        import beaverAutoCAD_cli

        beaverAutoCAD_cli.PyAPP()
    import beaverAutoCAD_gui

    beaverAutoCAD_gui.PyAPP()

except ImportError:
    print "Using textual interface"
    import beaverAutoCAD_cli

    beaverAutoCAD_cli.PyAPP()
