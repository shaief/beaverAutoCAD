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

import sys

import os.path
import datetime
import beaverAutoCAD_core

__author__ = "Shai Efrati"
__copyright__ = "Copyright 2013, Shai Efrati for NADRASH Ltd."
__credits__ = ["Shai Efrati"]
__license__ = "GPL"
__version__ = "0.0.1"
__maintainer__ = "Shai Efrati"
__email__ = "shaief@gmail.com"
__status__ = "Production"

currentDirectory = os.getcwd()
print currentDirectory
now = datetime.datetime.now()
today_date = "%04d%02d%02d_%02d-%02d" % (now.year, now.month, now.day, now.hour, now.minute)
today_date_designed = "%02d/%02d/%04d %02d:%02d" % (now.day, now.month, now.year, now.hour, now.minute)

try:
    with open('settings.txt') as settings_file:
        homeDir = settings_file.readline()
        print "File will be saved at: " + homeDir
except IOError:
    homeDir = os.path.expanduser("~" + '\My Documents')
    print 'Warning: No settings file found. Using default directory instead, and creating a settings file.'
    print "File will be saved at: " + homeDir


class PyAPP():
    # def __init__(self):

    def directory_settings(self):
        dir_name = self.dir_button.get_current_folder()
        print dir_name
        os.chdir(currentDirectory)
        with open("settings.txt", "w") as text_file:
            text_file.write(dir_name)
        os.chdir(dir_name)
        filename = self.entry.get_text()
        return filename

    def set_file_name(self, filename):
        # This method checks the existance of an XLS file, and allows the user to overwrite it,
        #or use a different file.
        tableFilename = self.dir_button.get_current_folder() + '\\' + filename + ".xls"
        print tableFilename
        if os.path.isfile(tableFilename):
            md = gtk.MessageDialog(self, gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_QUESTION,
                                   gtk.BUTTONS_CLOSE, "File " + filename + ".xls exist. Do you want to continue?")
            md.run()
            md.destroy()

    def callback_lines_lengths(self, widget, callback_data=None):
        # This method connects the gui to the relevant function in the app's core
        savingPath = self.dir_button.get_current_folder()
        filename = "AAC_lines_" + self.directory_settings()
        self.set_file_name(filename)
        draw_units = self.units.get_active_text()
        beaverAutoCAD_core.line_lengths_excel(filename, savingPath, draw_units)  # calls the function from the core
        print "Done."

    def callback_blocks_count(self, widget, callback_data=None):
        # This method connects the gui to the relevant function in the app's core
        savingPath = self.dir_button.get_current_folder()
        filename = "AAC_blocks_" + self.directory_settings()
        beaverAutoCAD_core.count_blocks_excel(filename, savingPath)  # calls the function from the core
        print "Done."

    def callback_exit(self, widget, callback_data=None):
        gtk.main_quit()


if __name__ == "__main__":
    app = PyAPP()