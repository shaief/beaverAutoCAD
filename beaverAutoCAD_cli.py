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

now = datetime.datetime.now()
today_date = "%04d%02d%02d_%02d-%02d" % (now.year, now.month, now.day, now.hour, now.minute)
today_date_designed = "%02d/%02d/%04d %02d:%02d" % (now.day, now.month, now.year, now.hour, now.minute)

ascii_art = r"""
 _                       _____     _       _____ _____ ____
| |_ ___ ___ _ _ ___ ___|  _  |_ _| |_ ___|     |  _  |    \
| . | -_| .'| | | -_|  _|     | | |  _| . |   --|     |  |  |
|___|___|__,|\_/|___|_| |__|__|___|_| |___|_____|__|__|____/

-============  Automatic AutoCAD Calculations  ============-
-===== Was made by Shai Efrati for NADRASH Ltd. (2013) ====-
-=====     shaief@gmail.com // http://shaief.com      =====-
-========= https://github.com/shaief/beaverAutoCAD ========-

"""
done = """
 ____
|    \ ___ ___ ___
|  |  | . |   | -_|_
|____/|___|_|_|___|_|

"""
try:
    with open('settings.txt') as settings_file:
        homeDir = settings_file.readline()
        print "File will be saved at: " + homeDir
except IOError:
    homeDir = os.path.expanduser("~" + '\My Documents')
    print 'Warning: No settings file found. Using default directory instead, and creating a settings file.'
    print "File will be saved at: " + homeDir


class PyAPP(object):
    def __init__(self):
        self.dir_name = os.path.expanduser('~\Desktop')
        self.filename = '{}_{}'.format(beaverAutoCAD_core.acad.ActiveDocument.Name[0:-4],
                                       today_date)

    def set_file_name(self, filename):
        # This method checks the existance of an XLS file, and allows the user to overwrite it,
        # or use a different file.
        tableFilename = self.dir_button.get_current_folder() + '\\' + filename + ".xls"
        print tableFilename
        if os.path.isfile(tableFilename):
            md = gtk.MessageDialog(self, gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_QUESTION,
                                   gtk.BUTTONS_CLOSE, "File " + filename + ".xls exist. Do you want to continue?")
            md.run()
            md.destroy()

    def lines_lengths(self, ):
        # This method connects the gui to the relevant function in the app's core
        savingPath = self.get_current_folder()
        filename = "AAC_lines_" + self.directory_settings()
        self.set_file_name(filename)
        draw_units = self.units.get_active_text()
        beaverAutoCAD_core.line_lengths_excel(filename, savingPath, draw_units)  # calls the function from the core
        print "Done."

    def blocks_count(self, ):
        # This method connects the gui to the relevant function in the app's core
        savingPath = self.dir_button.get_current_folder()
        filename = "AAC_blocks_" + self.directory_settings()
        beaverAutoCAD_core.count_blocks_excel(filename, savingPath)  # calls the function from the core
        print "Done."

    def user_interactions(self):
        print(ascii_art)
        print 'Hello and welcome to beaverAutoCAD textual interface!'
        print 'Files will be saved at: {}'.format(self.dir_name)
        print '(1) Sum Lines Lengths in a DWG to MS-Excel'
        print '(2) Count Blocks in a DWG to MS-Excel'
        print '(3) Count Blocks per layer in a DWG'
        user_chose = raw_input('Please choose what to do [1,2,3]: ')
        if user_chose == '1':
            user_string = raw_input('Enter a string to search in layer names: ')
            user_units = raw_input('Drawing units (m / [cm] / mm):')
            if not user_units:
                user_units = 'cm'
            beaverAutoCAD_core.line_lengths_excel(filename='AAC_lines_{}'.format(self.filename),
                                                  savingPath=self.dir_name,
                                                  draw_units=user_units,
                                                  layers_contain=user_string)
            print 'Done.'

        elif user_chose == '2':
            user_string = raw_input('Enter a string to search in layer names: ')
            user_layer0 = raw_input('Use layer 0? y/[n]')
            if not user_layer0 or user_layer0.lower() == 'n':
                user_layer0 = 'no'
            else:
                user_layer0 = 'yes'
            beaverAutoCAD_core.count_blocks_excel(filename="AAC_blocks_{}".format(self.filename),
                                                  savingPath=self.dir_name,
                                                  uselayer0=user_layer0,
                                                  layers_contain=user_string)
            print 'Done.'
        elif user_chose == '3':
            user_string = raw_input('Enter a string to search in layer names: ')
            user_layer0 = raw_input('Use layer 0? y/[n]')
            if not user_layer0 or user_layer0.lower() == 'n':
                user_layer0 = 'no'
            else:
                user_layer0 = 'yes'
            beaverAutoCAD_core.count_blocks_per_layer(filename="AAC_blocks_per_layer_{}".format(self.filename),
                                                      savingPath=self.dir_name,
                                                      uselayer0=user_layer0,
                                                      layers_contain=user_string)
            print 'Done.'
        else:
            print 'No option was chosen. Goodbye!'
        print """
                               __
 _____           _ _          |  |
|   __|___ ___ _| | |_ _ _ ___|  |
|  |  | . | . | . | . | | | -_|__|
|_____|___|___|___|___|_  |___|__|
                      |___|
"""


if __name__ == "__main__":
    app = PyAPP()
    app.user_interactions()