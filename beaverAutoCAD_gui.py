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
import pygtk
if not sys.platform == 'win32':
    pygtk.require('2.0')

import gtk,gobject,time
import os.path
import datetime
import beaverAutoCAD_core

if gtk.pygtk_version < (2,4,0):
	print "Please install an updated version of pygtk.\nCurrently using pygtk version " +str(gtk.pygtk_version[0]) +"."+str(gtk.pygtk_version[1]) +"."+str(gtk.pygtk_version[2])
	raise SystemExit
	
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
today_date = "%04d%02d%02d_%02d-%02d" %(now.year, now.month, now.day, now.hour, now.minute)
today_date_designed = "%02d/%02d/%04d %02d:%02d" %(now.day, now.month, now.year, now.hour, now.minute)

tooltiptext = "Automatic AutoCAD Calculations was made by Shai Efrati for NADRASH Ltd. (2013)"

try:
   with open('settings.txt') as settings_file:
		homeDir = settings_file.readline()
		print "File will be saved at: "+homeDir
except IOError:
   homeDir = os.path.expanduser("~"+'\My Documents')
   print 'Warning: No settings file found. Using default directory instead, and creating a settings file.'
   print "File will be saved at: "+homeDir

class PyAPP():
	def __init__(self):
		self.window = gtk.Window()
		self.window.set_title("beaverAutoCAD - Automating AutoCAD Calculations")
		self.window.set_position(gtk.WIN_POS_CENTER)
		self.create_widgets()
		self.connect_signals()
		try:
			self.window.set_icon_from_file("SE_Logo25.png")
		except Exception, e:
			print e.message
			sys.exit(1)
		self.window.show_all()
		gtk.main()
	
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
	#This method checks the existance of an XLS file, and allows the user to overwrite it, 
	#or use a different file.
		tableFilename = self.dir_button.get_current_folder()+'\\'+filename+".xls"
		print tableFilename
		if os.path.isfile(tableFilename):
			md = gtk.MessageDialog(self, gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_QUESTION,
										gtk.BUTTONS_CLOSE, "File "+filename+".xls exist. Do you want to continue?")
			md.run()
			md.destroy()
	
	def create_widgets(self):
		self.vbox = gtk.VBox(spacing=10)

		self.hbox_0 = gtk.HBox(spacing=10)		
		self.nadrashlogo = gtk.Image()
		self.nadrashlogo.set_from_file("Nadrash25mm90.png")
		self.nadrashlogo.set_tooltip_text(tooltiptext)
		self.hbox_0.pack_start(self.nadrashlogo)
		
		self.hbox_1 = gtk.HBox(spacing=10)
		self.dir_button = gtk.FileChooserButton(title = "Choose directory")
		self.dir_button.set_action(gtk.FILE_CHOOSER_ACTION_SELECT_FOLDER)
		self.dir_button.set_tooltip_text("This is where your file will be saved")
		self.hbox_1.pack_start(self.dir_button)
		self.label = gtk.Label("File Name: ")
		self.hbox_1.pack_start(self.label)
		self.entry = gtk.Entry()
		self.hbox_1.pack_start(self.entry)
		
		self.hbox_2 = gtk.HBox(spacing=10)
		self.use0Label = gtk.Label("Use layer 0 in calculation: ")
		self.hbox_2.pack_start(self.use0Label)
		self.use0 = gtk.combo_box_new_text()
		self.use0.append_text('yes')
		self.use0.append_text('no')
		self.use0.set_active(1)
		self.use0.set_tooltip_text("In cases where you don't want to count blocks in layer 0, choose no. Otherwise, choose yes")
		self.hbox_2.pack_start(self.use0)
		
		self.unitsLabel = gtk.Label("DWG units: ")
		self.hbox_2.pack_start(self.unitsLabel)
		self.units = gtk.combo_box_new_text()
		self.units.append_text('m')
		self.units.append_text('cm')
		self.units.append_text('mm')
		self.units.set_active(1)
		self.units.set_tooltip_text("Choose the units you used in the drawing")
		self.hbox_2.pack_start(self.units)
		
		self.hbox_3 = gtk.HBox(spacing=10)
		self.bLineLength = gtk.Button("Sum Lines Lengths in a DWG to MS-Excel")
		self.hbox_3.pack_start(self.bLineLength)
		self.bBlocksCount = gtk.Button("Count Blocks in a DWG to MS-Excel")
		self.hbox_3.pack_start(self.bBlocksCount)
 		self.bBlocksCountPerLayer = gtk.Button("Count Blocks per layer in a DWG")
		self.hbox_3.pack_start(self.bBlocksCountPerLayer)
		
		self.hbox_4 = gtk.HBox(spacing=10)
		self.pbar = gtk.ProgressBar()
		self.hbox_4.pack_start(self.pbar)
		self.button_exit = gtk.Button("Exit")
		self.hbox_4.pack_start(self.button_exit)
		self.button_exit.set_tooltip_text("Press to exit...")
 
		self.hbox_5 = gtk.HBox(spacing=10)
		self.se_logo = gtk.Image()
		self.se_logo.set_from_file("SE_Logo25.png")
		self.hbox_5.pack_start(self.se_logo)
		self.se_logo.set_tooltip_text(tooltiptext)
		self.verLabel = gtk.Label("GUI Ver. " + __version__ + " // Core Ver. " + beaverAutoCAD_core.__version__)
		self.hbox_5.pack_start(self.verLabel)

		self.vbox.pack_start(self.hbox_0)
		self.vbox.pack_start(self.hbox_1)
		self.vbox.pack_start(self.hbox_2)
		self.vbox.pack_start(self.hbox_3)
		self.vbox.pack_start(self.hbox_4)
		self.vbox.pack_start(self.hbox_5)
		
		self.window.add(self.vbox)
	
	def connect_signals(self):
	# This method connects signals to relevant actions
		self.dir_button.set_current_folder(homeDir)
		self.filename = self.entry.set_text('temp'+today_date)
		self.button_exit.connect("clicked", self.callback_exit)
		self.bLineLength.connect('clicked', self.callback_lines_lengths)
		self.bBlocksCount.connect('clicked', self.callback_blocks_count)
		self.bBlocksCountPerLayer.connect('clicked', self.callback_blocks_count_per_layer)
		self.window.connect("destroy", gtk.main_quit)
	
	def callback_lines_lengths(self, widget, callback_data=None):
	# This method connects the gui to the relevant function in the app's core
		savingPath = self.dir_button.get_current_folder()
		filename = "AAC_lines_"+self.directory_settings()
		self.set_file_name(filename)
		draw_units = self.units.get_active_text()
		beaverAutoCAD_core.line_lengths_excel(filename, savingPath, draw_units) # calls the function from the core
		print "Done."

	def callback_blocks_count(self, widget, callback_data=None):
	# This method connects the gui to the relevant function in the app's core
		savingPath = self.dir_button.get_current_folder()
		filename = "AAC_blocks_"+self.directory_settings()		
		uselayer0 = self.use0.get_active_text()
		beaverAutoCAD_core.count_blocks_excel(filename, savingPath, uselayer0) # calls the function from the core
		print "Done."
 
 	def callback_blocks_count_per_layer(self, widget, callback_data=None):
	# This method connects the gui to the relevant function in the app's core
		savingPath = self.dir_button.get_current_folder()
		filename = "AAC_blocks_per_layer"+self.directory_settings()		
		uselayer0 = self.use0.get_active_text()
		beaverAutoCAD_core.count_blocks_per_layer(filename, savingPath, uselayer0) # calls the function from the core
		print "Done."
 
	def callback_exit(self, widget, callback_data=None):
		gtk.main_quit()
 
 
if __name__ == "__main__":
    app = PyAPP()