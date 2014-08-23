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

__version__ = "0.0.2"

import sys
import time
import os.path
import datetime
import unicodedata
from pyautocad import Autocad, utils
from pyautocad.contrib.tables import Table

acad = Autocad()  # AutoCAD should be running with the analyzed drawing

currentDirectory = os.getcwd()
print 'Running from: {}'.format(currentDirectory)
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

def prompt_ascii_art():
    acad.prompt(ascii_art)

def line_lengths_excel(filename, savingPath, draw_units, layers_contain):
    '''
    This function iterates over all the layers in the opened DWG and write sum of line lengths of each layer
    into one MS-Excel sheet.
    Parameters needed:
    1. Name of an MS-Excel file (doesn't have to exist)
    2. Units of the drwaing
    '''
    os.chdir(savingPath)
    acad.prompt("Creating a table of line lengths")
    tableFilename = filename + '.xls'
    table = Table()
    layers = []
    total_length = []
    units_scale = {"m": 1, "cm": 100, "mm": 1000}
    units = draw_units
    scale = units_scale[units]

    for line in acad.iter_objects('polyline'):
        l1 = line.length
        if layers_contain in line.Layer:
            # print line.Layer
            if line.Layer in layers:
                i = layers.index(line.Layer)
                total_length[i] += l1
            else:
                layers.append(line.Layer)
                total_length.append(l1)
    print layers
    print total_length
    acad.prompt("Saving file AAC_lines_" + filename + ".xls at " + savingPath)
    # Add headlines to table
    table.writerow(["NADRASH LTD.", "Lines Lengths", "Created:", today_date_designed, acad.ActiveDocument.Name])
    table.writerow(["Layer", "Length [" + units + "]", "Length [m]", "", ""])
    # Add data to table
    for i in range(len(layers)):
        table.writerow([layers[i], total_length[i], total_length[i] / scale, "", ""])
    # Save table in xls
    table.save(tableFilename, 'xls')


def count_blocks_excel(filename, savingPath, uselayer0, layers_contain):
    '''
    This function iterates over all the layers in the opened DWG and summing up all the blocks in the file
    into one MS-Excel sheet.
    Parameters needed:
    1. Name of an MS-Excel file (doesn't have to exist)
    2. Should it count objects in Layer 0?
    3. Should it count objects only in specific layers?
    '''
    os.chdir(savingPath)
    tableFilename = filename + '.xls'
    table = Table()
    block_list = []
    total_blocks = []
    acad.prompt("Creating a table of blocks count")
    layer0counter = 0
    for block in acad.iter_objects('block'):
        ''' This if statement checks if the layer is Layer0.
        Some people workflow includes leaving "garbage" in layer 0,
        and we don't want it to count these objects.'''
        if (uselayer0 == "no") & (unicodedata.normalize('NFKD', block.Layer).encode('ascii', 'ignore') == "0"):
            # print "block was on layer 0"
            layer0counter += 1
            continue
        if layers_contain in block.Layer:
            # print block.Layer
            if block.name in block_list:
                i = block_list.index(block.name)
                total_blocks[i] += 1
            else:
                block_list.append(block.name)
                total_blocks.append(1)

    print block_list
    print total_blocks
    if (uselayer0 == "no"):
        print str(layer0counter) + " blocks counted and ignored on layer 0"
    acad.prompt("Saving file AAC_blocks_" + filename + ".xls at " + savingPath)
    # Add headlines to table
    table.writerow(["NADRASH LTD.", "Blocks Count", "Created:", today_date_designed, acad.ActiveDocument.Name])
    table.writerow(["Block", "Amount", "", "", "Blocks counted only in layers contain: {}".format(layers_contain)])
    # Add data to table
    for i in range(len(block_list)):
        table.writerow([block_list[i], total_blocks[i], "", "", ""])
    # Save table in xls
    table.save(tableFilename, 'xls')


def count_blocks_per_layer(filename, savingPath, uselayer0, layers_contain):
    '''
    This function iterates over all the layers in the opened DWG and summing up all the blocks in each layer
    into one MS-Excel sheet.
    Parameters needed:
    1. Name of an MS-Excel file (doesn't have to exist)
    2. Should it count objects in Layer 0?
    3. Should it count objects only in specific layers?
    '''
    os.chdir(savingPath)
    tableFilename = filename + '.xls'
    table = Table()
    block_list = []
    block_name_list = []
    block_layer = []
    total_blocks = []
    acad.prompt("Creating a table of blocks count")
    layer0counter = 0
    for block in acad.iter_objects('block'):
        ''' This if statement checks if the layer is Layer0.
        Some people workflow includes leaving "garbage" in layer 0,
        and we don't want it to count these objects.'''
        if (uselayer0 == "no") & (unicodedata.normalize('NFKD', block.Layer).encode('ascii', 'ignore') == "0"):
            # print "block was on layer 0"
            layer0counter += 1
            continue
        if layers_contain in block.Layer:
            # print block.Layer
            if block.Layer + " " + block.name in block_list:
                i = block_list.index(block.Layer + " " + block.name)
                total_blocks[i] += 1
            else:
                block_list.append(block.Layer + " " + block.name)
                block_name_list.append(block.name)
                block_layer.append(block.Layer)
                total_blocks.append(1)

    print block_list
    print total_blocks
    if (uselayer0 == "no"):
        print str(layer0counter) + " blocks counted and ignored on layer 0"
    acad.prompt("Saving file AAC_blocks_per_layer" + filename + ".xls at " + savingPath)
    # Add headlines to table
    table.writerow(["NADRASH LTD.", "Blocks Count", "Created:", today_date_designed, acad.ActiveDocument.Name])
    table.writerow(["Layer", "Block Name", "Amount", "",
                    "Blocks counted only in layers contain: {}".format(layers_contain)])
    # Add data to table
    for i in range(len(block_list)):
        table.writerow([block_layer[i], block_name_list[i], total_blocks[i], "", ""])
    # Save table in xls
    table.save(tableFilename, 'xls')
