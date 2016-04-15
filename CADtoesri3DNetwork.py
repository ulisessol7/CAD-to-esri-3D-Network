# -*- coding: utf-8 -*-
"""
CAD to esri 3D Network:
--------------------------------------------------------------------------------
This script...
--------------------------------------------------------------------------------
"""
from __future__ import print_function
import os
import glob
import re
import tempfile
from win32com import client

__author__ = 'Ulises  Guzman'
__date__ = '04/07/2016'

# this is only a temporary solution for the working directories.
PR_PATH = 'C:\Users\ulisesdario\CAD-to-esri-3D-Network'
# reformatting path strings to have forward slashes, otherwise AutoCAD fails.
PR_PATH = PR_PATH.replace('\\', '/') + '/'
print(PR_PATH)
# SCRATCH_PATH = 'C:\Users\ulisesdario\Desktop\scratch'
# SCRATCH_PATH = SCRATCH_PATH.replace('\\', '/') + '/'
# exploring the possibility of creating a temporary directory for geoprocessing
# this solution is multiplatform.
SCRATCH_PATH = tempfile.mkdtemp()
SCRATCH_PATH = SCRATCH_PATH.replace('\\', '/') + '/'
print(SCRATCH_PATH)
"""
--------------------------------------------------------------------------------
SCRIPT DESCRIPTION:





--------------------------------------------------------------------------------
"""


def path_retriever(ws):
    """ This helper function prompts the user for a folder or a gdb name,
    the string will then be use to construct a valid path string.

    Args:
    ws (string) = The name of the folder or the gdb that contains the data

    Returns:
    path (string) = A string representation of the folder,or geodatabase,
    location.

    Examples:
    >>> path_retriever('Guzman_lab3')
    Please enter a valid path for Guzman_lab3:
    """
    path = raw_input('Please enter a valid path for'
                     ' %s : ' % ws)
    # checking if the information provided by the user is a valid path
    while not (os.path.exists(path) and path.endswith('%s' % ws)):
        path = raw_input('Please enter a valid path for the %s: ' % ws)
    return path


def file_collector():
    """
    This function will grab the dwg files from their folders, for our project
    we will have two folders, one for each building, the naming convention
    for the folders should be as follows: bldgnumber_bdlgcode (for instance
    the library's folder should be named: 245_LIBR). I think the best option
    here would be to create a 'dictionary' that has the buildings numbers as
    keys and the dwg files as values, if you have questions about this please
    let me know.
    """
    # this function should return a python dictionary
    return


def cad_layer_name_simplifier(layername):
    """This function simplifies CAD layers' names by extracting their rightmost
    word characters, this logic follows the AIA CAD Layer Guidelines.

    Args:
    layername (str) = A string representation of the CAD layer name.

    Returns:
    simple_layer_name (str) = A 'simplified' version of the CAD layer name.

    Examples:
    >>> cad_layer_name_simplifier('A-SPAC-PPLN-AREA')
    AREA
    """
    match = re.search('\w+$', layername)
    simple_layer_name = match.group()
    return simple_layer_name


def autocadmap_to_shp(floor_plan, layer_on):
    """
    ...
    """
    # opening the last AutoCAD instance according to the windows registry.
    acad = client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    doc = acad.ActiveDocument
    doc.SendCommand("SDI 1\n")
    doc.SendCommand('(command "_.OPEN" "%s" "Y")\n' % floor_plan)
    doc.SendCommand("(ACAD-PUSH-DBMOD)\n")
    # turning off all the layers in the drawing.
    doc.SendCommand('(command "-LAYER" "OFF" "*" "Y" "")\n')
    # turning on the specified layer.
    doc.SendCommand('(command "-LAYER" "ON" "%s" "")\n' % layer_on)
    # thawing the specified layer.
    doc.SendCommand('(command "-LAYER" "THAW" "%s" "")\n' % layer_on)
    # switching to Model view.
    doc.SendCommand("MODEL\n")
    sl_name = cad_layer_name_simplifier(layer_on)
    mp = '-MAPEXPORT'
    # setting the parameters for the MAPEXPORT AutoCADMap command.
    out_name = '{0}{1}-{2}.shp'.format(SCRATCH_PATH,
                                       os.path.basename(floor_plan)[:-4],
                                       sl_name)
    ex_set = PR_PATH + 'mapexportsettings.epf'
    pr = 'PROCEED'
    ex_command = '(command "{0}" "SHP" "{1}" "Y" "{2}"' \
                 ' "{3}")'.format(mp, out_name, ex_set, pr)
    doc.SendCommand('%s\n' % ex_command)
    doc.SendCommand("(ACAD-POP-DBMOD)\n")
    # print(out_name)
    # print(ex_command)
    return

# tests
autocadmap_to_shp(
    'C:/Users/ulisesdario/Downloads/S-241E-01-DWG-BAS.dwg', 'A-SPAC-PPLN-AREA')

# cad_layer_name_simplifier('A-SPAC-PPLN-AREA')
