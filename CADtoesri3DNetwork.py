# -*- coding: utf-8 -*-
"""CAD to esri 3D Network"""
from __future__ import print_function
import os
import glob
from win32com import client

__author__ = 'Ulises  Guzman'
__author__ = 'Joseph Javadi'
date = '04/07/2016'

# this is only a temporary solution for the working directories.
PR_PATH = 'C:\Users\ulisesdario\CAD-to-esri-3D-Network'
# reformatting path strings to have forward slashes, otherwise AutoCAD fails.
PR_PATH = PR_PATH.replace('\\', '/') + '/'
print(PR_PATH)
SCRATCH_PATH = 'C:\Users\ulisesdario\Desktop\scratch'
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
    # Joseph
    # this function should return a python dictionary
    return


def cad_to_fc(floor_plan, layer_on):
    """
    """
    # Ulises
    # opening the last AutoCAD instance according to the windows registry.
    acad = client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    doc = acad.ActiveDocument
    doc.SendCommand("SDI 1\n")
    doc.SendCommand('(command "_.OPEN" "%s")\n' % floor_plan)
    # turning off all the layers in the drawing.
    doc.SendCommand('(command "-layer" "off" "*" "Y" "")\n')
    # turning on the specified layer.
    doc.SendCommand('(command "-layer" "on" "%s" "")\n' % layer_on)
    # thawing the specified layer.
    doc.SendCommand('(command "-layer" "thaw" "%s" "")\n' % layer_on)
    # switching to Model view.
    doc.SendCommand("Model\n")
    mp = '-mapexport'
    # setting the parameters for the MAPEXPORT AutoCADMap command.
    out_name = SCRATCH_PATH + os.path.basename(floor_plan)[:-4] + '.shp'
    ex_set = PR_PATH + 'mapexportsettings.epf'
    pr = 'Proceed'
    # MAPEXPORT AutoCAD Map string command.
    ex_command = '(command "{0}" "SHP" "{1}" "Y" "{2}" "S" "L" "All" "*" "*"' \
                 ' "No" "{3}")'.format(mp, out_name, ex_set, pr)
    print(out_name)
    print(ex_command)
    doc.SendCommand("'%s'\n" % ex_command)
    return


def topology_checks():
    """
    """
    # Ulises
    return

cad_to_fc(
    'C:/Users/ulisesdario/Downloads/S-241E-01-DWG-BAS.dwg', 'A-SPAC-PPLN-AREA')
