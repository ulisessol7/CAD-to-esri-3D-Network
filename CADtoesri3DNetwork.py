# -*- coding: utf-8 -*-
"""
Name:       CADtoesri3DNetwork.py
Purpose:    ...
Author:     Ulises  Guzman
Created:    04/07/2016
Copyright:   (c)
ArcGIS Version:   10.3.1
AutoCAD Version:
Python Version:   2.7x
--------------------------------------------------------------------------------
This script...
--------------------------------------------------------------------------------
"""
from __future__ import print_function
import os
import glob
import re
import tempfile
import inspect
from win32com import client
import arcpy
from arcpy import env


PR_PATH = os.getcwd()
# reformatting path strings to have forward slashes, otherwise AutoCAD fails.
PR_PATH = PR_PATH.replace('\\', '/') + '/'
print(PR_PATH)
# exploring the possibility of creating a temporary directory for geoprocessing
# this solution is multiplatform.
SCRATCH_PATH = tempfile.mkdtemp()
SCRATCH_PATH = SCRATCH_PATH.replace('\\', '/') + '/'
# env.scratchWorkspace = 'C:\Users\ulisesdario\Documents\ArcGIS'
print(SCRATCH_PATH)
temp_gdb = env.scratchGDB


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


def autocadmap_to_shp(floor_plan, outloc, layer_on):
    """This function transforms AutoCAD Map files into shapefiles, this function
    was developed as an alternative to the current workflows proposed by esri:
    http://desktop.arcgis.com/en/arcmap/10.3/tools/conversion-toolbox/
    cad-to-geodatabase.html. The main advantage is that this function can
    import spatial attributes and not only geometry, this is especially
    important on work environments that are heavily dependent on Munsys Ai:
    http://www.openspatial.com/products/munsys-ai
    Args:
    floor_plan (str) = The dwg file full path.
    layer_on (str) = The AutoCAD layer that is meant to be set 'ON'.

    Returns:
    A shapefile file based on the provided floor plan. This will only contain
    the geometry that is present in the layer represented by the layer_on
    parameter.

    Examples:
    >>> autocadmap_to_shp(
    'C:/Users/ulisesdario/S-241E-01-DWG-BAS.dwg', 'C:/Users/ulisesdario',
    'A-SPAC-PPLN-AREA')
    Executing autocadmap_to_shp...
    C:/Users/ulisesdario/S-241E-01-DWG-BAS-AREA.shp has been successfully
    created
    """
    # getting the name of the function programatically.
    print ('Executing {}... '.format(inspect.currentframe().f_code.co_name))
    # formatting outloc string to be compatible with AutoCAD.
    outloc = outloc.replace('\\', '/') + '/'
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
    out_name = '{0}{1}-{2}.shp'.format(outloc,
                                       os.path.basename(floor_plan)[:-4],
                                       sl_name)
    ex_set = PR_PATH + 'mapexportsettings.epf'
    pr = 'PROCEED'
    ex_command = '(command "{0}" "SHP" "{1}" "Y" "{2}"' \
                 ' "{3}")'.format(mp, out_name, ex_set, pr)
    doc.SendCommand('%s\n' % ex_command)
    doc.SendCommand("(ACAD-POP-DBMOD)\n")
    # doc.SendCommand("QQUIT\n")
    doc.Close(True)
    print('{} has been successfully created'.format(out_name))
    return


def shp_files_reader(location):
    """Returns a list of existing shp files in the provided location. This
    function was developed as an alternative to ListFeatureClasses(), the major
    advantage here is that this function is non ArcPy dependent.
    Args:
    location (str) = A string representation of the directory location.

    Returns:
    shapefiles (list) = A list that contains all the shapefiles' file names
    in the provided directory.
    shapefiles_full_path (list) = A list that contains all the shapefiles' real
    paths in the provided directory.

    Examples:
    >>> shp_files_reader('C:\Users\ulisesdario\Desktop\scratch')
    Executing shp_files_reader...
    2 shapefiles were found in scratch:
    ['thiessen.shp', 'simplify.shp']
    ['C:\\Users\\ulisesdario\\Desktop\\scratch\\thiessen.shp',
    'C:\\Users\\ulisesdario\\Desktop\\scratch\\simplify.shp']
    """
    # getting the name of the function programatically.
    print ('Executing {}... '.format(inspect.currentframe().f_code.co_name))
    original_workspace = os.getcwd()
    os.chdir(location)
    shapefiles = glob.glob("*.shp")
    shapefiles_full_path = [os.path.join(location, shp) for shp in shapefiles]
    print ('{} shapefiles were found in {}: '.format(
        (len(shapefiles)), os.path.basename(location)))
    print (shapefiles)
    os.chdir(original_workspace)
    return shapefiles, shapefiles_full_path


def shp_to_fc(shapefiles, gdblocation):
    """This function invokes the arcpy.FeatureClassToGeodatabase_conversion
    tool in each item in the shapefiles' list.The gdblocation argument serves
    as the output for the aforementioned tool.
    Args:
    shapefiles (list) = A list that contains all the shapefiles' real
    paths in the provided directory.
    gdblocation (str) = A string representation of the gdb location.

    Returns:
    A feature class file based on the provided shapefiles argument.

    Examples:
    >>>  s = ['C:\\Users\\ulisesdario\\Desktop\\scratch\\thiessen.shp',
    'C:\\Users\\ulisesdario\\Desktop\\scratch\\simplify.shp']
    >>> shp_to_fc(s, 'C:\Users\ulisesdario\Documents\ArcGIS\Default.gdb')
    Executing shp_to_fc...
    C:\Users\ulisesdario\Desktop\scratch\thiessen.shp Successfully
    converted: C:\Users\ulisesdario\Documents\ArcGIS\Default.gdb\thiessen
    ...
    2 feature classes were created in Default.gdb:
    """
    try:
        # getting the name of the function programatically.
        print ('Executing {}... '.format(
            inspect.currentframe().f_code.co_name))
        outLocation = gdblocation
        for shp in shapefiles:
            arcpy.FeatureClassToGeodatabase_conversion(shp, outLocation)
        print ('{} feature classes were created in {}: '.format(
            (len(shapefiles)), os.path.basename(gdblocation)))
        print (shapefiles)
    except arcpy.ExecuteError:
        print (arcpy.GetMessages(2))
    except Exception as e:
        print (e.args[0])


def skeletonizer():
    """Medial Axis
    """
    skeleton = None
    return skeleton


# tests
# autocadmap_to_shp(
#     'C:/Users/ulisesdario/Downloads/S-241E-01-DWG-BAS.dwg',
#     'C:\Users\ulisesdario\Desktop\scratch', 'A-SPAC-PPLN-AREA')
# cad_layer_name_simplifier('A-SPAC-PPLN-AREA')
# shp_files_reader('C:\Users\ulisesdario\Desktop\scratch')
# env.workspace = 'C:\Users\ulisesdario\Desktop\scratch'
# environments = arcpy.ListEnvironments()
# for environment in environments:
#     envSetting = eval("arcpy.env." + environment)
#     print ("%-30s: %s" % (environment, envSetting))
# arcpy.ResetEnvironments()
# environments = arcpy.ListEnvironments()
# for environment in environments:
#     envSetting = eval("arcpy.env." + environment)
#     print ("%-30s: %s" % (environment, envSetting))

# s = shp_files_reader('C:\Users\ulisesdario\Desktop\scratch')[1]
# shp_to_fc(s, 'C:\Users\ulisesdario\Documents\ArcGIS\Default.gdb')


mxd = arcpy.mapping.MapDocument(
    'C:\Users\ulisesdario\CAD-to-esri-3D-Network\scratch.mxd')
mxd.author = "Ulises Guzman"
mxd.save()
