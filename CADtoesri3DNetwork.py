# -*- coding: utf-8 -*-
"""
Name:       CADtoesri3DNetwork.py
Author:     Ulises  Guzman
Created:    04/07/2016
Copyright:   (c)
ArcGIS Version:   10.3.1
AutoCAD Version:  20.1
Python Version:   2.7.8
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
import repr as reprlib
from win32com import client
import arcpy
import arcpy.na
from arcpy import env
import pandas as pd
import pdb


PR_PATH = os.getcwd()
# reformatting path strings to have forward slashes, otherwise AutoCAD fails.
PR_PATH = PR_PATH.replace('\\', '/') + '/'
# print(PR_PATH)
# exploring the possibility of creating a temporary directory for geoprocessing
# this solution is multiplatform.
SCRATCH_PATH = tempfile.mkdtemp()
SCRATCH_PATH = SCRATCH_PATH.replace('\\', '/') + '/'
# env.scratchWorkspace = 'C:\Users\ulisesdario\Documents\ArcGIS'
# print(SCRATCH_PATH)
temp_gdb = env.scratchGDB


def path_retriever(directory_name=os.path.basename(os.getcwd())):
    """ This helper function prompts the user for a folder or a gdb name,
    the string will then be use to construct a valid path string.

    Args:
    directory_name(string) = The name of the folder or the gdb that contains
    the data.

    Returns:
    path (string) = A string representation of the folder,or geodatabase,
    location.

    Examples:
    >>> path_retriever('Guzman_lab3')
    Please enter a valid path for Guzman_lab3:
    """
    # getting the name of the function programatically.
    print ('Executing {}... '.format(inspect.currentframe().f_code.co_name))
    path = raw_input('Please enter a valid path for'
                     ' %s : ' % directory_name)
    # checking if the information provided by the user is a valid path
    while not (os.path.exists(path) and path.endswith('%s' % directory_name)):
        path = raw_input(
            'Please enter a valid path for the %s: ' % directory_name)
    return path


def bldgs_dict(qryAllBldgs_location='qryAllBldgs.xlsx'):
    """
    This function reads the 'qryAllBldgs.xlsx' file and returns
    a python dictionary in which the keys are the buildings' numbers
    and the values are the buildings' codes (i.e. {325 : 'MUEN'} ).
    This allows to rename files either based on building number or
    building code. The 'qryAllBldgs.xlsx' file was exported from
    the 'LiveCASP_SpaceDatabase.mdb' access file
    ('\\Cotterpin\CASPData\Extract\LiveDatabase').
    Args:
    qryAllBldgs_location(string) = The location of the 'qryAllBldgs.xlsx' file.

    Returns:
    bldgs (dict) = A python dictionary in which the keys are the buildings'
    numbers and the values are the buildings' codes.

    Examples:
    >>> bldgs_dict()
    Executing bldgs_dict...
    CU Boulder buildings: {u'131': u'ATHN', u'133': u'TB33',...}
    """
    # getting the name of the function programatically.
    print ('Executing {}... '.format(inspect.currentframe().f_code.co_name))
    bldgs = pd.read_excel(
        qryAllBldgs_location, header=0, index_col=0).to_dict()
    # getting building codes.
    bldgs = bldgs['BuildingCode']
    print('CU Boulder buildings: {}'.format(reprlib.repr(bldgs)))
    return bldgs
# bldgs_dict()


def dwg_file_collector(bldgs_dict, location=os.getcwd()):
    """
    ...
    Args:
    bldgs_dict (func) = A call to the bldgs_dict function.
    location (str) = A string representation of the directory location.

    Returns:
    dwg_bldg_code (dict) = A dictionary that contains  a list of every dwg
    per building folder.
    dwg_bldg_number (dict) = A dictionary that contains  a list of every dwg
    per building folder.
    in the subfolders.

    Examples:
    >>> dwg_file_collector(bldgs_dict('...\qryAllBldgs.xlsx'),
                           '...\CAD-to-esri-3D-Network\\floorplans')
    Executing bldgs_dict...
    CU Boulder buildings: {u'131': u'ATHN', u'133': u'TB33', ...}
    Executing dwg_file_collector...
    16 dwgs were found in: .../CAD-to-esri-3D-Network/floorplans/
    Buildings numbers dictionary: {'338': ['S-338-01-DWG-BAS.dwg',...}
    Buildings codes dictionary: {u'ADEN': ['S-339-01-DWG-BAS.dwg',...}
    """
    # getting the name of the function programatically.
    print ('Executing {}... '.format(inspect.currentframe().f_code.co_name))
    original_workspace = os.getcwd()
    # making the path compatible with python.
    location = location.replace('\\', '/') + '/'
    os.chdir(location)
    folders = [p.replace('\\', '') for p in glob.glob('*/')]
    # so the number of dwgs that were found can be reported.
    dwg_files = []
    dwg_bldg_number = {}
    dwg_bldg_code = {}
    for folder in folders:
        folder_path = ''.join([location, folder])
        os.chdir(folder_path)
        folder_dwg_files = glob.glob('*.dwg')
        # our current dwg naming convention is as follows:
        # 'bldg_number-floor_number-DWG-drawing_type (i.e.'325-01-DWG-BAS.dwg')
        # removes 'ROOF' files from the floorplans' list.
        folder_dwg_files = [dwg for dwg in folder_dwg_files if dwg[-7:] ==
                            'BAS.dwg' and 'ROOF' not in dwg]
        # dict where the buildings' numbers are the keys.
        dwg_bldg_number[folder] = folder_dwg_files
        # dict where the buildings' codes are the keys.
        dwg_bldg_code[bldgs_dict[folder]] = folder_dwg_files
        dwg_files += folder_dwg_files
    os.chdir(original_workspace)
    print ('{} dwgs were found in: {} '.format(
        (len(dwg_files)), location))
    print('Buildings numbers dictionary: {}'.format(
        reprlib.repr(dwg_bldg_number)))
    print('Buildings codes dictionary: {}'.format(
        reprlib.repr(dwg_bldg_code)))
    return dwg_bldg_number, dwg_bldg_code
dwg_file_collector(
    bldgs_dict('C:/Users/ulisesdario/CAD-to-esri-3D-Network/qryAllBldgs.xlsx'),
    'C:\Users\ulisesdario\CAD-to-esri-3D-Network\\floorplans')


def cad_layer_name_simplifier(layer_name):
    """This function simplifies CAD layers' names by extracting their rightmost
    word characters, this logic follows the AIA CAD Layer Guidelines.

    Args:
    layer_name (str) = A string representation of the CAD layer name.

    Returns:
    simple_layer_name (str) = A 'simplified' version of the CAD layer name.

    Examples:
    >>> cad_layer_name_simplifier('A-SPAC-PPLN-AREA')
    AREA
    """
    # getting the name of the function programatically.
    print ('Executing {}... '.format(inspect.currentframe().f_code.co_name))
    match = re.search('\w+$', layer_name)
    simple_layer_name = match.group()
    return simple_layer_name


def autocadmap_to_shp(floor_plan, out_loc, layer_on):
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
    out_loc (str) = A string representation of the output location.

    Returns:
    A shapefile based on the provided floor plan. This will only contain
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
    # formatting out_loc string to be compatible with AutoCAD.
    out_loc = out_loc.replace('\\', '/') + '/'
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
    out_name = '{0}{1}-{2}.shp'.format(out_loc,
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


def shp_files_reader(location=os.getcwd()):
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


def shp_to_fc(shapefiles, out_gdb_location='in_memory'):
    """This function invokes the arcpy.FeatureClassToGeodatabase_conversion
    tool in each item in the shapefiles' list.The out_gdb_location argument
    serves as the output for the aforementioned tool.
    Args:
    shapefiles (list) = A list that contains all the shapefiles' real
    paths in the provided directory.
    out_gdb_location (str) = A string representation of the gdb location.

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
        for shp in shapefiles:
            arcpy.FeatureClassToGeodatabase_conversion(shp, out_gdb_location)
        print ('{} feature classes were created in {}: '.format(
            (len(shapefiles)), os.path.basename(out_gdb_location)))
        print (shapefiles)
    except arcpy.ExecuteError:
        print (arcpy.GetMessages(2))


def skeletonizer(floor_plans_fc, out_location, workspace=env.workspace):
    """Medial Axis
    """
    try:
        # getting the name of the function programatically.
        print ('Executing {}... '.format(
            inspect.currentframe().f_code.co_name))
        # string manipulation
        skeleton_name = [floor_plans_fc[:] for floor_plan in floor_plans_fc]
        print(skeleton_name)
        # logic to create skeletons as feature classes

    except arcpy.ExecuteError:
        print (arcpy.GetMessages(2))


def build_network(egdb, feature_dataset, feature_type):
    """It keeps the master network up to date by re-building its source
    features.The master network and its source features will be assumed to live
    in an enterprise geodatabase created on PostgreSQL 9.3.
    """
    # 'Database Connections' points to the default location
    network_ws = 'Database Connections\\{}\\{}'.format(egdb, feature_dataset)
    env.workspace = network_ws
    print(network_ws)
    fc_list = arcpy.ListFeatureClasses('*', '%s' % feature_type)
    print(fc_list)

    # esri code, for debugging purposes
    # for environment in environments:
    # As the environment is passed as a variable, use Python's getattr
    # to evaluate the environment's value
    #     #
    #     env_value = getattr(arcpy.env, environment)

    # Format and print each environment and its current setting
    #     #
    #     print("{0:<30}: {1}".format(environment, env_value))
    return


# tests
# build_network(
#     'master_network.sde', 'network_scratch.ulisessol7.CU_Boulder_Networks',
#     'POLYLINE')
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


# mxd = arcpy.mapping.MapDocument(
#     'C:\Users\ulisesdario\CAD-to-esri-3D-Network\scratch.mxd')
# mxd.author = "Ulises Guzman"
# mxd.save()
# print(os.path.basename(os.getcwd()))
# dwg_file_collector('W:\\')
