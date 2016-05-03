# -*- coding: utf-8 -*-
"""
Name:       CADtoesri3DNetwork.py
Author:     Ulises  Guzman
Created:    05/04/2016
Copyright:   (c)
ArcGIS Version:   10.3.1
Python Version:   2.7.8
--------------------------------------------------------------------------------
This script outlines a possible workflow for the creation of a pedestrian
indoor & outdoor 3D Network for The University of Colorado at Boulder.
The logic is as follows:
* Gather buildings' information (building codes & numbers) from excel file
(that's how this information is currently storage).
* Read & filter dwg files (buildings floorplans).
* Create esri feature classes by automating AutoCAD Map.
* Create skeletons for each floorplan (A work in progress).
* Populate skeletons with the relevant attributes and transform them into 3D
feature classes.
*Finally update the network by invoking the 'build method'.
*Run centrality analysis to test the network topology.

The latter requires the creation of a 3D Network inside an Enterprise
Geodatabase in advance because,unfortunately,arcpy does not currently provide
a method to create a network dataset from scratch. For this example, the EGDB
was created on PostgreSQL.
--------------------------------------------------------------------------------
"""
from __future__ import print_function
import os
import glob
import re
import time
# import tempfile
import inspect
import repr as reprlib
from win32com import client
import pandas as pd
import arcpy
import arcpy.na
from arcpy import env
env.overwriteOutput = True
env.qualifiedFieldNames = "UNQUALIFIED"

# import pdb

# PR_PATH = os.getcwd()
# reformatting path strings to have forward slashes, otherwise AutoCAD fails.
# PR_PATH = PR_PATH.replace('\\', '/') + '/'
# print(PR_PATH)
# exploring the possibility of creating a temporary directory for geoprocessing
# this solution is multiplatform.
# SCRATCH_PATH = tempfile.mkdtemp()
# SCRATCH_PATH = SCRATCH_PATH.replace('\\', '/') + '/'
# print('Scratch {}'.format(SCRATCH_PATH))
# env.scratchWorkspace = 'C:\Users\ulisesdario\Documents\ArcGIS'
# print(SCRATCH_PATH)
# temp_gdb = env.scratchGDB


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
        for i, dwg in enumerate(folder_dwg_files):
            if dwg[-7:] == 'BAS.dwg' and 'ROOF' not in dwg:
                folder_dwg_files[i] = '/'.join([folder_path, dwg])
            else:
                folder_dwg_files.remove(dwg)
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


def autocadmap_to_shp(floor_plan, out_loc, layer_on, map_exp_set):
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
    map_exp_set (str) = A string representation of the map export settings
    location.

    Returns:
    A shapefile based on the provided floor plan. This will only contain
    the geometry that is present in the layer represented by the layer_on
    parameter.

    Examples:
    >>> autocadmap_to_shp(
    'C:/Users/ulisesdario/S-241E-01-DWG-BAS.dwg', 'C:/Users/ulisesdario',
    'A-SPAC-PPLN-AREA', 'G:/CAD-to-esri-3D-Network/mapexportsettings.epf')
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
    ex_set = map_exp_set.replace('\\', '/')
    pr = 'PROCEED'
    ex_command = '(command "{0}" "SHP" "{1}" "Y" "{2}"' \
                 ' "{3}")'.format(mp, out_name, ex_set, pr)
    doc.SendCommand('%s\n' % ex_command)
    doc.SendCommand("(ACAD-POP-DBMOD)\n")
    # this line avoids the 'Call was Rejected By Callee' Error.
    time.sleep(2)
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
    """Medial Axis, a work in progress.
    """
    try:
        # getting the name of the function programatically.
        print ('Executing {}... '.format(
            inspect.currentframe().f_code.co_name))
        # logic to create skeletons as feature classes
        skeletons = 'Medial axis algorithm'
    except Exception as e:
        print(e)
    return skeletons


def build_network(egdb, skeletons_list, network_skel_source, network, type):
    """It keeps the master network up to date by re-building its source
    features.The master network and its source features will be assumed to live
    in an enterprise geodatabase created on PostgreSQL 9.3.
    """
    # getting the name of the function programatically.
    print ('Executing {}... '.format(
        inspect.currentframe().f_code.co_name))
    env.workspace = egdb
    skeletons_list = arcpy.ListFeatureClasses("*{}".format(type))
    network_fields = {'name': 'TEXT', 'speed': 'FLOAT',
                      'minutes': 'DOUBLE', 'seconds': 'DOUBLE',
                      'elevation': 'SHORT'}
    # this is only being used as a proof of concept
    floor_heights = {'01': 0, '02': 10}
    for skeleton in skeletons_list:
        # network_scratch.ulisessol7.S_338_01_DWG_BAS_CENT
        # getting floor number.
        match = re.search('_(\d){2}_', skeleton)
        skeleton_floor = match.group()[1:-1]
        # getting field names
        field_names = [f.name for f in arcpy.ListFields(skeleton)]
        for field in network_fields:
            if field not in field_names:
                try:
                    arcpy.AddField_management(
                        skeleton, field, network_fields[field])
                except arcpy.ExecuteError:
                    print (arcpy.GetMessages(2))
    # to get a list of fields for the cursor object
    # map is slightly slower than a list comprehension, but clearer in this
    # particular case.
        cursor_fields = sorted(
            ['SHAPE@LENGTH'] + map(str.lower, network_fields.keys()))
        # print(cursor_fields)
        # ['SHAPE@LENGTH', 'elevation', 'minutes', 'name', 'seconds', 'speed']
        with arcpy.da.UpdateCursor(skeleton, cursor_fields) as cursor:
            for row in cursor:
                # updating the 'SPEED' field, the speed is expressed in fps
                row[5] = 4.11
                # updating the 'NAME' field
                row[3] = 'Placeholder'
                # updating the 'MINUTES' field
                row[2] = row[0] / (row[5] * 60)
                # updating the 'SECONDS' field
                row[4] = row[0] / row[5]
                # updating the 'ELEVATION' field
                row[1] = floor_heights[skeleton_floor]
                cursor.updateRow(row)
        arcpy.CheckOutExtension('3D')
        arcpy.CheckOutExtension('Network')
        # utilizing 'in_memory' workspace for faster performance.
        output_3d = 'in_memory' + '\\' + skeleton + '_3D'
        # important step, only 3D polylines can be connected in the Z dimension
        # in ArcScene
        arcpy.FeatureTo3DByAttribute_3d(skeleton, output_3d, 'ELEVATION')
    env.workspace = 'in_memory'
    append_candidates = arcpy.ListFeatureClasses('*_3D')
    # [arcpy.RegisterAsVersioned_management(sk) for sk in skeletons_list]
    # the network , and its sources, must be register as versioned beforehand,
    # the append candidates, on the other hand  must not.
    arcpy.Append_management(
        append_candidates, network_skel_source, 'NO_TEST')
    # not working properly, sometimes it does not re-build the network.
    arcpy.BuildNetwork_na(network)
    # cleanning up the 'in_memory' workspace.
    arcpy.Delete_management('in_memory')
    arcpy.CheckInExtension('3D')
    arcpy.CheckInExtension('Network')
    return

# build_network(
#     'master_network.sde', '', 'network_scratch.ulisessol7.pedestrian3D',
#     'network_scratch.ulisessol7.CU_Boulder_Network')


def centrality_calculator():
    """
    """

if __name__ == '__main__':
    bldgs_excel_path = 'G:\\CAD-to-esri-3D-Network\\qryAllBldgs.xlsx'
    floorplans_dir = 'G:\\CAD-to-esri-3D-Network\\floorplans'
    dwgs_dict = dwg_file_collector(bldgs_dict(bldgs_excel_path),
                                   floorplans_dir)[0]
    shapefiles_dir = 'G:\CAD-to-esri-3D-Network\shapefiles'
    network_loc_layer = 'A-SPAC-PPLN-AREA'
    network_skel_layer = 'A-CENT'
    autocad_exp_settings = 'G:\CAD-to-esri-3D-Network\mapexportsettings.epf'
    egdb = 'Database Connections\\master_network.sde'
    network_3d = egdb + '\\' + 'network_scratch.ulisessol7.CU_Boulder_Network'
    network_links_source = egdb + '\\' + \
        'network_scratch.ulisessol7.pedestrian3D'
    network_locs_source = egdb + '\\' + \
        'network_scratch.ulisessol7.location_3D'
    for k, v in dwgs_dict.iteritems():
        for floorplan in v:
            if ('338-01' in floorplan) or ('338-02' in floorplan):
                # this will create the network locations.
                # autocadmap_to_shp(floorplan, shapefiles_dir, network_loc_layer,
                #                   autocad_exp_settings)
                # the skeletons were created manually in AutoCAD MAp :(
                autocadmap_to_shp(floorplan, shapefiles_dir,
                                  network_skel_layer, autocad_exp_settings)
                shapefiles = shp_files_reader(shapefiles_dir)[1]
                shp_to_fc(shapefiles, egdb)
                # processing skeletons
                build_network(
                    egdb, '', network_links_source, network_3d, 'CENT')
            else:
                print('not processed: {}'.format(floorplan))
