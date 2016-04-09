# -*- coding: utf-8 -*-
"""CAD to esri 3D Network"""
from __future__ import print_function
import os
import glob
from win32com import client

__author__ = 'Ulises  Guzman'
__author__ = 'Joseph Javadi'
date = '04/07/2016'


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


def cad_to_fc():
    """
    """
    # Ulises
    return


def topology_checks():
    """
    """
    # Ulises
    return
