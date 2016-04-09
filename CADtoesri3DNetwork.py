from win32com import client

__author__ = 'Ulises  Guzman'
__author__ = 'Joseph Javadi'
date = '04/07/2016'


"""
Test
"""

acad = client.Dispatch("AutoCAD.Application")
acad.Visible = True
doc = acad.ActiveDocument
doc.SendCommand("ZOOM E\n")
print "Success!"
