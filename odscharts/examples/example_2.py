"""
This example demonstrates large data sets
"""
import sys
import os
from math import *

here = os.path.abspath(os.path.dirname(__file__))

sys.path.insert(0, os.path.abspath( os.path.join(here, "../../" ) ) )  # needed to find odscharts development version

from odscharts.spreadsheet import SpreadSheet


my_class = SpreadSheet()

list_of_rows = [['Angle','Sine','Cosine'], ['deg','','']]
for iang in xrange( 3601 ):
    ang_deg = float(iang) / 10.0
    ang = radians(ang_deg)
    
    list_of_rows.append( [ang_deg, sin(ang), cos(ang)] )

my_class.add_sheet('Trig_Data', list_of_rows)

my_class.add_scatter( 'Trig_Plot', 'Trig_Data',
                          title='Trig Functions', xlabel='Angle', 
                          ylabel='Trig Function', y2label='Unittest Y2 Axis',
                          xcol=1,
                          ycolL=[2,3], ycol2L=None,
                          labelL=None, label2L=None)
my_class.save( filename=os.path.join(here,'trig_plot'), debug=False)