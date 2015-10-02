"""
This example demonstrates large data sets
"""
import sys
import os

# These inserts are needed to find odscharts development version ONLY in special cases
#here = os.path.abspath(os.path.dirname(__file__))
#sys.path.insert(0, os.path.abspath( os.path.join(here, "../" ) ) )  
#sys.path.insert(0, os.path.abspath( os.path.join(here, "../../" ) ) )

from math import *
from odscharts.spreadsheet import SpreadSheet

mySprSht = SpreadSheet()

list_of_rows = [['Angle','Sine','Cosine'], ['deg','','']]
for iang in range( 3601 ):
    ang_deg = float(iang) / 10.0
    ang = radians(ang_deg)
    
    list_of_rows.append( [ang_deg, sin(ang), cos(ang)] )

mySprSht.add_sheet('Trig_Data', list_of_rows)

mySprSht.add_scatter( 'Trig_Plot', 'Trig_Data',
                          title='Trig Functions', 
                          xlabel='Angle', 
                          ylabel='Trig Function', 
                          xcol=1,
                          ycolL=[2,3], showMarkerL=[1,0])
                          
mySprSht.save( filename='trig_plot', launch=True)