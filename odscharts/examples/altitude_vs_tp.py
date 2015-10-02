"""
This example demonstrates the proper use of project: odscharts
"""
import sys
import os


# These inserts are needed to find odscharts development version ONLY in special cases
#here = os.path.abspath(os.path.dirname(__file__))
#sys.path.insert(0, os.path.abspath( os.path.join(here, "../" ) ) )  
#sys.path.insert(0, os.path.abspath( os.path.join(here, "../../" ) ) )

from odscharts.spreadsheet import SpreadSheet


mySprSht = SpreadSheet()

list_of_rows = [['Altitude','Pressure','Temp R','Temp K'], 
                ['feet','psia','degR','degK'], 
                [0, 14.7, 518.7, 288.1667], [5000, 12.23, 500.8, 278.2222], 
                [10000, 10.11, 483.0, 268.3333], [30000, 4.36, 411.8, 228.7778],
                [60000, 1.04, 390, 216.6667]]

mySprSht.add_sheet('Alt_Data', list_of_rows)

mySprSht.add_scatter( 'Alt_P_Plot', 'Alt_Data',
                          title='Pressure vs Altitude', xlabel='Altitude', 
                          ylabel='Pressure', y2label='Unittest Y2 Axis',
                          xcol=1,
                          ycolL=[2])

mySprSht.add_scatter( 'Alt_T_Plot', 'Alt_Data',
                          title='Temperature vs Altitude', xlabel='Altitude', 
                          ylabel='Temperature', y2label='Unittest Y2 Axis',
                          xcol=1,
                          ycolL=[3,4])
                          
mySprSht.add_scatter( 'Alt_PT_Plot', 'Alt_Data',
                          title='Temperature vs Altitude', xlabel='Altitude', 
                          ylabel='Temperature and Pressure', y2label='',
                          xcol=1,
                          ycolL=[2,3,4])

mySprSht.add_scatter( 'Alt_P_T2_Plot', 'Alt_Data',
                          title='T&P vs Altitude', xlabel='Altitude', 
                          ylabel='Pressure', y2label='Temperature',
                          xcol=1,
                          ycolL=[2], ycol2L=[3,4])
mySprSht.save( filename='altitude_vs_tp' )

mySprSht.launch_application()

