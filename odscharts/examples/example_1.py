"""
This example demonstrates the proper use of project: odscharts
"""
import sys
import os

here = os.path.abspath(os.path.dirname(__file__))

sys.path.insert(0, os.path.abspath( os.path.join(here, "../" ) ) )  # needed to find odscharts development version
sys.path.insert(0, os.path.abspath( os.path.join(here, "../../" ) ) )  # needed to find odscharts development version

from odscharts.spreadsheet import SpreadSheet


my_class = SpreadSheet()

list_of_rows = [['Altitude','Pressure','Temperature','Temperature'], 
                ['feet','psia','degR','degK'], 
                [0, 14.7, 518.7, 288.1667], [5000, 12.23, 500.8, 278.2222], 
                [10000, 10.11, 483.0, 268.3333], [30000, 4.36, 411.8, 228.7778],
                [60000, 1.04, 390, 216.6667]]

my_class.add_sheet('Alt_Data', list_of_rows)
#my_class.save( filename=os.path.join(here,'alt_table'), debug=False)


my_class.add_scatter( 'Alt_Plot', 'Alt_Data',
                          title='Pressure vs Altitude', xlabel='Altitude', 
                          ylabel='Pressure', y2label='Unittest Y2 Axis',
                          xcol=1,
                          ycolL=[2], ycol2L=None,
                          labelL=None, label2L=None)
#my_class.save( filename=os.path.join(here,'alt_1_plot'), debug=False)
                          
my_class.add_scatter( 'Alt_2_Plot', 'Alt_Data',
                          title='Temperature vs Altitude', xlabel='Altitude', 
                          ylabel='Temperature', y2label='Unittest Y2 Axis',
                          xcol=1,
                          ycolL=[3,4], ycol2L=None,
                          labelL=None, label2L=None)
#my_class.save( filename=os.path.join(here,'alt_2_plot'), debug=False)

my_class.add_scatter( 'Alt_3_Plot', 'Alt_Data',
                          title='T&P vs Altitude', xlabel='Altitude', 
                          ylabel='T and P', y2label='Unittest Y2 Axis',
                          xcol=1,
                          ycolL=[2,3,4], ycol2L=None,
                          labelL=None, label2L=None)
my_class.save( filename=os.path.join(here,'alt_3_plot'), debug=False)

