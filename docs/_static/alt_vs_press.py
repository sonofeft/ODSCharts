"""
This example demonstrates a simple chart
"""
from odscharts.spreadsheet import SpreadSheet

mySprSht = SpreadSheet()

list_of_rows = [['Altitude','Pressure'], ['feet','psia'], 
                [0, 14.7],    [5000, 12.23], [10000, 10.11],   
                [30000, 4.36],[60000, 1.04]]

mySprSht.add_sheet('Alt_Data', list_of_rows)

mySprSht.add_scatter( 'Alt_P_Plot', 'Alt_Data',
                          title='Pressure vs Altitude', 
                          xlabel='Altitude', 
                          ylabel='Pressure', 
                          xcol=1,
                          ycolL=[2])
mySprSht.save( filename='alt_vs_press' )

mySprSht.launch_application()

