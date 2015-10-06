"""
This example demonstrates the use of line styles
"""
from math import *
from odscharts.spreadsheet import SpreadSheet
from odscharts.line_styles import LINE_STYLE_LIST as sL

mySprSht = SpreadSheet()
rev_orderL = list( range(10,-1,-1) )

toprow = ['Line Style Index']
unitsrow = ['']
toprow.extend( ['%i) %s'%(i,sL[i]) for i in rev_orderL] )
unitsrow.extend( ['' for i in rev_orderL] )

xbegrow = [0]
xbegrow.extend( [i for i in rev_orderL] )
xendrow = [1]
xendrow.extend( [i for i in rev_orderL] )

list_of_rows = [toprow, unitsrow, xbegrow, xendrow]

mySprSht.add_sheet('Line_Style_Data', list_of_rows)

mySprSht.add_scatter( 'Line_Style', 'Line_Style_Data',
                          title='Line Styles', 
                          xlabel='', 
                          ylabel='Index of Line Style', 
                          xcol=1,
                          ycolL=[2,3,4,5,6,7,8,9,10,11,12], 
                          lineStyleL=[i for i in rev_orderL],
                          lineThkL = [1.5],
                          showMarkerL=[0])

mySprSht.setYrange( ymin=-1, ymax=None, plot_sheetname=None)
mySprSht.save( filename='line_style_plot', launch=True)