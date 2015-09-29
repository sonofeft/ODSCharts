#!/usr/bin/env python
# -*- coding: ascii -*-

r"""
ODSCharts creates ods spreadsheet files readable by either Microsoft Excel or OpenOffice.

The format is a very narrow subset of full spreadsheet support::
    * All sheets contain either a table of numbers or a chart object
        - A table of numbers: 
            - starts at "A1"
            - row 1 is labels
            - row 2 is units
            - row 3 through N is float or string entries
        - Chart objects are scatter plots
            - Each series is a column from a table
            - Each x axis is a column from a table

There is no attempt to supply a full API interface.



ODSCharts
Copyright (C) 2015  Charlie Taylor

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.

-----------------------

"""
import zipfile
import os
import time
from copy import deepcopy

import sys
if sys.version_info < (3,):
    import ElementTree_27OD as ET
else:
    import ElementTree_34OD as ET

from data_table_desc import DataTableDesc
from plot_table_desc import PlotTableDesc
from metainf import add_ObjectN
from object_content import build_chart_object_content
from template_xml_file import TemplateXML_File

here = os.path.abspath(os.path.dirname(__file__))


# for multi-file projects see LICENSE file for authorship info
# for single file projects, insert following information
__author__ = 'Charlie Taylor'
__copyright__ = 'Copyright (c) 2015 Charlie Taylor'
__license__ = 'GPL-3'
exec( open(os.path.join( here,'_version.py' )).read() )  # creates local __version__ variable
__email__ = "cet@appliedpython.com"
__status__ = "3 - Alpha" # "3 - Alpha", "4 - Beta", "5 - Production/Stable"

TABLE_INSERT_POINT = 1  # just after "table:calculation-settings" Element


def load_template_xml_from_ods(ods_fname, fname, subdir='' ):
    
    full_ods_path = os.path.join( here, 'templates', ods_fname )
    
    if subdir:
        inner_fname = subdir + '/' + fname
    else:
        inner_fname =  fname
    
    odsfile = zipfile.ZipFile( full_ods_path )
    src = odsfile.read( inner_fname ).decode('utf-8')
    
    return TemplateXML_File( src )


def zipfile_insert( zipfileobj, filename, data):
    """Create a file named filename, into the zip archive.
       "data" is the string that is placed into filename.
    """

    # zip seems to struggle with non-ascii characters
    #data = data.encode('utf-8')

    now = time.localtime(time.time())[:6]
    info = zipfile.ZipInfo(filename)
    info.date_time = now
    info.compress_type = zipfile.ZIP_DEFLATED
    zipfileobj.writestr(info, data)


class MySheetNameError(Exception):
    """Custom exception handler for duplicate sheet names"""
    def __init__(self, msg):
        self.msg = msg
    def __str__(self):
        return repr(self.msg)

class SpreadSheet(object):
    """Creates OpenDocument Spreadsheets with charts for Microsoft Excel and OpenOffice
    """

    def __init__(self):
        """Inits SpreadSheet with filename and blank content."""
        
        self.filename = None
        self.data_table_objD = {} # dict of data desc objects (DataTableDesc)
        self.plot_table_objD = {} # dict of plot desc objects (PlotTableDesc)
        self.ordered_plotL = [] # list of plot names in insertion order 
        self.ordered_dataL = [] # list of data sheet names in insertion order
        
        self.content_xml_obj = load_template_xml_from_ods( 'alt_chart.ods', 'content.xml' )
        self.meta_xml_obj = load_template_xml_from_ods( 'alt_chart.ods', 'meta.xml' )
        self.mimetype_str = 'application/vnd.oasis.opendocument.spreadsheet'
        self.styles_xml_obj = load_template_xml_from_ods( 'alt_chart.ods', 'styles.xml' )

        self.metainf_manifest_xml_obj = load_template_xml_from_ods('empty_sheets123.ods', 'manifest.xml', subdir='META-INF')
        
        self.template_ObjectN_styles_xml_obj = load_template_xml_from_ods( 'alt_chart.ods', 'styles.xml', subdir='Object 1')
        
        # Clean up template for content (remove default table and graph)
        
        self.spreadsheet_obj = self.content_xml_obj.find('office:body/office:spreadsheet')
        
        self.meta_creation_date_obj = self.meta_xml_obj.find('office:meta/meta:creation-date')
        self.meta_dc_date_obj = self.meta_xml_obj.find('office:meta/dc:date')
        
        self.meta_init_creator_obj = self.meta_xml_obj.find('office:meta/meta:initial-creator')
        self.meta_dc_creator_obj = self.meta_xml_obj.find('office:meta/dc:creator')
        
        # Remove the empty sheets from the template spreadsheet
        tableL = self.content_xml_obj.findall('office:body/office:spreadsheet/table:table')
                
        #tableL = self.spreadsheet_obj.findall('table:table')
        for table in tableL:
            #table.getparent().remove(table)
            parent = self.content_xml_obj.getparent( table )
            self.content_xml_obj.remove_child( table, parent )
        
        #table2 = tableL[1]
        
        #table1 = tableL[0]
        #table1.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name', 'Sheet1')
        #print( table1 )
        #print( table2.items() )
        

    def meta_time(self):
        "Return time string in meta data format"
        t = time.localtime()
        stamp = "%04d-%02d-%02dT%02d:%02d:%02d" % (t[0], t[1], t[2], t[3], t[4], t[5])
        return stamp


    def add_scatter(self, plot_sheetname, data_sheetname, 
                      title='', xlabel='', ylabel='', y2label='',
                      xcol=1,
                      ycolL=None, ycol2L=None,
                      labelL=None, label2L=None):
        """Add a scatter plot to the spread sheet.
        
           Use data from "data_sheetname" to create "plot_sheetname" with scatter plot.
           
           Assume index into columns is "1-based" such that column "A" is 1, "B" is 2, etc.
           
        """
        # Don't allow duplicate sheet names
        if (plot_sheetname in self.data_table_objD) or (plot_sheetname in self.plot_table_objD):
            raise  MySheetNameError('Duplicate sheet name submitted for new plot: "%s"'%plot_sheetname)

        # Data sheet must already exist in order to make a plot
        if (data_sheetname not in self.data_table_objD):
            raise  MySheetNameError('Data sheet for "%s" plot missing: "%s"'%(plot_sheetname, data_sheetname))

        num_chart = len(self.plot_table_objD) + 1
        add_ObjectN( num_chart, self.metainf_manifest_xml_obj)
        
        plot_desc = PlotTableDesc( plot_sheetname, num_chart, self.content_xml_obj)
        
        
        self.spreadsheet_obj.insert(TABLE_INSERT_POINT, plot_desc.table_obj)
        
        #obj_name = 'Object %i'%num_chart
        
        # ================= save PlotTableDesc object
        self.plot_table_objD[plot_sheetname] = plot_desc
        self.ordered_plotL.append( plot_sheetname )
        
        # Start making the chart object that goes onto the plot sheet 
        #  Assign plot parameters to PlotTableDesc object
        plot_desc.plot_sheetname = plot_sheetname
        plot_desc.data_sheetname = data_sheetname
        plot_desc.title = title
        plot_desc.xlabel = xlabel
        plot_desc.ylabel = ylabel
        plot_desc.y2label = y2label
        plot_desc.xcol = xcol
        plot_desc.ycolL = ycolL
        plot_desc.ycol2L = ycol2L
        plot_desc.labelL = labelL
        plot_desc.label2L = label2L
        
        # assigns attribute chart_obj to plot_desc
        table_desc = self.data_table_objD[data_sheetname]
        
        # create a new chart object
        chart_obj = load_template_xml_from_ods('alt_chart.ods', 'content.xml', subdir='Object 1')
        
        build_chart_object_content( chart_obj, plot_desc, table_desc )
        

    def add_sheet(self, data_sheetname, list_of_rows):
        """Create a new sheet in the spreadsheet with "data_sheetname" as its name.
           
           the list_of_rows will be placed at "A1" and should be: 
            - row 1 is labels
            - row 2 is units
            - row 3 through N is float or string entries
           
            for example:
            list_of_rows = 
            [['Altitude','Pressure','Temperature'], 
            ['feet','psia','degR'], 
            [0, 14.7, 518.7], [5000, 12.23, 500.8], 
            [10000, 10.11, 483.0], [60000, 1.04, 390]]
        """
        if (data_sheetname in self.data_table_objD) or (data_sheetname in self.plot_table_objD):
            raise  MySheetNameError('Duplicate sheet name submitted for new datasheet: "%s"'%data_sheetname)
                
                
        table_desc = DataTableDesc( data_sheetname, list_of_rows, self.content_xml_obj)
        self.spreadsheet_obj.insert(TABLE_INSERT_POINT, table_desc.table_obj)
        
        self.data_table_objD[data_sheetname] = table_desc
        self.ordered_dataL.append( table_desc )
            

    def save(self, filename='my_chart.ods', debug=False):
        """Saves SpreadSheet to an ods file readable by Microsoft Excel or OpenOffice"""
        
        if not filename.lower().endswith('.ods'):
            filename = filename + '.ods'
        
        print('Saving ods file: %s'%filename)
        self.filename = filename
        
        zipfileobj = zipfile.ZipFile(filename, "w")
        
        self.meta_creation_date_obj.text = self.meta_time()
        self.meta_dc_date_obj.text = self.meta_time()
        self.meta_init_creator_obj.text = 'ODSCharts'
        self.meta_dc_creator_obj.text = 'ODSCharts'
        
        
        
        zipfile_insert( zipfileobj, 'meta.xml', self.meta_xml_obj.tostring())
        
        zipfile_insert( zipfileobj, 'mimetype', self.mimetype_str.encode('UTF-8'))
        
        zipfile_insert( zipfileobj, 'META-INF/manifest.xml', self.metainf_manifest_xml_obj.tostring())
                                    
        for N, plot_sheetname in enumerate( self.ordered_plotL ):
            
            plot_desc = self.plot_table_objD[ plot_sheetname ]
            
            zipfile_insert( zipfileobj, 'Object %i/styles.xml'%(N+1,), self.template_ObjectN_styles_xml_obj.tostring())
            
            zipfile_insert( zipfileobj, 'Object %i/content.xml'%(N+1,), plot_desc.chart_obj.tostring())
                                        
                                    
        zipfile_insert( zipfileobj, 'content.xml', self.content_xml_obj.tostring())
        
        zipfile_insert( zipfileobj, 'styles.xml', self.styles_xml_obj.tostring())
        

if __name__ == '__main__':
    C = SpreadSheet()
    C.save( filename='performance', debug=False)
