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
from lxml import etree as ET

here = os.path.abspath(os.path.dirname(__file__))


# for multi-file projects see LICENSE file for authorship info
# for single file projects, insert following information
__author__ = 'Charlie Taylor'
__copyright__ = 'Copyright (c) 2015 Charlie Taylor'
__license__ = 'GPL-3'
exec( open(os.path.join( here,'_version.py' )).read() )  # creates local __version__ variable
__email__ = "cet@appliedpython.com"
__status__ = "3 - Alpha" # "3 - Alpha", "4 - Beta", "5 - Production/Stable"


def load_template_xml( fname, subdir='' ):
    if subdir:
        full_path = os.path.join( here, 'templates', subdir, fname )
    else:
        full_path = os.path.join( here, 'templates', fname )
    return ET.parse( full_path )


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

class SpreadSheet(object):
    """Creates OpenDocument Spreadsheets with charts for Microsoft Excel and OpenOffice
    """

    def __init__(self):
        """Inits SpreadSheet with filename and blank content."""
        
        self.filename = None
        self.sheet_objL = [] # list of sheet objects
        self.chart_objL = [] # list of chart objects
        
        self.content_xml_obj = load_template_xml( 'content.xml' )
        self.meta_xml_obj = load_template_xml( 'meta.xml' )
        self.mimetype_str = 'application/vnd.oasis.opendocument.spreadsheet'
        self.styles_xml_obj = load_template_xml( 'styles.xml' )

        self.metainf_manifest_xml_obj = load_template_xml( 'manifest.xml', subdir='META-INF')
        
        # Clean up template for content (remove default table and graph)
        self.content_root = self.content_xml_obj.getroot()
        
        tableL = self.content_root.findall('office:body/office:spreadsheet/table:table', self.content_root.nsmap)
        self.sheet_objL.extend( tableL )
        self.table_insert_point = 1  # just after "table:calculation-settings" Element
        
        self.spreadsheet_obj = self.content_root.find('office:body/office:spreadsheet', self.content_root.nsmap)
        #table2 = tableL[1]
        
        #table1 = tableL[0]
        #table1.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name', 'Sheet1')
        #print( table1 )
        #print( table2.items() )
        

    def add_sheet(self, sheetname, list_of_rows):
        """Create a new sheet in the spreadsheet with "sheetname" as its name.

           the list_of_rows will be placed at "A1" and should be: 
            - row 1 is labels
            - row 2 is units
            - row 3 through N is float or string entries
           
            for example:
            list_of_rows = [['Altitude','Pressure','Temperature'], 
                            ['feet','psia','degR'], 
                            [0, 14.7, 518.7], [5000, 12.23, 500.8], 
                            [10000, 10.11, 483.0], [60000, 1.04, 390]]
           
        """
        def fix_ns( shortname ): 
            """remove namespace shortcuts from name"""
            sL = shortname.split(':')
            if len(sL)!=2:
                return shortname
                
            return '{%s}'%self.content_root.nsmap[sL[0]] + sL[1]
                
        #print( self.content_root.nsmap )
        attribD = {fix_ns('table:name'):sheetname, fix_ns('table:style-name'):"ta1"}
        newsheet = ET.Element(fix_ns('table:table'), attrib=attribD, nsmap=self.content_root.nsmap)
        
        colD = {fix_ns('table:style-name'):"co1", fix_ns('table:number-columns-repeated'):"16384", 
                fix_ns('table:default-cell-style-name'):"ce1"}
        col_elm = ET.Element(fix_ns('table:table-column'), attrib=colD, nsmap=self.content_root.nsmap)
        newsheet.append( col_elm )
        
        str_cellD = {fix_ns('office:value-type'):"string", fix_ns('table:style-name'):"ce1"}
        float_cellD = {fix_ns('office:value-type'):"float", fix_ns('table:style-name'):"ce1"}
            
        row_rep = 1048576 # max number of rows
            
        for row in list_of_rows:
            row_obj = ET.Element(fix_ns('table:table-row'), 
                                 attrib={fix_ns('table:style-name'):"ro1"}, 
                                 nsmap=self.content_root.nsmap)
                
            col_rep = 16384 # max number of columns
                
            for value in row:
                try:
                    fval = float( value )
                    float_cellD[fix_ns('office:value')] = str(fval)
                    D = float_cellD
                    text_p = "%g"%fval
                except:
                    D = str_cellD
                    text_p = "%s"%value
                
                cell_obj = ET.Element(fix_ns('table:table-cell'), attrib=D, nsmap=self.content_root.nsmap)
                text_obj = ET.Element(fix_ns('text:p'),  nsmap=self.content_root.nsmap)
                text_obj.text = text_p
                cell_obj.append( text_obj )
                row_obj.append(cell_obj)
                col_rep -= 1 # decrement max available columns remaining

            last_cell_obj = ET.Element(fix_ns('table:table-cell'), 
                                       attrib={fix_ns('table:number-columns-repeated'):"%i"%col_rep}, 
                                       nsmap=self.content_root.nsmap)
            row_obj.append(last_cell_obj)
            newsheet.append( row_obj )
            row_rep -= 1 # decrement max remaining number of rows
            
        last_row_obj = ET.Element(fix_ns('table:table-row'), 
                             attrib={fix_ns('table:style-name'):"ro1", 
                                     fix_ns('table:number-rows-repeated'):"%i"%row_rep}, 
                             nsmap=self.content_root.nsmap)
        last_row_obj.append( ET.Element(fix_ns('table:table-cell'), 
                                        attrib={fix_ns('table:number-columns-repeated'):"16384"}, 
                                        nsmap=self.content_root.nsmap) )

        newsheet.append( last_row_obj )
        self.sheet_objL.append( newsheet )
        self.spreadsheet_obj.insert(self.table_insert_point, newsheet)
        self.table_insert_point += 1
            

    def save(self, filename='my_chart.ods', debug=False):
        """Saves SpreadSheet to an ods file readable by Microsoft Excel or OpenOffice"""
        
        if not filename.lower().endswith('.ods'):
            filename = filename + '.ods'
        
        print('Saving ods file: %s'%filename)
        self.filename = filename
        
        zipfileobj = zipfile.ZipFile(filename, "w")
        
        zipfile_insert( zipfileobj, 'meta.xml', 
                        ET.tostring(self.meta_xml_obj, xml_declaration=True, 
                                    encoding="UTF-8", standalone=True))
        zipfile_insert( zipfileobj, 'mimetype', self.mimetype_str.encode('UTF-8'))
        zipfile_insert( zipfileobj, 'META-INF/manifest.xml', 
                        ET.tostring(self.metainf_manifest_xml_obj, xml_declaration=True, 
                                    encoding="UTF-8", standalone=True))
        zipfile_insert( zipfileobj, 'content.xml', 
                        ET.tostring(self.content_xml_obj, xml_declaration=True, 
                                    encoding="UTF-8", standalone=True))
        zipfile_insert( zipfileobj, 'styles.xml', 
                        ET.tostring(self.styles_xml_obj, xml_declaration=True, 
                                    encoding="UTF-8", standalone=True))
        

if __name__ == '__main__':
    C = SpreadSheet()
    C.save( filename='performance', debug=False)
