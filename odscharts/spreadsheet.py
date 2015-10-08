#!/usr/bin/env python
# -*- coding: ascii -*-

r"""
ODSCharts creates ods spreadsheet files readable by either Microsoft Excel, 
LibreOffice or OpenOffice.

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
from copy import deepcopy, copy
import  subprocess

import sys
if sys.version_info < (3,):
    import odscharts.ElementTree_27OD as ET
else:
    import odscharts.ElementTree_34OD as ET

from odscharts.data_table_desc import DataTableDesc
from odscharts.plot_table_desc import PlotTableDesc
from odscharts.metainf import add_ObjectN
from odscharts.object_content import build_chart_object_content

from odscharts.template_xml_file import TemplateXML_File
from odscharts.find_obj import find_elem_w_attrib, elem_set, NS_attrib, NS

from odscharts.line_styles import gen_dash_elements_from_set_of_istyles

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
    """Create a file named filename, in the zip archive.
       "data" is the UTF-8 string that is placed into filename.

       (Not called by User)
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
    """Creates OpenDocument Spreadsheets with charts for Microsoft Excel, LibreOffice and OpenOffice
    """

    def launch_application(self):
        """
        Will launch Excel, LibreOffice or Openoffice using "os.startfile" or "open" or
        "xdg-open" depending on the platform.

        ONLY WORKS IF file has been saved.
        """

        if self.filename is None:
            print( '='*75 )
            print( '='*5, 'MUST SAVE FILE before launch_application will work.' , '='*5)
            print( '='*75 )
            return

        #os.startfile( self.filename )
        if sys.platform == "win32":
            os.startfile(self.filename)
        else:
            opener ="open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, self.filename])

    def __init__(self):
        """Inits SpreadSheet with filename and blank content."""

        self.filename = None
        self.data_table_objD = {} # dict of data desc objects (DataTableDesc), index=data_sheetname, value=Obj
        self.plot_sheet_objD = {} # dict of plot desc objects (PlotTableDesc), index=plot_sheetname, value=Obj
        self.ordered_plotL = [] # list of plot names in insertion order
        self.ordered_dataL = [] # list of data sheet names in insertion order

        self.plot_xMinMaxD = {} # index=plot sheet name, values=(xmin,xmax). can be (None,None)
        self.plot_yMinMaxD = {} # index=plot sheet name, values=(ymin,ymax). can be (None,None)
        self.plot_y2MinMaxD = {} # index=plot sheet name, values=(y2min,y2max). can be (None,None)

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

    # make sure any added Element objects are in nsOD, rev_nsOD and qnameOD of parent_obj
    def add_tag(self, tag, parent_obj ):
        sL = tag.split('}')
        uri = sL[0][1:]
        name = sL[1]

        parent_obj.qnameOD[tag] = parent_obj.nsOD[uri] + ':' + name


    def setAxisRange(self, axis_name, min_val=None, max_val=None, plot_sheetname=None):
        """
        Generic routine used by setXrange, setYrange and setY2range. (Not called by User)

        Set the Axis range of the named plot to min_val, max_val.

            * If no plot_sheetname is provided, use the latest plot (if any).
            * If no min_val is provided, ignore it.
            * If no max_val is provided, ignore it.
        """

        # use latest plot_sheetname if no name is provided
        if plot_sheetname is None:
            plot_sheetname = self.ordered_plotL[-1]

        plotSheetObj = self.plot_sheet_objD[plot_sheetname]
        chart_obj = plotSheetObj.chart_obj
        nsOD = chart_obj.rev_nsOD

        auto_styles = chart_obj.find('office:automatic-styles')
        minmax_style,ipos_minmax_style = find_elem_w_attrib('style:style', auto_styles, nsOD,
                                   attrib={'style:name':'%s'%axis_name}, nth_match=0)
        #print( 'FOUND:  ipos_minmax_style = ', ipos_minmax_style )

        chart_prop = minmax_style.find( NS('style:chart-properties', nsOD) )

        if not min_val is None:
            chart_prop.set( NS('chart:minimum', nsOD), '%g'%min_val )
            self.add_tag( NS('chart:minimum', nsOD), chart_obj )

        if not max_val is None:
            chart_prop.set( NS('chart:maximum', nsOD), '%g'%max_val )
            self.add_tag( NS('chart:maximum', nsOD), chart_obj )

    def setAxisRanges(self, plot_sheetname ):

        if plot_sheetname in self.plot_xMinMaxD:
            xmin, xmax = self.plot_xMinMaxD[ plot_sheetname ]
            self.setAxisRange( 'Axs0' ,min_val=xmin, max_val=xmax, plot_sheetname=plot_sheetname)

        if plot_sheetname in self.plot_yMinMaxD:
            ymin, ymax = self.plot_yMinMaxD[ plot_sheetname ]
            self.setAxisRange( 'Axs1' ,min_val=ymin, max_val=ymax, plot_sheetname=plot_sheetname)

        if plot_sheetname in self.plot_y2MinMaxD:
            ymin, ymax = self.plot_y2MinMaxD[ plot_sheetname ]
            self.setAxisRange( 'Axs2' ,min_val=ymin, max_val=ymax, plot_sheetname=plot_sheetname)


    def setXrange(self, xmin=None, xmax=None, plot_sheetname=None):
        """
        Set the X range of the named plot to xmin, xmax.

            * If no plot_sheetname is provided, use the latest plot (if any).
            * If no min_val is provided, ignore it.
            * If no max_val is provided, ignore it.

        :keyword float xmin: Minimum value on X Axis (default==None)
        :type  xmin: float or None
        :keyword float xmax: Maximum value on X Axis (default==None)
        :type  xmax: float or None
        :keyword None plot_sheetname:  (default==None)
        :type  plot_sheetname: None or str or unicode
        :return: None
        :rtype: None

        """
        #self.setAxisRange( 'Axs0' ,min_val=xmin, max_val=xmax, plot_sheetname=plot_sheetname)


        # use latest plot_sheetname if no name is provided
        if plot_sheetname is None:
            plot_sheetname = self.ordered_plotL[-1]

        self.plot_xMinMaxD[plot_sheetname] = (xmin, xmax)


    def setYrange(self, ymin=None, ymax=None, plot_sheetname=None):
        """
        Set the Y range of the named plot to ymin, ymax.

            * If no plot_sheetname is provided, use the latest plot (if any).
            * If no min_val is provided, ignore it.
            * If no max_val is provided, ignore it.

        :keyword float ymin: Minimum value on Primary Y Axis (default==None)
        :type  ymin: float or None
        :keyword float ymax: Maximum value on Primary Y Axis (default==None)
        :type  ymax: float or None
        :keyword None plot_sheetname:  (default==None)
        :type  plot_sheetname: None or str or unicode
        :return: None
        :rtype: None
        """
        #self.setAxisRange( 'Axs1' ,min_val=ymin, max_val=ymax, plot_sheetname=plot_sheetname)

        # use latest plot_sheetname if no name is provided
        if plot_sheetname is None:
            plot_sheetname = self.ordered_plotL[-1]

        self.plot_yMinMaxD[plot_sheetname] = (ymin, ymax)


    def setY2range(self, ymin=None, ymax=None, plot_sheetname=None):
        """
        Set the Y2 range of the named plot to ymin, ymax.

            * If no plot_sheetname is provided, use the latest plot (if any).
            * If no min_val is provided, ignore it.
            * If no max_val is provided, ignore it.

        :keyword float ymin: Minimum value on Secondary Y Axis (default==None)
        :type  ymin: float or None
        :keyword float ymax: Maximum value on Secondary Y Axis (default==None)
        :type  ymax: float or None
        :keyword None plot_sheetname:  (default==None)
        :type  plot_sheetname: None or str or unicode
        :return: None
        :rtype: None
        """
        #self.setAxisRange( 'Axs2' ,min_val=ymin, max_val=ymax, plot_sheetname=plot_sheetname)

        # use latest plot_sheetname if no name is provided
        if plot_sheetname is None:
            plot_sheetname = self.ordered_plotL[-1]

        self.plot_y2MinMaxD[plot_sheetname] = (ymin, ymax)


    def add_curve(self, plot_sheetname, data_sheetname, xcol=1,
                    ycolL=None, ycol2L=None,
                    showMarkerL=None, showMarker2L=None,
                    showLineL=None, showLine2L=None,
                    lineThkL=None, lineThk2L=None,
                    lineStyleL=None, lineStyle2L=None,
                    colorL=None, color2L=None,
                    labelL=None, label2L=None):
        """        
        Use data from "data_sheetname" to add a plot or plots to scatter plot on
        page  "plot_sheetname".

        Most of the inputs are exactly like the "add_scatter" method.

        Assume index into columns is "1-based" such that column "A" is 1, "B" is 2, etc.

        NOTE: marker color is NOT supported in Excel from ODS format.
        Markers will appear, but will be some other default color.
        For this reason, ODSCharts uses Excel default color list unless the user
        overrides them with colorL or color2L.

        :param plot_sheetname: Name of the plot's tabbed window in Excel,LibreOffice or OpenOffice
        :type  plot_sheetname: str or unicode
        :param data_sheetname: Name of the data's tabbed window in Excel, LibreOffice or OpenOffice
        :type  data_sheetname: str or unicode

        :keyword int xcol: 1-based column index on data sheet of X Axis data (default==1)

        :keyword ycolL: List of 1-based column indeces on data sheet of Primary Y Axis data (default==None)
        :type  ycolL: None or int list
        :keyword ycol2L: List of 1-based column indeces on data sheet of Secondary Y Axis data  (default==None)
        :type  ycol2L: None or int list

        :keyword showMarkerL: List of boolean values indicating marker or no-marker for ycolL data (default==None)
        :type  showMarkerL: None or bool list
        :keyword showMarker2L: List of boolean values indicating marker or no-marker for ycol2L data (default==None)
        :type  showMarker2L: None or bool list

        :keyword showLineL: List of boolean values indicating line or no-line for ycolL data (default==None)
        :type  showLineL: None or bool list
        :keyword showLine2L: List of boolean values indicating line or no-line for ycol2L data (default==None)
        :type  showLine2L: None or bool list

        :keyword lineThkL: List of thickness values for primary y axis curves.
            If the list is not defined, then the thickness is set to "0.8mm".
            For example, lineThkL=[0.5, 1, 2] gives thin, about normal and thick lines.
            (Units=="mm")(default==None)
        :type  lineThkL: list of float
        :keyword lineThk2L: List of thickness values for secondary y axis curves
            If the list is not defined, then the thickness is set to "0.8mm".
            For example, lineThkL=[0.5, 1, 2] gives thin, about normal and thick lines.
            (Units=="mm")(default==None)
        :type  lineThk2L: list of float

        :keyword colorL: List of color values to use for ycolL curves. Format is "#FFCC33". Any curve not explicitly set will use Excel default colors (default==None)
        :type  colorL: None or str list
        :keyword color2L: List of color values to use for ycol2L curves. Format is "#FFCC33". Any curve not explicitly set will use Excel default colors (default==None)
        :type  color2L: None or str list
        :keyword labelL: Not yet implemented (default==None)
        :type  labelL: None or str list
        :keyword label2L: Not yet implemented (default==None)
        :type  label2L: None or str list

        :return: None
        :rtype: None
        
        
        """

        # use latest plot_sheetname if no name is provided
        if plot_sheetname is None:
            plot_sheetname = self.ordered_plotL[-1]

        # use latest data_sheetname if no name is provided
        if data_sheetname is None:
            data_sheetname = self.ordered_dataL[-1]

        # make sure that plot_sheetname and data_sheetname exist
        if plot_sheetname not in self.plot_sheet_objD:
            raise  MySheetNameError('Named plot sheet does NOT exist: "%s"'%plot_sheetname)

        # Data sheet must already exist in order to make a plot
        if data_sheetname not in self.data_table_objD:
            raise  MySheetNameError('Named data sheet does NOT exist: "%s"'%data_sheetname)

        # Make of copy of original plotSheetObj
        plotSheetObj = self.plot_sheet_objD[plot_sheetname]


        plotSheetObj.add_to_primary_y(data_sheetname, xcol, ycolL,
                                      showMarkerL=showMarkerL, colorL=colorL,
                                      lineThkL=lineThkL,
                                      lineStyleL=lineStyleL,
                                      labelL=labelL)

        plotSheetObj.add_to_secondary_y( data_sheetname, xcol, ycol2L,
                                         showMarker2L=showMarker2L, color2L=color2L,
                                         lineThk2L=lineThk2L,
                                         lineStyle2L=lineStyle2L,
                                         label2L=label2L)



    def add_scatter(self, plot_sheetname, data_sheetname,
                      title='', xlabel='', ylabel='', y2label='',
                      xcol=1, logx=False, logy=False, log2y=False, excel_colors=True,
                      ycolL=None, ycol2L=None,
                      showMarkerL=None, showMarker2L=None,
                      showLineL=None, showLine2L=None,
                      showUnits=True,
                      lineThkL=None, lineThk2L=None,
                      lineStyleL=None, lineStyle2L=None,
                      colorL=None, color2L=None,
                      labelL=None, label2L=None):
        """
        Add a scatter plot to the spread sheet.

        Use data from "data_sheetname" to create "plot_sheetname" with scatter plot.

        Assume index into columns is "1-based" such that column "A" is 1, "B" is 2, etc.

        NOTE: marker color is NOT supported in Excel from ODS format.
        Markers will appear, but will be some other default color.
        For this reason, ODSCharts uses Excel default color list unless the user
        overrides them with colorL or color2L.

        :param plot_sheetname: Name of the plot's tabbed window in Excel, LibreOffice or OpenOffice
        :type  plot_sheetname: str or unicode
        :param data_sheetname: Name of the data's tabbed window in Excel, LibreOffice or OpenOffice
        :type  data_sheetname: str or unicode

        :keyword title: Title of the plot as displayed in  Excel, LibreOffice or OpenOffice (default=='')
        :type  title: str or unicode
        :keyword xlabel: X Axis label as displayed in  Excel, LibreOffice or OpenOffice (default=='')
        :type  xlabel: str or unicode
        :keyword ylabel: Primary Y Axis label as displayed in  Excel, LibreOffice or OpenOffice (default=='')
        :type  ylabel: str or unicode
        :keyword y2label: Secondary Y Axis Label as displayed in  Excel, LibreOffice or OpenOffice  (default=='')
        :type  y2label: str or unicode

        :keyword int xcol: 1-based column index on data sheet of X Axis data (default==1)

        :keyword ycolL: List of 1-based column indeces on data sheet of Primary Y Axis data (default==None)
        :type  ycolL: None or int list
        :keyword ycol2L: List of 1-based column indeces on data sheet of Secondary Y Axis data  (default==None)
        :type  ycol2L: None or int list

        :keyword showMarkerL: List of boolean values indicating marker or no-marker for ycolL data (default==None)
        :type  showMarkerL: None or bool list
        :keyword showMarker2L: List of boolean values indicating marker or no-marker for ycol2L data (default==None)
        :type  showMarker2L: None or bool list
        
        :keyword showLineL: List of boolean values indicating line or no-line for ycolL data (default==None)
        :type  showLineL: None or bool list
        :keyword showLine2L: List of boolean values indicating line or no-line for ycol2L data (default==None)
        :type  showLine2L: None or bool list

        :keyword lineThkL: List of thickness values for primary y axis curves.
            If the list is not defined, then the thickness is set to "0.8mm".
            For example, lineThkL=[0.5, 1, 2] gives thin, about normal and thick lines.
            (Units=="mm")(default==None)
        :type  lineThkL: list of float
        :keyword lineThk2L: List of thickness values for secondary y axis curves
            If the list is not defined, then the thickness is set to "0.8mm".
            For example, lineThkL=[0.5, 1, 2] gives thin, about normal and thick lines.
            (Units=="mm")(default==None)
        :type  lineThk2L: list of float

        :keyword bool logx: Flag for X Axis type. True=log, False=linear (default==False)
        :keyword bool logy: Flag for Primary Y Axis type. True=log, False=linear (default==False)
        :keyword bool log2y:  Flag for Secondary Y Axis type. True=log, False=linear (default==False)
        :keyword bool showUnits: Flag to control units display. True=show units on axes, False=show labels only (default==True)
        :keyword bool excel_colors: Flag to control default color palette. True=use Excel colors, False=use custome colors.

        :keyword colorL: List of color values to use for ycolL curves. Format is "#FFCC33". Any curve not explicitly set will use Excel default colors (default==None)
        :type  colorL: None or str list
        :keyword color2L: List of color values to use for ycol2L curves. Format is "#FFCC33". Any curve not explicitly set will use Excel default colors (default==None)
        :type  color2L: None or str list
        :keyword labelL: Not yet implemented (default==None)
        :type  labelL: None or str list
        :keyword label2L: Not yet implemented (default==None)
        :type  label2L: None or str list

        :return: None
        :rtype: None
        """
        # Don't allow duplicate sheet names
        if (plot_sheetname in self.data_table_objD) or (plot_sheetname in self.plot_sheet_objD):
            raise  MySheetNameError('Duplicate sheet name submitted for new plot: "%s"'%plot_sheetname)

        # Data sheet must already exist in order to make a plot
        if (data_sheetname not in self.data_table_objD):
            raise  MySheetNameError('Data sheet for "%s" plot missing: "%s"'%(plot_sheetname, data_sheetname))

        num_chart = len(self.plot_sheet_objD) + 1

        # Add new chart object and tab page in Excel/LibreOffice/OpenOffice
        add_ObjectN( num_chart, self.metainf_manifest_xml_obj)

        plotSheetObj = PlotTableDesc( plot_sheetname, num_chart, self.content_xml_obj,
                                      excel_colors=excel_colors)
        plotSheetObj.document = self

        self.spreadsheet_obj.insert(TABLE_INSERT_POINT, plotSheetObj.xmlSheetObj)

        #obj_name = 'Object %i'%num_chart

        # ================= save PlotTableDesc object
        self.plot_sheet_objD[plot_sheetname] = plotSheetObj
        self.ordered_plotL.append( plot_sheetname )

        plotSheetObj.add_to_primary_y(data_sheetname, xcol, ycolL,
                                      showMarkerL=showMarkerL,
                                      showLineL=showLineL,
                                      colorL=colorL,
                                      lineThkL=lineThkL,
                                      lineStyleL=lineStyleL,
                                      labelL=labelL)

        plotSheetObj.add_to_secondary_y( data_sheetname, xcol, ycol2L,
                                         showMarker2L=showMarker2L,
                                         showLine2L=showLine2L,
                                         color2L=color2L,
                                         lineThk2L=lineThk2L,
                                         lineStyle2L=lineStyle2L,
                                         label2L=label2L)

        # Start making the chart object that goes onto the plot sheet
        #  Assign plot parameters to PlotTableDesc object
        plotSheetObj.plot_sheetname = plot_sheetname
        plotSheetObj.data_sheetname = data_sheetname
        plotSheetObj.title = title
        plotSheetObj.xlabel = xlabel
        plotSheetObj.ylabel = ylabel
        plotSheetObj.y2label = y2label
        plotSheetObj.ycolL = ycolL
        plotSheetObj.ycol2L = ycol2L

        plotSheetObj.logx = logx
        plotSheetObj.logy = logy
        plotSheetObj.log2y = log2y



    def add_sheet(self, data_sheetname, list_of_rows):
        """Create a new sheet in the spreadsheet with "data_sheetname" as its name.

           the list_of_rows will be placed at "A1" and should be:
            - row 1 is labels
            - row 2 is units
            - row 3 through N is float or string entries

            for example:
            list_of_rows = ::

                [['Altitude','Pressure','Temperature'],
                ['feet','psia','degR'],
                [0, 14.7, 518.7],
                [5000, 12.23, 500.8],
                [10000, 10.11, 483.0],
                [60000, 1.04, 390]]

        :param str data_sheetname:  Name of the data's tabbed window in Excel, LibreOffice or OpenOffice
        :type  data_sheetname: str or unicode
        :param list list_of_rows: the list_of_rows will be placed at "A1" and should be:
            - row 1 is labels
            - row 2 is units
            - row 3 through N is float or string entries
        :return: None
        :rtype: None

        """
        if (data_sheetname in self.data_table_objD) or (data_sheetname in self.plot_sheet_objD):
            raise  MySheetNameError('Duplicate sheet name submitted for new datasheet: "%s"'%data_sheetname)


        dataTableObj = DataTableDesc( data_sheetname, list_of_rows, self.content_xml_obj)
        self.spreadsheet_obj.insert(TABLE_INSERT_POINT, dataTableObj.xmlSheetObj)

        self.data_table_objD[data_sheetname] = dataTableObj
        self.ordered_dataL.append( dataTableObj )


    def save(self, filename='my_chart.ods', launch=False):
        """
        Saves SpreadSheet to an ods file readable by Microsoft Excel, LibreOffice or OpenOffice.

        If the launch flag is set, will launch Excel, LibreOffice or Openoffice using "os.startfile"
        or "open" or "xdg-open" depending on the platform.

        :keyword filename: Name of ods file to save (default=='my_chart.ods')
        :type  filename: str or unicode
        :keyword bool launch: If True, will launch Excel, LibreOffice or OpenOffice (default==False)
        :return: None
        :rtype: None
        """

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

            plotSheetObj = self.plot_sheet_objD[ plot_sheetname ]

            # create a new chart object
            if plotSheetObj.ycol2L:
                chart_obj = load_template_xml_from_ods('alt_chart_y2.ods', 'content.xml', subdir='Object 1')
            else:
                chart_obj = load_template_xml_from_ods('alt_chart.ods', 'content.xml', subdir='Object 1')

            build_chart_object_content( chart_obj, plotSheetObj )

            self.setAxisRanges( plot_sheetname )
            
            if len(plotSheetObj.set_of_line_styles) > 0:
                ObjectN_styles_xml_obj = deepcopy( self.template_ObjectN_styles_xml_obj )
                nsOD = chart_obj.rev_nsOD
                office_styles_obj = ObjectN_styles_xml_obj.root.find("office:styles", nsOD)
                gen_dash_elements_from_set_of_istyles( office_styles_obj, 
                                                       plotSheetObj.set_of_line_styles )
            else:
                ObjectN_styles_xml_obj = self.template_ObjectN_styles_xml_obj
                

            zipfile_insert( zipfileobj, 'Object %i/styles.xml'%(N+1,), ObjectN_styles_xml_obj.tostring())

            zipfile_insert( zipfileobj, 'Object %i/content.xml'%(N+1,), plotSheetObj.chart_obj.tostring())


        zipfile_insert( zipfileobj, 'content.xml', self.content_xml_obj.tostring())

        zipfile_insert( zipfileobj, 'styles.xml', self.styles_xml_obj.tostring())

        zipfileobj.close()

        if launch:
            #os.startfile( self.filename )
            self.launch_application()


if __name__ == '__main__':
    C = SpreadSheet()
    C.save( filename='performance')
