"""
Start with an XML template of the "Object 1/content.xml" file and modify it
for a new chart
"""
from copy import deepcopy
from collections import OrderedDict

import sys
if sys.version_info < (3,):
    import ElementTree_27OD as ET
else:
    import ElementTree_34OD as ET


def get_col_letters_from_number(num):
    letters = ''
    while num:
        mod = (num - 1) % 26
        letters += chr(mod + 65)
        num = (num - 1) // 26
    return ''.join(reversed(letters))


# These colors generated with color_utils.py
#   Abandoned when I discovered that Excel marker colors are not supported.
#COLOR_LIST = ['#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#00FFFF', '#FF00FF', 
#              '#FF7F00', '#00FF7F', '#7F00FF', '#7FFF00', '#007FFF', '#FF007F', 
#              '#B23535', '#35B235', '#3535B2', '#B2B235', '#35B2B2', '#B235B2', 
#              '#B27435', '#35B274', '#7435B2', '#74B235', '#3574B2', '#B23574', 
#              '#663D3D', '#3D663D', '#3D3D66', '#66663D', '#3D6666', '#663D66', 
#              '#66513D', '#3D6651', '#513D66', '#51663D', '#3D5166', '#663D51']

# These default colors are taken from Excel
#  Since Excel does not support marker colors in ODS format, try to just use
#  the Excel default colors so the marker color will match the line color
COLOR_LIST = ['#325885', '#873331', '#6b833a', '#574271', '#2f788c', '#b0672b', 
              '#3c679a', '#9d3d3a', '#7d9844', '#664e83', '#388ca2', '#cb7833', 
              '#4373ab', '#ae4441', '#8ba94d', '#725792', '#3f9bb4', '#e1853a', 
              '#4a7ebb', '#be4b48', '#98b954', '#7d60a0', '#46aac5', '#f69240', 
              '#809cc7', '#ca817f', '#aec685', '#9c8bb3', '#7ebace', '#f6a97b', 
              '#a5b6d3', '#d5a5a4', '#c1d2a7', '#b6abc5', '#a4cad8', '#f6bea2']

def get_color( index, inp_color=None):
    """Get colors for plot lines"""
    list_color = COLOR_LIST[ index % len(COLOR_LIST) ]
    
    s = '%s'%inp_color
    if not s.startswith('#'):
        return list_color
    if not len(s)==7:
        return list_color
    return inp_color

def build_chart_object_content( chart_obj, plot_desc, table_desc ):
    """When chart_obj is input, it still holds values from the original template"""
    

    def node_print( n ):
        for ch in n:
            print( ch)
        print( '='*55)

    chart = chart_obj.find('office:body/office:chart/chart:chart')

    title = chart.find('chart:title/text:p', chart_obj.rev_nsOD)
    #print( title.text)
    title.text = plot_desc.title

    plot_area = chart.find('chart:plot-area', chart_obj.rev_nsOD)

    axisL = plot_area.findall('chart:axis', chart_obj.rev_nsOD)
    xaxis = None
    yaxis = None
    for axis in axisL:
        dim_name = axis.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}dimension')
        if dim_name == 'x':
            xaxis = axis
        if dim_name == 'y':
            yaxis = axis

    xtitle = xaxis.find('chart:title/text:p', chart_obj.rev_nsOD)
    ytitle = yaxis.find('chart:title/text:p', chart_obj.rev_nsOD)
    xtitle.text = plot_desc.xlabel
    ytitle.text = plot_desc.ylabel
    
    
    chart_series = plot_area.find('chart:series', chart_obj.rev_nsOD)
    node_print( chart_series )

    series_label_cell_address = chart_series.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address')
    print( 'series_label_cell_address =',series_label_cell_address)

    series_value_cell_range = chart_series.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address')
    print( 'series_value_cell_range =',series_value_cell_range)
    ycol = plot_desc.ycolL[0]
    col_letter = get_col_letters_from_number( ycol )
    sht_name = plot_desc.data_sheetname
    val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,table_desc.nrows)
    print( 'val_cell_range =',val_cell_range)
    
    lab_cell = '%s.$%s$1'%(sht_name, col_letter)
    chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
    chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)

    chart_domain = chart_series.find('chart:domain', chart_obj.rev_nsOD)
    chart_data_point = chart_series.find('chart:data-point', chart_obj.rev_nsOD)
    #print( chart_domain.items())
    #print( chart_data_point.items())

    xcol = plot_desc.xcol
    xcol_letter = get_col_letters_from_number( xcol )
    xval_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,xcol_letter,xcol_letter,table_desc.nrows)
    chart_domain.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}cell-range-address',xval_cell_range)
    
    chart_data_point.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}repeated','%i'%(table_desc.nrows-2,))
    
    if plot_desc.ycolL and len(plot_desc.ycolL) > 1:
        auto_styles = chart_obj.find('office:automatic-styles')
        autostyleL = auto_styles.findall('style:style', chart_obj.rev_nsOD)
        ref_series_style = None
        
        istyle_loc = 0
        for iloc,astyle in enumerate(autostyleL):
            if astyle.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name') == 'G0S0':
                ref_series_style = astyle
                istyle_loc = iloc + 2
                
                elem = ref_series_style.find("style:graphic-properties", chart_obj.rev_nsOD)
                c = plot_desc.colorL[0]  # get_color( 0, inp_color=plot_desc.colorL[0] )
                if not c is None:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c)                
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)                
                
                elem = ref_series_style.find("style:chart-properties", chart_obj.rev_nsOD)
                if plot_desc.showMarkerL[0]:  # showMarkerL is guaranteed to have same dimension as ycolL
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
                else:    
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")
                    
                
        #print( ref_series_style.items())
        
        nG0S = 1
        for ycol in plot_desc.ycolL[1:]:
            new_style = deepcopy( ref_series_style )
            new_style.set('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'G0S%i'%nG0S )

            elem = new_style.find("style:graphic-properties", chart_obj.rev_nsOD)
            c = plot_desc.colorL[nG0S]  #get_color( nG0S, inp_color=plot_desc.colorL[nG0S] )
            if not c is None:
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c )
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)                


            elem = new_style.find("style:chart-properties", chart_obj.rev_nsOD)
            if plot_desc.showMarkerL[nG0S]:  # showMarkerL is guaranteed to have same dimension as ycolL
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
            else:
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")
            
            
            auto_styles.insert(istyle_loc, new_style )
            istyle_loc += 1
            
            new_chart_series = deepcopy( chart_series )
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}style-name','G0S%i'%nG0S)
            
            col_letter = get_col_letters_from_number( ycol )
            sht_name = plot_desc.data_sheetname
            val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,table_desc.nrows)
            print( 'val_cell_range =',val_cell_range)
    
            lab_cell = '%s.$%s$1'%(sht_name, col_letter)
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)

            ipos = len(plot_area)-1
            plot_area.insert(ipos, new_chart_series )

            nG0S += 1
    
    plot_desc.chart_obj = chart_obj
    
    
