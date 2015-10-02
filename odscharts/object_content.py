"""
Start with an XML template of the "Object 1/content.xml" file and modify it
for a new chart
"""
from copy import deepcopy
from collections import OrderedDict
from find_obj import find_elem_w_attrib, elem_set, NS_attrib, NS

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
    

    chart = chart_obj.find('office:body/office:chart/chart:chart')
    
    nsOD = chart_obj.rev_nsOD

    title = chart.find('chart:title/text:p', nsOD)
    #print( 'title.text = ',title.text)
    title.text = plot_desc.title

    plot_area = chart.find('chart:plot-area', nsOD)
    #print( 'plot_area =',plot_area )

    axisL = plot_area.findall('chart:axis', nsOD)
    #print( 'axisL =',axisL )
    xaxis = None
    yaxis = None
    x2axis = None
    y2axis = None
    for axis in axisL:
        dim_xy = axis.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}dimension')
        dim_name = axis.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}name')
        #print( 'dim_xy =',dim_xy )
        if dim_xy == 'x' and dim_name == 'primary-x':
            xaxis = axis
        if dim_xy == 'y' and dim_name == 'primary-y':
            yaxis = axis

        # ODDLY... Excel uses capital X in "secondary-X" label
        #if dim_xy == 'x' and dim_name.startswith('secondary'):
        #    x2axis = axis
        if dim_xy == 'y' and dim_name.startswith('secondary'):
            y2axis = axis

    xtitle = xaxis.find('chart:title/text:p', nsOD)
    ytitle = yaxis.find('chart:title/text:p', nsOD)
    
    xtitle.text = plot_desc.xlabel
    if plot_desc.showUnits:
        xcol_units = table_desc.list_of_rows[1][plot_desc.xcol-1]
        xtitle.text = plot_desc.xlabel + ' (%s)'%xcol_units
    
    ytitle.text = plot_desc.ylabel
    if plot_desc.showUnits and plot_desc.ycolL:
        unitSet = set( [table_desc.list_of_rows[1][ycol-1] for ycol in plot_desc.ycolL] )
        unitL = sorted( list(unitSet) )
        ytitle.text = plot_desc.ylabel + ' (%s)'%( ', '.join(unitL), )
    
    # Look for secondary y axis 
    if plot_desc.ycol2L:
        #x2title = x2axis.find('chart:title/text:p', nsOD)
        y2title = y2axis.find('chart:title/text:p', nsOD)
        #x2title.text = plot_desc.xlabel # NOTE: I only allow one X axis label
        y2title.text = plot_desc.y2label
        if plot_desc.showUnits:
            unitSet = set( [table_desc.list_of_rows[1][ycol2-1] for ycol2 in plot_desc.ycol2L] )
            unitL = sorted( list(unitSet) )
            y2title.text = plot_desc.y2label + ' (%s)'%( ', '.join(unitL), )
    
    
    chart_seriesL = plot_area.findall('chart:series', nsOD)
    y_series = None
    y2_series = None
    for c_series in chart_seriesL:
        axis_name = c_series.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}attached-axis')
        if axis_name == 'primary-y':
            y_series = c_series
        if axis_name == 'secondary-y':
            y2_series = c_series
    #print( 'y2_series =',y2_series )

    # =============== Primary Y Axis =========================
    series_label_cell_address = y_series.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address')
    #print( 'series_label_cell_address =',series_label_cell_address)

    series_value_cell_range = y_series.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address')
    #print( 'series_value_cell_range =',series_value_cell_range)
    ycol = plot_desc.ycolL[0]
    col_letter = get_col_letters_from_number( ycol )
    sht_name = plot_desc.data_sheetname
    val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,table_desc.nrows)
    #print( 'val_cell_range =',val_cell_range)
    
    lab_cell = '%s.$%s$1'%(sht_name, col_letter)
    y_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
    y_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)

    chart_domain = y_series.find('chart:domain', nsOD)
    chart_data_point = y_series.find('chart:data-point', nsOD)
    #print( chart_domain.items())
    #print( chart_data_point.items())

    xcol = plot_desc.xcol
    xcol_letter = get_col_letters_from_number( xcol )
    xval_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,xcol_letter,xcol_letter,table_desc.nrows)
    chart_domain.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}cell-range-address',xval_cell_range)
    
    chart_data_point.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}repeated','%i'%(table_desc.nrows-2,))
    
    # Look for logarithmic X Axis
    if plot_desc.logx:
        auto_styles = chart_obj.find('office:automatic-styles')
        
        # Find "Axs0"
        logx_style,ipos_logx_style = find_elem_w_attrib('style:style', auto_styles, nsOD, 
                                   attrib={'style:name':'Axs0'}, nth_match=0)
        #print( 'FOUND:  ipos_logx_style = ', ipos_logx_style )
        
        chart_prop = logx_style.find( NS('style:chart-properties', nsOD) )
        chart_prop.attrib = NS_attrib({ "chart:display-label":"true", "chart:link-data-style-to-source":"true",
            "chart:logarithmic":"true", "chart:tick-marks-major-inner":"false", "chart:tick-marks-major-outer":"true",
            "chart:tick-marks-minor-inner":"true", "chart:tick-marks-minor-outer":"true", "chart:visible":"true"}, nsOD)

        # Find "GMa0"
        logx_style,ipos_logx_style = find_elem_w_attrib('style:style', auto_styles, nsOD, 
                                   attrib={'style:name':'GMa0'}, nth_match=0)
        #print( 'FOUND:  ipos_logx_style = ', ipos_logx_style )
        stroke_style = logx_style.find( NS('style:graphic-properties', nsOD) )
        stroke_style.set( NS('draw:stroke', nsOD), 'solid' )
        del stroke_style.attrib[ NS('draw:stroke-dash', nsOD) ]
        
        # Find "G0S0"
        logx_style,ipos_logx_style = find_elem_w_attrib('style:style', auto_styles, nsOD, 
                                   attrib={'style:name':'G0S0'}, nth_match=0)
        #print( 'FOUND:  ipos_logx_style = ', ipos_logx_style )
        new_elem_1 = ET.SubElement(auto_styles,"{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style", 
            attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family', 'chart'), 
            ('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'GMi0')]))
        
        new_elem_2 = ET.SubElement(new_elem_1,"{urn:oasis:names:tc:opendocument:xmlns:style:1.0}graphic-properties", 
            attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill', 'none'), 
            ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke', 'dash'), 
            ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash', 'a4'), 
            ('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-width', '0.01042in'), 
            ('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color', '#000000'), 
            ('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-opacity', '100%'), 
            ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-linejoin', 'round')]))
            
        # Add minor grid to xaxis
        new_elem_3 = ET.SubElement(xaxis,"{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}grid", 
            attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}class', 'minor'), 
            ('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}style-name', 'GMi0')]))
        
        
    
    if plot_desc.ycolL and len(plot_desc.ycolL) >=1:
        auto_styles = chart_obj.find('office:automatic-styles')
        autostyleL = auto_styles.findall('style:style', nsOD)
        ref_series_style = None
        
        istyle_loc = 0
        for iloc,astyle in enumerate(autostyleL):
            if astyle.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name') == 'G0S0':
                ref_series_style = astyle
                istyle_loc = iloc + 2
                
                elem = ref_series_style.find("style:graphic-properties", nsOD)
                c = plot_desc.get_next_color()
                if not c is None:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c)                
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)                
                
                elem = ref_series_style.find("style:chart-properties", nsOD)
                if plot_desc.showMarkerL[0]:  # showMarkerL is guaranteed to have same dimension as ycolL
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
                else:    
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")
        
        # Look for logarithmic scale on primary y
        if plot_desc.logy:
            #print( 'Got log y' )
        
            # Find "Axs1"
            logy_style,ipos_logy_style = find_elem_w_attrib('style:style', auto_styles, nsOD, 
                                       attrib={'style:name':'Axs1'}, nth_match=0)
            #print( 'FOUND:  ipos_logy_style = ', ipos_logy_style )
            chart_prop = logy_style.find( NS('style:chart-properties', nsOD) )
            
            chart_prop.set( NS("chart:logarithmic", nsOD), "true" )
            chart_prop.set( NS("chart:tick-marks-minor-inner", nsOD), "true" )
            chart_prop.set( NS("chart:tick-marks-minor-outer", nsOD), "true" )

        
            # Find "GMa1"
            logy_style,ipos_logy_style = find_elem_w_attrib('style:style', auto_styles, nsOD, 
                                       attrib={'style:name':'GMa1'}, nth_match=0)
            #print( 'FOUND:  ipos_logy_style = ', ipos_logy_style )
            logy_style.set( NS("draw:stroke", nsOD), "dash" )
        
            # Find "G0S1"
            logy_style,ipos_logy_style = find_elem_w_attrib('style:style', auto_styles, nsOD, 
                                       attrib={'style:name':'G0S1'}, nth_match=0)
            #print( 'FOUND:  ipos_logy_style = ', ipos_logy_style )
            
            new_elem_1 = ET.SubElement(auto_styles,"{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style", 
                attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family', 'chart'), 
                ('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'GMi1')]))
            
            new_elem_2 = ET.SubElement(new_elem_1,"{urn:oasis:names:tc:opendocument:xmlns:style:1.0}graphic-properties", 
                attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill', 'none'), 
                ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke', 'dash'), 
                ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash', 'a4'), 
                ('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-width', '0.01042in'), 
                ('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color', '#000000'), 
                ('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-opacity', '100%'), 
                ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-linejoin', 'round')]))
            
            # Add minor grid to yaxis
            new_elem_3 = ET.SubElement(yaxis,"{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}grid", 
                attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}class', 'minor'), 
                ('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}style-name', 'GMi1')]))

                
        #print( ref_series_style.items())
        
        nG0S = 1
        for ycol in plot_desc.ycolL[1:]:
            new_style = deepcopy( ref_series_style )
            new_style.set('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'G0S%i'%nG0S )

            elem = new_style.find("style:graphic-properties", nsOD)
            c = plot_desc.get_next_color()
            if not c is None:
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c )
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)                


            elem = new_style.find("style:chart-properties", nsOD)
            if plot_desc.showMarkerL[nG0S]:  # showMarkerL is guaranteed to have same dimension as ycolL
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
            else:
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")
            
            
            auto_styles.insert(istyle_loc, new_style )
            istyle_loc += 1
            
            new_chart_series = deepcopy( y_series )
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}style-name','G0S%i'%nG0S)
            
            col_letter = get_col_letters_from_number( ycol )
            sht_name = plot_desc.data_sheetname
            val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,table_desc.nrows)
            #print( 'val_cell_range =',val_cell_range)
    
            lab_cell = '%s.$%s$1'%(sht_name, col_letter)
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)

            ipos = len(plot_area)-1
            plot_area.insert(ipos, new_chart_series )

            nG0S += 1

    # =============== Secondary Y Axis =========================
    # Look for secondary y axis 
    if plot_desc.ycol2L:
        series_label_cell_address = y2_series.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address')
        #print( 'Y2 series_label_cell_address =',series_label_cell_address)

        series_value_cell_range = y2_series.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address')
        #print( 'Y2 series_value_cell_range =',series_value_cell_range)
        ycol = plot_desc.ycol2L[0]
        col_letter = get_col_letters_from_number( ycol )
        sht_name = plot_desc.data_sheetname
        val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,table_desc.nrows)
        #print( 'Y2 val_cell_range =',val_cell_range)
        
        lab_cell = '%s.$%s$1'%(sht_name, col_letter)
        y2_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
        y2_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)

        chart_domain = y2_series.find('chart:domain', nsOD)
        chart_data_point = y2_series.find('chart:data-point', nsOD)
        #print( chart_domain.items())
        #print( chart_data_point.items())

        xcol = plot_desc.xcol
        xcol_letter = get_col_letters_from_number( xcol )
        xval_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,xcol_letter,xcol_letter,table_desc.nrows)
        chart_domain.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}cell-range-address',xval_cell_range)
        
        chart_data_point.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}repeated','%i'%(table_desc.nrows-2,))

        
        # Look for logarithmic scale on primary y
        if plot_desc.log2y:
            #print( 'Got log on secondary y' )
        
            # Find "Axs2"
            logy_style,ipos_logy_style = find_elem_w_attrib('style:style', auto_styles, nsOD, 
                                       attrib={'style:name':'Axs2'}, nth_match=0)
            #print( 'FOUND:  ipos_logy_style = ', ipos_logy_style )
            chart_prop = logy_style.find( NS('style:chart-properties', nsOD) )
            
            chart_prop.set( NS("chart:logarithmic", nsOD), "true" )
            chart_prop.set( NS("chart:tick-marks-minor-inner", nsOD), "true" )
            chart_prop.set( NS("chart:tick-marks-minor-outer", nsOD), "true" )
        
        
            # Make "GMi2"
            new_elem_1 = ET.SubElement(auto_styles,"{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style", 
                attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family', 'chart'), 
                ('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'GMi2')]))
            
            # Looks better with only one set of y grids... Seems obvious now
            if 0:#plot_desc.ycolL and (not plot_desc.logy):
                new_elem_2 = ET.SubElement(new_elem_1,"{urn:oasis:names:tc:opendocument:xmlns:style:1.0}graphic-properties", 
                    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill', 'none'), 
                    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke', 'dash'), 
                    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash', 'a4'), 
                    ('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-width', '0.01042in'), 
                    ('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color', '#000000'), 
                    ('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-opacity', '100%'), 
                    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-linejoin', 'round')]))
                
                # Add minor grid to yaxis
                new_elem_3 = ET.SubElement(y2axis,"{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}grid", 
                    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}class', 'minor'), 
                    ('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}style-name', 'GMi2')]))



        #  If more than one curve on secondary y...
        if len(plot_desc.ycol2L) > 1:
            auto_styles = chart_obj.find('office:automatic-styles')
            autostyleL = auto_styles.findall('style:style', nsOD)
            ref_series_style = None
            
            istyle_loc = 0
            for iloc,astyle in enumerate(autostyleL):
                if astyle.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name') == 'G1S0':
                    ref_series_style = astyle
                    istyle_loc = iloc + 2
                    
                    elem = ref_series_style.find("style:graphic-properties", nsOD)
                    c = plot_desc.get_next_color()
                    if not c is None:
                        elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c)                
                        elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)                
                    
                    elem = ref_series_style.find("style:chart-properties", nsOD)
                    if plot_desc.showMarkerL[0]:  # showMarkerL is guaranteed to have same dimension as ycol2L
                        elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
                    else:    
                        elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")
                        
                    
            #print( ref_series_style.items())
            
            nG1S = 1
            for ycol in plot_desc.ycol2L[1:]:
                new_style = deepcopy( ref_series_style )
                new_style.set('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'G1S%i'%nG1S )

                elem = new_style.find("style:graphic-properties", nsOD)
                c = plot_desc.get_next_color()
                if not c is None:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c )
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)                


                elem = new_style.find("style:chart-properties", nsOD)
                if plot_desc.showMarker2L[nG1S]:  # showMarkerL is guaranteed to have same dimension as ycol2L
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
                else:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")
                
                
                auto_styles.insert(istyle_loc, new_style )
                istyle_loc += 1
                
                new_chart_series = deepcopy( y2_series )
                new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}style-name','G1S%i'%nG1S)
                
                col_letter = get_col_letters_from_number( ycol )
                sht_name = plot_desc.data_sheetname
                val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,table_desc.nrows)
                #print( 'val_cell_range =',val_cell_range)
        
                lab_cell = '%s.$%s$1'%(sht_name, col_letter)
                new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
                new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)

                ipos = len(plot_area)-1
                plot_area.insert(ipos, new_chart_series )

                nG1S += 1
        
    
    plot_desc.chart_obj = chart_obj
    
    
