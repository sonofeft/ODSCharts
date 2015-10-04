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

N_LABELS = 0
N_UNITS = 1

def get_all_units_on_chart( plotSheetObj ):

    doc = plotSheetObj.document

    xUnitsL = []
    yUnitsL = []
    y2UnitsL = []

    # Get units for x and primary y
    for i,xcol in enumerate( plotSheetObj.xcolL ):
        data_sheetname = plotSheetObj.ycolDataSheetNameL[i]
        dataTableObj = doc.data_table_objD[ data_sheetname ]
        xcol_units = dataTableObj.list_of_rows[N_UNITS][xcol-1]
        if xcol_units:
            xUnitsL.append( xcol_units )

        ycol = plotSheetObj.ycolL[i]
        ycol_units = dataTableObj.list_of_rows[N_UNITS][ycol-1]
        if ycol_units:
            yUnitsL.append( ycol_units )

    # get units for x and secondary y
    for i,xcol in enumerate( plotSheetObj.xcol2L ):
        data_sheetname = plotSheetObj.ycol2_DataSheetNameL[i]
        dataTableObj = doc.data_table_objD[ data_sheetname ]
        xcol_units = dataTableObj.list_of_rows[N_UNITS][xcol-1]
        if xcol_units:
            xUnitsL.append( xcol_units )

        ycol2 = plotSheetObj.ycol2L[i]
        ycol2_units = dataTableObj.list_of_rows[N_UNITS][ycol2-1]
        if ycol2_units:
            y2UnitsL.append( ycol2_units )

    xUnitsL = sorted( list( set(xUnitsL) ) )
    yUnitsL = sorted( list( set(yUnitsL) ) )
    y2UnitsL = sorted( list( set(y2UnitsL) ) )

    if xUnitsL:
        xUnitsStr = ' (%s)'%', '.join(xUnitsL)
    else:
        xUnitsStr = ''

    if yUnitsL:
        yUnitsStr = ' (%s)'%', '.join(yUnitsL)
    else:
        yUnitsStr = ''

    if y2UnitsL:
        y2UnitsStr = ' (%s)'%', '.join(y2UnitsL)
    else:
        y2UnitsStr = ''

    return xUnitsStr, yUnitsStr, y2UnitsStr

def build_chart_object_content( chart_obj, plotSheetObj ):
    """When chart_obj is input, it still holds values from the original template"""


    chart = chart_obj.find('office:body/office:chart/chart:chart')
    plotSheetObj.chart_obj = chart_obj

    nsOD = chart_obj.rev_nsOD
    doc = plotSheetObj.document
    
    def set_new_tag(elem, shortname, value):
        """
        Adds shortname to qnames and sets attrib variable on elem
        shortname format is "chart:symbol-type"
        """

        doc.add_tag(NS(shortname, nsOD), plotSheetObj.chart_obj)
        elem.set(NS(shortname, nsOD), value)
        

    title = chart.find('chart:title/text:p', nsOD)
    #print( 'title.text = ',title.text)
    title.text = plotSheetObj.title

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

    xUnitsStr, yUnitsStr, y2UnitsStr = get_all_units_on_chart( plotSheetObj )


    if plotSheetObj.showUnits:
        xtitle.text = plotSheetObj.xlabel + xUnitsStr
    else:
        xtitle.text = plotSheetObj.xlabel

    if plotSheetObj.showUnits and plotSheetObj.ycolL:
        ytitle.text = plotSheetObj.ylabel + yUnitsStr
    else:
        ytitle.text = plotSheetObj.ylabel

    # Look for secondary y axis
    if plotSheetObj.ycol2L:
        y2title = y2axis.find('chart:title/text:p', nsOD)
        if plotSheetObj.showUnits:
            y2title.text = plotSheetObj.y2label + y2UnitsStr
        else:
            y2title.text = plotSheetObj.y2label


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
    ycol = plotSheetObj.ycolL[0]
    col_letter = get_col_letters_from_number( ycol )
    sht_name = plotSheetObj.ycolDataSheetNameL[0]
    dataTableObj = doc.data_table_objD[ sht_name ]
    val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,dataTableObj.nrows)
    #print( 'val_cell_range =',val_cell_range)

    lab_cell = '%s.$%s$1'%(sht_name, col_letter)
    y_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
    y_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)

    chart_domain = y_series.find('chart:domain', nsOD)
    chart_data_point = y_series.find('chart:data-point', nsOD)
    #print( chart_domain.items())
    #print( chart_data_point.items())

    xcol = plotSheetObj.xcolL[0]
    xcol_letter = get_col_letters_from_number( xcol )
    xval_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,xcol_letter,xcol_letter,dataTableObj.nrows)
    chart_domain.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}cell-range-address',xval_cell_range)

    chart_data_point.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}repeated','%i'%(dataTableObj.nrows-2,))

    # Look for logarithmic X Axis
    if plotSheetObj.logx:
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



    if plotSheetObj.ycolL and len(plotSheetObj.ycolL) >=1:
        auto_styles = chart_obj.find('office:automatic-styles')
        autostyleL = auto_styles.findall('style:style', nsOD)
        ref_series_style = None

        istyle_loc = 0
        for iloc,astyle in enumerate(autostyleL):
            if astyle.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name') == 'G0S0':
                ref_series_style = astyle
                istyle_loc = iloc + 2

                # graphic-properties
                elem = ref_series_style.find("style:graphic-properties", nsOD)
                c = plotSheetObj.colorL[0]
                if not c is None:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c)
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)


                w = plotSheetObj.lineThkL[0]
                if not w is None:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-width", w)

                if not plotSheetObj.showLineL[0]:
                    #elem.attrib.clear()
                    elem.set(NS("draw:fill", nsOD), "none")
                    elem.set(NS("draw:stroke", nsOD), "none")
                else:
                    elem.set(NS("draw:stroke", nsOD), "solid")


                # .............. Failed Experiment ...................
                #doc.add_tag(NS("draw:symbol-color", nsOD), plotSheetObj.chart_obj)
                #elem.set(NS("draw:symbol-color", nsOD), "#999999")
                #doc.add_tag(NS("draw:fill-color", nsOD), plotSheetObj.chart_obj)
                #elem.set(NS("draw:fill-color", nsOD), "#999999")
                #elem.set(NS("draw:fill", nsOD), "solid")

                # chart-properties
                elem = ref_series_style.find("style:chart-properties", nsOD)
                if plotSheetObj.showMarkerL[0]:  # showMarkerL is guaranteed to have same dimension as ycolL
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
                else:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")

                # elem is  chart-properties
                if not plotSheetObj.showLineL[0]:
                    set_new_tag(elem, "chart:symbol-type", "named-symbol")
                    set_new_tag(elem, "chart:symbol-name", plotSheetObj.markerTypeL[0])
                    hw = plotSheetObj.markerHtWdL[0]
                    set_new_tag(elem, "chart:symbol-width", hw)
                    set_new_tag(elem, "chart:symbol-height", hw)


        # Look for logarithmic scale on primary y
        if plotSheetObj.logy:
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
        for ycol in plotSheetObj.ycolL[1:]:
            new_style = deepcopy( ref_series_style )
            new_style.set('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'G0S%i'%nG0S )

            elem = new_style.find("style:graphic-properties", nsOD)
            c = plotSheetObj.colorL[nG0S]
            if not c is None:
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c )
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)

            w = plotSheetObj.lineThkL[nG0S]
            if not w is None:
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-width", w)

            if not plotSheetObj.showLineL[nG0S]:
                #elem.attrib.clear()
                elem.set(NS("draw:fill", nsOD), "none")
                elem.set(NS("draw:stroke", nsOD), "none")
            else:
                elem.set(NS("draw:stroke", nsOD), "solid")

            elem = new_style.find("style:chart-properties", nsOD)
            if plotSheetObj.showMarkerL[nG0S]:  # showMarkerL is guaranteed to have same dimension as ycolL
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
            else:
                elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")

            # elem is  chart-properties
            if not plotSheetObj.showLineL[nG0S]:
                set_new_tag(elem, "chart:symbol-type", "named-symbol")
                set_new_tag(elem, "chart:symbol-name", plotSheetObj.markerTypeL[nG0S])
                hw = plotSheetObj.markerHtWdL[nG0S]
                set_new_tag(elem, "chart:symbol-width", hw)
                set_new_tag(elem, "chart:symbol-height", hw)

            auto_styles.insert(istyle_loc, new_style )
            istyle_loc += 1

            new_chart_series = deepcopy( y_series )
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}style-name','G0S%i'%nG0S)

            col_letter = get_col_letters_from_number( ycol )
            sht_name = plotSheetObj.ycolDataSheetNameL[nG0S]
            dataTableObj = doc.data_table_objD[ sht_name ]
            val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,dataTableObj.nrows)
            #print( 'val_cell_range =',val_cell_range)

            lab_cell = '%s.$%s$1'%(sht_name, col_letter)
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)


            xcol = plotSheetObj.xcolL[nG0S]
            xcol_letter = get_col_letters_from_number( xcol )
            xval_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,xcol_letter,xcol_letter,dataTableObj.nrows)
            new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}cell-range-address',xval_cell_range)


            ipos = len(plot_area)-1
            plot_area.insert(ipos, new_chart_series )

            nG0S += 1

    # =============== Secondary Y Axis =========================
    # Look for secondary y axis
    if (plotSheetObj.ycol2L is not None) and (len(plotSheetObj.ycol2L)>0):
        series_label_cell_address = y2_series.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address')
        #print( 'Y2 series_label_cell_address =',series_label_cell_address)

        series_value_cell_range = y2_series.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address')
        #print( 'Y2 series_value_cell_range =',series_value_cell_range)
        ycol = plotSheetObj.ycol2L[0]
        col_letter = get_col_letters_from_number( ycol )
        sht_name = plotSheetObj.ycol2_DataSheetNameL[0]
        dataTableObj = doc.data_table_objD[ sht_name ]
        val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,dataTableObj.nrows)
        #print( 'Y2 val_cell_range =',val_cell_range)

        lab_cell = '%s.$%s$1'%(sht_name, col_letter)
        y2_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
        y2_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)

        chart_domain = y2_series.find('chart:domain', nsOD)
        chart_data_point = y2_series.find('chart:data-point', nsOD)
        #print( chart_domain.items())
        #print( chart_data_point.items())

        xcol = plotSheetObj.xcol2L[0]
        xcol_letter = get_col_letters_from_number( xcol )
        xval_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,xcol_letter,xcol_letter,dataTableObj.nrows)
        chart_domain.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}cell-range-address',xval_cell_range)

        chart_data_point.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}repeated','%i'%(dataTableObj.nrows-2,))


        # Look for logarithmic scale on primary y
        if plotSheetObj.log2y:
            print( 'Got log on secondary y' )

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
            if 0:#plotSheetObj.ycolL and (not plotSheetObj.logy):
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
        if len(plotSheetObj.ycol2L) >= 1:
            auto_styles = chart_obj.find('office:automatic-styles')
            autostyleL = auto_styles.findall('style:style', nsOD)
            ref_series_style = None

            istyle_loc = 0
            for iloc,astyle in enumerate(autostyleL):
                if astyle.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name') == 'G1S0':
                    ref_series_style = astyle
                    istyle_loc = iloc + 2

                    elem = ref_series_style.find("style:graphic-properties", nsOD)
                    c = plotSheetObj.color2L[0]
                    if not c is None:
                        elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c)
                        elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)


                    w = plotSheetObj.lineThk2L[0]
                    if not w is None:
                        elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-width", w)


                    if not plotSheetObj.showLine2L[0]:
                        elem.set(NS("draw:fill", nsOD), "none")
                        elem.set(NS("draw:stroke", nsOD), "none")
                    else:
                        elem.set(NS("draw:stroke", nsOD), "solid")

                    elem = ref_series_style.find("style:chart-properties", nsOD)
                    if plotSheetObj.showMarkerL[0]:  # showMarkerL is guaranteed to have same dimension as ycol2L
                        elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
                    else:
                        elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")

                    # elem is  chart-properties
                    if not plotSheetObj.showLine2L[0]:
                        set_new_tag(elem, "chart:symbol-type", "named-symbol")
                        set_new_tag(elem, "chart:symbol-name", plotSheetObj.markerType2L[0])
                        hw = plotSheetObj.markerHtWd2L[0]
                        set_new_tag(elem, "chart:symbol-width", hw)
                        set_new_tag(elem, "chart:symbol-height", hw)



            plotSheetObj.ref_series_style2 = ref_series_style

            #print( ref_series_style.items())

            nG1S = 1
            for ycol in plotSheetObj.ycol2L[1:]:
                new_style = deepcopy( ref_series_style )
                new_style.set('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'G1S%i'%nG1S )

                elem = new_style.find("style:graphic-properties", nsOD)
                c = plotSheetObj.color2L[nG1S]
                if not c is None:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color", c )
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color", c)


                w = plotSheetObj.lineThk2L[nG1S]
                if not w is None:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-width", w)

                if not plotSheetObj.showLine2L[nG1S]:
                    elem.set(NS("draw:fill", nsOD), "none")
                    elem.set(NS("draw:stroke", nsOD), "none")
                else:
                    elem.set(NS("draw:stroke", nsOD), "solid")

                elem = new_style.find("style:chart-properties", nsOD)
                if plotSheetObj.showMarker2L[nG1S]:  # showMarkerL is guaranteed to have same dimension as ycol2L
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "automatic")
                else:
                    elem.set("{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}symbol-type", "none")

                # elem is  chart-properties
                if not plotSheetObj.showLine2L[nG1S]:
                    set_new_tag(elem, "chart:symbol-type", "named-symbol")
                    set_new_tag(elem, "chart:symbol-name", plotSheetObj.markerType2L[nG1S])
                    hw = plotSheetObj.markerHtWd2L[nG1S]
                    set_new_tag(elem, "chart:symbol-width", hw)
                    set_new_tag(elem, "chart:symbol-height", hw)


                auto_styles.insert(istyle_loc, new_style )
                istyle_loc += 1

                new_chart_series = deepcopy( y2_series )
                new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}style-name','G1S%i'%nG1S)

                col_letter = get_col_letters_from_number( ycol )
                sht_name = plotSheetObj.ycol2_DataSheetNameL[nG1S]
                dataTableObj = doc.data_table_objD[ sht_name ]
                val_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,col_letter,col_letter,dataTableObj.nrows)
                #print( 'val_cell_range =',val_cell_range)

                lab_cell = '%s.$%s$1'%(sht_name, col_letter)
                new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}label-cell-address',lab_cell)
                new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}values-cell-range-address',val_cell_range)


                xcol = plotSheetObj.xcol2L[nG1S]
                xcol_letter = get_col_letters_from_number( xcol )
                xval_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,xcol_letter,xcol_letter,dataTableObj.nrows)
                new_chart_series.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}cell-range-address',xval_cell_range)


                ipos = len(plot_area)-1
                plot_area.insert(ipos, new_chart_series )

                nG1S += 1

    plotSheetObj.istyle_loc = istyle_loc
    plotSheetObj.y_series = y_series
    plotSheetObj.chart_obj = chart_obj


