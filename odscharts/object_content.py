"""
Start with an XML template of the "Object 1/content.xml" file and modify it
for a new chart
"""
from copy import deepcopy

from lxml import etree as ET

def get_col_letters_from_number(num):
    letters = ''
    while num:
        mod = (num - 1) % 26
        letters += chr(mod + 65)
        num = (num - 1) // 26
    return ''.join(reversed(letters))
    

def build_chart_object_content( chart_obj, plot_desc, table_desc ):
    """When chart_obj is input, it still holds values from the original template"""
    
    root = chart_obj.getroot()
    

    def node_print( n ):
        for ch in n:
            print( ch)
        print( '='*55)

    chart = root.find('office:body/office:chart/chart:chart', root.nsmap)

    title = chart.find('chart:title/text:p', root.nsmap)
    #print( title.text)
    title.text = plot_desc.title

    plot_area = chart.find('chart:plot-area', root.nsmap)

    axisL = plot_area.findall('chart:axis', root.nsmap)
    xaxis = None
    yaxis = None
    for axis in axisL:
        dim_name = axis.get('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}dimension')
        if dim_name == 'x':
            xaxis = axis
        if dim_name == 'y':
            yaxis = axis

    xtitle = xaxis.find('chart:title/text:p', root.nsmap)
    ytitle = yaxis.find('chart:title/text:p', root.nsmap)
    xtitle.text = plot_desc.xlabel
    ytitle.text = plot_desc.ylabel
    
    
    chart_series = plot_area.find('chart:series', root.nsmap)
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

    chart_domain = chart_series.find('chart:domain', root.nsmap)
    chart_data_point = chart_series.find('chart:data-point', root.nsmap)
    #print( chart_domain.items())
    #print( chart_data_point.items())

    xcol = plot_desc.xcol
    xcol_letter = get_col_letters_from_number( xcol )
    xval_cell_range = '%s.$%s$3:.$%s$%i'%(sht_name,xcol_letter,xcol_letter,table_desc.nrows)
    chart_domain.set('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}cell-range-address',xval_cell_range)
    
    chart_data_point.set('{urn:oasis:names:tc:opendocument:xmlns:chart:1.0}repeated','%i'%(table_desc.nrows-2,))
    
    if len(plot_desc.ycolL) > 1:
        auto_styles = root.find('office:automatic-styles', root.nsmap)
        autostyleL = auto_styles.findall('style:style', root.nsmap)
        ref_series_style = None
        
        istyle_loc = 0
        for iloc,astyle in enumerate(autostyleL):
            if astyle.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name') == 'G0S0':
                ref_series_style = astyle
                istyle_loc = iloc + 2
        #print( ref_series_style.items())
        
        nG0S = 1
        for ycol in plot_desc.ycolL[1:]:
            new_style = deepcopy( ref_series_style )
            new_style.set('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'G0S%i'%nG0S )
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
    
    
