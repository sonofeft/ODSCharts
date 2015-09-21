"""
A PlotTableDesc object holds the XML logic as well as all info about a 
scatter plot table
"""

from lxml import etree as ET

class PlotTableDesc(object):
    """Holds a description of a scatter plot sheet.
    """

    def __init__(self, plot_sheetname, num_chart, NS, nsmap):
        """Inits SpreadSheet with filename and blank content.
        
        NS: cleans up the namespace callouts of the lxml Element
        nsmap: dictionary containing documents namespace definitions
        
        Attributes::
        
            plot_sheetname: name of data sheet
        
        """
        try:
            register_namespace = ET.register_namespace
        except AttributeError:
            def register_namespace(prefix, uri):
                ET._namespace_map[uri] = prefix
        
        for prefix, uri in nsmap.items():
            ET.register_namespace(prefix, uri)
        
        # Start building new lxml Element to be new Sheet in spreadsheet
        attribD = {NS('table:name'):plot_sheetname}
        newsheet = ET.Element(NS('table:table'), attrib=attribD, nsmap=nsmap)

        table_shapes  = ET.Element(NS('table:shapes'), nsmap=nsmap)
        
        def NS_attrib( attD ):
            D = {}
            for key,val in attD.items():
                D[ NS(key) ] = val
            return D
        
        attribD = NS_attrib({ 'draw:z-index':"1", 'draw:id':"id0", 'draw:style-name':"a0", 'draw:name':"Chart 1", 'svg:x':"0in",
                     'svg:y':"0in", 'svg:width':"9.47551in", 'svg:height':"6.88048in", 'style:rel-width':"scale",
                     'style:rel-height':"scale"})
        draw_frame = ET.Element(NS('draw:frame'), attrib=attribD, nsmap=nsmap)

        attribD = NS_attrib({'xlink:href':"Object %i/"%num_chart, 'xlink:type':"simple", 'xlink:show':"embed", 'xlink:actuate':"onLoad"})
        draw_obj = ET.Element(NS('draw:object'), attrib=attribD, nsmap=nsmap)
        
        svg_title  = ET.Element(NS('svg:title'), nsmap=nsmap)
        svg_desc  = ET.Element(NS('svg:desc'), nsmap=nsmap)
        
        draw_frame.append( draw_obj )
        draw_frame.append( svg_title )
        draw_frame.append( svg_desc )
        table_shapes.append( draw_frame )
        
        
        tab_col = ET.Element(NS('table:table-column'), attrib={NS('table:number-columns-repeated'):"16384"}, nsmap=nsmap)
        tab_row = ET.Element(NS('table:table-row'), attrib={NS('table:number-rows-repeated'):"1048576"}, nsmap=nsmap)
        tab_cell = ET.Element(NS('table:table-cell'), attrib={NS('table:number-columns-repeated'):"16384"}, nsmap=nsmap)
        tab_row.append( tab_cell )
        
        
        newsheet.append( table_shapes )
        newsheet.append( tab_col )
        newsheet.append( tab_row )
        
                
        self.table_obj = newsheet
