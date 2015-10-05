"""
A PlotTableDesc object holds the XML logic as well as all info about a
scatter plot table
"""
import sys
if sys.version_info < (3,):
    import ElementTree_27OD as ET
else:
    import ElementTree_34OD as ET

from color_utils import BIG_COLOR_HEXSTR_LIST,  EXCEL_COLOR_LIST, getValidHexStr

SYMBOL_TYPE_LIST = ["diamond", "square", "arrow-up",  "arrow-down", "arrow-right", 
                    "arrow-left", "bow-tie", "hourglass", "circle", "star", "x", 
                    "plus", "asterisk"]

class PlotTableDesc(object):
    """Holds a description of a scatter plot sheet.
    """

    def get_next_color(self):
        """
        Return the next color in the list.
        Increment counter for next color.
        """
        i = self.i_color
        self.i_color += 1
        c = self.color_list[ i % len(self.color_list) ]

        #print( 'i_color = %i,  color = "%s"'%(i,c) )

        return c

    def get_next_symbol_type(self):
        """
        Return the next symbol type in SYMBOL_TYPE_LIST
        """
        i = self.i_symbol_type
        self.i_symbol_type += 1
        
        return SYMBOL_TYPE_LIST[i % len(SYMBOL_TYPE_LIST)]
        

    def __init__(self, plot_sheetname, num_chart, parent_obj, excel_colors=True):
        """Inits SpreadSheet with filename and blank content.

        NS: cleans up the namespace callouts of the xml Element

        Attributes::

            plot_sheetname: name of data sheet

        """
        self.i_color = 0 # index into color chart for next curve
        self.i_prim_color = 0 # index into colorL
        self.i_sec_color = 0  # index into color2L
        
        self.i_symbol_type = 0 # index into SYMBOL_TYPE_LIST

        if excel_colors:
            self.color_list = EXCEL_COLOR_LIST
        else:
            self.color_list = BIG_COLOR_HEXSTR_LIST

        NS = parent_obj.NS

        # Start building new xml Element to be new Sheet in spreadsheet
        attribD = {NS('table:name'):plot_sheetname}
        newsheet = ET.Element(NS('table:table'), attrib=attribD)

        table_shapes  = ET.Element(NS('table:shapes'))
        #print( 'table_shapes.text =',table_shapes.text )

        def NS_attrib( attD ):
            D = {}
            for key,val in attD.items():
                D[ NS(key) ] = val
            return D

        attribD = NS_attrib({ 'draw:z-index':"1", 'draw:id':"id0", 'draw:style-name':"a0", 'draw:name':"Chart 1", 'svg:x':"0in",
                     'svg:y':"0in", 'svg:width':"9.47551in", 'svg:height':"6.88048in", 'style:rel-width':"scale",
                     'style:rel-height':"scale"})
        draw_frame = ET.Element(NS('draw:frame'), attrib=attribD)

        attribD = NS_attrib({'xlink:href':"Object %i/"%num_chart, 'xlink:type':"simple", 'xlink:show':"embed", 'xlink:actuate':"onLoad"})
        draw_obj = ET.Element(NS('draw:object'), attrib=attribD)

        svg_title  = ET.Element(NS('svg:title'))
        svg_desc  = ET.Element(NS('svg:desc'))

        draw_frame.append( draw_obj )
        draw_frame.append( svg_title )
        draw_frame.append( svg_desc )
        table_shapes.append( draw_frame )


        tab_col = ET.Element(NS('table:table-column'), attrib={NS('table:number-columns-repeated'):"16384"})
        tab_row = ET.Element(NS('table:table-row'), attrib={NS('table:number-rows-repeated'):"1048576"})
        tab_cell = ET.Element(NS('table:table-cell'), attrib={NS('table:number-columns-repeated'):"16384"})
        tab_row.append( tab_cell )


        newsheet.append( table_shapes )
        newsheet.append( tab_col )
        newsheet.append( tab_row )

        # make sure any added Element objects are in nsOD, rev_nsOD and qnameOD of parent_obj
        def add_tag( tag ):
            sL = tag.split('}')
            uri = sL[0][1:]
            name = sL[1]
            parent_obj.qnameOD[tag] = parent_obj.nsOD[uri] + ':' + name

        def add_tags( obj ):
            if hasattr(obj,'tag'):
                add_tag( obj.tag )
            if hasattr(obj, 'attrib'):
                for q,v in obj.attrib.items():
                    add_tag( q )

        for parent in newsheet.iter():
            add_tags( parent )
            for child in parent.getchildren():
                add_tags( child )


        self.xmlSheetObj = newsheet


        self.plot_sheetname = ''
        self.data_sheetname = ''
        self.title = ''
        self.xlabel = ''
        self.ylabel = ''
        self.y2label = ''
        self.ycolL = None
        self.ycol2L = None

        self.logx = False
        self.logy = False
        self.log2y = False

        self.xcolL = []
        self.ycolDataSheetNameL = []

        self.xcol2L = []
        self.ycol2_DataSheetNameL = []

        self.showMarkerL = None
        self.showMarker2L = None
        self.markerTypeL = None
        self.markerType2L = None
        
        self.showLineL = None
        self.showLine2L = None

        self.lineThkL = None
        self.lineThk2L = None
        
        self.lineStyleL = None
        self.lineStyle2L = None
        self.set_of_line_styles = set()
        
        self.markerHtWdL = None
        self.markerHtWd2L = None
        
        self.showUnits = True

        self.colorL = None
        self.color2L = None

        self.labelL = None
        self.label2L = None

    def add_to_primary_y(self, data_sheetname, xcol, ycolL,
                            showMarkerL=None, showLineL=None,
                            lineThkL=None, lineStyleL=None, 
                            colorL=None, labelL=None):
        """
        Add new curves to primary y axis
        """
        if ycolL is None:
            return

        #print('colorL =',colorL)
        # self.colorL will always be None 1st time through
        if self.colorL is None:
            self.colorL = []
            self.labelL = []
            self.showMarkerL = []
            self.markerTypeL = []
            self.showLineL = []
            self.lineThkL = []
            self.lineStyleL = []
            self.markerHtWdL = []
            self.ycolL = []
            self.xcolL = []
            self.ycolDataSheetNameL = []

        # Fill attribute lists with input list values
        for i,ycol in enumerate(ycolL):
            self.ycolL.append( ycol )
            self.xcolL.append( xcol )
            self.ycolDataSheetNameL.append( data_sheetname )

            # See if there's an input color
            c_inp = get_ith_value( colorL, i, None )
            c_palette = self.get_next_color()
            c = getValidHexStr( c_inp, c_palette)
            self.colorL.append( c )
            
            self.markerTypeL.append( self.get_next_symbol_type() )

            self.labelL.append( get_ith_value( labelL, i, None ) )

            if len(self.showMarkerL):
                self.showMarkerL.append( get_ith_value( showMarkerL, i, self.showMarkerL[-1] ) )
            else:
                self.showMarkerL.append( get_ith_value( showMarkerL, i, True ) )

            if len(self.showLineL):
                self.showLineL.append( get_ith_value( showLineL, i, self.showLineL[-1] ) )
            else:
                self.showLineL.append( get_ith_value( showLineL, i, True ) )

            # first get number into line thickness
            if len(self.lineThkL):
                self.lineThkL.append( get_ith_value( lineThkL, i, self.lineThkL[-1] ) )
            else:
                self.lineThkL.append( get_ith_value( lineThkL, i, 0.8 ) )
                
            # then convert to mm
            try:
                v = float(self.lineThkL[-1])
                self.markerHtWdL.append("%gmm"%(v*3,))
                self.lineThkL[-1] = "%gmm"%v
            except:
                self.markerHtWdL.append("%2.4mm")
                self.lineThkL[-1] = "0.8mm"

            # Set line style to solid unless otherwise indicated
            if len(self.lineStyleL):
                self.lineStyleL.append( get_ith_value( lineStyleL, i, self.lineStyleL[-1] ) )
            else:
                self.lineStyleL.append( get_ith_value( lineStyleL, i, 0 ) )
                
            # include all y curves (primary and secondary) in set_of_line_styles
            self.set_of_line_styles.update( set(self.lineStyleL) )
            self.set_of_line_styles.discard(0)  # Make sure that 0 does not appear
            

    def add_to_secondary_y(self, data_sheetname, xcol, ycol2L,
                              showMarker2L=None, showLine2L=None,
                              lineThk2L=None, lineStyle2L=None, 
                              color2L=None, label2L=None):
        """
        Add new curves to primary y axis
        """
        if ycol2L is None:
            return

        # self.color2L will always be None 1st time through
        if self.color2L is None:
            self.color2L = []
            self.label2L = []
            self.showMarker2L = []
            self.markerType2L = []
            self.showLine2L = []
            self.lineThk2L = []
            self.lineStyle2L = []
            self.markerHtWd2L = []
            self.ycol2L = []
            self.xcol2L = []
            self.ycol2_DataSheetNameL = []

        # Fill attribute lists with input list values
        for i,ycol in enumerate(ycol2L):
            self.ycol2L.append( ycol )
            self.xcol2L.append( xcol )
            self.ycol2_DataSheetNameL.append( data_sheetname )

            # See if there's an input color
            c_inp = get_ith_value( color2L, i, None )
            c_palette = self.get_next_color()
            c = getValidHexStr( c_inp, c_palette)
            self.color2L.append( c )
            
            self.markerType2L.append( self.get_next_symbol_type() )

            self.label2L.append( get_ith_value( label2L, i, None ) )

            if len(self.showMarker2L):
                self.showMarker2L.append( get_ith_value( showMarker2L, i, self.showMarker2L[-1] ) )
            else:
                self.showMarker2L.append( get_ith_value( showMarker2L, i, True ) )

            if len(self.showLine2L):
                self.showLine2L.append( get_ith_value( showLine2L, i, self.showLine2L[-1] ) )
            else:
                self.showLine2L.append( get_ith_value( showLine2L, i, True ) )

            # first get number into line thickness
            if len(self.lineThk2L):
                self.lineThk2L.append( get_ith_value( lineThk2L, i, self.lineThk2L[-1] ) )
            else:
                self.lineThk2L.append( get_ith_value( lineThk2L, i, 0.8 ) )
                
            # then convert to mm
            try:
                v = float(self.lineThk2L[-1])
                self.markerHtWd2L.append("%gmm"%(v*3,))
                self.lineThk2L[-1] = "%gmm"%v
            except:
                self.markerHtWd2L.append("%2.4mm")
                self.lineThk2L[-1] = "0.8mm"

            # Set line style to solid unless otherwise indicated
            if len(self.lineStyle2L):
                self.lineStyle2L.append( get_ith_value( lineStyle2L, i, self.lineStyle2L[-1] ) )
            else:
                self.lineStyle2L.append( get_ith_value( lineStyle2L, i, 0 ) )
                
            # include all y curves (primary and secondary) in set_of_line_styles
            self.set_of_line_styles.update( set(self.lineStyle2L) )
            self.set_of_line_styles.discard(0)  # Make sure that 0 does not appear

def get_ith_value( valL, i, default_val ):
    """Return the ith value in valL if possible, otherwise default_val"""
    #print('valL=',valL,'  default_val=',default_val)
    try:
        val = valL[i]
    except:
        val = default_val
    return val


