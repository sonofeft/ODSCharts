"""
A DataTableDesc object holds the XML logic as well as all info about a data table
"""

from lxml import etree as ET


class DataTableDesc(object):
    """Holds a description of a data table sheet.
    """

    def __init__(self, data_sheetname, list_of_rows, NS, nsmap):
        """Inits SpreadSheet with filename and blank content.
        
        NS: cleans up the namespace callouts of the lxml Element
        nsmap: dictionary containing documents namespace definitions
        
        Attributes::
        
            data_sheetname: name of data sheet
            nrows: number of rows including label and units rows
            ncols: max number of cols
            labelL: list of labels
            unitsL: list of units 
        
        """
        
        self.data_sheetname = data_sheetname
        self.list_of_rows = list_of_rows
        
        # calc number of rows and columns in data
        self.nrows = len( list_of_rows )
        self.ncols = 0
        for row in list_of_rows:
            self.ncols = max(self.ncols, len(row))
        
        # make sure that row1 and row2 contain labels and units
        self.labelL = []
        self.unitsL = []
        for n in range( self.ncols ):
            try:
                s = str( list_of_rows[0][n] ).strip()
            except:
                s = 'Column %i'%(n+1,)
            self.labelL.append( s )
            
            try:
                s = str( list_of_rows[1][n] ).strip()
            except:
                s = ''
            self.unitsL.append( s )

        # Start building new lxml Element to be new Sheet in spreadsheet
        #print( nsmap )
        attribD = {NS('table:name'):data_sheetname, NS('table:style-name'):"ta1"}
        newsheet = ET.Element(NS('table:table'), attrib=attribD, nsmap=nsmap)
        
        colD = {NS('table:style-name'):"co1", NS('table:number-columns-repeated'):"16384", 
                NS('table:default-cell-style-name'):"ce1"}
        col_elm = ET.Element(NS('table:table-column'), attrib=colD, nsmap=nsmap)
        newsheet.append( col_elm )
        
        str_cellD = {NS('office:value-type'):"string", NS('table:style-name'):"ce1"}
        float_cellD = {NS('office:value-type'):"float", NS('table:style-name'):"ce1"}
            
        row_rep = 1048576 # max number of rows
            
        for row in list_of_rows:
            row_obj = ET.Element(NS('table:table-row'), 
                                 attrib={NS('table:style-name'):"ro1"}, 
                                 nsmap=nsmap)
                
            col_rep = 16384 # max number of columns
                
            for value in row:
                try:
                    fval = float( value )
                    float_cellD[NS('office:value')] = str(fval)
                    D = float_cellD
                    text_p = "%g"%fval
                except:
                    D = str_cellD
                    text_p = "%s"%value
                
                cell_obj = ET.Element(NS('table:table-cell'), attrib=D, nsmap=nsmap)
                text_obj = ET.Element(NS('text:p'),  nsmap=nsmap)
                text_obj.text = text_p
                cell_obj.append( text_obj )
                row_obj.append(cell_obj)
                col_rep -= 1 # decrement max available columns remaining

            last_cell_obj = ET.Element(NS('table:table-cell'), 
                                       attrib={NS('table:number-columns-repeated'):"%i"%col_rep}, 
                                       nsmap=nsmap)
            row_obj.append(last_cell_obj)
            newsheet.append( row_obj )
            row_rep -= 1 # decrement max remaining number of rows
            
        last_row_obj = ET.Element(NS('table:table-row'), 
                             attrib={NS('table:style-name'):"ro1", 
                                     NS('table:number-rows-repeated'):"%i"%row_rep}, 
                             nsmap=nsmap)
        last_row_obj.append( ET.Element(NS('table:table-cell'), 
                                        attrib={NS('table:number-columns-repeated'):"16384"}, 
                                        nsmap=nsmap) )

        newsheet.append( last_row_obj )
        
        self.table_obj = newsheet
