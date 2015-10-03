from odscharts.color_utils import BIG_COLORNAME_LIST,  EXCEL_COLOR_LIST, getValidHexStr,\
            VERY_SHORT_NAME_DICT, COLOR_NAME_DICT, get_best_gray_text
            
            
fOut = open('./_static/odscharts_colors.html', 'w')

def start_table():
    sL.append( '<P><TABLE bgcolor="#666666">' )

def end_table():
    sL.append( '</TABLE>' )
    
def start_row():
    sL.append('<TR>')
    
def end_row():
    sL.append('</TR>')
   
def start_cell( bgcolor="#FFFFFF" ):
    sL.append('<TD bgcolor="%s">'%bgcolor)
    
def end_cell():
    sL.append('</TD>')
   
def add_text(text, color="#000000", br=False):
    sL.append('<SPAN style="color:%s">%s </SPAN>'%(color, text))
    if br:
        sL[-1] = sL[-1] + '<br>'
    
def do_xxx():
    sL.append('')

def build_color_table( colorL, ncol=6, show_full_name=False, hex_only=False ):
    start_table()
    
    for i,c in enumerate(colorL):
        if i == 0:
            start_row()
        hexstr = getValidHexStr(c, c)
        start_cell(bgcolor=hexstr)
        c_best = get_best_gray_text( hexstr )[0]
        
        if not hex_only:
            add_text(c, color=c_best, br=True)
        
        if show_full_name:
            c_full = VERY_SHORT_NAME_DICT[c]
            add_text(c_full, color=c_best, br=True)
        
        add_text(hexstr, color=c_best, br=False)
        end_cell()
        
        if i % ncol == ncol-1:
            end_row()
            start_row()
            
    if i % ncol  != ncol-1:
        end_row()
    end_table()
        

# Show Excel Color List
sL = ['<br><br><H1>Excel Default Color Table</H1>']
build_color_table( EXCEL_COLOR_LIST, ncol=6, show_full_name=False, hex_only=True )

# Show VERY_SHORT_NAME_DICT
sL.append( '<br><br><H1>Short Name Color Table</H1>' )
nameL = sorted( VERY_SHORT_NAME_DICT.keys() )
build_color_table( nameL, ncol=8, show_full_name=True )


# Show Alternate Color List
sL.append( '<br><br><H1>Alternate Color Table</H1>' )
build_color_table( BIG_COLORNAME_LIST, ncol=6, show_full_name=False, hex_only=False )


fOut.write( '\n'.join( sL ) )