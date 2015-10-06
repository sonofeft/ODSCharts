

from copy import deepcopy
from collections import OrderedDict
from find_obj import find_elem_w_attrib, elem_set, NS_attrib, NS

import sys
if sys.version_info < (3,):
    import ElementTree_27OD as ET
else:
    import ElementTree_34OD as ET

# 'Solid' is a place-holder 
#    such that lineStyle==0 will not generate a new style of draw:stroke-dash
#                     0      1      2         3            4               5
LINE_STYLE_LIST = ['Solid','Dot','Dash', 'ShortDash','ShortDot','ShortDashDot',
                   'ShortDashDotDot','LongDash','LongDashDot','LongDashDotDot','DashDot']
#                          6            7             8             9              10


def get_display_style_name(i):
    """Returns the attribute value of draw:display-name (e.g. "LongDash")"""
    
    ii = i % len(LINE_STYLE_LIST)
    if ii==0:
        return ''
    else:
        return LINE_STYLE_LIST[ii]
        
def get_dash_a_name( i_style ):
    return "a%i"%(i_style+4,)  # names start at "a5" since i_style==0 is not allowed

def gen_dash_elements_from_set_of_istyles( parent, set_of_i_styles ):
    """
    Add a bunch of new draw:stroke-dash SubElement to parent
    """
    i_styleL = sorted( list(set_of_i_styles) )
    for i_style in i_styleL:
        generate_xml_drawStrokeDash_element( parent, i_style )

def generate_xml_drawStrokeDash_element( parent, i_style ):
    """
    Adds a new draw:stroke-dash SubElement to parent containing draw:name="a%i"%(i_style+4,)
    
    i_style is an index into LINE_STYLE_LIST 
    such that draw:display-name="%s"%LINE_STYLE_LIST[i_style]
    """
    
    i_style = i_style % len(LINE_STYLE_LIST)
    if i_style==0:
        return # No SubElement genertated for solid line
    
    display_name = LINE_STYLE_LIST[i_style]
    
    a_name = get_dash_a_name( i_style )  # names start at "a5" since i_style==0 is not allowed
    
    def make_subelem(ndot1=1, ldot1='0.25in', ndot2=0, ldot2='0in', dist='0.03125in', style='rect'):
        """
        style can be "rect" or "round" (hard to tell the difference)
        """
        new_elem_2 = ET.SubElement(parent,"{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash", 
        attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name', '%s'%a_name), 
        ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}display-name', '%s'%display_name), 
        ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style', '%s'%style), 
        ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1', '%i'%ndot1), 
        ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1-length', '%s'%ldot1), 
        ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2', '%i'%ndot2), 
        ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2-length', '%s'%ldot2), 
        ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}distance', '%s'%dist)]))  

    # Guess distances from Excel and OpenOffice files
    # short dot="0.015625in", dot="0.03125in"
    # short dash="0.0625in",  dash="0.125in",  long dash="0.25in"
    
    if display_name=='LongDash': # 2 for Excel
        make_subelem(ndot1=1, ldot1='0.25in', ndot2=0, ldot2='0in')
    if display_name=='LongDashDot': # 3 for Excel
        make_subelem(ndot1=1, ldot1='0.25in', ndot2=1, ldot2='0.03125in')
    if display_name=='LongDashDotDot': # 4 for Excel
        make_subelem(ndot1=1, ldot1='0.25in', ndot2=2, ldot2='0.03125in')
    if display_name=='DashDot': # guess from OO
        make_subelem(ndot1=1, ldot1='0.125in', ndot2=1, ldot2='0.03125in')
    if display_name=='Dot': # guess from Excel
        make_subelem(ndot1=1, ldot1='0.03125in', ndot2=0, ldot2='0in')
    if display_name=='Dash':
        make_subelem(ndot1=1, ldot1='0.125in', ndot2=0, ldot2='0in')
    if display_name=='ShortDash':
        make_subelem(ndot1=1, ldot1='0.0625in', ndot2=0, ldot2='0in')
    if display_name=='ShortDot':
        make_subelem(ndot1=1, ldot1='0.015625in', ndot2=0, ldot2='0in')
    if display_name=='ShortDashDot':
        make_subelem(ndot1=1, ldot1='0.0625in', ndot2=1, ldot2='0.015625in')
    if display_name=='ShortDashDotDot':
        make_subelem(ndot1=1, ldot1='0.0625in', ndot2=2, ldot2='0.015625in')
        
    return

"""
    # some of OpenOffice line styles
    new_elem_3 = ET.SubElement(parent,"{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash", 
    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name', '_33__20_Dashes_20_3_20_Dots_20__28_var_29_'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}display-name', '3 Dashes 3 Dots (var)'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style', 'rect'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1', '3'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1-length', '197%'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2', '3'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}distance', '100%')]))
    
    new_elem_4 = ET.SubElement(parent,"{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash", 
    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name', 'Dash_20_2'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}display-name', 'Dash 2'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style', 'rect'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1', '1'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1-length', '0.02cm'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2', '1'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2-length', '0.02cm'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}distance', '0.02cm')]))
    
    new_elem_5 = ET.SubElement(parent,"{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash", 
    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name', 'Fine_20_Dashed_20__28_var_29_'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}display-name', 'Fine Dashed (var)'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style', 'rect'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1', '1'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1-length', '197%'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}distance', '197%')]))
    
    new_elem_6 = ET.SubElement(parent,"{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash", 
    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name', 'Ultrafine_20_Dashed'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}display-name', 'Ultrafine Dashed'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style', 'rect'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1', '1'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1-length', '0.051cm'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2', '1'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2-length', '0.051cm'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}distance', '0.051cm')]))
     
    # 2
    new_elem_2 = ET.SubElement(parent,"{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash", 
    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name', 'a5'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}display-name', 'LongDash'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style', 'rect'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1', '1'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1-length', '0.25in'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2', '0'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2-length', '0in'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}distance', '0.03125in')]))    

    # 3
    new_elem_1 = ET.SubElement(parent,"{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash", 
    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name', 'a6'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}display-name', 'LongDashDot'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style', 'rect'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1', '1'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1-length', '0.25in'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2', '1'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2-length', '0.03125in'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}distance', '0.03125in')]))
   
   # 4
    new_elem_1 = ET.SubElement(parent,"{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash", 
    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name', 'a7'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}display-name', 'LongDashDotDot'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style', 'rect'),
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1', '1'),
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1-length', '0.25in'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2', '2'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2-length', '0.03125in'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}distance', '0.03125in')]))
  
    # 5
    new_elem_1 = ET.SubElement(parent,"{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke-dash", 
    attrib=OrderedDict([('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name', 'a8'),
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}display-name', 'SysDash'),
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style', 'rect'),
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1', '1'),
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots1-length', '0.03125in'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2', '0'),
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}dots2-length', '0in'), 
    ('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}distance', '0.03125in')]))
"""