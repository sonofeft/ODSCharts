# Python 2 and 3
from __future__ import unicode_literals
from __future__ import absolute_import
from __future__ import print_function
import re
import colorsys

# These default colors are taken from Excel
#  Since Excel does not support marker colors in ODS format, try to just use
#  the Excel default colors so the marker color will match the line color
EXCEL_COLOR_LIST = ['#325885', '#873331', '#6b833a', '#574271', '#2f788c', '#b0672b',
                    '#3c679a', '#9d3d3a', '#7d9844', '#664e83', '#388ca2', '#cb7833',
                    '#4373ab', '#ae4441', '#8ba94d', '#725792', '#3f9bb4', '#e1853a',
                    '#4a7ebb', '#be4b48', '#98b954', '#7d60a0', '#46aac5', '#f69240',
                    '#809cc7', '#ca817f', '#aec685', '#9c8bb3', '#7ebace', '#f6a97b',
                    '#a5b6d3', '#d5a5a4', '#c1d2a7', '#b6abc5', '#a4cad8', '#f6bea2']

def w3_luminance(r,g,b): # float values of r,g,b
    def condition( c ): # c is r,g or b
        if c<0.03928:
            val = c / 12.92
        else:
            val = ((c+0.055)/1.055)**2.4
        return val

    return 0.2126*condition(r) + 0.7152*condition(g) + 0.0722*condition(b)

def w3_contrast( webStr_1, webStr_2 ):
    L1 = w3_luminance( *hex_to_rgbfloat(webStr_1) )
    L2 = w3_luminance( *hex_to_rgbfloat(webStr_2) )

    if L1>L2:
        return (L1 + 0.05) / (L2 + 0.05)
    else:
        return (L2 + 0.05) / (L1 + 0.05)

def hex_to_rgb(hex_str):
    """Returns a tuple representing the given hex string as RGB.

    >>> hex_to_rgb('CC0000')
    (204, 0, 0)
    """
    if hex_str.startswith('#'):
        hex_str = hex_str[1:]
    ans_tup = tuple([int(hex_str[i:i + 2], 16) for i in xrange(0, len(hex_str), 2)])
    #print 'for',hex_str,'ans_tup=',ans_tup
    return ans_tup


def hex_to_rgbfloat(hex_str):
    """Returns a tuple representing the given hex string as RGB.

    >>> hex_to_rgb('CC0000')
    (204, 0, 0)
    """
    if hex_str.startswith('#'):
        hex_str = hex_str[1:]
    ans_tup = tuple([int(hex_str[i:i + 2], 16) for i in xrange(0, len(hex_str), 2)])
    #print 'for',hex_str,'ans_tup=',ans_tup
    return tuple([round(float(c)/255, 8) for c in ans_tup])


def rgb_to_hex(rgb):
    """Converts an rgb tuple to hex string for web.

    >>> rgb_to_hex((204, 0, 0))
    '#CC0000'
    """
    return '#' + ''.join(["%0.2X" % c for c in rgb])

def rgbfloat_to_hex(rgb_float):
    """Converts an rgb tuple to hex string for web.

    >>> rgb_to_hex((0.8, 0.0, 0.0))
    '#CC0000'
    """
    rgb = tuple([int(c*255) for c in rgb_float])
    return '#' + ''.join(["%0.2X" % c for c in rgb])


def scale_rgb_tuple(rgb, down=True):
    """Scales an RGB tuple up or down to/from values between 0 and 1.

    >>> scale_rgb_tuple((204, 0, 0))
    (.80, 0, 0)
    >>> scale_rgb_tuple((.80, 0, 0), False)
    (204, 0, 0)
    """
    if not down:
        return tuple([int(c*255) for c in rgb])
    return tuple([round(float(c)/255, 8) for c in rgb])

def lighten_color(web_hex_str, ratio=0.5):
    hue, sat, val = colorsys.rgb_to_hsv(*scale_rgb_tuple(hex_to_rgb(web_hex_str)))
    val = val + (1.0-val) * ratio
    sat = sat * ratio

    lighter_str = rgb_to_hex( scale_rgb_tuple(colorsys.hsv_to_rgb(hue, sat, val), False) )
    return lighter_str

def darken_color(web_hex_str, ratio=0.5):
    hue, sat, val = colorsys.rgb_to_hsv(*scale_rgb_tuple(hex_to_rgb(web_hex_str)))
    val = val * ratio
    sat = sat + (1.0-sat) * ratio
    darker_str = rgb_to_hex( scale_rgb_tuple(colorsys.hsv_to_rgb(hue, sat, val), False) )
    return darker_str

def get_best_same_color_text(web_hex_str, lum_min=0.1):
    '''Returns biggest YIQ contrast'''
    r,g,b = hex_to_rgb(web_hex_str)
    r,g,b = scale_rgb_tuple((r,g,b), down=True)

    y,i,q =  colorsys.rgb_to_yiq( r,g,b )
    if y>0.5:
        rgb_float = colorsys.yiq_to_rgb(lum_min,i,q)
    else:
        rgb_float = colorsys.yiq_to_rgb(1.0-lum_min,i,q)

    #print 'y,i,q=',y,i,q, '    rgb_float=',rgb_float

    return rgbfloat_to_hex(rgb_float)


def get_best_gray_text(web_hex_str):
    '''Returns 3 choices: most contrast to least contrast'''
    r,g,b = hex_to_rgb(web_hex_str)

    # Counting the perceptive luminance - human eye favors green color...
    lum =  1.0 - ( 0.299 * r + 0.587 * g + 0.114 * b)/255.0

    if lum < 0.5: # bright color use black font
        return '#000000','#111111','#222222'
    else:         # dark color use white font
        return '#ffffff','#eeeeee','#dddddd'

def get_complement(web_hex_str):
    r,g,b = hex_to_rgb(web_hex_str)
    return rgb_to_hex((255-r,255-g,255-b))

def make_color_rotation(web_hex_str):
    r,g,b = hex_to_rgb(web_hex_str)
    c2 = rgb_to_hex((g,b,r))
    c3 = rgb_to_hex((b,r,g))
    return [web_hex_str,c2,c3]

def make_triad(web_hex_str, rot_deg=30.0):
    """Returns a list 3 of hex strings in web format to be used for triad
    color schemes from the given base color.

    make_triad('#CC0000')
    ['#CC0000', '#A30000', '#660000']
    """
    if not web_hex_str.startswith('#'):
        web_hex_str = '#' + web_hex_str

    colors = [web_hex_str]
    orig_rgb = hex_to_rgb(web_hex_str)
    hue, sat, val = colorsys.rgb_to_hsv(*scale_rgb_tuple(orig_rgb))

    # rotate +rot_deg
    d_hue = rot_deg / 360.0
    d20 = ( (hue+d_hue)%1.0, sat, val)
    colors.append(rgb_to_hex( scale_rgb_tuple(colorsys.hsv_to_rgb(*d20), False) ))

    d20 = ( (1.0 + hue-d_hue)%1.0, sat, val)
    colors.append(rgb_to_hex( scale_rgb_tuple(colorsys.hsv_to_rgb(*d20), False) ))

    return colors



stdList = ['red','green','blue','cyan','magenta','olive','brown','gold','coral',
    'darkred','darkgreen','darkblue']
def getColorName(i):
    """Returns name of color in position i.  for example 0="red", 2="blue"  """
    return BIG_COLORNAME_LIST[ i % len(BIG_COLORNAME_LIST) ]


def getColorHexStr(i):
    """Returns hex string of color in position i.  for example 0="#FF0000", 2="#0000FF"  """
    return BIG_COLOR_HEXSTR_LIST[i % len(BIG_COLOR_HEXSTR_LIST)]
    #return COLOR_NAME_DICT[ getColorName(i) ]


hexstr_pattern = re.compile( r'#(?:[a-fA-F0-9]{3}|[a-fA-F0-9]{6})\b' )

def getValidHexStr( name_or_hex, c_default):
    """
    Returns a valid hex string.
    name_or_hex can be an index name in COLOR_NAME_DICT or a hex string.
    If name_or_hex is not valid, return c_default
    """
    c = '%s'%name_or_hex
    #print('(c=%s)'%c, end='')
    c = c.lower()

    if c in VERY_SHORT_NAME_DICT:
        c = VERY_SHORT_NAME_DICT[c]

    if c in COLOR_NAME_DICT:
        #print('(in dict)', end='')
        return COLOR_NAME_DICT[ c ]

    c = c.upper()
    if not c.startswith('#'):
        c = '#' + c

    # change from short form to full length form
    if len(c)==4:
        c = c[0] + c[1] + c[1] + c[2] + c[2] + c[3] + c[3]

    if c in BIG_COLOR_HEXSTR_LIST:
        #print('(in big list)', end='')
        return c

    if hexstr_pattern.match( c ):
        #print('(re match)', end='')
        return c

    # Nothing looks right so return the default
    #print('(Default Color)')
    return c_default

VERY_SHORT_NAME_DICT = {'r':'red', 'g':'green', 'b':'blue', 'c':'cyan',
                        'm':'magenta', 'k':'black', 'y':'yellow', 'p':'purple',
                        't':'tan', 'n':'navy', 'i':'indigo', 'v':'violet',
                        's':'salmon', 'o':'olive', 'dr':'darkred', 'db':'darkblue',
                        'dg':'darkgreen', 'dm':'darkmagenta', 'dc':'darkcyan',
                        'dv':'darkviolet', 's':'sienna',
                        'do':'darkolivegreen', 'f':'firebrick', 'ds':'darksalmon', 
                        'dk':'darkkhaki' }

COLOR_NAME_DICT = {
    'aliceblue'            : '#F0F8FF',
    'antiquewhite'         : '#FAEBD7',
    'aqua'                 : '#00FFFF',
    'aquamarine'           : '#7FFFD4',
    'azure'                : '#F0FFFF',
    'beige'                : '#F5F5DC',
    'bisque'               : '#FFE4C4',
    'black'                : '#000000',
    'blanchedalmond'       : '#FFEBCD',
    'blue'                 : '#0000FF',
    'blueviolet'           : '#8A2BE2',
    'brown'                : '#A52A2A',
    'burlywood'            : '#DEB887',
    'cadetblue'            : '#5F9EA0',
    'chartreuse'           : '#7FFF00',
    'chocolate'            : '#D2691E',
    'coral'                : '#FF7F50',
    'cornflowerblue'       : '#6495ED',
    'cornsilk'             : '#FFF8DC',
    'crimson'              : '#DC143C',
    'cyan'                 : '#00FFFF',
    'darkblue'             : '#00008B',
    'darkcyan'             : '#008B8B',
    'darkgoldenrod'        : '#B8860B',
    'darkgray'             : '#A9A9A9',
    'darkgreen'            : '#006400',
    'darkkhaki'            : '#BDB76B',
    'darkmagenta'          : '#8B008B',
    'darkolivegreen'       : '#556B2F',
    'darkorange'           : '#FF8C00',
    'darkorchid'           : '#9932CC',
    'darkred'              : '#8B0000',
    'darksalmon'           : '#E9967A',
    'darkseagreen'         : '#8FBC8F',
    'darkslateblue'        : '#483D8B',
    'darkslategray'        : '#2F4F4F',
    'darkturquoise'        : '#00CED1',
    'darkviolet'           : '#9400D3',
    'deeppink'             : '#FF1493',
    'deepskyblue'          : '#00BFFF',
    'dimgray'              : '#696969',
    'dodgerblue'           : '#1E90FF',
    'firebrick'            : '#B22222',
    'floralwhite'          : '#FFFAF0',
    'forestgreen'          : '#228B22',
    'fuchsia'              : '#FF00FF',
    'gainsboro'            : '#DCDCDC',
    'ghostwhite'           : '#F8F8FF',
    'gold'                 : '#FFD700',
    'goldenrod'            : '#DAA520',
    'gray'                 : '#808080',
    'green'                : '#008000',
    'greenyellow'          : '#ADFF2F',
    'honeydew'             : '#F0FFF0',
    'hotpink'              : '#FF69B4',
    'indigo'               : '#4B0082',
    'ivory'                : '#FFFFF0',
    'khaki'                : '#F0E68C',
    'lavender'             : '#E6E6FA',
    'lavenderblush'        : '#FFF0F5',
    'lawngreen'            : '#7CFC00',
    'lemonchiffon'         : '#FFFACD',
    'lightblue'            : '#ADD8E6',
    'lightcoral'           : '#F08080',
    'lightcyan'            : '#E0FFFF',
    'lightgoldenrodyellow' : '#FAFAD2',
    'lightgreen'           : '#90EE90',
    'lightgrey'            : '#D3D3D3',
    'lightpink'            : '#FFB6C1',
    'lightsalmon'          : '#FFA07A',
    'lightseagreen'        : '#20B2AA',
    'lightskyblue'         : '#87CEFA',
    'lightslategray'       : '#778899',
    'lightsteelblue'       : '#B0C4DE',
    'lightyellow'          : '#FFFFE0',
    'lime'                 : '#00FF00',
    'limegreen'            : '#32CD32',
    'linen'                : '#FAF0E6',
    'magenta'              : '#FF00FF',
    'maroon'               : '#800000',
    'mediumaquamarine'     : '#66CDAA',
    'mediumblue'           : '#0000CD',
    'mediumorchid'         : '#BA55D3',
    'mediumpurple'         : '#9370DB',
    'mediumseagreen'       : '#3CB371',
    'mediumslateblue'      : '#7B68EE',
    'mediumspringgreen'    : '#00FA9A',
    'mediumturquoise'      : '#48D1CC',
    'mediumvioletred'      : '#C71585',
    'midnightblue'         : '#191970',
    'mintcream'            : '#F5FFFA',
    'mistyrose'            : '#FFE4E1',
    'moccasin'             : '#FFE4B5',
    'navajowhite'          : '#FFDEAD',
    'navy'                 : '#000080',
    'oldlace'              : '#FDF5E6',
    'olive'                : '#808000',
    'olivedrab'            : '#6B8E23',
    'orange'               : '#FFA500',
    'orangered'            : '#FF4500',
    'orchid'               : '#DA70D6',
    'palegoldenrod'        : '#EEE8AA',
    'palegreen'            : '#98FB98',
    'palevioletred'        : '#AFEEEE',
    'peru'                 : '#CD853F',
    'pink'                 : '#FFC0CB',
    'plum'                 : '#DDA0DD',
    'powderblue'           : '#B0E0E6',
    'purple'               : '#800080',
    'red'                  : '#FF0000',
    'rosybrown'            : '#BC8F8F',
    'royalblue'            : '#4169E1',
    'saddlebrown'          : '#8B4513',
    'salmon'               : '#FA8072',
    'sandybrown'           : '#FAA460',
    'seagreen'             : '#2E8B57',
    'seashell'             : '#FFF5EE',
    'sienna'               : '#A0522D',
    'silver'               : '#C0C0C0',
    'skyblue'              : '#87CEEB',
    'slateblue'            : '#6A5ACD',
    'slategray'            : '#708090',
    'snow'                 : '#FFFAFA',
    'springgreen'          : '#00FF7F',
    'steelblue'            : '#4682B4',
    'tan'                  : '#D2B48C',
    'teal'                 : '#008080',
    'thistle'              : '#D8BFD8',
    'tomato'               : '#FF6347',
    'turquoise'            : '#40E0D0',
    'violet'               : '#EE82EE',
    'wheat'                : '#F5DEB3',
    'white'                : '#FFFFFF',
    'whitesmoke'           : '#F5F5F5',
    'yellow'               : '#FFFF00',
    'yellowgreen'          : '#9ACD32',
    'black'                : '#000000',
    'navy'                 : '#000080',
    'mediumblue'           : '#0000CD',
    'teal'                 : '#008080',
    'darkcyan'             : '#008B8B',
    'deepskyblue'          : '#00BFFF',
    'darkturquoise'        : '#00CED1',
    'mediumspringgreen'    : '#00FA9A',
    'lime'                 : '#00FF00',
    'springgreen'          : '#00FF7F',
    'aqua'                 : '#00FFFF',
    'midnightblue'         : '#191970',
    'dodgerblue'           : '#1E90FF',
    'lightseagreen'        : '#20B2AA',
    'forestgreen'          : '#228B22',
    'seagreen'             : '#2E8B57',
    'darkslategray'        : '#2F4F4F',
    'limegreen'            : '#32CD32',
    'mediumseagreen'       : '#3CB371',
    'turquoise'            : '#40E0D0',
    'royalblue'            : '#4169E1',
    'steelblue'            : '#4682B4',
    'darkslateblue'        : '#483D8B',
    'mediumturquoise'      : '#48D1CC',
    'indigo'               : '#4B0082',
    'darkolivegreen'       : '#556B2F',
    'cadetblue'            : '#5F9EA0',
    'cornflowerblue'       : '#6495ED',
    'mediumaquamarine'     : '#66CDAA',
    'dimgray'              : '#696969',
    'slateblue'            : '#6A5ACD',
    'olivedrab'            : '#6B8E23',
    'slategray'            : '#708090',
    'lightslategray'       : '#778899',
    'mediumslateblue'      : '#7B68EE',
    'lawngreen'            : '#7CFC00',
    'chartreuse'           : '#7FFF00',
    'aquamarine'           : '#7FFFD4',
    'maroon'               : '#800000',
    'purple'               : '#800080',
    'gray'                 : '#808080',
    'grey'                 : '#808080',
    'skyblue'              : '#87CEEB',
    'lightskyblue'         : '#87CEFA',
    'blueviolet'           : '#8A2BE2',
    'darkmagenta'          : '#8B008B',
    'saddlebrown'          : '#8B4513',
    'darkseagreen'         : '#8FBC8F',
    'lightgreen'           : '#90EE90',
    'mediumpurple'         : '#9370DB',
    'darkviolet'           : '#9400D3',
    'palegreen'            : '#98FB98',
    'darkorchid'           : '#9932CC',
    'yellowgreen'          : '#9ACD32',
    'sienna'               : '#A0522D',
    'darkgray'             : '#A9A9A9',
    'lightblue'            : '#ADD8E6',
    'greenyellow'          : '#ADFF2F',
    'palevioletred'        : '#AFEEEE',
    'lightsteelblue'       : '#B0C4DE',
    'powderblue'           : '#B0E0E6',
    'firebrick'            : '#B22222',
    'darkgoldenrod'        : '#B8860B',
    'mediumorchid'         : '#BA55D3',
    'rosybrown'            : '#BC8F8F',
    'darkkhaki'            : '#BDB76B',
    'silver'               : '#C0C0C0',
    'mediumvioletred'      : '#C71585',
    'peru'                 : '#CD853F',
    'chocolate'            : '#D2691E',
    'tan'                  : '#D2B48C',
    'lightgrey'            : '#D3D3D3',
    'thistle'              : '#D8BFD8',
    'orchid'               : '#DA70D6',
    'goldenrod'            : '#DAA520',
    'crimson'              : '#DC143C',
    'gainsboro'            : '#DCDCDC',
    'plum'                 : '#DDA0DD',
    'burlywood'            : '#DEB887',
    'lightcyan'            : '#E0FFFF',
    'lavender'             : '#E6E6FA',
    'darksalmon'           : '#E9967A',
    'violet'               : '#EE82EE',
    'palegoldenrod'        : '#EEE8AA',
    'lightcoral'           : '#F08080',
    'khaki'                : '#F0E68C',
    'aliceblue'            : '#F0F8FF',
    'honeydew'             : '#F0FFF0',
    'azure'                : '#F0FFFF',
    'wheat'                : '#F5DEB3',
    'beige'                : '#F5F5DC',
    'whitesmoke'           : '#F5F5F5',
    'mintcream'            : '#F5FFFA',
    'ghostwhite'           : '#F8F8FF',
    'salmon'               : '#FA8072',
    'sandybrown'           : '#FAA460',
    'antiquewhite'         : '#FAEBD7',
    'linen'                : '#FAF0E6',
    'lightgoldenrodyellow' : '#FAFAD2',
    'oldlace'              : '#FDF5E6',
    'fuchsia'              : '#FF00FF',
    'deeppink'             : '#FF1493',
    'orangered'            : '#FF4500',
    'tomato'               : '#FF6347',
    'hotpink'              : '#FF69B4',
    'darkorange'           : '#FF8C00',
    'lightsalmon'          : '#FFA07A',
    'orange'               : '#FFA500',
    'lightpink'            : '#FFB6C1',
    'pink'                 : '#FFC0CB',
    'navajowhite'          : '#FFDEAD',
    'moccasin'             : '#FFE4B5',
    'bisque'               : '#FFE4C4',
    'mistyrose'            : '#FFE4E1',
    'blanchedalmond'       : '#FFEBCD',
    'lavenderblush'        : '#FFF0F5',
    'seashell'             : '#FFF5EE',
    'cornsilk'             : '#FFF8DC',
    'lemonchiffon'         : '#FFFACD',
    'floralwhite'          : '#FFFAF0',
    'snow'                 : '#FFFAFA',
    'yellow'               : '#FFFF00',
    'lightyellow'          : '#FFFFE0',
    'ivory'                : '#FFFFF0',
    }


# Make color name list w/o the stdList colors (in whatever order the hash delivers)
BIG_COLORNAME_LIST = stdList[:] # make a copy of stdList
BIG_COLORNAME_LIST.extend( list( set(COLOR_NAME_DICT.keys()) - set(stdList) ) )
BIG_COLOR_HEXSTR_LIST = [COLOR_NAME_DICT[key] for key in BIG_COLORNAME_LIST]

if __name__=="__main__":
    import sys

    def test_valid( c ):
        print('Testing valid %8s '%c, end='')
        print( getValidHexStr( c, "#000000") )

    for c in ['red','#3ac', '#AA0099', '55AAFF', 'goober']:
        test_valid(c)

    sys.exit()

    print( "get_best_gray_text('#121212') =",get_best_gray_text('#121212') )

    print()
    s1 = '#000000'
    for s in ['#000000','#ffffff','#669900','#cc0000']:
        print( 'contrast of',s1,'and',s,'=',w3_contrast(s1,s))

    print()
    print()
    for s in ['#FF0000','#00FF00','#0000FF']:
        hue, sat, val = colorsys.rgb_to_hsv(*scale_rgb_tuple(hex_to_rgb(s)))
        print( '%s,  h=%s,  s=%s,  v=%s'%(s,hue, sat, val))

    colorL = []
    hue, sat, val = 0., 1., 1.
    golden_ratio = 0.618033988749895
    dh = 1.0/3.0
    #hstartL = [0.0, 1.0/6.0, 1.0/12.0, 3.0/12.0, 1.0/24.0, 3.0/24.0, 5.0/24.0, 7.0/24.0]
    hstartL = [0.0, 1.0/6.0, 1.0/12.0, 3.0/12.0]
    satL = [1., .7, .4]

    for sat in satL:
        val = sat
        for hstart in hstartL:
            hue = hstart
            for j in range(3):

                c = rgb_to_hex( scale_rgb_tuple(colorsys.hsv_to_rgb(hue, sat, val), False) )
                if c not in colorL:
                    colorL.append( c )
                hue += dh

    print( colorL)

    for i in range( 2*len(BIG_COLORNAME_LIST) ):
        if i==len(BIG_COLORNAME_LIST):
            print()
            print()
        print( getColorName(i), getColorHexStr(i), end='')


