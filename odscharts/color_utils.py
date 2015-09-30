import colorsys

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
    
if __name__=="__main__":
    
    print "get_best_gray_text('#121212') =",get_best_gray_text('#121212')
    
    print
    s1 = '#000000'
    for s in ['#000000','#ffffff','#669900','#cc0000']:
        print 'contrast of',s1,'and',s,'=',w3_contrast(s1,s)
        
    print 
    print
    for s in ['#FF0000','#00FF00','#0000FF']:
        hue, sat, val = colorsys.rgb_to_hsv(*scale_rgb_tuple(hex_to_rgb(s)))
        print '%s,  h=%s,  s=%s,  v=%s'%(s,hue, sat, val)
        
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
            
    print colorL
            
    
    
        
        