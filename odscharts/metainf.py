"""
Make changes to META-INF data and Object 1-N subdirectories
"""

from lxml import etree as ET

def NS( shortname, nsmap): 
    """remove namespace shortcuts from name"""
    
    sL = shortname.split(':')
    if len(sL)!=2:
        return shortname
        
    return '{%s}'%nsmap[sL[0]] + sL[1]


def add_ObjectN( N, metainf_manifest_xml_obj):
    """Add Object manifest:file-entry object to manifest.xml in META-INF of ods file"""
    
    M = metainf_manifest_xml_obj.getroot() # just to reduce line lengths
    nsmap = M.nsmap
    #print( 'nsmap=',nsmap)
    
    # put subdirectory object into manifest
    obj_name = 'Object %i'%N
    attribD = {NS('manifest:full-path', nsmap):obj_name+'/', 
               NS('manifest:media-type', nsmap):"application/vnd.oasis.opendocument.chart"}
                   
    mfe = ET.Element(NS('manifest:file-entry', nsmap), attrib=attribD, nsmap=nsmap)
    M.insert(N ,mfe) # place in order
    
    # put xml files into subdirectory
    attribD = {NS('manifest:full-path', nsmap):obj_name+'/content.xml', 
               NS('manifest:media-type', nsmap):"text/xml"}
    mfe = ET.Element(NS('manifest:file-entry', nsmap), attrib=attribD, nsmap=nsmap)
    M.insert(N*3 ,mfe) # place in order
    
    attribD = {NS('manifest:full-path', nsmap):obj_name+'/styles.xml', 
               NS('manifest:media-type', nsmap):"text/xml"}
    mfe = ET.Element(NS('manifest:file-entry', nsmap), attrib=attribD, nsmap=nsmap)
    M.insert(N*3 ,mfe) # place in order
    
    