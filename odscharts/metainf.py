"""
Make changes to META-INF data and Object 1-N subdirectories
"""

import sys
if sys.version_info < (3,):
    import ElementTree_27OD as ET
else:
    import ElementTree_34OD as ET



def add_ObjectN( N, metainf_manifest_xml_obj):
    """Add Object manifest:file-entry object to manifest.xml in META-INF of ods file"""
    
    NS = metainf_manifest_xml_obj.NS
    
    M = metainf_manifest_xml_obj.getroot() # just to reduce line lengths
    
    # put subdirectory object into manifest
    obj_name = 'Object %i'%N
    attribD = {NS('manifest:full-path'):obj_name+'/', 
               NS('manifest:media-type'):"application/vnd.oasis.opendocument.chart"}
                   
    mfe = ET.Element(NS('manifest:file-entry'), attrib=attribD)
    M.insert(N ,mfe) # place in order
    
    # put xml files into subdirectory
    attribD = {NS('manifest:full-path'):obj_name+'/content.xml', 
               NS('manifest:media-type'):"text/xml"}
    mfe = ET.Element(NS('manifest:file-entry'), attrib=attribD)
    M.insert(N*3 ,mfe) # place in order
    
    attribD = {NS('manifest:full-path'):obj_name+'/styles.xml', 
               NS('manifest:media-type'):"text/xml"}
    mfe = ET.Element(NS('manifest:file-entry'), attrib=attribD)
    M.insert(N*3 ,mfe) # place in order
    
    