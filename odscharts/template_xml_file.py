import sys
from collections import OrderedDict
import io

if sys.version_info < (3,):
    import ElementTree_27OD as ET
else:
    import ElementTree_34OD as ET


# get StringIO for either python 2.x or 3.x
try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO

import re
from copy import deepcopy

header_re = re.compile( '\<\?.*\?\>', flags=re.MULTILINE | re.UNICODE )

class TemplateXML_File(object):
    
    def __init__(self, xml_file_name_or_src):
        """
        Read and parse xml file using modified version of standard python 
        xml.etree.ElementTree.
        
        xml_file_name_or_src can be a file name like: "content.xml" OR 
        can be xml source.
        """
        
        if xml_file_name_or_src.endswith('.xml') and len(xml_file_name_or_src)<256:
            self.xml_file_name_or_src = xml_file_name_or_src
            
            fInp = io.open(xml_file_name_or_src, 'rt', encoding='utf-8')
            xml_src = fInp.read()
            fInp.close()
        else:
            self.xml_file_name_or_src = None
            xml_src = xml_file_name_or_src

        self.xml_header = '' # Assume no header unless found at head of file
        match = header_re.match(xml_src)
        if match:
            #print( 'Found XML Header: ' + match.group(0) )
            self.xml_header = match.group(0) # will need \n when serialized

        # ns entries like: ('urn:oasis:names:tc:opendocument:xmlns:table:1.0', u'table')
        self.nsOD = OrderedDict()
        
        # rev_ns entries like: (u'table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0')
        self.rev_nsOD = OrderedDict()
        
        # qname entries like: ('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}document-content', u'office:document-content')
        self.qnameOD = OrderedDict()
                
        events = ("start", "end", "start-ns", "end-ns")
        context = ET.iterparse(StringIO(xml_src), events=events)

        for event, elem in context:
            if event=="start":
                #print('type elem.tag =', type(elem.tag))
                #print('     type("}") =',type("}"))
                sL = elem.tag.split('}')
                if len(sL) == 2:
                    name = sL[1]
                    uri = sL[0][1:]
                    self.qnameOD[elem.tag] = '%s:%s'%(self.nsOD[uri], name)
                
                for qname,v in elem.attrib.items():
                    sL = qname.split('}')
                    if len(sL) == 2:
                        name = sL[1]
                        uri = sL[0][1:]
                        self.qnameOD[qname] = '%s:%s'%(self.nsOD[uri], name)
            if event=="start-ns":
                self.nsOD[elem[1]] = elem[0]
                self.rev_nsOD[elem[0]] = elem[1] # reverse index into namespace

        self.context = context
        #self.root = ET.ElementTree( context.root )
        self.root = context.root
        
        self.parentD = {} # index=child Element object, value=parent Element object
        self.depthD = {}  # index=Element object, value = depth in xml tree
        self.max_depth = 0
        self.short_pathD = {} # index=Element, value = short name (like: "ns0:name1/ns1:xyz/ns3:abc")
        # After building tree, create self.parentD for all Elements
        self.parentD[self.root] = None
        self.depthD[self.root] = 0
        self.short_pathD[self.root] = self.qnameOD[ self.root.tag ] # no calc req'd... just = qname
        
        temp_short_path_counterD = {} # just used here to help count occurances of short path
        self.short_path_counterD = {} # index=Element, value=short path counter value
        self.short_path_parent_counterD = {} #  index=Element, value=parent's short path counter value
        self.short_path_counterD[self.root] = 1 # 1st (and only) occurance
        self.short_path_parent_counterD[self.root] = 1 # 1st (and only) occurance
        
        for parent in self.root.iter():
            try:
                for child in parent.getchildren():
                    self.parentD[child] = parent
                    self.depthD[child] = self.depthD[parent] + 1
                    self.max_depth = max(self.max_depth, self.depthD[child])
                    
                    short_path = self.get_short_path( child )
                    self.short_pathD[child] = short_path
                    
                    temp_short_path_counterD[(parent,short_path)] = temp_short_path_counterD.get((parent,short_path), 0) + 1
                    self.short_path_counterD[child] = '%s,%s'%(self.short_path_counterD[parent], 
                                                       temp_short_path_counterD[(parent,short_path)])
                    self.short_path_parent_counterD[child] = '%s'%self.short_path_counterD[parent]
                    
            except:
                print( 'NOTICE: No children for:', parent )

    def get_short_path(self, elem):
        
        sL = [ self.qnameOD[ elem.tag ] ]
        while self.parentD[elem] is not None:
            elem = self.parentD[elem]
            sL.append( self.qnameOD[ elem.tag ] )
        sL.reverse()
        return "/".join(sL)

    def getroot(self):
        return self.root
    
    def getparent(self, child):
        return self.parentD[child]
        
    def remove_child(self, child, parent):
        parent.remove( child )
        del self.parentD[child]
    
    def add_child(self, child, parent):
        self.parentD[child] = parent
        parent.append( child )
    
    def add_children(self, childrenL, parent):
        for child in childrenL:
            self.add_child( child, parent )

    def tostring(self):

        class dummy:
            pass
        xml_dataL = []
        if self.xml_header:
            xml_dataL = [self.xml_header + '\n']

        dummy_file = dummy()
        dummy_file.write = xml_dataL.append

        if sys.version_info < (3,):
            ET._serialize_xml(dummy_file.write, self.root, "utf-8", self.qnameOD, self.nsOD)
        else:
            short_empty_elements = True # use short format for empty elements
            ET._serialize_xml(dummy_file.write, self.root, self.qnameOD, self.nsOD, short_empty_elements)
        
        return "".join(xml_dataL)
        
    def write(self, out_file_name):
        #fOut = open(out_file_name, "w")
        fOut = io.open(out_file_name, 'wt', encoding='utf-8')
        fOut.write( self.tostring() )
        fOut.close()
        
        
    def find(self, path, elem_obj=None):
        """Find Element within root OR elem_obj.  
           path format is: 'office:body/office:spreadsheet'
        """
        if elem_obj:
            return elem_obj.find( path, self.rev_nsOD )
        else:
            return self.root.find( path, self.rev_nsOD )
        
    def findall(self, path, elem_obj=None):
        """Find All Element objects within root OR elem_obj.  
           path format is: 'office:body/office:spreadsheet'
        """
        if elem_obj:
            return elem_obj.findall( path, self.rev_nsOD )
        else:
            return self.root.findall( path, self.rev_nsOD )


    def NS(self, path_or_tag ): 
        """force into tag format like: '{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table' """
        if path_or_tag.startswith('{'):
            return path_or_tag
            
        sL = path_or_tag.split(':')
        if len(sL)!=2:
            return path_or_tag
            
        return '{%s}%s'%( self.rev_nsOD[sL[0]], sL[1] )

    def NS_attrib(self, attD ):
        D = OrderedDict()
        for key,val in attD.items():
            D[ self.NS(key) ] = val
        return D
        
    def new_elem(self, name, attribOD=None):
        """
        Create a new Element object.
        name format can be 'table:table' OR '{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table'
        (i.e. can be path format OR tab format)
        attribOD is an OrderedDict in order to preserve attribute order.
        """
        
        tag = self.NS( name )
        #print( 'tag =',tag )


        if attribOD:
            OD = self.NS_attrib( attribOD )
            new_elem = ET.Element(tag, attrib=OD)
        else:
            new_elem = ET.Element(tag)

        sL = new_elem.tag.split('}')
        if len(sL) == 2:
            name = sL[1]
            uri = sL[0][1:]
            self.qnameOD[new_elem.tag] = '%s:%s'%(self.nsOD[uri], name)
        
        for qname,v in new_elem.attrib.items():
            sL = qname.split('}')
            if len(sL) == 2:
                name = sL[1]
                uri = sL[0][1:]
                self.qnameOD[qname] = '%s:%s'%(self.nsOD[uri], name)
        
        return new_elem

if __name__ == "__main__":
    
    TFile = TemplateXML_File(r'D:\temp\open_office\content.xml')
    
    #TFile.write( r'D:\temp\open_office\content_v11.xml' )
    
    ss = TFile.find( 'office:body/office:spreadsheet' )
    print( 'ss =', ss )
    print( 'ss.tag =', ss.tag )
    print()
    
    ss2 = TFile.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}spreadsheet')
    print( 'ss2 =',ss2,'  (Should be None since can NOT search on tag)' )
    
    #print( TFile.findall( 'table:table', elem_obj=ss ) )
    print
    #print( TFile.qnameOD.items()[:2] )
    
    newtab1 = TFile.new_elem( '{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table', attribOD=None)
    newtab2 = TFile.new_elem( 'table:table', attribOD=None)
    
    print( 'newtab1 =',newtab1 )
    print( 'newtab2 =',newtab2 )
    
    newtab3 = TFile.new_elem( 'table:table', attribOD={'table:name':'Sheet1'})
    print( 'newtab3 =',newtab3 )
    print( 'newtab3 =',newtab3.items() )
