# Python 2 and 3
from __future__ import unicode_literals
from __future__ import absolute_import
from __future__ import print_function

"""
Represents an xml file as a tree of set objects
"""
from collections import OrderedDict
from odscharts.template_xml_file import TemplateXML_File
from copy import copy, deepcopy

def make_attr_sets_from_setrepL( e_srL ):
    """
    Make a list of sets of just attribute names (not values)
    Used to see if two elements share a large number of attributes
    """
    
    attr_sL = [] # list of sets
    for e_sr in e_srL: # look at each set in setrep list
        aL = []
        for item in e_sr:
            sL = item.split('|')
            if len(sL)==3 and sL[0]=='ATTR':
                aL.append( sL[1] )
        attr_sL.append( set(aL) )
    return attr_sL # a list of sets containing attribute names

def get_set_member( setrep, startswith='COUNT' ):
    """Looks through set and returns a list of set members satisfying startswith"""
    
    returnL = []
    for s in setrep:
        if s.startswith( startswith ):
            returnL.append( s )
    return returnL

def strip_short_path( short_path, num_beg=0, num_end=0):
    """
    Strip number of positions from short_path as indicated by num_beg and num_end
    """        
    sL = short_path.split('/')
    return '/'.join(sL[num_beg: len(sL) - num_end])
    
    
    

class SetRep( TemplateXML_File ):
    
    def __init__(self, xml_file_name_or_src ):
        """
        Initialize from the TemplateXML_File class
        
        xml_file_name_or_src can be a file name like: "content.xml" OR 
        can be xml source.        
        """
        
        if not xml_file_name_or_src.strip():
            xml_file_name_or_src = '<office:document-styles  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ></office:document-styles>'
        
        TemplateXML_File.__init__(self, xml_file_name_or_src)
        
        self.e_setrepOD = OrderedDict()  # index=Element object, value=set representation
        self.elem_from_idD = {} # index=id(e_setrep), value=Element object from which it's made
        
        # The lists within depth_elemD maintain the order from original file
        self.depth_elemD = {} # index=n depth, value=list of Element objects at that depth
        for n in range(self.max_depth +1):
            self.depth_elemD[n] = [] # init the depth lists
        
        def build_elem_setrep( elem ):
            descL = [ "PATH|"+self.short_pathD[elem] ]
            
            # COUNT is only used to help put nodes in correct tree limb
            #descL.append(  'COUNT|'+'%s'%self.short_path_counterD[elem] )
            descL.append(  'PARENT|'+'%s'%self.short_path_parent_counterD[elem] )
            
            descL.append( 'TEXT|'+'%s'%elem.text )
            descL.append( 'TAIL|'+'%s'%elem.tail )
            
            for qname,v in elem.attrib.items():
                short_name = self.qnameOD[qname]
                descL.append( 'ATTR|'+short_name+'|'+v )
            
            for child in elem.getchildren():
                descL.append( 'CHILD|'+self.short_pathD[child] )
            
            s = set( descL )
            self.elem_from_idD[ id(s) ] = elem # use python object id of set to get Element
            return s

        self.e_setrepOD[self.root] = build_elem_setrep( self.root )
        self.depth_elemD[0].append( self.root )
        
        for parent in self.root.iter():
            for child in parent.getchildren():
                self.e_setrepOD[child] = build_elem_setrep( child )
                self.depth_elemD[ self.depthD[child] ].append( child )

    

    def find_diff(self, setrep2):
        """Find the differences between SetRep objects self and setrep2"""
        
        print self.e_setrepOD[self.root] == setrep2.e_setrepOD[setrep2.root], self.short_pathD[self.root]
        
        if self.max_depth == setrep2.max_depth:
            print 'Both SetRep Objects have the Same Depth'
        else:
            print 'WARNING... Different Depths: self=%i,  setrep2=%i'%(self.max_depth, setrep2.max_depth)

        not_found_in1LL = []  # List of Lists, 1st index is depth, gives list of not found
        not_found_in2LL = []  #   at that depth
        total_nf1 = 0
        total_nf2 = 0

        for n in range(self.max_depth +1):
            nfL1 = [] # list of not-found at this depth
            nfL2 = [] # list of not-found at this depth
            
            
            print
            print '='*20,'Depth =%i'%n,'='*20
            eL1 = self.depth_elemD[n]
            eL2 = setrep2.depth_elemD[n]
            
            e_srL1 = [self.e_setrepOD[e] for e in eL1]
            e_srL2 = [setrep2.e_setrepOD[e] for e in eL2]
            
            # Look for non-found elements in self
            for i1,e_sr1 in enumerate(e_srL1):
                if e_sr1 in e_srL2:
                    print '.',
                else:
                    print '?',
                    #print 'NOT FOUND',eL1[i1]
                    nfL1.append( eL1[i1] )
                    total_nf1 += 1
            print
            # Look for non-found elements in setrep2
            for i2,e_sr2 in enumerate(e_srL2):
                if e_sr2 in e_srL1:
                    print '.',
                else:
                    print '?',
                    #print '2 NOT FOUND',eL2[i2]
                    nfL2.append( eL2[i2] )
                    total_nf2 += 1
                    
            not_found_in1LL.append(nfL1)
            not_found_in2LL.append(nfL2)
            
            if nfL1 or nfL2:
                print
                print 'at depth=%i, not found 1=%i, not found 2=%i, #element diff=%i'%\
                       (n, len(nfL1), len(nfL2), len(e_srL1)-len(e_srL2))
        
        print
        print '='*20,'Not Found 1 = %i'%total_nf1,'='*5,'Not Found 2 = %i'%total_nf2,'='*20
        
        total_diff = total_nf1 + total_nf2
        return total_diff
        

    def find_changes_to_make_equal(self, setrep2):
        """Find the changes required in self ElementTree to make it equal
           to the ElementTree from setrep2
        
           Find the Element objects in self and setrep2 that match.
           Matching Element objects may still have attrib, child or text differences,
           but they appear in each document at the same point and perform the same task.
           They can match ONLY IF they are at the same depth with same PATH
           
           Find Element objects that are close, but require changes
           Find the Element objects that need to be added and those that need 
           to be removed
        """
        
        self.setrep2 = setrep2 # save SetRep of ElementTree used to find changes to make equal
        
        # holds either matching elem in setrep2 or None
        self.matching_elemD = {} # index=self.elem, value=setrep2.elem
        self.rev_matching_elemD = {} # index=setrep2.elem, value=self.elem
        
        #  match holds tuples of (elem, elem2)
        self.match_elemLL = [] # list of lists by depth of partially matching Element Objects

        # unmatched holds tuples of Element and SetRep objects
        self.unmatched_elemLL = [] # list of lists by depth of unmatched Element objects in self
        self.unmatched_elemLL2 = [] # list of lists by depth of unmatched Element objects in setrep2
        
        # partial match holds tuples of (elem, elem2)
        self.partial_match_elemLL = [] # list of lists by depth of partially matching Element Objects
        
        #multi_match_elemLL = [] # list of lists of multiple matches

        for n in range(self.max_depth +1): # iterate over depth values
            print '='*40 + '> Working Depth = %i'%n
            elemL = self.depth_elemD[n]

            self.match_elemLL.append( [] )
            self.unmatched_elemLL.append( [] )
            self.unmatched_elemLL2.append( [] )
            self.partial_match_elemLL.append( [] )
            #multi_match_elemLL.append( [] )

            eL2  = setrep2.depth_elemD[n]
            e_srL2 = [setrep2.e_setrepOD[e2] for e2 in eL2] # build list of SetRep objects in setrep2 at depth
            for e2 in eL2:
                self.rev_matching_elemD[e2] = None # init to None
            
            for elem in elemL:
                self.matching_elemD[elem] = None # initialize to None (perhaps reset later)
                e_sr = self.e_setrepOD[elem]
                
                # Count the number of EXACT MATCHES for Element self setreps in setrep2
                match_count = e_srL2.count( e_sr )
                # These are complete and TOTAL matches at EVERY location
                if match_count == 0:
                    self.unmatched_elemLL[-1].append( (elem, e_sr) )
                elif match_count >= 1: # look for first match
                    i2 = e_srL2.index( e_sr )
                    e_sr2 = e_srL2[i2]
                    e_srL2.pop(i2)  # remove the matched Element from e_srL2
                    
                    e2 = setrep2.elem_from_idD[ id(e_sr2) ]
                    self.matching_elemD[elem] = e2
                    self.rev_matching_elemD[e2] = elem
                    self.match_elemLL[-1].append( (elem, e2) )
                #else:
                #    multi_match_elemLL[-1].append( (elem, e_sr) )
            
            # =========== Now look at partial matches =====================
            # With complete and total matches handled, now look for partial matches
            #   (PATH must match)
            # Find best match for members of unmatched_elemLL (at this depth)
            if e_srL2:
                attr_sL2 = make_attr_sets_from_setrepL( e_srL2 )
                
                unLL = copy( self.unmatched_elemLL[-1] )
                for elem, e_sr in unLL:
                    i2 = self.find_best_setrep_match( elem, e_sr, e_srL2, attr_sL2 )
                    if i2>=0:
                        e_sr2 = e_srL2.pop(i2) # removes e_sr2 from from list of sets of elements
                        attr_sL2.pop(i2) # throw away attr set at i2
                        
                        e2 = setrep2.elem_from_idD[ id(e_sr2) ]
                        
                        self.partial_match_elemLL[-1].append( (elem, e2) )
                        print 'Found best partial match', e2
                        print '     ',get_set_member( e_sr, startswith='PARENT' ),get_set_member( e_sr2, startswith='PARENT' )
                        
                        # While not a perfect match, it is a partial match
                        self.unmatched_elemLL[-1].remove( (elem, e_sr) )
                
            # save the unmatched elements from setrep2
            e2L = [setrep2.elem_from_idD[ id(e_sr2) ] for e_sr2  in e_srL2]
            self.unmatched_elemLL2[-1].extend( e2L )
            
            #for elem, e_sr in self.unmatched_elemLL[-1]:
            #    print '    UN:',elem
            #for e_sr2 in e_srL2:
            #    e2 = setrep2.elem_from_idD[ id(e_sr2) ]
            #    print '   UN2:',elem
                
            #print '%2i) unmatch=%2i,  multi=%2i'%(n, len(self.unmatched_elemLL[-1]), len(multi_match_elemLL[-1]))
            
            
            print '%2i) delete=%2i copy=%2i modify=%2i, exact=%2i'%(n, 
                    len(self.unmatched_elemLL[-1]), 
                    len(self.unmatched_elemLL2[-1]), 
                    len(self.partial_match_elemLL[-1]), 
                    len(self.match_elemLL[-1]) )
            
            if self.unmatched_elemLL[-1]:
                print '==>Delete These Element objects from self'
                for elem, e_sr in self.unmatched_elemLL[-1]:
                    print '.....',elem
                    print '.......',e_sr  #self.short_pathD[elem]
                    
            if self.unmatched_elemLL2[-1]:
                print '==>Copy these Elements from setrep2'
                for e2, e_sr2 in zip( e2L, e_srL2 ):
                    print '.....',e2
                    print '.......',e_sr2 #setrep2.short_pathD[e2]

                    
            if self.partial_match_elemLL[-1]:
                print '==>Modify these elements in self to match setrep2'
                for e, e2 in self.partial_match_elemLL[-1]:
                    print '...1...',self.short_pathD[e]
                    print '....2..',setrep2.short_pathD[e2]

    def get_attrib_assignment_commands(self, elem, e2):
        """
        Find elem and e2 in each ElementTree, and generate commands to make 
        attrib the same in self as in setrep2
        """
        commandL = []
        
        short_path = strip_short_path( self.short_pathD[elem], num_beg=1, num_end=0)
                
        #print 'short_path =',short_path
        for i, e in enumerate( self.findall( short_path ) ):
            #print 'findall result:',e
            if e==elem:
                commandL.append( '    elem = self.content_xml_obj.findall("%s")[%i]'%(short_path, i) )
        
        for a in e2.attrib:
            if e2.get(a) != elem.get(a):
                commandL.append( '    elem.set("%s", "%s")'%(a, e2.get(a)) )
        
        for a in elem.attrib:
            if a not in e2.attrib:
                commandL.append( '    del elem.attrib["%s"]'%a )
        
        return commandL
        

    def get_e2_copy_commands(self, e2):
        """
        Find e2 in setrep2 ElementTree, and generate commands to make 
        a copy of it in self
        """
        commandL = []
        
        parent2 = self.setrep2.parentD[e2]
        
        if parent2 in self.e2_object_nameD:
            
            parent_name = self.e2_object_nameD[parent2]
            e2_name = 'new_elem_%i'%( len(self.e2_object_nameD) + 1 )
            self.e2_object_nameD[e2] = e2_name
            
            commandL.append( '    %s = ET.SubElement(%s,"%s", attrib=%s)'%(e2_name, parent_name, e2.tag, repr(e2.attrib)) )
            if e2.text:
                commandL.append( '    %s.text = "%s")'%(e2_name, e2.text) )
            
        else:
            # Assume same parent findall index ONLY for top level attachment point
            e_sr2 = self.setrep2.e_setrepOD[e2]
            parent_str = get_set_member( e_sr2, startswith='PARENT' )
            #print 'parent_str =',parent_str
            #print '   e2 attr =',get_set_member( e_sr2, startswith='ATTR' ),get_set_member( e_sr2, startswith='TEXT' )

            short_path_parent2 = strip_short_path( self.setrep2.short_pathD[e2], num_beg=1, num_end=1)
            parent2 = self.setrep2.parentD[e2]
            iparent = -1
            for i, p in enumerate( self.setrep2.findall( short_path_parent2 ) ):
                #print 'findall result:',e
                if p==parent2:
                    iparent = i # assume that parent in setrep2 is in same location as self
            #print '  Found e2 parent at index =',iparent
            #print '     ',short_path_parent2
                    
            parent_self = self.findall( short_path_parent2 )[iparent]
            
            commandL.append( '    parent = self.content_xml_obj.findall("%s")[%i]'%(short_path_parent2, iparent) )
            
            e2_name = 'new_elem_%i'%( len(self.e2_object_nameD) + 1 )
            self.e2_object_nameD[e2] = e2_name
            
            commandL.append( '    %s = ET.SubElement(parent,"%s", attrib=%s)'%(e2_name, e2.tag, repr(e2.attrib)) )
            if e2.text:
                commandL.append( '    %s.text = "%s")'%(e2_name, e2.text) )
            
        
        #print 'Short Paths:', short_path2, short_path_parent2
        
        return commandL

    def get_remove_command(self, elem):
        """
        Find elem and parent in ElementTree, and generate commands to remove elem
        """
        commandL = []
        
        short_path = strip_short_path( self.short_pathD[elem], num_beg=1, num_end=0)
        
        for i, e in enumerate( self.findall( short_path ) ):
            if e==elem:
                commandL.append( '    elem = self.content_xml_obj.findall("%s")[%i]'%(short_path, i) )
                
        parent = self.parentD[elem]
        short_path_parent = strip_short_path( self.short_pathD[elem], num_beg=1, num_end=1)
        
        for i, p in enumerate( self.findall( short_path_parent ) ):
            if p==parent:
                commandL.append( '    parent = self.content_xml_obj.findall("%s")[%i]'%(short_path_parent, i) )
                
        commandL.append( '    parent.remove( elem )' )
        return commandL
        
        
    def create_ET_commands_to_make_equal(self):
        """
        Assume that "find_changes_to_make_equal" has already been called.
        
        Now create ET commands that will create setrep2 from self.
        """
        print '^'*77
        print 'v'*77
        
        commandL = [] # list of commands that will make self equal to setrep2
        
        # 1) Unchanged Element objects in match_elemLL are left alone.
        # 2) Modify Elements in partial_match_elemLL
        
        for n in range(self.max_depth +1): # iterate over depth values
            if self.partial_match_elemLL[n]:
                print len(self.partial_match_elemLL[n]),'Modifications self at depth =',n
            else:
                print 'No Modifications of self at depth =',n
                
            for elem, e2 in self.partial_match_elemLL[n]:
                commandL.extend( self.get_attrib_assignment_commands(elem, e2) )
                
        # 3) Copy Elements from setrep2 that are in unmatched_elemLL2
        
        self.e2_object_nameD = {} # index=e2, value=name used for create command
        
        for n in range(self.max_depth +1): # iterate over depth values
            if self.unmatched_elemLL2[n]:
                print len(self.unmatched_elemLL2[n]),'Copies of e2 at depth =',n
            else:
                print 'No Copying of e2 at depth =',n
                
            for e2 in self.unmatched_elemLL2[n]:
                commandL.extend( self.get_e2_copy_commands(e2) )
        
        # 4) Delete Any remaining self Elements in unmatched_elemLL
        LL = copy( self.unmatched_elemLL )
        LL.reverse()  # delete from the bottom up
        for L in LL:
            if L:
                for elem, e_sr in L:
                    commandL.extend( self.get_remove_command(elem) )

        
        print '\n'.join(commandL)
        
        

    def find_best_setrep_match(self, elem, e_sr, e_srL2, attr_sL2 ):
        """
        Find the index of the first, best match of element setrep in the list.
        Calc a score for each member of list, e_srL.
        If all scores are zero, return -1 as index of best match.
        """
        attr_s = make_attr_sets_from_setrepL( [e_sr] )[0]
        
        path = "PATH|"+self.short_pathD[elem]
        scoreL = []
        max_score = 0
        for e_sr2, attr_s2 in zip(e_srL2, attr_sL2):
            if path in e_sr2:
                score = 0
                
                big_set = e_sr | e_sr2
                size_big_set = float( len(big_set) ) + 1.0 # make sure no divide by zero
                score += len( e_sr & e_sr2 ) / size_big_set
                
                big_set = attr_s | attr_s2
                size_big_set = float( len(big_set) ) + 1.0 # make sure no divide by zero
                score += len( attr_s & attr_s2 ) / size_big_set
                #print 'ATTR Score =',len( attr_s & attr_s2 ) / size_big_set
                
                if get_set_member( e_sr, startswith='PARENT' ) == get_set_member( e_sr, startswith='PARENT' ):
                    score += 1.0
                
                scoreL.append( score )
                max_score = max(max_score, score)
            else:
                scoreL.append( 0 )
                #print 'PATH MISMATCH'
                #print path
        
        #print path,scoreL.index( max_score ), scoreL
        
        if e_srL2 and max_score>0:
            i = scoreL.index( max_score ) # index in e_srL2 of best match to elem
            return i
        else:
            return -1

if __name__ == "__main__":
    import sys

    
    fileB = r'D:\py_proj_2015\ss_format\alt_chart_3series\styles.xml'
    fileA = r'D:\py_proj_2015\ss_format\alt_chart_2series\styles.xml'
    
    fileB = r'D:\py_proj_2015\ss_format\alt_chart_3series\meta.xml'
    fileA = r'D:\py_proj_2015\ss_format\alt_chart_2series\meta.xml'

    fileB = r'D:\py_proj_2015\ss_format\alt_chart_3series\Object 1\content.xml'
    fileA = r'D:\py_proj_2015\ss_format\alt_chart_2series\Object 1\content.xml'

    fileB = r'D:\py_proj_2015\ss_format\alt_chart_3series\Object 1\styles.xml'
    fileA = r'D:\py_proj_2015\ss_format\alt_chart_2series\Object 1\styles.xml'

    fileB = r'D:\py_proj_2015\ss_format\alt_chart_3series\content.xml'
    fileA = r'D:\py_proj_2015\ss_format\alt_chart_2series\content.xml'
    
    #fileB = fileA

    srA = SetRep(fileA)
    srB = SetRep(fileB)
    print 'Max Depth =',srA.max_depth
    print '0',srA.short_pathD[srA.root]
    for elem in srA.root:
        depth = srA.depthD[elem]
        print '    '*depth,depth,srA.short_pathD[elem]

    print  '+'*55
    
    srA.find_changes_to_make_equal( srB )
    
    srA.create_ET_commands_to_make_equal()
    
    
    #sys.exit()
    print '#'*75
    print ' '*10,'Following is Results from "find_diff" method.'
    print '#'*75
    srA.find_diff( srB )
    #print srA.e_setrepOD[srA.root] - srB.e_setrepOD[srB.root]
    
    #for k,v in srA.e_setrepOD.items():
    #    print
    #    print v
    
 