import unittest
# import unittest2 as unittest # for versions of python < 2.7

"""
        Method                  Checks that
self.assertEqual(a, b)           a == b   
self.assertNotEqual(a, b)        a != b   
self.assertTrue(x)               bool(x) is True  
self.assertFalse(x)              bool(x) is False     
self.assertIs(a, b)              a is b
self.assertIsNot(a, b)           a is not b
self.assertIsNone(x)             x is None 
self.assertIsNotNone(x)          x is not None 
self.assertIn(a, b)              a in b
self.assertNotIn(a, b)           a not in b
self.assertIsInstance(a, b)      isinstance(a, b)  
self.assertNotIsInstance(a, b)   not isinstance(a, b)  

See:
      https://docs.python.org/2/library/unittest.html
         or
      https://docs.python.org/dev/library/unittest.html
for more assert options
"""

import sys, os

here = os.path.abspath(os.path.dirname(__file__)) # Needed for py.test
up_one = os.path.split( here )[0]  # Needed to find odscharts development version
if here not in sys.path[:2]:
    sys.path.insert(0, here)
if up_one not in sys.path[:2]:
    sys.path.insert(0, up_one)

from odscharts.spreadsheet import SpreadSheet

class MyTest(unittest.TestCase):

    def setUp(self):
        unittest.TestCase.setUp(self)
        self.myclass = SpreadSheet()

    def tearDown(self):
        unittest.TestCase.tearDown(self)
        del( self.myclass )

    def test_should_always_pass_cleanly(self):
        """Should always pass cleanly."""
        pass

    def test_myclass_existence(self):
        """Check that myclass exists"""
        result = self.myclass

        # See if the self.myclass object exists
        self.assertTrue(result)
    

    def test_save(self):
        """Check that save operates cleanly"""
        list_of_rows = [['Altitude','Pressure','Temperature','Temperature'], 
                        ['feet','psia','degR','degK'], 
                        [0, 14.7, 518.7, 288.1667], [5000, 12.23, 500.8, 278.2222], 
                        [10000, 10.11, 483.0, 268.3333], [30000, 4.36, 411.8, 228.7778],
                        [60000, 1.04, 390, 216.6667]]
        
        self.myclass.add_sheet('Alt_Data', list_of_rows)
        self.myclass.add_scatter( 'Alt_Plot', 'Alt_Data',
                                  title='Unittest Title', xlabel='Unittest X Axis', 
                                  ylabel='Unittest Y Axis', y2label='Unittest Y2 Axis',
                                  xcol=1,
                                  ycolL=[2,3,4], ycol2L=None,
                                  labelL=None, label2L=None)
                                  
        #self.myclass.add_scatter( 'Alt_Plot2', 'Alt_Data')
        #self.myclass.add_scatter( 'Alt_Plot3', 'Alt_Data')
        
        
        self.myclass.save( filename=os.path.join(here,'performance'), debug=False)
        


if __name__ == '__main__':
    # Can test just this file from command prompt
    #  or it can be part of test discovery from nose, unittest, pytest, etc.
    unittest.main()

