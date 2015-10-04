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

ALT_DATA = [['Altitude','Pressure','Temp R','Temp K'],
            ['feet','psia','degR','degK'],
            [1,      14.7, 518.7, 288.1667],
            [5000,  12.23, 500.8, 278.2222],
            [10000, 10.11, 483.0, 268.3333],
            [30000,  4.36, 411.8, 228.7778],
            [60000,  1.04, 390.0, 216.6667]]

ALT_DATA_WIDE = [ALT_DATA[0][:], ALT_DATA[1][:], ALT_DATA[2][:], ALT_DATA[3][:],
                 ALT_DATA[4][:], ALT_DATA[5][:], ALT_DATA[6][:]]
ALT_DATA_WIDE[0].extend(['Temp F','Temp C'])
ALT_DATA_WIDE[1].extend(['degF','degC'])
for i in range(2,7):
    ALT_DATA_WIDE[i].extend( [ALT_DATA_WIDE[i][2]-459.67, ALT_DATA_WIDE[i][3]-273.15] )

class MyTest(unittest.TestCase):

    def setUp(self):
        unittest.TestCase.setUp(self)
        self.mySprSht = SpreadSheet()

    def tearDown(self):
        unittest.TestCase.tearDown(self)
        del( self.mySprSht )

    def test_should_always_pass_cleanly(self):
        """Should always pass cleanly."""
        pass

    def test_myclass_existence(self):
        """Check that myclass exists"""
        result = self.mySprSht

        # See if the self.mySprSht object exists
        self.assertTrue(result)


    def test_save(self):
        """Check that save operates cleanly"""
        list_of_rows = ALT_DATA

        self.mySprSht.add_sheet('Alt_Data', list_of_rows)
        self.mySprSht.add_scatter( 'Alt_Plot', 'Alt_Data',
                                  title='Unittest Title', xlabel='Unittest X Axis',
                                  ylabel='Unittest Y Axis', y2label='Unittest Y2 Axis',
                                  xcol=1,
                                  ycolL=[2,3,4], ycol2L=None,
                                  showMarkerL=[1,1,1], showMarker2L=None,
                                  lineThkL=["0.1in","0.01in"],
                                  colorL=['cyan','GRaY','#69a'], #['#666666'],
                                  labelL=None, label2L=None)

        self.mySprSht.setXrange( -10000, 70000, plot_sheetname=None)
        self.mySprSht.setYrange(-100, 700)
        #self.mySprSht.add_scatter( 'Alt_Plot2', 'Alt_Data')
        #self.mySprSht.add_scatter( 'Alt_Plot3', 'Alt_Data')
        self.mySprSht.save( filename=os.path.join(here,'alt') )


    def test_save_secondary_y(self):
        """Check that save operates for a second y axis"""
        list_of_rows = ALT_DATA_WIDE

        self.mySprSht.add_sheet('Alt_Data', list_of_rows)
        self.mySprSht.add_scatter( 'Alt_Plot', 'Alt_Data',
                                  title='Unittest Title', xlabel='Unittest X Axis',
                                  ylabel='Unittest Y Axis', y2label='Unittest Y2 Axis',
                                  xcol=1,
                                  ycolL=[2,5,6], ycol2L=[3,4],
                                  lineThkL=[1,2,3],
                                  lineThk2L=[3.5, 1.5],
                                  showMarkerL=[1,1,1], showMarker2L=None,
                                  showLineL=[0,1,0], showLine2L=[0,1],
                                  #colorL=['r','g','b'], excel_colors=False, color2L=['c','dc'],
                                  labelL=None, label2L=None)

        self.mySprSht.setY2range(50, 550)
        self.mySprSht.save( filename=os.path.join(here,'alt_y2') )


    def test_logx_save(self):
        """Check that save operates cleanly"""
        list_of_rows = ALT_DATA

        self.mySprSht.add_sheet('Alt_Data', list_of_rows)
        self.mySprSht.add_scatter( 'Alt_Plot', 'Alt_Data',
                                  title='Unittest Title', xlabel='Unittest X Axis',
                                  ylabel='Unittest Y Axis', y2label='Unittest Y2 Axis',
                                  xcol=1, logx=True,
                                  ycolL=[3,4], ycol2L=None,
                                  showMarkerL=[1,1,1], showMarker2L=None,
                                  colorL=None,
                                  labelL=None, label2L=None)

        self.mySprSht.save( filename=os.path.join(here,'alt_logx') )


    def test_logy_save(self):
        """Check that save operates cleanly"""
        list_of_rows = ALT_DATA

        self.mySprSht.add_sheet('Alt_Data', list_of_rows)
        self.mySprSht.add_scatter( 'Alt_Plot', 'Alt_Data',
                                  title='Unittest Title', xlabel='Unittest X Axis',
                                  ylabel='Unittest Y Axis', y2label='Unittest Y2 Axis',
                                  xcol=1, logy=True,
                                  ycolL=[3,4], ycol2L=None,
                                  showMarkerL=[1,1,1], showMarker2L=None,
                                  colorL=None,
                                  labelL=None, label2L=None)

        self.mySprSht.save( filename=os.path.join(here,'alt_logy') )


    def test_log2y_save(self):
        """Check that save operates for a second y axis"""
        list_of_rows = ALT_DATA_WIDE

        self.mySprSht.add_sheet('Alt_Data', list_of_rows)
        self.mySprSht.add_scatter( 'Alt_Plot', 'Alt_Data',
                                  title='Unittest Title', xlabel='Unittest X Axis',
                                  ylabel='Unittest Y Axis', y2label='Unittest Y2 Axis',
                                  xcol=1, log2y=True,
                                  ycolL=[2,5,6], ycol2L=[3,4],
                                  showMarkerL=[1,1,1], showMarker2L=None,
                                  colorL=None,
                                  labelL=None, label2L=None)

        self.mySprSht.save( filename=os.path.join(here,'alt_log2y') )

    def test_all_log_save(self):
        """Check that save operates for a second y axis"""
        list_of_rows = ALT_DATA_WIDE

        self.mySprSht.add_sheet('Alt_Data', list_of_rows)
        self.mySprSht.add_scatter( 'Alt_Plot', 'Alt_Data',
                                  title='Unittest Title', xlabel='Unittest X Axis',
                                  ylabel='Unittest Y Axis', y2label='Unittest Y2 Axis',
                                  xcol=1, log2y=True, logx=True, logy=True,
                                  ycolL=[2,5,6], ycol2L=[3,4],
                                  showMarkerL=[1,1,1], showMarker2L=None,
                                  colorL=None,
                                  labelL=None, label2L=None)

        self.mySprSht.save( filename=os.path.join(here,'alt_all_log') )



if __name__ == '__main__':
    # Can test just this file from command prompt
    #  or it can be part of test discovery from nose, unittest, pytest, etc.
    unittest.main()

