.. 2015-10-03 sonofeft 0a7332442bb149c2a6c09bd42dc13e3ce3f39091
   Maintain spacing of "History" and "GitHub Log" titles

History
=======

GitHub Log
----------

* Oct 03, 2015
    - (by: sonofeft) 
        - refactored so add_curve would work.
        - Name Refactoring
            Started add_curves_to_chart but I need to refactor that out and make it
            build-on-save
* Oct 02, 2015
    - (by: sonofeft) 
        - start version 0.0.8
            Cleaned up Excel default colors
        - Changed to Beta
            Added custom color palette with excel_colors flag.
            Starting BMI example
        - "Click Images" tweek to docs
        - Added field name descriptions to doc string of spreadsheet.py
        - Added platform independent os.startfile
        - Version 0.0.6
            Took docstring of _version.py out of automodule sphinx doc
        - Screen Shots added to docs
        - more doc updates

* Oct 02, 2015
    - (by: sonofeft) 
        - made quickstart doc page
        - version 0.0.5
            Updated MANIFEST.in, removed debug print statements, changed default
            colorL handling, cleaned up examples
* Oct 01, 2015
    - (by: sonofeft) 
        - Set Axis Range added for X, Y and Y2
        - Added Xmin Xmax to charts
        - Added log scale to secondary Y
        - added log Y Axis
        - Added logx axis
        - Added units to plot
        - Added Secondary Y Axis
* Sep 30, 2015
    - (by: sonofeft) 
        - version 0.0.4, added Excel default colors
* Sep 29, 2015
    - (by: sonofeft) 
        - Added Examples, fixed python 3 error
        - changed to ODS file templates
* Sep 27, 2015
    - (by: sonofeft) 
        - Set to Pre-Alpha
        - Added some element depth analysis and file differencing
* Sep 23, 2015
    - (by: sonofeft) 
        - changed .gitignore for test ods file
        - removed test ods chart
        - Got both python 2 and 3 to pass tox
            A bit gooned up with using ElementTree.py from both python 2.7 and 3.4
            as starting points for OrderedDict implementations.
            Still much testing to be done.
        - Removed lxml as a requirement
            Modified standard xml library to use OrderedDict so that attribute order
            is maintained with xml file templates
            Wrote an ElementTree wrapper to take care of all the namespace stuff
            that vanilla xml.etree did not.
* Sep 21, 2015
    - (by: sonofeft) 
        - Some success with simple scatter charts
* Sep 20, 2015
    - (by: sonofeft) 
        - VERY CRUDE... Kinda Working
* Sep 19, 2015
    - (by: sonofeft) 
        - put Alpha Code Warning into docs
        - Some TOX and doc changes
        - fix LICENSE file
        - First Working Spreadsheet w/o Charts
    - (by: Charlie Taylor) 
        - Initial commit


* Sep 18, 2015
    - (by: sonofeft)
        - First Created ODSCharts with PyHatch