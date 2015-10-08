

.. image:: https://travis-ci.org/sonofeft/ODSCharts.svg?branch=master
    :target: https://travis-ci.org/sonofeft/ODSCharts

.. image:: https://img.shields.io/pypi/v/ODSCharts.svg
    :target: https://pypi.python.org/pypi/odscharts
        
.. image:: https://img.shields.io/pypi/pyversions/ODSCharts.svg
    :target: https://wiki.python.org/moin/Python2orPython3

.. image:: https://img.shields.io/pypi/l/ODSCharts.svg
    :target: https://pypi.python.org/pypi/odscharts




ODSCharts For Microsoft Excel, LibreOffice And OpenOffice
=========================================================

ODSCharts creates Opendocument spreadsheets `(*.ods files)` with scatter charts that either 
Microsoft Excel, LibreOffice or OpenOffice can open and manipulate.


See the Code at: `<https://github.com/sonofeft/ODSCharts>`_

See the Docs at: `<http://odscharts.readthedocs.org/en/latest/>`_

See PyPI page at:`<https://pypi.python.org/pypi/odscharts>`_


ODSCharts will create ods files readable by Microsoft Excel, LibreOffice or OpenOffice.

The format is a very narrow subset of full spreadsheet support. 
*There is no attempt to supply a full API interface*::

    #. All sheets contain either a table of numbers or a chart object
        - A table of numbers: 
            - starts at "A1"
            - row 1 is labels
            - row 2 is units
            - row 3 through N is float or string entries
        - Chart objects are scatter plots
            - Each series is a column from a table
            - Each x axis is a column from a table


`Click images to see full size`

.. image:: ./_static/alt_vs_tp_excel.png
    :width: 40%
.. image:: ./_static/alt_vs_tp_oo.png
    :width: 50%


What I Think I Know
-------------------

    * Matplotlib is very good when you want to publish a chart **without** the data.

    * Spreadsheets are very good when you want to publish **both** the chart **and** the data.
    
    * Python is a great general-purpose programming language for science and engineering 
    
    * Therefore the world needs a cross-platform, open-source, python solution to generate  cross-platform, open-source spreadsheet files.



What I Know About ODS
---------------------
    
    * ``*.ods`` files are cross-platform, open-source spreadsheet files.
    
    * LibreOffice and OpenOffice read ``*.ods`` files created by Excel much better than Excel reads ``*.ods`` files created by LibreOffice or OpenOffice.

    * Excel ``*.ods`` files are more simple than LibreOffice or OpenOffice ``*.ods`` files (Excel only partially supports ``ods``)
    
    * It makes sense to reverse-engineer Excel-generated ``*.ods`` files as cross-platform, open-source spreadsheet files.

    * ODSCharts generates ``*.ods`` files by reverse-engineering ``*.ods`` files created by Excel.
    
That's It... That's how ODSCharts was born.
