

.. image:: https://travis-ci.org/sonofeft/ODSCharts.svg?branch=master
    :target: https://travis-ci.org/sonofeft/ODSCharts

.. image:: https://img.shields.io/pypi/v/ODSCharts.svg
    :target: https://pypi.python.org/pypi/odscharts
        
.. image:: https://img.shields.io/pypi/pyversions/ODSCharts.svg
    :target: https://wiki.python.org/moin/Python2orPython3

.. image:: https://img.shields.io/pypi/l/ODSCharts.svg
    :target: https://pypi.python.org/pypi/odscharts


Creates Opendocument Spreadsheets With Charts For Microsoft Excel And Openoffice
================================================================================


See the Code at: `<https://github.com/sonofeft/ODSCharts>`_

See the Docs at: `<http://odscharts.readthedocs.org/en/latest/>`_

See PyPI page at:`<https://pypi.python.org/pypi/odscharts>`_


ODSCharts will create ods files readable by either Microsoft Excel or OpenOffice.

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



What I Know About ODF
---------------------

    * Matplotlib is very good when you want to publish a chart **without** the data.

    * Spreadsheets are very good when you want to publish **both** the chart **and** the data.
    
    * I got motivated to make a cross-platform, open-source, python solution to generate  ``*.odf`` files.
    
    * OpenOffice reads ``*.odf`` files created by Excel much better than Excel reads ``*.odf`` files created by OpenOffice.

    * Excel ``*.odf`` files are more simple than OpenOffice ``*.odf`` files (Excel only partially supports odf)

    * ODSCharts generates ``*.odf`` files by reverse-engineering ``*.odf`` files created by Excel.
    
That's It... That's pretty much all I know.
