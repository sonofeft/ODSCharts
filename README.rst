

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

The format is a very narrow subset of full spreadsheet support::

    #. All sheets contain either a table of numbers or a chart object
        - A table of numbers: 
            - starts at "A1"
            - row 1 is labels
            - row 2 is units
            - row 3 through N is float or string entries
        - Chart objects are scatter plots
            - Each series is a column from a table
            - Each x axis is a column from a table

There is no attempt to supply a full API interface.

