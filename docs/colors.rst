
.. colors

Colors in ODSCharts
===================

If no colors are specified in ODSCharts, curves will be assigned color
from one of two color tables, 
the Excel Default Color Table or the Alternate Color Table (see below). 

In the ``add_scatter``
call, the option ``excel_colors=True`` selects the Excel table whereas 
``excel_colors=False`` selects the Alternate.

Because Excel does not support marker colors in ods files, ODSCharts defaults to
using the Excel color table so that the marker colors 
will match when opened in Excel. (i.e. ``excel_colors=True``)

OpenOffice will properly support any specified colors, so the Alternate Table may
be preferred when using OpenOffice (i.e. ``excel_colors=False``)

Users may also assign colors to each curve.
When using ``add_scatter`` or ``add_curve`` to create curves on a plot, the two inputs
``colorL`` and ``color2L`` will specify colors desired for primary or secondary Y Axis.

The python snippet below, for example, will make the first 5 primary Y Axis curves 
red, cyan, grey, #6699AA and darkgreen respectively.

.. code:: python

    xcol = 1,
    ycolL=[2,3,4,5,6],
    colorL=['r','cyan','GRaY','#69a','#006400']

Any color not recognized by name should be input in "#AACCFF" format.

Some common colors such as red, green or blue can be specified in shorthand as r, g or b. 
(see Short Name Color Table below)


.. raw:: html
   :file: ./_static/odscharts_colors.html

