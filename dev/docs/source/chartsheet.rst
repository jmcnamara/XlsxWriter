.. _chartsheet:

The Chartsheet Class
====================

In Excel a chartsheet is a worksheet that only contains a chart.

.. image:: _images/chartsheet.png

The **Chartsheet** class has some of the functionality of data
:ref:`Worksheets <Worksheet>` such as tab selection, headers, footers, margins
and print properties but its primary purpose is to display a single chart.
This makes it different from ordinary data worksheets which can have one or
more *embedded* charts.

Like a data worksheet a chartsheet object isn't instantiated directly. Instead
a new chartsheet is created by calling the :func:`add_chartsheet()` method
from a :ref:`Workbook <Workbook>` object::

    workbook   = xlsxwriter.Workbook('filename.xlsx')
    worksheet  = workbook.add_worksheet()  # Required for the chart data.
    chartsheet = workbook.add_chartsheet()
    #...
    workbook.close()


A chartsheet object functions as a worksheet and not as a chart. In order to
have it display data a :ref:`Chart <chart_class>` object must be created and
added to the chartsheet::

    chartsheet = workbook.add_chartsheet()
    chart      = workbook.add_chart({'type': 'bar'})

    # Configure the chart.

    chartsheet.set_chart(chart)

The data for the chartsheet chart must be contained on a separate worksheet.
That is why it is always created in conjunction with at least one data
worksheet, as shown above.


chartsheet.set_chart()
----------------------

.. py:function:: set_chart(chart)

   Add a chart to a chartsheet.

   :param chart:       A chart object.

The ``set_chart()`` method is used to insert a chart into a chartsheet. A chart
object is created via the Workbook :func:`add_chart()` method where the chart
type is specified::

    chart = workbook.add_chart({type, 'column'})

    chartsheet.set_chart(chart)

Only one chart can be added to an individual chartsheet.

See :ref:`chart_class`, :ref:`working_with_charts` and :ref:`chart_examples`.


Worksheet methods
-----------------

The following :ref:`Worksheet` methods are also available through a chartsheet:


* :func:`activate()`
* :func:`select()`
* :func:`hide()`
* :func:`set_first_sheet()`
* :func:`protect()`
* :func:`set_zoom()`
* :func:`set_tab_color()`
* :func:`set_landscape()`
* :func:`set_portrait()`
* :func:`set_paper()`
* :func:`set_margins()`
* :func:`set_header()`
* :func:`set_footer()`
* :func:`get_name()`


For example::

    chartsheet.set_tab_color('#FF9900')

The :func:`set_zoom()` method can be used to modify the displayed size of the
chart.


Chartsheet Example
------------------

See :ref:`ex_chartsheet`.

