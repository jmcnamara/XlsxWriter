.. _ex_chart_pie:

Example: Pie Chart
==================

Example of creating Excel Pie charts. Chart 1 in the following example is:

.. image:: _images/chart_pie1.png
   :scale: 75 %

Chart 2 shows how to set segment colors.

It is possible to define chart colors for most types of XlsxWriter charts via
the :func:`add_series()` method. However, Pie charts are a special case since
each segment is represented as a point and as such it is necessary to assign
formatting to each point in the series.

.. image:: _images/chart_pie2.png
   :scale: 75 %

Chart 3 shows how to rotate the segments of the chart:

.. image:: _images/chart_pie3.png
   :scale: 75 %


.. literalinclude:: ../../../examples/chart_pie.py
