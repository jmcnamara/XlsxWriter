.. _chart_class:

The Chart Class
===============

The ``Chart`` module is a base class for modules that implement charts in
XlsxWriter. The information in this section is applicable to all of the
available chart subclasses, such as Area, Bar, Column, Line, Pie, Scatter,
Stock and Radar.

A chart object is created via the Workbook :func:`add_chart()` method where the
chart type is specified::

    chart = workbook.add_chart({type, 'column'})

It is then inserted into a worksheet as an embedded chart using the
:func:`insert_chart` Worksheet method::

    worksheet.insert_chart('A7', chart)

.. Note::
   Chart sheets, i.e, charts that occupy an entire worksheet in a workbook
   are not supported yet. They will be added in the next 1-2 releases.

The following is a small working example or adding an embedded chart::

    from xlsxwriter.workbook import Workbook

    workbook = Workbook('chart.xlsx')
    worksheet = workbook.add_worksheet()

    # Create a new Chart object.
    chart = workbook.add_chart({'type': 'column'})

    # Write some data to add to plot on the chart.
    data = [
        [1, 2, 3, 4, 5],
        [2, 4, 6, 8, 10],
        [3, 6, 9, 12, 15],
    ]

    worksheet.write_column('A1', data[0])
    worksheet.write_column('B1', data[1])
    worksheet.write_column('C1', data[2])

    # Configure the chart. In simplest case we add one or more data series.
    chart.add_series({'values': '=Sheet1!$A$1:$A$5'})
    chart.add_series({'values': '=Sheet1!$B$1:$B$5'})
    chart.add_series({'values': '=Sheet1!$C$1:$C$5'})

    # Insert the chart into the worksheet.
    worksheet.insert_chart('A7', chart)

    workbook.close()

.. image:: _static/chart_simple.png
   :scale: 75 %


The supported chart types are:

* ``area``: Creates an Area (filled line) style chart.

* ``bar``: Creates a Bar style (transposed histogram) chart.

* ``column``: Creates a column style (histogram) chart.

* ``line``: Creates a Line style chart.

* ``pie``: Creates a Pie style chart.

* ``scatter``: Creates a Scatter style chart.

* ``stock``: Creates a Stock style chart.

* ``radar``: Creates a Radar style chart.


Chart subtypes are also supported for some chart types::

    workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})

The available subtypes are::

    area
        stacked
        percent_stacked

    bar
        stacked
        percent_stacked

    column
        stacked
        percent_stacked

    scatter
        straight_with_markers
        straight
        smooth_with_markers
        smooth

    radar
        with_markers
        filled


Methods that are common to all chart types are documented below. See
:ref:`working_with_charts` for chart specific information.


chart.add_series()
------------------

.. py:function:: add_series(options)

   Add a data series to a chart.

   :param dict options: A dictionary of chart series options.

In Excel a chart **series** is a collection of information hat define which
data is plotted such as values, axis labels and formatting.

For an XlsxWriter chart object the ``add_series()`` method is used to set the
properties for a series::

    chart.add_series({
        'categories': '=Sheet1!$A$1:$A$5',
        'values':     '=Sheet1!$B$1:$B$5',
        'line':       {'color': 'red'},
    })

    # Or using a list of values instead of category/value formulas:
    #     [sheetname, first_row, first_col, last_row, last_col]
    chart1.add_series({
        'categories': ['Sheet1', 0, 0, 4, 0],
        'values':     ['Sheet1', 0, 1, 4, 1],
        'line':       {'color': 'red'},
    })

As shown above the ``categories`` and ``values`` can take either a range
formula such as ``=Sheet1!$A$2:$A$7`` or, more usefully when generating the
range programmatically, a list with zero indexed row/column values.

The series options that can be set are:

* ``values``: This is the most important property of a series and is the only
  mandatory option for every chart object. This option links the chart with
  the worksheet data that it displays. The data range can be set using a
  formula as shown in the first example above or using a list of values as
  shown in the second example.

* ``categories``: This sets the chart category labels. The category is more
  or less the same as the X axis. In most chart types the ``categories``
  property is optional and the chart will just assume a sequential series from
  ``1..n``.

* ``name``: Set the name for the series. The name is displayed in the chart
  legend and in the formula bar. The name property is optional and if it isn't
  supplied it will default to ``Series 1..n``.

* ``line``: Set the properties of the series line type such as colour and
  width. See :ref:`chart_formatting_line`.

* ``border``: Set the border properties of the series such as colour and
  style. See :ref:`chart_formatting_border`.

* ``fill``: Set the fill properties of the series such as colour. See
  :ref:`chart_formatting_fill`.

* ``marker``: Set the properties of the series marker such as style and
  colour. See :ref:`chart_series_option_marker`.

* ``trendline``: Set the properties of the series trendline such as linear,
  polynomial and moving average types. See
  :ref:`chart_series_option_trendline`.

* ``y_error_bars``: Set vertical error bounds for a chart series. See
  :ref:`chart_series_option_error_bars`.

* ``x_error_bars``: Set horizontal error bounds for a chart series. See
  :ref:`chart_series_option_error_bars`.

* ``data_labels``: Set data labels for the series. See
  :ref:`chart_series_option_data_labels`.

* ``points``: Set properties for individual points in a series. See
  :ref:`chart_series_option_points`.

* ``invert_if_negative``: Invert the fill colour for negative values. Usually
  only applicable to column and bar charts.

* ``overlap``: Set the overlap between series in a Bar/Column chart. The
  range is +/- 100. The default is 0::

    chart.add_series({
        'categories': '=Sheet1!$A$1:$A$5',
        'values':     '=Sheet1!$B$1:$B$5',
        'overlap':    10,
    })

  Note, it is only necessary to apply the ``overlap`` property to one series
  in the chart.

* ``gap``: Set the gap between series in a Bar/Column chart. The range is 0
  to 500. The default is 150::

    chart.add_series({
        'categories': '=Sheet1!$A$1:$A$5',
        'values':     '=Sheet1!$B$1:$B$5',
        'gap':        200,
    })

  Note, it is only necessary to apply the ``gap`` property to one series in
  the chart.

* ``smooth``: Ask the points in the series to be fit to a smooth curve,
  rather than straight line segments.  This currently applies only to ``line``
  and ``scatter`` charts.  Setting this option will override the behavior
  suggested by the ``subtype`` option for a ``scatter`` chart.  The default value
  is ``None``, in which case the default behavior depends on the chart type and 
  subtype::
    
    chart.add_series({
        'categories': '=Sheet1!$A$1:$A$5',
        'values':     '=Sheet1!$B$1:$B$5',
        'smooth':     True,
    })


More than one series can be added to a chart. In fact, some chart types such as
``stock`` require it. The series numbering and order in the Excel chart will
be the same as the order in which they are added in XlsxWriter.


chart.set_x_axis()
------------------

.. py:function:: set_x_axis(options)

   Set the chart X axis options.

   :param dict options: A dictionary of axis options.

The ``set_x_axis()`` method is used to set properties of the X axis::

    chart.set_x_axis({
        'name': 'Earnings per Quarter',
        'name_font': {'size': 14, 'bold': True},
        'num_font':  {'italic': True },
    })

.. image:: _static/chart_x_axis.png
   :scale: 75 %

The options that can be set are::

    name
    name_font
    num_font
    num_format
    min
    max
    minor_unit
    major_unit
    crossing
    reverse
    log_base
    label_position
    major_gridlines
    minor_gridlines
    visible

These options are explained below. Some properties are only applicable to
**value** or **category** axes (this is noted in each case). See
:ref:`chart_val_cat_axes` for an explanation of Excel's distinction between
the axis types.

* ``name``: Set the name (also known as title or caption) for the axis. The
  name is displayed below the X axis. (Applicable to category and value
  axes.)::

    chart.set_x_axis({'name': 'Earnings per Quarter'})

  This property is optional. The default is to have no axis name.

  The name can also be a formula such as ``=Sheet1!$A$1``.

* ``name_font``: Set the font properties for the axis name. (Applicable to
  category and value axes.)::

    chart.set_x_axis({'name_font': {'bold': True, 'italic': True}})

  See the :ref:`chart_fonts` section for more details on font properties.

* ``num_font``: Set the font properties for the axis numbers. (Applicable to
  category and value axes)::

    chart.set_x_axis({'name_font': {'bold': True, 'italic': True}})

  See the :ref:`chart_fonts` section for more details on font properties.

* ``num_format``: Set the number format for the axis. (Applicable to category
  and value axes)::

    chart.set_x_axis({'num_format': '#,##0.00'})
    chart.set_y_axis({'num_format': '0.00%'})

  The number format is similar to the Worksheet Cell Format ``num_format``
  apart from the fact that a format index cannot be used. An explicit format
  string must be used as shown above. See :func:`set_num_format()` for more
  information.

* ``min``: Set the minimum value for the axis range. (Applicable to value
  axes only.)::

    chart.set_x_axis({'min': 3, 'max': 6})

  .. image:: _static/chart_max_min.png
     :scale: 75 %

* ``max``: Set the maximum value for the axis range. (Applicable to value
  axes only.)

* ``minor_unit``: Set the increment of the minor units in the axis range.
  (Applicable to value axes only.)::

    chart.set_x_axis({'minor_unit': 0.4, 'major_unit': 2})

* ``major_unit``: Set the increment of the major units in the axis range.
  (Applicable to value axes only.)

* ``crossing``: Set the position where the y axis will cross the x axis.
  (Applicable to category and value axes.)

  The ``crossing`` value can either be the string ``'max'`` to set the
  crossing at the maximum axis value or a numeric value::

    chart.set_x_axis({'crossing': 3})
    chart.set_y_axis({'crossing': 'max'})

  **For category axes the numeric value must be an integer** to represent the
  category number that the axis crosses at. For value axes it can have any
  value associated with the axis. See also :ref:`chart_val_cat_axes`.

  If crossing is omitted (the default) the crossing will be set automatically
  by Excel based on the chart data.

* ``reverse``: Reverse the order of the axis categories or values.
  (Applicable to category and value axes.)::

    chart.set_y_axis({'reverse': True})

  .. image:: _static/chart_reverse.png
     :scale: 75 %

* ``log_base``: Set the log base of the axis range. (Applicable to value axes
  only.)::

    chart.set_y_axis({'log_base': 10})

* ``label_position``: Set the "Axis labels" position for the axis. The
  following positions are available::

    next_to (the default)
    high
    low
    none

  For example::

    chart.set_x_axis({'label_position': 'high'})
    chart.set_y_axis({'label_position': 'low'})

* ``major_gridlines``: Configure the major gridlines for the axis. The
  available properties are::

    visible
    line

  For example::

    chart.set_x_axis({
        'major_gridlines': {
            'visible': True,
            'line': {'width': 1.25, 'dash_type': 'dash'}
        },
    })

  .. image:: _static/chart_gridlines.png
     :scale: 75 %

  The ``visible`` property is usually on for the X axis but it depends on the
  type of chart.

  The ``line`` property sets the gridline properties such as colour and
  width. See :ref:`chart_formatting`.

* ``minor_gridlines``: This takes the same options as ``major_gridlines``
  above.

  The minor gridline ``visible`` property is off by default for all chart
  types.

* ``visible``: Configure the visibility of the axis::

    chart.set_y_axis({'visible': False})

  Axes are visible by default.


chart.set_y_axis()
------------------

.. py:function:: set_y_axis(options)

   Set the chart Y axis options.

   :param dict options: A dictionary of axis options.

The ``set_y_axis()`` method is used to set properties of the Y axis.

The properties that can be set are the same as for ``set_x_axis``, see above.


chart.set_x2_axis()
-------------------

.. py:function:: set_x2_axis(options)

   Set the chart secondary X axis options.

   :param dict options: A dictionary of axis options.

The ``set_x2_axis()`` method is used to set properties of the secondary X axis,
see :func:`chart_secondary_axes`.

The properties that can be set are the same as for ``set_x_axis``, see above.

The default properties for this axis are::

    'label_position': 'none',
    'crossing':       'max',
    'visible':        False,


chart.set_y2_axis()
-------------------

.. py:function:: set_y2_axis(options)

   Set the chart secondary Y axis options.

   :param dict options: A dictionary of axis options.

The ``set_y2_axis()`` method is used to set properties of the secondary Y axis,
see :func:`chart_secondary_axes`.

The properties that can be set are the same as for ``set_x_axis``, see above.

The default properties for this axis are::

    'major_gridlines': {'visible': True}


chart.set_size()
----------------

.. py:function:: set_size(options)

   Set the size or scale of the chart.

   :param dict options: A dictionary of chart size options.

The ``set_size()`` method is used to set the dimensions of the chart. The size
properties that can be set are::

     width
     height
     x_scale
     y_scale
     x_offset
     y_offset

The ``width`` and ``height`` are in pixels. The default chart width x height is
480 x 288 pixels. The size of the chart can be modified by setting the
``width`` and ``height`` or by setting the ``x_scale`` and ``y_scale``::

    chart.set_size({'width': 720, 'height': 576})
    # Same as:
    chart.set_size({'x_scale', 1.5, 'y_scale', 2})

The ``x_offset`` and ``y_offset`` position the top left corner of the chart in
the cell that it is inserted into.

Note: the ``x_offset`` and ``y_offset`` parameters can also be set via the
:func:`insert_chart()` method::

    worksheet.insert_chart('E2', chart, {'x_offset': 25, 'y_offset': 10})


chart.set_title()
-----------------

.. py:function:: set_title(options)

   Set the chart title options.

   :param dict options: A dictionary of chart size options.

The ``set_title()`` method is used to set properties of the chart title::

    chart.set_title({'name': 'Year End Results'})

.. image:: _static/chart_title.png
   :scale: 75 %

The properties that can be set are:

* ``name``: Set the name (title) for the chart. The name is displayed above
  the chart. The name can also be a formula such as ``=Sheet1!$A$1``. The name
  property is optional. The default is to have no chart title.

* ``name_font``: Set the font properties for the chart title. See
  :ref:`chart_fonts`.


chart.set_legend()
------------------

.. py:function:: set_legend(options)

   Set the chart legend options.

   :param dict options: A dictionary of chart legend options.

The ``set_legend()`` method is used to set properties of the chart legend. For
example it can be used to turn off the default chart legend::

    chart.set_legend({'position': 'none'})

.. image:: _static/chart_legend_none.png
   :scale: 75 %

The options that can be set are::

    position
    delete_series


* ``position``: Set the position of the chart legend::

    chart.set_legend({'position': 'bottom'})

  .. image:: _static/chart_legend_bottom.png
     :scale: 75 %

  The default legend position is ``right``. The available positions are::

    none
    top
    bottom
    left
    right
    overlay_left
    overlay_right

* ``delete_series``: This allows you to remove one or more series from the
  the legend (the series will still display on the chart). This property takes
  a list as an argument and the series are zero indexed::

    # Delete/hide series index 0 and 2 from the legend.
    chart.set_legend({'delete_series': [0, 2]})

  .. image:: _static/chart_legend_delete.png
     :scale: 75 %


chart.set_chartarea()
---------------------

.. py:function:: set_chartarea(options)

   Set the chart area options.

   :param dict options: A dictionary of chart area options.

The ``set_chartarea()`` method is used to set the properties of the chart area::

    chart.set_chartarea({
        'border': {'none': True},
        'fill':   {'color': 'red'}
    })

.. image:: _static/chart_chartarea.png
   :scale: 75 %

The properties that can be set are:

* ``border``: Set the border properties of the chartarea such as colour and
  style. See :ref:`chart_formatting`.

* ``fill``: Set the fill properties of the chartarea such as colour. See
  :ref:`chart_formatting`.



chart.set_plotarea()
--------------------

.. py:function:: set_plotarea(options)

   Set the plot area options.

   :param dict options: A dictionary of plot area options.

The ``set_plotarea()`` method is used to set properties of the plot area of a
chart::

    chart.set_plotarea({
        'border': {'color': 'red', 'width': 2, 'dash_type': 'dash'},
        'fill':   {'color': '#FFFFC2'}
    })

.. image:: _static/chart_plotarea.png
   :scale: 75 %

The properties that can be set are:

* ``border``: Set the border properties of the plotarea such as colour and
  style. See :ref:`chart_formatting`.

* ``fill``: Set the fill properties of the plotarea such as colour. See
  :ref:`chart_formatting`.



chart.set_style()
-----------------

.. py:function:: set_style(style_id)

   Set the chart style type.

   :param int style_id: An index representing the chart style.

The ``set_style()`` method is used to set the style of the chart to one of the
48 built-in styles available on the 'Design' tab in Excel::

    chart.set_style(37)

.. image:: _static/chart_style.png
   :scale: 75 %

The style index number is counter from 1 on the top left. The default style is
2.


chart.set_table()
-----------------

.. py:function:: set_table(options)

   Set properties for an axis data table.

   :param dict options: A dictionary of axis table options.

The ``set_table()`` method adds a data table below the horizontal axis with the
data used to plot the chart::

    chart.set_table()

.. image:: _static/chart_table.png
   :scale: 75 %

The available options, with default values are::

    'horizontal': True   # Display vertical lines in the table.
    'vertical':   True   # Display horizontal lines in the table.
    'outline':    True   # Display an outline in the table.
    'show_keys':  False  # Show the legend keys with the table data.

    chart.set_table({'show_keys': True})

The data table can only be shown with Bar, Column, Line, Area and stock charts.


chart.set_up_down_bars()
------------------------

.. py:function:: set_up_down_bars(options)

   Set properties for the chart up-down bars.

   :param dict options: A dictionary of options.

The ``set_up_down_bars()`` method adds Up-Down bars to Line charts to indicate
the difference between the first and last data series::

    chart.set_up_down_bars()

It is possible to format the up and down bars to add ``fill`` and ``border``
properties if required. See :ref:`chart_formatting`::

    chart.set_up_down_bars({
        'up': {
            'fill':   {'color': '#00B050'},
            'border': {'color': 'black'}
        },
        'down': {
            'fill':   {'color': 'red'},
            'border': {'color': 'black'},
        },
    })

.. image:: _static/chart_up_down_bars.png
   :scale: 75 %

Up-down bars can only be applied to Line charts and to Stock charts (by
default).


chart.set_drop_lines()
----------------------

.. py:function:: set_drop_lines(options)

   Set properties for the chart drop lines.

   :param dict options: A dictionary of options.

The ``set_drop_lines()`` method adds Drop Lines to charts to show the Category
value of points in the data::

    chart.set_drop_lines()

.. image:: _static/chart_drop_lines.png
   :scale: 75 %

It is possible to format the Drop Line ``line`` properties if required. See
:ref:`chart_formatting`::

    chart.set_drop_lines({'line': {'color': 'red',
                                   'dash_type': 'square_dot'}})

Drop Lines are only available in Line, Area and Stock charts.


chart.set_high_low_lines()
--------------------------

.. py:function:: set_high_low_lines(options)

   Set properties for the chart high-low lines.

   :param dict options: A dictionary of options.

The ``set_high_low_lines()`` method adds High-Low lines to charts to show the
maximum and minimum values of points in a Category::

    chart.set_high_low_lines()

.. image:: _static/chart_high_low_lines.png
   :scale: 75 %

It is possible to format the High-Low Line ``line`` properties if required. See
:ref:`chart_formatting`::

    chart.set_high_low_lines({
        'line': {
            'color': 'red',
            'dash_type': 'square_dot'
        }
    })

High-Low Lines are only available in Line and Stock charts.


chart.show_blanks_as()
----------------------

.. py:function:: show_blanks_as(option)

   Set the option for displaying blank data in a chart.

   :param string option: A string representing the display option.

The ``show_blanks_as()`` method controls how blank data is displayed in a
chart::

    chart.show_blanks_as('span')

The available options are::

    'gap'   # Blank data is shown as a gap. The default.
    'zero'  # Blank data is displayed as zero.
    'span'  # Blank data is connected with a line.


chart.show_hidden_data()
------------------------

.. py:function:: show_hidden_data()

   Display data on charts from hidden rows or columns.


Display data in hidden rows or columns on the chart::

    chart.show_hidden_data()


See also :ref:`working_with_charts` and :ref:`chart_examples`.
