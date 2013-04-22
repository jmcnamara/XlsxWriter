.. _working_with_charts:

Working with Charts
===================

TODO Add intro

.. _chart_val_cat_axes:

Value and Category Axes
-----------------------

Excel differentiates between a chart axis that is used for series
**categories** and an axis that is used for series **values**.

In the example above the X axis is the category axis and each of the values is
evenly spaced. The Y axis (in this case) is the value axis and points are
displayed according to their value.

Since Excel treats the axes differently it also handles their formatting
differently and exposes different properties for each.

As such some of ``XlsxWriter`` axis properties can be set for a value axis,
some can be set for a category axis and some properties can be set for both.

For example the ``min`` and ``max`` properties can only be set for value axes
and ``reverse`` can be set for both. The type of axis that a property applies
to is shown in the ``set_x_axis()`` section of the documentation above.

Some charts such as ``Scatter`` and ``Stock`` have two value axes.


.. _chart_series_options:

Chart Series Options
--------------------

This section details the following properties of ``add_series()`` in more
detail::

    marker
    trendline
    y_error_bars
    x_error_bars
    data_labels
    points


.. _chart_series_option_marker:

Chart series option: Marker
---------------------------

The marker format specifies the properties of the markers used to distinguish
series on a chart. In general only Line and Scatter chart types and trendlines
use markers.

The following properties can be set for ``marker`` formats in a chart::

    type 
    size 
    border 
    fill

The ``type`` property sets the type of marker that is used with a series::

        chart.add_series({
            'values': '=Sheet1!$B$1:$B$5',
            'marker': {'type': 'diamond'},
        })

The following ``type`` properties can be set for ``marker`` formats in a chart.
These are shown in the same order as in the Excel format dialog::

    automatic 
    none 
    square 
    diamond 
    triangle 
    x 
    star 
    short_dash 
    long_dash
    circle plus

The ``automatic`` type is a special case which turns on a marker using the
default marker style for the particular series number::

    chart.add_series({
        'values': '=Sheet1!$B$2:$B$7',
        'marker': {'type': 'automatic'},
    })

If ``automatic`` is on then other marker properties such as size, border or
fill cannot be set.

The ``size`` property sets the size of the marker and is generally used in
conjunction with ``type``::

        chart.add_series({
            'values': '=Sheet1!$B$1:$B$5',
            'marker': {'type': 'diamond', 'size': 7},
        })
        
Nested ``border`` and ``fill`` properties can also be set for a marker. See the
"CHART FORMATTING" section below::

        chart.add_series({
            'categories': '=Sheet1!$A$1:$A$5',
            'values': '=Sheet1!$B$1:$B$5',
            'marker': {
                'type': 'square',
                'size': 5,
                'line': {'color': 'yellow'},
                'fill': {'color': 'red'},
            },
        })

.. _chart_series_option_trendline:

Chart series option: Trendline
------------------------------

A trendline can be added to a chart series to indicate trends in the data such
as a moving average or a polynomial fit.

The following properties can be set for trendlines in a chart series::

    type 
    order      (for polynomial trends) 
    period     (for moving average) 
    forward    (for all except moving average) 
    backward   (for all except moving average)
    name 
    line

The ``type`` property sets the type of trendline in the series::

    chart.add_series({
        values, '=Sheet1not B1:B5', trendline, { type, 'linear' },
    })

The available ``trendline`` types are::

    exponential
    linear
    log
    moving_average
    polynomial
    power

A ``polynomial`` trendline can also specify the ``order`` of the polynomial.
The default value is 2::

    chart.add_series({
        values, '=Sheet1!B1:B5', trendline,:
            type, 'polynomial', order, 3,
        },
    })

A ``moving_average`` trendline can also specify the ``period`` of the moving
average. The default value is 2::

    chart.add_series({
        values, '=Sheet1!B1:B5', trendline,:
            type, 'moving_average', period, 3,
        },
    })

The ``forward`` and ``backward`` properties set the forecast period of the
trendline::

    chart.add_series({
        values, '=Sheet1!B1:B5', trendline,:
            type, 'linear', forward, 0.5, backward, 0.5,
        },
    })

The ``name`` property sets an optional name for the trendline that will appear
in the chart legend. If it isn't specified the Excel default name will be
displayed. This is usually a combination of the trendline type and the series
name::

    chart.add_series({
        values, '=Sheet1!B1:B5', trendline,:
            type, 'linear', name, 'Interpolated trend',
        },
    })

Several of these properties can be set in one go::

    chart.add_series({
        values, '=Sheet1!B1:B5',
        trendline,:
            type, 'linear',
            name, 'My trend name',
            forward, 0.5,
            backward, 0.5,
            line,:
                color, 'red',
                width, 1,
                dash_type, 'long_dash',
            },
        },
    })

Trendlines cannot be added to series in a stacked chart or pie chart, radar
chart or (when implemented) to 3D, surface, or doughnut charts.

.. _chart_series_option_error_bars:

Chart series option: Error Bars
-------------------------------

Error bars can be added to a chart series to indicate error bounds in the data.
The error bars can be vertical ``y_error_bars`` (the most common type) or
horizontal ``x_error_bars`` (for Bar and Scatter charts only).

The following properties can be set for error bars in a chart series::

    type value (for all types except standard error) direction end_style
    line

The ``type`` property sets the type of error bars in the series::

    chart.add_series({
        values, '=Sheet1!B1:B5', y_error_bars, { type,
        'standard_error' },
    })

The available error bars types are available::

    fixed
    percentage
    standard_deviation
    standard_error

Note, the "custom" error bars type is not supported.

All error bar types, except for ``standard_error`` must also have a value
associated with it for the error bounds::

    chart.add_series({
        values, '=Sheet1!B1:B5',
        y_error_bars,:
            type, 'percentage',
            value, 5,
        },
    })

The ``direction`` property sets the direction of the error bars. It should be
one of the following::

    plus # Positive direction only.
    minus # Negative direction only.
    both # Plus and minus directions, The default.

The ``end_style`` property sets the style of the error bar end cap. The options
are 1 (the default) or 0 (for no end cap)::

    chart.add_series({
        values, '=Sheet1!B1:B5',
        y_error_bars,:
            type, 'fixed',
            value, 2,
            end_style, 0,
            direction, 'minus'
        },
    })


.. _chart_series_option_data_labels:

Chart series option: Data Labels
--------------------------------

Data labels can be added to a chart series to indicate the values of the
plotted data points.

The following properties can be set for ``data_labels`` formats in a chart::

    value category series_name position leader_lines percentage

The ``value`` property turns on the *Value* data label for a series::

    chart.add_series({
        values, '=Sheet1!B1:B5', data_labels, { value, 1 },
    })

The ``category`` property turns on the *Category Name* data label for a series::

    chart.add_series({
        values, '=Sheet1!B1:B5', data_labels, { category, 1 },
    })

The ``series_name`` property turns on the *Series Name* data label for a series::

    chart.add_series({
        values, '=Sheet1!B1:B5', data_labels, { series_name, 1 },
    })

The ``position`` property is used to position the data label for a series::

    chart.add_series({
        values, '=Sheet1!B1:B5', data_labels, { value, 1, position,
        'center' },
    })

Valid positions are::

    center
    right
    left
    top
    bottom
    above # Same as top
    below # Same as bottom
    inside_end # Pie chart mainly.
    outside_end # Pie chart mainly.
    best_fit # Pie chart mainly.

The ``percentage`` property is used to turn on the display of data labels as a
*Percentage* for a series. It is mainly used for pie charts::

    chart.add_series({
        values, '=Sheet1!B1:B5', data_labels, { percentage, 1 },
    })

The ``leader_lines`` property is used to turn on *Leader Lines* for the data
label for a series. It is mainly used for pie charts::

    chart.add_series({
        values, '=Sheet1!B1:B5', data_labels, { value, 1,
        leader_lines, 1 },
    })

Note: Even when leader lines are turned on they aren't automatically visible in
Excel or XlsxWriter. Due to an Excel limitation (or design) leader lines only
appear if the data label is moved manually or if the data labels are very
close and need to be adjusted automatically.


.. _chart_series_option_points:

Chart series option: Points
---------------------------

In general formatting is applied to an entire series in a chart. However, it is
occasionally required to format individual points in a series. In particular
this is required for Pie charts where each segment is represented by a point.

In these cases it is possible to use the ``points`` property of
``add_series()``::

    chart.add_series({
        values, '=Sheet1!A1:A3',
        points, [
            { fill, { color, '#FF0000' } },
            { fill, { color, '#CC0000' } },
            { fill, { color, '#990000' } },
        ],
    })

The ``points`` property takes an array ref of format options (see the "CHART
FORMATTING" section below). To assign default properties to points in a series
pass ``undef`` values in the array ref::

    # Format point 3 of 3 only.
    chart.add_series({
        values, '=Sheet1!A1:A3',
        points, [
            None,
            None,
            { fill, { color, '#990000' } },
        ],
    })

    # Format the first point only. chart.add_series({
        values, '=Sheet1!A1:A3', points, [ { fill, { color,
        '#FF0000' } } ],
    })




.. _chart_formatting:

Chart Formatting
----------------

The following chart formatting properties can be set for any chart object that
they apply to (and that are supported by XlsxWriter) such as chart lines,
column fill areas, plot area borders, markers, gridlines and other chart
elements documented above::

    line 
    border 
    fill

Chart formatting properties are generally set using hash refs::

    chart.add_series({
        values, '=Sheet1!B1:B5', line, { color, 'blue' },
    })

In some cases the format properties can be nested. For example a ``marker`` may
contain ``border`` and ``fill`` sub-properties::

    chart.add_series({
        values, '=Sheet1!B1:B5', line, { color, 'blue' }, marker,:
            type, 'square', size, 5, border, { color, 'red' },
            fill, { color, 'yellow' },
        },
    })

.. _chart_formatting_line:

Chart formatting: Line
----------------------

The line format is used to specify properties of line objects that appear in a
chart such as a plotted line on a chart or a border.

The following properties can be set for ``line`` formats in a chart::

    none color width dash_type

The ``none`` property is uses to turn the ``line`` off (it is always on by default except in Scatter charts). This is useful if you wish to plot a series with markers but without a line::

    chart.add_series({
        values, '=Sheet1!B1:B5', line, { none, 1 },
    })

The ``color`` property sets the color of the ``line``::

    chart.add_series({
        values, '=Sheet1!B1:B5', line, { color, 'red' },
    })

The available colours are shown in the main XlsxWriter documentation. It is
also possible to set the colour of a line with a HTML style RGB colour::

    chart.add_series({
        line, { color, '#FF0000' },
    })

The ``width`` property sets the width of the ``line``. It should be specified
in increments of 0.25 of a point as in Excel::

    chart.add_series({
        values, '=Sheet1!B1:B5', line, { width, 3.25 },
    })

The ``dash_type`` property sets the dash style of the line::

    chart.add_series({
        values, '=Sheet1!B1:B5', line, { dash_type, 'dash_dot' },
    })

The following ``dash_type`` values are available. They are shown in the order
that they appear in the Excel dialog::

    solid 
    round_dot 
    square_dot 
    dash 
    dash_dot 
    long_dash 
    long_dash_dot
    long_dash_dot_dot

The default line style is ``solid``    })

More than one ``line`` property can be specified at a time::

    chart.add_series({
        values, '=Sheet1!B1:B5',
        line,:
            color, 'red',
            width, 1.25,
            dash_type, 'square_dot',
        },
    })

.. _chart_formatting_border:    

Chart formatting: Border
------------------------

The ``border`` property is a synonym for ``line``.

It can be used as a descriptive substitute for ``line`` in chart types such as
Bar and Column that have a border and fill style rather than a line style. In
general chart objects with a ``border`` property will also have a fill
property.

.. _chart_formatting_fill:

Chart formatting: Fill
----------------------

The fill format is used to specify filled areas of chart objects such as the
interior of a column or the background of the chart itself.

The following properties can be set for ``fill`` formats in a chart::

    none color

The ``none`` property is used to turn the ``fill`` property off (it is
generally on by default)::

    chart.add_series({
        values, '=Sheet1!B1:B5', fill, { none, 1 },
    })

The ``color`` property sets the colour of the ``fill`` area::

    chart.add_series({
        values, '=Sheet1!B1:B5', fill, { color, 'red' },
    })

The available colours are shown in the main XlsxWriter documentation. It is
also possible to set the colour of a fill with a HTML style RGB colour::

    chart.add_series({
        fill, { color, '#FF0000' },
    })

The ``fill`` format is generally used in conjunction with a ``border`` format
which has the same properties as a ``line`` format::

    chart.add_series({
        values, '=Sheet1!B1:B5', border, { color, 'red' }, fill, {
        color, 'yellow' },
    })


.. _chart_fonts:

Chart Fonts
-----------

The following font properties can be set for any chart object that they apply
to (and that are supported by XlsxWriter) such as chart titles, axis labels
and axis numbering. They correspond to the equivalent Worksheet cell Format
object properties. See "FORMAT_METHODS" in XlsxWriter for more information::

    name size bold italic underline color

The following explains the available font properties::

* ``name``: Set the font name::

    chart.set_x_axis(num_font, { name, 'Arial' })

* ``size``: Set the font size::

    chart.set_x_axis(num_font, { name, 'Arial', size, 10 })

* ``bold``: Set the font bold property, should be 0 or 1::

    chart.set_x_axis(num_font, { bold, 1 })

* ``italic``: Set the font italic property, should be 0 or 1::

    chart.set_x_axis(num_font, { italic, 1 })

* ``underline``: Set the font underline property, should be 0 or 1::

    chart.set_x_axis(num_font, { underline, 1 })

* ``color``: Set the font color property. Can be a color index, a color name
  or HTML style RGB colour::

    chart.set_x_axis(num_font, { color, 'red' })
    chart.set_y_axis(num_font, { color, '#92D050' })


Here is an example of Font formatting in a Chart program::

    # Format the chart title.
    chart.set_title({
        name, 'Sales Results Chart',
        name_font,:
            name, 'Calibri',
            color, 'yellow',
        },
    })

    # Format the X-axis. chart.set_x_axis({
        name, 'Month', name_font,:
            name, 'Arial', color, '#92D050'
        }, num_font,:
            name, 'Courier New', color, '#00B0F0',
        },
    })

    # Format the Y-axis. chart.set_y_axis({
        name, 'Sales (1000 units)', name_font,:
            name, 'Century', underline, 1, color, 'red'
        }, num_font,:
            bold, 1, italic, 1, color, '#7030A0',
        },
    })


.. _chart_secondary_axes:

Secondary Axes
--------------

TODO




Chart Limitations
-----------------

The chart feature in XlsxWriter is under active development. More chart types
and features will be added in time.

Features that are on the TODO list and will be added are::

* Add more chart sub-types.
* Additional formatting options.
* More axis controls.
* 3D charts.
* Additional chart types such as Bubble or Doughnut.

