.. _working_with_formats:

Working with Formats
====================

The methods and properties used to add formatting to a cell are shown in
:ref:`format`. This section provides some additional information about working
with formats.


Creating and using a Format object
----------------------------------

Cell formatting is defined through a :ref:`Format object <format>`. Format
objects are created by calling the workbook ``add_format()`` method as
follows::

    format1 = workbook.add_format()       # Set properties later.
    format2 = workbook.add_format(props)  # Set properties at creation.

Once a Format object has been constructed and its properties have been set it
can be passed as an argument to the worksheet ``write`` methods as follows::

    worksheet.write       (0, 0, 'Foo', format)
    worksheet.write_string(1, 0, 'Bar', format)
    worksheet.write_number(2, 0, 3,     format)
    worksheet.write_blank (3, 0, '',    format)

Formats can also be passed to the worksheet ``set_row()`` and ``set_column()``
methods to define the default formatting properties for a row or column::

    worksheet.set_row(0, 18, format)
    worksheet.set_column('A:D', 20, format)


Format methods and Format properties
------------------------------------

The following table shows the Excel format categories, the formatting
properties that can be applied and the equivalent object method:

+------------+------------------+----------------------+------------------------------+
| Category   | Description      | Property             | Method Name                  |
+============+==================+======================+==============================+
| Font       | Font type        | ``'font_name'``      | :func:`set_font_name()`      |
+------------+------------------+----------------------+------------------------------+
|            | Font size        | ``'font_size'``      | :func:`set_font_size()`      |
+------------+------------------+----------------------+------------------------------+
|            | Font color       | ``'font_color'``     | :func:`set_font_color()`     |
+------------+------------------+----------------------+------------------------------+
|            | Bold             | ``'bold'``           | :func:`set_bold()`           |
+------------+------------------+----------------------+------------------------------+
|            | Italic           | ``'italic'``         | :func:`set_italic()`         |
+------------+------------------+----------------------+------------------------------+
|            | Underline        | ``'underline'``      | :func:`set_underline()`      |
+------------+------------------+----------------------+------------------------------+
|            | Strikeout        | ``'font_strikeout'`` | :func:`set_font_strikeout()` |
+------------+------------------+----------------------+------------------------------+
|            | Super/Subscript  | ``'font_script'``    | :func:`set_font_script()`    |
+------------+------------------+----------------------+------------------------------+
| Number     | Numeric format   | ``'num_format'``     | :func:`set_num_format()`     |
+------------+------------------+----------------------+------------------------------+
| Protection | Lock cells       | ``'locked'``         | :func:`set_locked()`         |
+------------+------------------+----------------------+------------------------------+
|            | Hide formulas    | ``'hidden'``         | :func:`set_hidden()`         |
+------------+------------------+----------------------+------------------------------+
| Alignment  | Horizontal align | ``'align'``          | :func:`set_align()`          |
+------------+------------------+----------------------+------------------------------+
|            | Vertical align   | ``'valign'``         | :func:`set_align()`          |
+------------+------------------+----------------------+------------------------------+
|            | Rotation         | ``'rotation'``       | :func:`set_rotation()`       |
+------------+------------------+----------------------+------------------------------+
|            | Text wrap        | ``'text_wrap'``      | :func:`set_text_wrap()`      |
+------------+------------------+----------------------+------------------------------+
|            | Justify last     | ``'text_justlast'``  | :func:`set_text_justlast()`  |
+------------+------------------+----------------------+------------------------------+
|            | Center across    | ``'center_across'``  | :func:`set_center_across()`  |
+------------+------------------+----------------------+------------------------------+
|            | Indentation      | ``'indent'``         | :func:`set_indent()`         |
+------------+------------------+----------------------+------------------------------+
|            | Shrink to fit    | ``'shrink'``         | :func:`set_shrink()`         |
+------------+------------------+----------------------+------------------------------+
| Pattern    | Cell pattern     | ``'pattern'``        | :func:`set_pattern()`        |
+------------+------------------+----------------------+------------------------------+
|            | Background color | ``'bg_color'``       | :func:`set_bg_color()`       |
+------------+------------------+----------------------+------------------------------+
|            | Foreground color | ``'fg_color'``       | :func:`set_fg_color()`       |
+------------+------------------+----------------------+------------------------------+
| Border     | Cell border      | ``'border'``         | :func:`set_border()`         |
+------------+------------------+----------------------+------------------------------+
|            | Bottom border    | ``'bottom'``         | :func:`set_bottom()`         |
+------------+------------------+----------------------+------------------------------+
|            | Top border       | ``'top'``            | :func:`set_top()`            |
+------------+------------------+----------------------+------------------------------+
|            | Left border      | ``'left'``           | :func:`set_left()`           |
+------------+------------------+----------------------+------------------------------+
|            | Right border     | ``'right'``          | :func:`set_right()`          |
+------------+------------------+----------------------+------------------------------+
|            | Border color     | ``'border_color'``   | :func:`set_border_color()`   |
+------------+------------------+----------------------+------------------------------+
|            | Bottom color     | ``'bottom_color'``   | :func:`set_bottom_color()`   |
+------------+------------------+----------------------+------------------------------+
|            | Top color        | ``'top_color'``      | :func:`set_top_color()`      |
+------------+------------------+----------------------+------------------------------+
|            | Left color       | ``'left_color'``     | :func:`set_left_color()`     |
+------------+------------------+----------------------+------------------------------+
|            | Right color      | ``'right_color'``    | :func:`set_right_color()`    |
+------------+------------------+----------------------+------------------------------+


There are two ways of setting Format properties: by using the object interface
or by setting the property as a dictionary of key/value pairs in the
constructor. For example, a typical use of the object interface would be as
follows::

    format = workbook.add_format()
    format.set_bold()
    format.set_font_color('red')

By comparison the properties can be set by passing a dictionary of properties
to the `add_format()` constructor::

    format = workbook.add_format({'bold': True, 'font_color': 'red'})

The object method interface is mainly provided for backward compatibility. The
key/value interface has proved to be more flexible in real world programs and
is the recommended method for setting format properties.

Format Colors
-------------

Format property colors are specified using a Html sytle ``#RRGGBB`` value or a
imited number of named colors::

    cell_format1.set_font_color('#FF0000')
    cell_format2.set_font_color('red')

See :ref:`colors` for more details.


Format Defaults
---------------

The default Excel 2007+ cell format is Calibri 11 with all other properties off.

In general a format method call without an argument will turn a property on,
for example::

    format1 = workbook.add_format()

    format1.set_bold()   # Turns bold on.
    format1.set_bold(1)  # Also turns bold on.


Since most properties are already off by default it isn't generally required to
turn them off. However, it is possible if required::

    format1.set_bold(0); # Turns bold off.


Modifying Formats
-----------------

Each unique cell format in an XlsxWriter spreadsheet must have a corresponding
Format object. It isn't possible to use a Format with a ``write()`` method and
then redefine it for use at a later stage. This is because a Format is applied
to a cell not in its current state but in its final state. Consider the
following example::

    format = workbook.add_format({'bold': True, 'font_color': 'red'})
    worksheet.write('A1', 'Cell A1', format)

    # Later...
    format.set_font_color('green')
    worksheet.write('B1', 'Cell B1', format)

Cell A1 is assigned a format which is initially has the font set to the colour
red. However, the colour is subsequently set to green. When Excel displays
Cell A1 it will display the final state of the Format which in this case will
be the colour green.

