.. _working_with_formats:

Working with Formats
====================

This section TODO...

Creating and using a Format object
----------------------------------

Cell formatting is defined through a Format object. Format objects are created
by calling the workbook ``add_format()`` method as follows::

    format1 = workbook.add_format()       # Set properties later
    format2 = workbook.add_format(props)  # Set at creation

The format object holds all the formatting properties that can be applied to a
cell, a row or a column. The process of setting these properties is discussed
in the next section.

Once a Format object has been constructed and its properties have been set it
can be passed as an argument to the worksheet ``write`` methods as follows::

    worksheet.write(0, 0, 'One', format)
    worksheet.write_string(1, 0, 'Two', format)
    worksheet.write_number(2, 0, 3, format)
    worksheet.write_blank(3, 0, format)

Formats can also be passed to the worksheet ``set_row()`` and ``set_column()``
methods to define the default property for a row or column.

    worksheet.set_row(0, 15, format) worksheet.set_column(0, 0, 15, format)


Format methods and Format properties
------------------------------------

The following table shows the Excel format categories, the formatting
properties that can be applied and the equivalent object method:

+----------------------------------------------------------------------------+------------------+----------------------+------------------------------+
| Category                                                                   | Description      | Property             | Method Name                  |
+============================================================================+==================+======================+==============================+
| Font                                                                       | Font type        | ``'font_name'``      | :func:`set_font_name()`      |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Font size        | ``'font_size'``      | :func:`set_font_size()`      |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Font color       | ``'font_color'``     | :func:`set_font_color()`     |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Bold             | ``'bold'``           | :func:`set_bold()`           |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Italic           | ``'italic'``         | :func:`set_italic()`         |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Underline        | ``'underline'``      | :func:`set_underline()`      |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Strikeout        | ``'font_strikeout'`` | :func:`set_font_strikeout()` |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Super/Subscript  | ``'font_script'``    | :func:`set_font_script()`    |
+----------------------------------------------------------------------------+------------------+----------------------+------------------------------+
| Number                                                                     | Numeric format   | ``'num_format'``     | :func:`set_num_format()`     |
+----------------------------------------------------------------------------+------------------+----------------------+------------------------------+
| Protection                                                                 | Lock cells       | ``'locked'``         | :func:`set_locked()`         |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Hide formulas    | ``'hidden'``         | :func:`set_hidden()`         |
+----------------------------------------------------------------------------+------------------+----------------------+------------------------------+
| Alignment                                                                  | Horizontal align | ``'align'``          | :func:`set_align()`          |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Vertical align   | ``'valign'``         | :func:`set_align()`          |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Rotation         | ``'rotation'``       | :func:`set_rotation()`       |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Text wrap        | ``'text_wrap'``      | :func:`set_text_wrap()`      |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Justify last     | ``'text_justlast'``  | :func:`set_text_justlast()`  |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Center across    | ``'center_across'``  | :func:`set_center_across()`  |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Indentation      | ``'indent'``         | :func:`set_indent()`         |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Shrink to fit    | ``'shrink'``         | :func:`set_shrink()`         |
+----------------------------------------------------------------------------+------------------+----------------------+------------------------------+
| Pattern                                                                    | Cell pattern     | ``'pattern'``        | :func:`set_pattern()`        |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Background color | ``'bg_color'``       | :func:`set_bg_color()`       |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Foreground color | ``'fg_color'``       | :func:`set_fg_color()`       |
+----------------------------------------------------------------------------+------------------+----------------------+------------------------------+
| Border                                                                     | Cell border      | ``'border'``         | :func:`set_border()`         |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Bottom border    | ``'bottom'``         | :func:`set_bottom()`         |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Top border       | ``'top'``            | :func:`set_top()`            |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Left border      | ``'left'``           | :func:`set_left()`           |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Right border     | ``'right'``          | :func:`set_right()`          |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Border color     | ``'border_color'``   | :func:`set_border_color()`   |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Bottom color     | ``'bottom_color'``   | :func:`set_bottom_color()`   |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Top color        | ``'top_color'``      | :func:`set_top_color()`      |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Left color       | ``'left_color'``     | :func:`set_left_color()`     |
| +------------------+----------------------+------------------------------+ |                                                                        |
|                                                                            | Right color      | ``'right_color'``    | :func:`set_right_color()`    |
+----------------------------------------------------------------------------+------------------+----------------------+------------------------------+


There are two ways of setting Format properties: by using the object method
interface or by setting the property directly. For example, a typical use of
the method interface would be as follows::

    format = workbook.add_format()
    format.set_bold()
    format.set_color('red')

By comparison the properties can be set directly by passing a hash of
properties to the Format constructor::

    format = workbook.add_format(bold, 1, color, 'red')

or after the Format has been constructed by means of the
``set_format_properties()`` method as follows::

    format = workbook.add_format()
    format.set_format_properties(bold, 1, color, 'red')

You can also store the properties in one or more named hashes and pass them to
the required method::

    font = (
        font, 'Calibri',
        size, 12,
        color, 'blue',
        bold, 1,
     )

    shading = (
        bg_color, 'green',
        pattern, 1,
     )


    format1 = workbook.add_format(font); # Font only
    format2 = workbook.add_format(font, shading); # Font and shading

The provision of two ways of setting properties might lead you to wonder which
is the best way. The method mechanism may be better if you prefer setting
properties via method calls (which the author did when the code was first
written) otherwise passing properties to the constructor has proved to be a
little more flexible and self documenting in practice. An additional advantage
of working with property hashes is that it allows you to share formatting
between workbook objects as shown in the example above.


.. _format_colors:

Format Colors
-------------

                        'black' 'blue' 'brown' 'cyan'
                        'gray' 'green' 'lime' 'magenta' 'navy' 'orange' 'pink'
                        'purple' 'red' 'silver' 'white' 'yellow'


Tips for working with formats
-----------------------------

The default format is Calibri 11 with all other properties off.

Each unique format in XlsxWriter must have a corresponding Format object. It
isn't possible to use a Format with a write() method and then redefine the
Format for use at a later stage. This is because a Format is applied to a cell
not in its current state but in its final state. Consider the following
example::

    format = workbook.add_format()
    format.set_bold()
    format.set_color('red')
    worksheet.write('A1', 'Cell A1', format)
    format.set_color('green')
    worksheet.write('B1', 'Cell B1', format)

Cell A1 is assigned the Format ``$format`` which is initially set to the colour
red. However, the colour is subsequently set to green. When Excel displays
Cell A1 it will display the final state of the Format which in this case will
be the colour green.

In general a method call without an argument will turn a property on, for
example::

    format1 = workbook.add_format()
    format1.set_bold(); # Turns bold on
    format1.set_bold(1); # Also turns bold on
    format1.set_bold(0); # Turns bold off
