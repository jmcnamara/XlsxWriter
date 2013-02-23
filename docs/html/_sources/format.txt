.. _format:

The Format Class
================

This section describes the methods and properties that are available for
formatting cells in Excel.

The properties of a cell that can be formatted include: fonts, colours,
patterns, borders, alignment and number formatting.

.. image:: _static/formats_intro.png

format.set_font_name()
----------------------

.. py:function:: set_font_name(fontname)

   Set the font used in the cell.

   :param string fontname: Cell font.

Specify the font used used in the cell format::

    cell_format.set_font_name('Times New Roman')

Excel can only display fonts that are installed on the system that it is
running on. Therefore it is best to use the fonts that come as standard such
as 'Calibri', 'Times New Roman' and 'Courier New'.

The default font for an unformatted cell in Excel 2007+ is 'Calibri'.


format.set_font_size()
----------------------

.. py:function:: set_font_size(size)

   Set the size of the font used in the cell.

   :param int size: The cell font size.

Set the font size of the cell format::

    format = workbook.add_format()
    format.set_font_size(30)

Excel adjusts the height of a row to accommodate the largest font size in the
row. You can also explicitly specify the height of a row using the
:func:`set_row()` worksheet method.


format.set_font_color()
-----------------------

.. py:function:: set_font_color(color)

   Set the color of the font used in the cell.

   :param string color: The cell font color.


Set the font colour::

    format = workbook.add_format()
    
    format.set_font_color('red')
    
    worksheet.write(0, 0, 'wheelbarrow', format)

The color can be a Html style ``#RRGGBB`` string or a limited number of named
colors, see :ref:`format_colors`.

Note: The ``set_font_color()`` method is used to set the colour of the font in
a cell. To set the colour of a cell use the :func:`set_bg_color()` and
:func:`set_pattern()` methods.


format.set_bold()
-----------------

.. py:function:: set_bold()

   Turn on bold for the format font.

Set the bold property of the font::

    format.set_bold()


format.set_italic()
-------------------

.. py:function:: set_italic()

   Turn on italic for the format font.

Set the italic property of the font::

    format.set_italic()


format.set_underline()
----------------------

.. py:function:: set_underline()

   Turn on underline for the format.
   
   :param int style: Underline style.

Set the underline property of the format::

    format.set_underline()

The available underline styles are:

* 1 = Single underline (the default)
* 2 = Double underline
* 33 = Single accounting underline
* 34 = Double accounting underline


format.set_font_strikeout()
---------------------------

.. py:function:: set_font_strikeout()

   Set the strikeout property of the font.


format.set_font_script()
------------------------

.. py:function:: set_font_script()

   Set the superscript/subscript property of the font.

The available options are:

* 1 = Superscript
* 2 = Subscript

format.set_num_format()
-----------------------

.. py:function:: set_num_format(format_string)

   Set the number format for a cell.
   
   :param string format_string: The cell number format.

This method is used to define the numerical format of a number in Excel. It
controls whether a number is displayed as an integer, a floating point number,
a date, a currency value or some other user defined format.

The numerical format of a cell can be specified by using a format string or an
index to one of Excel's built-in formats::

    format1 = workbook.add_format()
    format2 = workbook.add_format()
    
    format1.set_num_format('d mmm yyyy')  # Format string.
    format2.set_num_format(0x0F)          # Format index.

Format strings can control any aspect of number formatting allowed by Excel::

    format01.set_num_format('0.000')
    worksheet.write(1, 0, 3.1415926, format01)       # -> 3.142

    format02.set_num_format('#,##0')
    worksheet.write(2, 0, 1234.56, format02)         # -> 1,235

    format03.set_num_format('#,##0.00')
    worksheet.write(3, 0, 1234.56, format03)         # -> 1,234.56

    format04.set_num_format('0.00')
    worksheet.write(4, 0, 49.99, format04)           # -> 49.99

    format05.set_num_format('mm/dd/yy')
    worksheet.write(5, 0, 36892.521, format05)       # -> 01/01/01

    format06.set_num_format('mmm d yyyy')
    worksheet.write(6, 0, 36892.521, format06)       # -> Jan 1 2001

    format07.set_num_format('d mmmm yyyy')
    worksheet.write(7, 0, 36892.521, format07)       # -> 1 January 2001

    format08.set_num_format('dd/mm/yyyy hh:mm AM/PM')
    worksheet.write(8, 0, 36892.521, format08)      # -> 01/01/2001 12:30 AM

    format09.set_num_format('0 "dollar and" .00 "cents"')
    worksheet.write(9, 0, 1.87, format09)           # -> 1 dollar and .87 cents

    # Conditional numerical formatting.
    format10.set_num_format('[Green]General;[Red]-General;General')
    worksheet.write(10, 0, 123, format10)  # > 0 Green
    worksheet.write(11, 0, -45, format10)  # < 0 Red
    worksheet.write(12, 0,   0, format10)  # = 0 Default colour

    # Zip code.
    format11.set_num_format('00000')
    worksheet.write(13, 0, 1209, format11)

.. image:: _static/formats_num_str.png

The number system used for dates is described in
:ref:`working_with_dates_and_time`.

The colour format should have one of the following values::

    [Black] [Blue] [Cyan] [Green] [Magenta] [Red] [White] [Yellow]

For more information refer to the
`Microsoft documentation on cell formats <http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx>`_.

Excel's built-in formats are shown in the following table:

+-------+-------+--------------------------------------------------------+
| Index | Index | Format String                                          |
+=======+=======+========================================================+
| 0     | 0x00  | ``General``                                            |
+-------+-------+--------------------------------------------------------+
| 1     | 0x01  | ``0``                                                  |
+-------+-------+--------------------------------------------------------+
| 2     | 0x02  | ``0.00``                                               |
+-------+-------+--------------------------------------------------------+
| 3     | 0x03  | ``#,##0``                                              |
+-------+-------+--------------------------------------------------------+
| 4     | 0x04  | ``#,##0.00``                                           |
+-------+-------+--------------------------------------------------------+
| 5     | 0x05  | ``($#,##0_);($#,##0)``                                 |
+-------+-------+--------------------------------------------------------+
| 6     | 0x06  | ``($#,##0_);[Red]($#,##0)``                            |
+-------+-------+--------------------------------------------------------+
| 7     | 0x07  | ``($#,##0.00_);($#,##0.00)``                           |
+-------+-------+--------------------------------------------------------+
| 8     | 0x08  | ``($#,##0.00_);[Red]($#,##0.00)``                      |
+-------+-------+--------------------------------------------------------+
| 9     | 0x09  | ``0%``                                                 |
+-------+-------+--------------------------------------------------------+
| 10    | 0x0a  | ``0.00%``                                              |
+-------+-------+--------------------------------------------------------+
| 11    | 0x0b  | ``0.00E+00``                                           |
+-------+-------+--------------------------------------------------------+
| 12    | 0x0c  | ``# ?/?``                                              |
+-------+-------+--------------------------------------------------------+
| 13    | 0x0d  | ``# ??/??``                                            |
+-------+-------+--------------------------------------------------------+
| 14    | 0x0e  | ``m/d/yy``                                             |
+-------+-------+--------------------------------------------------------+
| 15    | 0x0f  | ``d-mmm-yy``                                           |
+-------+-------+--------------------------------------------------------+
| 16    | 0x10  | ``d-mmm``                                              |
+-------+-------+--------------------------------------------------------+
| 17    | 0x11  | ``mmm-yy``                                             |
+-------+-------+--------------------------------------------------------+
| 18    | 0x12  | ``h:mm AM/PM``                                         |
+-------+-------+--------------------------------------------------------+
| 19    | 0x13  | ``h:mm:ss AM/PM``                                      |
+-------+-------+--------------------------------------------------------+
| 20    | 0x14  | ``h:mm``                                               |
+-------+-------+--------------------------------------------------------+
| 21    | 0x15  | ``h:mm:ss``                                            |
+-------+-------+--------------------------------------------------------+
| 22    | 0x16  | ``m/d/yy h:mm``                                        |
+-------+-------+--------------------------------------------------------+
| ...   | ...   | ...                                                    |
+-------+-------+--------------------------------------------------------+
| 37    | 0x25  | ``(#,##0_);(#,##0)``                                   |
+-------+-------+--------------------------------------------------------+
| 38    | 0x26  | ``(#,##0_);[Red](#,##0)``                              |
+-------+-------+--------------------------------------------------------+
| 39    | 0x27  | ``(#,##0.00_);(#,##0.00)``                             |
+-------+-------+--------------------------------------------------------+
| 40    | 0x28  | ``(#,##0.00_);[Red](#,##0.00)``                        |
+-------+-------+--------------------------------------------------------+
| 41    | 0x29  | ``_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)``            |
+-------+-------+--------------------------------------------------------+
| 42    | 0x2a  | ``_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)``         |
+-------+-------+--------------------------------------------------------+
| 43    | 0x2b  | ``_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)``    |
+-------+-------+--------------------------------------------------------+
| 44    | 0x2c  | ``_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`` |
+-------+-------+--------------------------------------------------------+
| 45    | 0x2d  | ``mm:ss``                                              |
+-------+-------+--------------------------------------------------------+
| 46    | 0x2e  | ``[h]:mm:ss``                                          |
+-------+-------+--------------------------------------------------------+
| 47    | 0x2f  | ``mm:ss.0``                                            |
+-------+-------+--------------------------------------------------------+
| 48    | 0x30  | ``##0.0E+0``                                           |
+-------+-------+--------------------------------------------------------+
| 49    | 0x31  | ``@``                                                  |
+-------+-------+--------------------------------------------------------+

.. note::
   Numeric formats 23 to 36 are not documented by Microsoft and
   may differ in international versions.
.. note::
   The dollar sign appears as the defined local currency symbol.


format.set_locked()
-------------------

.. py:function:: set_locked(state)

   Set the cell locked state.
   
   :param bool state: Turn cell locking on or off. Defaults to True.

This property can be used to prevent modification of a cells contents.
Following Excel's convention, cell locking is turned on by default. However,
it only has an effect if the worksheet has been protected, see the worksheet
``protect()`` method (not implemented yet)::

    locked = workbook.add_format()
    locked.set_locked(True)

    unlocked = workbook.add_format()
    locked.set_locked(False)

    # Enable worksheet protection
    worksheet.protect()

    # This cell cannot be edited.
    worksheet.write('A1', '=1+2', locked)

    # This cell can be edited.
    worksheet.write('A2', '=1+2', unlocked)


format.set_hidden()
-------------------

.. py:function:: set_hidden()

   Hide formulas in a cell.
  

This property is used to hide a formula while still displaying its result. This
is generally used to hide complex calculations from end users who are only
interested in the result. It only has an effect if the worksheet has been
protected, see the worksheet ``protect()`` method (not implemented yet)::

    hidden = workbook.add_format()
    hidden.set_hidden()

    # Enable worksheet protection
    worksheet.protect()

    # The formula in this cell isn't visible
    worksheet.write('A1', '=1+2', hidden)


format.set_align()
------------------

.. py:function:: set_align(alignment)

   Set the alignment for data in the cell.

   :param string alignment: The vertical and or horizontal alignment direction.

This method is used to set the horizontal and vertical text alignment within a
cell. The following are the available horizontal alignments:

+----------------------+
| Horizontal alignment |
+======================+
| center               |
+----------------------+
| right                |
+----------------------+
| fill                 |
+----------------------+
| justify              |
+----------------------+
| center_across        |
+----------------------+

The following are the available vertical alignments:

+----------------------+
| Vertical alignment   |
+======================+
| top                  |
+----------------------+
| vcenter              |
+----------------------+
| bottom               |
+----------------------+
| vjustify             |
+----------------------+


As in Excel, vertical and horizontal alignments can be combined::

    format = workbook.add_format()
    
    format.set_align('center')
    format.set_align('vcenter')
    
    worksheet.set_row(0, 30)
    worksheet.write(0, 0, 'Some Text', format)

Text can be aligned across two or more adjacent cells using the
``'center_across'`` property. However, for genuine merged cells it is better
to use the ``merge_range()`` worksheet method (not implemented yet).

The ``'vjustify'`` (vertical justify) option can be used to provide automatic
text wrapping in a cell. The height of the cell will be adjusted to
accommodate the wrapped text. To specify where the text wraps use the
``set_text_wrap()`` method.


format.set_center_across()
--------------------------

.. py:function:: set_center_across()

   Centre text across adjacent cells.

Text can be aligned across two or more adjacent cells using the
``set_center_across()`` method. This is an alias for the
``set_align('center_across')`` method call.

Only one cell should contain the text, the other cells should be blank::

    format = workbook.add_format()
    format.set_center_across()

    worksheet.write(1, 1, 'Center across selection', format)
    worksheet.write_blank(1, 2, format)

For actual merged cells it is better to use the ``merge_range()`` worksheet
method.


format.set_text_wrap()
----------------------

.. py:function:: set_text_wrap()

   Wrap text in a cell.

Turn text wrapping on for text in a cell::

    format = workbook.add_format()
    format.set_text_wrap()

    worksheet.write(0, 0, "Some long text to wrap in a cell", format)

If you wish to control where the text is wrapped you can add newline characters
to the string::

    format = workbook.add_format()
    format.set_text_wrap()

    worksheet.write(0, 0, "It's\na bum\nwrap", format)

Excel will adjust the height of the row to accommodate the wrapped text. A
similar effect can be obtained without newlines using the
``set_align('vjustify')`` method.


format.set_rotation()
---------------------

.. py:function:: set_rotation(angle)

   Set the rotation of the text in a cell.

   :param int angle: Rotation angle in the range -90 to 90 and 270.

Set the rotation of the text in a cell. The rotation can be any angle in the
range -90 to 90 degrees::

    format = workbook.add_format()
    format.set_rotation(30)

    worksheet.write(0, 0, 'This text is rotated', format)

The angle 270 is also supported. This indicates text where the letters run from
top to bottom.


format.set_indent()
-------------------

.. py:function:: set_indent(level)

   Set the cell text indentation level.

   :param int level: Indentation level.

This method can be used to indent text in a cell. The argument, which should be
an integer, is taken as the level of indentation::

    format = workbook.add_format()
    format.set_indent(2)

    worksheet.write(0, 0, 'This text is indented', format)

Indentation is a horizontal alignment property. It will override any other
horizontal properties but it can be used in conjunction with vertical
properties.


format.set_shrink()
-------------------

.. py:function:: set_shrink()

   Turn on the text "shrink to fit" for a cell.

This method can be used to shrink text so that it fits in a cell::

    format = workbook.add_format()
    format.set_shrink()

    worksheet.write(0, 0, 'Honey, I shrunk the text!', format)


format.set_text_justlast()
--------------------------

.. py:function:: set_text_justlast()

   Turn on the justify last text property.

Only applies to Far Eastern versions of Excel.


format.set_pattern()
--------------------

.. py:function:: set_pattern(index)

   :param int index: Pattern index. 0 - 18.

Set the background pattern of a cell.

The most common pattern is 1 which is a solid fill of the background color.


format.set_bg_color()
---------------------

.. py:function:: set_bg_color(color)

   Set the color of the background pattern in a cell.

   :param string color: The cell font color.

The ``set_bg_color()`` method can be used to set the background colour of a
pattern. Patterns are defined via the ``set_pattern()`` method. If a pattern
hasn't been defined then a solid fill pattern is used as the default.

Here is an example of how to set up a solid fill in a cell::

    format = workbook.add_format()

    format.set_pattern(1)  # This is optional when using a solid fill.
    format.set_bg_color('green')
    
    worksheet.write('A1', 'Ray', format)

.. image:: _static/formats_set_bg_color.png

The color can be a Html style ``#RRGGBB`` string or a limited number of named
colors, see :ref:`format_colors`.



format.set_fg_color()
---------------------

.. py:function:: set_fg_color(color)

   Set the color of the foreground pattern in a cell.

   :param string color: The cell font color.

The ``set_fg_color()`` method can be used to set the foreground colour of a
pattern.

The color can be a Html style ``#RRGGBB`` string or a limited number of named
colors, see :ref:`format_colors`.



format.set_border()
-------------------

.. py:function:: set_border(style)
   
   Set the cell border style.

   :param int style: Border style index. Default is 1.

Individual border elements can be configured using the following methods with
the same parameters:

* :func:`set_bottom()`
* :func:`set_top()`
* :func:`set_left()`
* :func:`set_right()`

A cell border is comprised of a border on the bottom, top, left and right.
These can be set to the same value using ``set_border()`` or individually
using the relevant method calls shown above.

The following shows the border styles sorted by XlsxWriter index number:

+-------+---------------+--------+-----------------+
| Index | Name          | Weight | Style           |
+=======+===============+========+=================+
| 0     | None          | 0      |                 |
+-------+---------------+--------+-----------------+
| 1     | Continuous    | 1      | ``-----------`` |
+-------+---------------+--------+-----------------+
| 2     | Continuous    | 2      | ``-----------`` |
+-------+---------------+--------+-----------------+
| 3     | Dash          | 1      | ``- - - - - -`` |
+-------+---------------+--------+-----------------+
| 4     | Dot           | 1      | ``. . . . . .`` |
+-------+---------------+--------+-----------------+
| 5     | Continuous    | 3      | ``-----------`` |
+-------+---------------+--------+-----------------+
| 6     | Double        | 3      | ``===========`` |
+-------+---------------+--------+-----------------+
| 7     | Continuous    | 0      | ``-----------`` |
+-------+---------------+--------+-----------------+
| 8     | Dash          | 2      | ``- - - - - -`` |
+-------+---------------+--------+-----------------+
| 9     | Dash Dot      | 1      | ``- . - . - .`` |
+-------+---------------+--------+-----------------+
| 10    | Dash Dot      | 2      | ``- . - . - .`` |
+-------+---------------+--------+-----------------+
| 11    | Dash Dot Dot  | 1      | ``- . . - . .`` |
+-------+---------------+--------+-----------------+
| 12    | Dash Dot Dot  | 2      | ``- . . - . .`` |
+-------+---------------+--------+-----------------+
| 13    | SlantDash Dot | 2      | ``/ - . / - .`` |
+-------+---------------+--------+-----------------+

The following shows the borders in the order shown in the Excel Dialog:

+-------+-----------------+-------+-----------------+
| Index | Style           | Index | Style           |
+=======+=================+=======+=================+
| 0     | None            | 12    | ``- . . - . .`` |
+-------+-----------------+-------+-----------------+
| 7     | ``-----------`` | 13    | ``/ - . / - .`` |
+-------+-----------------+-------+-----------------+
| 4     | ``. . . . . .`` | 10    | ``- . - . - .`` |
+-------+-----------------+-------+-----------------+
| 11    | ``- . . - . .`` | 8     | ``- - - - - -`` |
+-------+-----------------+-------+-----------------+
| 9     | ``- . - . - .`` | 2     | ``-----------`` |
+-------+-----------------+-------+-----------------+
| 3     | ``- - - - - -`` | 5     | ``-----------`` |
+-------+-----------------+-------+-----------------+
| 1     | ``-----------`` | 6     | ``===========`` |
+-------+-----------------+-------+-----------------+


format.set_bottom()
-------------------

.. py:function:: set_bottom(style)
   
   Set the cell bottom border style.

   :param int style: Border style index. Default is 1.

Set the cell bottom border style. See :func:`set_border` for details on the
border styles.


format.set_top()
----------------

.. py:function:: set_top(style)
   
   Set the cell top border style.

   :param int style: Border style index. Default is 1.

Set the cell top border style. See :func:`set_border` for details on the border
styles.


format.set_left()
-----------------

.. py:function:: set_left(style)
   
   Set the cell left border style.

   :param int style: Border style index. Default is 1.

Set the cell left border style. See :func:`set_border` for details on the
border styles.


format.set_right()
------------------

.. py:function:: set_right(style)
   
   Set the cell right border style.

   :param int style: Border style index. Default is 1.

Set the cell right border style. See :func:`set_border` for details on the
border styles.


format.set_border_color()
-------------------------

.. py:function:: set_border_color(color)

   Set the color of the cell border.

   :param string color: The cell border color.
   
Individual border elements can be configured using the following methods with
the same parameters:

* :func:`set_bottom_color()`
* :func:`set_top_color()`
* :func:`set_left_color()`
* :func:`set_right_color()`

Set the colour of the cell borders. A cell border is comprised of a border on
the bottom, top, left and right. These can be set to the same colour using
``set_border_color()`` or individually using the relevant method calls shown
above.

The color can be a Html style ``#RRGGBB`` string or a limited number of named
colors, see :ref:`format_colors`.


format.set_bottom_color()
-------------------------

.. py:function:: set_bottom_color(color)

   Set the color of the bottom cell border.

   :param string color: The cell border color.

See :func:`set_border_color` for details on the border colors.


format.set_top_color()
----------------------

.. py:function:: set_top_color(color)

   Set the color of the top cell border.

   :param string color: The cell border color.

See :func:`set_border_color` for details on the border colors.


format.set_left_color()
-----------------------

.. py:function:: set_left_color(color)

   Set the color of the left cell border.

   :param string color: The cell border color.

See :func:`set_border_color` for details on the border colors.


format.set_right_color()
------------------------

.. py:function:: set_right_color(color)

   Set the color of the right cell border.

   :param string color: The cell border color.

See :func:`set_border_color` for details on the border colors.


