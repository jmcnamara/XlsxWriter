.. _worksheet:

The Worksheet Class
===================

The worksheet class represents an Excel worksheet. It handles operations such
as writing data to cells or formatting worksheet layout.

A worksheet object isn't instantiated directly. Instead a new worksheet is
created by calling the :func:`add_worksheet()` method from a :func:`Workbook`
object::

    workbook   = Workbook('filename.xlsx')

    worksheet1 = workbook.add_worksheet()
    worksheet2 = workbook.add_worksheet()


worksheet.write()
-----------------

.. py:function:: write(row, col, data[, cell_format])

   Write generic data to a worksheet cell.

   :param row:         The cell row (zero indexed).
   :param col:         The cell column (zero indexed).
   :param data:        Cell data to write. Variable types.
   :param cell_format: Optional Format object.
   :type  row:         integer
   :type  col:         integer
   :type  cell_format: :ref:`Format <format>`

Excel makes a distinction between data types such as strings, numbers,
blanks, formulas and hyperlinks. To simplify the process of writing
data to an XlsxWriter file the ``write()`` method acts as a general
alias for several more specific methods:

* :func:`write_string()`
* :func:`write_number()`
* :func:`write_blank()`
* :func:`write_formula()`

The general rule is that if the data looks like a *something* then a
*something* is written. Here are some examples::


    worksheet.write(0, 0, 'Hello')          # write_string()
    worksheet.write(1, 0, 'One')            # write_string()
    worksheet.write(2, 0, 2)                # write_number()
    worksheet.write(3, 0, 3.00001)          # write_number()
    worksheet.write(4, 0, '')               # write_blank()
    worksheet.write(5, 0, None)             # write_blank()
    worksheet.write(6, 0, '=SIN(PI()/4)')   # write_formula()


The ``write()`` method supports two forms of notation to designate the
position of cells: **Row-column** notation and **A1** notation::

    # These are equivalent.
    worksheet.write(0, 0, 'Hello')
    worksheet.write('A1', 'Hello')

See :ref:`cell_notation` for more details.


The ``cell_format`` parameter is optional. It should be a valid :ref:`Format
<format>` object::

    cell_format = workbook.add_format({'bold': True, 'italic': True})

    worksheet.write(0, 0, 'Hello', cell_format)  # Cell is bold and italic.

The ``write()`` method will ignore empty strings or ``None`` unless a
format is also supplied. As such you needn't worry about special
handling for empty or ``None`` values in your data. See also the
:func:`write_blank()` method.


One problem with the ``write()`` method is that occasionally data
looks like a number but you don't want it treated as a number. For
example, zip codes or ID numbers often start with a leading zero. If
you write this data as a number then the leading zero(s) will be
stripped. In this case you shouldn't use the ``write()`` method and
should use ``write_string()`` instead.


worksheet.write_string()
------------------------

.. py:function:: write_string(row, col, string[, cell_format])

   Write a string to a worksheet cell.

   :param row:         The cell row (zero indexed).
   :param col:         The cell column (zero indexed).
   :param string:      String to write to cell.
   :param cell_format: Optional Format object.
   :type  row:         integer
   :type  col:         integer
   :type  string:      string
   :type  cell_format: :ref:`Format <format>`

The ``write_string()`` method writes a string to the cell specified by
``row`` and ``column``::

    worksheet.write_string(0, 0, 'Your text here')
    worksheet.write_string('A2', 'or here')

Both row-column and A1 style notation are support. See
:ref:`cell_notation` for more details.

The ``cell_format`` parameter is optional. It should be a valid :ref:`Format
<format>` object.

The maximum string size supported by Excel is 32,767 characters. Strings
longer than this will be truncated by ``write_string()``.

.. note::
   Even though Excel allows strings of 32,767 characters in a cell, the
   maximum number of characters that Excel can **display** in a cell is 1000.
   However, all 32,767 characters can be displayed in the formula bar.

In general it is sufficient to use the ``write()`` method when dealing
with string data. However, you may sometimes need to use
``write_string()`` to write data that looks like a number but that
you don't want treated as a number. For example, zip codes or phone
numbers::

    # Write ID number as a plain string.
    worksheet.write_string('A1', '01209')

However, if the user edits this string Excel may convert it back to a
number. To get around this you can use the Excel text format ``'@'``::

    # Format as a string. Doesn't change to a number when edited
    str_format = workbook.add_format({'num_format', '@'})
    worksheet.write_string('A1', '01209', str_format)

This behaviour, while slightly tedious, is unfortunately consistent
with the way Excel handles string data that looks like numbers.


worksheet.write_number()
------------------------

.. py:function:: write_number(row, col, number[, cell_format])

   Write a number to a worksheet cell.

   :param row:         The cell row (zero indexed).
   :param col:         The cell column (zero indexed).
   :param number:      Number to write to cell.
   :param cell_format: Optional Format object.
   :type  row:         integer
   :type  col:         integer
   :type  number:      int or float
   :type  cell_format: :ref:`Format <format>`

The ``write_number()`` method writes an integer or a float to the cell
specified by ``row`` and ``column``::

    worksheet.write_number(0, 0, 123456)
    worksheet.write_number('A2', 2.3451)

Both row-column and A1 style notation are support. See
:ref:`cell_notation` for more details.

The ``cell_format`` parameter is optional. It should be a valid :ref:`Format
<format>` object.


worksheet.write_formula()
-------------------------

.. py:function:: write_formula(row, col, formula[, cell_format[, value]])

   Write a formula to a worksheet cell.

   :param row:         The cell row (zero indexed).
   :param col:         The cell column (zero indexed).
   :param formula:     Formula to write to cell.
   :param cell_format: Optional Format object.
   :type  row:         integer
   :type  col:         integer
   :type  formula:     string
   :type  cell_format: :ref:`Format <format>`

The ``write_formula()`` method writes a formula or function to the
cell specified by ``row`` and ``column``::

    worksheet.write_formula(0, 0, '=B3 + B4')
    worksheet.write_formula(1, 0, '=SIN(PI()/4)')
    worksheet.write_formula(2, 0, '=SUM(B1:B5)')
    worksheet.write_formula('A4', '=IF(A3>1,"Yes", "No")')
    worksheet.write_formula('A5', '=AVERAGE(1, 2, 3, 4)')
    worksheet.write_formula('A6', '=DATEVALUE("1-Jan-2001")')

Array formulas are also supported::

    worksheet.write_formula('A7', '{=SUM(A1:B1*A2:B2)}')

See also the ``write_array_formula()`` method below.

Both row-column and A1 style notation are support. See
:ref:`cell_notation` for more details.

The ``cell_format`` parameter is optional. It should be a valid :ref:`Format
<format>` object.

XlsxWriter doesn't calculate the value of a formula and instead stores
the value 0 as the formula results. It then sets a global flag in the
Xlsx file to say that all formulas and functions should be
recalculated when the file is opened. This is the method recommended
in the Excel documentation and in general it works fine with
spreadsheet applications. However, applications that don't have a
facility to calculate formulas, such as Excel Viewer, or some mobile
applications will only display the 0 results.

If required, it is also possible to specify the calculated result of
the formula using the options ``value`` parameter. This is
occasionally necessary when working with non-Excel applications that
don't calculate the value of the formula. The calculated ``value`` is
added at the end of the argument list::

    worksheet.write('A1', '=2+2', num_format, 4)

.. note::
   Some versions of Excel 2007 do not display the calculated values of
   formulas written by XlsxWriter. Applying all available Service
   Packs to Excel should fix this.


worksheet.write_array_formula()
-------------------------------

.. py:function:: write_array_formula(first_row, first_col, last_row, \
                                    last_col, formula[, cell_format[, value]])

   Write an array formula to a worksheet cell.

   :param first_row:   The first row of the range. (All zero indexed.)
   :param first_col:   The first column of the range.
   :param last_row:    The last row of the range.
   :param last_col:    The last col of the range.
   :param formula:     Array formula to write to cell.
   :param cell_format: Optional Format object.
   :type  first_row:   integer
   :type  first_col:   integer
   :type  last_row:    integer
   :type  last_col:    integer
   :type  formula:     string
   :type  cell_format: :ref:`Format <format>`

The ``write_array_formula()`` method write an array formula to a cell
range. In Excel an array formula is a formula that performs a
calculation on a set of values. It can return a single value or a
range of values.

An array formula is indicated by a pair of braces around the formula:
``{=SUM(A1:B1*A2:B2)}``. If the array formula returns a single value
then the ``first_`` and ``last_`` parameters should be the same::

    worksheet.write_array_formula('A1:A1', '{=SUM(B1:C1*B2:C2)}')

It this case however it is easier to just use the ``write_formula()``
or ``write()`` methods::


    # Same as above but more concise.
    worksheet.write('A1', '{=SUM(B1:C1*B2:C2)}')
    worksheet.write_formula('A1', '{=SUM(B1:C1*B2:C2)}')

For array formulas that return a range of values you must specify the
range that the return values will be written to::



    worksheet.write_array_formula('A1:A3',    '{=TREND(C1:C3,B1:B3)}')
    worksheet.write_array_formula(0, 0, 2, 0, '{=TREND(C1:C3,B1:B3)}')

As shown above, both row-column and A1 style notation are support. See
:ref:`cell_notation` for more details.

The ``cell_format`` parameter is optional. It should be a valid :ref:`Format
<format>` object.

If required, it is also possible to specify the calculated value of
the formula. This is occasionally necessary when working with
non-Excel applications that don't calculate the value of the
formula. The calculated ``value`` is added at the end of the argument
list::


    worksheet.write_array_formula('A1:A3', '{=TREND(C1:C3,B1:B3)}', format, 105)

In addition, some early versions of Excel 2007 don't calculate the
values of array formulas when they aren't supplied. Installing the
latest Office Service Pack should fix this issue.



worksheet.write_blank()
-----------------------

.. py:function:: write_blank(row, col, blank[, cell_format])

   Write a blank worksheet cell.

   :param row:         The cell row (zero indexed).
   :param col:         The cell column (zero indexed).
   :param blank:       None or empty string. The value is ignored.
   :param cell_format: Optional Format object.
   :type  row:         integer
   :type  col:         integer
   :type  cell_format: :ref:`Format <format>`

Write a blank cell specified by ``row`` and ``column``::

    worksheet.write_blank(0, 0, None, format)

This method is used to add formatting to a cell which doesn't contain
a string or number value.


Excel differentiates between an "Empty" cell and a "Blank" cell. An
"Empty" cell is a cell which doesn't contain data whilst a "Blank"
cell is a cell which doesn't contain data but does contain
formatting. Excel stores "Blank" cells but ignores "Empty" cells.


As such, if you write an empty cell without formatting it is ignored::

    worksheet.write('A1', None, format)  # write_blank()
    worksheet.write('A2', None)  # Ignored

This seemingly uninteresting fact means that you can write arrays of
data without special treatment for ``None`` or empty string values.


See the note about "Cell notation".


worksheet.write_datetime()
--------------------------

.. py:function:: write_datetime(row, col, datetime [, cell_format])

   Write a date or time to a worksheet cell.

   :param row:         The cell row (zero indexed).
   :param col:         The cell column (zero indexed).
   :param datetime:    A datetime object.
   :param cell_format: Optional Format object.
   :type  row:         integer
   :type  col:         integer
   :type  formula:     string
   :type  datetime:    :py:mod:`datetime`
   :type  cell_format: :ref:`Format <format>`


The ``write_datetime()`` method can be used to write a date or time to
the cell specified by ``row`` and ``column``::


    worksheet.write_datetime('A1', '2004-05-13T23:20', date_format)

The ``date_string`` should be in the following format::

    yyyy-mm-ddThh:mm:ss.sss

This conforms to an ISO8601 date but it should be noted that the full
range of ISO8601 formats are not supported.


The following variations on the ``date_string`` parameter are permitted::

    yyyy-mm-ddThh:mm:ss.sss # Standard format
    yyyy-mm-ddT # No time
              Thh:mm:ss.sss # No date
    yyyy-mm-ddThh:mm:ss.sssZ # Additional Z (but not time zones)
    yyyy-mm-ddThh:mm:ss # No fractional seconds
    yyyy-mm-ddThh:mm # No seconds

Note that the ``T`` is required in all cases.

A date should always have a ``cell_format``, otherwise it will appear as
a number, see "DATES AND TIME IN EXCEL" and :ref:`format`. Here is a
typical example::

    date_format = workbook.add_format(num_format, 'mm/dd/yy')
    worksheet.write_datetime('A1', '2004-05-13T23:20', date_format)

Valid dates should be in the range 1900-01-01 to 9999-12-31, for the
1900 epoch and 1904-01-01 to 9999-12-31, for the 1904 epoch. As with
Excel, dates outside these ranges will be written as a string.


See also the datetime.pl program in the ``examples`` directory of the
distro.



worksheet.set_row()
-------------------

.. py:function:: set_row( row, height, cell_format, hidden, level, collapsed )

This method can be used to change the default properties of a row. All
parameters apart from ``row`` are optional.


The most common use for this method is to change the height of a row::

    worksheet.set_row(0, 20)  # Row 1 height set to 20

If you wish to set the format without changing the height you can pass
``None`` as the height parameter::


    worksheet.set_row(0, None, format)

The ``cell_format`` parameter will be applied to any cells in the row
that don't have a format. For example::


    worksheet.set_row(0, None, format1)  # Set the format for row 1
    worksheet.write('A1', 'Hello')  # Defaults to format1
    worksheet.write('B1', 'Hello', format2)  # Keeps format2

If you wish to define a row format in this way you should call the
method before any calls to ``write()``. Calling it afterwards will
overwrite any format that was previously specified.


The ``hidden`` parameter should be set to 1 if you wish to hide a
row. This can be used, for example, to hide intermediary steps in a
complicated calculation::


    worksheet.set_row(0, 20, format, 1)
    worksheet.set_row(1, None, None, 1)

The ``level`` parameter is used to set the outline level of the
row. Outlines are described in "OUTLINES AND GROUPING IN
EXCEL". Adjacent rows with the same outline level are grouped together
into a single outline.


The following example sets an outline level of 1 for rows 1 and 2 (zero-indexed)::

    worksheet.set_row(1, None, None, 0, 1)
    worksheet.set_row(2, None, None, 0, 1)

The ``hidden`` parameter can also be used to hide collapsed outlined
rows when used in conjunction with the ``level`` parameter.


    worksheet.set_row(1, None, None, 1, 1)
    worksheet.set_row(2, None, None, 1, 1)

For collapsed outlines you should also indicate which row has the
collapsed ``+`` symbol using the optional ``collapsed`` parameter.


    worksheet.set_row(3, None, None, 0, 0, 1)

For a more complete example see the ``outline.pl`` and
``outline_collapsed.pl`` programs in the examples directory of the
distro.


Excel allows up to 7 outline levels. Therefore the ``level`` parameter
should be in the range ``0 <= level <= 7``.



worksheet.set_column()
----------------------

.. py:function:: set_column( first_col, last_col, width, cell_format, hidden, level, collapsed )

This method can be used to change the default properties of a single
column or a range of columns. All parameters apart from ``first_col``
and ``last_col`` are optional.


If ``set_column()`` is applied to a single column the value of
``first_col`` and ``last_col`` should be the same. In the case where
``last_col`` is zero it is set to the same value as ``first_col``.


It is also possible, and generally clearer, to specify a column range
using the form of A1 notation used for columns. See the note about
"Cell notation".


Examples::

    worksheet.set_column(0, 0, 20)  # Column A   width set to 20
    worksheet.set_column(1, 3, 30)  # Columns B-D width set to 30
    worksheet.set_column('E:E', 20)  # Column E   width set to 20
    worksheet.set_column('F:H', 30)  # Columns F-H width set to 30

The width corresponds to the column width value that is specified in
Excel. It is approximately equal to the length of a string in the
default font of Calibri 11. Unfortunately, there is no way to specify
"AutoFit" for a column in the Excel file format. This feature is only
available at runtime from within Excel.


As usual the ``cell_format`` parameter is optional, for additional
information, see "CELL FORMATTING". If you wish to set the format
without changing the width you can pass ``None`` as the width
parameter::


    worksheet.set_column(0, 0, None, format)

The ``cell_format`` parameter will be applied to any cells in the
column that don't have a format. For example::


    worksheet.set_column('A:A', None, format1)  # Set format for col 1
    worksheet.write('A1', 'Hello')  # Defaults to format1
    worksheet.write('A2', 'Hello', format2)  # Keeps format2

If you wish to define a column format in this way you should call the
method before any calls to ``write()``. If you call it afterwards it
won't have any effect.


A default row format takes precedence over a default column format::

    worksheet.set_row(0, None, format1)  # Set format for row 1
    worksheet.set_column('A:A', None, format2)  # Set format for col 1
    worksheet.write('A1', 'Hello')  # Defaults to format1
    worksheet.write('A2', 'Hello')  # Defaults to format2

The ``hidden`` parameter should be set to 1 if you wish to hide a
column. This can be used, for example, to hide intermediary steps in a
complicated calculation::


    worksheet.set_column('D:D', 20, format, 1)
    worksheet.set_column('E:E', None, None, 1)

The ``level`` parameter is used to set the outline level of the
column. Outlines are described in "OUTLINES AND GROUPING IN
EXCEL". Adjacent columns with the same outline level are grouped
together into a single outline.


The following example sets an outline level of 1 for columns B to G::

    worksheet.set_column('B:G', None, None, 0, 1)

The ``hidden`` parameter can also be used to hide collapsed outlined
columns when used in conjunction with the ``level`` parameter.


    worksheet.set_column('B:G', None, None, 1, 1)

For collapsed outlines you should also indicate which row has the
collapsed ``+`` symbol using the optional ``collapsed`` parameter.


    worksheet.set_column('H:H', None, None, 0, 0, 1)

For a more complete example see the ``outline.pl`` and
``outline_collapsed.pl`` programs in the examples directory of the
distro.


Excel allows up to 7 outline levels. Therefore the ``level`` parameter
should be in the range ``0 <= level <= 7``.


worksheet.activate()
--------------------

.. py:function:: activate()

The ``activate()`` method is used to specify which worksheet is
initially visible in a multi-sheet workbook::


    worksheet1 = workbook.add_worksheet('To')
    worksheet2 = workbook.add_worksheet('the')
    worksheet3 = workbook.add_worksheet('wind')

    worksheet3.activate()

This is similar to the Excel VBA activate method. More than one
worksheet can be selected via the ``select()`` method, see below,
however only one worksheet can be active.


The default active worksheet is the first worksheet.


worksheet.select()
------------------

.. py:function:: select()

The ``select()`` method is used to indicate that a worksheet is
selected in a multi-sheet workbook::


    worksheet1.activate()
    worksheet2.select()
    worksheet3.select()

A selected worksheet has its tab highlighted. Selecting worksheets is
a way of grouping them together so that, for example, several
worksheets could be printed in one go. A worksheet that has been
activated via the ``activate()`` method will also appear as selected.
