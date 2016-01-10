.. _cell_notation:

Working with Cell Notation
==========================

XlsxWriter supports two forms of notation to designate the position of cells:
**Row-column** notation and **A1** notation.

Row-column notation uses a zero based index for both row and column while A1
notation uses the standard Excel alphanumeric sequence of column letter and
1-based row. For example::


    (0, 0)      # Row-column notation.
    ('A1')      # The same cell in A1 notation.

    (6, 2)      # Row-column notation.
    ('C7')      # The same cell in A1 notation.

Row-column notation is useful if you are referring to cells programmatically::

    for row in range(0, 5):
        worksheet.write(row, 0, 'Hello')

A1 notation is useful for setting up a worksheet manually and for working with
formulas::

    worksheet.write('H1', 200)
    worksheet.write('H2', '=H1+1')

In general when using the XlsxWriter module you can use A1 notation anywhere
you can use row-column notation.

XlsxWriter supports Excels worksheet limits of 1,048,576 rows by 16,384
columns.

.. note::
   Ranges in A1 notation must be in uppercase, like in Excel.

.. note::
   In Excel it is also possible to use R1C1 notation. This is not
   supported by XlsxWriter.


.. _abs_reference:

Relative and Absolute cell references
-------------------------------------

When dealing with Excel cell references it is important to distinguish between
relative and absolute cell references in Excel.

**Relative** cell references change when they are copied while **Absolute**
references maintain fixed row and/or column references. In Excel absolute
references are prefixed by the dollar symbol as shown below::

    A1    # Column and row are relative.
    $A1   # Column is absolute and row is relative.
    A$1   # Column is relative and row is absolute.
    $A$1  # Column and row are absolute.

See the Microsoft Office documentation for
`more information on relative and absolute references <http://office.microsoft.com/en-001/excel-help/switch-between-relative-absolute-and-mixed-references-HP010342940.aspx>`_.

Some functions such as :func:`conditional_format()` require absolute
references.


Defined Names and Named Ranges
------------------------------

It is also possible to define and use "Defined names/Named ranges" in
workbooks and worksheets, see :func:`define_name`::

    workbook.define_name('Exchange_rate', '=0.96')
    worksheet.write('B3', '=B2*Exchange_rate')

See also :ref:`ex_defined_name`.


.. _cell_utility:

Cell Utility Functions
----------------------

The ``XlsxWriter`` ``utility`` module contains several helper functions for
dealing with A1 notation as shown below. These functions can be imported as
follows::

    from xlsxwriter.utility import xl_rowcol_to_cell

    cell = xl_rowcol_to_cell(1, 2)  # C2


xl_rowcol_to_cell()
~~~~~~~~~~~~~~~~~~~

.. py:function:: xl_rowcol_to_cell(row, col[, row_abs, col_abs])

   Convert a zero indexed row and column cell reference to a A1 style string.

   :param int row:      The cell row.
   :param int col:      The cell column.
   :param bool row_abs: Optional flag to make the row absolute.
   :param bool col_abs: Optional flag to make the column absolute.
   :rtype:              A1 style string.


The ``xl_rowcol_to_cell()`` function converts a zero indexed row and column
cell values to an ``A1`` style string::

    cell = xl_rowcol_to_cell(0, 0)   # A1
    cell = xl_rowcol_to_cell(0, 1)   # B1
    cell = xl_rowcol_to_cell(1, 0)   # A2

The optional parameters ``row_abs`` and ``col_abs`` can be used to indicate
that the row or column is absolute::

    str = xl_rowcol_to_cell(0, 0, col_abs=True)                # $A1
    str = xl_rowcol_to_cell(0, 0, row_abs=True)                # A$1
    str = xl_rowcol_to_cell(0, 0, row_abs=True, col_abs=True)  # $A$1


xl_cell_to_rowcol()
~~~~~~~~~~~~~~~~~~~

.. py:function:: xl_cell_to_rowcol(cell_str)

   Convert a cell reference in A1 notation to a zero indexed row and column.

   :param string cell_str: A1 style string, absolute or relative.
   :rtype:                 Tuple of ints for (row, col).


The ``xl_cell_to_rowcol()`` function converts an Excel cell reference in ``A1``
notation to a zero based row and column. The function will also handle Excel's
absolute, ``$``, cell notation::

    (row, col) = xl_cell_to_rowcol('A1')    # (0, 0)
    (row, col) = xl_cell_to_rowcol('B1')    # (0, 1)
    (row, col) = xl_cell_to_rowcol('C2')    # (1, 2)
    (row, col) = xl_cell_to_rowcol('$C2')   # (1, 2)
    (row, col) = xl_cell_to_rowcol('C$2')   # (1, 2)
    (row, col) = xl_cell_to_rowcol('$C$2')  # (1, 2)


xl_col_to_name()
~~~~~~~~~~~~~~~~

.. py:function:: xl_col_to_name(col[, col_abs])

   Convert a zero indexed column cell reference to a string.

   :param int col:      The cell column.
   :param bool col_abs: Optional flag to make the column absolute.
   :rtype:              Column style string.


The ``xl_col_to_name()`` converts a zero based column reference to a string::

    column = xl_col_to_name(0)    # A
    column = xl_col_to_name(1)    # B
    column = xl_col_to_name(702)  # AAA

The optional parameter ``col_abs`` can be used to indicate if the column is
absolute::

    column = xl_col_to_name(0, False)  # A
    column = xl_col_to_name(0, True)   # $A
    column = xl_col_to_name(1, True)   # $B


xl_range()
~~~~~~~~~~

.. py:function:: xl_range(first_row, first_col, last_row, last_col)

   Converts zero indexed row and column cell references to a A1:B1 range
   string.

   :param int first_row:     The first cell row.
   :param int first_col:     The first cell column.
   :param int last_row:      The last cell row.
   :param int last_col:      The last cell column.
   :rtype:                   A1:B1 style range string.


The ``xl_range()`` function converts zero based row and column cell references
to an ``A1:B1`` style range string::

    cell_range = xl_range(0, 0, 9, 0)  # A1:A10
    cell_range = xl_range(1, 2, 8, 2)  # C2:C9
    cell_range = xl_range(0, 0, 3, 4)  # A1:E4


xl_range_abs()
~~~~~~~~~~~~~~

.. py:function:: xl_range_abs(first_row, first_col, last_row, last_col)

   Converts zero indexed row and column cell references to a $A$1:$B$1
   absolute range string.

   :param int first_row:     The first cell row.
   :param int first_col:     The first cell column.
   :param int last_row:      The last cell row.
   :param int last_col:      The last cell column.
   :rtype:                   $A$1:$B$1 style range string.


The ``xl_range_abs()`` function converts zero based row and column cell
references to an absolute ``$A$1:$B$1`` style range string::

    cell_range = xl_range_abs(0, 0, 9, 0)  # $A$1:$A$10
    cell_range = xl_range_abs(1, 2, 8, 2)  # $C$2:$C$9
    cell_range = xl_range_abs(0, 0, 3, 4)  # $A$1:$E$4
