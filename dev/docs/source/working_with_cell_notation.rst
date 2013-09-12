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
you can use row-column notation::

    # These are equivalent.
    worksheet.write(0, 7, 200)
    worksheet.write('H1', 200)


The ``XlsxWriter`` ``utility`` contains several helper functions for dealing
with A1 notation, for example::

    from utility import xl_cell_to_rowcol, import xl_rowcol_to_cell

    (row, col) = xl_cell_to_rowcol('C2')  # -> (1, 2)
    string     = xl_rowcol_to_cell(1, 2)  # -> C2

.. note::
   In Excel it is also possible to use R1C1 notation. This is not
   supported by XlsxWriter.
