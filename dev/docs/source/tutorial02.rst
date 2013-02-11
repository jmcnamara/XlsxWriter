.. _tutorial2:

Tutorial 2: Adding formatting to the XLSX File
==============================================

.. highlight:: python

In the previous section we created a simple spreadsheet using Python and the
XlsxWriter module.

This presented the data that we wanted but it looked a little bare. In order to
make it clearer we would like to add some simple formatting, like this:

.. image:: _static/tutorial02.png

The differences here are that we have added **Item** and **Cost** header
columns in a bold font, we have formatted the currency in the second column
and we have made the **Total** string bold.

To do this we can extend our program like this (the significant changes are
shown with a red line):

.. code-block:: python
   :emphasize-lines: 7-15, 32, 36-37
      
    from xlsxwriter.workbook import Workbook

    # Create a workbook and add a worksheet.
    workbook = Workbook('Expenses02.xlsx')
    worksheet = workbook.add_worksheet()
    
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    
    # Add a number format for cells with money.
    money = workbook.add_format({'num_format': '$#,##0'})
    
    # Write some data header.
    worksheet.write('A1', 'Item', bold)
    worksheet.write('B1', 'Cost', bold)
    
    # Some data we want to write to the worksheet.
    expenses = (
        ['Rent', 1000],
        ['Gas',   100],
        ['Food',  300],
        ['Gym',    50],
    )
    
    # Start from the first cell below the headers.
    row = 1
    col = 0
    
    # Iterate over the data and write it out row by row.
    for item, cost in (expenses):
        worksheet.write(row, col,     item)
        worksheet.write(row, col + 1, cost, money)
        row += 1
    
    # Write a total using a formula.
    worksheet.write(row, 0, 'Total',       bold)
    worksheet.write(row, 1, '=SUM(B2:B5)', money)
    
    workbook.close()

The main difference between this and the previous program is that we have added
two :ref:`Format <Format>` objects that we can use to format cells in the
spreadsheet::

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    
    # Add a number format for cells with money.
    money = workbook.add_format({'num_format': '$#,##0'})

We then pass this :ref:`Format <Format>` as an optional third parameter to the
:ref:`worksheet. <Worksheet>`:func:`write()` method::

    write(row, column, token, [format])   

Like this::

    worksheet.write(row, 0, 'Total', bold)

Which leads us to another new feature in this program. To add the headers in
the first row of the worksheet we used :func:`write()` like this::

    worksheet.write('A1', 'Item', bold)
    worksheet.write('B1', 'Cost', bold)

So, instead of ``(row, col)`` we used the Excel ``'A1'``  style notation. See
:ref:`cell_notation` for more details but don't be too concerned about it for
now. It is just a little syntactic sugar to help with laying out worksheets.










