Tutorial 3: Writing different types of data to the XLSX File
============================================================

.. highlight:: python

In the previous section we created a simple spreadsheet with formatting using
Python and the XlsxWriter module.

This time let's extend the data we want to write to include some dates::

    expenses = (
        ['Rent', '2013-01-13', 1000],
        ['Gas',  '2013-01-14',  100],
        ['Food', '2013-01-16',  300],
        ['Gym',  '2013-01-20',   50],
    )

The corresponding spreadsheet will look like this:

.. image:: _static/tutorial03.png

The differences here are that we have added the **Date** column, formatted the
dates and made column 'B' a little wider to accommodate the dates.

To do this we can extend our program like this (the significant changes are
shown with a red line):

.. code-block:: python
   :emphasize-lines: 14, 17, 21, 26-29, 38-42

    from datetime import datetime
    from xlsxwriter.workbook import Workbook

    # Create a workbook and add a worksheet.
    workbook = Workbook('Expenses03.xlsx')
    worksheet = workbook.add_worksheet()

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': 1})

    # Add a number format for cells with money.
    money_format = workbook.add_format({'num_format': '$#,##0'})

    # Add an Excel date format.
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

    # Adjust the column width.
    worksheet.set_column(1, 1, 15)

    # Write some data headers.
    worksheet.write('A1', 'Item', bold)
    worksheet.write('B1', 'Date', bold)
    worksheet.write('C1', 'Cost', bold)

    # Some data we want to write to the worksheet.
    expenses = (
        ['Rent', '2013-01-13', 1000],
        ['Gas',  '2013-01-14',  100],
        ['Food', '2013-01-16',  300],
        ['Gym',  '2013-01-20',   50],
    )

    # Start from the first cell below the headers.
    row = 1
    col = 0

    for item, date_str, cost in (expenses):
        # Convert the date string into a datetime object.
        date = datetime.strptime(date_str, "%Y-%m-%d")

        worksheet.write_string(  row, col,     item              )
        worksheet.write_datetime(row, col + 1, date, date_format )
        worksheet.write_number(  row, col + 2, cost, money_format)
        row += 1

    # Write a total using a formula.
    worksheet.write(row, 0, 'Total', bold)
    worksheet.write(row, 2, '=SUM(C2:C5)', money_format)

    workbook.close()

The main difference between this and the previous program is that we have added
a new :ref:`Format <Format>` object for dates and we have additional handling
for data types.

Excel treats different types of input data differently, although it generally
does it transparently to the user. To illustrate this, open up a new Excel
spreadsheet, make the first column wider and enter the following data::

    123
    123.456
    1234567890123456
    Hello
    World
    2013/01/01
    2013/01/01          (But change the format from Date to General)
    01234

You should see something like the following:

.. image:: _static/tutorial03_2.png

There are a few things to notice here. The first is that the numbers in the
first three rows are stored as numbers and are aligned to the right of the
cell. The second is that the strings in the following rows are stored as
strings and are aligned to the left. The third is that the date string format
has changed and that it is aligned to the right. The final thing to notice is
that Excel has stripped the leading 0 from 012345.

Let's look at each of these in more detail.

**Numbers are stored as numbers**: In general Excel stores data as either
strings or numbers. So it shouldn't be surprising that it stores numbers as
numbers. Within a cell a number is right aligned by default. Internally Excel
handles numbers as IEEE-754 64-bit double-precision floating point. This means
that, in most cases, the maximum number of digits that can be stored in Excel
without losing precision is 15. This can be seen in cell ``'A3'`` where the 16
digit number has lost precision in the last digit.


**Strings are stored as strings**: Again not so surprising. Within a cell a
string is left aligned by default. Excel 2007+ stores strings internally as
UTF-8.

**Dates are stored as numbers**: The first clue to this is that the dates are
right aligned like numbers. More explicitly, the data in cell ``'A7'`` shows
that if you remove the date format the underlying data is a number. When you
enter a string that looks like a date Excel converts it to a number and
applies the default date format to it so that it is displayed as a date. This
is explained in more detail in :ref:`working_with_dates_and_time`.

**Things that look like numbers are stored as numbers**: In cell ``'A8'`` we
entered ``012345`` but Excel converted it to the number ``12345``. This is
something to be aware of if you are writing ID numbers or zip codes. In order
to preserve the leading zero(es) you need to store the data as either a string
or a number with a format.

XlsxWriter tries to mimic the way Excel works via the
:ref:`worksheet. <Worksheet>`:func:`write()` method and separates Python data
into types that Excel recognises. The ``write()`` method acts as a general
alias for several more specific methods:

* :func:`write_string()`
* :func:`write_number()`
* :func:`write_datetime()`
* :func:`write_blank()`
* :func:`write_formula()`

So, let's see how all this affects our program.

The main change in our example program is the addition of date handling. As we
saw above Excel stores dates as numbers. XlsxWriter makes the required
conversion if the date and time are in Python ``datetime`` format. To convert
the date strings in our example to ``datetime`` objects we use the
``datetime.strptime`` function. We then use the ``write_datetime()`` function
to write it to a file. However, since the date is converted to a number we
also need to add a number format to ensure that Excel displays it as as date::

    from datetime import datetime
    ...

    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
    ...

    for item, date_str, cost in (expenses):
        # Convert the date string into a datetime object.
        date = datetime.strptime(date_str, "%Y-%m-%d")
        ...
        worksheet.write_datetime(row, col + 1, date, date_format )
        ...

The other thing to notice in our program is that we have used explicit write
methods for different types of data::

        worksheet.write_string(  row, col,     item              )
        worksheet.write_datetime(row, col + 1, date, date_format )
        worksheet.write_number(  row, col + 2, cost, money_format)

This is mainly to show that if you need more control over the type of data you
write to a worksheet you can use the appropriate method. In this simplified
example the :func:`write()` method would have worked as well but it is
important to note that in cases where ``write()`` doesn't do the right thing
you will need to be explicit.

Finally, the last addition to our program is the :func:`set_column` method to
adjust the width of column 'B' so that the dates are more clearly visible::

    # Adjust the column width.
    worksheet.set_column('B:B', 15)

The :func:`set_column` and corresponding :func:`set_row` methods are explained
in more detail in :ref:`worksheet`.
