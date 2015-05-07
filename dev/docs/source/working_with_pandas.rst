.. _ewx_pandas:

Working with Python Pandas
==========================

Python `Pandas <http://pandas.pydata.org/>`_ is a Python data analysis
library. It can read, filter and re-arrange small and large data sets and
output them in a range of formats including Excel.

Pandas writes Excel files using the `Xlwt
<https://pypi.python.org/pypi/xlwt>`_ module for xls files and the `Openpyxl
<https://pypi.python.org/pypi/openpyxl>`_ or XlsxWriter modules for xlsx
files.


Using XlsxWriter with Pandas
----------------------------

To use XlsxWriter with Pandas you specify it as the Excel writer *engine*::

    import pandas as pd

    # Create a Pandas dataframe from the data.
    df = pd.DataFrame({'Data': [10, 20, 30, 20, 15, 30, 45]})

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Sheet1')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

The output from this would look like the following:

.. image:: _images/pandas_simple.png

See the full example at :ref:`ex_pandas_simple`


Accessing XlsxWriter from Pandas
--------------------------------

In order to apply XlsxWriter features such as Charts, Conditional Formatting
and Column Formatting to the Pandas output we need to access the underlying
:ref:`workbook <Workbook>` and :ref:`worksheet <Worksheet>` objects. After
that we can treat them as normal XlsxWriter objects.

Continuing on from the above example we do that as follows::

    import pandas as pd

    # Create a Pandas dataframe from the data.
    df = pd.DataFrame({'Data': [10, 20, 30, 20, 15, 30, 45]})

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Sheet1')

    # Get the xlsxwriter objects from the dataframe writer object.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

This is equivalent to the following code when using XlsxWriter on its own::

    workbook  = xlsxwriter.Workbook('filename.xlsx')
    worksheet = workbook.add_worksheet()

Once we have the Workbook and Worksheet objects we we can use them to apply
other features such as adding a chart::

    # Create a chart object.
    chart = workbook.add_chart({'type': 'column'})

    # Configure the series of the chart from the dataframe data.
    chart.add_series({'values': '=Sheet1!$B$2:$B$8'})

    # Insert the chart into the worksheet.
    worksheet.insert_chart('D2', chart)

The output would look like this:

.. image:: _images/pandas_chart.png

See the full example at :ref:`ex_pandas_chart`

Alternatively, we could apply a conditional format like this::

    # Apply a conditional format to the cell range.
    worksheet.conditional_format('B2:B8', {'type': '3_color_scale'})

Which would give:

.. image:: _images/pandas_conditional.png

See the full example at :ref:`ex_pandas_conditional`

It isn't possible to format any cells that already have a format applied to
them such as the header and index cells and any cells that contain dates of
datetimes. However, it is possible to set default date and datetime formats::

    writer = pd.ExcelWriter("pandas_datetime.xlsx",
                            engine='xlsxwriter',
                            datetime_format='mmm d yyyy hh:mm:ss',
                            date_format='mmmm dd yyyy')

Which would give:

.. image:: _images/pandas_datetime.png

See the full example at :ref:`ex_pandas_datetime`

It is possible to format any other column data using :func:`set_column()`::

    # Add some cell formats.
    format1 = workbook.add_format({'num_format': '#,##0.00'})
    format2 = workbook.add_format({'num_format': '0%'})

    # Set the column width and format.
    worksheet.set_column('B:B', 18, format1)

    # Set the format but not the column width.
    worksheet.set_column('C:C', None, format2)

.. image:: _images/pandas_column_formats.png

See the full example at :ref:`ex_pandas_column_formats`


Further Pandas and Excel Information
------------------------------------

See all the XlsxWriter Pandas examples at :ref:`pandas_examples`.

See also `Using Pandas and XlsxWriter to create Excel charts
<http://pandas-xlsxwriter-charts.readthedocs.org/en/latest/index.html>`_ and
the series of articles on the Practical Business Python website about `Using
Pandas and Excel <http://pbpython.com/tag/excel.html>`_.
