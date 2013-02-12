.. _working_with_dates_and_time:

Working with Dates and Time
===========================

Dates and times in Excel are represented by real numbers, for example "Jan 1
2013 12:00 PM" is represented by the number 41275.5.

The integer part of the number stores the number of days since the epoch and
the fractional part stores the percentage of the day.

A date or time in Excel is just like any other number. To display the number as
a date you must apply an Excel number format to it. Here are some examples:

.. code-block:: python

    from xlsxwriter.workbook import Workbook
    
    workbook = Workbook('date_examples.xlsx')
    worksheet = workbook.add_worksheet()

    # Widen column A for extra visibility.
    worksheet.set_column('A:A', 30)

    # A number to convert to a date.
    number = 41333.5

    # Write it as a number without formatting.
    worksheet.write('A1', number)                # 41333.5

    format2 = workbook.add_format({'num_format': 'dd/mm/yy'})
    worksheet.write('A2', number, format2)       # 28/02/13

    format3 = workbook.add_format({'num_format': 'mm/dd/yy'})
    worksheet.write('A3', number, format3)       # 02/28/13

    format4 = workbook.add_format({'num_format': 'd-m-yyyy'})
    worksheet.write('A4', number, format4)       # 28-2-2013

    format5 = workbook.add_format({'num_format': 'dd/mm/yy hh:mm'})
    worksheet.write('A5', number, format5)       # 28/02/13 12:00

    format6 = workbook.add_format({'num_format': 'd mmm yyyy'})
    worksheet.write('A6', number, format6)       # 28 Feb 2013

    format7 = workbook.add_format({'num_format': 'mmm d yyyy hh:mm AM/PM'})
    worksheet.write('A7', number, format7)       # Feb 28 2008 12:00 PM

    workbook.close()

.. image:: _static/working_with_dates_and_times01.png

To make working with dates and times a little easier the XlsxWriter module
provides a :func:`write_datetime` method to write dates in
:class:`datetime.datetime` format.

The :class:`datetime.datetime` class is part of the standard Python
:mod:`datetime` library.

There are many way to create a a datetime object but the most common is to use
the :meth:`datetime.strptime <datetime.datetime.strptime>` method.

TODO.





