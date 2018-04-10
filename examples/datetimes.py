##############################################################################
#
# A simple program to write some dates and times to an Excel file
# using the XlsxWriter Python module.
#
# Copyright 2013-2018, John McNamara, jmcnamara@cpan.org
#
from datetime import datetime
import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('datetimes.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})

# Expand the first columns so that the dates are visible.
worksheet.set_column('A:B', 30)

# Write the column headers.
worksheet.write('A1', 'Formatted date', bold)
worksheet.write('B1', 'Format', bold)

# Create a datetime object to use in the examples.

date_time = datetime.strptime('2013-01-23 12:30:05.123',
                              '%Y-%m-%d %H:%M:%S.%f')

# Examples date and time formats. In the output file compare how changing
# the format codes change the appearance of the date.
date_formats = (
    'dd/mm/yy',
    'mm/dd/yy',
    'dd m yy',
    'd mm yy',
    'd mmm yy',
    'd mmmm yy',
    'd mmmm yyy',
    'd mmmm yyyy',
    'dd/mm/yy hh:mm',
    'dd/mm/yy hh:mm:ss',
    'dd/mm/yy hh:mm:ss.000',
    'hh:mm',
    'hh:mm:ss',
    'hh:mm:ss.000',
)

# Start from first row after headers.
row = 1

# Write the same date and time using each of the above formats.
for date_format_str in date_formats:

    # Create a format for the date or time.
    date_format = workbook.add_format({'num_format': date_format_str,
                                      'align': 'left'})

    # Write the same date using different formats.
    worksheet.write_datetime(row, 0, date_time, date_format)

    # Also write the format string for comparison.
    worksheet.write_string(row, 1, date_format_str)

    row += 1

workbook.close()
