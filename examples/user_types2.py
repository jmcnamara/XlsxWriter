##############################################################################
#
# An example of adding support for user defined types to the XlsxWriter write()
# method.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter
import math


# Create a function that will behave like a worksheet write() method.
#
# This function takes a float and if it is NaN then it writes a blank cell
# instead. It should take the parameters shown below and return the return
# value from the called worksheet write_*() method.
#
def ignore_nan(worksheet, row, col, number, format=None):
    if math.isnan(number):
        return worksheet.write_blank(row, col, None, format)
    else:
        # Return control to the calling write() method for any other number.
        return None


# Set up the workbook as usual.
workbook = xlsxwriter.Workbook("user_types2.xlsx")
worksheet = workbook.add_worksheet()


# Add the write() handler/callback to the worksheet.
worksheet.add_write_handler(float, ignore_nan)

# Create some data to write.
my_data = [1, 2, float("nan"), 4, 5]

# Write the data. Note that write_row() calls write() so this will work as
# expected. Writing NaN values would raise a TypeError without the handler.
worksheet.write_row("A1", my_data)

workbook.close()
