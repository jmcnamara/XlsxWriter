##############################################################################
#
# An example of adding support for user defined types to the XlsxWriter write()
# method.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter
import uuid


# Create a function that will behave like a worksheet write() method.
#
# This function takes a UUID and writes it as as string. It should take the
# parameters shown below and return the return value from the called worksheet
# write_*() method. In this case it changes the UUID to a string and calls
# write_string() to write it.
#
def write_uuid(worksheet, row, col, token, format=None):
    return worksheet.write_string(row, col, str(token), format)


# Set up the workbook as usual.
workbook = xlsxwriter.Workbook("user_types1.xlsx")
worksheet = workbook.add_worksheet()

# Make the first column wider for clarity.
worksheet.set_column("A:A", 40)

# Add the write() handler/callback to the worksheet.
worksheet.add_write_handler(uuid.UUID, write_uuid)

# Create a UUID.
my_uuid = uuid.uuid3(uuid.NAMESPACE_DNS, "python.org")

# Write the UUID. This would raise a TypeError without the handler.
worksheet.write("A1", my_uuid)

workbook.close()
