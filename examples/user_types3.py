##############################################################################
#
# An example of adding support for user defined types to the XlsxWriter write()
# method.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter


# Create a function that changes the worksheet write() method so that it
# hides/replaces user passwords when writing string data. The password data,
# based on the sample data structure, will be data in the second column, apart
# from the header row.
def hide_password(worksheet, row, col, string, format=None):
    if col == 1 and row > 0:
        return worksheet.write_string(row, col, "****", format)
    else:
        return worksheet.write_string(row, col, string, format)


# Set up the workbook as usual.
workbook = xlsxwriter.Workbook("user_types3.xlsx")
worksheet = workbook.add_worksheet()

# Make the headings in the first row bold.
bold = workbook.add_format({"bold": True})
worksheet.set_row(0, None, bold)

# Add the write() handler/callback to the worksheet.
worksheet.add_write_handler(str, hide_password)

# Create some data to write.
my_data = [
    ["Name", "Password", "City"],
    ["Sara", "$5%^6&", "Rome"],
    ["Michele", "123abc", "Milano"],
    ["Maria", "juvexme", "Torino"],
    ["Paolo", "qwerty", "Fano"],
]

# Write the data. Note that write_row() calls write() so this will work as
# expected.
for row_num, row_data in enumerate(my_data):
    worksheet.write_row(row_num, 0, row_data)

workbook.close()
