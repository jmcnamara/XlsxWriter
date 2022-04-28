##############################################################################
#
# A simple example of converting some Unicode text to an Excel file using
# the XlsxWriter Python module.
#
# This example generates a spreadsheet with some Polish text from a file
# with UTF8 encoded text.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2022, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Open the input file with the correct encoding.
textfile = open('unicode_polish_utf8.txt', mode='r', encoding='utf-8')

# Create an new Excel file and convert the text data.
workbook = xlsxwriter.Workbook('unicode_polish_utf8.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 50)

# Start from the first cell.
row = 0
col = 0

# Read the text file and write it to the worksheet.
for line in textfile:
    # Ignore the comments in the text file.
    if line.startswith('#'):
        continue

    # Write any other lines to the worksheet.
    worksheet.write(row, col, line.rstrip("\n"))
    row += 1

workbook.close()
