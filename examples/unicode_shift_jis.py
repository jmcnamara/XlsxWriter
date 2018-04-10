##############################################################################
#
# A simple example of converting some Unicode text to an Excel file using
# the XlsxWriter Python module.
#
# This example generates a spreadsheet with some Japanese text from a file
# with Shift-JIS encoded text.
#
# Copyright 2013-2018, John McNamara, jmcnamara@cpan.org
#
import codecs
import xlsxwriter

# Open the input file with the correct encoding.
textfile = codecs.open('unicode_shift_jis.txt', 'r', 'shift_jis')

# Create an new Excel file and convert the text data.
workbook = xlsxwriter.Workbook('unicode_shift_jis.xlsx')
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
