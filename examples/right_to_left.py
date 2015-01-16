#######################################################################
#
# Example of how to use Python and the XlsxWriter module to change the
# default worksheet direction from left-to-right to right-to-left as
# required by some middle eastern versions of Excel.
#
# Copyright 2013-2015, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('right_to_left.xlsx')

# Add two worksheets.
worksheet1 = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()

# Change the direction for worksheet2.
worksheet2.right_to_left()

# Write some data to show the difference.

# Standard direction:      A1, B1, C1, ...
worksheet1.write('A1', 'Hello')

# Right to left direction: ..., C1, B1, A1
worksheet2.write('A1', 'Hello')

workbook.close()
