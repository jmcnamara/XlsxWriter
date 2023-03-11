#######################################################################
#
# Example of how to use Python and the XlsxWriter module to change the default
# worksheet and cell text direction from left-to-right to right-to-left as
# required by some middle eastern versions of Excel.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook("right_to_left.xlsx")

# Add two worksheets.
worksheet1 = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()

# Add the cell formats.
format_left_to_right = workbook.add_format({"reading_order": 1})
format_right_to_left = workbook.add_format({"reading_order": 2})

# Make the columns wider for clarity.
worksheet1.set_column("A:A", 25)
worksheet2.set_column("A:A", 25)

# Change the direction for worksheet2.
worksheet2.right_to_left()

# Write some data to show the difference.

# Standard direction:         | A1 | B1 | C1 | ...
worksheet1.write("A1", "نص عربي / English text")  # Default direction.
worksheet1.write("A2", "نص عربي / English text", format_left_to_right)
worksheet1.write("A3", "نص عربي / English text", format_right_to_left)

# Right to left direction:    ... | C1 | B1 | A1 |
worksheet2.write("A1", "نص عربي / English text")  # Default direction.
worksheet2.write("A2", "نص عربي / English text", format_left_to_right)
worksheet2.write("A3", "نص عربي / English text", format_right_to_left)

workbook.close()
