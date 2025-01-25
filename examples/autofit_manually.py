##############################################################################
#
# The following example demonstrates manually auto-fitting the the width of a
# column in Excel based on the maximum string width. The worksheet ``autofit()``
# method will do this automatically but occasionally you may need to control the
# maximum and minimum column widths yourself.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
from functools import reduce

import xlsxwriter
from xlsxwriter.utility import cell_autofit_width

workbook = xlsxwriter.Workbook("autofit.xlsx")
worksheet = workbook.add_worksheet()

# Some string data to write.
cities = ["Addis Ababa", "Buenos Aires", "Cairo", "Dhaka"]

# Write the strings:
worksheet.write_column(0, 0, cities)

# Find the maximum column width in pixels.
max_width = reduce(max, map(cell_autofit_width, cities))

# Set the column width as if it was auto-fitted.
worksheet.set_column_pixels(0, 0, max_width)

workbook.close()
