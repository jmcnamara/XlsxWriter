##############################################################################
#
# Example used in the XlsxWriter documentation.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter
from xlsxwriter.color import Color

# Create a new Excel file object.
workbook = xlsxwriter.Workbook("example.xlsx")

# Add a worksheet.
worksheet = workbook.add_worksheet()

# Widen the text column and set the font to a fixed width for clarity
font_format = workbook.add_format({"font_name": "Courier New"})
worksheet.set_column(0, 0, 20, font_format)

# A Color instance using the HTML string constructor.
color_format = workbook.add_format({"bg_color": Color("#FF7F50")})
worksheet.write_string(0, 0, 'Color("#FF7F50")')
worksheet.write_blank(0, 1, None, color_format)

# A Color instance using a named color string constructor.
color_format = workbook.add_format({"bg_color": Color("Green")})
worksheet.write_string(2, 0, 'Color("Green")')
worksheet.write_blank(2, 1, None, color_format)

# A Color instance using the Theme tuple constructor.
color_format = workbook.add_format({"bg_color": Color((7, 3))})
worksheet.write_string(4, 0, "Color((7, 3))")
worksheet.write_blank(4, 1, None, color_format)


workbook.close()
