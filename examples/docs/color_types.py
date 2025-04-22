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
worksheet.set_column(0, 0, 35, font_format)

#
# Color examples
#

# A Color instance using the Html string constructor.
color_format = workbook.add_format({"bg_color": Color("#FF7F50")})
worksheet.write_string(0, 0, 'Color("#FF7F50")')
worksheet.write_blank(0, 1, None, color_format)

# A Color instance using the Html integer constructor.
color_format = workbook.add_format({"bg_color": Color(0xDCDCDC)})
worksheet.write_string(1, 0, "Color(0xDCDCDC)")
worksheet.write_blank(1, 1, None, color_format)

# A Color instance using the a named color string constructor.
color_format = workbook.add_format({"bg_color": Color("Green")})
worksheet.write_string(2, 0, 'Color("Green")')
worksheet.write_blank(2, 1, None, color_format)

# A Color instance using the Theme tuple constructor.
color_format = workbook.add_format({"bg_color": Color((7, 3))})
worksheet.write_string(3, 0, "Color((7, 3))")
worksheet.write_blank(3, 1, None, color_format)

# A Color instance using the rgb() method.
color_format = workbook.add_format({"bg_color": Color.rgb("#6495ED")})
worksheet.write_string(5, 0, 'Color.rgb("#6495ED")')
worksheet.write_blank(5, 1, None, color_format)

# A Color instance using the rgb_integer() method.
color_format = workbook.add_format({"bg_color": Color.rgb_integer(0xDAA520)})
worksheet.write_string(6, 0, "Color.rgb_integer(0xDAA520)")
worksheet.write_blank(6, 1, None, color_format)

# A Color instance using the theme() method.
color_format = workbook.add_format({"bg_color": Color.theme(4, 2)})
worksheet.write_string(7, 0, "Color.theme(4, 2)")
worksheet.write_blank(7, 1, None, color_format)

# A implicit Color instance using a Html string.
color_format = workbook.add_format({"bg_color": "#E59E83"})
worksheet.write_string(9, 0, '"#E59E83"')
worksheet.write_blank(9, 1, None, color_format)

# A implicit Color instance using a named color string.
color_format = workbook.add_format({"bg_color": "Orange"})
worksheet.write_string(10, 0, '"Orange"')
worksheet.write_blank(10, 1, None, color_format)


workbook.close()
