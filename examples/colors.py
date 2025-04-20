##############################################################################
#
# A demonstration of the RGB and Theme colors palettes available in the
# XlsxWriter Python module.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter
from xlsxwriter.color import Color

# Create a new Excel file object.
workbook = xlsxwriter.Workbook("colors.xlsx")

# Add a worksheet for the RGB colors.
worksheet = workbook.add_worksheet("RGB Colors")

# Write some predefined colors to cells.
named_colors = [
    "Black",
    "Blue",
    "Brown",
    "Cyan",
    "Gray",
    "Green",
    "Lime",
    "Magenta",
    "Navy",
    "Orange",
    "Pink",
    "Purple",
    "Red",
    "Silver",
    "White",
    "Yellow",
]

# Write the named colors.
for row, name in enumerate(named_colors):
    color_format = workbook.add_format({"bg_color": Color(name)})
    worksheet.write_string(row, 0, name)
    worksheet.write_blank(row, 1, None, color_format)

# Write some user-defined RGB colors to cells.
user_defined_colors = [
    "#FF7F50",
    "#DCDCDC",
    "#6495ED",
    "#DAA520",
]

for row, name in enumerate(user_defined_colors, start=len(named_colors)):
    color_format = workbook.add_format({"bg_color": Color(name)})
    worksheet.write_string(row, 0, name)
    worksheet.write_blank(row, 1, None, color_format)

# Add a worksheet for the Theme colors.
worksheet = workbook.add_worksheet("Theme Colors")

# Add alternative colors for the cell text.
black_text = Color("#000000")
white_text = Color("#FFFFFF")

# Create a cell with each of the theme colors.
for row in range(6):
    for col in range(10):
        # Use the theme color for the background.
        theme_color = Color.theme(col, row)

        # Use white font color for better contrast on dark backgrounds.
        if col != 0:
            font_color = white_text
        else:
            font_color = black_text

        color_format = workbook.add_format(
            {
                "align": "center",
                "bg_color": theme_color,
                "font_color": font_color,
            }
        )
        worksheet.write_string(row, col, f"({col}, {row})", color_format)

# Save the file to disk.
workbook.close()
