##############################################################################
#
# An example of changing the default theme for a workbook to a user supplied
# custom theme using XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2026, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Set the default format to match the custom theme. Note, the font scheme is set
# to "minor" to indicate that the font is part of the theme.
workbook = xlsxwriter.Workbook(
    "theme_custom.xlsx",
    {
        "default_format_properties": {
            "font_name": "Arial",
            "font_size": 11,
            "font_scheme": "minor",
        },
        "default_row_height": 19,
        "default_column_width": 72,
    },
)

# Add a custom theme to the workbook. This can be a .thmx file, an .xlsx file,
# or a theme1.xml file extracted from an unzipped xlsx file.
workbook.use_custom_theme("Technic.thmx")

worksheet = workbook.add_worksheet()

# Write some text to demonstrate the changed theme.
worksheet.write("A1", "Hello")

workbook.close()
