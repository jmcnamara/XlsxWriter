##############################################################################
#
# A simple example of adding checkboxes to an Excel worksheet with Python
# and XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2024, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook("checkbox.xlsx")
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column("A:A", 20)

# Add a checkbox format.
checkbox = workbook.add_format({"checkbox": True})

# Add a checkbox to A1.
worksheet.write("A1", "", checkbox)

# Add a checked checkbox to A2.
worksheet.write("A2", True, checkbox)

# Add a checkbox format, colored red.
red_checkbox = workbook.add_format({"font_color": "#FF0000", "checkbox": True})

# Add a red checkbox to A3.
worksheet.write("A3", True, red_checkbox)

workbook.close()
