##############################################################################
#
# An example of adding checkbox boolean values to a worksheet using the the
# XlsxWriter Python module.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Create a new Excel file object.
workbook = xlsxwriter.Workbook("checkbox.xlsx")

# Add a worksheet to the workbook.
worksheet = workbook.add_worksheet()

# Create some formats to use in the worksheet.
bold = workbook.add_format({"bold": True})
light_red = workbook.add_format({"bg_color": "#FFC7CE"})
light_green = workbook.add_format({"bg_color": "#C6EFCE"})

# Set the column width for clarity.
worksheet.set_column(0, 0, 30)

# Write some descriptions.
worksheet.write(1, 0, "Some simple checkboxes:", bold)
worksheet.write(4, 0, "Some checkboxes with cell formats:", bold)

# Insert some boolean checkboxes to the worksheet.
worksheet.insert_checkbox(1, 1, False)
worksheet.insert_checkbox(2, 1, True)

# Insert some checkboxes with cell formats.
worksheet.insert_checkbox(4, 1, False, light_red)
worksheet.insert_checkbox(5, 1, True, light_green)

# Close the workbook.
workbook.close()
