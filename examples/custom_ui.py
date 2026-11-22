#######################################################################
#
# An example of adding a custom ribbon tab to an XlsxWriter file using a
# Custom UI XML file.
#
# The custom tab contains built-in Excel commands and a button that runs a
# macro from a VBA project file.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2026, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Note the file extension should be .xlsm.
workbook = xlsxwriter.Workbook("custom_ui.xlsm")
worksheet = workbook.add_worksheet()

worksheet.set_column("A:A", 50)

# Add the Custom UI XML file for the Excel 2010+ ribbon.
workbook.add_custom_ui("./customUI14.xml")

# Show some instructions for the end user.
bold = workbook.add_format({"bold": True})
worksheet.write("A1", "Use the 'My Ribbon' tab to sort the selected data.", bold)

# Add some data to sort.
for row, fruit in enumerate(["Pear", "Apple", "Orange", "Banana", "Grape"], 2):
    worksheet.write(row, 0, fruit)

# Preselect the data so it is ready to sort.
worksheet.set_selection("A3:A7")

# Add the VBA project for the macro on the custom ribbon tab. This was extracted
# from an Excel file using the vba_extract tool.
workbook.add_vba_project("./vbaProject_custom_ui.bin")

workbook.close()
