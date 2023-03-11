#######################################################################
#
# An example of adding macros to an XlsxWriter file using a VBA project
# file extracted from an existing Excel xlsm file.
#
# The vba_extract.py utility supplied with XlsxWriter can be used to extract
# the vbaProject.bin file.
#
# An embedded macro is connected to a form button on the worksheet.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Note the file extension should be .xlsm.
workbook = xlsxwriter.Workbook("macros.xlsm")
worksheet = workbook.add_worksheet()

worksheet.set_column("A:A", 30)

# Add the VBA project binary.
workbook.add_vba_project("./vbaProject.bin")

# Show text for the end user.
worksheet.write("A3", "Press the button to say hello.")

# Add a button tied to a macro in the VBA project.
worksheet.insert_button(
    "B3", {"macro": "say_hello", "caption": "Press Me", "width": 80, "height": 30}
)

workbook.close()
