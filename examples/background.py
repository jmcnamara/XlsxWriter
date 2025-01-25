##############################################################################
#
# An example of setting a worksheet background image with the XlsxWriter
# Python module.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook("background.xlsx")
worksheet = workbook.add_worksheet()

# Set the background image.
worksheet.set_background("logo.png")

workbook.close()
