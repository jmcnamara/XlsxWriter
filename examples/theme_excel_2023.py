##############################################################################
#
# An example of changing the default theme for a workbook to the Excel 2023
# Office/Aptos theme using XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2026, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Use the Excel 2023 Office/Aptos theme in the workbook.
workbook = xlsxwriter.Workbook("theme_excel_2023.xlsx", {"excel_2023_theme": True})

worksheet = workbook.add_worksheet()

# Write some text to demonstrate the changed theme.
worksheet.write("A1", "Hello")

workbook.close()
