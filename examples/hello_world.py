##############################################################################
#
# A hello world spreadsheet using the XlsxWriter Python module.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook("hello_world.xlsx")
worksheet = workbook.add_worksheet()

worksheet.write("A1", "Hello world")

workbook.close()
