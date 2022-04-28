##############################################################################
#
# A simple formatting example using XlsxWriter.
#
# This program demonstrates the indentation cell format.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2022, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('cell_indentation.xlsx')

worksheet = workbook.add_worksheet()

indent1 = workbook.add_format({'indent': 1})
indent2 = workbook.add_format({'indent': 2})

worksheet.set_column('A:A', 40)

worksheet.write('A1', "This text is indented 1 level", indent1)
worksheet.write('A2', "This text is indented 2 levels", indent2)

workbook.close()
