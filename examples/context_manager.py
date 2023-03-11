##############################################################################
#
# A simple example using the XlsxWriter Python module and the "with" context
# manager. This doesn't require an explicit close().
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

with xlsxwriter.Workbook("hello_world.xlsx") as workbook:
    worksheet = workbook.add_worksheet()

    worksheet.write("A1", "Hello world")
