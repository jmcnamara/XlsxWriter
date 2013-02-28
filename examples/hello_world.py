##############################################################################
#
# A hello world spreadsheet using the XlsxWriter Python module.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#
from xlsxwriter.workbook import Workbook

workbook = Workbook('hello_world.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Hello world')

workbook.close()
