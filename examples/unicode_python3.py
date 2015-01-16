###############################################################################
#
#
# A simple Unicode spreadsheet in Python 3 using the XlsxWriter Python module.
#
# Copyright 2013-2015, John McNamara, jmcnamara@cpan.org
#

# To write Unicode text in UTF-8 to a xlsxwriter file in Python 3:
#
# 1. Encode the file as UTF-8.
#
#

import xlsxwriter

workbook = xlsxwriter.Workbook('unicode_python3.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('B3', 'Это фраза на русском!')

workbook.close()
