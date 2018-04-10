###############################################################################
# _*_ coding: utf-8
#
# A simple Unicode spreadsheet in Python 2 using the XlsxWriter Python module.
#
# Copyright 2013-2018, John McNamara, jmcnamara@cpan.org
#

# To write Unicode text in UTF-8 to a xlsxwriter file in Python 2:
#
# 1. Encode the file as UTF-8.
# 2. Include the "coding" directive at the start of the file.
# 3. Use u'' to indicate a Unicode string.

import xlsxwriter

workbook = xlsxwriter.Workbook('unicode_python2.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('B3', u'Это фраза на русском!')

workbook.close()
