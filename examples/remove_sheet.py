#######################################################################
#
# Example of how to remove a worksheet with XlsxWriter.
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('remove_sheet.xlsx')
worksheet1 = workbook.add_worksheet('Sheet1')
worksheet2 = workbook.add_worksheet('Sheet2')
worksheet3 = workbook.add_worksheet('Sheet3')

worksheet1.set_column('A:A', 30)
worksheet2.set_column('A:A', 30)
worksheet3.set_column('A:A', 30)

# Remove Sheet2. It won't be present after this call.
workbook.remove_sheet('Sheet2')

workbook.close()
