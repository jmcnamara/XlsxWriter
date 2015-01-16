#######################################################################
#
# Example of how to set Excel worksheet tab colours using Python
# and the XlsxWriter module..
#
# Copyright 2013-2015, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('tab_colors.xlsx')

# Set up some worksheets.
worksheet1 = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()
worksheet3 = workbook.add_worksheet()
worksheet4 = workbook.add_worksheet()

# Set tab colours
worksheet1.set_tab_color('red')
worksheet2.set_tab_color('green')
worksheet3.set_tab_color('#FF9900')  # Orange

# worksheet4 will have the default colour.

workbook.close()
