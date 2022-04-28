#######################################################################
#
# An example of using Python and XlsxWriter to write some "rich strings",
# i.e., strings with multiple formats.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2022, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('rich_strings.xlsx')
worksheet = workbook.add_worksheet()

worksheet.set_column('A:A', 30)

# Set up some formats to use.
bold = workbook.add_format({'bold': True})
italic = workbook.add_format({'italic': True})
red = workbook.add_format({'color': 'red'})
blue = workbook.add_format({'color': 'blue'})
center = workbook.add_format({'align': 'center'})
superscript = workbook.add_format({'font_script': 1})

# Write some strings with multiple formats.
worksheet.write_rich_string('A1',
                            'This is ',
                            bold, 'bold',
                            ' and this is ',
                            italic, 'italic')

worksheet.write_rich_string('A3',
                            'This is ',
                            red, 'red',
                            ' and this is ',
                            blue, 'blue')

worksheet.write_rich_string('A5',
                            'Some ',
                            bold, 'bold text',
                            ' centered',
                            center)

worksheet.write_rich_string('A7',
                            italic,
                            'j = k',
                            superscript, '(n-1)',
                            center)

# If you have formats and segments in a list you can add them like this:
segments = ['This is ', bold, 'bold', ' and this is ', blue, 'blue']
worksheet.write_rich_string('A9', *segments)

workbook.close()
