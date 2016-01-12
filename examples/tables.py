###############################################################################
#
# Example of how to add tables to an XlsxWriter worksheet.
#
# Tables in Excel are used to group rows and columns of data into a single
# structure that can be referenced in a formula or formatted collectively.
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('tables.xlsx')
worksheet1 = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()
worksheet3 = workbook.add_worksheet()
worksheet4 = workbook.add_worksheet()
worksheet5 = workbook.add_worksheet()
worksheet6 = workbook.add_worksheet()
worksheet7 = workbook.add_worksheet()
worksheet8 = workbook.add_worksheet()
worksheet9 = workbook.add_worksheet()
worksheet10 = workbook.add_worksheet()
worksheet11 = workbook.add_worksheet()
worksheet12 = workbook.add_worksheet()

currency_format = workbook.add_format({'num_format': '$#,##0'})

# Some sample data for the table.
data = [
    ['Apples', 10000, 5000, 8000, 6000],
    ['Pears', 2000, 3000, 4000, 5000],
    ['Bananas', 6000, 6000, 6500, 6000],
    ['Oranges', 500, 300, 200, 700],

]


###############################################################################
#
# Example 1.
#
caption = 'Default table with no data.'

# Set the columns widths.
worksheet1.set_column('B:G', 12)

# Write the caption.
worksheet1.write('B1', caption)

# Add a table to the worksheet.
worksheet1.add_table('B3:F7')


###############################################################################
#
# Example 2.
#
caption = 'Default table with data.'

# Set the columns widths.
worksheet2.set_column('B:G', 12)

# Write the caption.
worksheet2.write('B1', caption)

# Add a table to the worksheet.
worksheet2.add_table('B3:F7', {'data': data})


###############################################################################
#
# Example 3.
#
caption = 'Table without default autofilter.'

# Set the columns widths.
worksheet3.set_column('B:G', 12)

# Write the caption.
worksheet3.write('B1', caption)

# Add a table to the worksheet.
worksheet3.add_table('B3:F7', {'autofilter': 0})

# Table data can also be written separately, as an array or individual cells.
worksheet3.write_row('B4', data[0])
worksheet3.write_row('B5', data[1])
worksheet3.write_row('B6', data[2])
worksheet3.write_row('B7', data[3])


###############################################################################
#
# Example 4.
#
caption = 'Table without default header row.'

# Set the columns widths.
worksheet4.set_column('B:G', 12)

# Write the caption.
worksheet4.write('B1', caption)

# Add a table to the worksheet.
worksheet4.add_table('B4:F7', {'header_row': 0})

# Table data can also be written separately, as an array or individual cells.
worksheet4.write_row('B4', data[0])
worksheet4.write_row('B5', data[1])
worksheet4.write_row('B6', data[2])
worksheet4.write_row('B7', data[3])


###############################################################################
#
# Example 5.
#
caption = 'Default table with "First Column" and "Last Column" options.'

# Set the columns widths.
worksheet5.set_column('B:G', 12)

# Write the caption.
worksheet5.write('B1', caption)

# Add a table to the worksheet.
worksheet5.add_table('B3:F7', {'first_column': 1, 'last_column': 1})

# Table data can also be written separately, as an array or individual cells.
worksheet5.write_row('B4', data[0])
worksheet5.write_row('B5', data[1])
worksheet5.write_row('B6', data[2])
worksheet5.write_row('B7', data[3])


###############################################################################
#
# Example 6.
#
caption = 'Table with banded columns but without default banded rows.'

# Set the columns widths.
worksheet6.set_column('B:G', 12)

# Write the caption.
worksheet6.write('B1', caption)

# Add a table to the worksheet.
worksheet6.add_table('B3:F7', {'banded_rows': 0, 'banded_columns': 1})

# Table data can also be written separately, as an array or individual cells.
worksheet6.write_row('B4', data[0])
worksheet6.write_row('B5', data[1])
worksheet6.write_row('B6', data[2])
worksheet6.write_row('B7', data[3])


###############################################################################
#
# Example 7.
#
caption = 'Table with user defined column headers'

# Set the columns widths.
worksheet7.set_column('B:G', 12)

# Write the caption.
worksheet7.write('B1', caption)

# Add a table to the worksheet.
worksheet7.add_table('B3:F7', {'data': data,
                               'columns': [{'header': 'Product'},
                                           {'header': 'Quarter 1'},
                                           {'header': 'Quarter 2'},
                                           {'header': 'Quarter 3'},
                                           {'header': 'Quarter 4'},
                                           ]})


###############################################################################
#
# Example 8.
#
caption = 'Table with user defined column headers'

# Set the columns widths.
worksheet8.set_column('B:G', 12)

# Write the caption.
worksheet8.write('B1', caption)

# Formula to use in the table.
formula = '=SUM(Table8[@[Quarter 1]:[Quarter 4]])'

# Add a table to the worksheet.
worksheet8.add_table('B3:G7', {'data': data,
                               'columns': [{'header': 'Product'},
                                           {'header': 'Quarter 1'},
                                           {'header': 'Quarter 2'},
                                           {'header': 'Quarter 3'},
                                           {'header': 'Quarter 4'},
                                           {'header': 'Year',
                                            'formula': formula},
                                           ]})


###############################################################################
#
# Example 9.
#
caption = 'Table with totals row (but no caption or totals).'

# Set the columns widths.
worksheet9.set_column('B:G', 12)

# Write the caption.
worksheet9.write('B1', caption)

# Formula to use in the table.
formula = '=SUM(Table9[@[Quarter 1]:[Quarter 4]])'

# Add a table to the worksheet.
worksheet9.add_table('B3:G8', {'data': data,
                               'total_row': 1,
                               'columns': [{'header': 'Product'},
                                           {'header': 'Quarter 1'},
                                           {'header': 'Quarter 2'},
                                           {'header': 'Quarter 3'},
                                           {'header': 'Quarter 4'},
                                           {'header': 'Year',
                                            'formula': formula
                                            },
                                           ]})


###############################################################################
#
# Example 10.
#
caption = 'Table with totals row with user captions and functions.'

# Set the columns widths.
worksheet10.set_column('B:G', 12)

# Write the caption.
worksheet10.write('B1', caption)

# Options to use in the table.
options = {'data': data,
           'total_row': 1,
           'columns': [{'header': 'Product', 'total_string': 'Totals'},
                       {'header': 'Quarter 1', 'total_function': 'sum'},
                       {'header': 'Quarter 2', 'total_function': 'sum'},
                       {'header': 'Quarter 3', 'total_function': 'sum'},
                       {'header': 'Quarter 4', 'total_function': 'sum'},
                       {'header': 'Year',
                        'formula': '=SUM(Table10[@[Quarter 1]:[Quarter 4]])',
                        'total_function': 'sum'
                        },
                       ]}

# Add a table to the worksheet.
worksheet10.add_table('B3:G8', options)


###############################################################################
#
# Example 11.
#
caption = 'Table with alternative Excel style.'

# Set the columns widths.
worksheet11.set_column('B:G', 12)

# Write the caption.
worksheet11.write('B1', caption)

# Options to use in the table.
options = {'data': data,
           'style': 'Table Style Light 11',
           'total_row': 1,
           'columns': [{'header': 'Product', 'total_string': 'Totals'},
                       {'header': 'Quarter 1', 'total_function': 'sum'},
                       {'header': 'Quarter 2', 'total_function': 'sum'},
                       {'header': 'Quarter 3', 'total_function': 'sum'},
                       {'header': 'Quarter 4', 'total_function': 'sum'},
                       {'header': 'Year',
                        'formula': '=SUM(Table11[@[Quarter 1]:[Quarter 4]])',
                        'total_function': 'sum'
                        },
                       ]}


# Add a table to the worksheet.
worksheet11.add_table('B3:G8', options)


###############################################################################
#
# Example 12.
#
caption = 'Table with column formats.'

# Set the columns widths.
worksheet12.set_column('B:G', 12)

# Write the caption.
worksheet12.write('B1', caption)

# Options to use in the table.
options = {'data': data,
           'total_row': 1,
           'columns': [{'header': 'Product', 'total_string': 'Totals'},
                       {'header': 'Quarter 1',
                        'total_function': 'sum',
                        'format': currency_format,
                        },
                       {'header': 'Quarter 2',
                        'total_function': 'sum',
                        'format': currency_format,
                        },
                       {'header': 'Quarter 3',
                        'total_function': 'sum',
                        'format': currency_format,
                        },
                       {'header': 'Quarter 4',
                        'total_function': 'sum',
                        'format': currency_format,
                        },
                       {'header': 'Year',
                        'formula': '=SUM(Table12[@[Quarter 1]:[Quarter 4]])',
                        'total_function': 'sum',
                        'format': currency_format,
                        },
                       ]}

# Add a table to the worksheet.
worksheet12.add_table('B3:G8', options)

workbook.close()
