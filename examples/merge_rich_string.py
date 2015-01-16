##############################################################################
#
# An  example of merging cells which contain a rich string using the
# XlsxWriter Python module.
#
# Copyright 2013-2015, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('merge_rich_string.xlsx')
worksheet = workbook.add_worksheet()

# Set up some formats to use.
red = workbook.add_format({'color': 'red'})
blue = workbook.add_format({'color': 'blue'})
cell_format = workbook.add_format({'align': 'center',
                                   'valign': 'vcenter',
                                   'border': 1})

# We can only write simple types to merged ranges so we write a blank string.
worksheet.merge_range('B2:E5', "", cell_format)

# We then overwrite the first merged cell with a rich string. Note that we
# must also pass the cell format used in the merged cells format at the end.
worksheet.write_rich_string('B2',
                            'This is ',
                            red, 'red',
                            ' and this is ',
                            blue, 'blue',
                            cell_format)

workbook.close()
