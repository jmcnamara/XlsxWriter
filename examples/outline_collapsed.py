###############################################################################
#
# Example of how to use Python and XlsxWriter to generate Excel outlines and
# grouping.
#
# These examples focus mainly on collapsed outlines. See also the
# outlines.py example program for more general examples.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2022, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Create a new workbook and add some worksheets
workbook = xlsxwriter.Workbook('outline_collapsed.xlsx')
worksheet1 = workbook.add_worksheet('Outlined Rows')
worksheet2 = workbook.add_worksheet('Collapsed Rows 1')
worksheet3 = workbook.add_worksheet('Collapsed Rows 2')
worksheet4 = workbook.add_worksheet('Collapsed Rows 3')
worksheet5 = workbook.add_worksheet('Outline Columns')
worksheet6 = workbook.add_worksheet('Collapsed Columns')

# Add a general format
bold = workbook.add_format({'bold': 1})


# This function will generate the same data and sub-totals on each worksheet.
# Used in the first 4 examples.
#
def create_sub_totals(worksheet):
    # Adjust the column width for clarity.
    worksheet.set_column('A:A', 20)

    # Add the data, labels and formulas.
    worksheet.write('A1', 'Region', bold)
    worksheet.write('A2', 'North')
    worksheet.write('A3', 'North')
    worksheet.write('A4', 'North')
    worksheet.write('A5', 'North')
    worksheet.write('A6', 'North Total', bold)

    worksheet.write('B1', 'Sales', bold)
    worksheet.write('B2', 1000)
    worksheet.write('B3', 1200)
    worksheet.write('B4', 900)
    worksheet.write('B5', 1200)
    worksheet.write('B6', '=SUBTOTAL(9,B2:B5)', bold)

    worksheet.write('A7', 'South')
    worksheet.write('A8', 'South')
    worksheet.write('A9', 'South')
    worksheet.write('A10', 'South')
    worksheet.write('A11', 'South Total', bold)

    worksheet.write('B7', 400)
    worksheet.write('B8', 600)
    worksheet.write('B9', 500)
    worksheet.write('B10', 600)
    worksheet.write('B11', '=SUBTOTAL(9,B7:B10)', bold)

    worksheet.write('A12', 'Grand Total', bold)
    worksheet.write('B12', '=SUBTOTAL(9,B2:B10)', bold)

###############################################################################
#
# Example 1: A worksheet with outlined rows. It also includes SUBTOTAL()
# functions so that it looks like the type of automatic outlines that are
# generated when you use the Excel Data->SubTotals menu item.
#
worksheet1.set_row(1, None, None, {'level': 2})
worksheet1.set_row(2, None, None, {'level': 2})
worksheet1.set_row(3, None, None, {'level': 2})
worksheet1.set_row(4, None, None, {'level': 2})
worksheet1.set_row(5, None, None, {'level': 1})

worksheet1.set_row(6, None, None, {'level': 2})
worksheet1.set_row(7, None, None, {'level': 2})
worksheet1.set_row(8, None, None, {'level': 2})
worksheet1.set_row(9, None, None, {'level': 2})
worksheet1.set_row(10, None, None, {'level': 1})

# Write the sub-total data that is common to the row examples.
create_sub_totals(worksheet1)


###############################################################################
#
# Example 2: Create a worksheet with collapsed outlined rows.
# This is the same as the example 1  except that the all rows are collapsed.
# Note: We need to indicate the rows that contains the collapsed symbol '+'
# with the optional parameter, 'collapsed'.
#
worksheet2.set_row(1, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(2, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(3, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(4, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(5, None, None, {'level': 1, 'hidden': True})

worksheet2.set_row(6, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(7, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(8, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(9, None, None, {'level': 2, 'hidden': True})
worksheet2.set_row(10, None, None, {'level': 1, 'hidden': True})

worksheet2.set_row(11, None, None, {'collapsed': True})

# Write the sub-total data that is common to the row examples.
create_sub_totals(worksheet2)


###############################################################################
#
# Example 3: Create a worksheet with collapsed outlined rows.
# Same as the example 1  except that the two sub-totals are collapsed.
#
worksheet3.set_row(1, None, None, {'level': 2, 'hidden': True})
worksheet3.set_row(2, None, None, {'level': 2, 'hidden': True})
worksheet3.set_row(3, None, None, {'level': 2, 'hidden': True})
worksheet3.set_row(4, None, None, {'level': 2, 'hidden': True})
worksheet3.set_row(5, None, None, {'level': 1, 'collapsed': True})

worksheet3.set_row(6, None, None, {'level': 2, 'hidden': True})
worksheet3.set_row(7, None, None, {'level': 2, 'hidden': True})
worksheet3.set_row(8, None, None, {'level': 2, 'hidden': True})
worksheet3.set_row(9, None, None, {'level': 2, 'hidden': True})
worksheet3.set_row(10, None, None, {'level': 1, 'collapsed': True})

# Write the sub-total data that is common to the row examples.
create_sub_totals(worksheet3)


###############################################################################
#
# Example 4: Create a worksheet with outlined rows.
# Same as the example 1  except that the two sub-totals are collapsed.
#
worksheet4.set_row(1, None, None, {'level': 2, 'hidden': True})
worksheet4.set_row(2, None, None, {'level': 2, 'hidden': True})
worksheet4.set_row(3, None, None, {'level': 2, 'hidden': True})
worksheet4.set_row(4, None, None, {'level': 2, 'hidden': True})
worksheet4.set_row(5, None, None, {'level': 1, 'hidden': True,
                                   'collapsed': True})

worksheet4.set_row(6, None, None, {'level': 2, 'hidden': True})
worksheet4.set_row(7, None, None, {'level': 2, 'hidden': True})
worksheet4.set_row(8, None, None, {'level': 2, 'hidden': True})
worksheet4.set_row(9, None, None, {'level': 2, 'hidden': True})
worksheet4.set_row(10, None, None, {'level': 1, 'hidden': True,
                                    'collapsed': True})

worksheet4.set_row(11, None, None, {'collapsed': True})


# Write the sub-total data that is common to the row examples.
create_sub_totals(worksheet4)


###############################################################################
#
# Example 5: Create a worksheet with outlined columns.
#
data = [
    ['Month', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Total'],
    ['North', 50, 20, 15, 25, 65, 80, '=SUM(B2:G2)'],
    ['South', 10, 20, 30, 50, 50, 50, '=SUM(B3:G3)'],
    ['East', 45, 75, 50, 15, 75, 100, '=SUM(B4:G4)'],
    ['West', 15, 15, 55, 35, 20, 50, '=SUM(B5:G5)']]

# Add bold format to the first row.
worksheet5.set_row(0, None, bold)

# Set column formatting and the outline level.
worksheet5.set_column('A:A', 10, bold)
worksheet5.set_column('B:G', 5, None, {'level': 1})
worksheet5.set_column('H:H', 10)

# Write the data and a formula.
for row, data_row in enumerate(data):
    worksheet5.write_row(row, 0, data_row)

worksheet5.write('H6', '=SUM(H2:H5)', bold)

###############################################################################
#
# Example 6: Create a worksheet with collapsed outlined columns.
# This is the same as the previous example except with collapsed columns.
#

# Reuse the data from the previous example.

# Add bold format to the first row.
worksheet6.set_row(0, None, bold)

# Set column formatting and the outline level.
worksheet6.set_column('A:A', 10, bold)
worksheet6.set_column('B:G', 5, None, {'level': 1, 'hidden': True})
worksheet6.set_column('H:H', 10, None, {'collapsed': True})

# Write the data and a formula.
for row, data_row in enumerate(data):
    worksheet6.write_row(row, 0, data_row)

worksheet6.write('H6', '=SUM(H2:H5)', bold)

workbook.close()
