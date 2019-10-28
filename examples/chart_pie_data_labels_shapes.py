import xlsxwriter

workbook = xlsxwriter.Workbook('chart_pie_data_labels_shapes.xlsx')

worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})

# Add the worksheet data that the charts will refer to.
headings = ['Category', 'Values']
data = [
    ['Apple', 'Cherry', 'Pecan'],
    [60, 30, 10],
]

worksheet.write_row('A1', headings, bold)
worksheet.write_column('A2', data[0])
worksheet.write_column('B2', data[1])

#######################################################################
#
# Create a new chart object.
#
chart = workbook.add_chart({'type': 'pie'})

# Configure the series.
chart.add_series({
    'name': 'Food',
    'categories': ['Sheet1', 1, 0, 3, 0],
    'values':     ['Sheet1', 1, 1, 3, 1],
    # Format label options.
    'data_labels': {'category': True, 'percentage': True,
                    'separator': '\n', 'position': 'outside_end',
                    'shape': 'rectangular_callout'}
})

chart.set_legend({'none': True})

# Insert the chart into the worksheet.
worksheet.insert_chart('E2', chart)

workbook.close()
