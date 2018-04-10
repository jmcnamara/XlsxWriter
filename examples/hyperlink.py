###############################################################################
#
# Example of how to use the XlsxWriter module to write hyperlinks
#
# Copyright 2013-2018, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

# Create a new workbook and add a worksheet
workbook = xlsxwriter.Workbook('hyperlink.xlsx')
worksheet = workbook.add_worksheet('Hyperlinks')

# Format the first column
worksheet.set_column('A:A', 30)

# Add a sample alternative link format.
red_format = workbook.add_format({
    'font_color': 'red',
    'bold':       1,
    'underline':  1,
    'font_size':  12,
})

# Write some hyperlinks
worksheet.write_url('A1', 'http://www.python.org/')  # Implicit format.
worksheet.write_url('A3', 'http://www.python.org/', string='Python Home')
worksheet.write_url('A5', 'http://www.python.org/', tip='Click here')
worksheet.write_url('A7', 'http://www.python.org/', red_format)
worksheet.write_url('A9', 'mailto:jmcnamara@cpan.org', string='Mail me')

# Write a URL that isn't a hyperlink
worksheet.write_string('A11', 'http://www.python.org/')

workbook.close()
