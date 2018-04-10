#######################################################################
#
# An example of inserting textboxes into an Excel worksheet using
# Python and XlsxWriter.
#
# Copyright 2013-2018, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('textbox.xlsx')
worksheet = workbook.add_worksheet()
row = 4
col = 1

# The examples below show different textbox options and formatting. In each
# example the text describes the formatting.


# Example
text = 'A simple textbox with some text'
worksheet.insert_textbox(row, col, text)
row += 10

# Example
text = 'A textbox with changed dimensions'
options = {
    'width': 256,
    'height': 100,
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'A textbox with an offset in the cell'
options = {
    'x_offset': 10,
    'y_offset': 10,
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'A textbox with scaling'
options = {
    'x_scale': 1.5,
    'y_scale': 0.8,
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'A textbox with some long text that wraps around onto several lines'
worksheet.insert_textbox(row, col, text)
row += 10

# Example
text = 'A textbox\nwith some\nnewlines\n\nand paragraphs'
worksheet.insert_textbox(row, col, text)
row += 10

# Example
text = 'A textbox with a solid fill background'
options = {
    'fill': {'color': 'red'},
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'A textbox with a no fill background'
options = {
    'fill': {'none': True},
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'A textbox with a gradient fill background'
options = {
    'gradient': {'colors': ['#DDEBCF',
                            '#9CB86E',
                            '#156B13']},
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'A textbox with a user defined border line'
options = {
    'border': {'color': 'red',
               'width': 3,
               'dash_type': 'round_dot'},
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'A textbox with no border line'
options = {
    'border': {'none': True},
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'Default alignment: top - left'
worksheet.insert_textbox(row, col, text)
row += 10

# Example
text = 'Alignment: top - center'
options = {
    'align': {'horizontal': 'center'},
}
worksheet.insert_textbox(row, col, text)
row += 10

# Example
text = 'Alignment: top - center'
options = {
    'align': {'horizontal': 'center'},
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'Alignment: middle - center'
options = {
    'align': {'vertical': 'middle',
               'horizontal': 'center'},
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'Font properties: bold'
options = {
    'font': {'bold': True},
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'Font properties: various'
options = {
    'font': {'bold': True},
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'Font properties: various'
options = {
    'font': {'bold': True,
             'italic': True,
             'underline': True,
             'name': 'Arial',
             'color': 'red',
             'size': 12}
}
worksheet.insert_textbox(row, col, text, options)
row += 10

# Example
text = 'Some text in a textbox with formatting'
options = {
    'font': {'color': 'white'},
    'align': {'vertical': 'middle',
              'horizontal': 'center'
              },
    'gradient': {'colors': ['red', 'blue']},
}
worksheet.insert_textbox(row, col, text, options)
row += 10


workbook.close()
