#######################################################################
#
# An example of using the new Excel LAMBDA() function with the XlsxWriter
# module. Note, this function is only currently available if you are
# subscribed to the Microsoft Office Beta program.
#
# Copyright 2013-2021, John McNamara, jmcnamara@cpan.org
#
from xlsxwriter.workbook import Workbook

workbook = Workbook('lambda.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1',
                'Note: Lambda functions currently only work with '
                'the Beta Channel versions of Excel 365')

# Write a Lambda function to convert Fahrenheit to Celsius to a cell.
#
# Note that the lambda function parameters must be prefixed with
# "_xlpm.". These prefixes won't show up in Excel.
worksheet.write_da_formula('A2',
                           '=LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(32)')

# Create the same formula (without an argument) as a defined name and use that
# to calculate a value.
#
# Note that the formula name is prefixed with "_xlfn." (this is normally
# converted automatically by write_da_formula() but isn't for defined names)
# and note that the lambda function parameters are prefixed with
# "_xlpm.". These prefixes won't show up in Excel.
workbook.define_name('ToCelsius',
                     '=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))')

worksheet.write_da_formula('A3', '=ToCelsius(212)')

workbook.close()
