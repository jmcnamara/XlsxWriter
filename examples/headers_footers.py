######################################################################
#
# This program shows several examples of how to set up headers and
# footers with XlsxWriter.
#
# The control characters used in the header/footer strings are:
#
#     Control             Category            Description
#     =======             ========            ===========
#     &L                  Justification       Left
#     &C                                      Center
#     &R                                      Right
#
#     &P                  Information         Page number
#     &N                                      Total number of pages
#     &D                                      Date
#     &T                                      Time
#     &F                                      File name
#     &A                                      Worksheet name
#
#     &fontsize           Font                Font size
#     &"font,style"                           Font name and style
#     &U                                      Single underline
#     &E                                      Double underline
#     &S                                      Strikethrough
#     &X                                      Superscript
#     &Y                                      Subscript
#
#     &[Picture]          Images              Image placeholder
#     &G                                      Same as &[Picture]
#
#     &&                  Miscellaneous       Literal ampersand &
#
# See the main XlsxWriter documentation for more information.
#
# Copyright 2013-2015, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('headers_footers.xlsx')
preview = 'Select Print Preview to see the header and footer'

######################################################################
#
# A simple example to start
#
worksheet1 = workbook.add_worksheet('Simple')
header1 = '&CHere is some centred text.'
footer1 = '&LHere is some left aligned text.'

worksheet1.set_header(header1)
worksheet1.set_footer(footer1)

worksheet1.set_column('A:A', 50)
worksheet1.write('A1', preview)


######################################################################
#
# Insert a header image.
#
worksheet2 = workbook.add_worksheet('Image')
header2 = '&L&G'

# Adjust the page top margin to allow space for the header image.
worksheet2.set_margins(top=1.3)

worksheet2.set_header(header2, {'image_left': 'python-200x80.png'})

worksheet2.set_column('A:A', 50)
worksheet2.write('A1', preview)


######################################################################
#
# This is an example of some of the header/footer variables.
#
worksheet3 = workbook.add_worksheet('Variables')
header3 = '&LPage &P of &N' + '&CFilename: &F' + '&RSheetname: &A'
footer3 = '&LCurrent date: &D' + '&RCurrent time: &T'

worksheet3.set_header(header3)
worksheet3.set_footer(footer3)

worksheet3.set_column('A:A', 50)
worksheet3.write('A1', preview)
worksheet3.write('A21', 'Next sheet')
worksheet3.set_h_pagebreaks([20])

######################################################################
#
# This example shows how to use more than one font
#
worksheet4 = workbook.add_worksheet('Mixed fonts')
header4 = '&C&"Courier New,Bold"Hello &"Arial,Italic"World'
footer4 = '&C&"Symbol"e&"Arial" = mc&X2'

worksheet4.set_header(header4)
worksheet4.set_footer(footer4)

worksheet4.set_column('A:A', 50)
worksheet4.write('A1', preview)

######################################################################
#
# Example of line wrapping
#
worksheet5 = workbook.add_worksheet('Word wrap')
header5 = "&CHeading 1\nHeading 2"

worksheet5.set_header(header5)

worksheet5.set_column('A:A', 50)
worksheet5.write('A1', preview)

######################################################################
#
# Example of inserting a literal ampersand &
#
worksheet6 = workbook.add_worksheet('Ampersand')
header6 = '&CCuriouser && Curiouser - Attorneys at Law'

worksheet6.set_header(header6)

worksheet6.set_column('A:A', 50)
worksheet6.write('A1', preview)

workbook.close()
