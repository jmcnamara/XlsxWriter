##############################################################################
#
# A simple program demonstrating a check for exceptions when closing the file.
#
# Copyright 2013-2021, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook('check_close.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Hello world')

# Try to close() the file in a loop so that if there is an exception, such as
# if the file is open in Excel, we can ask the user to close the file, and
# try again to overwrite it.
while True:
    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError as e:
        # For Python 3 use input() instead of raw_input().
        decision = raw_input("Exception caught in workbook.close(): %s\n"
                             "Please close the file if it is open in Excel.\n"
                             "Try to write file again? [Y/n]: " % e)
        if decision != 'n':
            continue

    break
