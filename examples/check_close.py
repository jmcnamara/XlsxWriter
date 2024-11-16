##############################################################################
#
# A simple program demonstrating a check for exceptions when closing the file.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2024, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook("check_close.xlsx")
worksheet = workbook.add_worksheet()

worksheet.write("A1", "Hello world")

# Try to close() the file in a loop so that if there is an exception, such as
# if the file is open in Excel, we can ask the user to close the file, and
# try again to overwrite it.
while True:
    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError as e:
        decision = input(
            f"Exception caught in workbook.close(): {e}\n"
            f"Please close the file if it is open in Excel.\n"
            f"Try to write file again? [Y/n]: "
        )
        if decision != "n":
            continue

    break
