##############################################################################
#
# Example of how to subclass the Workbook and Worksheet objects. We also
# override the default worksheet.write() method to show how that is done.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.worksheet import convert_cell_args


class MyWorksheet(Worksheet):
    """
    Subclass of the XlsxWriter Worksheet class to override the default
    write() method.

    """

    @convert_cell_args
    def write(self, row, col, *args):
        data = args[0]

        # Reverse strings to demonstrate the overridden method.
        if isinstance(data, str):
            data = data[::-1]
            return self.write_string(row, col, data)
        else:
            # Call the parent version of write() as usual for other data.
            return super(MyWorksheet, self).write(row, col, *args)


class MyWorkbook(Workbook):
    """
    Subclass of the XlsxWriter Workbook class to override the default
    Worksheet class with our custom class.

    """

    def add_worksheet(self, name=None):
        # Overwrite add_worksheet() to create a MyWorksheet object.
        worksheet = super(MyWorkbook, self).add_worksheet(name, MyWorksheet)

        return worksheet


# Create a new MyWorkbook object.
workbook = MyWorkbook("inheritance1.xlsx")

# The code from now on will be the same as a normal "Workbook" program.
worksheet = workbook.add_worksheet()

# Write some data to test the subclassing.
worksheet.write("A1", "Hello")
worksheet.write("A2", "World")
worksheet.write("A3", 123)
worksheet.write("A4", 345)

workbook.close()
