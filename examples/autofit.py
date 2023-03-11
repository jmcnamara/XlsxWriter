#######################################################################
#
# An example of using simulated autofit to automatically adjust the width of
# worksheet columns based on the data in the cells.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
from xlsxwriter.workbook import Workbook

workbook = Workbook("autofit.xlsx")
worksheet = workbook.add_worksheet()

# Write some worksheet data to demonstrate autofitting.
worksheet.write(0, 0, "Foo")
worksheet.write(1, 0, "Food")
worksheet.write(2, 0, "Foody")
worksheet.write(3, 0, "Froody")

worksheet.write(0, 1, 12345)
worksheet.write(1, 1, 12345678)
worksheet.write(2, 1, 12345)

worksheet.write(0, 2, "Some longer text")

worksheet.write(0, 3, "http://ww.google.com")
worksheet.write(1, 3, "https://github.com")

# Autofit the worksheet.
worksheet.autofit()

workbook.close()
