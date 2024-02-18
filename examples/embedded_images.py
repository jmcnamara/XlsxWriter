##############################################################################
#
# An example of embedding images into a worksheet cells using the XlsxWriter
# Python module. This is equivalent to Excel's "Place in cell" image insert.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook("embedded_images.xlsx")
worksheet = workbook.add_worksheet()

# Widen the first column to make the caption clearer.
worksheet.set_column(0, 0, 30)
worksheet.write(0, 0, "Embed images that scale to cell size")

# Embed an images in cells of different widths/heights.
worksheet.set_column(1, 1, 14)

worksheet.set_row(1, 60)
worksheet.embed_image(1, 1, "python.png")

worksheet.set_row(3, 120)
worksheet.embed_image(3, 1, "python.png")

workbook.close()
