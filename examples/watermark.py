##############################################################################
#
# An example of adding a worksheet watermark image using the XlsxWriter Python
# module. This is based on the method of putting an image in the worksheet
# header as suggested in the Microsoft documentation:
# https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2023, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter

workbook = xlsxwriter.Workbook("watermark.xlsx")
worksheet = workbook.add_worksheet()

# Set a worksheet header with the watermark image.
worksheet.set_header("&C&[Picture]", {"image_center": "watermark.png"})

workbook.close()
