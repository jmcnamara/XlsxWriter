##############################################################################
#
# An example of inserting images from a Python BytesIO byte stream into a
# worksheet using the XlsxWriter module.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2022, John McNamara, jmcnamara@cpan.org
#

from io import BytesIO
from urllib.request import urlopen


import xlsxwriter

# Create the workbook and add a worksheet.
workbook  = xlsxwriter.Workbook('images_bytesio.xlsx')
worksheet = workbook.add_worksheet()


# Read an image from a remote url.
url = 'https://raw.githubusercontent.com/jmcnamara/XlsxWriter/' + \
      'master/examples/logo.png'

image_data = BytesIO(urlopen(url).read())

# Write the byte stream image to a cell. Note, the filename must be
# specified. In this case it will be read from url string.
worksheet.insert_image('B2', url, {'image_data': image_data})


# Read a local image file into a a byte stream. Note, the insert_image()
# method can do this directly. This is for illustration purposes only.
filename   = 'python.png'

image_file = open(filename, 'rb')
image_data = BytesIO(image_file.read())
image_file.close()


# Write the byte stream image to a cell. The filename must  be specified.
worksheet.insert_image('B8', filename, {'image_data': image_data})


workbook.close()
