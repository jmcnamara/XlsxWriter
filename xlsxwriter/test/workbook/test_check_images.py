###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.exceptions import UndefinedImageSize, UnsupportedImageFormat
from xlsxwriter.workbook import Workbook


class TestInsertImage(unittest.TestCase):
    """
    Test exceptions with insert_image().

    """

    def test_undefined_image_size(self):
        """Test adding an image with no height/width data."""

        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        worksheet = workbook.add_worksheet()

        with self.assertRaises(UndefinedImageSize):
            worksheet.insert_image(
                "B13", "xlsxwriter/test/comparison/images/nosize.png"
            )

        workbook.fileclosed = True

    def test_unsupported_image(self):
        """Test adding an unsupported image type."""

        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        worksheet = workbook.add_worksheet()

        with self.assertRaises(UnsupportedImageFormat):
            worksheet.insert_image(
                "B13", "xlsxwriter/test/comparison/images/unsupported.txt"
            )

        workbook.fileclosed = True
