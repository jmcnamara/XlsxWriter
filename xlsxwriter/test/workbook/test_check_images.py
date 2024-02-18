###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...workbook import Workbook
from ...exceptions import UndefinedImageSize
from ...exceptions import UnsupportedImageFormat


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

        worksheet.insert_image("B13", "xlsxwriter/test/comparison/images/nosize.png")

        self.assertRaises(UndefinedImageSize, workbook._prepare_drawings)

        workbook.fileclosed = True

    def test_unsupported_image(self):
        """Test adding an unsupported image type."""

        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        worksheet = workbook.add_worksheet()

        worksheet.insert_image(
            "B13", "xlsxwriter/test/comparison/images/unsupported.txt"
        )

        self.assertRaises(UnsupportedImageFormat, workbook._prepare_drawings)

        workbook.fileclosed = True
