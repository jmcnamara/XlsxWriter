###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#

import unittest
from io import BytesIO

from xlsxwriter.workbook import Workbook


class TestPixelOffsets(unittest.TestCase):
    """
    Test the Worksheet pixel offset caches.

    """

    def setUp(self):
        self.workbook = Workbook(BytesIO(), {"in_memory": True})
        self.worksheet = self.workbook.add_worksheet()

    def tearDown(self):
        self.workbook.close()

    def test_default_pixel_offsets(self):
        """Test offsets with default row heights and column widths."""
        worksheet = self.worksheet

        self.assertEqual(worksheet.col_pixel_offsets, [0])
        self.assertEqual(worksheet.row_pixel_offsets, [0])

        for index in (0, 1, 5, 2, 5):
            self.assertEqual(worksheet._col_pixel_offset(index), index * 64)
            self.assertEqual(worksheet._row_pixel_offset(index), index * 20)

        self.assertEqual(worksheet.col_pixel_offsets, [0, 64, 128, 192, 256, 320])
        self.assertEqual(worksheet.row_pixel_offsets, [0, 20, 40, 60, 80, 100])

    def test_custom_pixel_offsets(self):
        """Test custom sizes with cache growth and non-sequential lookups."""
        worksheet = self.worksheet
        worksheet.set_column_pixels(2, 2, 100)
        worksheet.set_row_pixels(3, 40)

        cases = (
            (0, 0, 0),
            (1, 64, 20),
            (2, 128, 40),
            (5, 356, 120),
            (10, 676, 220),
            (4, 292, 100),
        )

        for index, col_offset, row_offset in cases:
            self.assertEqual(worksheet._col_pixel_offset(index), col_offset)
            self.assertEqual(worksheet._row_pixel_offset(index), row_offset)
