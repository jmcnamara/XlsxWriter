###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.format import Format
from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.worksheet import Worksheet


class TestAutofit(unittest.TestCase):
    """
    Test the Worksheet autofit() method.

    """

    def setUp(self):
        self.worksheet = Worksheet()
        self.worksheet.str_table = SharedStringTable()

    def autofit_widths(self):
        """Get the autofit column widths in column order."""
        return sorted(
            (col_num, col_info.width)
            for col_num, col_info in self.worksheet.col_info.items()
            if col_info.autofit
        )

    def test_autofit_empty_worksheet(self):
        """Check for no column changes with an empty worksheet."""
        self.worksheet.autofit()

        self.assertEqual(self.autofit_widths(), [])

    def test_autofit_sparse_data(self):
        """Check for the widths of the stored cells only with sparse data."""
        worksheet = self.worksheet

        # Write a diagonal of strings and some cells far from the others.
        for i in range(20):
            worksheet.write_string(i, i, "x" * (i + 1))

        worksheet.write_number(100, 30, 12345)
        worksheet.write_string(2, 16000, "Far away")
        worksheet.write_blank(3, 16383, None, Format())

        worksheet.autofit()

        exp = [(col_num, 14 + 7 * col_num) for col_num in range(20)]
        exp += [(30, 42), (16000, 61)]

        self.assertEqual(self.autofit_widths(), exp)
