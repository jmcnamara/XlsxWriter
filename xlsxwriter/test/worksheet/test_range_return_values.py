###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.worksheet import Worksheet


class TestRangeReturnValues(unittest.TestCase):
    """
    Test the return value for various functions that handle 1 or 2D ranges.

    """

    MAX_ROW = 1048576
    MAX_COL = 16384
    BOUND_ERROR = -1

    def setUp(self):
        self.ws = Worksheet()

    def test_write_string_row_out_of_bounds(self):
        got = self.ws.write_string(self.MAX_ROW, 0, "Foo")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_string_col_out_of_bounds(self):
        got = self.ws.write_string(0, self.MAX_COL, "Foo")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_string_row_and_col_out_of_bounds(self):
        got = self.ws.write_string(self.MAX_ROW, self.MAX_COL, "Foo")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_number_row_out_of_bounds(self):
        got = self.ws.write_number(self.MAX_ROW, 0, 123)
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_number_col_out_of_bounds(self):
        got = self.ws.write_number(0, self.MAX_COL, 123)
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_number_row_and_col_out_of_bounds(self):
        got = self.ws.write_number(self.MAX_ROW, self.MAX_COL, 123)
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_blank_row_out_of_bounds(self):
        got = self.ws.write_blank(self.MAX_ROW, 0, None, "format")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_blank_col_out_of_bounds(self):
        got = self.ws.write_blank(0, self.MAX_COL, None, "format")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_blank_row_and_col_out_of_bounds(self):
        got = self.ws.write_blank(self.MAX_ROW, self.MAX_COL, None, "format")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_formula_row_out_of_bounds(self):
        got = self.ws.write_formula(self.MAX_ROW, 0, "=A1")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_formula_col_out_of_bounds(self):
        got = self.ws.write_formula(0, self.MAX_COL, "=A1")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_formula_row_and_col_out_of_bounds(self):
        got = self.ws.write_formula(self.MAX_ROW, self.MAX_COL, "=A1")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_array_formula_last_col_out_of_bounds(self):
        got = self.ws.write_array_formula(0, 0, 0, self.MAX_COL, "=A1")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_array_formula_last_row_out_of_bounds(self):
        got = self.ws.write_array_formula(0, 0, self.MAX_ROW, 0, "=A1")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_array_formula_first_col_out_of_bounds(self):
        got = self.ws.write_array_formula(0, self.MAX_COL, 0, 0, "=A1")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_array_formula_first_row_out_of_bounds(self):
        got = self.ws.write_array_formula(self.MAX_ROW, 0, 0, 0, "=A1")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_write_array_formula_all_out_of_bounds(self):
        got = self.ws.write_array_formula(
            self.MAX_ROW, self.MAX_COL, self.MAX_ROW, self.MAX_COL, "=A1"
        )
        self.assertEqual(got, self.BOUND_ERROR)

    def test_merge_range_last_col_out_of_bounds(self):
        got = self.ws.merge_range(0, 0, 0, self.MAX_COL, "Foo")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_merge_range_last_row_out_of_bounds(self):
        got = self.ws.merge_range(0, 0, self.MAX_ROW, 0, "Foo")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_merge_range_first_col_out_of_bounds(self):
        got = self.ws.merge_range(0, self.MAX_COL, 0, 0, "Foo")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_merge_range_first_row_out_of_bounds(self):
        got = self.ws.merge_range(self.MAX_ROW, 0, 0, 0, "Foo")
        self.assertEqual(got, self.BOUND_ERROR)

    def test_set_column_last_col_out_of_bounds(self):
        got = self.ws.set_column(6, self.MAX_COL, 17)
        self.assertEqual(got, self.BOUND_ERROR)

    def test_set_column_first_col_out_of_bounds(self):
        got = self.ws.set_column(self.MAX_COL, 6, 17)
        self.assertEqual(got, self.BOUND_ERROR)

    def test_set_row_out_of_bounds(self):
        got = self.ws.set_row(self.MAX_ROW, 30)
        self.assertEqual(got, self.BOUND_ERROR)

    def test_string_length_return_values(self):
        """Test writing strings at Excel's length limit in UTF-16 units."""
        worksheet = Worksheet()
        worksheet.str_table = SharedStringTable()

        # Excel's limit is in UTF-16 units, so these emojis count as 2.
        got = worksheet.write_string(0, 0, "\U0001f600" * 16383 + "a")
        self.assertEqual(got, 0)

        got = worksheet.write_string(1, 0, "\U0001f600" * 16384)
        self.assertEqual(got, -2)

        # The string is truncated without splitting an emoji.
        self.assertIn("\U0001f600" * 16383, worksheet.str_table.string_table)
