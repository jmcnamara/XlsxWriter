###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...worksheet import Worksheet


class TestRangeReturnValues(unittest.TestCase):
    """
    Test the return value for various functions that handle 1 or 2D ranges.

    """

    def test_range_return_values(self):
        """Test writing a worksheet with data out of bounds."""
        worksheet = Worksheet()

        max_row = 1048576
        max_col = 16384
        bound_error = -1

        # Test some out of bound values.
        got = worksheet.write_string(max_row, 0, "Foo")
        self.assertEqual(got, bound_error)

        got = worksheet.write_string(0, max_col, "Foo")
        self.assertEqual(got, bound_error)

        got = worksheet.write_string(max_row, max_col, "Foo")
        self.assertEqual(got, bound_error)

        got = worksheet.write_number(max_row, 0, 123)
        self.assertEqual(got, bound_error)

        got = worksheet.write_number(0, max_col, 123)
        self.assertEqual(got, bound_error)

        got = worksheet.write_number(max_row, max_col, 123)
        self.assertEqual(got, bound_error)

        got = worksheet.write_blank(max_row, 0, None, "format")
        self.assertEqual(got, bound_error)

        got = worksheet.write_blank(0, max_col, None, "format")
        self.assertEqual(got, bound_error)

        got = worksheet.write_blank(max_row, max_col, None, "format")
        self.assertEqual(got, bound_error)

        got = worksheet.write_formula(max_row, 0, "=A1")
        self.assertEqual(got, bound_error)

        got = worksheet.write_formula(0, max_col, "=A1")
        self.assertEqual(got, bound_error)

        got = worksheet.write_formula(max_row, max_col, "=A1")
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(0, 0, 0, max_col, "=A1")
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(0, 0, max_row, 0, "=A1")
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(0, max_col, 0, 0, "=A1")
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(max_row, 0, 0, 0, "=A1")
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(max_row, max_col, max_row, max_col, "=A1")
        self.assertEqual(got, bound_error)

        got = worksheet.merge_range(0, 0, 0, max_col, "Foo")
        self.assertEqual(got, bound_error)

        got = worksheet.merge_range(0, 0, max_row, 0, "Foo")
        self.assertEqual(got, bound_error)

        got = worksheet.merge_range(0, max_col, 0, 0, "Foo")
        self.assertEqual(got, bound_error)

        got = worksheet.merge_range(max_row, 0, 0, 0, "Foo")
        self.assertEqual(got, bound_error)

        # Column out of bounds.
        got = worksheet.set_column(6, max_col, 17)
        self.assertEqual(got, bound_error)

        got = worksheet.set_column(max_col, 6, 17)
        self.assertEqual(got, bound_error)

        # Row out of bounds.
        worksheet.set_row(max_row, 30)

        # Reverse man and min column numbers
        worksheet.set_column(0, 3, 17)
