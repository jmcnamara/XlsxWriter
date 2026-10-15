###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#

import unittest
from unittest.mock import patch

from xlsxwriter.worksheet import Worksheet


class TestPositionObjectPixels(unittest.TestCase):
    """Test object positions with uniform and individually sized rows."""

    def setUp(self):
        self.worksheet = Worksheet()
        self.worksheet.set_default_row(24)

    def position(self, row, y_offset=0, anchor=0):
        return self.worksheet._position_object_pixels(
            2, row, 10, y_offset, 128, 74, anchor
        )

    def test_uniform_row_height(self):
        """A custom default height gives the same complete object geometry."""
        self.assertEqual(
            self.position(100_000, y_offset=7),
            [2, 100_000, 10, 7, 4, 100_002, 10, 17, 138, 3_200_007],
        )

    def test_uniform_height_skips_earlier_rows(self):
        """Positioning a distant object doesn't visit every preceding row."""
        with patch.object(
            self.worksheet, "_size_row", wraps=self.worksheet._size_row
        ) as size_row:
            self.position(100_000)

        self.assertLess(size_row.call_count, 20)
        self.assertTrue(
            all(call.args[0] >= 100_000 for call in size_row.call_args_list)
        )

    def test_row_overrides_and_hidden_rows(self):
        """Keep the original calculation for individually sized or hidden rows."""
        self.worksheet.set_row(0, 30)
        self.worksheet.set_row(2, 0)
        self.worksheet.set_row_pixels(4, 12.25)

        for anchor in range(5):
            with self.subTest(anchor=anchor):
                self.assertEqual(self.position(10, anchor=anchor)[-1], 276.25)

    def test_default_changed_after_row_override(self):
        """An earlier row override keeps its height when the default changes."""
        self.worksheet.set_row(0, 15)
        self.worksheet.set_default_row(36)

        self.assertEqual(self.position(10)[-1], 452)

    def test_default_changed_twice(self):
        """Uniform positioning uses the latest default height."""
        self.worksheet.set_default_row(36)

        self.assertEqual(self.position(10)[-1], 480)

    def test_negative_row_index(self):
        """Negative indexes retain the earlier empty-range offset."""
        self.assertEqual(self.position(-1)[-1], 0)

    def test_invalid_row_indexes(self):
        """The uniform path retains range()'s errors for non-integer indexes."""
        for row in (None, "3", 5.25, [], {}, (1,)):
            with self.subTest(row=row):
                with self.assertRaises(TypeError) as expected:
                    range(row)
                with self.assertRaises(TypeError) as actual:
                    self.position(row)
                self.assertEqual(str(actual.exception), str(expected.exception))

    def test_negative_offset(self):
        """Moving the start cell for a negative offset keeps the same distance."""
        self.assertEqual(self.position(5, y_offset=-40)[-1], 120)

    def test_row_override_after_positioning(self):
        """Changing a row after an earlier calculation is picked up immediately."""
        self.assertEqual(self.position(20)[-1], 640)
        self.worksheet.set_row(10, 48)
        self.assertEqual(self.position(20)[-1], 672)

    def test_hidden_unused_rows(self):
        """Keep the existing positioning behavior when unused rows are hidden."""
        self.worksheet.set_default_row(24, hide_unused_rows=True)

        for anchor in range(5):
            with self.subTest(anchor=anchor):
                self.assertEqual(self.position(10, anchor=anchor)[-1], 320)

    def test_unchanged_default_height(self):
        """The existing fast path still handles the original default height."""
        self.worksheet = Worksheet()
        self.assertEqual(self.position(10)[-1], 200)
