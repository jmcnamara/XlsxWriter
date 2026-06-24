###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import os
import tempfile
import unittest
from dataclasses import dataclass

from xlsxwriter.workbook import Workbook


class TestWriteUnknownTypes(unittest.TestCase):
    """
    Test write() handling of types that are not natively supported.

    When float() raises TypeError for an unsupported type, write() should fall
    through to the str() fallback rather than immediately re-raising. Both
    paths are tested here.
    """

    def _make_workbook(self, tmp_path):
        """Create a Workbook backed by a temp file."""
        return Workbook(tmp_path)

    def test_write_dataclass_falls_back_to_str(self):
        """write() should use str() for a dataclass instead of raising TypeError."""

        @dataclass
        class Point:
            x: int
            y: int

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            tmp = f.name
        try:
            workbook = Workbook(tmp)
            worksheet = workbook.add_worksheet()
            # Should NOT raise; should write str(Point(1, 2)) as a string cell.
            worksheet.write(0, 0, Point(1, 2))
            workbook.close()
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)

    def test_write_custom_object_falls_back_to_str(self):
        """write() should use str() for objects that do not support float()."""

        class MyObj:
            def __str__(self):
                return "custom_value"

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            tmp = f.name
        try:
            workbook = Workbook(tmp)
            worksheet = workbook.add_worksheet()
            worksheet.write(0, 0, MyObj())
            workbook.close()
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)

    def test_write_object_with_float_value_error_falls_back_to_str(self):
        """write() str() fallback is used when __float__ raises ValueError."""

        class WeirdFloat:
            def __float__(self):
                raise ValueError("not a number")

            def __str__(self):
                return "weird_object"

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            tmp = f.name
        try:
            workbook = Workbook(tmp)
            worksheet = workbook.add_worksheet()
            worksheet.write(0, 0, WeirdFloat())
            workbook.close()
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)

    def test_write_truly_unsupported_type_raises_type_error(self):
        """write() should raise TypeError when neither float() nor str() works."""

        class Unsupported:
            def __float__(self):
                raise TypeError("no float")

            def __str__(self):
                raise TypeError("no str")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            tmp = f.name
        try:
            workbook = Workbook(tmp)
            worksheet = workbook.add_worksheet()
            with self.assertRaises(TypeError):
                worksheet.write(0, 0, Unsupported())
            workbook.close()
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)

    def test_write_row_with_mixed_types_including_dataclass(self):
        """write_row() should handle a row mixing standard types and a dataclass."""

        @dataclass
        class Payload:
            value: str

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            tmp = f.name
        try:
            workbook = Workbook(tmp)
            worksheet = workbook.add_worksheet()
            worksheet.write_row(
                0, 0, [1, True, 1.0, "hello", Payload(value="world")]
            )
            workbook.close()
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)


if __name__ == "__main__":
    unittest.main()
