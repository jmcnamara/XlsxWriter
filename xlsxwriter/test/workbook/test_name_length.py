###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
import warnings
from io import BytesIO, StringIO

from xlsxwriter.workbook import Workbook


class TestNameLength(unittest.TestCase):
    """
    Test Excel's 255-character name length limit for define_name() and
    add_table().

    """

    def test_define_name_rejects_overlong_name(self):
        """define_name() should reject names longer than 255 characters."""

        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)
        workbook.add_worksheet()

        with warnings.catch_warnings(record=True) as warned:
            warnings.simplefilter("always")
            result = workbook.define_name("n" * 256, "=Sheet1!$A$1")

        self.assertEqual(result, -1)
        self.assertEqual(len(warned), 1)
        self.assertTrue(issubclass(warned[0].category, UserWarning))
        self.assertIn("255", str(warned[0].message))
        self.assertEqual(workbook.defined_names, [])

        workbook.fileclosed = True

    def test_define_name_accepts_max_length_name(self):
        """define_name() should accept names of exactly 255 characters."""

        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)
        workbook.add_worksheet()

        with warnings.catch_warnings(record=True) as warned:
            warnings.simplefilter("always")
            result = workbook.define_name("n" * 255, "=Sheet1!$A$1")

        self.assertEqual(result, 0)
        self.assertEqual(warned, [])
        self.assertEqual(len(workbook.defined_names), 1)
        self.assertEqual(workbook.defined_names[0][0], "n" * 255)

        workbook.fileclosed = True

    def test_add_table_rejects_overlong_name(self):
        """add_table() should reject names longer than 255 characters."""

        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)
        worksheet = workbook.add_worksheet()

        with warnings.catch_warnings(record=True) as warned:
            warnings.simplefilter("always")
            result = worksheet.add_table("A1:B2", {"name": "T" * 256})

        self.assertEqual(result, -2)
        self.assertEqual(len(warned), 1)
        self.assertTrue(issubclass(warned[0].category, UserWarning))
        self.assertIn("255", str(warned[0].message))
        self.assertEqual(worksheet.tables, [])

        workbook.fileclosed = True

    def test_add_table_accepts_max_length_name(self):
        """add_table() should accept names of exactly 255 characters."""

        out = BytesIO()
        workbook = Workbook(out)
        worksheet = workbook.add_worksheet()

        with warnings.catch_warnings(record=True) as warned:
            warnings.simplefilter("always")
            result = worksheet.add_table("A1:B2", {"name": "T" * 255})

        self.assertEqual(result, 0)
        self.assertEqual(warned, [])
        self.assertEqual(len(worksheet.tables), 1)
        self.assertEqual(worksheet.tables[0]["name"], "T" * 255)

        workbook.close()
