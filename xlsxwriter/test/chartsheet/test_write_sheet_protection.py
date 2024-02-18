###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...chartsheet import Chartsheet


class TestWriteSheetProtection(unittest.TestCase):
    """
    Test the Chartsheet _write_sheet_protection() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.chartsheet = Chartsheet()
        self.chartsheet._set_filehandle(self.fh)

    def test_write_sheet_protection_1(self):
        """Test the _write_sheet_protection() method."""

        password = ""
        options = {}

        self.chartsheet.protect(password, options)
        self.chartsheet._write_sheet_protection()

        exp = """<sheetProtection content="1" objects="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_2(self):
        """Test the _write_sheet_protection() method."""

        password = "password"
        options = {}

        self.chartsheet.protect(password, options)
        self.chartsheet._write_sheet_protection()

        exp = """<sheetProtection password="83AF" content="1" objects="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_3(self):
        """Test the _write_sheet_protection() method."""

        password = ""
        options = {"objects": False}

        self.chartsheet.protect(password, options)
        self.chartsheet._write_sheet_protection()

        exp = """<sheetProtection content="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_4(self):
        """Test the _write_sheet_protection() method."""

        password = ""
        options = {"content": False}

        self.chartsheet.protect(password, options)
        self.chartsheet._write_sheet_protection()

        exp = """<sheetProtection objects="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_5(self):
        """Test the _write_sheet_protection() method."""

        password = ""
        options = {"content": False, "objects": False}

        self.chartsheet.protect(password, options)
        self.chartsheet._write_sheet_protection()

        exp = ""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_6(self):
        """Test the _write_sheet_protection() method."""

        password = "password"
        options = {"content": False, "objects": False}

        self.chartsheet.protect(password, options)
        self.chartsheet._write_sheet_protection()

        exp = """<sheetProtection password="83AF"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_protection_7(self):
        """Test the _write_sheet_protection() method."""

        password = "password"
        options = {
            "objects": True,
            "scenarios": True,
            "format_cells": True,
            "format_columns": True,
            "format_rows": True,
            "insert_columns": True,
            "insert_rows": True,
            "insert_hyperlinks": True,
            "delete_columns": True,
            "delete_rows": True,
            "select_locked_cells": False,
            "sort": True,
            "autofilter": True,
            "pivot_tables": True,
            "select_unlocked_cells": False,
        }

        self.chartsheet.protect(password, options)
        self.chartsheet._write_sheet_protection()

        exp = """<sheetProtection password="83AF" content="1" objects="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
